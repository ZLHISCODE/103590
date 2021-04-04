Attribute VB_Name = "mdl吉林市"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

'以下是刘兴宏加入20040924
Private mblnInit As Boolean

Public Enum 业务类型_吉林
    初始化服务调用 = 0
    读卡
    取药品信息
    取服务信息
    取诊疗信息
    设置发票数据
    设置门诊收据明细
    设置门诊大类信息
    取得门诊收据计算信息
    取消结算
    入院登记
    取消入院登记
    出院登记
    取消出院登记
    设置住院帐单
    设置记帐单明细数据
    设置结算单
    设置住院大类信息
    取得住院结算计算信息
    服务数据提交
    政策服务提交
    结束服务调用
End Enum
Private Type InitbaseInfor
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
    明细时实上传 As Boolean
End Type
Public InitInfor_吉林 As InitbaseInfor

Private Type 病人身份
        中心 As String
        卡号  As String
        身份证号 As String
        姓名     As String
        性别     As String
        出生日期  As String
        医保号    As String
        单位编码    As String
        个人身份    As String
        公务员标志  As String   '(0-非公务员1-公务员，其他参照公务员)
        补充保险    As String   '0-不参加1-参加
        大病医保    As String   '0-不参加1-参加
        隶属关系    As String   '1-市属财政0-其他
        照顾级别    As String   '0-无1-一级2-二级3-三级
        职工属地    As String   '0本地1常驻外地2-异地安置
        是否慢性病  As String   '0-不是1-是
        重大疾病    As String   '0-不是1-是
        住院标志    As String   '0-不住院 1-住院
        起付段医疗费累计  As Double
        统筹支付累计 As Double
        进入统筹累计  As Double
        起付线金额累计 As Double
        慢病统筹支付累计 As Double
        大病累计    As Double
        帐户余额    As Double
        住院次数    As Integer
        支付序列 As Integer
        
        费用总额    As Double
        诊断编码    As String
        诊断名称    As String
        病种代码    As String
        门诊类型  As Integer
End Type
Public g病人身份_吉林 As 病人身份



'-----------------------------------------------------------------------------------------------------------------

Private str门诊类型 As String * 1, str保险病种 As String * 10, strTempArr(50, 1) As String

'===============================================================================================================
'功能: 初始服务调用
'入口参数: 服务类型(2)
'说明: 10-门诊收费,11-门诊退费,20-入院登记,21-医嘱录入,22-住院记帐,23-住院结算,24-出院登记,25-取消入院登记
'      26-冲销住院记帐,27-取消结算,28-取消出院登记,29-取消医嘱录入,其他-系统功能/字典功能
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function InitCalc Lib "YhYbClient.dll" (ByVal str服务类型 As String) As Long

'===============================================================================================================
'功能: 结束服务调用
'入口参数: 无
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function FinalCalc Lib "YhYbClient.dll" () As Long

'===============================================================================================================
'功能: 取医保服务名
'入口参数: 无
'出口参数: 无
'返回: 系统正在使用的应用服务器名称
'===============================================================================================================
Public Declare Function GetAppServerName Lib "YhYbClient.dll" () As String

'===============================================================================================================
'功能: 指定提供服务的服务器
'入口参数: 应用服务器名称(windows限制长度)
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetAppServerName Lib "YhYbClient.dll" (ByVal str服务器名 As String) As Long

'===============================================================================================================
'功能: 设置卡读写器端口
'入口参数: 端口号
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetCardPort Lib "YhYbClient.dll" (ByVal str端口号 As String) As Long

'===============================================================================================================
'功能: 获取错误信息
'入口参数: 无
'出口参数: 无
'返回: 错误信息
'===============================================================================================================
Public Declare Function GetErrMsg Lib "YhYbClient.dll" () As String

'===============================================================================================================
'功能: 取医疗机构信息
'入口参数: 无
'出口参数: 中心代码(4位),机构代码(4位),机构名称(40位),医院级别(2位),级别名称(20位)
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function GetInfo_MediOrgan Lib "YhYbClient.dll" (ByVal str中心代码 As String, _
    ByVal str机构代码 As String, ByVal str机构名称 As String, ByVal str医院级别 As String, _
    ByVal str级别名称 As String) As Long

'===============================================================================================================
'功能: 取药品字典信息
'入口参数: 项目编码(10)
'出口参数: 项目名称(40位),费用大类(2位),项目类别(1位),是否医保(1位),是否限价(1位),自付比例(4位宽3位小数)
'          标准单价(8位，2位小数)
'说明: 项目类别:0-甲或普通,1-乙或高精尖,2-自费
'      是否医保:0-不是,1-是
'      是否限价:0-不是,1-是
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function GetInfo_MediDic Lib "YhYbClient.dll" (ByVal str项目编码 As String, _
    str项目名称 As String, str费用大类 As String, str项目类别 As String, _
    str是否医保 As String, str是否限价 As String, dbl自付比例 As Double, _
    dbl标准单价 As Double) As Long

'===============================================================================================================
'功能: 取诊疗字典信息
'入口参数: 项目编码(10)
'出口参数: 项目名称(40位),费用大类(2位),项目类别(1位),是否医保(1位),是否限价(1位),自付比例(4位宽3位小数)
'          标准单价(8位，2位小数)
'说明: 项目类别:0-甲或普通,1-乙或高精尖,2-自费
'      是否医保:0-不是,1-是
'      是否限价:0-不是,1-是
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function GetInfo_ItemDic Lib "YhYbClient.dll" (ByVal str项目编码 As String, _
    str项目名称 As String, str费用大类 As String, str项目类别 As String, _
    str是否医保 As String, str是否限价 As String, dbl自付比例 As Double, _
    ByRef dbl标准单价 As Double) As Long

'===============================================================================================================
'功能: 取服务设施字典信息
'入口参数: 项目编码(10)
'出口参数: 项目名称(40位),费用大类(2位),项目类别(1位),是否医保(1位),是否限价(1位),自付比例(4位宽3位小数)
'          标准单价(8位，2位小数)
'说明: 项目类别:0-甲或普通,1-乙或高精尖,2-自费
'      是否医保:0-不是,1-是
'      是否限价:0-不是,1-是
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function GetInfo_ServerDic Lib "YhYbClient.dll" (ByVal str项目编码 As String, _
    str项目名称 As String, str费用大类 As String, str项目类别 As String, _
    str是否医保 As String, str是否限价 As String, dbl自付比例 As Double, _
    dbl标准单价 As Double) As Long

'===============================================================================================================
'功能: 取病种信息
'入口参数: 病种编码(10)
'出口参数: 病种主码(10位),病种附码(10位),病种名称(40位),病种统计码(3位),注释(200位),病种标志(1位),病种类别(2位)
'          科别(3位)
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function GetInfo_SickDic Lib "YhYbClient.dll" (ByVal str病种编码 As String, _
     str病种主码 As String, str病种附码 As String, str病种名称 As String, _
    str病种统计码 As String, str注释 As String, str病种标志 As String, _
    str病种类别 As String, str科别 As String) As Long

'===============================================================================================================
'功能: 取费用大类信息
'入口参数: 费用大类编码(2位)
'出口参数: 费用大类名称(20位)
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function GetInfo_MediKind Lib "YhYbClient.dll" (ByVal str大类编码 As String, _
     str大类名称 As String) As Long

'===============================================================================================================
'功能: 读卡
'入口参数: 无
'出口参数: 中心代码(4位),卡号(16位),身份证号(18位),姓名(10位),性别(1位),出生日期(8位),医保号(8位)
'          用人单位编码(5位),个人身份,公务员标志(1位),是否参加补充保险(1位),参加大病医保(1位),隶属关系(1位)
'          照顾级别(1位),职工属地(1位),是否慢性病(1位),是否重大疾病(1位),住院标志(1位)
'          起付段以上医疗费累计(8位,2位小数),本年统筹支付累计(8位,2位小数),年起付线金额累计(8位,2位小数)
'          慢性病本年统筹支付累计(8位,2位小数),重大疾病本年统筹支付累计(8位,2位小数),帐户余额(8位,2位小数)
'          本年有效住院次数(3位)
'说明: 性别:0-女,1-男
'      出生日期:格式yyyymmdd
'      个人身份:0-在职,1-退休
'      公务员标志:0-非公务员,1-公务员,其他参照公务员
'      是否参加补充保险:0-不参加,1-参加
'      参加大病医保:0-不参加,1-参加
'      隶属关系:1-市属财政,0-其他
'      照顾级别:0-无,1-一级,2-二级,3-三级
'      职工属地:0-本地,1-常驻外地,2-异地安置
'      是否慢性病:0-不是,1-是
'      是否重大疾病:0-不是,1-是
'      住院标志:0-不住院,1-住院
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function ReadCard Lib "YhYbClient.dll" (str中心代码 As String, str卡号 As String, _
     str身份证号 As String, STR姓名 As String, str性别 As String, _
     str出生日期 As String, str医保帐号 As String, str用人单位编码 As String, _
     str个人身份 As String, str公务员标志 As String, str补充保险 As String, _
     str大病医保 As String, str隶属关系 As String, str是否慢性病 As String, str是否大病 As String, str照顾级别 As String, _
     str职工属地 As String, _
     str住院标志 As String, dbl进入统筹累计 As Double, dbl本年统筹累计 As Double, _
     dbl年起付线累计 As Double, dbl特病本年累计 As Double, dbl大病本年累计 As Double, _
     dbl帐户余额 As Double, int本年住院次数 As Long, int支付序列号 As Long) As Long            '文档中声明与说明不一致
    
    

'===============================================================================================================
'功能: 设置门诊收据
'入口参数: 保留,收据号(13位),门诊计算类型(1位),医保病种代码(10位),诊断说明(200位),科室名称(20位),医生名称(10位)
'          药师(10位),中药付数(2位),金额(8位,2位小数)
'说明: 收据号:Not Null
'      门诊计算类型:Not Null,1-普通,2-慢性病,3-重大疾病,4-照顾对象,5-特种人,6-计划生育,7-工伤
'      金额:>0
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetClinicBill Lib "YhYbClient.dll" (ByVal lng保留 As Long, ByVal str收据号 As String, ByVal str门诊计算类型 As String, ByVal str医保病种代码 As String, _
    ByVal str诊断说明 As String, ByVal str科室名称 As String, ByVal str医生名称 As String, _
    ByVal str药师 As String, ByVal dbl中药付数 As Double, ByVal dbl金额 As Double) As Long

'===============================================================================================================
'功能: 设置门诊收据明细
'入口参数: 保留,医保项目编号(10位),医院项目名称(40位),剂型名称(20位),单位含量(14位),用法用量(40位)
'          费用大类代码(2位),费用类别(1位),是否医保(1位),是否药品(1位),单价(8位,2位小数),数量(8位,2位小数)
'          金额(8位,2位小数)
'说明: 医保项目编号:Not Null
'      医院项目名称:Not Null
'      费用大类代码:Not Null
'      费用类别:0-甲或普通,1-乙或高精尖,2-自费,Not Null
'      是否医保:0-不是,1-是[为与旧接口兼容而保留未用]
'      是否药品:0-项目,1-药品,2-服务设施[床],Not Null
'      单价,数量,金额:>0
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetClinicBillDetail Lib "YhYbClient.dll" (ByVal lng保留 As Long, _
    ByVal str医保项目编号 As String, ByVal str医院项目名称 As String, ByVal str剂型名称 As String, _
    ByVal str单位含量 As String, ByVal str用法用量 As String, ByVal str费用大类代码 As String, _
    ByVal str费用类别 As String, ByVal str是否医保 As String, ByVal str是否药品 As String, _
    ByVal dbl单价 As Double, ByVal dbl数量 As Double, ByVal dbl金额 As Double) As Long

'===============================================================================================================
'功能: 设置门诊大类信息
'入口参数: 保留,医保费用大类代码(2位),相应费用大类金额(8位,2位小数)
'说明: 医保费用大类代码:Not Null
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetClinicMediKind Lib "YhYbClient.dll" (ByVal lng保留 As Long, _
    ByVal str大类代码 As String, ByVal dbl大类金额 As Double) As Long

'===============================================================================================================
'功能: 入院登记
'入口参数: 个人句柄(由ReadCard服务调用返回),住院号(13位),入院日期(8位),住院科室(20位),病区(20位),房间(10位)
'          床号(3位),门诊医生(10位),入院诊断代码(10位),str诊断名称(200)
'说明: 住院号:Not Null
'      入院日期:格式yyyymmdd,Not Null
'      入院诊断代码:病种代码,Not Null
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function InhosRegister Lib "YhYbClient.dll" (ByVal int个人句柄 As Integer, _
    ByVal str住院号 As String, ByVal str入院日期 As String, ByVal str住院科室 As String, _
    ByVal str病区 As String, ByVal str房间 As String, ByVal str床号 As String, ByVal str门诊医生 As String, _
    ByVal str入院诊断代码 As String, ByVal str诊断名称 As String) As Long

'===============================================================================================================
'功能: 出院登记
'入口参数: 住院号(13位),出院日期(8位),住院科室(20位),病区(20位),房间(10位),床号(3位),主治医生(10位)
'          出院诊断代码(10位),住院天数(3位),出院治疗情况(1位)
'说明: 住院号:Not Null
'      出院日期:格式yyyymmdd,Not Null
'      出院诊断代码:病种代码,Not Null
'      出院治疗情况:1-治愈,2-好转,3-未愈,4-死亡,9-其他,Not Null
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function OuthosRegister Lib "YhYbClient.dll" (ByVal str住院号 As String, _
    ByVal str出院日期 As String, ByVal str住院科室 As String, ByVal str病区 As String, _
    ByVal str房间 As String, ByVal str床号 As String, ByVal str医生 As String, _
    ByVal str出院诊断代码 As String, ByVal str出院治疗情况 As String, ByVal lng住院天数 As Long, ByVal str保留 As String) As Long
    '文档中声明与说明的参数不一致

'===============================================================================================================
'功能: 设置医嘱
'入口参数: 住院号(13位),医嘱号(13位),开嘱医生姓名(10位),停嘱医生姓名(10位),执行姓名(10位),录入人(10位)
'          是否长期医嘱(1位),医嘱开始日期(8位),医嘱开始时间(8位),医嘱停止日期(8位),医嘱停止时间(8位)
'          执行日期(8位),执行时间(8位),录入日期(8位),医嘱描述(200位)
'说明: 住院号:Not Null
'      医嘱号:Not Null
'      是否长期医嘱:0-不是,1-是
'      医嘱开始日期,医嘱停止日期,执行日期,录入日期:格式yyyymmdd,not null
'      医嘱开始时间,医嘱停止时间,执行时间:格式hh:mi:ss,not null
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetDoctorAdvice Lib "YhYbClient.dll" (ByVal int住院号 As Integer, _
    ByVal str医嘱号 As String, ByVal str开嘱医生姓名 As String, ByVal str停嘱医生姓名 As String, _
    ByVal str执行姓名 As String, ByVal str录入人 As String, ByVal str是否长期医嘱 As String, _
    ByVal str医嘱开始日期 As String, ByVal str医嘱开始时间 As String, ByVal str医嘱停止日期 As String, _
    ByVal str医嘱停止时间 As String, ByVal str执行日期 As String, ByVal str执行时间 As String, _
    ByVal str录入日期 As String, ByVal str医嘱描述 As String) As Long

'===============================================================================================================
'功能: 设置住院帐单
'入口参数: 住院号(13),住院帐单号(13位),病种代码(10位),科室(20位),医生(10位),中草药付数(2位),金额(8位,2位小数)
'说明: 住院号:Not Null
'      住院帐单号:Not Null
'      病种代码:Not Null
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetInHosBill Lib "YhYbClient.dll" (ByVal str住院登记号 As String, _
    ByVal str帐单号 As String, ByVal str病种代码 As String, ByVal str科室 As String, ByVal str医生 As String, _
    ByVal dbl中草药付数 As Double, ByVal dbl金额 As Double) As Long

'===============================================================================================================
'功能: 设置住院帐单明细
'入口参数: 保留,医保项目编号(10位),医院项目名称(40位),剂型名称(20位),单位含量(14位),用法用量(40位)
'          费用大类代码(2位),费用类别(1位),是否医保(1位),是否药品(1位),单价(8位,2位小数),数量(8位,2位小数)
'          金额(8位,2位小数)
'说明: 医保项目编号:Not Null
'      医院项目名称:Not Null
'      费用大类代码:Not Null
'      费用类别:0-甲或普通,1-乙或高精尖,2-自费,Not Null
'      是否医保:0-不是,1-是[为与旧接口兼容而保留未用]
'      是否药品:0-项目,1-药品,2-服务设施[床],Not Null
'      单价,数量,金额:>0
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetInHosBillDetail Lib "YhYbClient.dll" (ByVal int保留 As Integer, _
    ByVal str医保项目编号 As String, ByVal str医院项目名称 As String, ByVal str剂型名称 As String, _
    ByVal str单位含量 As String, ByVal str用法用量 As String, ByVal str费用大类代码 As String, _
     ByVal str是否医保 As String, ByVal str费用类别 As String, ByVal str是否药品 As String, _
    ByVal dbl单价 As Double, ByVal dbl数量 As Double, ByVal dbl金额 As Double) As Long

'===============================================================================================================
'功能: 设置结算单
'入口参数: 保留,结算单号(13位),费用总额(8位,2位小数),入院日期(8位),出院日期(8位),科室(20位),医生(10位)
'          出院病种编码(10位),并发症说明(200位),出院情况(1位),住院天数(3位)
'说明: 结算单号:Not Null
'      入院日期,出院日期:格式yyyymmdd,Not Null
'      出院情况:1-治愈,2-好转,3-未愈,4-死亡,9-其他,Not Null
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetCheckOutBill Lib "YhYbClient.dll" (ByVal lng保留 As Long, _
    ByVal str结算单号 As String, ByVal str入院日期 As String, _
    ByVal str出院日期 As String, ByVal str科室 As String, ByVal str医生 As String, _
    ByVal str病种编码 As String, ByVal str并发症 As String, ByVal STR出院情况 As String, _
    ByVal int住院天数 As Long, ByVal dbl费用总额 As Double) As Long

'===============================================================================================================
'功能: 设置结算单明细
'入口参数: 保留,医保项目编号(10位),医院项目名称(40位),剂型名称(20位),单位含量(14位),用法用量(40位)
'          费用大类代码(2位),费用类别(1位),是否医保(1位),是否药品(1位),单价(8位,2位小数),数量(8位,2位小数)
'          金额(8位,2位小数)
'说明: 医保项目编号:Not Null
'      医院项目名称:Not Null
'      费用大类代码:Not Null
'      费用类别:0-甲或普通,1-乙或高精尖,2-自费,Not Null
'      是否医保:0-不是,1-是[为与旧接口兼容而保留未用]
'      是否药品:0-项目,1-药品,2-服务设施[床],Not Null
'      单价,数量,金额:>0
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetCheckOutBillDetailX Lib "YhYbClient.dll" (ByVal lng保留 As Long, _
    ByVal str医保项目编号 As String, ByVal str医院项目名称 As String, ByVal str剂型名称 As String, _
    ByVal str单位含量 As String, ByVal str用法用量 As String, ByVal str费用大类代码 As String, _
    ByVal str费用类别 As String, ByVal str是否医保 As String, ByVal str是否药品 As String, _
    ByVal dbl单价 As Double, ByVal dbl数量 As Double, ByVal dbl金额 As Double) As Long

'===============================================================================================================
'功能: 设置住院大类信息
'入口参数: 保留,医保费用大类代码(2位),相应费用大类金额(8位,2位小数)
'说明: 医保费用大类代码:Not Null
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function SetInHosMediKind Lib "YhYbClient.dll" (ByVal lng保留 As Long, ByVal str大类代码 As String, ByVal dbl大类金额 As Double) As Long

'===============================================================================================================
'功能: 取消入院登记
'入口参数: 保留
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function uInhosRegister Lib "YhYbClient.dll" (ByVal int保留 As Integer) As Long

'===============================================================================================================
'功能: 取消出院登记
'入口参数: 住院号(13位)
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function uOuthosRegister Lib "YhYbClient.dll" (ByVal str住院号 As String) As Long

'===============================================================================================================
'功能: 门诊、药店退费,住院退费, 住院取消结算
'入口参数: 保留,退费新单据号(13位),要退费的单据号(13位)
'说明: 退费新单据号,要退费的单据号:Not Null
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function ReturnCharge Lib "YhYbClient.dll" (ByVal int保留 As Integer, ByVal str新单号 As String, _
    ByVal str原单号 As String) As Long

'===============================================================================================================
'功能: 取消医嘱录入
'入口参数: 医嘱号(13位)
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function uDoctorAdvice Lib "YhYbClient.dll" (ByVal str医嘱号 As String) As Long

'===============================================================================================================
'功能: 取得门诊收据计算信息(药店)
'入口参数: 保留
'出口参数: 帐户余额(8位,2位小数),个人帐户支付(8位,2位小数),个人现金支付(8位,2位小数),个人比例负担(8位,2位小数)
'          统筹支付(8位,2位小数),照顾支付(8位,2位小数),照顾垫付(8位,2位小数),自付段支付(8位,2位小数)
'          商保支付(8位,2位小数),甲类药品(8位,2位小数),自费药品(8位,2位小数),乙类药品(8位,2位小数)
'          甲类诊疗(普通)(8位,2位小数),自费诊疗(8位,2位小数),乙类诊疗(高精尖)(8位,2位小数),甲类设施(8位,2位小数)
'          自费设施(8位,2位小数),乙类设施(8位,2位小数),其他自费(8位,2位小数),自付段累计(8位,2位小数)
'          统筹支付累计(8位,2位小数),重病支付累计(8位,2位小数),慢病支付累计(8位,2位小数)
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function GetClinicBillData Lib "YhYbClient.dll" (ByVal lng保留 As Long, _
    ByRef dbl帐户余额 As Double, ByRef dbl个人帐户支付 As Double, ByRef dbl个人现金支付 As Double, _
    ByRef dbl个人比例负担 As Double, ByRef dbl统筹支付 As Double, ByRef dbl照顾支付 As Double, _
    ByRef dbl照顾垫付 As Double, ByRef dbl自付段支付 As Double, ByRef dbl商保支付 As Double, _
    ByRef dbl甲类药品 As Double, ByRef dbl自费药品 As Double, ByRef dbl乙类药品 As Double, _
    ByRef dbl甲类诊疗 As Double, ByRef dbl自费诊疗 As Double, ByRef dbl乙类诊疗 As Double, _
    ByRef dbl甲类设施 As Double, ByRef dbl自费设施 As Double, ByRef dbl乙类设施 As Double, _
    ByRef dbl其他自费 As Double, ByRef dbl自付段累计 As Double, ByRef dbl统筹支付累计 As Double, _
    ByRef dbl重病支付累计 As Double, ByRef dbl慢病支付累计 As Double, ByRef dbl非基本医疗费 As Double) As Long

'===============================================================================================================
'功能: 取得住院结算计算信息
'入口参数: 保留
'出口参数: 帐户余额(8位,2位小数),个人帐户支付(8位,2位小数),个人现金支付(8位,2位小数),个人比例负担(8位,2位小数)
'          统筹支付(8位,2位小数),照顾支付(8位,2位小数),照顾垫付(8位,2位小数),自付段支付(8位,2位小数)
'          商保支付(8位,2位小数),甲类药品(8位,2位小数),自费药品(8位,2位小数),乙类药品(8位,2位小数)
'          甲类诊疗(普通)(8位,2位小数),自费诊疗(8位,2位小数),乙类诊疗(高精尖)(8位,2位小数),甲类设施(8位,2位小数)
'          自费设施(8位,2位小数),乙类设施(8位,2位小数),其他自费(8位,2位小数),自付段累计(8位,2位小数)
'          统筹支付累计(8位,2位小数),重病支付累计(8位,2位小数),慢病支付累计(8位,2位小数)
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function GetCheckOutBillData Lib "YhYbClient.dll" (ByVal lng保留 As Long, _
    ByRef dbl帐户余额 As Double, ByRef dbl个人帐户支付 As Double, ByRef dbl个人现金支付 As Double, _
    ByRef dbl个人比例负担 As Double, ByRef dbl统筹支付 As Double, ByRef dbl照顾支付 As Double, _
    ByRef dbl照顾垫付 As Double, ByRef dbl自付段支付 As Double, ByRef dbl商保支付 As Double, _
    ByRef dbl甲类药品 As Double, ByRef dbl自费药品 As Double, ByRef dbl乙类药品 As Double, _
    ByRef dbl甲类诊疗 As Double, ByRef dbl自费诊疗 As Double, ByRef dbl乙类诊疗 As Double, _
    ByRef dbl甲类设施 As Double, ByRef dbl自费设施 As Double, ByRef dbl乙类设施 As Double, _
    ByRef dbl其他自费 As Double, ByRef dbl自付段累计 As Double, ByRef dbl统筹支付累计 As Double, _
    ByRef dbl重病支付累计 As Double, ByRef dbl慢病支付累计 As Double, ByRef dbl非基本医疗费 As Double) As Long

'===============================================================================================================
'功能: 政策服务提交数据(门诊收费、住院结算用)
'入口参数: 个人帐户支付(8位,2位小数),现金支付(8位,2位小数)
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function CommitDataX Lib "YhYbClient.dll" (ByVal dbl人帐支付 As Double, _
    ByVal dbl现金支付 As Double) As Long

'===============================================================================================================
'功能: 其他服务提交数据
'入口参数: 无
'出口参数: 无
'返回: 0成功,-1失败
'===============================================================================================================
Public Declare Function CommitData Lib "YhYbClient.dll" () As Long

Private Function Get交易代码(ByVal intType As 业务类型_吉林, Optional bln读名称 As Boolean = False) As String
    Select Case intType
        Case 初始化服务调用
            Get交易代码 = IIf(bln读名称, "初始化服务调用", "01")
        Case 读卡
            Get交易代码 = IIf(bln读名称, "读卡", "02")
        Case 取药品信息
            Get交易代码 = IIf(bln读名称, "取药品信息", "03")
        Case 取服务信息
            Get交易代码 = IIf(bln读名称, "取服务信息", "04")
        Case 取诊疗信息
            Get交易代码 = IIf(bln读名称, "取诊疗信息", "05")
        Case 设置发票数据
            Get交易代码 = IIf(bln读名称, "设置发票数据", "06")
        Case 设置门诊收据明细
            Get交易代码 = IIf(bln读名称, "设置门诊收据明细", "07")
        Case 设置门诊大类信息
            Get交易代码 = IIf(bln读名称, "设置门诊大类信息", "08")
        Case 取得门诊收据计算信息
            Get交易代码 = IIf(bln读名称, "取得门诊收据计算信息", "09")
        Case 取消结算
            Get交易代码 = IIf(bln读名称, "取消结算", "10")
        Case 入院登记
            Get交易代码 = IIf(bln读名称, "入院登记", "11")
        Case 取消入院登记
            Get交易代码 = IIf(bln读名称, "取消入院登记", "12")
        Case 出院登记
            Get交易代码 = IIf(bln读名称, "出院登记", "13")
        Case 取消出院登记
            Get交易代码 = IIf(bln读名称, "取消出院登记", "14")
        Case 设置住院帐单
            Get交易代码 = IIf(bln读名称, "设置住院帐单", "15")
        Case 设置记帐单明细数据
            Get交易代码 = IIf(bln读名称, "设置记帐单明细数据", "16")
        Case 设置结算单
            Get交易代码 = IIf(bln读名称, "设置结算单", "17")
        Case 设置住院大类信息
            Get交易代码 = IIf(bln读名称, "设置住院大类信息", "18")
        Case 取得住院结算计算信息
            Get交易代码 = IIf(bln读名称, "取得住院结算计算信息", "19")
        Case 服务数据提交
            Get交易代码 = IIf(bln读名称, "服务数据提交", "20")
        Case 政策服务提交
            Get交易代码 = IIf(bln读名称, "政策服务提交", "21")
        Case 结束服务调用
            Get交易代码 = IIf(bln读名称, "结束服务调用", "22")
        Case Else
            Get交易代码 = IIf(bln读名称, "错误的交易代码", "-1")
    End Select
End Function

Public Function CheckReturn吉林() As Boolean
    CheckReturn吉林 = True
    If glngReturn = -1 Then
        MsgBox "在进行医保调用时，医保返回以下错误：" & vbCrLf & "    " & GetErrMsg(), vbInformation, "接口错误"
        CheckReturn吉林 = False
    End If
End Function

Public Sub delArrar()
    Dim iLoop As Long
    For iLoop = 0 To 50
        strTempArr(iLoop, 0) = ""
        strTempArr(iLoop, 1) = "0"
    Next
End Sub

Public Sub setArrar(str大类 As String, dbl费用 As Double)
    Dim iLoop As Long
    For iLoop = 0 To 50
        If strTempArr(iLoop, 0) = str大类 Then
            strTempArr(iLoop, 1) = CLng(strTempArr(iLoop, 1)) + dbl费用
            Exit Sub
        ElseIf strTempArr(iLoop, 0) = "" Then
            strTempArr(iLoop, 0) = str大类
            strTempArr(iLoop, 1) = dbl费用
            Exit Sub
        End If
    Next
End Sub

Public Function 医保初始化_吉林() As Boolean
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    If mblnInit = True Then
        医保初始化_吉林 = True
        Exit Function
    End If
    DebugTool "进入医保初始化接口"
    
    '刘兴宏:20040923加入
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_吉林.模拟数据 = True
    Else
        InitInfor_吉林.模拟数据 = False
    End If
    
    
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_吉林
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取参数"
    
    InitInfor_吉林.明细时实上传 = False
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "明细时实上传"
                InitInfor_吉林.明细时实上传 = Nvl(!参数值, 1) = 1
            End Select
            .MoveNext
        Loop
    End With
    
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_吉林)
    InitInfor_吉林.医院编码 = Nvl(rsTemp!医院编码)
    mblnInit = True
    医保初始化_吉林 = True
    DebugTool "医保初始化接口成功"
End Function

Public Function 身份标识_吉林(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Err = 0
    On Error GoTo errHand:
    身份标识_吉林 = frmIdentify吉林.GetPatient(bytType, lng病人ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_吉林 = ""
End Function
Private Function IS是否刷卡病人(ByVal lng病人ID As Long) As Boolean
    '判断当前的病人是否刷卡的病人
    Dim rsTemp As New ADODB.Recordset
    IS是否刷卡病人 = False
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select * from 保险帐户 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医保病人信息"
    If rsTemp.EOF Then
        ShowMsgbox "不存在当前的医保病人"
        Exit Function
    End If
    If g病人身份_吉林.医保号 <> Trim(Nvl(rsTemp!医保号)) Then
        ShowMsgbox "卡有错误,不是当前病人的.请确定插入的卡是否存确!"
        Exit Function
    End If
    IS是否刷卡病人 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 个人余额_吉林(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_吉林)
    
    If rsTemp.EOF Then
        个人余额_吉林 = 0
    Else
        个人余额_吉林 = rsTemp("帐户余额")
    End If
End Function

Public Function 门诊结算_吉林(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency, Optional ByRef strAdvance As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strArr
    Dim dbl中药付数 As Double, dbl金额 As Double, str医保病种代码 As String, str是否药品 As String, str结算方式 As String
    Dim StrInput As String, strOutput As String
    Dim iLoop As Integer
    Dim blnOld As Boolean '是否需要填写校正字段
    On Error GoTo errHandle
    
    gstrSQL = " " & _
        "  Select Rownum 标识号,A.ID,A.病人ID,A.收费细目id,A.NO,A.序号,A.操作员姓名,A.记录性质,A.记录状态,A.登记时间,A.开单人 as 医生,H.编号 as 医生编号, " & _
        "      A.付数,A.数次*A.付数 as 数量,A.是否上传,A.计算单位,B.规格,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额 as 实收金额, " & _
        "      A.医嘱序号,A.收费类别,B.编码 as 项目编码,B.名称 as 项目名称,decode(J.标识码,null,B.标识主码||B.标识子码,nvl(J.标识码,' ')) as 国家编码, " & _
        "      D.项目编码 医保编码,D.项目名称 as 医保名称,J.名称 as 剂型,D.是否医保,C.名称 开单部门,E.名称 受单部门, " & _
        "      L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.单位编码,L.顺序号,L.退休证号,L.帐户余额,L.当前状态,L.病种ID,L.在职,L.年龄段,L.灰度级,L.就诊时间 " & _
        "  From (Select * From 门诊费用记录 Where nvl(实收金额,0)<>0 and  记录状态<>0 and 结帐ID=[2] and  Nvl(附加标志,0)<>9 ) A,收费细目 B,部门表 C,保险支付项目 D,部门表 E,  " & _
        "       (Select distinct Q.药品id,Q.标识码,T.名称 From 药品目录 Q,药品信息 R,药品剂型 T  Where  Q.药名id=R.药名id and R.剂型=T.编码 ) J, " & _
        "       人员表 H,保险帐户 L" & _
        "  Where A.收费细目ID=B.ID And A.开单部门ID=C.ID(+)  and A.病人id=L.病人id  and L.险类=[1] and a.收费细目id=J.药品id(+) " & _
        "        and A.执行部门ID=E.ID(+) And A.收费细目ID=D.收费细目ID And D.险类=[1] and a.开单人=H.姓名(+) " & _
        "  Order by A.ID"
                        
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_吉林, lng结帐ID)
    
    If rs明细.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "没有填写收费记录"
        Exit Function
    End If
    dbl中药付数 = 0
    dbl金额 = 0
    While Not rs明细.EOF
        dbl金额 = dbl金额 + rs明细!实收金额
        If Nvl(rs明细!收费类别) = "6" Or Nvl(rs明细!收费类别) = "7" Then
            dbl中药付数 = dbl中药付数 + Nvl(rs明细!付数, 0)
        End If
        rs明细.MoveNext
    Wend
    
    rs明细.MoveFirst
    If dbl金额 = 0 Then
        Err.Raise 9000, gstrSysName, "病人没有发生费用,不能进行医保处理"
        Exit Function
    End If
    
    '取病人ID和操作员
    lng病人ID = rs明细!病人ID
    
    gstrSQL = "Select * From 保险病种 Where ID=" & Nvl(rs明细!病种ID, 0) & " And 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_吉林)
        
    '刘兴宏:20040923可以没有病种
    If rsTemp.EOF Then
        str医保病种代码 = ""
    Else
        str医保病种代码 = Substr(rsTemp!编码, 1, 10)
    End If
    
    'If 业务请求_吉林(初始化服务调用, "10", strOutPut) = False Then Exit Function
    '刘兴宏:需重新刷一次卡
    If 身份鉴别_吉林(0, "10") = False Then
        Exit Function
    End If
    If IS是否刷卡病人(lng病人ID) = False Then
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
        Exit Function
    End If

    
    '设置发票数据
    StrInput = 1
    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!NO), 1, 13)
    StrInput = StrInput & vbTab & g病人身份_吉林.门诊类型
    '陈宏悦于20060512修改
    StrInput = StrInput & vbTab & Substr(g病人身份_吉林.病种代码, 1, 10)
    StrInput = StrInput & vbTab & Substr(g病人身份_吉林.诊断名称, 1, 200)
    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!开单部门), 1, 20)
    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!医生), 1, 10)
    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!操作员姓名), 1, 10)
    StrInput = StrInput & vbTab & Substr(dbl中药付数, 1, 2)
    StrInput = StrInput & vbTab & dbl金额
    If 业务请求_吉林(设置发票数据, StrInput, strOutput) = False Then Exit Function
    
    delArrar            '清除费用大类记录数组
    
    '设置发票明细数据：SetClinicBillDetail
    Do While Not rs明细.EOF
        If Nvl(rs明细!医保编码, "") = "" Then
            Err.Raise 9000, gstrSysName, "项目[" & Nvl(rs明细!项目名称) & "]未设置对应的医保项目,不能用于医保"
            Exit Function
        End If
        
        If Nvl(rs明细!是否上传, 0) = 0 And rs明细!实收金额 <> 0 Then
            str是否药品 = "0"
            StrInput = Substr(Nvl(rs明细!医保编码), 1, 10)
            
            Select Case UCase(Nvl(rs明细!收费类别))
                Case "5", "6", "7"
                    str是否药品 = "1"
                    'aItemName：项目名称(40位)
                    'aMediKindCode：费用大类(2位)
                    'aIsCityRich：(0-甲或普通1-乙或高精尖)(1位)[新增2自费]
                    'aIsCityMedi：是否医保(1位)(0-不是1-是)[为与旧接口兼容而保留]
                    'aIsLimit：是否医保限价项目(1位)(0-不是1-是)
                    'aCitySelfPayRate：自付比例(4位宽3位小数)
                    'aPrice：标准单价(8位，2位小数)
                    
                    If 业务请求_吉林(取药品信息, StrInput, strOutput) = False Then Exit Function
                Case "J", "H", "I"
                    If 业务请求_吉林(取服务信息, StrInput, strOutput) = False Then Exit Function
                    str是否药品 = "2"
                Case Else
                    If 业务请求_吉林(取诊疗信息, StrInput, strOutput) = False Then Exit Function
            End Select
            If strOutput = "" Then Exit Function
            strArr = Split(strOutput, vbTab)
            
            '记录费用大类金额
            setArrar CStr(strArr(1)), Nvl(rs明细!实收金额, 0)
            
            '调用接口,写入明细
            'aInvoiceHandle: [为与旧接口兼容而保留未用]
            StrInput = "1"
            'aCityMediCareNo：医保项目编号。(10位)not null
            StrInput = StrInput & vbTab & Nvl(rs明细!医保编码, "")
            'aItemName：医院项目名称(40位)not null
            StrInput = StrInput & vbTab & Nvl(rs明细!项目名称, "")
            'aConformationName：剂型名称(20位)
            StrInput = StrInput & vbTab & Nvl(rs明细!剂型, "")
            'aUnitContent：单位含量(14位)
            StrInput = StrInput & vbTab & Substr(Nvl(rs明细!规格, ""), 1, 14)
                '刘兴宏:暂且屏蔽
                'gstrSQL = "Select 单量,频次,用法 From 药品收发记录 where 费用id=" & Nvl(!ID, 0)
                'zlDataBase.OpenRecordset rsTemp, gstrSQL, "获取病人单理及频次"
                'If Not rsTemp.EOF Then
                'strInput = strInput & vbTab & Substr("单量:" & Nvl(rsTemp!单量, "") & " 频次:" & Nvl(rsTemp!频次) & "用法:" & Nvl(rsTemp!用法), 1, 14)
                'Else
                'strInput = strInput & vbTab & ""
                'End If
            'aDosage：用法用量(40位)
            StrInput = StrInput & vbTab & ""
            'aMediKindCode：费用大类代码(2位)not null
            StrInput = StrInput & vbTab & strArr(1)
            'aIsRich：(0-甲或普通1-乙或高精尖)(1位)[新增2自费]not null
            StrInput = StrInput & vbTab & strArr(2)
            'aIsCityMedi：是否医保(1位)(0-不是1-是)[为与旧接口兼容而保留未用]
            StrInput = StrInput & vbTab & strArr(3)
            'aIsMedi:是否药品(0-项目1-药品2服务设施[床] )not null
            StrInput = StrInput & vbTab & str是否药品
            'aPrice：单价(8位，2位小数)>0
            StrInput = StrInput & vbTab & Nvl(rs明细!实际价格, 0)
            'aQuantity：数量(8位，2位小数)>0
            StrInput = StrInput & vbTab & Nvl(rs明细!数量, 0)
            'aAmount：金额(8位，2位小数)>0
            StrInput = StrInput & vbTab & Nvl(rs明细!实收金额, 0)
            
            If 业务请求_吉林(设置门诊收据明细, StrInput, strOutput) = False Then Exit Function
            gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细!ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        End If
        rs明细.MoveNext
    Loop
    
    For iLoop = 0 To 50
        If strTempArr(iLoop, 0) <> "" Then
            StrInput = "1"
            StrInput = StrInput & vbTab & strTempArr(iLoop, 0)
            StrInput = StrInput & vbTab & strTempArr(iLoop, 1)
            If 业务请求_吉林(设置门诊大类信息, StrInput, strOutput) = False Then Exit Function
        Else
            Exit For
        End If
    Next
        
    If 业务请求_吉林(取得门诊收据计算信息, "1", strOutput) = False Then Exit Function
    strArr = Split(strOutput, vbTab)
    
    
    If Val(strArr(1)) <> 0 Then
        str结算方式 = str结算方式 & "||个人帐户|" & Format(Val(strArr(1)), "####0.00;-####0.00; ;")
    End If
    
    If Val(strArr(4)) <> 0 Then
        str结算方式 = str结算方式 & "||统筹支付|" & Format(Val(strArr(4)), "####0.00;-####0.00; ;")
    End If
    
    '如果存在
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        #If gverControl < 2 Then
            blnOld = True
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
        #Else
            strAdvance = str结算方式
            gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    End If
    
    If blnOld Then
        If frm结算信息.ShowME(lng结帐ID, True) = False Then
            Exit Function
        End If
    End If

   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(自付段累计),帐户累计支出_IN(统筹支付累计),累计进入统筹_IN(重病支付累计),累计统筹报销_IN(慢病支付累计),住院次数_IN,起付线(非基本医疗费),封顶线_IN(个人现金支付),实际起付线_IN(个人比例负担),
    '   发生费用金额_IN(费用总额),全自付金额_IN(自付段支付),首先自付金额_IN(自费药品),
    '   进入统筹金额_IN(甲类药品),统筹报销金额_IN(统筹支付),    大病自付金额_IN(照顾支付),超限自付金额_IN(照顾垫付),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(交易类型),主页ID_IN,中途结帐_IN,备注_IN
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    
    With g病人身份_吉林
        gstrSQL = "zl_保险结算记录_insert( 1," & lng结帐ID & "," & TYPE_吉林 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          Val(strArr(19)) & "," & Val(strArr(20)) & "," & Val(strArr(21)) & "," & Val(strArr(22)) & ",NULL," & Val(strArr(23)) & "," & Val(strArr(2)) & "," & Val(strArr(3)) & "," & _
         dbl金额 & "," & Val(strArr(7)) & "," & Val(strArr(10)) & "," & _
          Val(strArr(9)) & "," & Val(strArr(4)) & "," & Val(strArr(5)) & "," & Val(strArr(6)) & "," & Val(strArr(1)) & ",'" & _
          .门诊类型 & "',Null,Null,NULl" & IIf(blnOld, "", ",1") & ")"
              
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        '附加信息参数:
            '性质_IN ,记录ID_IN,帐户余额_IN ,商保支付_IN ,乙类药品_IN ,甲类诊疗_IN ,自费诊疗_IN ,乙类诊疗_IN ,甲类设施_IN ,自费设施_IN ,乙类设施_IN ,其他自费_IN
        gstrSQL = "zl_保险结算记录_附加信息( 1," & lng结帐ID & "," & Val(strArr(0)) & "," & Val(strArr(8)) & "," & Val(strArr(11)) & "," & Val(strArr(12)) & "," & Val(strArr(13)) & "," & Val(strArr(14)) & "," & Val(strArr(15)) & "," & Val(strArr(16)) & "," & Val(strArr(17)) & "," & Val(strArr(18)) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录附加信息")
    End With

    StrInput = strArr(1)    '个人帐户支付
    StrInput = StrInput & vbTab & strArr(2) '现金支付
    
    If 业务请求_吉林(政策服务提交, StrInput, strOutput) = False Then Exit Function
    If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function


    门诊结算_吉林 = True
    DebugTool "门诊结算成功"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_吉林(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, rs明细 As New ADODB.Recordset
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency, int保留 As Integer, datCurr As Date
    Dim StrInput As String, strOutput As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select *  From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If rs明细.EOF Then
        Err.Raise 9000, gstrSysName, "没有找到原单据的明细数据，不能进行冲销", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do While Not rs明细.EOF
        cur票据总金额 = cur票据总金额 + Nvl(rs明细("结帐金额"), 0)
        rs明细.MoveNext
    Loop
    
    rs明细.MoveFirst
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
    

    '调用接口数冲销
    If 身份鉴别_吉林(0, "11") = False Then Exit Function
    If IS是否刷卡病人(lng病人ID) = False Then
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
        Exit Function
    End If
    
    StrInput = "1"
    StrInput = StrInput & vbTab & Nvl(rs明细!NO) & "R"
    StrInput = StrInput & vbTab & Nvl(rs明细!NO)
    If 业务请求_吉林(取消结算, StrInput, strOutput) = False Then Exit Function
    If 业务请求_吉林(服务数据提交, "", strOutput) = False Then Exit Function
    If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
    
    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(自付段累计),帐户累计支出_IN(统筹支付累计),累计进入统筹_IN(重病支付累计),累计统筹报销_IN(慢病支付累计),住院次数_IN,起付线(非基本医疗费),封顶线_IN(个人现金支付),实际起付线_IN(个人比例负担),
    '   发生费用金额_IN(费用总额),全自付金额_IN(自付段支付),首先自付金额_IN(自费药品),
    '   进入统筹金额_IN(甲类药品),统筹报销金额_IN(统筹支付),    大病自付金额_IN(照顾支付),超限自付金额_IN(照顾垫付),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(),主页ID_IN,中途结帐_IN,备注_IN
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    
    gstrSQL = "Select * From 保险结算记录 Where 记录ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID, TYPE_吉林)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "没有找到原有的保险结算记录", vbInformation, gstrSysName
        Exit Function
    End If
        
    With g病人身份_吉林
        gstrSQL = "zl_保险结算记录_insert( 1," & lng冲销ID & "," & TYPE_吉林 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          -1 * Nvl(rsTemp!帐户累计增加, 0) & "," & -1 * Nvl(rsTemp!帐户累计支出, 0) & "," & -1 * Nvl(rsTemp!累计进入统筹, 0) & "," & -1 * Nvl(rsTemp!累计统筹报销, 0) & ",NULL," & -1 * Nvl(rsTemp!起付线, 0) & "," & -1 * Nvl(rsTemp!封顶线, 0) & "," & -1 * Nvl(rsTemp!实际起付线, 0) & "," & _
         -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * Nvl(rsTemp!首先自付金额, 0) & "," & _
          -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & "," & -1 * Nvl(rsTemp!大病自付金额, 0) & "," & -1 * Nvl(rsTemp!超限自付金额, 0) & "," & -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & _
          rsTemp!支付顺序号 & "',Null,Null,NULl)"
              
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        '附加信息参数:
        '性质_IN ,记录ID_IN,帐户余额_IN ,商保支付_IN ,乙类药品_IN ,甲类诊疗_IN ,自费诊疗_IN ,乙类诊疗_IN ,甲类设施_IN ,自费设施_IN ,乙类设施_IN ,其他自费_IN
        gstrSQL = "zl_保险结算记录_附加信息( 1," & lng冲销ID & "," & -1 * Nvl(rsTemp!帐户余额, 0) & "," & -1 * Nvl(rsTemp!商保支付, 0) & "," & -1 * Nvl(rsTemp!乙类药品, 0) & "," & -1 * Nvl(rsTemp!甲类诊疗, 0) & "," & -1 * Nvl(rsTemp!自费诊疗, 0) & "," & -1 * Nvl(rsTemp!乙类诊疗, 0) & "," & -1 * Nvl(rsTemp!甲类设施, 0) & "," & -1 * Nvl(rsTemp!自费设施, 0) & "," & -1 * Nvl(rsTemp!乙类设施, 0) & "," & -1 * Nvl(rsTemp!其他自费, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录附加信息")
    End With

    门诊结算冲销_吉林 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function 入院登记_吉林(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
        
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用,请先进行结帐!"
        Exit Function
    End If
    
    
    On Error GoTo errHand:
    If 身份鉴别_吉林(1, "20") = False Then Exit Function
    
    '获取相关病人信息
    gstrSQL = "Select C.住院号,C.当前床号,L.名称 as 病区,A.当前病区id,to_char(A.确诊日期,'yyyyMMdd') as 确诊日期,A.门诊医师,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyyMMdd') 入院经办时间," & _
        " to_char(A.入院日期,'yyyyMMdd') 入院日期,to_char(A.入院日期,'ss') as 序号 ,to_char(A.登记时间,'yyyyMMdd') 入院时间,D.入院诊断 " & _
        " From 病案主页 A,部门表 B,部门表 L,病人信息 C, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码,'')) AS 入院诊断 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   D" & _
        " Where A.病人id=C.病人id and a.当前病区ID=L.iD(+) and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        ""
    If g病人身份_吉林.诊断名称 = "" Then
        ShowMsgbox "没有输入诊断情况,请在身份窗体中输入!"
        Exit Function
    End If
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取入院信息"
    
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    'aPrnHandle：个人句柄，由ReadCard服务调用返回。
    StrInput = "1"
    'aInHosNo：住院号(13位)not null
    StrInput = StrInput & vbTab & lng病人ID & "-" & lng主页ID & "-" & Nvl(rsTemp!序号)
    'aInHosDate：入院日期(8位)(YYYYMMDD)not null
    StrInput = StrInput & vbTab & Nvl(rsTemp!入院日期)
    'aDepartmentName：住院科室(20位)
    StrInput = StrInput & vbTab & Nvl(rsTemp!入院科室)
    'aSickArea：病区(20位)
    StrInput = StrInput & vbTab & Nvl(rsTemp!病区)
    
    gstrSQL = "Select * From 床位状况记录 D where 病区ID=[1] And 床号=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "读取床位信息", CLng(Nvl(rsTemp!当前病区ID, 0)), CLng(Nvl(rsTemp!当前床号)))
    If rsData.EOF Then
        'aRoom：房间(10位)
        StrInput = StrInput & vbTab & ""
        'aBedNo：床号(3位)
        StrInput = StrInput & vbTab & ""
    Else
        'aRoom：房间(10位)
        StrInput = StrInput & vbTab & Nvl(rsData!房间号)
        'aBedNo：床号(3位)
        StrInput = StrInput & vbTab & Right(Nvl(rsData!床号), 3)
    End If
    'aClinicDoctorCode：门诊医生(10位)
    StrInput = StrInput & vbTab & Nvl(rsTemp!门诊医师)
    'aInHosDiagnoseCode：入院诊断代码(病种)(10位)not null
    'strInput = strInput & vbTab & Substr(g病人身份_吉林.病种代码, 1, 10)
    StrInput = StrInput & vbTab & Substr(g病人身份_吉林.病种代码, 1, 10)
    StrInput = StrInput & vbTab & Substr(g病人身份_吉林.诊断名称, 1, 200)
    
    If 业务请求_吉林(入院登记, StrInput, strOutput) = False Then Exit Function
    If 业务请求_吉林(服务数据提交, "", strOutput) = False Then Exit Function
    If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
    
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_吉林 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_吉林 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_吉林 = False
End Function
Public Function 业务请求_吉林(ByVal intType As 业务类型_吉林, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strReturn As String
    Dim strOutput(0 To 20) As String, dblOutPut(0 To 25) As Double, intOutPut(0 To 5) As Integer, lngOutPut(0 To 5) As Long
    Dim strArr1
    Dim strArr(0 To 20) As String
    Dim strReg As String
    Dim str名称 As String
    
    Dim i As Integer
    str名称 = Get交易代码(intType, True)
    DebugTool "进入业务请求函数(业务类型为:" & intType & " 业务名称:" & str名称 & ")," & vbCrLf & "   输入参数为" & strInputString
    
    业务请求_吉林 = False
    
    StrInput = strInputString
    
    If InitInfor_吉林.模拟数据 Then
        '读取模拟数据
        Read模拟数据 intType, strInputString, strOutPutstring
         业务请求_吉林 = True
        Exit Function
    End If
   
    strArr1 = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr1)
        strArr(i) = strArr1(i)
    Next
        
    
    Err = 0
    On Error GoTo errHand:
    
    Select Case intType
        Case 读卡
           lngReturn = ReadCard(strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), strOutput(5), strOutput(6), strOutput(7), strOutput(8), strOutput(9), strOutput(10), strOutput(11), strOutput(12), strOutput(13), strOutput(14), strOutput(15), strOutput(16), strOutput(17), dblOutPut(0), dblOutPut(1), dblOutPut(2), dblOutPut(3), dblOutPut(4), dblOutPut(5), lngOutPut(0), lngOutPut(1))
           If lngReturn < 0 Then
                ShowMsgbox "在进行医保读卡时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                Exit Function
           End If
           '构建返回串
           strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & strOutput(5) & vbTab & strOutput(6) & vbTab & strOutput(7) & vbTab & strOutput(8) & vbTab & strOutput(9) & vbTab & strOutput(10) & vbTab & strOutput(11) & vbTab & strOutput(12) & vbTab & strOutput(13) & vbTab & strOutput(14) & vbTab & strOutput(15) & vbTab & strOutput(16) & vbTab & strOutput(17) & vbTab & dblOutPut(0) & vbTab & dblOutPut(1) & vbTab & dblOutPut(2) & vbTab & dblOutPut(3) & vbTab & dblOutPut(4) & vbTab & dblOutPut(5) & vbTab & lngOutPut(0) & vbTab & lngOutPut(1)
        Case 初始化服务调用
            '门诊收费(10),门诊退费(11),入院登记(20),医嘱录入(21),住院记帐(22),住院结算(23),出院登记(24),取消入院登记(25),冲销住院记帐(26),取消结算(27),取消出院登记(28),取消医嘱录入(29),其他
            lngReturn = InitCalc(strArr(0))
           If lngReturn < 0 Then
                ShowMsgbox "在进行医保初始化服务时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                Exit Function
           End If
        Case 入院登记
           lngReturn = InhosRegister(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9))
           If lngReturn < 0 Then
                ShowMsgbox "在进行入院登记时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                Exit Function
           End If
        Case 取消入院登记
           lngReturn = uInhosRegister(0)
           If lngReturn < 0 Then
                ShowMsgbox "在进行取消入院登记时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                Exit Function
           End If
        Case 出院登记
           lngReturn = OuthosRegister(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), Val(strArr(9)), strArr(10))
           If lngReturn < 0 Then
                ShowMsgbox "在进行出院登记时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                Exit Function
           End If
            
        Case 服务数据提交
           lngReturn = CommitData()
           If lngReturn < 0 Then
                ShowMsgbox "在进行服务数据提交时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                Exit Function
           End If
        Case 结束服务调用
           lngReturn = FinalCalc()
           If lngReturn < 0 Then
                ShowMsgbox "在进行结束服务调用时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                Exit Function
           End If
        Case 设置发票数据
            lngReturn = SetClinicBill(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), Val(strArr(8)), Val(strArr(9)))
            If lngReturn < 0 Then
                ShowMsgbox "在进行设置发票数据时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                Exit Function
           End If
        Case 取药品信息
            lngReturn = GetInfo_MediDic(strArr(0), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), dblOutPut(0), dblOutPut(1))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取药品信息时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
           strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & dblOutPut(0) & vbTab & dblOutPut(1)
           
        Case 取服务信息
            lngReturn = GetInfo_ServerDic(strArr(0), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), dblOutPut(0), dblOutPut(1))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取服务信息时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & dblOutPut(0) & vbTab & dblOutPut(1)
                   
        Case 取诊疗信息
            lngReturn = GetInfo_ItemDic(strArr(0), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), dblOutPut(0), dblOutPut(1))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取诊疗信息时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & dblOutPut(0) & vbTab & dblOutPut(1)
        Case 设置门诊收据明细
            lngReturn = SetClinicBillDetail(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), Val(strArr(10)), Val(strArr(11)), Val(strArr(12)))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取诊疗信息时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 设置门诊大类信息
            lngReturn = SetClinicMediKind(Val(strArr(0)), strArr(1), Val(strArr(2)))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行设置门诊大类信息时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 取得门诊收据计算信息
            lngReturn = GetClinicBillData(1, dblOutPut(0), dblOutPut(1), dblOutPut(2), dblOutPut(3), dblOutPut(4), dblOutPut(5), dblOutPut(6), dblOutPut(7), dblOutPut(8), dblOutPut(9), dblOutPut(10), dblOutPut(11), dblOutPut(12), dblOutPut(13), dblOutPut(14), dblOutPut(15), dblOutPut(16), dblOutPut(17), dblOutPut(18), dblOutPut(19), dblOutPut(20), dblOutPut(21), dblOutPut(22), dblOutPut(23))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取得门诊收据计算信息时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
            strReturn = ""
            For i = 0 To 23
                '构建返回串
                strReturn = strReturn & dblOutPut(i) & vbTab
            Next
        Case 政策服务提交
            lngReturn = CommitDataX(Val(strArr(0)), Val(strArr(1)))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行政策服务提交时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 取消结算
            lngReturn = ReturnCharge(Val(strArr(0)), strArr(1), strArr(2))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取消结算时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 设置住院帐单
            lngReturn = SetInHosBill(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), Val(strArr(5)), Val(strArr(6)))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取消结算时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 设置记帐单明细数据
            lngReturn = SetInHosBillDetail(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), Val(strArr(10)), Val(strArr(11)), Val(strArr(12)))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行设置记帐单明细数据时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 取消出院登记
            lngReturn = uOuthosRegister(strArr(0))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取消出院登记时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 设置结算单
            lngReturn = SetCheckOutBill(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), Val(strArr(9)), Val(strArr(10)))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行设置结算单时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 设置住院大类信息
            lngReturn = SetInHosMediKind(Val(strArr(0)), strArr(1), Val(strArr(2)))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行设置住院大类信息时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
        Case 取得住院结算计算信息
            lngReturn = GetCheckOutBillData(1, dblOutPut(0), dblOutPut(1), dblOutPut(2), dblOutPut(3), dblOutPut(4), dblOutPut(5), dblOutPut(6), dblOutPut(7), dblOutPut(8), dblOutPut(9), dblOutPut(10), dblOutPut(11), dblOutPut(12), dblOutPut(13), dblOutPut(14), dblOutPut(15), dblOutPut(16), dblOutPut(17), dblOutPut(18), dblOutPut(19), dblOutPut(20), dblOutPut(21), dblOutPut(22), dblOutPut(23))
            If lngReturn < 0 Then
                 ShowMsgbox "在进行取得住院结算计算信息时发生如下错误：" & vbCrLf & "错误号:" & lngReturn & vbCrLf & "错误描述:" & GetErrMsg()
                Call 业务请求_吉林(结束服务调用, "", "")
                 Exit Function
            End If
            strReturn = ""
            For i = 0 To 23
                '构建返回串
                strReturn = strReturn & dblOutPut(i) & vbTab
            Next
    End Select
    strOutPutstring = strReturn
    业务请求_吉林 = True
    DebugTool "     输出参数为:" & strReturn
    DebugTool "业务请求成功(业务类型为:" & intType & " 业务名称:" & str名称 & ")"
     Exit Function
errHand:
    DebugTool "业务请求失败(业务类型为:" & intType & " 业务名称:" & str名称 & ")"
    If ErrCenter = 1 Then
        Resume
    End If
End Function
    
    
    



Public Function 入院登记撤销_吉林(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
            
    '刘兴宏:20040923增加的
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim str医保号  As String
    
    Err = 0
    On Error GoTo errHand
    
    DebugTool "进入扩院登撤消接口"
    
    入院登记撤销_吉林 = False
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用，不能撤消入院登记"
        Exit Function
    End If
    
    '先初始化服务
    If 业务请求_吉林(初始化服务调用, "25", strOutput) = False Then Exit Function
    
    '需读卡验正
    If 业务请求_吉林(读卡, "", strOutput) = False Then Exit Function
    If 业务请求_吉林(取消入院登记, "", strOutput) = False Then Exit Function
    If 业务请求_吉林(服务数据提交, "", strOutput) = False Then Exit Function
    If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
    DebugTool "调用医保的取消业务成功,并开始更新保险帐户的相关状态！"
    
    '更新医保帐户
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_吉林 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    
    DebugTool "取消成功"
    入院登记撤销_吉林 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 出院登记_吉林(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    '个人状态的修改
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Dim datCurr As Date, bln零费用出院 As Boolean, str住院号 As String, _
        strInNote As String, str病种编码 As String, str医保号 As String
    Dim StrInput As String, strOutput As String
    
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    '刘兴宏:20040924更改
    bln零费用出院 = Not 存在未结费用(lng病人ID, lng主页ID)
    
    
    
    If bln零费用出院 = True Then
        '刘兴宏:20040924更正
       If 入院登记撤销_吉林(lng病人ID, lng主页ID) = True Then
            出院登记_吉林 = True
       End If
        Exit Function
    End If
        
    '先确定是否已经存在出院病种
    Dim str出院病种 As String, str并发症 As String
    
Go病种:
    If frm病种选择_吉林.ShowSelect(TYPE_吉林, lng病人ID, lng主页ID, str出院病种, str并发症) = False Then Exit Function
    
    gstrSQL = "" & _
        "   Select a.*,b.编码,b.名称,a.并发症 From 保险帐户 a,保险病种 b where a.出院病种ID=b.ID(+) and a.病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取出院病种"
    str出院病种 = Nvl(rsTemp!编码)
    If Nvl(rsTemp!编码) = "" Or Nvl(rsTemp!并发症) = "" Then
           If MsgBox("由于没有病种或并发症，所以不能出院登记,是否重新录入?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                GoTo Go病种:
            Else
                Exit Function
            End If
    End If
    
        
    If 业务请求_吉林(初始化服务调用, "24", strOutput) = False Then Exit Function
    
    '获取出院诊断
    'strInNote = 获取入出院诊断(lng病人id, lng主页ID, False, False, True)
    
    '获取住院医师
    gstrSQL = "" & _
        "   Select A.入院日期,(sysdate-a.入院日期)/365 as 住院天数,b.当前床号,B.住院号," & _
        "           to_char(A.入院日期,'ss') as 序号,A.当前病区id,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
        "           C.密码,G.出院编码,D.编码 As 科室编码,J.名称 as 病区,A.出院方式 " & _
        "   from 病案主页 A,病人信息 B,保险帐户 C,部门表 D,部门表 J, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码,'')) AS 出院编码 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 = 3 and a.主页id=[2] and a.病人id=[1] Group by 病人id,主页id)   G" & _
        "   Where   A.病人ID = B.病人ID And A.病人ID = C.病人ID And A.当前病区ID=J.id(+) and " & _
        "           A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]" & _
        "           and A.主页id=G.主页id(+) and a.病人id=G.病人id(+) " & _
        ""
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    
    If rsTemp.EOF Then
        MsgBox "不能取得病人的入院登记信息", vbInformation, gstrSysName
        Exit Function
    End If
    gstrSQL = "Select * From 床位状况记录 D where 病区ID=[1] And 床号=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "读取床位信息", CLng(Nvl(rsTemp!当前病区ID, 0)), CLng(Nvl(rsTemp!当前床号)))
    
    '刘兴宏:20040923更证了住院号
    'str住院号 = Format(lng病人ID, "0#########") & Format(lng主页ID, "0##")
    'aInHosNo：住院号(13位)not null
    StrInput = lng病人ID & "-" & lng主页ID & "-" & Nvl(rsTemp!序号) & vbTab
    'aOutHosDate：出院日期(8位)(YYYYMMDD)not null
    StrInput = StrInput & Format(datCurr, "yyyymmdd") & vbTab
    'aDepartmentName：住院科室(20位)
    StrInput = StrInput & Substr(Nvl(rsTemp!住院科室), 1, 20) & vbTab
    'aSickArea：病区(20位)
    StrInput = StrInput & Substr(Nvl(rsTemp!病区), 1, 20) & vbTab
    'aRoom：房间(10位)
    'aBedNo：床号(3位)
    If rsData.EOF Then
        StrInput = StrInput & "" & vbTab
        StrInput = StrInput & "" & vbTab
    Else
        StrInput = StrInput & Substr(Nvl(rsData!房间号), 1, 10) & vbTab
        StrInput = StrInput & Right(Nvl(rsData!床号), 3) & vbTab
    End If
    
    'aDoctorCode：医生(主治)(10位)
    StrInput = StrInput & Substr(Nvl(rsTemp!住院医师), 1, 10) & vbTab
    'aoutHosDiagnoseCode：出院诊断代码(病种)(10位)not null
    'strInput = strInput & Substr(g病人身份_吉林.病种代码, 1, 10) & vbTab
    
    StrInput = StrInput & Substr(str出院病种, 1, 10) & vbTab
    '1-治愈2-好转 3-未愈 4-死亡9-其他
    'aOutHosCure：出院治疗情况(1-治愈2-好转 3-未愈 4-死亡9-其他)(1位)not null
     StrInput = StrInput & Substr(Get治渝情况_吉林(lng病人ID, lng主页ID), 1, 1) & vbTab
    'aInHosDays：住院天数(3位)
    StrInput = StrInput & Substr(Int(Nvl(rsTemp!住院天数, 0)), 1, 3) & vbTab
    'aoutHostYPE: 出院类型
    StrInput = StrInput & "1" & vbTab

    If 业务请求_吉林(出院登记, StrInput, strOutput) = False Then Exit Function
    If 业务请求_吉林(服务数据提交, StrInput, strOutput) = False Then Exit Function
    If 业务请求_吉林(结束服务调用, StrInput, strOutput) = False Then Exit Function
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_吉林 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_吉林 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_吉林 = False
End Function
Private Function Get治渝情况_吉林(lng病人ID As Long, lng主页ID As Long) As String
    '功能:获取治渝情况标识
    '     A-治愈、B-好转、C-未愈、D-死亡、E-其他
    '??49  治愈情况标识    CHAR    439 1   1治愈、2好转、3未愈、4死亡、5其他，住院必添 院端
    'A-治愈、B-好转、C-未愈、D-死亡、E-其他
    
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.出院情况" & _
             " From 诊断情况 A,疾病编码目录 B " & _
             " Where A.病人ID=[1] And A.疾病ID=B.ID(+) And A.主页ID=[2]" & _
             "       And A.诊断类型 in (2,3)" & _
             " Order by A.诊断类型 Desc"
    Set rsInNote = zlDatabase.OpenSQLRecord(strTmp, "医保接口", lng病人ID, lng主页ID)
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = Nvl(rsInNote!出院情况)
    End If
    strTmp = Decode(strTmp, "治愈", "1", "好转", "2", "未愈", "3", "死亡", "4", "其他", "9", "1")
    Get治渝情况_吉林 = strTmp
   Call WriteDebugInfor_大连("Get治渝情况_吉林", lng病人ID)
End Function

Public Function 出院登记撤销_吉林(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '出院登记撤消
    Dim StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    gstrSQL = "Select to_char(入院日期,'ss') as 序号 From 病案主页 where 病人id= " & lng病人ID & " and 主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "出院登记"
    If rsTemp.EOF Then Exit Function
    
    StrInput = lng病人ID & "-" & lng主页ID & "-" & rsTemp!序号
    
    If 业务请求_吉林(初始化服务调用, "28", strOutput) = False Then Exit Function
    If 业务请求_吉林(取消出院登记, StrInput, strOutput) = False Then Exit Function
    If 业务请求_吉林(服务数据提交, StrInput, strOutput) = False Then Exit Function
    If 业务请求_吉林(结束服务调用, StrInput, strOutput) = False Then Exit Function
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_吉林 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    出院登记撤销_吉林 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_吉林(lng结帐ID As Long, Optional ByRef strAdvance As String) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
    '        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
    '        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str病种代码 As String, str结算方式 As String, str并发证 As String
    Dim dbl中药付数  As Double, dbl金额 As Double
    Dim lng主页ID As Long
    Dim lng病人ID As Long, iLoop As Integer
    Dim StrInput  As String, strOutput As String, str是否药品 As String
    Dim strArr
    Dim blnOld As Boolean '是否需要填写校正字段
    
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select 当前状态 from 保险帐户 where 病人id=" & lng病人ID
    

    gstrSQL = " " & _
        "        select a.实收金额,a.id,a.记录性质,a.主页id,a.记录状态,a.发生时间,a.登记时间,a.no,a.病人病区id,a.床号,a.序号,a.标识号 as 住院号,a.病人科室id,a.病人id,a.收费类别,b.类别,a.计算单位, " & _
        "               A.计算单位,A.付数,A.数次*Nvl(A.付数,1) 数量,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额 ,a.开单人 as 医生,c.编号 as 医生编号, " & _
        "               a.医嘱序号,nvl(a.婴儿费,0) as 婴儿费,to_char(F1.入院日期,'yyyyMMdd') as 入院日期,(F1.出院日期-F1.入院日期)/365 AS 住院天数,to_char(F1.出院日期,'yyyyMMDD') as 出院日期, A.实收金额,nvl(A.是否上传,0) as 是否上传, " & _
        "               D.编码 as 项目编码,D.名称 as 项目名称,decode(J.标识码,null,D.标识主码||D.标识子码,nvl(J.标识码,' ')) as 国家编码, " & _
        "               E.项目编码 as 医保编码,E.项目名称 as 医保名称,e.是否医保,e.大类id,H.名称 as 开单部门,J.名称 as 剂型, " & _
        "               L.险类,l.中心 , l.卡号, l.医保号, l.人员身份, l.单位编码, l.顺序号, l.退休证号, l.帐户余额, l.当前状态, l.病种ID, l.在职, l.年龄段, l.灰度级, l.就诊时间 " & _
        "        from 住院费用记录 a,收费类别 b,病案主页 F1,人员表 c,收费细目 D,保险支付项目 E,保险帐户 L,部门表 H, " & _
        "             (Select distinct Q.药品id,Q.标识码,T.名称 From 药品目录 Q,药品信息 R,药品剂型 T  Where  Q.药名id=R.药名id and R.剂型=T.编码 ) J " & _
        "        where  a.记录状态<>0 and  a.收费类别=b.编码 and a.收费细目id=J.药品id(+)   and  Nvl(a.附加标志,0)<>9 and a.收费细目id=D.id and a.开单人=c.姓名(+)  and " & _
        "              a.收费细目id=E.收费细目ID and A.病人id=F1.病人id and A.主页id=F1.主页ID and a.病人id=L.病人ID and a.开单部门id=h.id  and " & _
        "              a.结帐ID = " & lng结帐ID & " And E.险类 = " & TYPE_吉林
        
    zlDatabase.OpenRecordset rs明细, gstrSQL, "提取住院结帐明细"
    If rs明细.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "没有结算明细,不能结算"
        Exit Function
    End If
    
    dbl中药付数 = 0
    dbl金额 = 0
    While Not rs明细.EOF
        dbl金额 = dbl金额 + Nvl(rs明细!结帐金额, 0)
        If Nvl(rs明细!收费类别) = "6" Or Nvl(rs明细!收费类别) = "7" Then
            dbl中药付数 = dbl中药付数 + Nvl(rs明细!付数, 0)
        End If
        rs明细.MoveNext
    Wend
    rs明细.MoveFirst
    If dbl金额 = 0 Then
        Err.Raise 9000, gstrSysName, "病人没有发生费用,不能进行医保处理"
        Exit Function
    End If
    
    lng病人ID = Nvl(rs明细!病人ID, 0)
    lng主页ID = Nvl(rs明细!主页ID, 0)
   
   gstrSQL = "" & _
        "   Select a.密码,b.编码,b.名称,a.并发症 From 保险帐户 a,保险病种 b where a.出院病种ID=b.ID(+) and a.病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取出院病种"
    str病种代码 = Nvl(rsTemp!编码)
    
    If Nvl(rsTemp!编码) = "" Or Nvl(rsTemp!并发症) = "" Then
        Err.Raise 9000, gstrSysName, "必需输入病种代码和并发症!"
        Exit Function
    End If
    g病人身份_吉林.门诊类型 = Nvl(rsTemp!密码, "1")
    
    'gstrSQL = "Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码,'')) AS 出院编码 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 = 3 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人id & " Group by 病人id,主页id"
'    gstrSQL = "Select * From 保险病种 where id=" & Nvl(rs明细!病种ID, 0)
    str病种代码 = Nvl(rsTemp!编码)
    str并发证 = Nvl(rsTemp!并发症)
    
    'aPrnHandle: [为与旧接口兼容而保留未用] ?
    StrInput = g病人身份_吉林.门诊类型
    'aCheckOutBillNo：结算单据号。(13位)not null
    StrInput = StrInput & vbTab & lng结帐ID
    'aInHosDate：入院日期(8位)(yyyymmdd)not null
    StrInput = StrInput & vbTab & Nvl(rs明细!入院日期)
    'aOutHosDate：出院日期(8位)(yyyymmdd)not null
    StrInput = StrInput & vbTab & Nvl(rs明细!出院日期)
    'aDepartmentName：科室(20位)
    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!开单部门), 1, 20)
    'aDoctorName：医生(10位)
    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!医生), 1, 10)
    'aoutHosDiagnoseCode：出院病种编码(10位)
    StrInput = StrInput & vbTab & Substr(Nvl(str病种代码), 1, 20)
    'aSubDiagnose：并发症说明(200位)
    StrInput = StrInput & vbTab & Substr(str并发证, 1, 200)
    'aOutHosCure：出院情况(1-治愈2-好转 3-未愈 4-死亡9-其他)(1位)not null
    StrInput = StrInput & vbTab & Substr(Get治渝情况_吉林(lng病人ID, lng主页ID), 1, 1)
    'aInHosDays：住院天数(3位)
    StrInput = StrInput & vbTab & Substr(Int(Nvl(rs明细!住院天数, 0)), 1, 3)
    'aAmount：相应结算费用总金额。(8位，2位小数)
    StrInput = StrInput & vbTab & dbl金额
    
    If 身份鉴别_吉林(4, "23") = False Then Exit Function
    If IS是否刷卡病人(lng病人ID) = False Then
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
        Exit Function
    End If
    
    If 业务请求_吉林(设置结算单, StrInput, strOutput) = False Then Exit Function
    
    delArrar            '清除费用大类记录数组
    
    
    Do While Not rs明细.EOF
        If Nvl(rs明细!医保编码, "") = "" Then
                Err.Raise 9000, gstrSysName, "项目[" & Nvl(rs明细!项目名称) & "]未设置对应的医保项目,不能用于医保"
                Exit Function
        End If
        
        If Nvl(rs明细!是否上传, 0) = 0 And Nvl(rs明细!实收金额, 0) <> 0 Then
            Err.Raise 9000, gstrSysName, " 存在未上传的明细,请重新预结一次!"
            Exit Function
        End If
        
        str是否药品 = "0"
        StrInput = Substr(Nvl(rs明细!医保编码), 1, 10)
        
        Select Case UCase(Nvl(rs明细!收费类别))
            Case "5", "6", "7"
                str是否药品 = "1"
                'aItemName：项目名称(40位)
                'aMediKindCode：费用大类(2位)
                'aIsCityRich：(0-甲或普通1-乙或高精尖)(1位)[新增2自费]
                'aIsCityMedi：是否医保(1位)(0-不是1-是)[为与旧接口兼容而保留]
                'aIsLimit：是否医保限价项目(1位)(0-不是1-是)
                'aCitySelfPayRate：自付比例(4位宽3位小数)
                'aPrice：标准单价(8位，2位小数)
                
                If 业务请求_吉林(取药品信息, StrInput, strOutput) = False Then Exit Function
            Case "J", "H", "I"
                If 业务请求_吉林(取服务信息, StrInput, strOutput) = False Then Exit Function
                str是否药品 = "2"
            Case Else
                If 业务请求_吉林(取诊疗信息, StrInput, strOutput) = False Then Exit Function
        End Select
        If strOutput = "" Then
            Err.Raise 9000, gstrSysName, "在获取药品等信息时出现返回值为空,请与医保提供商联系!" & vbCrLf & " 输入参数为:" & StrInput
            
            Exit Function
        End If
        strArr = Split(strOutput & vbTab & vbTab & vbTab, vbTab)
        
        '记录费用大类金额
        setArrar CStr(strArr(1)), Nvl(rs明细!结帐金额, 0)
        
        rs明细.MoveNext
    Loop
    
    
    For iLoop = 0 To 50
        If strTempArr(iLoop, 0) <> "" Then
            StrInput = "1"
            StrInput = StrInput & vbTab & strTempArr(iLoop, 0)
            StrInput = StrInput & vbTab & strTempArr(iLoop, 1)
            If 业务请求_吉林(设置住院大类信息, StrInput, strOutput) = False Then Exit Function
        Else
            Exit For
        End If
    Next
    
    If 业务请求_吉林(取得住院结算计算信息, "", strOutput) = False Then Exit Function
    'aAccRemain：帐户余额(8位，2位小数)
    'aPayAcc：个人帐户支付(8位，2位小数)
    'aPayCash：个人现金支付(8位，2位小数)
    'aPayPer：个人比例负担(8位，2位小数)
    'aPayPlan：统筹支付(8位，2位小数)
    'aPayCarePlan：照顾支付(8位，2位小数)
    'aPayCareSelf：照顾垫付(8位，2位小数)
    'aPaySelfPart：自付段支付(8位，2位小数)
    'aPayBusiness：商保支付(8位，2位小数)
    'aCompMediFir：甲类药品(8位，2位小数)
    'aCompMediSelf：自费药品(8位，2位小数)
    'aCompMediSec：乙类药品(8位，2位小数)
    'aCompTreatFir：甲类诊疗(普通)(8位，2位小数)
    'aCompTreatSelf：自费诊疗(8位，2位小数)
    'aCompTreatSec：乙类诊疗(高精尖)(8位，2位小数)
    'aCompBedFir：甲类设施(8位，2位小数)
    'aCompBedSelf：自费设施(8位，2位小数)
    'aCompBedSec：乙类设施(8位，2位小数)
    'aCompOtherSelf：其他自费(8位，2位小数)
    'aAccSelfPayPart：自付段累计(8位，2位小数)
    'aAccPlanPay：统筹支付累计(8位，2位小数)
    'aAccHeavyIll：重病支付累计(8位，2位小数)
    'aAccDeferIll：慢病支付累计(8位，2位小数)
    'aUBasePay  非基本医疗费（8位,2位小数）刘兴宏：补充
    If strOutput = "" Then
        Err.Raise 9000, gstrSysName, "在取得住院结算计算信息时，返回了空值!"
        Exit Function
    End If
    strArr = Split(strOutput, vbTab)
    
    
    If Val(strArr(1)) <> 0 Then
        str结算方式 = str结算方式 & "||个人帐户|" & Format(Val(strArr(1)), "####0.00;-####0.00; ;")
    End If
    
    If Val(strArr(4)) <> 0 Then
        str结算方式 = str结算方式 & "||统筹支付|" & Format(Val(strArr(4)), "####0.00;-####0.00; ;")
    End If
    
    '如果存在
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        #If gverControl < 2 Then
            blnOld = True
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
        #Else
            strAdvance = str结算方式
            gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    End If
    Dim intMouse As Integer
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If blnOld Then
        If frm结算信息.ShowME(lng结帐ID, True) = False Then
            Exit Function
        End If
    End If
    Screen.MousePointer = intMouse
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(自付段累计),帐户累计支出_IN(统筹支付累计),累计进入统筹_IN(重病支付累计),累计统筹报销_IN(慢病支付累计),住院次数_IN,起付线(非基本医疗费),封顶线_IN(个人现金支付),实际起付线_IN(个人比例负担),
    '   发生费用金额_IN(费用总额),全自付金额_IN(自付段支付),首先自付金额_IN(自费药品),
    '   进入统筹金额_IN(甲类药品),统筹报销金额_IN(统筹支付),    大病自付金额_IN(照顾支付),超限自付金额_IN(照顾垫付),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(),主页ID_IN,中途结帐_IN,备注_IN
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    
    With g病人身份_吉林
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_吉林 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          Val(strArr(19)) & "," & Val(strArr(20)) & "," & Val(strArr(21)) & "," & Val(strArr(22)) & ",NULL," & Val(strArr(23)) & "," & Val(strArr(2)) & "," & Val(strArr(3)) & "," & _
         dbl金额 & "," & Val(strArr(7)) & "," & Val(strArr(10)) & "," & _
          Val(strArr(9)) & "," & Val(strArr(4)) & "," & Val(strArr(5)) & "," & Val(strArr(6)) & "," & Val(strArr(1)) & "," & _
          "'" & g病人身份_吉林.门诊类型 & "',Null,Null,NULL" & IIf(blnOld, "", ",1") & ")"
              
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        '附加信息参数:
            '性质_IN ,记录ID_IN,帐户余额_IN ,商保支付_IN ,乙类药品_IN ,甲类诊疗_IN ,自费诊疗_IN ,乙类诊疗_IN ,甲类设施_IN ,自费设施_IN ,乙类设施_IN ,其他自费_IN
        gstrSQL = "zl_保险结算记录_附加信息( 2," & lng结帐ID & "," & Val(strArr(0)) & "," & Val(strArr(8)) & "," & Val(strArr(11)) & "," & Val(strArr(12)) & "," & Val(strArr(13)) & "," & Val(strArr(14)) & "," & Val(strArr(15)) & "," & Val(strArr(16)) & "," & Val(strArr(17)) & "," & Val(strArr(18)) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录附加信息")
    End With
    
    StrInput = strArr(1)    '个人帐户支付
    StrInput = StrInput & vbTab & strArr(2) '现金支付
    
    If 业务请求_吉林(政策服务提交, StrInput, strOutput) = False Then Exit Function
    If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
    
    住院结算_吉林 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

    Public Function 住院结算冲销_吉林(lng结帐ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput  As String
    Dim strArr
    Dim lng冲销ID As Long
    
    Dim lng病人ID As Long
    
    On Error GoTo errHandle
    
    
    '退费
    gstrSQL = "select distinct A.病人id, A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "大连医保", lng结帐ID)
    'gstrSQL = "select distinct A.病人id,A.结帐ID from 病人费用记录 A,病人费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取冲销ID"
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "没有相关单据"
        Exit Function
    End If
    
    lng冲销ID = Nvl(rsTemp!ID)
    lng病人ID = Nvl(rsTemp!病人ID)
   
    StrInput = "1"
    StrInput = StrInput & vbTab & lng结帐ID & "R"
    StrInput = StrInput & vbTab & lng结帐ID
    If 身份鉴别_吉林(4, "27") = False Then Exit Function
    If IS是否刷卡病人(lng病人ID) = False Then
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
        Exit Function
    End If
    
    If 业务请求_吉林(取消结算, StrInput, strOutput) = False Then Exit Function
       
    gstrSQL = "Select * From 保险结算记录 Where 记录ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID, TYPE_吉林)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "没有找到原有的保险结算记录", vbInformation, gstrSysName
        Exit Function
    End If
        
    With g病人身份_吉林
        gstrSQL = "zl_保险结算记录_insert( 2," & lng冲销ID & "," & TYPE_吉林 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          -1 * Nvl(rsTemp!帐户累计增加, 0) & "," & -1 * Nvl(rsTemp!帐户累计支出, 0) & "," & -1 * Nvl(rsTemp!累计进入统筹, 0) & "," & -1 * Nvl(rsTemp!累计统筹报销, 0) & ",NULL," & -1 * Nvl(rsTemp!起付线, 0) & "," & -1 * Nvl(rsTemp!封顶线, 0) & "," & -1 * Nvl(rsTemp!实际起付线, 0) & "," & _
         -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * Nvl(rsTemp!首先自付金额, 0) & "," & _
          -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & "," & -1 * Nvl(rsTemp!大病自付金额, 0) & "," & -1 * Nvl(rsTemp!超限自付金额, 0) & "," & -1 * Nvl(rsTemp!个人帐户支付, 0) & "," & _
          "NULL,Null,Null,NULl)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        '附加信息参数:
        '性质_IN ,记录ID_IN,帐户余额_IN ,商保支付_IN ,乙类药品_IN ,甲类诊疗_IN ,自费诊疗_IN ,乙类诊疗_IN ,甲类设施_IN ,自费设施_IN ,乙类设施_IN ,其他自费_IN
        gstrSQL = "zl_保险结算记录_附加信息( 2," & lng冲销ID & "," & -1 * Nvl(rsTemp!帐户余额, 0) & "," & -1 * Nvl(rsTemp!商保支付, 0) & "," & -1 * Nvl(rsTemp!乙类药品, 0) & "," & -1 * Nvl(rsTemp!甲类诊疗, 0) & "," & -1 * Nvl(rsTemp!自费诊疗, 0) & "," & -1 * Nvl(rsTemp!乙类诊疗, 0) & "," & -1 * Nvl(rsTemp!甲类设施, 0) & "," & -1 * Nvl(rsTemp!自费设施, 0) & "," & -1 * Nvl(rsTemp!乙类设施, 0) & "," & -1 * Nvl(rsTemp!其他自费, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录附加信息")
    End With
    If 业务请求_吉林(服务数据提交, "", strOutput) = False Then Exit Function
    If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
    住院结算冲销_吉林 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function 处方登记_吉林(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim str病种代码 As String
    Dim dbl付数 As Double, dbl金额 As Double
    Dim StrInput As String, strOutput As String
    Dim str是否药品  As String
    Dim strArr
    Dim collData  As Collection
    
    
    Err = 0
    On Error GoTo errHand:
    
    处方登记_吉林 = False
    DebugTool "进入处方登记:" & Time
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then

        gstrSQL = " " & _
            " Select A.id,A.病人ID,F.住院号,A.NO,A.序号,A.医嘱序号,A.记录性质,A.记录状态,A.收费类别,D.类别,to_char(A.登记时间,'yyyyMMddhh24miss') 登记时间, " & _
            "        A.开单人 医生,F1.住院医师 住院医生,to_char(f1.入院日期,'ss') as 登记序号,a.是否上传,V.编号 AS 医生编号,B.名称 开单部门,A.收费细目ID,A.计算单位,A.付数,Round(A.实收金额/(A.数次*A.付数),2) as 实际价格,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传, " & _
            "        C.项目编码 医保项目编码,G.规格 ,F.住院次数 AS 主页id, " & _
            "        G.标识主码||G.标识子码 AS 国家编码,G.名称 AS 项目名称,K.名称 AS 剂型, " & _
            "        E.险类,E.中心,E.卡号,E.医保号,E.密码,E.人员身份,E.单位编码,E.顺序号,E.退休证号,E.帐户余额,E.当前状态, " & _
            "        E.病种ID,E.在职,E.年龄段,E.灰度级,to_char(E.就诊时间,'yyyyMMddhh24miss') 就诊时间 " & _
            " From 住院费用记录 A,部门表 B,收费类别 D,保险帐户 E,病人信息 F,病案主页 F1,收费细目 G,人员表 V," & _
            "       (Select J.名称,O.药品id From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K, " & _
            "       (Select M.项目编码,M.收费细目id From 保险支付项目 M Where M.险类=[4]) C " & _
            " Where   a.记录状态<>0 and   a.病人id=E.病人ID AND a.病人id=F.病人ID AND A.病人id=F1.病人id and F1.险类=" & TYPE_吉林 & " AND F.主页id= F1.主页id  AND  a.开单人=V.姓名(+) AND a.收费细目id=k.药品id(+) AND a.收费细目id=G.id AND E.险类=" & TYPE_吉林 & "   AND A.收费类别=D.编码 AND  " & _
            "           A.记录性质=[1] and  A.记录状态=[2] And A.NO=[3]" & _
            "           And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 " & _
            " order by A.病人ID,A.记录状态"
    Else
        gstrSQL = " " & _
            " Select A.id,A.病人ID,F.住院号,A.NO,A.序号,A.医嘱序号,A.记录性质,A.记录状态,A.收费类别,D.类别,to_char(A.登记时间,'yyyyMMddhh24miss') 登记时间, " & _
            "        A.开单人 医生,F1.住院医师 住院医生,to_char(f1.入院日期,'ss') as 登记序号,a.是否上传,V.编号 AS 医生编号,B.名称 开单部门,A.收费细目ID,A.计算单位,A.付数,Round(A.实收金额/(A.数次*A.付数),2) as 实际价格,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传, " & _
            "        C.项目编码 医保项目编码,G.规格 ,F.住院次数 AS 主页id, " & _
            "        G.标识主码||G.标识子码 AS 国家编码,G.名称 AS 项目名称,K.名称 AS 剂型, " & _
            "        E.险类,E.中心,E.卡号,E.医保号,E.密码,E.人员身份,E.单位编码,E.顺序号,E.退休证号,E.帐户余额,E.当前状态, " & _
            "        E.病种ID,E.在职,E.年龄段,E.灰度级,to_char(E.就诊时间,'yyyyMMddhh24miss') 就诊时间 " & _
            " From 住院费用记录 A,部门表 B,收费类别 D,保险帐户 E,病人信息 F,病案主页 F1,收费细目 G,人员表 V," & _
            "       (Select J.名称,O.药品id From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K, " & _
            "       (Select M.项目编码,M.收费细目id From 保险支付项目 M Where M.险类=[4]) C " & _
            " Where   a.记录状态<>0 and   a.病人id=E.病人ID AND a.病人id=F.病人ID AND A.病人id=F1.病人id and F1.险类=" & TYPE_吉林 & " AND F.住院次数= F1.主页id  AND  a.开单人=V.姓名(+) AND a.收费细目id=k.药品id(+) AND a.收费细目id=G.id AND E.险类=" & TYPE_吉林 & "   AND A.收费类别=D.编码 AND  " & _
            "           A.记录性质=[1] and  A.记录状态=[2] And A.NO=[3]" & _
            "           And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 " & _
            " order by A.病人ID,A.记录状态"

    End If
    
    '第一步: 读取费用明细记录
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取处方明细", lng记录性质, lng记录状态, str单据号, TYPE_吉林)
    
    If rs明细.RecordCount = 0 Then
        ShowMsgbox "没有明细记录!"
        Exit Function
    End If
    dbl付数 = 0
    dbl金额 = 0
    Dim lngCount As Long
    lngCount = 0
    lng病人ID = 0
    Set collData = New Collection
    
    While Not rs明细.EOF
        If Nvl(rs明细!医保项目编码, "") = "" Then
                ShowMsgbox "项目[" & Nvl(rs明细!项目名称) & "]未设置对应的医保项目,不能用于医保"
                Exit Function
        End If
        If lng病人ID <> Nvl(rs明细!病人ID, 0) Then
            lng病人ID = Nvl(rs明细!病人ID, 0)
            dbl金额 = 0: dbl付数 = 0
            collData.Add Array(dbl付数, dbl金额), "K" & lng病人ID
            lngCount = lngCount + 1
        End If
        collData.Remove "K" & lng病人ID
        
        dbl金额 = dbl金额 + rs明细!金额
        If Nvl(rs明细!收费类别) = "6" Or Nvl(rs明细!收费类别) = "7" Then
            dbl付数 = dbl付数 + Nvl(rs明细!付数, 0)
        End If
        collData.Add Array(dbl付数, dbl金额), "K" & lng病人ID
        
        rs明细.MoveNext
    Wend
    
    If lngCount > 1 Then
        ShowMsgbox "不能同时对多个病人进行记帐,病人数为:" & lngCount
        Exit Function
    End If
    If InitInfor_吉林.明细时实上传 = False Then
        处方登记_吉林 = True
        Exit Function
    End If
    rs明细.MoveFirst
    
    
    lng病人ID = Nvl(rs明细!病人ID, 0)
    lng主页ID = Nvl(rs明细!主页ID, 0)
    
    gstrSQL = "Select * From 保险病种 Where ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(Nvl(rs明细!病种ID, 0)), TYPE_吉林)
    str病种代码 = ""
    If Not rsTemp.EOF Then
        str病种代码 = Nvl(rsTemp!编码)
    End If
    
    If lng记录状态 = 1 Then
        If 业务请求_吉林(初始化服务调用, "22", strOutput) = False Then Exit Function
    End If
    
    lng病人ID = 0
   Do While Not rs明细.EOF
        
        If Nvl(rs明细!是否上传, 0) = 0 And rs明细!金额 <> 0 Then
            If lng记录状态 = 1 Then
                If lng病人ID <> Nvl(rs明细!病人ID, 0) Then
                    lng病人ID = Nvl(rs明细!病人ID, 0)
                    lng主页ID = Nvl(rs明细!主页ID, 0)
                    
                    
                    DebugTool "上传处方记帐单 开始:" & Time
                    'aInHosRegisterNo：本次处理的住院登记号。
                    StrInput = lng病人ID & "-" & lng主页ID & "-" & Nvl(rs明细!登记序号)
                    'aSerialNo：住院帐单号(13位)not null
                    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!NO) & "-" & Nvl(rs明细!记录性质, 0), 1, 13)
                    'aDiagnoseCode：病种代码(10位)not null
                    StrInput = StrInput & vbTab & Substr(str病种代码, 1, 10)
                    'aDepartmentName：科室(20位)
                    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!开单部门), 1, 20)
                    'aDoctorName：医生(10位)
                    StrInput = StrInput & vbTab & Substr(Nvl(rs明细!医生), 1, 10)
                    'aHerbalCopy：中草药付数(2位)
                    StrInput = StrInput & vbTab & Substr(collData("K" & lng病人ID)(0), 1, 10)
                            
                    'aAmount：金额(8位，2位小数)
                    StrInput = StrInput & vbTab & Substr(collData("K" & lng病人ID)(1), 1, 10)
                    If 业务请求_吉林(设置住院帐单, StrInput, strOutput) = False Then Exit Function
                    DebugTool "上传处方记帐单 结束:" & Time
                End If
                str是否药品 = "0"
                DebugTool "上传处方明细 开始:" & Time
                
                StrInput = Substr(Nvl(rs明细!医保项目编码), 1, 10)
                
                Select Case UCase(Nvl(rs明细!收费类别))
                    Case "5", "6", "7"
                        str是否药品 = "1"
                        If 业务请求_吉林(取药品信息, StrInput, strOutput) = False Then Exit Function
                        Case "J", "H", "I"
                        If 业务请求_吉林(取服务信息, StrInput, strOutput) = False Then Exit Function
                        str是否药品 = "2"

                    Case Else
                        If 业务请求_吉林(取诊疗信息, StrInput, strOutput) = False Then Exit Function
                End Select
                If strOutput = "" Then
                    strOutput = "" & vbTab & vbTab & vbTab & vbTab
'                    Exit Function
                End If
                strArr = Split(strOutput, vbTab)
                '调用接口,写入明细
                'aBillHandle: [为与旧接口兼容而保留未用] ?
                StrInput = "1"
                'aCityMediCareNo：医保项目编号。(10位)not null
                StrInput = StrInput & vbTab & Nvl(rs明细!医保项目编码, "")
                'aItemName：医院项目名称(40位)not null
                StrInput = StrInput & vbTab & Nvl(rs明细!项目名称, "")
                'aConformationName：剂型名称(20位)
                StrInput = StrInput & vbTab & Substr(Nvl(rs明细!剂型, ""), 1, 20)
                'aUnitContent：单位含量(14位)
                StrInput = StrInput & vbTab & Substr(Nvl(rs明细!规格, ""), 1, 14)
                    '刘兴宏:暂且屏蔽
                    'gstrSQL = "Select 单量,频次,用法 From 药品收发记录 where 费用id=" & Nvl(!ID, 0)
                    'zlDataBase.OpenRecordset rsTemp, gstrSQL, "获取病人单理及频次"
                    'If Not rsTemp.EOF Then
                    'strInput = strInput & vbTab & Substr("单量:" & Nvl(rsTemp!单量, "") & " 频次:" & Nvl(rsTemp!频次) & "用法:" & Nvl(rsTemp!用法), 1, 14)
                    'Else
                    'strInput = strInput & vbTab & ""
                    'End If
                'aDosage：用法用量(40位)
                StrInput = StrInput & vbTab & ""
                'aMediKindCode：费用大类代码(2位)not null
                StrInput = StrInput & vbTab & strArr(1)
                'aIsCityMedi：是否医保(1位)(0-不是1-是)[为与旧接口兼容而保留未用]
                StrInput = StrInput & vbTab & strArr(3)
                'aIsRich：(0-甲或普通1-乙或高精尖)(1位)[新增2自费]not null
                StrInput = StrInput & vbTab & strArr(2)
                'aIsMedi:是否药品(0-项目1-药品2服务设施[床] )not null
                StrInput = StrInput & vbTab & str是否药品
                'aPrice：单价(8位，2位小数)>0
                StrInput = StrInput & vbTab & Nvl(rs明细!实际价格, 0)
                'aQuantity：数量(8位，2位小数)>0
                StrInput = StrInput & vbTab & Nvl(rs明细!数量, 0)
                'aAmount：金额(8位，2位小数)>0
                StrInput = StrInput & vbTab & Nvl(rs明细!金额, 0)
                
                If 业务请求_吉林(设置记帐单明细数据, StrInput, strOutput) = False Then Exit Function
                DebugTool "上传处方明细 结束:" & Time
            End If
            gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细!ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        End If
        rs明细.MoveNext
    Loop
    
    If lng记录状态 <> 1 Then
        '冲销单据
        If 业务请求_吉林(初始化服务调用, "26", strOutput) = False Then Exit Function
        
'        If 身份鉴别_吉林(0, "26") = False Then Exit Function
        StrInput = "1"
        StrInput = StrInput & vbTab & str单据号 & "-" & lng记录性质 & "R"
        StrInput = StrInput & vbTab & str单据号 & "-" & lng记录性质
        If 业务请求_吉林(取消结算, StrInput, strOutput) = False Then Exit Function
        If 业务请求_吉林(服务数据提交, "", strOutput) = False Then Exit Function
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
    Else
        DebugTool "上传处方明细 数据提交开始:" & Time
        If 业务请求_吉林(服务数据提交, "", strOutput) = False Then Exit Function
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
        DebugTool "上传处方明细 数据提交失败:" & Time
    End If
    
    
    处方登记_吉林 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Function Read模拟数据(ByVal int业务类型 As 业务类型_吉林, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以便测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim STRNAME As String
    
    strFile = App.Path & "\模拟提交串.txt"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Select Case int业务类型
    Case 读卡
        STRNAME = "读卡"
    Case 初始化服务调用
        Exit Function
    Case 取消入院登记
        Exit Function
    Case 服务数据提交
        Exit Function
    Case 结束服务调用
        Exit Function
    Case 出院登记
        Exit Function
    End Select
   
    Dim blnStart As Boolean
    Dim strArr
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                    
                If blnStart Then
                    If strText = "" Then
                        strText = "" & vbTab
                    End If
                    strArr = Split(strText, "|")
                    
                    If Val(strArr(0)) = 1 Then
                        str = strArr(1)
                        Exit Do
                    End If
                Else
                     If "<" & STRNAME & ">" = strText Then
                         blnStart = True
                     End If
                End If
                If "</" & STRNAME & ">" = strText Then
                    Exit Do
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function
Public Function 身份鉴别_吉林(ByVal bytType As Byte, Optional str类型 As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:远程身份鉴别
    '--入参数:bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '--出参数:
    '--返  回:成功true,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim blnReturn As Boolean
    Dim StrInput As String
    Dim strOutput As String
    Dim strArr
    
        
    Err = 0
    On Error GoTo errHand:
    
    身份鉴别_吉林 = False
    Select Case bytType
    Case 0      '门诊
        StrInput = "10"
    Case 1      '入院登记
        StrInput = "20"
    Case Else
        StrInput = "23"
    End Select
    If str类型 <> "" Then
        '采取传入参数:
        StrInput = str类型
    End If
    
    If 业务请求_吉林(初始化服务调用, StrInput, strOutput) = False Then Exit Function
    
    If 业务请求_吉林(读卡, "", strOutput) = False Then Exit Function
    
    If strOutput = "" Then
        '刘兴宏 /*200408*/
        DebugTool "读取个人信息时出现了传出串为空了!"
        Exit Function
    End If
    
    strArr = Split(strOutput, vbTab)
    
    
    '给公用变量赋值
    With g病人身份_吉林
        .中心 = strArr(0)
        .卡号 = strArr(1)
        .身份证号 = strArr(2)
        .姓名 = strArr(3)
        .性别 = IIf(Val(strArr(4)) = 0, "女", "男")
        .出生日期 = GetStringToDate(strArr(5))
        .医保号 = strArr(6)
        .单位编码 = strArr(7)
        .个人身份 = strArr(8)
        .公务员标志 = strArr(9)  '(0-非公务员1-公务员，其他参照公务员)
        .补充保险 = strArr(10)    '0-不参加1-参加
        .大病医保 = strArr(11)    '0-不参加1-参加
        .隶属关系 = strArr(12)     '1-市属财政0-其他
        .是否慢性病 = strArr(13)   '0-不是1-是
        .重大疾病 = strArr(14)     '0-不是1-是
        .照顾级别 = strArr(15)    '0-无1-一级2-二级3-三级
        .职工属地 = strArr(16)     '0本地1常驻外地2-异地安置
        .住院标志 = strArr(17)     '0-不住院 1-住院
        
        .起付段医疗费累计 = Val(strArr(18))
        .统筹支付累计 = Val(strArr(19))
        .起付线金额累计 = Val(strArr(20))
        
        .慢病统筹支付累计 = Val(strArr(21))
        .大病累计 = Val(strArr(22))
        .帐户余额 = Val(strArr(23))
        .住院次数 = Val(strArr(24))
        .支付序列 = Val(strArr(25))
    End With
    身份鉴别_吉林 = True
    DebugTool "身份鉴别成功"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    身份鉴别_吉林 = False
End Function

Private Function GetStringToDate(ByVal StrInput As String) As String
    '功能:将形如"20040404"转换成"2004-04-04"
    Dim intLen As Integer
    Dim strTemp As String
    intLen = Len(StrInput)
    Select Case intLen
    Case 6
        strTemp = Left(StrInput, 4) & "-0" & Mid(StrInput, 5, 1) & "-0" & Mid(StrInput, 6, 1)
    Case 8
        strTemp = Left(StrInput, 4) & "-" & Mid(StrInput, 5, 2) & "-" & Mid(StrInput, 7, 2)
    Case Else
        strTemp = StrInput
    End Select
    GetStringToDate = strTemp
    
    
End Function

Public Function 住院虚拟结算_吉林(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    
    'rsExse:字符集
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    
    Dim rsTemp As New ADODB.Recordset
    Dim dbl总额 As Double
    Dim lng主页ID As Long
    Dim lng病人id1 As Long
    Dim intMouse  As Integer
    Err = 0
    On Error GoTo errHand:
    住院虚拟结算_吉林 = ""
    DebugTool "进入虚拟结算:" & Time
    
    gstrSQL = "Select 当前状态 From 保险帐户 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断是否出院没有"
    If Nvl(rsTemp!当前状态, 0) = 1 Then
        ShowMsgbox "该病人还办理出院,所以不能结算!"
        DebugTool "退出虚拟结算失败,病人未出院:" & Time
        Exit Function
    End If
    
    With rsExse
        dbl总额 = 0
        lng主页ID = Nvl(!主页ID, 0)
        Do While Not .EOF
            g病人身份_吉林.费用总额 = g病人身份_吉林.费用总额 + Nvl(!金额, 0)
            .MoveNext
        Loop
    End With
    
    If bln结帐处 Then
        Screen.MousePointer = 1
        If 身份标识_吉林(4, lng病人id1) = "" Then
            Screen.MousePointer = intMouse
            住院虚拟结算_吉林 = ""
            Exit Function
        End If
        Screen.MousePointer = intMouse
        If lng病人ID <> lng病人id1 Then
            ShowMsgbox "不是当前要结算的病人!"
            住院虚拟结算_吉林 = ""
            Exit Function
        End If
    End If
    
    Call 补传明细记录(lng病人ID, lng主页ID)


    DebugTool "退出虚拟结算成功:" & Time
    住院虚拟结算_吉林 = "统筹支付;" & 0 & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 补传明细记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '补传相关明细记录
    Dim rsTemp As New ADODB.Recordset
    Dim strNO As String, str病种代码 As String
    Dim dbl中药付数 As Double, dbl总额 As Double
    Dim StrInput  As String, strOutput As String
    Dim strArr
    Dim str是否药品  As String
    Dim bln提交 As Boolean
    
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select B.编码 From 保险帐户 A,保险病种 B where a.病种ID=B.ID and B.险类=" & TYPE_吉林 & " and a.病人ID=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病种"
    If Not rsTemp.EOF Then
        str病种代码 = Nvl(rsTemp!编码)
    End If
    补传明细记录 = False
    DebugTool "进入补明细接口:" & Time
    
    gstrSQL = "" & _
            "   Select A.ID,A.病人ID,a.主页id,A.开单人 as 医生,A.收费类别,A.记录性质,B.名称 as 开单部门,to_char(f.入院日期,'ss') as 登记序号,A.记录状态,A.NO,Decode(A.收费类别,'6',A.付数,'7',A.付数,0) as 中药付数," & _
            "           C.项目编码 as 医保项目编码,G.编码 as 项目编码,G.名称 as 收费项目,G.规格,K.名称 剂型," & _
            "           A.数次*A.付数 as 数量,Round(Nvl(A.实收金额,0)/(A.数次*A.付数),2) as 实际价格,Nvl(A.实收金额,0) as 实收金额" & _
            "   From 住院费用记录 A,部门表 B,病案主页 F,收费细目 G," & _
            "       (Select M.项目编码,M.收费细目id From 保险支付项目 M Where M.险类=" & TYPE_吉林 & ") C," & _
            "       (Select J.名称,O.药品id From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K " & _
            "   where nvl(A.是否上传,0)=0 and mod(nvl(A.记录状态,0),3)<>2 and nvl(a.记录状态,0)<>0  and A.记帐费用=1 and A.结帐ID is null and nvl(A.实收金额,0)<>0  and  " & _
            "       nvl(a.婴儿费,0)=0 and  A.收费细目id=K.药品id(+) And A.开单部门id+0=B.ID and A.收费细目id=G.id and A.收费细目id=C.收费细目id(+)  and A.病人id=F.病人id and A.主页id=F.主页id   and A.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & _
            "   ORDER BY a.记录性质,A.NO,A.序号"
            
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "补充上传明细"
    
    If rsTemp.EOF Then
        DebugTool "无补传记录,返回:" & Time
        GoTo go120:
    End If
    strNO = ""
    
    DebugTool "存在被传记录,返回:" & Time
    
    If 身份鉴别_吉林(1, "22") = False Then Exit Function
    
    If IS是否刷卡病人(Nvl(rsTemp!病人ID, 0)) = False Then
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
        Exit Function
    End If
  
    With rsTemp
        DebugTool "开始检查医保项目的对码情况,返回:" & Time
        Do While Not .EOF
            If Nvl(!医保项目编码) = "" Then
                ShowMsgbox "项目[" & Nvl(!项目名称) & "]未设置对应的医保项目,不能用于医保"
                Exit Function
            End If
            .MoveNext
        Loop
        DebugTool "结束检查医保项目的对码情况,返回:" & Time
        .MoveFirst
        '上传所有明细
        bln提交 = False
        DebugTool "开始补传处方明细,返回:" & Time
        strNO = ""
        Do While Not .EOF
            If strNO <> Nvl(!记录性质) & "-" & Nvl(!NO) & "-" & Nvl(!记录状态) Then
                strNO = Nvl(!记录性质) & "-" & Nvl(!NO) & "-" & Nvl(!记录状态)
                If GetSumJe(Nvl(!记录性质, 0), Nvl(!NO), Nvl(!记录状态, 1), dbl中药付数, dbl总额) = False Then Exit Function
                'aInHosRegisterNo：本次处理的住院登记号。
                StrInput = lng病人ID & "-" & lng主页ID & "-" & Nvl(!登记序号)
                'aSerialNo：住院帐单号(13位)not null
                StrInput = StrInput & vbTab & Substr(Nvl(!NO) & "-" & Nvl(!记录性质, 0), 1, 13)
                'aDiagnoseCode：病种代码(10位)not null
                StrInput = StrInput & vbTab & Substr(str病种代码, 1, 10)
                'aDepartmentName：科室(20位)
                StrInput = StrInput & vbTab & Substr(Nvl(!开单部门), 1, 20)
                'aDoctorName：医生(10位)
                StrInput = StrInput & vbTab & Substr(Nvl(!医生), 1, 10)
                'aHerbalCopy：中草药付数(2位)
                StrInput = StrInput & vbTab & Substr(dbl中药付数, 1, 10)
                'aAmount：金额(8位，2位小数)
                StrInput = StrInput & vbTab & Substr(dbl总额, 1, 10)
                
                If 业务请求_吉林(设置住院帐单, StrInput, strOutput) = False Then Exit Function
            End If
            
            str是否药品 = "0"
            StrInput = Substr(Nvl(!医保项目编码), 1, 10)
            
            Select Case UCase(Nvl(!收费类别))
                Case "5", "6", "7"
                    str是否药品 = "1"
                    If 业务请求_吉林(取药品信息, StrInput, strOutput) = False Then Exit Function
                Case "J", "H", "I"
                    If 业务请求_吉林(取服务信息, StrInput, strOutput) = False Then Exit Function
                    str是否药品 = "2"

                Case Else
                    If 业务请求_吉林(取诊疗信息, StrInput, strOutput) = False Then Exit Function
            End Select
            
            If strOutput = "" Then
                ShowMsgbox "在获取药品等信息时返回了空值，请与医保提供商联系!" & vbCrLf & " 输入参数为:" & StrInput
                Exit Function
            End If
            
            strArr = Split(strOutput, vbTab)
            
            '调用接口,写入明细
            'aBillHandle: [为与旧接口兼容而保留未用] ?
            StrInput = "1"
            'aCityMediCareNo：医保项目编号。(10位)not null
            StrInput = StrInput & vbTab & Nvl(!医保项目编码, "")
            'aItemName：医院项目名称(40位)not null
            StrInput = StrInput & vbTab & Nvl(!收费项目, "")
            'aConformationName：剂型名称(20位)
            StrInput = StrInput & vbTab & Substr(Nvl(!剂型, ""), 1, 20)
            'aUnitContent：单位含量(14位)
            StrInput = StrInput & vbTab & Substr(Nvl(!规格, ""), 1, 14)
                '刘兴宏:暂且屏蔽
                'gstrSQL = "Select 单量,频次,用法 From 药品收发记录 where 费用id=" & Nvl(!ID, 0)
                'zlDataBase.OpenRecordset rsTemp, gstrSQL, "获取病人单理及频次"
                'If Not rsTemp.EOF Then
                'strInput = strInput & vbTab & Substr("单量:" & Nvl(rsTemp!单量, "") & " 频次:" & Nvl(rsTemp!频次) & "用法:" & Nvl(rsTemp!用法), 1, 14)
                'Else
                'strInput = strInput & vbTab & ""
                'End If
            'aDosage：用法用量(40位)
            StrInput = StrInput & vbTab & ""
            'aMediKindCode：费用大类代码(2位)not null
            StrInput = StrInput & vbTab & strArr(1)
            'aIsRich：(0-甲或普通1-乙或高精尖)(1位)[新增2自费]not null
            StrInput = StrInput & vbTab & strArr(2)
            'aIsCityMedi：是否医保(1位)(0-不是1-是)[为与旧接口兼容而保留未用]
            StrInput = StrInput & vbTab & strArr(3)
            'aIsMedi:是否药品(0-项目1-药品2服务设施[床] )not null
            StrInput = StrInput & vbTab & str是否药品
            'aPrice：单价(8位，2位小数)>0
            StrInput = StrInput & vbTab & Nvl(!实际价格, 0)
            'aQuantity：数量(8位，2位小数)>0
            StrInput = StrInput & vbTab & Nvl(!数量, 0)
            'aAmount：金额(8位，2位小数)>0
            StrInput = StrInput & vbTab & Nvl(!实收金额, 0)
            If 业务请求_吉林(设置记帐单明细数据, StrInput, strOutput) = False Then Exit Function
            bln提交 = True
            gstrSQL = "zl_病人记帐记录_上传 ('" & !ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            .MoveNext
        Loop
        DebugTool "结束补传处方明细,返回:" & Time
        
        If bln提交 Then
            '有数据,需提交
            If 业务请求_吉林(服务数据提交, "", strOutput) = False Then Exit Function
        End If
        
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
    End With
    
go120:
    gstrSQL = "" & _
            "   Select A.ID,A.病人ID,a.主页id,A.开单人 as 医生,A.收费类别,A.记录性质,B.名称 as 开单部门,to_char(f.入院日期,'ss') as 登记序号,A.记录状态,A.NO,Decode(A.收费类别,'6',A.付数,'7',A.付数,0) as 中药付数," & _
            "           C.项目编码 as 医保项目编码,G.编码 as 项目编码,G.名称 as 收费项目,G.规格,K.名称 剂型," & _
            "           A.数次*A.付数 as 数量,Round(Nvl(A.实收金额,0)/(A.数次*A.付数),2) as 实际价格,Nvl(A.实收金额,0) as 实收金额" & _
            "   From 住院费用记录 A,部门表 B,病案主页 F,收费细目 G," & _
            "       (Select M.项目编码,M.收费细目id From 保险支付项目 M Where M.险类=" & TYPE_吉林 & ") C," & _
            "       (Select J.名称,O.药品id From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K " & _
            "   where nvl(A.是否上传,0)=0 and mod(nvl(A.记录状态,0),3)=2 and A.记帐费用=1 and A.结帐ID is null and nvl(A.实收金额,0)<>0  and  " & _
            "       nvl(a.婴儿费,0)=0 and  A.收费细目id=K.药品id(+) And A.开单部门id+0=B.ID and A.收费细目id=G.id and A.收费细目id=C.收费细目id(+)  and A.病人id=F.病人id and A.主页id=F.主页id   and A.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & _
            "   ORDER BY a.记录性质,A.NO,A.序号"
            
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "补充上传明细"
    
    If rsTemp.EOF Then
        补传明细记录 = True
        DebugTool "无补传的退费记录,返回:" & Time
        Exit Function
    End If
    
    With rsTemp
        DebugTool "开始检查医保项目的对码情况,返回:" & Time
        Do While Not .EOF
            If Nvl(!医保项目编码) = "" Then
                ShowMsgbox "项目[" & Nvl(!项目名称) & "]未设置对应的医保项目,不能用于医保"
                Exit Function
            End If
            .MoveNext
        Loop
        DebugTool "结束检查医保项目的对码情况,返回:" & Time
        .MoveFirst
        '冲销相关的单据
        strNO = ""
        If 身份鉴别_吉林(0, "26") = False Then Exit Function
        
        bln提交 = False
        Do While Not .EOF
            If strNO <> Nvl(!记录性质) & "-" & Nvl(!NO) & "-" & Nvl(!记录状态) Then
                strNO = Nvl(!记录性质) & "-" & Nvl(!NO) & "-" & Nvl(!记录状态)
                '冲销单据
                StrInput = "1"
                StrInput = StrInput & vbTab & Nvl(!NO) & "-" & Nvl(!记录性质, 0) & "R"
                StrInput = StrInput & vbTab & Nvl(!NO) & "-" & Nvl(!记录性质, 0)
                If 业务请求_吉林(取消结算, StrInput, strOutput) = False Then Exit Function
                bln提交 = True
            End If
            gstrSQL = "zl_病人记帐记录_上传 ('" & !ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            .MoveNext
        Loop
        If bln提交 Then
            If 业务请求_吉林(服务数据提交, "", strOutput) = False Then Exit Function
        End If
        
        If 业务请求_吉林(结束服务调用, "", strOutput) = False Then Exit Function
    End With
    
    DebugTool "补传记录上传成功,返回:" & Time
    补传明细记录 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetSumJe(ByVal lng记录性质 As Long, ByVal strNO As String, ByVal lng记录状态 As Long, dbl中药付数 As Double, dbl总额 As Double) As Boolean
    '功能:获取指定单据的汇总额
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
            "   Select  SUM(Decode(A.收费类别,'6',NVL(A.付数,0),'7',NVL(A.付数,0),0)) as 中药付数," & _
            "           Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 金额 From 住院费用记录 a " & _
            "   where nvl(A.是否上传,0)=0 and A.记录状态=" & lng记录状态 & " and A.记帐费用=1 and " & _
            "       nvl(a.婴儿费,0)=0  and a.记录性质=" & lng记录性质 & " and a.No='" & strNO & "'"
                
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取单据汇总额"
    dbl中药付数 = Nvl(rsTemp!中药付数, 0)
    dbl总额 = Nvl(rsTemp!金额, 0)
    GetSumJe = True
    Exit Function
errHand:
    dbl中药付数 = 0
    dbl总额 = 0
    GetSumJe = False
End Function
Public Function 挂号结算_吉林(ByVal lng结帐ID As Long) As Boolean
     挂号结算_吉林 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 挂号冲销_吉林(ByVal lng结帐ID As Long) As Boolean
    挂号冲销_吉林 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function 更新病种_吉林(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim str出院病种 As String, str并发症 As String
    
    Err = 0
    On Error GoTo errHand:
    
    更新病种_吉林 = frm病种选择_吉林.ShowSelect(TYPE_吉林, lng病人ID, lng主页ID, str出院病种, str并发症)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function 医保设置_吉林() As Boolean
    医保设置_吉林 = frmSet吉林.参数设置
End Function
'
