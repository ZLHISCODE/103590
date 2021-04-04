VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:所有调用服务接口
'一、病人域接口
'       1.zlPatiSvr_GetPatiInfo-根据条件获取病人详细信息服务接口
'       2.zl_PatiSvr_GetPatiInfsByRange-按条件范围批量获取病人信息服务接口
'       3.zlPatiSvr_GetCardTypes-获取医疗卡类别信息服务接口
'       4.zlPatiSvr_GetPatiID-根据条件获取病人ID值服务接口
'       5.zlPatiSvr_GetPatiExtendInfo-获取病人信息从表信息服务接口
'       6.ZlPatisvr_GetPatiCardInfo-获取病人的有效卡服务接口
'       7.zlPatiSvr_GetInsureByPatiID-获取病人的医保信息服务接口
'二、费用域接口
'  1.zlExseSvr_GetPatiSurplusInfo-获取病人费用余额信息
'  2.zlExseSvr_CardSquareTypes-获取消费卡类别服务接口
'三、临床域接口
'  1.zlCisSvr_GetPatPageInfByRange-根据一条范围条件，获取病人病案信息
'  2.zlCisSvr_GetPatiID-根据床号或住院号，获取病人ID
'  3.zlCIsSvr_GetPatiPageInfo-根据病人id或主页id,获取病案信息
'编制:刘兴洪
'日期:2019-11-02 15:49:29
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlPatiSvr_GetPatiInfsByRange(ByVal intQueryStatus As Integer, ByVal cllFilter As Variant, _
    ByRef cllPatiInfos_out As Collection, Optional ByVal str病人Ids As String, Optional ByRef str病区IDs As String, _
    Optional ByVal blnExpendInfo As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息集
    '入参:intQueryStatus-查询类型(0-仅门诊;1-在院 ;2-门诊及在院)
    '     cllFilter-过滤条件
    '     str病人Ids-病人ID
    '     rsPatiPage-主页信息
    '     str病区IDs-当前病区Ids
    '出参:cllPatiInfos_out-返回的数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTmp As Collection
    Dim cllData As Collection, cllSubData As Collection, cllContacts As Collection
    Dim i As Long
    
    Set cllPatiInfos_out = New Collection
    If zl_PatiSvr_GetPatiInfsByRange(intQueryStatus, cllFilter, cllData, str病人Ids, str病区IDs, blnExpendInfo) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then zlPatiSvr_GetPatiInfsByRange = True: Exit Function
    For i = 1 To cllData.count
        Set cllSubData = cllData(i): Set cllTmp = New Collection
        cllTmp.Add cllSubData("_pati_id"), "病人ID"
        
        cllTmp.Add NVL(cllSubData("_outpatient_num")), "门诊号"
        cllTmp.Add NVL(cllSubData("_fee_category")), "费别"
        cllTmp.Add NVL(cllSubData("_pati_name")), "姓名"
        cllTmp.Add NVL(cllSubData("_pati_sex")), "性别"
        cllTmp.Add NVL(cllSubData("_pati_age")), "年龄"
        cllTmp.Add NVL(cllSubData("_pati_birthdate")), "出生日期"
        cllTmp.Add NVL(cllSubData("_pati_idcard")), "身份证号"
        cllTmp.Add NVL(cllSubData("_pati_birthplace")), "出生地点"
        cllTmp.Add NVL(cllSubData("_pat_home_phno")), "家庭电话"
        cllTmp.Add NVL(cllSubData("_pat_home_addr")), "家庭地址"
        cllTmp.Add NVL(cllSubData("_phone_number")), "手机号"
        cllTmp.Add NVL(cllSubData("_emp_name")), "工作单位名称"
        cllTmp.Add NVL(cllSubData("_pati_identity")), "身份"
        cllTmp.Add NVL(cllSubData("_ocpt_name")), "职业"
        cllTmp.Add NVL(cllSubData("_pati_education")), "学历"
        cllTmp.Add NVL(cllSubData("_vcard_no")), "就诊卡号"
        cllTmp.Add NVL(cllSubData("_create_time")), "登记时间"
        cllTmp.Add NVL(cllSubData("_inpatient_num")), "住院号"
        cllTmp.Add NVL(cllSubData("_pati_bed")), "床号"
        cllTmp.Add NVL(cllSubData("_insurance_type")), "险类"

        If blnExpendInfo Then
            If CollectionExitsValue(cllData, "_contacts") Then
                Set cllContacts = cllData("_contacts")
                cllTmp.Add NVL(cllContacts("_name")), "联系人姓名"
                cllTmp.Add NVL(cllContacts("_phone")), "联系人电话"
            Else
                cllTmp.Add "", "联系人姓名"
                cllTmp.Add "", "联系人电话"
            End If
            cllTmp.Add NVL(cllSubData("_mdlpay_mode_name")), "医疗付款方式名称"
            cllTmp.Add NVL(cllSubData("_visit_time")), "就诊时间"
            cllTmp.Add NVL(cllSubData("_insurance_num")), "医保号"
        End If
        cllPatiInfos_out.Add cllTmp, "_" & cllTmp("病人ID")
    Next
    
    zlPatiSvr_GetPatiInfsByRange = True
End Function
Public Function zlPatiSvr_GetCardTypes(ByRef cllCardTypes_out As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡服务
    '入参:
    '出参:cllCardTypes_out-返回的卡集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 16:53:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetCardTypes = zl_PatiSvr_GetCardTypes(cllCardTypes_out)
End Function

Public Function zlPatiSvr_GetPatiID(ByVal cllFindCons As Collection, ByVal cllOtherFindCons As Collection, _
    ByRef cllPatiDatas_Out As Collection, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg As String, _
    Optional ByVal bln检查使用时间 As Boolean = True, Optional ByVal bln检查停用或挂失 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人ID信息
    '       cllFindCons-查找条件(array(接点名称,接点值))
    '                接点名称包含:卡类别ID,卡号,二维码,社区名,社区号)
    '       cllOtherFindCons-其他查找条件:array(查询的名称,查询的内容)
    '                   查询的名称:如:门诊号,就诊卡号，身份证号等
    '       blnNotShowErrMsg-不显示错误的提示信息
    '      bln检查使用时间-按卡类别查找有效
    '      bln检查停用或挂失-按卡类别查找有效
    '出参:strErrMsg-返回的错误信息
    '        lng病人ID-返回的病人ID
    '        cllPatiDatas_Out-返回病人信息数据
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-19 09:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlPatiSvr_GetPatiID = zl_PatiSvr_GetPatiID(cllFindCons, cllOtherFindCons, cllPatiDatas_Out, blnNotShowErrMsg, strErrMsg, bln检查使用时间, bln检查停用或挂失)
End Function

Public Function zlPatiSvr_GetPatiIDByComminuty(ByVal int社区序号 As Integer, ByVal str社区号 As String, _
                ByRef lng病人ID As Long, Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据社区号查找病人ID
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/15 17:50
    '---------------------------------------------------------------------------------------
    Dim cllFindCons As Collection, cllPati As Collection
    Set cllFindCons = New Collection
    cllFindCons.Add Array("社区序号", int社区序号)
    cllFindCons.Add Array("社区号", str社区号)
    zlPatiSvr_GetPatiIDByComminuty = zl_PatiSvr_GetPatiID(cllFindCons, Nothing, cllPati, blnNotShowErrMsg, strErrMsg)
    If cllPati Is Nothing Then Exit Function
    If cllPati.count = 0 Then Exit Function
    lng病人ID = cllPati("_pati_id")
End Function

Public Function zlCisSvr_PatiIsInhospital(ByVal lng病人ID As Long, ByRef blnInhospital As Boolean, _
                Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查病人是否在院就诊
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 14:35
    '---------------------------------------------------------------------------------------
    zlCisSvr_PatiIsInhospital = Zl_CisSvr_PatiIsInhospital(lng病人ID, blnInhospital, blnNotShowErrMsg, strErrMsg)
End Function

Public Function zlPatiSvr_GetPatiInfo(ByVal lng病人ID As Long, _
    ByVal cllOtherFindCons As Collection, ByRef cllPatiDatas_Out As Collection, _
    Optional ByVal int查询类型 As Integer = 0, _
    Optional ByVal bln包含家属 As Boolean, _
    Optional ByVal bln包含过敏药物 As Boolean, _
    Optional ByVal bln包含免疫信息 As Boolean, _
    Optional ByVal bln包含卡信息 As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人详细信息服务接口
    '入参:cllOtherFindCons-其他查找条件(array(查询名称,查询值)
    '             查询名称:病人IDS,姓名,性别,出生日期等,见query_cons_list[]列表中的描述部分
    '      int查询类型-0-基本;1-基本+联系人;2-所有
    '出参:cllPatiDatas_Out-返回病人信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 18:02:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetPatiInfo = zl_PatiSvr_GetPatiInfo(lng病人ID, cllOtherFindCons, cllPatiDatas_Out, int查询类型, bln包含家属, bln包含过敏药物, bln包含免疫信息, bln包含卡信息, blnNotShowErrMsg, strErrMsg)
End Function

Public Function zlPatiSvr_GetPatiExtendInfo(ByVal lng病人ID As Long, ByVal str信息名集 As String, ByRef cllPatiData_Out As Collection, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息从表信息服务接口
    '入参:str信息名集-多个用逗号分离,如：医学警示,联系人2,联系人3等
    '
    '出参:cllPatiData_Out-返回病人从表信息数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 20:10:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetPatiExtendInfo = zl_PatiSvr_GetPatiExtendInfo(lng病人ID, str信息名集, cllPatiData_Out, blnNotShowErrMsg, strErrMsg)
End Function
Public Function ZlPatisvr_GetPatiCardInfo(ByVal strCardTypeIDs As String, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从给定卡类别中检索指定病人持有有效卡的卡类别
    '入参: strCardTypeIDs 给定卡类别，多个用逗号分隔
    '返回:返回病人持有有效卡的卡类别，多个用逗号分隔
    '编制:刘兴洪
    '日期:2018-12-03 15:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    ZlPatisvr_GetPatiCardInfo = Zl_Patisvr_GetPatiCardInfo(strCardTypeIDs, lng病人ID)
End Function
Public Function zlPatiSvr_GetInsureByPatiID(lng病人ID As Long, Optional ByRef int险类_Out As Integer, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医保病人是否存在未结费用
    '入参:lng病人ID
    '     blnNotShowErrMsg-是否不显示错误信息
    '出参:int险类_Out-险类
    '     strErrMsg_out-返回的错误信息值
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 16:40:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetInsureByPatiID = zl_PatiSvr_GetInsureByPatiID(lng病人ID, int险类_Out, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlPatiSvr_CheckOutNoIsExist(ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查门诊号是否被他人使用
    ' 入参 : str门诊号-传入检查的门诊号
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlPatiSvr_CheckOutNoIsExist = zl_PatiSvr_CheckOutNoIsExist(lng病人ID, str门诊号, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlPatiSvr_PhoneNumberExist(ByVal lng病人ID As Long, ByVal str手机号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查手机号是否被他人使用
    ' 入参 :
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlPatiSvr_PhoneNumberExist = zl_PatiSvr_PhoneNumberExist(lng病人ID, str手机号, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlPatiSvr_CheckInsNoIsExist(ByVal str医保号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查医保号是否被他人使用
    ' 入参 :
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlPatiSvr_CheckInsNoIsExist = zl_PatiSvr_CheckInsNoIsExist(str医保号, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function


Public Function zlPatiSvr_GetPatiIdsByRange(ByVal strCondition As String, ByRef strPatiIds As String, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String, _
    Optional ByVal blnFindByFilter As Boolean, Optional ByVal cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件值获取符合条件的病人ID
    '入参:
    '   strCondition=可能是就诊卡号、身份证号、IC卡号、门诊号
    '   blnFindByFilter=True:按过滤条件(cllFilter)获取;False:按strCondition获取
    '   cllFilter=过滤条件:Array(Key,Value),Key:合同单位ID
    '出参:
    '返回:执行成功返回True，失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetPatiIdsByRange = zl_PatiSvr_GetPatiIdsByRange(strCondition, strPatiIds, _
        blnNotShowErrMsg, strErrMsg_Out, blnFindByFilter, cllFilter)
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'临床相关服务

Public Function zlCisSvr_GetPatPageInfByRange(ByVal intQueryStatus As Integer, ByVal cllFilter As Collection, Optional ByVal str病人Ids As String, Optional ByRef str病区IDs As String, _
    Optional ByRef cllPatiPages_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一条范围条件，获取病人病案信息
    '入参:intQueryStatus-查询类型(0-在院病人;1-出院病人;2-在院或出院 )
    '     cllFilter-过滤条件
    '     str病人Ids-多个用逗号:病人ID或病人ID:主页ID
    '     rsPatiPage-主页信息
    '     str病区IDs-当前病区Ids
    '出参:rsPatiPageInfo_Out-返回的病人信息集
    '     strPatiIds_Out-返回当前所涉及的病人IDs
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCisSvr_GetPatPageInfByRange = zl_CisSvr_GetPatPageInfByRange(intQueryStatus, cllFilter, str病人Ids, str病区IDs, cllPatiPages_Out)
    
End Function
Public Function zlCisSvr_GetPatiID(ByVal lng病区ID As Long, ByVal str床号 As String, ByVal str住院号 As String, Optional ByRef lng主页ID_out As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号及住院号获取病人ID
    '      str住院号-住院号
    '出参:lng主页ID-返回当前床号的主页ID
    '返回:成功返回病人ID,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection
    
    Set cllFilter = New Collection
    If Val(str住院号) <> 0 Then
        cllFilter.Add Array("住院号", str住院号)
    Else
        cllFilter.Add Array("病区ID", lng病区ID)
        cllFilter.Add Array("床号", str床号)
    End If
    zlCisSvr_GetPatiID = zl_CisSvr_GetPatiID(cllFilter, lng主页ID_out)
End Function

Public Function zlCIsSvr_GetPatiPageInfo(ByVal int查询类型 As Integer, ByVal str病人主页IDs As String, ByRef cllPatiPage_Out As Variant, _
    Optional ByRef bln仅取最后住院 As Boolean, Optional bln含婴儿信息 As Boolean, Optional ByRef bln含转科信息 As Boolean, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID或主页id信息来获取病案主页信息
    '入参:int查询类型-0-只获取基本信息;1-获取基本信息+扩展信息;2-仅获取取主页ID字段
    '     str病人主页IDs-两种格式:
    '           1.病人id1:主页id1,病人id2:主页id2...
    '           2.病人id1,病人id2,...病人idn
    '      bln仅取最后住院:主读取病人最后一次的病案,(str病人主页IDs第二种格式有效)
    '      bln含婴儿信息:是否包含婴儿信息
    '      bln含转科信息:是否包转科信息
    '出参:cllPatiPageInfo_Out-返回的病案信息集
    '     strErrMsg_Out-返回的错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    zlCIsSvr_GetPatiPageInfo = zl_CIsSvr_GetPatiPageInfo(int查询类型, str病人主页IDs, cllPatiPage_Out, bln仅取最后住院, bln含婴儿信息, bln含转科信息, blnNotShowErrMsg, strErrMsg_Out)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'三、费用相关服务
Public Function zlExseSvr_GetPatiSurplusInfo(ByVal str病人Ids As String, ByRef cllSurplusData_Out As Collection, _
    Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByVal strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人费用余额信息
    '入参:str病人Ids-病人ID,多个用逗号分离
    '     blnNotShowErrMsg-是否显示错误信息框,true-不显示;false-显示
    '出参:cllSurplusData_Out-返回病人信息集
    '     strErrMsg_out-不显示时，返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlExseSvr_GetPatiSurplusInfo = zl_ExseSvr_GetPatiSurplusInfo(str病人Ids, cllSurplusData_Out, blnNotShowErrMsg, strErrMsg_Out)
End Function
Public Function zlExseSvr_GetConsumerCardType(ByRef cllTypesData_out As Collection, Optional ByVal blnOnlyStart As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡类别信息服务接口
    '入参:blnOnlyStart-只获取启用的卡类别
    '出参:cllTypesData_out-返回卡类别信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:47:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlExseSvr_GetConsumerCardType = zl_ExseSvr_GetConsumerCardType(cllTypesData_out, blnOnlyStart)
End Function

Public Function zlPatiSvr_GetPatiAddrssInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
                ByVal str地址类别 As String, ByRef cllAddrList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人结构化地址信息
    ' 入参 : str地址类别:查询的地址类别：1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址；为0时表示查询所有类型的地址信息
    '        多个用逗号分隔，例如："3,4"
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim cllData As Collection, cllSubData As Collection, cllTmp As Collection
    Dim i As Integer
    
    On Error GoTo Errhand
    Set cllAddrList = New Collection
    If zl_PatiSvr_GetPatiAddrssInfo(lng病人ID, lng主页ID, str地址类别, cllData) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then zlPatiSvr_GetPatiAddrssInfo = True: Exit Function
    
    For i = 1 To cllData.count
        Set cllSubData = cllData(i): Set cllTmp = New Collection
            cllTmp.Add NVL(cllSubData("_pat_addr_type")), "地址类别"
            cllTmp.Add NVL(cllSubData("_pat_addr_state")), "省"
            cllTmp.Add NVL(cllSubData("_pat_addr_city")), "市"
            cllTmp.Add NVL(cllSubData("_pat_addr_county")), "县"
            cllTmp.Add NVL(cllSubData("_pat_addr_township")), "乡"
            cllTmp.Add NVL(cllSubData("_pat_addr_other")), "其他"
            cllTmp.Add NVL(cllSubData("_pat_region_code")), "区划代码"
        cllAddrList.Add cllTmp, "_" & cllTmp("地址类别")
    Next
    
    zlPatiSvr_GetPatiAddrssInfo = True
    Exit Function
Errhand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function ZlCissvr_GetPatiBaseInfoByPatiId(ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long = -1, _
                Optional ByRef cllPatiList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人结构化地址信息
    ' 入参 : lng病人ID；lng主页ID
    ' 出参 : cllPati(Collect):成员()
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim cllData As Collection, cllSubData As Collection, cllTmp As Collection
    Dim i As Integer
    
    On Error GoTo Errhand
    Set cllPatiList = New Collection
    If Zl_Cissvr_GetPatiBaseInfo(lng病人ID, lng主页ID, cllData) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then ZlCissvr_GetPatiBaseInfoByPatiId = True: Exit Function
    
    For i = 1 To cllData.count
        Set cllSubData = cllData(i): Set cllTmp = New Collection
            cllTmp.Add Val(NVL(cllSubData("_pati_id"))), "病人ID"
            cllTmp.Add NVL(cllSubData("_page_id")), "主页ID"
            cllTmp.Add NVL(cllSubData("_pati_name")), "姓名"
            cllTmp.Add NVL(cllSubData("_inpatient_num")), "住院号"
            cllTmp.Add NVL(cllSubData("_registration_time")), "入院时间"
            cllTmp.Add NVL(cllSubData("_adtd_time")), "出院时间"
            
        cllPatiList.Add cllTmp, "_" & cllTmp("病人ID") & "_" & cllTmp("主页ID")
    Next
    
    ZlCissvr_GetPatiBaseInfoByPatiId = True
    Exit Function
Errhand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

