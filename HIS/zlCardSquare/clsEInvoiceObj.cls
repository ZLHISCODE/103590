VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEInvoiceObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*********************************************************************************************************************************************
'功能:电子票据交易相关接口
'一、公共相关接口
'   1.zlReReadParaAndVar:重新读取参数值及变量值
'   2.zlInitialize:初始化接口
'   3.zlTerminate:终止接口
'   4.zlEInvoiceSet:电子发票设备或参数配置接口
'二、票据判断相关接口
'   1.zlIsStartEInvoice:判断是否启用了电子票据的
'   2.zlIsStartEInvoiceFromBalanceID:判断指定结算信息是否启用了电子发票
'   3.zlIsHisManagerInvoice:判断是否HIS管理纸质发票
'   4.zlIsStartEinvoicFromNO-根据费用据号，判断是否当前结帐单是否启用了电子票据的
'三、业务处理相关接口
'   1.zlCreateEInvoice-开具电子票据
'   2.zlCancelEInvoice-作废电子票据
'     zlCancelEInvoiceFromBalanceInfor-根据结帐对象，执行电子票据票据作废操作
'   3.zlPrintEInvoice-打印电子票据A4接口
'   4.zlTurnPaperInvoice-换开纸质票据
'   5.zlCancelPaperInvoice-作废纸质票据
'   6.zlGetNextInvoiceNo-获取下一张有效的发票号
'   7.zlGetTranPaperInvoiceModule-获取纸质发票的换开方式
'四、获取业务数据
'   1.zlGetEInvoiceInforFromBalanceID-根据结算ID,获取电子票据信息(返回记录集)
'编制:刘兴洪
'日期:2020-03-19 10:43:04
'*********************************************************************************************************************************************
Private mlngMoudle As Long, mstrPrivs As String
Private mfrmMain As Object
Private mobjEInvoice As Object
Private Function GetEInvoiceObject(ByRef objEInvoice As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取电子票据对应接对象
    '入参:
    '出参:objEinvoice-返回Invoice对象
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 14:32:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mobjEInvoice Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
        If Err <> 0 Then
            MsgBox "不存在可用的电子票据接口部件(zlPublicExpense.clsPubEInvoice)，请与系统管理员联系,详细的错误信息为:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrProductName
            Exit Function
        End If
        mobjEInvoice.IsExseRegion = False
    End If
    Set objEInvoice = mobjEInvoice
    GetEInvoiceObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetNextInvoiceNo(ByVal frmMain As Object, ByRef strInvoiceNO_Out As String, _
    Optional cllPati As Collection, Optional lng领用ID As Long, Optional ByVal byt场合 As Byte = 5) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一下张发票号
    '入参:byt场合：1-收费, 2-预交, 3-结帐, 4-挂号;5-就诊卡
    '     cllPati-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称):目前无用，主要是以后扩展HIS管理票据
    '     lng领用ID-票据领用ID,以后HIS管理票据可能要用,暂无用处
    '出参:strInvoiceNO_Out-返回下一张有效发票号
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-23 16:53:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objEInvoice As Object
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlGetNextInvoiceNo = objEInvoice.zlGetNextInvoiceNo(frmMain, byt场合, strInvoiceNO_Out, cllPati, lng领用ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReReadParaAndVar() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新读取参数及变量值
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-19 11:30:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlReReadParaAndVar = objEInvoice.zlReReadParaAndVar
End Function
Public Function zlInitialize(ByVal frmMain As Object, Optional ByVal byt场合 As Byte = 5, Optional ByVal blnDeviceSet As Boolean, Optional ByVal lngModule As Long = 1107) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口部件
    '入参:
    '    byt场合：1-收费,2-预交,3-结帐,4-挂号
    '    frmMain -调用的主窗体
    '    cnOracle -数据库连接对象
    '    lngSys -当前调用系统编号
    '    lngModule-当前调用模块号:
    '    blnDeviceSet -设备设置调用的初始化
    '出参:strExtend-扩展参数,暂留,现传为空
    '返回:初始化成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '    1.使用本部件前,必须先调用本接口进行初始化
    '    2.初始化接口,在HIS进入模块时调用(例如：进入收费管理界面),可在本接口中完成第三方电子票据打印接口的初始化、登录、打开连接等时使用
    '    3.模块号：1111=门诊挂号,1121=门诊收费,1124=保险补充结算,1137=住院结帐,1107=医疗卡发放管理;
    '           1103=预交款管理;1504=票据入库管理;1501=票据使用监控;2128=体检结帐管理
    '调用者:
    '   票据入库或票据领用?挂号?收费?结帐?补充结算?设备配置?体检团体结帐等
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    Dim strExtend As String
    Set mfrmMain = frmMain: mlngMoudle = lngModule
    
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlInitialize = objEInvoice.zlInitialize(frmMain, byt场合, gcnOracle, glngSys, lngModule, blnDeviceSet, strExtend)
End Function

Public Function zlTerminate() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:终止接口
    '返回:终止成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   终止接口,在HIS退出模块时调用(例如：退出收费管理界面),可在本接口中完成第三方电子票据接口的资源释放、断开连接等时使用
    '调用者:
    '    票据入库或票据领用,挂号,收费,结帐,补充结算,设备配置,体检团体结帐等
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlTerminate = objEInvoice.zlTerminate()
    Set objEInvoice = Nothing
End Function


Public Function zlIsStartEInvoice(Optional ByVal byt场合 As Byte = 5, Optional ByVal intInsure As Integer = 0, Optional int类别 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否启用了电子票据的
    '入参:byt场合：1-收费,2-预交,3-结帐,4-挂号
    '     intInsure-当前医保:0-代表普通病人;>0表示具体医保
    '     int类别-0-不区分类别;1-门诊;2-住院
    '出参:
    '返回:启用返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-19 14:32:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlIsStartEInvoice = objEInvoice.zlIsStartEInvoice(byt场合, intInsure, int类别)
    
    Set objEInvoice = Nothing
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlEInvoiceSet(ByVal frmMain As Object, Optional ByRef strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:电子发票设备或参数配置接口
    '入参:
    '    frmMain -调用的主窗体
    '    strExtend-1-当前仅支持开具电子票据不支持换开
    '出参:
    '    strExtend-扩展参数,暂留,现传为空
    '返回:配置成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '    一般针对电子票据设备（主要是打印设备、电子票据网络相关配置等)一些配置进行设置,比如：打印机,端口号、IP、用户名、服务地址等的配置
    '调用者:
    '   在各业务模块(如门诊、收费、结帐等)中的参数设置  : 设备配置—>电子票据设备, 点击“设置”按钮时,调用本接口.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlEInvoiceSet = objEInvoice.zlEInvoiceSet(frmMain, strExtend)
End Function

Public Function zlIsHisManagerInvoice(Optional ByVal byt场合 As Byte = 5) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否HIS管理纸质发票
    '入参:byt场合：1-收费, 2-预交, 3-结帐, 4-挂号;5-就诊卡
    '出参:
    '返回:HIS管理纸质的发票返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-19 14:32:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlIsHisManagerInvoice = objEInvoice.zlIsHisManagerInvoice(byt场合)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsStartEInvoiceFromBalanceID(ByVal lng结算ID As Long, Optional ByVal byt场合 As Byte = 5) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前结算是否启用了电子票据
    '入参:byt场合-1-收费, 2-预交, 3-结帐, 4-挂号;5-就诊卡
    '     lng结算ID-byt场合=2时,预交ID,否者结帐ID
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-23 17:17:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlIsStartEInvoiceFromBalanceID = objEInvoice.zlIsStartEInvoiceFromBalanceID(byt场合, lng结算ID)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 

Public Function zlCreateEInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, _
    Optional ByRef blnTurnPaperInvoice_Out As Boolean, Optional ByVal blnOlnyCreateEInv As Boolean = False, Optional ByVal byt场合 As Byte = 5) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:开具电子票据接口
    '入参:
    '    frmMain-调用主窗口
    '    byt场合-1-收费, 2-预交, 3-结帐, 4-挂号;5-就诊卡
    '    blnOlnyCreateEInv-是否只开具电子票据,true-表示只开具电子票据，False-根据参数来控制是否换开纸质发票
    '    cllSwapData-交易数据集
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    '出参:
    '    blnTurnPaperInvoice_Out-纸质发票换开是否成功：true-表示成功，false-表示失败
    '返回:开具电子票据成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '步聚:
    '   1.
    '调用者:
    '   1.门诊收费(部分退费)、结帐、挂号、 体检结帐、缴预交等
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlCreateEInvoice = objEInvoice.zlCreateEInvoice(frmMain, byt场合, blnOlnyCreateEInv, cllSwapData, blnTurnPaperInvoice_Out)
    
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlTurnPaperInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, ByVal lngEInvoiceID As Long, _
    ByVal strEInvoiceCode As String, ByVal strEInvoiceNO As String, ByRef strInvoiceNO_Out As String, _
        Optional ByVal int操作状态 As Integer = 0, Optional ByVal strUseDate As String = "", Optional ByVal int场合 As Byte = 5) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:换开处理
    '入参:
    '    frmMain-调用主窗口
    '    int场合-1-收费, 2-预交, 3-结帐, 4-挂号;5-医疗发卡
    '    strInvoiceNO-发票号
    '    lngEInvoiceID-电子票据使用记录ID
    '    strEInvoiceCode-电子票据代码
    '    strEInvoiceNO-电子票据号码
    '    strUseDate-票据使用时间,不传入时，也当前时间为准
    '    int操作状态:0-换开;1-重新换开;2-作废票据;3-回收票据
    '    cllSwapData-交易数据集
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)

    '出参:
    '   strInvoiceNO_Out:返回本次的发票号信息,多个用逗号分离;返回空时，将以传入的发票号为准
    '返回:换开成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.主要是针对电子票据换开纸质票据业务
    '调用者:
    '   1.门诊收费(部分退费)、结帐、挂号、 体检结帐、缴预交
    '   2.换开、重新换开
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlTurnPaperInvoice = objEInvoice.zlTurnPaperInvoice(frmMain, int场合, cllSwapData, lngEInvoiceID, strEInvoiceCode, strEInvoiceNO, strInvoiceNO_Out, int操作状态, strUseDate)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCancelEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, _
    ByVal cllPati As Collection, Optional ByVal str作废时间 As String, Optional byt场合 As Byte = 5, _
    Optional ByVal str操作 As String = "结帐作废") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废电子票据接口
    '入参:
    '    frmMain-调用主窗口
    '    int场合-1-收费, 2-预交, 3-结帐, 4-挂号;5-医疗发卡
    '    lngEInvoiceID-需要作废的电子票据记录ID: 电子票据使用记录.ID
    '    strExtend-扩展参数，暂留，现传为空
    '    cllPati-病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '    str作废时间-yyyy-mm-dd hh24:mi:ss
    '出参:
    '    strEInvoiceClientCode-电子票据开票点
    '    strSysSouceName_Out-系统来源名称
    '    strEInvoiceCode_out-电子票据代码(红票)
    '    strEInvoiceNO_out-电子票据号码(红票)
    '    strCheckCode_out-电子票据校验码(红票)
    '    strCreateTime_Out-电子票据生成时间(红票),格式:yyyy-mm-dd HH:MM:SS
    '    strEInvQRCode_Out-电子票据二维码图片数据(红票),base64加密存储,使用时，需解码
    '    strEInvUrl_Out-电子票据URL地址(红票)
    '    strEinvRemark_Out-电子票据的其他信息(红票)
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数，暂留，现传为空
    '返回:作废成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.对已经成生的电子票据进行作废
    '   2.返回的电子票据信息存储在“电子票据使用记录”中
    '调用者:
    '   1.门诊退费、结帐作废、退号、 体检结帐作废、预交退款等
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlCancelEInvoice = objEInvoice.zlCancelEInvoice(frmMain, byt场合, lngEInvoiceID, cllPati, str作废时间, str操作)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlPrintEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用打印票据接口
    '入参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-25 10:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function

    zlPrintEInvoice = objEInvoice.zlPrintEInvoice(frmMain, lngEInvoiceID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetEInvoiceInforFromBalanceID(ByVal lng结算ID As Long, ByRef rsEInvoce_Out As ADODB.Recordset, _
    Optional ByVal byt场合 As Byte = 5, Optional ByVal int查询范围 As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,获取指定电子票据信息
    '入参:byt场合-1-收费, 2-预交, 3-结帐, 4-挂号;5-医疗发卡
    '     int查询范围-0-所有;1-只查询有效的电子票据
    '
    '出参:rsEInvoce_Out-返回指定的电子票据信息数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-25 11:28:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlGetEInvoiceInforFromBalanceID = objEInvoice.zlGetEInvoiceInforFromBalanceID(byt场合, lng结算ID, rsEInvoce_Out, int查询范围)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlCancelPaperInvoice(ByVal frmMain As Object, ByVal strInvoiceNo As String, ByVal lngBalanceID As Long, ByVal cllPati As Collection, ByVal strNO As String, _
    ByVal lngEInvoiceID As Long, Optional ByVal str作废时间 As String = "", Optional ByVal byt场合 As Byte = 5) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废纸质票据接口
    '入参:
    '    frmMain-调用主窗口
    '    byt场合-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '    strInvoiceNO-发票号
    '    lngBalanceID-"1.byt场合=2(预交)时：原预交ID,2.byt场合<>2(预交)时：原结帐ID"
    '    lngEInvoiceID-电子票据使用记录ID
    '    strExtend -扩展参数，暂留，现传为空
    '    cllPati-病人信息集, 非费域必传此参数,格式（病人ID,姓名,性别,年龄,门诊号,住院号）,Key("_名称")
    '    strNO-单据号
    '出参:
    '返回:作废纸质票据成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.针对打印机打印卡纸、撕纸等造成的纸质票据损坏，而平台又换纸成功的，需要作废换卡的纸质票据
    '调用者:
    '   1.门诊收费、结帐、挂号、 体检结帐、缴预交等的作废功能
    '   2.重新换开及作废
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    Dim bln补结算 As Boolean, bln已作废电子票据 As Boolean, blnShowMsg As Boolean, strErrmsg_Out As String, int预交红票 As Integer
    On Error GoTo errHandle
    
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    
    blnShowMsg = True
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废纸质票据接口
    '入参:
    '    frmMain-调用主窗口
    '    byt场合-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '    strInvoiceNO-发票号
    '    lngBalanceID-"1.byt场合=2(预交)时：原预交ID"
    '    2.byt场合<>2(预交)时：原结帐ID"
    '    lngEInvoiceID-电子票据使用记录ID
    '    strOperatorCode-操作员编号
    '    strOperatorName-操作员姓名
    '    strExtend -扩展参数，暂留，现传为空
    '    bln补结算-是否费用补结算
    '    cllPati-病人信息集, 非费域必传此参数,格式（病人ID,姓名,性别,年龄,门诊号,住院号）,Key("_名称")
    '    strNO-单据号,非费域必传此参数,分别为收费、预交、结帐或就诊卡费用单据号
    '出参: bln作废电子票据-电子票据是否被同时作废
    zlCancelPaperInvoice = objEInvoice.zlCancelPaperInvoice(frmMain, byt场合, strInvoiceNo, lngBalanceID, _
      lngEInvoiceID, UserInfo.编号, UserInfo.姓名, str作废时间, bln补结算, bln已作废电子票据, blnShowMsg, strErrmsg_Out, int预交红票, cllPati, strNO)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetTranPaperInvoiceModule(Optional ByVal byt场合 As Byte = 5) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取换开纸质发票方式
    '入参:byt场合：1-收费, 2-预交, 3-结帐, 4-挂号;5-就诊卡
    '出参:
    '返回:'0-不换开，1-自动换开，2-提示换开
    '编制:刘兴洪
    '日期:2020-03-27 11:31:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
     zlGetTranPaperInvoiceModule = objEInvoice.zlGetTranPaperInvoiceModule(byt场合)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'Public Function zlGetEInvoiceSwapCollectFromBalanceID(ByVal lng结帐ID As Long, ByRef cllSwapData_Out As Collection, ByRef blnStartEInvoice_Out As Boolean, _
'    Optional ByVal strInvoiceNo As String, Optional ByVal lng领用ID As Long, Optional ByVal bln预交 As Boolean) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:获取结帐算相关信息
'    '入参:lng结帐ID-结帐ID(bln预交=true时，为预交ID)
'    '     strInvoiceNO-当前发票号
'    '     lng领用ID-发票的领用ID
'    '     bln预交-当前是预交款
'    '     strNo-单号号不为空时，按单据号处理
'    '出参:
'    '   blnStartEInvoice_Out-是否启用电子发票
'    '   cllSwapData_Out-返回结算信息
'    '      |-PatiInfo   Key="_PatiInfo"
'    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
'    '      |-BalanceInfo Key="_BalanceInfo"
'    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
'    '返回:成功返回true,否则返回False
'    '编制:刘兴洪
'    '日期:2020-03-27 17:21:08
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cllPati As Collection, cllBalanceInfo As Collection
'    Dim strSQL As String, rsTemp As ADODB.Recordset
'
'
'    On Error GoTo errHandle
'
'
'    If bln预交 Then
'        strSQL = " " & _
'        "   Select a.Id, a.No, a.病人id, a.主页id, Sum(A.金额) As 结帐金额, Max(A.预交电子票据) As 是否电子票据, " & _
'        "          Max(Nvl(d.姓名, c.姓名)) As 姓名, " & _
'        "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
'        "          max(M.险类) as 险类,to_char(max(A.收款时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(a.预交类别) as 结帐类型" & _
'        "   From  病人预交记录 A, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=3  and 记录ID=[1] ) M" & _
'        "   Where a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
'        "   Group By a.Id, a.No, a.病人id, a.主页id"
'
'    Else
'    strSQL = " " & _
'        "   Select a.结帐id As ID, b.No, a.病人id, a.主页id, Sum(a.冲预交) As 结帐金额, Max(a.是否电子票据) As 是否电子票据, Max(c.姓名) As 姓名, Max(c.性别) As 性别, " & _
'        "          Max(c.年龄) As 年龄, Max(c.住院号) As 住院号, Max(c.门诊号) As 门诊号, 0 As 险类, " & _
'        "          To_Char(Max(a.收款时间), 'yyyy-mm-dd hh24:mi:ss') As 收费时间, 1 as 结帐类型 " & _
'        "   From 病人预交记录 A, (Select  结帐id, From 住院费用记录 Where 结帐id = [1]) B, 病人信息 C  " & _
'        "   Where a.结帐id = b.结帐id And a.病人id = c.病人id(+)  And a.Id = [1] " & _
'        "   Group By a.结帐id, b.No, a.病人id, a.主页id"
'    End If
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID构建电子票据信息", lng结帐ID)
'    If rsTemp.EOF Then Exit Function
'    Set cllPati = New Collection
'
'    blnStartEInvoice_Out = Val(nvl(rsTemp!是否电子票据)) = 1
'
'    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
'    Set cllPati = New Collection
'    cllPati.Add Val(nvl(rsTemp!病人ID)), "_病人ID"
'    cllPati.Add Val(nvl(rsTemp!主页ID)), "_主页ID"
'
'    cllPati.Add nvl(rsTemp!姓名), "_姓名"
'    cllPati.Add nvl(rsTemp!性别), "_性别"
'    cllPati.Add nvl(rsTemp!年龄), "_年龄"
'    cllPati.Add nvl(rsTemp!门诊号), "_门诊号"
'    cllPati.Add nvl(rsTemp!住院号), "_住院号"
'    cllPati.Add Val(nvl(rsTemp!险类)), "_险类"
'
'    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID,合约单位结帐,结帐类型)
'    Set cllBalanceInfo = New Collection
'    cllBalanceInfo.Add strInvoiceNo, "_发票号"
'    cllBalanceInfo.Add lng结帐ID, "_结算ID"
'    cllBalanceInfo.Add 0, "_冲销ID"
'    cllBalanceInfo.Add nvl(rsTemp!NO), "_单据号"
'    cllBalanceInfo.Add Format(rsTemp!收费时间, "yyyy-mm-dd HH:MM:SS"), "_登记时间"
'    cllBalanceInfo.Add 0, "_是否补结算"
'    cllBalanceInfo.Add 0, "_是否部分退款"
'    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
'    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
'    cllBalanceInfo.Add RoundEx(Val(nvl(rsTemp!结帐金额)), 5), "_结算金额"
'    cllBalanceInfo.Add lng领用ID, "_领用ID"
'    cllBalanceInfo.Add IIf(Val(nvl(rsTemp!病人ID)) = 0, 1, 0), "_合约单位结帐"
'
'    If bln预交 Then
'        cllBalanceInfo.Add decode(Val(nvl(rsTemp!结帐类型)) = 0, 3, Val(nvl(rsTemp!结帐类型))), "_结算类型" '预交类别:1-门诊;2-住院 ;3-门诊和住院;
'    Else
'        cllBalanceInfo.Add 1, "_结算类型"  '结帐类型:1-门诊;2-住院 ;3-门诊和住院;
'    End If
'    Set cllSwapData_Out = New Collection
'    cllSwapData_Out.Add cllPati, "_PatiInfo"
'    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
'    zlGetEInvoiceSwapCollectFromBalanceID = True
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function
  
Public Function zlGetEInvoiceSwapCollectFromNo(ByVal strNos As String, ByRef cllSwapData_Out As Collection, ByRef blnStartEInvoice_Out As Boolean, _
    Optional ByVal strInvoiceNo As String, Optional ByVal lng领用ID As Long, Optional ByVal byt场合 As Byte = 5, _
    Optional ByRef dblEInvoice_Out As Double, Optional ByRef lng原结算ID_Out As Long, Optional str登记时间 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号，获取需要结算的信息
    '入参:lng结帐ID-结帐ID(bln预交=true时，为预交ID)
    '     strInvoiceNO-当前发票号
    '     lng领用ID-发票的领用ID
    '     strNo-单号号不为空时，按单据号处理
    '     byt场合-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '出参:
    '   dblEInvoice_Out-电子票据可以开具的金额
    '   lng原结算ID_Out-原结算ID
    '   blnStartEInvoice_Out-是否启用电子发票
    '   cllSwapData_Out-返回结算信息
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID,是否电子票据)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-27 17:21:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Zl_Exsesvr_Getbalanceinfo(strNos, cllSwapData_Out, blnStartEInvoice_Out, strInvoiceNo, lng领用ID, byt场合, dblEInvoice_Out, lng原结算ID_Out, str登记时间) = False Then Exit Function
    zlGetEInvoiceSwapCollectFromNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsStartEinvoicFromNO(ByVal strNO As String, Optional ByRef lng原结帐ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用单号判断是否启用了电子票据的
    '入参:strNo-结帐单号
    '出参：lng原结帐ID_Out-返回的结帐ID
    '返回:启用返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-31 17:48:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnStartEinvice As Boolean, dblMoney As Double
    On Error GoTo errHandle
    
    If Not Zl_Exsesvr_GetbalanceinfoFromNos(strNO, 5, dblMoney, lng原结帐ID_Out, blnStartEinvice) Then Exit Function
 
    zlIsStartEinvoicFromNO = blnStartEinvice
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetBalanceInforFromNO(ByVal strNO As String, ByRef dblTotalMoney_Out As Double, ByRef lng原结帐ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用单号原结帐ID
    '入参:strNo-费用单号
    '出参：lng原结帐ID_Out-返回的结帐ID
    '返回:启用返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-31 17:48:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnStartEinvice As Boolean
    On Error GoTo errHandle
    
    If Not Zl_Exsesvr_GetbalanceinfoFromNos(strNO, 5, dblTotalMoney_Out, lng原结帐ID_Out, blnStartEinvice) Then Exit Function
    zlGetBalanceInforFromNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCancelEInvoiceFromNO(ByVal frmMain As Object, ByVal strNO As String, Optional bln显示回收票据 As Boolean, Optional strInvoiceNO_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医疗卡单号发起对电子票据冲正
    '入参:strNO-单据号
    '     bln显示回收票据-是否显示需要回收的票据信息
    '出参:
    '     strInvoiceNo_Out-显示回收的发票号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-27 12:06:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsEInvoice As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim cllPati As Collection, strDate As String, lng结帐ID As Long
    Dim strSQL As String, lngEInvoiceID As Long
    
    On Error GoTo errHandle
         
    If Zl_Exsesvr_Geteinvoicesinfo(Me, strNO, cllPati, lngEInvoiceID, strInvoiceNO_Out) = False Then Exit Function
    
    
    If lngEInvoiceID = 0 Then
        zlCancelEInvoiceFromNO = True: Exit Function
    End If
    If bln显示回收票据 And strInvoiceNO_Out <> "" Then
        MsgBox "请注意回收以下发票号：" & vbCrLf & strInvoiceNO_Out, vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    End If
    
    
    '先检查数据是否合fc
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If zlCancelEInvoice(frmMain, lngEInvoiceID, cllPati, strDate) = False Then Exit Function
    zlCancelEInvoiceFromNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mfrmMain = Nothing
    Set mobjEInvoice = Nothing
End Sub

Public Function zlPrintNotice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, Optional ByVal byt场合 As Byte = 5) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印告知单
    '入参:frmMain-调用主窗体
    '     lngEInvoiceID-电子票据ID
    '     byt场合:：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-25 10:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlPrintNotice = objEInvoice.zlPrintNotice(frmMain, byt场合, lngEInvoiceID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExcuteErrDataEInvoice(ByVal frmMain As Object, ByVal lng结算ID As Long, Optional ByVal lng病人ID As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行异常的电子票据
    '入参:  lng结算ID-结帐ID
    '       lng病人ID-按病人ID查
    '出参:
    '返回:0-表示无异常数据
    '     1-表示存在异常数据，但处理成功!
    '     2-表示存在异常数据，但处理失败!
    '编制:刘兴洪
    '日期:2020-03-27 17:21:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object, rsTemp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errHandle
    
    strWhere = ""
    If lng病人ID <> 0 Then strWhere = strWhere & " And  病人ID=[1] "
    If lng结算ID <> 0 Then strWhere = strWhere & " And 业务id=[2] "
    
    If strWhere = "" Then
        '只处理本人的
        If lng结算ID <> 0 Then strWhere = strWhere & " And 操作员姓名=[3] "
    End If
     ' --记录标志    number(2)       0-开具电子票据;1-冲红电子票据;2-纸质票据;3-作废纸质票据
    
    strSQL = " " & _
    "  Select ID,  业务类型, 记录标志" & _
    "  From 电子票据异常记录 " & _
    "  where 操作场景=1 " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID构建电子票据信息", lng病人ID, lng结算ID, UserInfo.姓名)
    
    '0-表示无异常数据
    If rsTemp.EOF Then zlExcuteErrDataEInvoice = 0: Exit Function
    If GetEInvoiceObject(objEInvoice) = False Then zlExcuteErrDataEInvoice = 2: Exit Function
    With rsTemp
        Do While Not .EOF
            Select Case Val(nvl(rsTemp!记录标志))
            Case 0  '开具电子票据失败
                If MsgBox("你上次开具电子票据失败，你需要优先处理异常的电子票据信息才能继续操作,你是否需要处理?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then zlExcuteErrDataEInvoice = 2: Exit Function
            Case 1  '冲红电子票据;
                If MsgBox("你上次冲红电子票据失败，你需要优先处理异常的冲红电子票据信息才能继续操作,你是否需要处理?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then zlExcuteErrDataEInvoice = 2: Exit Function
            Case 2 '纸质票据;
                If MsgBox("你上次换开票据失败，你需要优先处理异常的票据信息才能继续操作,你是否需要处理?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then zlExcuteErrDataEInvoice = 2: Exit Function
            Case 3 '作废纸质票据
                If MsgBox("你上次作废纸抽票据失败，你需要优先处理异常的票据信息才能继续操作,你是否需要处理?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then zlExcuteErrDataEInvoice = 2: Exit Function
            Case Else
                MsgBox "不能识别的异常数据，请检查!", vbOKOnly + vbInformation, gstrSysName
                zlExcuteErrDataEInvoice = 2: Exit Function
            End Select
            '处理异常
            If objEInvoice.zlExcuteErrDataFromErrID(Val(nvl(rsTemp!ID))) = False Then zlExcuteErrDataEInvoice = 2: Exit Function
            rsTemp.MoveNext
        Loop
    End With
    zlExcuteErrDataEInvoice = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
  

Public Function zlExcuteErrDataEInvoiceFromErrID(ByVal frmMain As Object, ByVal lng异常ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据异常ID来处理异常数据
    '入参:  lng异常ID-异常ID
    '出参:
    '返回: 成功返回true, 否则返回False
    '编制:刘兴洪
    '日期:2020-03-27 17:21:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlExcuteErrDataEInvoiceFromErrID = objEInvoice.zlExcuteErrDataFromErrID(lng异常ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
  

