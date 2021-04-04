VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'   2.zlGetEinvoiceSwapCollect-根据结算信息，构建电子票据相关交易信息集
'   3.zlGetPatiCollectFromPatiObject-根据病人信息对象，返回病人信息集
'   4.zlGetEInvoiceSwapCollectFromDepositID-根据预交ID，构建电子票据相关交易信息集
'编制:刘兴洪
'日期:2020-03-19 10:43:04
'*********************************************************************************************************************************************
Private mlngMoudle As Long, mstrPrivs As String
Private mcnOracle As ADODB.Connection, mfrmMain As Object
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
    Optional cllPati As Collection, Optional lng领用ID As Long, Optional ByVal byt场合 As Byte = 2) As Boolean
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

Public Function zlInitialize(ByVal frmMain As Object, Optional ByVal byt场合 As Byte = 2, Optional ByVal blnDeviceSet As Boolean, Optional ByVal lngModule As Long = 1103) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口部件
    '入参:
    '    byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
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

Public Function zlIsStartEInvoice(Optional ByVal intInsure As Integer = 0, Optional ByVal int类别 As Integer = 0, _
                                                  Optional ByVal byt场合 As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否启用了电子票据的
    '入参:intInsure-当前医保:0-代表普通病人;>0表示具体医保
    '       int类别-:0-门诊及住院预交;1-门诊预交;2-住院预交
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
    '    strExtend-扩展参数,暂留,现传为空
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

Public Function zlIsHisManagerInvoice(Optional ByVal byt场合 As Byte = 2) As Boolean
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

Public Function zlIsStartEInvoiceFromBalanceID(ByVal lng结算ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前结算是否启用了电子票据
    '入参:lng结算ID-byt场合=2时,预交ID,否者结帐ID
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-23 17:17:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlIsStartEInvoiceFromBalanceID = objEInvoice.zlIsStartEInvoiceFromBalanceID(2, lng结算ID)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCreateEInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, _
    Optional ByRef blnTurnPaperInvoice_Out As Boolean, Optional ByVal blnOlnyCreateEInv As Boolean = False, _
    Optional ByVal byt场合 As Byte = 2, Optional ByVal int预交红票 As Integer, Optional ByVal blnShowMsg As Boolean = True) As Boolean
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
    zlCreateEInvoice = objEInvoice.zlCreateEInvoice(frmMain, byt场合, blnOlnyCreateEInv, cllSwapData, blnTurnPaperInvoice_Out, blnShowMsg, , , int预交红票)
    
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
Public Function zlTurnPaperInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, ByVal lngEInvoiceID As Long, _
    ByVal strEInvoiceCode As String, ByVal strEInvoiceNO As String, ByRef strInvoiceNO_Out As String, _
        Optional ByVal int操作状态 As Integer = 0, Optional ByVal strUseDate As String = "", Optional ByVal int场合 As Byte = 2, _
        Optional ByVal int红票 As Integer) As Boolean
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
    zlTurnPaperInvoice = objEInvoice.zlTurnPaperInvoice(frmMain, int场合, cllSwapData, lngEInvoiceID, strEInvoiceCode, strEInvoiceNO, strInvoiceNO_Out, int操作状态, strUseDate, , , int红票)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCancelEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, _
    ByVal cllPati As Collection, Optional ByVal str作废时间 As String, Optional byt场合 As Byte = 2, _
    Optional ByVal str操作 As String = "预交作废", Optional ByVal bln预交退款 As Boolean = True) As Boolean
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
    zlCancelEInvoice = objEInvoice.zlCancelEInvoice(frmMain, byt场合, lngEInvoiceID, cllPati, str作废时间, str操作, , , , bln预交退款)
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
    Optional ByVal byt场合 As Byte = 2, Optional ByVal int查询范围 As Integer = 1, Optional ByVal bln仅读取原始单据 As Boolean = False) As Boolean
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
    zlGetEInvoiceInforFromBalanceID = objEInvoice.zlGetEInvoiceInforFromBalanceID(byt场合, lng结算ID, rsEInvoce_Out, int查询范围, bln仅读取原始单据)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetEInvoiceIDFromDepositNO(ByVal strNO As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据预交NO,获取有效的电子票据ID
    '入参:
    '返回:返回电子票据ID
    '编制:刘兴洪
    '日期:2020-03-25 11:28:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsEInvoce As ADODB.Recordset
    On Error GoTo errHandle
     
    strSQL = "" & _
    "  Select max(ID) as ID " & _
    "  From 电子票据使用记录" & _
    "  Where 结算ID in (Select ID From 病人预交记录 where NO=[1] and 记录性质=1  And 记录状态 in (1,3) ) and 票种=2 And 记录状态=1 And nvl(原票据id,0)=0"
    Set rsEInvoce = zlDatabase.OpenSQLRecord(strSQL, "获取电子票据使用情况", strNO)
    zlGetEInvoiceIDFromDepositNO = Nvl(rsEInvoce!ID, 0)
    Exit Function
errHandle:
End Function

Public Function zlCancelPaperInvoice(ByVal frmMain As Object, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, _
    ByVal lngEInvoiceID As Long, Optional ByVal str作废时间 As String = "", Optional ByVal byt场合 As Byte = 2, Optional ByVal int红票 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废纸质票据接口
    '入参:
    '    frmMain-调用主窗口
    '    byt场合-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '    strInvoiceNO-发票号
    '    lngBalanceID-"1.byt场合=2(预交)时：原预交ID"
    '    2.byt场合<>2(预交)时：原结帐ID"
    '    lngEInvoiceID-电子票据使用记录ID
    '    strExtend -扩展参数，暂留，现传为空
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
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    
    zlCancelPaperInvoice = objEInvoice.zlCancelPaperInvoice(frmMain, byt场合, strInvoiceNO, lngBalanceID, _
      lngEInvoiceID, UserInfo.编号, UserInfo.姓名, str作废时间, , , , , int红票)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetTranPaperInvoiceModule(Optional ByVal byt场合 As Byte = 2) As Integer
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


Public Function zlGetEInvoiceSwapCollectFromDepositID(ByVal lng预交ID As Long, ByRef cllSwapData_Out As Collection, ByRef blnStartEInvoice_Out As Boolean, _
    Optional ByVal strInvoiceNO As String, Optional ByVal lng领用ID As Long, Optional lng冲销ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否正确
    '入参:lng预交ID-预交ID
    '     strInvoiceNO-当前发票号
    '     lng领用ID-发票的领用ID
    '     bln预交-当前是预交款
    '出参:
    '   blnStartEInvoice_Out-是否启用电子发票
    '   cllSwapData_Out-返回结算信息
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-27 17:21:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, cllBalanceInfo As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim rsTemp1  As ADODB.Recordset, dbl结算金额 As Double
    
    On Error GoTo errHandle
    
    strSQL = " " & _
    "   Select a.Id, a.No, a.病人id, a.主页id, Sum(A.金额) As 结帐金额, Max(A.预交电子票据) As 是否电子票据, " & _
    "          Max(Nvl(d.姓名, c.姓名)) As 姓名, " & _
    "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
    "          max(M.险类) as 险类,to_char(max(A.收款时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(a.预交类别) as 结帐类型" & _
    "   From  病人预交记录 A, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=3  and 记录ID=[1] ) M" & _
    "   Where a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
    "   Group By a.Id, a.No, a.病人id, a.主页id"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据预交ID构建电子票据信息", lng预交ID)
    If rsTemp.EOF Then Exit Function
    dbl结算金额 = roundEx(Val(Nvl(rsTemp!结帐金额)), 2)
    Set cllPati = New Collection
    If lng冲销ID <> 0 Then
         strSQL = " " & _
    "   Select a.Id, a.No, a.病人id, a.主页id, Sum(-1*A.冲预交) As 结帐金额, Max(A.预交电子票据) As 是否电子票据, " & _
    "          Max(Nvl(d.姓名, c.姓名)) As 姓名, " & _
    "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
    "          max(M.险类) as 险类,to_char(max(A.收款时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(a.预交类别) as 结帐类型" & _
    "   From  病人预交记录 A, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=3  and 记录ID=[1] ) M" & _
    "   Where a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
    "   Group By a.Id, a.No, a.病人id, a.主页id"
        Set rsTemp1 = zlDatabase.OpenSQLRecord(strSQL, "根据预交ID构建电子票据信息", lng冲销ID)
        If rsTemp1.EOF Then Exit Function
        dbl结算金额 = roundEx(Val(Nvl(rsTemp1!结帐金额)), 2)
    End If
    blnStartEInvoice_Out = Val(Nvl(rsTemp!是否电子票据)) = 1
     
    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    Set cllPati = New Collection
    cllPati.Add Val(Nvl(rsTemp!病人ID)), "_病人ID"
    cllPati.Add Val(Nvl(rsTemp!主页ID)), "_主页ID"
    
    cllPati.Add Nvl(rsTemp!姓名), "_姓名"
    cllPati.Add Nvl(rsTemp!性别), "_性别"
    cllPati.Add Nvl(rsTemp!年龄), "_年龄"
    cllPati.Add Nvl(rsTemp!门诊号), "_门诊号"
    cllPati.Add Nvl(rsTemp!住院号), "_住院号"
    cllPati.Add Val(Nvl(rsTemp!险类)), "_险类"

    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID,合约单位结帐,结帐类型)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_发票号"
    cllBalanceInfo.Add lng预交ID, "_结算ID"
    cllBalanceInfo.Add lng冲销ID, "_冲销ID"
    cllBalanceInfo.Add Nvl(rsTemp!NO), "_单据号"
    cllBalanceInfo.Add Format(rsTemp!收费时间, "yyyy-mm-dd HH:MM:SS"), "_登记时间"
    cllBalanceInfo.Add 0, "_是否补结算"
    cllBalanceInfo.Add 0, "_是否部分退款"
    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
    cllBalanceInfo.Add dbl结算金额, "_结算金额"
    cllBalanceInfo.Add lng领用ID, "_领用ID"
    cllBalanceInfo.Add decode(Val(Nvl(rsTemp!结帐类型)) = 0, 3, Val(Nvl(rsTemp!结帐类型))), "_结算类型" '预交类别:1-门诊;2-住院 ;3-门诊和住院;
    cllBalanceInfo.Add IIf(Val(Nvl(rsTemp!病人ID)) = 0, 1, 0), "_合约单位结帐"

    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    zlGetEInvoiceSwapCollectFromDepositID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetEinvoiceSwapCollect(ByVal objPati As clsPatientInfo, ByVal lng预交ID As Long, ByVal strNO As String, _
    ByVal dbMoney As Double, ByVal strDate As String, ByVal strInvoiceNO As String, ByRef cllSwapData_Out As Collection, _
    Optional ByVal lng冲销ID As Long, Optional lng领用ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前交易信息集
    '入参:objPati-病人信息对象
    '     strNo-预交单据号
    '     strDate-预交单据产生的时间
    '     strInvoiceNo-当前发票号
    '     lng领用ID-领用ID
    '     bln合约单位结账
    '出参:cllSwapData_Out-交易信息集
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-25 12:02:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiInfo As Collection, cllBalanceInfo As Collection, cllSwapData As Collection
    Dim i As Long, bytFunc As Byte
    
    On Error GoTo errHandle
    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    If zlGetPatiCollectFromPatiObject(objPati, cllPatiInfo) = False Then Exit Function
    
    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID,合约单位结帐,结帐类型)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_发票号"
    cllBalanceInfo.Add lng预交ID, "_结算ID"
    cllBalanceInfo.Add lng冲销ID, "_冲销ID"
    cllBalanceInfo.Add strNO, "_单据号"
    cllBalanceInfo.Add Format(strDate, "yyyy-mm-dd HH:MM:SS"), "_登记时间"
    cllBalanceInfo.Add 0, "_是否补结算"
    cllBalanceInfo.Add 0, "_是否部分退款"
    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
    cllBalanceInfo.Add dbMoney, "_结算金额"
    cllBalanceInfo.Add lng领用ID, "_领用ID"
    cllBalanceInfo.Add 2, "_结算类型" '结帐类型:1-门诊;2-住院 ;3-门诊和住院;
    cllBalanceInfo.Add 0, "_合约单位结帐"
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPatiInfo, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    zlGetEinvoiceSwapCollect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiCollectFromPatiObject(ByVal objPati As clsPatientInfo, ByRef cllPatiInfo_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人信息对象，构建病人信息集
    '入参:objPati-病人信息集
    '出参:cllPatiInfo_Out-病人信息集(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '返回:成功返回true,否则返回False
    '编制:焦博
    '日期:2020-06-28 14:44:46
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    
    If objPati Is Nothing Then Exit Function
    If objPati.病人ID = 0 Then Exit Function
    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    Set cllPatiInfo_Out = New Collection
    cllPatiInfo_Out.Add objPati.病人ID, "_病人ID"
    cllPatiInfo_Out.Add objPati.主页ID, "_主页ID"
    
    cllPatiInfo_Out.Add objPati.姓名, "_姓名"
    cllPatiInfo_Out.Add objPati.性别, "_性别"
    cllPatiInfo_Out.Add objPati.年龄, "_年龄"
    
    cllPatiInfo_Out.Add objPati.门诊号, "_门诊号"
    cllPatiInfo_Out.Add objPati.住院号, "_住院号"
    cllPatiInfo_Out.Add objPati.险类, "_险类"
    
    zlGetPatiCollectFromPatiObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCancelEInvoiceFromBalanceInfor(ByVal frmMain As Object, ByVal objPati As clsPatientInfo, _
                         ByVal strNO As String, Optional ByVal bln预交退款 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算信息，发起对电子票据冲正
    '入参:objPati-病人信息集
    '     frmMain-调用的主窗口
    '     objBalanceInfor-结帐信息集
    '出参:objBalanceInfor(是否电子票据作废)改为true
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-27 12:06:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsEInvoice As ADODB.Recordset
    Dim cllPati As Collection, strDate As String, lng结帐ID As Long
    Dim strSQL As String
    
    On Error GoTo errHandle

    strSQL = " " & _
    "Select a.Id, Max(a.预交电子票据) as 预交电子票据, Max(b.Id) As 电子票据id" & vbNewLine & _
    "From 病人预交记录 A, 电子票据使用记录 B" & vbNewLine & _
    "Where a.No = [1] And a.记录状态 In (1, 3) And a.Id = b.结算id(+) And b.票种(+) = 2 And b.记录状态(+) = 1 And a.记录性质=1" & vbNewLine & _
    "Group By a.Id"
     
    Set rsEInvoice = zlDatabase.OpenSQLRecord(strSQL, "获取预交信息", strNO)
    If rsEInvoice.EOF Then
        MsgBox "未找到单据号为【" & strNO & "】的原始预交数据，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Val(Nvl(rsEInvoice!电子票据ID)) = 0 Then
        zlCancelEInvoiceFromBalanceInfor = True: Exit Function
    End If
    
    If zlGetPatiCollectFromPatiObject(objPati, cllPati) = False Then Exit Function
    
    '先检查数据是否合法
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If zlCancelEInvoice(frmMain, Val(Nvl(rsEInvoice!电子票据ID)), cllPati, strDate, 2, , bln预交退款) = False Then Exit Function
    zlCancelEInvoiceFromBalanceInfor = True
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

Public Function zlIsStartEinvoicFromNO(ByVal strNO As String, Optional ByRef lng原预交ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用单号判断是否启用了电子票据的
    '入参:strNo-结帐单号
    '出参：lng原结帐ID_Out-返回的结帐ID
    '返回:启用返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-31 17:48:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = " " & _
    "Select a.Id as 预交ID, Max(a.预交电子票据) as 预交电子票据" & vbNewLine & _
    "From 病人预交记录 A" & vbNewLine & _
    "Where a.No = [1] And a.记录状态 In (1, 3) And a.记录性质= 1" & vbNewLine & _
    "Group By a.Id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断结算数据是否启用了电子票据的", strNO)
    If rsTemp.EOF Then Exit Function
    lng原预交ID_Out = Val(Nvl(rsTemp!预交ID))
    zlIsStartEinvoicFromNO = Val(Nvl(rsTemp!预交电子票据)) = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPrintNotice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, Optional ByVal byt场合 As Byte = 2) As Boolean
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


