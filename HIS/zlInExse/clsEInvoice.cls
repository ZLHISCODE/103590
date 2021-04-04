VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEInvoice"
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
'   4.zlIsStartEinvoiceBalanceNO-根据结帐单据号，判断是否当前结帐单是否启用了电子票据的
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
'   2.zlGetEinvoiceSwapCollectFromBalanceInfo-根据结算信息，构建电子票据相关交易信息集
'   3.zlGetPatiCollectFromPatiObject-根据病人信息对象，返回病人信息集
'   4.zlGetEInvoiceSwapCollectFromBalanceID-根据结算ID，构建电子票据相关交易信息集
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
    Optional cllPati As Collection, Optional lng领用ID As Long, Optional ByVal byt场合 As Byte = 3) As Boolean
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
Public Function zlInitialize(ByVal frmMain As Object, Optional ByVal byt场合 As Byte = 3, Optional ByVal blnDeviceSet As Boolean, Optional ByVal lngModule As Long = 1137) As Boolean
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


Public Function zlIsStartEInvoice(ByVal byt场合 As Byte, Optional ByVal intInsure As Integer = 0, Optional ByVal int类别 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否启用了电子票据的
    '入参:byt场合：1-收费,2-预交,3-结帐,4-挂号
    '     intInsure-当前医保:0-代表普通病人;>0表示具体医保
    '     int类别-对预交有效:0-不区分类别;1-门诊 2-住院
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

Public Function zlIsHisManagerInvoice(Optional ByVal byt场合 As Byte = 3) As Boolean
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

Public Function zlIsStartEInvoiceFromBalanceID(ByVal byt场合 As Byte, ByVal lng结算ID As Long) As Boolean
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
    Optional ByRef blnTurnPaperInvoice_Out As Boolean, Optional ByVal blnOlnyCreateEInv As Boolean = False, Optional ByVal byt场合 As Byte = 3) As Boolean
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
        Optional ByVal int操作状态 As Integer = 0, Optional ByVal strUseDate As String = "", Optional ByVal int场合 As Byte = 3) As Boolean
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
    ByVal cllPati As Collection, Optional ByVal str作废时间 As String, Optional byt场合 As Byte = 3, _
    Optional ByVal str操作 As String = "结帐作废", _
    Optional ByVal bln补结算 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废电子票据接口
    '入参:
    '    frmMain-调用主窗口
    '    int场合-1-收费, 2-预交, 3-结帐, 4-挂号;5-医疗发卡
    '    lngEInvoiceID-需要作废的电子票据记录ID: 电子票据使用记录.ID
    '    strExtend-扩展参数，暂留，现传为空
    '    cllPati-病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '    str作废时间-yyyy-mm-dd hh24:mi:ss
    '    bln补结算-是否费用补结算
    '出参:
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
    zlCancelEInvoice = objEInvoice.zlCancelEInvoice(frmMain, byt场合, lngEInvoiceID, cllPati, str作废时间, str操作, bln补结算)
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
    Optional ByVal byt场合 As Byte = 3, Optional ByVal int查询范围 As Integer = 1) As Boolean
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


Public Function zlCancelPaperInvoice(ByVal frmMain As Object, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, _
    ByVal lngEInvoiceID As Long, Optional ByVal str作废时间 As String = "", Optional ByVal byt场合 As Byte = 3) As Boolean
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
      lngEInvoiceID, UserInfo.编号, UserInfo.姓名, str作废时间)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetTranPaperInvoiceModule(Optional ByVal byt场合 As Byte = 3) As Integer
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


Public Function zlGetEInvoiceSwapCollectFromBalanceID(ByVal lng结帐ID As Long, ByRef cllSwapData_Out As Collection, ByRef blnStartEInvoice_Out As Boolean, _
    Optional ByVal strInvoiceNO As String, Optional ByVal lng领用ID As Long, Optional ByVal bln预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否正确
    '入参:lng结帐ID-结帐ID(bln预交=true时，为预交ID)
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
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    
    On Error GoTo errHandle
    If bln预交 Then
        strSql = " " & _
        "   Select a.Id, a.No, a.病人id, a.主页id, Sum(A.金额) As 结帐金额, Max(A.预交电子票据) As 是否电子票据, " & _
        "          Max(Nvl(d.姓名, c.姓名)) As 姓名, " & _
        "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
        "          max(M.险类) as 险类,to_char(max(A.收款时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(a.预交类别) as 结帐类型" & _
        "   From  病人预交记录 A, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=3  and 记录ID=[1] ) M" & _
        "   Where a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
        "   Group By a.Id, a.No, a.病人id, a.主页id"
    Else
        
        strSql = " " & _
        "   Select a.Id, a.No, a.病人id, a.主页id, Sum(b.冲预交) As 结帐金额, Max(b.是否电子票据) As 是否电子票据, " & _
        "          Max(decode(nvl(A.病人ID,0),0,A.原因,Nvl(d.姓名, c.姓名))) As 姓名, " & _
        "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
        "          max(M.险类) as 险类,to_char(max(A.收费时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(A.结帐类型) as 结帐类型" & _
        "   From 病人结帐记录 A, 病人预交记录 B, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=2  and 记录ID=[1] ) M" & _
        "   Where a.id=b.结帐ID and  a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
        "   Group By a.Id, a.No, a.病人id, a.主页id"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "根据结帐ID构建电子票据信息", lng结帐ID)
    If rsTemp.EOF Then Exit Function
    Set cllPati = New Collection
    
    blnStartEInvoice_Out = Val(nvl(rsTemp!是否电子票据)) = 1
     
    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    Set cllPati = New Collection
    cllPati.Add Val(nvl(rsTemp!病人ID)), "_病人ID"
    cllPati.Add Val(nvl(rsTemp!主页ID)), "_主页ID"
    
    cllPati.Add nvl(rsTemp!姓名), "_姓名"
    cllPati.Add nvl(rsTemp!性别), "_性别"
    cllPati.Add nvl(rsTemp!年龄), "_年龄"
    cllPati.Add nvl(rsTemp!门诊号), "_门诊号"
    cllPati.Add nvl(rsTemp!住院号), "_住院号"
    cllPati.Add Val(nvl(rsTemp!险类)), "_险类"

    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID,合约单位结帐,结帐类型)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_发票号"
    cllBalanceInfo.Add lng结帐ID, "_结算ID"
    cllBalanceInfo.Add 0, "_冲销ID"
    cllBalanceInfo.Add nvl(rsTemp!NO), "_单据号"
    cllBalanceInfo.Add Format(rsTemp!收费时间, "yyyy-mm-dd HH:MM:SS"), "_登记时间"
    cllBalanceInfo.Add 0, "_是否补结算"
    cllBalanceInfo.Add 0, "_是否部分退款"
    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
    cllBalanceInfo.Add RoundEx(Val(nvl(rsTemp!结帐金额)), 5), "_结算金额"
    cllBalanceInfo.Add lng领用ID, "_领用ID"
    If bln预交 Then
        cllBalanceInfo.Add decode(Val(nvl(rsTemp!结帐类型)) = 0, 3, Val(nvl(rsTemp!结帐类型))), "_结算类型" '预交类别:1-门诊;2-住院 ;3-门诊和住院;
        cllBalanceInfo.Add IIf(Val(nvl(rsTemp!病人ID)) = 0, 1, 0), "_合约单位结帐"
    Else
        cllBalanceInfo.Add decode(Val(nvl(rsTemp!结帐类型)) = 0, 3, Val(nvl(rsTemp!结帐类型))), "_结算类型"  '结帐类型:1-门诊;2-住院 ;3-门诊和住院;
        cllBalanceInfo.Add IIf(Val(nvl(rsTemp!病人ID)) = 0, 1, 0), "_合约单位结帐"
    End If
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    zlGetEInvoiceSwapCollectFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetEinvoiceSwapCollectFromBalanceInfo(ByVal objPati As clsPatiInfo, ByVal objBalanceInfor As clsBalanceInfo, _
    ByVal strInvoiceNO As String, ByRef cllSwapData_Out As Collection, Optional lng领用ID As Long, Optional ByVal bln合约单位结账 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前交易信息集
    '入参:objPati-病人信息对象
    '     objBalanceInfor-当前结算对象
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
    If zlGetPatiCollectFromPatiObject(objPati, objBalanceInfor.objInsure.险类, cllPatiInfo) = False Then Exit Function
    
    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID,合约单位结帐,结帐类型)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_发票号"
    cllBalanceInfo.Add objBalanceInfor.结帐ID, "_结算ID"
    cllBalanceInfo.Add objBalanceInfor.冲销ID, "_冲销ID"
    cllBalanceInfo.Add objBalanceInfor.结帐单据号, "_单据号"
    cllBalanceInfo.Add Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS"), "_登记时间"
    cllBalanceInfo.Add 0, "_是否补结算"
    cllBalanceInfo.Add 0, "_是否部分退款"
    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
    cllBalanceInfo.Add RoundEx(objBalanceInfor.当前结帐, 5), "_结算金额"
    cllBalanceInfo.Add lng领用ID, "_领用ID"
    cllBalanceInfo.Add IIf(objBalanceInfor.结算类型 = 1, 1, 2), "_结算类型" '结帐类型:1-门诊;2-住院 ;3-门诊和住院;
    cllBalanceInfo.Add IIf(bln合约单位结账, 1, 0), "_合约单位结帐"
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPatiInfo, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    zlGetEinvoiceSwapCollectFromBalanceInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiCollectFromPatiObject(ByVal objPati As clsPatiInfo, ByVal int险类 As Integer, ByRef cllPatiInfo_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人信息对象，构建病人信息集
    '入参:objPati-病人信息集
    '     int险类-险类
    '出参:cllPatiInfo_Out-病人信息集(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-27 11:44:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPatiTemp As clsPatientInfo
    
    On Error GoTo errHandle
    
    If objPati Is Nothing Then Exit Function
    
    If objPatiTemp Is Nothing Then
        Set objPatiTemp = New clsPatientInfo
    Else
        Set objPatiTemp = objPati.Patient
    End If
    
    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    Set cllPatiInfo_Out = New Collection
    cllPatiInfo_Out.Add objPatiTemp.病人ID, "_病人ID"
    cllPatiInfo_Out.Add objPatiTemp.主页ID, "_主页ID"
    
    cllPatiInfo_Out.Add objPatiTemp.姓名, "_姓名"
    cllPatiInfo_Out.Add objPatiTemp.性别, "_性别"
    cllPatiInfo_Out.Add objPatiTemp.年龄, "_年龄"
    cllPatiInfo_Out.Add objPatiTemp.门诊号, "_门诊号"
    cllPatiInfo_Out.Add objPatiTemp.住院号, "_住院号"
    cllPatiInfo_Out.Add int险类, "_险类"
    zlGetPatiCollectFromPatiObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCancelEInvoiceFromBalanceInfor(ByVal frmMain As Object, ByVal objPati As clsPatiInfo, ByRef objBalanceInfor As clsBalanceInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算信息，发起对电子票据冲正
    '入参:objPati-病人信息集
    '     objBalanceInfor-结帐信息集
    '出参:objBalanceInfor(是否电子票据作废)改为true
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-27 12:06:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsEInvoice As ADODB.Recordset
    Dim cllPati As Collection, strDate As String, lng结帐ID As Long
    Dim strSql As String
    
    On Error GoTo errHandle
    
    If objBalanceInfor.是否电子票据作废 Then zlCancelEInvoiceFromBalanceInfor = True: Exit Function
    strSql = " " & _
    "   Select C.Id, Max(b.是否电子票据) as 是否电子票据, Max(c.Id) As 电子票据id " & _
    "   From 病人结帐记录 A, 病人预交记录 B, 电子票据使用记录 C " & _
    "   Where a.No = [1] And a.记录状态 In (1, 3) And a.Id = b.结帐id And a.Id = c.结算id(+) And c.票种(+) = 3 And c.记录状态(+) = 1 " & _
    "   Group By C.Id"
     
    Set rsEInvoice = zlDatabase.OpenSQLRecord(strSql, "获取结帐信息", objBalanceInfor.结帐单据号)
    If rsEInvoice.EOF Then
        MsgBox "未找到结帐单【" & objBalanceInfor.结帐单据号 & "】的原始结帐数据，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    objBalanceInfor.是否电子票据 = Val(nvl(rsEInvoice!是否电子票据)) = 1
    If Val(nvl(rsEInvoice!电子票据Id)) = 0 Then
        objBalanceInfor.是否电子票据作废 = True
        
        
        zlCancelEInvoiceFromBalanceInfor = True: Exit Function
    End If
    If zlGetPatiCollectFromPatiObject(objPati, objBalanceInfor.objInsure.险类, cllPati) = False Then Exit Function
    
    '先检查数据是否合fc
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If zlCancelEInvoice(frmMain, Val(nvl(rsEInvoice!ID)), cllPati, strDate) Then
        objBalanceInfor.是否电子票据作废 = True
        zlCancelEInvoiceFromBalanceInfor = True: Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlIsStartEinvoiceBalanceNO(strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐单据号判断是否启用了电子票据的
    '入参:strNo-结帐单号
    '返回:启用返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-31 17:48:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSql = "Select Max(是否电子票据) as 是否电子票据 From 病人预交记录 Where 结帐ID IN(Select ID From 病人结帐记录 where NO=[1] and 记录状态 in (1,3))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "判断结算数据是否启用了电子票据的", strNO)
    zlIsStartEinvoiceBalanceNO = Val(nvl(rsTemp!是否电子票据)) = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function ZlGetBalanceIDByFeeNo(ByVal strNos As String, Optional ByVal bln补结算 As Boolean, _
    Optional ByRef byt场合 As Byte = 1) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用NO,获取原始的结帐ID
    '入参:
    '   strNos-单据号，多个逗号分隔
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡；非补结算时传入
    '传参：
    '   byt场合-补结算且单个NO时返回：1-收费，4-挂号
    '返回:如果传入一个单据号，则返回原始的结帐ID；如果传入多个单据号，则返回多个结帐ID，以逗号分隔
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim str结帐IDs As String
    
    On Error GoTo ErrHandler
    If InStr(strNos, ",") = 0 Then
        If bln补结算 Then
            '部分退重结记录与原纪录相似
            strSql = _
                " Select 结算id As 结帐ID,附加标志" & _
                " From (Select b.结算id, b.附加标志 From 费用补充记录 B" & _
                "        Where b.No = [1] And b.记录性质 = 1 And b.记录状态 In (1, 3)" & _
                "        Order By b.登记时间)" & _
                " Where Rownum < 2"
        Else
            strSql = "Select 结帐ID From 门诊费用记录 Where NO=[1] And 记录性质=[2] And 记录状态 in (1,3)"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "根据收费单获取原结帐ID", strNos, byt场合)
        If rsTemp.EOF Then Exit Function
        
        If bln补结算 Then byt场合 = IIf(Val(nvl(rsTemp!附加标志)) = 1, 4, 1)
        ZlGetBalanceIDByFeeNo = Val(nvl(rsTemp!结帐ID))
        Exit Function
    End If
    
    If bln补结算 Then
        '部分退重结记录与原纪录相似
        strSql = _
            " Select 结算id As 结帐id" & _
            " From (Select /*+Cardinality(j,10)*/" & _
            "         a.结算id, Row_Number() Over(Partition By a.记录性质, a.No Order By a.登记时间) As 组号" & _
            "        From 费用补充记录 A, Table(f_Str2List([1])) J" & _
            "        Where a.No = j.Column_Value And 记录性质 = 1 And 记录状态 In (1, 3))" & _
            " Where 组号 = 1"
    Else
        strSql = _
            " Select /*+Cardinality(j,10)*/ a.结帐ID" & _
            " From 门诊费用记录 A,Table(F_Str2List([1])) J" & _
            " Where a.NO=j.Column_Value And 记录性质=[2] And 记录状态 in (1,3)" & _
            " Group By a.结帐ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "根据收费单获取原结帐ID", strNos, byt场合)
    
    str结帐IDs = ""
    Do While Not rsTemp.EOF
        str结帐IDs = str结帐IDs & "," & rsTemp!结帐ID
        rsTemp.MoveNext
    Loop
    ZlGetBalanceIDByFeeNo = Mid(str结帐IDs, 2)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLCancelEInvoiceFromBalance_MZ(ByVal frmMain As Object, _
    ByVal lng原结算ID As Long, Optional ByVal intInsure As Integer, Optional ByVal bln补结算 As Boolean, _
    Optional ByVal str操作 As String = "门诊退费", Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算信息，发起对电子票据冲正
    '入参:
    '   cllPatiInfo：病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类)
    '   lng原结算ID：补结算时为原结算ID，普通收费时为原结帐ID
    '   str操作：操作功能，提示使用
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:objBalanceInfor(是否电子票据作废)改为true
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsEInvoice As ADODB.Recordset, strSql As String
    Dim cllPati As Collection, strDate As String, strInvoice As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        "  Select a.ID, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号, a.是否换开, a.纸质发票号" & _
        "  From 电子票据使用记录 A,Table(F_Num2List([1])) J" & _
        "  Where a.结算ID=j.Column_Value And a.票种=[2] And a.记录状态=1" & _
        "   Order By a.登记时间"
    Set rsEInvoice = zlDatabase.OpenSQLRecord(strSql, "获取电子票据使用情况", lng原结算ID, byt场合)
    If rsEInvoice.EOF Then ZLCancelEInvoiceFromBalance_MZ = True: Exit Function
    
'    If IsStartEInvoice(intInsure, byt场合) = False Then
'        MsgBox "本次结算已开具了电子票据，但当前窗口不支持电子票据。" & vbCrLf & _
'            "请到支持电子票据的窗口进行" & str操作 & "。", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    Do While Not rsEInvoice.EOF
        If cllPati Is Nothing Then
            Set cllPati = New Collection
            cllPati.Add Val(nvl(rsEInvoice!病人ID)), "_病人ID"
            cllPati.Add 0, "_主页ID"
            cllPati.Add nvl(rsEInvoice!姓名), "_姓名"
            cllPati.Add nvl(rsEInvoice!性别), "_性别"
            cllPati.Add nvl(rsEInvoice!年龄), "_年龄"
            cllPati.Add nvl(rsEInvoice!门诊号), "_门诊号"
            cllPati.Add nvl(rsEInvoice!住院号), "_住院号"
            cllPati.Add intInsure, "_险类"
        End If
        
        If Val(nvl(rsEInvoice!是否换开)) = 1 Then
            strInvoice = strInvoice & vbCrLf & nvl(rsEInvoice!纸质发票号)
        End If
        
        rsEInvoice.MoveNext
    Loop
    rsEInvoice.MoveFirst
    
'    If strInvoice <> "" Then
'        strInvoice = Mid(strInvoice, 3)
'        If MsgBox("本次结算包含如下已换开纸质票据，是否回收这些票据？" & _
'            vbCrLf & strInvoice, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
'    End If
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Do While Not rsEInvoice.EOF
        If zlCancelEInvoice(frmMain, Val(nvl(rsEInvoice!ID)), cllPati, strDate, byt场合, str操作, bln补结算) = False Then Exit Function
        rsEInvoice.MoveNext
    Loop
    ZLCancelEInvoiceFromBalance_MZ = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mfrmMain = Nothing
    Set mobjEInvoice = Nothing
End Sub

