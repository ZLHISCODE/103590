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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mlngModule As Long, mfrmMain As Object
Private mobjEInvoice As zlPublicExpense.clsPubEInvoice

Public Function Initialize(ByVal frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal byt场合 As Byte = 1, _
    Optional ByVal blnDeviceSet As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口部件
    '入参:
    '   frmMain：调用的主窗体
    '   lngModule：当前调用模块号
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '   blnDeviceSet：设备设置调用的初始化
    '出参:
    '返回:初始化成功返回true,否则返回False
    '说明:
    '   1.使用本部件前,必须先调用本接口进行初始化
    '   2.初始化接口,在HIS进入模块时调用(例如：进入收费管理界面)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExtend As String
    
    Set mfrmMain = frmMain: mlngModule = lngModule
    
    If mobjEInvoice Is Nothing Then
        On Error Resume Next
        Set mobjEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
        If Err <> 0 Then
            MsgBox "不存在可用的电子票据接口部件(zlPublicExpense.clsPubEInvoice)，请与系统管理员联系。详细的错误信息为:" & vbCrLf & Err.Description, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mobjEInvoice Is Nothing Then Exit Function
    
    Initialize = mobjEInvoice.zlInitialize(frmMain, byt场合, gcnOracle, glngSys, lngModule, blnDeviceSet, strExtend)
End Function

Public Function ReReadParaAndVar() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新读取参数及变量值
    '返回:成功返回true,否则返回False
    '说明：
    '   1.在参数设置后调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ReReadParaAndVar = mobjEInvoice.zlReReadParaAndVar
End Function

Public Function GetNextInvoiceNo(ByVal frmMain As Object, ByRef strInvoiceNO_Out As String, _
    Optional ByVal cllPati As Collection, Optional ByRef lng领用ID As Long, Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取下一张发票号
    '入参:
    '   frmMain -调用的主窗体
    '   cllPati-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称):目前无用，主要是以后扩展HIS管理票据
    '   lng领用ID-票据领用ID,以后HIS管理票据可能要用,暂无用处
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:
    '   strInvoiceNO_Out-返回下一张有效发票号
    '返回:获取成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetNextInvoiceNo = mobjEInvoice.zlGetNextInvoiceNo(frmMain, byt场合, strInvoiceNO_Out, cllPati, lng领用ID)
End Function

Public Function IsStartEInvoice(Optional ByVal intInsure As Integer, Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否启用了电子票据的
    '入参:
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '   intInsure：当前医保:0-代表普通病人;>0表示具体医保
    '出参:
    '返回:启用返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    IsStartEInvoice = mobjEInvoice.zlIsStartEInvoice(byt场合, intInsure)
End Function

Public Function EInvoiceSet(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:电子发票设备或参数配置接口
    '入参:
    '   frmMain：调用的主窗体
    '出参:
    '返回:配置成功返回true,否则返回False
    '说明:
    '    一般针对电子票据设备（主要是打印设备、电子票据网络相关配置等)一些配置进行设置,比如：打印机,端口号、IP、用户名、服务地址等的配置
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExtend As String
    
    EInvoiceSet = mobjEInvoice.zlEInvoiceSet(frmMain, strExtend)
End Function

Public Function IsHisManagerInvoice(Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否HIS管理纸质发票
    '入参:
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:
    '返回:HIS管理纸质的发票返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    IsHisManagerInvoice = mobjEInvoice.zlIsHisManagerInvoice(byt场合)
End Function

Public Function IsStartEInvoiceFromBalanceID(ByVal lng结算ID As Long, Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前结算是否启用了电子票据
    '入参:
    '   lng结算ID-结帐ID
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    IsStartEInvoiceFromBalanceID = mobjEInvoice.zlIsStartEInvoiceFromBalanceID(byt场合, lng结算ID)
End Function

Public Function CreateEInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, _
    Optional ByRef blnTurnPaperInvoice_Out As Boolean, _
    Optional ByVal blnOlnyCreateEInv As Boolean, Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:开具电子票据接口
    '入参:
    '   frmMain-调用主窗口
    '    cllSwapData-交易数据集
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    '   blnOlnyCreateEInv-是否只开具电子票据,true-表示只开具电子票据，False-根据参数来控制是否换开纸质发票
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:
    '    blnTurnPaperInvoice_Out-纸质发票换开是否成功：true-表示成功，false-表示失败
    '返回:开具电子票据成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CreateEInvoice = mobjEInvoice.zlCreateEInvoice(frmMain, byt场合, blnOlnyCreateEInv, cllSwapData, blnTurnPaperInvoice_Out)
End Function

Public Function GetNextPaperInvoice(ByVal frmMain As Object, ByVal cllPatiInfo As Collection, _
    ByRef lng领用ID As Long, Optional ByVal byt场合 As Byte = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一下张发票号
    '入参:
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '返回:发票号
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceNO As String
    
    On Error GoTo errHandle
    '根据票据领用读取
    If GetNextInvoiceNo(frmMain, strInvoiceNO, cllPatiInfo, lng领用ID, byt场合) = False Then Exit Function
    If strInvoiceNO = "" Then
        If frmInputBox.InputBox(frmMain, "发票号确认", "无法获取将要使用的发票号，" & _
                        vbCrLf & "请你输入换开将要使用的发票号码：", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
    Else
        If frmInputBox.InputBox(frmMain, "发票号确认", "请确认换开将要使用的发票号：", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
    End If
    GetNextPaperInvoice = strInvoiceNO
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function TurnPaperInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, ByVal lngEInvoiceID As Long, _
    ByVal strEInvoiceCode As String, ByVal strEInvoiceNo As String, ByRef strInvoiceNO_Out As String, _
    Optional ByVal int操作状态 As Integer, Optional ByVal strUseDate As String, Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:换开处理
    '入参:
    '   frmMain-调用主窗口
    '   cllSwapData-交易数据集
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    '   lngEInvoiceID-电子票据使用记录ID
    '   strEInvoiceCode-电子票据代码
    '   strEInvoiceNO-电子票据号码
    '   int操作状态:0-换开;1-重新换开;2-作废票据;3-回收票据
    '   strUseDate-票据使用时间,不传入时，也当前时间为准
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:
    '   strInvoiceNO_Out:返回本次的发票号信息,多个用逗号分离;返回空时，将以传入的发票号为准
    '返回:换开成功返回true,否则返回False
    '说明:
    '   1.主要是针对电子票据换开纸质票据业务
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceNO As String, lng领用ID As Long
    Dim cllPatiInfo As Collection, cllBalance As Collection
    
    If strInvoiceNO_Out <> "" Then
        If MsgBox("本次结算包含如下已换开纸质票据，是否回收这些票据？" & _
            vbCrLf & strInvoiceNO_Out, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    Set cllPatiInfo = cllSwapData("_PatiInfo")
    Set cllBalance = cllSwapData("_BalanceInfo")
    
    lng领用ID = cllBalance("_领用ID")
    strInvoiceNO = GetNextPaperInvoice(frmMain, cllPatiInfo, lng领用ID, byt场合)
    If strInvoiceNO = "" Then Exit Function '用户取消输入，不换开
    
    cllBalance.Remove "_发票号"
    cllBalance.Add strInvoiceNO, "_发票号"
    cllBalance.Remove "_领用ID"
    cllBalance.Add lng领用ID, "_领用ID"
    
    TurnPaperInvoice = mobjEInvoice.zlTurnPaperInvoice(frmMain, byt场合, cllSwapData, _
        lngEInvoiceID, strEInvoiceCode, strEInvoiceNo, strInvoiceNO_Out, int操作状态, strUseDate)
End Function

Private Function CancelEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, _
    ByVal cllPatiInfo As Collection, Optional ByVal str作废时间 As String, _
    Optional ByVal str操作 As String = "门诊退费", Optional ByVal byt场合 As Byte = 1, _
    Optional ByVal bln补结算 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废电子票据接口
    '入参:
    '   frmMain-调用主窗口
    '   lngEInvoiceID-需要作废的电子票据记录ID: 电子票据使用记录.ID
    '   cllPatiInfo - 病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类)
    '   str作废时间-yyyy-mm-dd hh24:mi:ss
    '   str操作：操作功能，提示使用
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:
    '返回:作废成功返回true,否则返回False
    '说明:
    '   1.对已经成生的电子票据进行作废
    '   2.返回的电子票据信息存储在“电子票据使用记录”中
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CancelEInvoice = mobjEInvoice.zlCancelEInvoice(frmMain, byt场合, lngEInvoiceID, cllPatiInfo, str作废时间, str操作, bln补结算)
End Function

Public Function PrintEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用打印票据接口
    '入参:
    '   lngEInvoiceID-需要作废的电子票据记录ID: 电子票据使用记录.ID
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PrintEInvoice = mobjEInvoice.zlPrintEInvoice(frmMain, lngEInvoiceID)
End Function

Public Function GetEInvoiceInforFromBalanceID(ByVal lng结算ID As Long, ByRef rsEInvoce_Out As ADODB.Recordset, _
    Optional ByVal int查询范围 As Integer = 1, Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算ID,获取指定电子票据信息
    '入参:
    '   lng结算ID-收费时为原始结帐ID，补结算时为原始结算ID
    '   int查询范围-0-所有;1-只查询有效的电子票据
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:rsEInvoce_Out-返回指定的电子票据信息数据
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetEInvoiceInforFromBalanceID = mobjEInvoice.zlGetEInvoiceInforFromBalanceID(byt场合, lng结算ID, rsEInvoce_Out, int查询范围)
End Function

Public Function CancelPaperInvoice(ByVal frmMain As Object, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, _
    ByVal lngEInvoiceID As Long, Optional ByVal str作废时间 As String, Optional ByVal byt场合 As Byte = 1, _
    Optional ByVal bln补结算 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废纸质票据接口
    '入参:
    '   frmMain-调用主窗口
    '   strInvoiceNO-发票号
    '   lngBalanceID-byt场合=2(预交)时：原预交ID；byt场合<>2(预交)时：原结帐ID
    '   lngEInvoiceID-电子票据使用记录ID
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:
    '返回:作废纸质票据成功返回true,否则返回False
    '说明:
    '   1.针对打印机打印卡纸、撕纸等造成的纸质票据损坏，而平台又换纸成功的，需要作废换卡的纸质票据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If MsgBox("本次结算包含如下已换开纸质票据，是否回收这些票据？" & _
        vbCrLf & strInvoiceNO, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    CancelPaperInvoice = mobjEInvoice.zlCancelPaperInvoice(frmMain, byt场合, strInvoiceNO, lngBalanceID, _
        lngEInvoiceID, UserInfo.编号, UserInfo.姓名, str作废时间, bln补结算)
End Function

Public Function GetTranPaperInvoiceModule(Optional ByVal byt场合 As Byte = 1) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取换开纸质发票方式
    '入参:
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:
    '返回:'0-不换开，1-自动换开，2-提示换开
    '---------------------------------------------------------------------------------------------------------------------------------------------
     GetTranPaperInvoiceModule = mobjEInvoice.zlGetTranPaperInvoiceModule(byt场合)
End Function

Public Function GetSwapCollectFromBalanceID(ByVal lng原结算ID As Long, _
    ByRef cllSwapData_Out As Collection, _
    Optional ByVal strInvoiceNO As String, Optional ByVal lng领用ID As Long, _
    Optional ByVal lng冲销ID As Long, Optional ByVal bln部分退费 As Boolean, _
    Optional ByVal strUseDate As String, Optional ByVal bln补结算 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算ID获取交易结算信息
    '入参:
    '   lng结算ID-收费时为原始结帐ID，补结算时为原始结算ID
    '   strInvoiceNO-当前发票号
    '   lng领用ID-发票的领用ID
    '出参:
    '   cllSwapData_Out-返回结算信息
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类),key(_节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, cllBalanceInfo As Collection
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, strInsureSql As String
    
    On Error GoTo ErrHandler
    If bln补结算 Then
        strWhere = " And b.结帐id In(Select 收费结帐ID From 费用补充记录 Where 结算ID=[1])"
    Else
        strWhere = " And b.结帐id = [1]"
    End If

    strSql = _
        " Select Max(a.病人id) As 病人ID, Max(a.主页id) As 主页ID, Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄," & _
        "        f_List2Str(Cast(Collect(a.No) As t_StrList)) As NO, Sum(a.结帐金额) As 结帐金额, Max(a.登记时间) As 登记时间" & _
        " From (Select a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.No, a.序号, Sum(a.结帐金额) As 结帐金额, Max(b.登记时间) As 登记时间" & _
        "        From 门诊费用记录 A, 门诊费用记录 B" & _
        "        Where Mod(a.记录性质, 10) = Mod(b.记录性质, 10) And a.No = b.No And a.序号 = b.序号" & strWhere & _
        "        Group By a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.No, a.序号" & _
        "        Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0) A"
    
    strInsureSql = "Select Max(险类) As 险类 From 保险结算记录 Where 性质 = 1 And 记录id = [1]"
    
    strSql = _
        " Select a.病人id, a.主页id, a.姓名, a.性别, a.年龄, m.门诊号, Nvl(n.住院号, m.住院号) As 住院号," & _
        "           a.No, a.结帐金额, a.登记时间, b.险类" & _
        " From (" & strSql & ") A, (" & strInsureSql & ") B, 病人信息 M, 病案主页 N" & _
        " Where a.病人id = m.病人id(+) And a.病人id = n.病人id(+) And a.主页id = n.主页id(+) And a.No Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "根据结帐ID构建电子票据信息", lng原结算ID)
    If rsTemp.EOF Then Exit Function

    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类)
    Set cllPati = New Collection
    cllPati.Add Val(Nvl(rsTemp!病人ID)), "_病人ID"
    cllPati.Add Val(Nvl(rsTemp!主页ID)), "_主页ID"
    cllPati.Add Nvl(rsTemp!姓名), "_姓名"
    cllPati.Add Nvl(rsTemp!性别), "_性别"
    cllPati.Add Nvl(rsTemp!年龄), "_年龄"
    cllPati.Add Nvl(rsTemp!门诊号), "_门诊号"
    cllPati.Add Nvl(rsTemp!住院号), "_住院号"
    cllPati.Add Val(Nvl(rsTemp!险类)), "_险类"

    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_发票号"
    cllBalanceInfo.Add lng原结算ID, "_结算ID"
    cllBalanceInfo.Add lng冲销ID, "_冲销ID"
    cllBalanceInfo.Add Nvl(rsTemp!NO), "_单据号"
    If strUseDate = "" Then
        cllBalanceInfo.Add Format(Nvl(rsTemp!登记时间), "yyyy-mm-dd HH:MM:SS"), "_登记时间"
    Else
        cllBalanceInfo.Add Format(strUseDate, "yyyy-mm-dd HH:MM:SS"), "_登记时间"
    End If
    cllBalanceInfo.Add IIf(bln补结算, 1, 0), "_是否补结算"
    cllBalanceInfo.Add IIf(bln部分退费, 1, 0), "_是否部分退款"
    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
    cllBalanceInfo.Add Val(Nvl(rsTemp!结帐金额)), "_结算金额"
    cllBalanceInfo.Add lng领用ID, "_领用ID"
    cllBalanceInfo.Add 1, "_结算类型"
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    
    GetSwapCollectFromBalanceID = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CancelEInvoiceFromBalance(ByVal frmMain As Object, _
    ByVal strNos As String, Optional ByVal intInsure As Integer, Optional ByVal bln补结算 As Boolean, _
    Optional ByVal str操作 As String = "门诊退费", Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算信息，发起对电子票据冲正
    '入参:
    '   cllPatiInfo：病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类)
    '   strNos：补结算时为结算单号，普通收费时为一次结算中的任一张单据号;多个时逗号分隔
    '   str操作：操作功能，提示使用
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '出参:objBalanceInfor(是否电子票据作废)改为true
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsEInvoice As ADODB.Recordset, strSql As String, str原结算ID As String
    Dim cllPati As Collection, strDate As String, strInvoice As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    str原结算ID = GetBalanceIDByFeeNo(strNos, bln补结算, byt场合)
    
    strSql = _
        "  Select a.ID, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号, a.是否换开, a.纸质发票号" & _
        "  From 电子票据使用记录 A,Table(F_Num2List([1])) J" & _
        "  Where a.结算ID=j.Column_Value And a.票种=[2] And a.记录状态=1" & _
        "   Order By a.登记时间"
    Set rsEInvoice = zlDatabase.OpenSQLRecord(strSql, "获取电子票据使用情况", str原结算ID, byt场合)
    If rsEInvoice.EOF Then CancelEInvoiceFromBalance = True: Exit Function
    
'    If IsStartEInvoice(intInsure, byt场合) = False Then
'        MsgBox "本次结算已开具了电子票据，但当前窗口不支持电子票据。" & vbCrLf & _
'            "请到支持电子票据的窗口进行" & str操作 & "。", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    Do While Not rsEInvoice.EOF
        If cllPati Is Nothing Then
            Set cllPati = New Collection
            cllPati.Add Val(Nvl(rsEInvoice!病人ID)), "_病人ID"
            cllPati.Add 0, "_主页ID"
            cllPati.Add Nvl(rsEInvoice!姓名), "_姓名"
            cllPati.Add Nvl(rsEInvoice!性别), "_性别"
            cllPati.Add Nvl(rsEInvoice!年龄), "_年龄"
            cllPati.Add Nvl(rsEInvoice!门诊号), "_门诊号"
            cllPati.Add Nvl(rsEInvoice!住院号), "_住院号"
            cllPati.Add intInsure, "_险类"
        End If
        
        If Val(Nvl(rsEInvoice!是否换开)) = 1 Then
            strInvoice = strInvoice & vbCrLf & Nvl(rsEInvoice!纸质发票号)
        End If
        
        rsEInvoice.MoveNext
    Loop
    rsEInvoice.MoveFirst
    
    If strInvoice <> "" Then
        strInvoice = Mid(strInvoice, 3)
        If MsgBox("本次结算包含如下已换开纸质票据，是否回收这些票据？" & _
            vbCrLf & strInvoice, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Do While Not rsEInvoice.EOF
        If CancelEInvoice(frmMain, Val(Nvl(rsEInvoice!ID)), cllPati, strDate, str操作, byt场合, bln补结算) = False Then Exit Function
        rsEInvoice.MoveNext
    Loop
    CancelEInvoiceFromBalance = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Terminate() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:终止接口
    '返回:终止成功返回true,否则返回False
    '说明:
    '   终止接口,在HIS退出模块时调用(例如：退出收费管理界面),可在本接口中完成第三方电子票据接口的资源释放、断开连接等时使用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Terminate = mobjEInvoice.zlTerminate()
End Function

Public Function SplitFeeNoByEInvoice(ByVal strNos As String, ByVal int记录性质 As Integer, _
    ByRef strEInvoiceNos As String, ByRef strNoEInvoiceNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:分解单据号，将启用电子票据和未启用电子票据的分开
    '入参:
    '   strNos-单据号，多个逗号分隔
    '出参:
    '   strNoEInvoiceNos-未启用电子票据的单据，格式：NO1,NO2,...
    '   strEInvoiceNos-启用电子票据的单据，格式：NO1,NO2,...
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim varNos As Variant, str结帐IDs As String
    
    On Error GoTo ErrHandler
    strEInvoiceNos = "": strNoEInvoiceNos = ""
    
    strSql = "Select /*+Cardinality(j,10)*/ a.NO, Max(b.是否电子票据) As 电子票据" & _
                " From 门诊费用记录 A, 病人预交记录 B, Table(F_Str2List([1])) J" & _
                " Where a.NO=j.Column_Value And a.结帐ID = b.结帐ID And b.记录性质 <> 1" & _
                "            And a.记录性质=[2] And a.记录状态 in (1,3)" & _
                " Group By a.NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "根据收费单获取原结帐ID", strNos, int记录性质)
    
    Do While Not rsTemp.EOF
        If Val(Nvl(rsTemp!电子票据)) = 0 Then
            strNoEInvoiceNos = strNoEInvoiceNos & "," & rsTemp!NO
        Else
            strEInvoiceNos = strEInvoiceNos & "," & rsTemp!NO
        End If
        rsTemp.MoveNext
    Loop
    If strEInvoiceNos <> "" Then strEInvoiceNos = Mid(strEInvoiceNos, 2)
    If strNoEInvoiceNos <> "" Then strNoEInvoiceNos = Mid(strNoEInvoiceNos, 2)
    
    SplitFeeNoByEInvoice = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBalanceIDByFeeNo(ByVal strNos As String, Optional ByVal bln补结算 As Boolean, _
    Optional ByVal byt场合 As Byte = 1) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用NO,获取原始的结帐ID
    '入参:
    '   strNos-单据号，多个逗号分隔
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡；非补结算时传入
    '返回:如果传入一个单据号，则返回原始的结帐ID；如果传入多个单据号，则返回多个结帐ID，以逗号分隔
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim str结帐IDs As String
    
    On Error GoTo ErrHandler
    If InStr(strNos, ",") = 0 Then
        If bln补结算 Then
            '部分退重结记录与原纪录相似
            strSql = _
                " Select 结算id As 结帐ID" & _
                " From (Select b.结算id From 费用补充记录 B" & _
                "        Where b.No = [1] And b.记录性质 = 1 And b.记录状态 In (1, 3)" & _
                "        Order By b.登记时间)" & _
                " Where Rownum < 2"
        Else
            strSql = "Select 结帐ID From 门诊费用记录 Where NO=[1] And 记录性质=[2] And 记录状态 in (1,3)"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "根据收费单获取原结帐ID", strNos, byt场合)
        If rsTemp.EOF Then Exit Function
        
        GetBalanceIDByFeeNo = Val(Nvl(rsTemp!结帐ID))
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
    GetBalanceIDByFeeNo = Mid(str结帐IDs, 2)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintNotice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印告知单
    '入参:frmMain-调用主窗体
    '     lngEInvoiceID-电子票据ID
    '     byt场合:：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PrintNotice = mobjEInvoice.zlPrintNotice(frmMain, byt场合, lngEInvoiceID)
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Set mfrmMain = Nothing
    Set mobjEInvoice = Nothing
End Sub
