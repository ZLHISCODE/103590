VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEinvoice_BS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsEInvoiceModule
'*********************************************************************************************************************************************
'电子票据相关接口
'一、电子票据公共接口类:
'   1.zlInitialize :初始化部件
'   2.zlTerminate:终止接口
'   3.zlEInvoiceSet:电子发票设备或参数配置接口
'   4.zlEInvoiceManager:电子票据管理接口
'二、票据库存管理相关接口(暂不处理)
'   1.zlGetStockInvoices:获取当前库存有效的发票接口
'   2.zlOutStockInvoices:下库存票据接口
'   3.zlReturnStockInvoices :票据申退接口
'   4.zlCancelBlankInvoice:报损空白票据接口
'三、电子票据相关接口
'   1.zlCreateEInvoice:开具电子票据接口
'   2.zlCheckCancelEInvoiceValied:检查作废电子票据的合法性
'   3.zlCancelEInvoice:作废电子票据接口
'   4.zlPrintEInvoice:打印电子票据(A4纸)
'   5.zlSendEinvoiceMsg:发送电子票据通知信息
'四、纸质票据相关接口
'   1.zlGetNextInvoiceNo:获取下一张有效发票号
'   2.zlTurnPaperInvoice:换开纸质发票接口
'   3.zlPrintPaperInvoice:打印纸质发票接口
'   4.zlCancelPaperInvoice:作废纸质票据接口
'五、对象获取
'   1.zlGetParaFrom() :获取参数配置界面
'编制:刘兴洪
'日期:2020-03-03 14:11:42
'*********************************************************************************************************************************************
Private mfrmMain As Object, mlngSys As Long, mlngModule As Long
Private mstrOperatorCode As String, mstrOperatorName As String, mstrDBUser As String
Private mstrEInvoiceClientCode As String '发票开票点编号
Private mlngCurEInvoiceID As Long  '当前票据ID
Private mstrCurReason As String '当前票据作废原因
Private mblnInitPrintSet As Boolean
'博思打印的服务地址和服务端口
Private Const mstrPrintIp As String = "127.0.0.1"
Private Const mstrPrintPort As String = "13526"

Private mcnOracle As ADODB.Connection
Private mfrmParaForm As Object
Private mobjJson As clsJson
Private mfrmEInvoiceManage As frmEInvoiceManage

Private Function clsEInvoiceModule_zlGetEInvoiceProductName(ByVal frmMain As Object, strProductName As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取电子票据系统名称,缺省从电子票据类别中获取，如果用户有不同的设置需自行调整
    '入参:
    '    frmMain-调用主窗口
    '    strExtend-扩展参数，暂留，现传为空
    '出参:
    '    strProductName-电子票据系统名称
    '    strExtend-扩展参数，暂留，现传为
    '返回:获取成功返回true,否则返回False
    '编制:李南春
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.在开具电子票据时，需要获取电子票据来源
    '调用者:
    '   1.门诊收费(部分退费)、结帐、挂号、 体检结帐、缴预交等开具电子票据时
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand

    If gstrProductName = "" Then
        strSql = "Select 名称 From 电子票据类别 Where 是否启用 = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取电子票据系统来源")
        If Not rsTmp.EOF Then
            gstrProductName = NVL(rsTmp!名称)
        End If
    End If
    strProductName = gstrProductName
    clsEInvoiceModule_zlGetEInvoiceProductName = True
    Exit Function
ErrHand:
    Err.Clear
End Function

Private Function clsEInvoiceModule_zlGetParaFrom() As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取参数配置界面
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '    1.主要是在导航台管理中调用电子票据管理,类似于医保部件中的“保险账户管理”模块
    '    2.该接口功能：对电子票据的对账、作废、查询等
    '调用者:
    '   1.新增模块："1144-电子票据管理"
    '   2.入口程序：由部件zL9CashBill调用本接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmParaForm Is Nothing Then
        Set mfrmParaForm = New frmParaSet_BS
        Load mfrmParaForm
    End If
    Set clsEInvoiceModule_zlGetParaFrom = mfrmParaForm
End Function

Private Function clsEInvoiceModule_zlInitialize(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection, _
                ByVal lngSys As Long, ByVal lngModule As Long, ByVal strOperatorCode As String, _
                ByVal strOperatorName As String, ByVal strDBUser As String, Optional ByVal blnDeviceSet As Boolean, Optional ByRef strEInvoiceClientCode As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口部件
    '入参:
    '    frmMain -调用的主窗体
    '    cnOracle -数据库连接对象
    '    lngSys -当前调用系统编号
    '    lngModule-当前调用模块号:
    '    strOperatorCode -操作员编号
    '    strOperatorName -操作员姓名
    '    strDBUser -数据库用户名
    '    blnDeviceSet -设备设置调用的初始化
    '    strExtend-扩展参数,暂留,现传为空
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
    Set mcnOracle = cnOracle: Set mfrmMain = frmMain
    mlngSys = lngSys: mlngModule = lngModule: mstrDBUser = strDBUser
    mstrOperatorCode = strOperatorCode: mstrOperatorName = strOperatorName
    mstrEInvoiceClientCode = strEInvoiceClientCode
    
    If Not blnDeviceSet And strEInvoiceClientCode = "" Then
        MsgBox "未传入开票点编号，无法进行接口初始化！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If zlInitIFacePara(lngSys, strOperatorCode, strOperatorName) = False Then Exit Function
    Call clsEInvoiceModule_zlGetEInvoiceProductName(frmMain, gstrProductName)
    clsEInvoiceModule_zlInitialize = True
End Function

Private Function clsEInvoiceModule_zlPrintEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印电子票据接口
    '入参:
    '    frmMain-调用主窗口
    '    lngEInvoiceID-电子票据使用记录.ID
    '    strExtend-扩展参数,暂留,现传为空
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数,暂留,现传为空\
    '返回:打印电子票据成功返回true,否则返回False
    '编制:李南春
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.向医疗电子票据管理平台发起电子票据打印请求
    '调用者:
    '   1.门诊收费(部分退费)、结帐、挂号、 体检结帐、缴预交等
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson_In As String, strJson_Out As String
    Dim blnResult As Boolean
    
    If Not InitPrintSet(strErrMsg_Out) Then Exit Function  '初始化打印
    '组织业务数据
    If mdlEInvoice_BS.zlGetJson_PrintEInvoice(lngEInvoiceID, strJson_In, strErrMsg_Out) = False Then Exit Function
    If Not PrintEInvoiceRequest("printElectBill", strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    clsEInvoiceModule_zlPrintEInvoice = True
End Function

Private Function clsEInvoiceModule_zlTerminate() As Boolean
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
    Set mfrmMain = Nothing
    Set mcnOracle = Nothing
    clsEInvoiceModule_zlTerminate = True
End Function
Private Function clsEInvoiceModule_zlEInvoiceSet(ByVal frmMain As Object, Optional strExtend As String) As Boolean
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
    clsEInvoiceModule_zlEInvoiceSet = True
End Function
Private Function clsEInvoiceModule_zlGetStockInvoices(ByVal frmMain As Object, ByVal bytOccasion As Byte, ByVal bytInvoiceType As Byte, strBatchNo_out As String, strPrefix_out As String, strStartInvoice_out As String, strEndInvoice_out As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前库存有效的发票接口
    '入参:
    '    frmMain-调用主窗口
    '    bytOccasion-场合:0-票据入库发起;1-票据领用发起
    '    bytInvoiceType-发票种类:1-收费,2-预交,3-结帐,4-挂号
    '    strExtend-扩展参数,暂留,现传为空
    '出参:
    '    strBatchNo_out-批次
    '    strPrefix_out-前缀文本
    '    strStartInvoice_out-开始发票号
    '    strEndInvoice_out-终止发票号
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数,暂留,现传为空
    '返回:获取当前库存成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.只返回当前有效发票信息.
    '   2.接口返回true时,如果没有返回开始发票号或终止发票号,则允许手工录入具体的发票信息,否则禁止入库或领用
    '   3.如果未有此接口,则缺省为True;
    '   4.如果从三方平台接口中返回多条有效的票据,请在接口中弹出选择器,选择一条合法的入库票据
    '调用者:
    '   1.票据入库管理:入库
    '   2.票据使用监控:领用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_zlGetStockInvoices = True
End Function
Private Function clsEInvoiceModule_zlOutStockInvoices(ByVal frmMain As Object, ByVal lngID As Long, ByVal bytOccasion As Byte, ByVal bytInvoiceType As Byte, ByVal strUserName As String, ByVal strRemark As String, ByVal strBatchNo As String, ByVal strStartInvoice As String, ByVal strEndInvoice As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:下库存票据接口
    '入参:
    '    frmMain-调用主窗口
    '    lngID-领用ID或入库ID
    '    bytOccasion-场合:0-票据入库发起;1-票据领用发起
    '    bytInvoiceType -发票种类:1-收费,2-预交,3-结帐,4-挂号
    '    strUserName-使用人
    '    strRemark-备注
    '    strBatchNo-批次
    '    strStartInvoice-开始发票号
    '    strEndInvoice-终止发票号
    '    strExtend -扩展参数,暂留,现传为空
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend -扩展参数,暂留,现传为空
    '返回:下库存成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.本接口主要是针对三方平台票据库存的下库存,HIS表示入库或领用
    '调用者:
    '   1.票据入库管理:入库
    '   2.票据使用监控:领用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_zlOutStockInvoices = True
End Function
Private Function clsEInvoiceModule_zlReturnStockInvoices(ByVal frmMain As Object, ByVal lngID As Long, ByVal bytOccasion As Byte, ByVal bytInvoiceType As Byte, ByVal strReturnerName As String, ByVal strRemark As String, ByVal strBatchNo As String, ByVal strStartInvoice As String, ByVal strEndInvoice As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据申退接口
    '入参:
    '    frmMain-调用主窗口
    '    lngID-领用ID或入库ID
    '    bytOccasion-场合:0-票据入库发起;1-票据领用发起
    '    bytInvoiceType-发票种类:1-收费,2-预交,3-结帐,4-挂号
    '    strReturnerName-申退人
    '    strRemark-申退原因
    '    strBatchNo-批次
    '    strStartInvoice-开始发票号
    '    strEndInvoice-终止发票号
    '    strExtend -扩展参数,暂留,现传为空
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend -扩展参数,暂留,现传为空
    '返回:票据申退成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.票据申退主要提供各个窗口由于票多领、错领了,回退到三方平台的库存
    '调用者:
    '   1.票据入库管理:申退
    '   2.票据使用监控:申退
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    clsEInvoiceModule_zlReturnStockInvoices = True
End Function


Private Function clsEInvoiceModule_zlCancelBlankInvoice(ByVal frmMain As Object, ByVal lngID As Long, _
                ByVal bytInvoiceType As Byte, ByVal strAuthorName As String, ByVal strReason As String, _
                ByVal strBatchNo As String, ByVal strStartInvoice As String, ByVal strEndInvoice As String, _
                ByVal strHappenTime As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:报损空白票据接口
    '入参:
    '    frmMain-调用主窗口
    '    lngID-领用ID
    '    bytInvoiceType-发票种类:1-收费,2-预交,3-结帐,4-挂号
    '    strAuthorName-作废人
    '    strReason-作废原因
    '    strBatchNo-批次
    '    strStartInvoice-开始发票号
    '    strEndInvoice-终止发票号
    '    strHappenTime -发生时间:yyyy-mm-dd HH:MM:SS
    '    strExtend -扩展参数,暂留,现传为空
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend -扩展参数,暂留,现传为空
    '返回:报损空白票据成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.主要是针对收费员手头的空白票据未使用但因其他原因造成发票的破损的处理操作
    '调用者:
    '   1.票据入库管理:报损
    '   2.票据使用监控:报损
    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strJson_In As String, strJson_Out As String
'    Dim blnResult As Boolean
'
'    '组织业务数据
'    If mdlEInvoice_BS.zlGetJson_CancelBlankInvoice(strBatchNo, strStartInvoice, strEndInvoice, mstrEInvoiceClientCode, _
'                    strAuthorName, strReason, strHappenTime, strJson_In, strErrMsg_Out) = False Then Exit Function
'    If Not EInvoiceRequest("invalidBlankBillNo", strJson_In, blnResult, strJson_Out) Then
'        strErrMsg_Out = strJson_Out
'        Exit Function
'    End If
    clsEInvoiceModule_zlCancelBlankInvoice = True
End Function
Private Function clsEInvoiceModule_zlCreateEInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, _
    ByVal lngEInvoiceID As Long, ByVal lngBalanceID As Long, ByVal lngWriteOffID As Long, _
    ByRef strEInvoiceClientCode As String, ByRef strSysSouceName_Out As String, _
    ByRef strEInvoiceCode_out As String, ByRef strEInvoiceNo_Out As String, ByRef strCheckCode_out As String, ByRef strCreateTime_Out As String, _
    ByRef strEInvQRCode_Out As String, ByRef strEInvUrl_Out As String, ByRef strEInvUrl1_Out As String, ByRef strEinvRemark_Out As String, ByRef strErrMsg_Out As String, _
    Optional ByRef strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:开具电子票据接口
    '入参:
    '    frmMain-调用主窗口
    '    bytInvoiceType-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '    lngEInvoiceID-电子票据使用记录.ID
    '    lngBalanceID-bytInvoiceType=2(预交)时：原预交ID ;bytInvoiceType<>2(预交)时：原结帐ID
    '    lngWriteOffID-bytInvoiceType<>2(预交)时：当前冲销ID,部分退费时有效
    '    strEInvoiceClientCode-传入客户端开票点编号
    '    strExtend-扩展参数,暂留,现传为空
    '出参:
    '    strEInvoiceClientCode-返回的客户端开票点编号
    '    strSysSouceName_Out-系统来源名称
    '    strEInvoiceCode_out-电子票据代码
    '    strEInvoiceNO_out-电子票据号码
    '    strCheckCode_out-电子票据校验码
    '    strCreateTime_Out-电子票据生成时间,格式:yyyyMMddHHmmssSSS
    '    strEInvQRCode_Out-电子票据二维码图片数据,base64加密存储,使用时,需解码
    '    strEInvUrl_Out-电子票据URL地址
    '    strEInvUrl1_Out-电子票据URL外网地址
    '    strEinvRemark_Out-电子票据的其他信息
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数,/Json格式
    '       totalamt  number 票据金额
    '       返回预交凭证
    '        voucherbatchcode    String  50  预交金凭证代码
    '        voucherno   String  20  预交金凭证号码
    '        voucherrandom   String  20  预交金凭证校验码
    '返回:开具电子票据成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.向医疗电子票据管理平台发起电子票据开具请求,生成电子票据,并返回电子票据相关信息
    '   2.返回的电子票据信息存储在“电子票据使用记录”中
    '调用者:
    '   1.门诊收费(部分退费)、结帐、挂号、 体检结帐、缴预交等
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim blnResult As Boolean
    Dim dbl票据金额 As Double
    Dim str接口标识 As String
    Dim strData As String, strMsg As String
    On Error GoTo ErrHand
    
    strErrMsg_Out = "": strExtend = ""
    If Not mdlEInvoice_BS.zlGetJson_CreateEInvoice(bytInvoiceType, lngEInvoiceID, lngBalanceID, lngWriteOffID, strEInvoiceClientCode, str接口标识, dbl票据金额, strData, strErrMsg_Out) Then Exit Function
    If Not EInvoiceRequest(str接口标识, strData, blnResult, strMsg) Then
        strErrMsg_Out = strMsg
        Exit Function
    End If
    
    If mobjJson.SetJsonString(strMsg) = False Then
        strErrMsg_Out = "博思接口返回的Json解析错误,请检查"
        Exit Function
    End If
    
    If bytInvoiceType = 2 And lngWriteOffID <> 0 Then
        strEInvoiceCode_out = mobjJson.GetJsonNodeValue("eScarletBillBatchCode")
        strEInvoiceNo_Out = mobjJson.GetJsonNodeValue("eScarletBillNo")
        strCheckCode_out = mobjJson.GetJsonNodeValue("eScarletRandom")
    Else
        strEInvoiceCode_out = mobjJson.GetJsonNodeValue("billBatchCode")
        strEInvoiceNo_Out = mobjJson.GetJsonNodeValue("billNo")
        strCheckCode_out = mobjJson.GetJsonNodeValue("random")
    End If
    strCreateTime_Out = mobjJson.GetJsonNodeValue("createTime")
    strEInvQRCode_Out = mobjJson.GetJsonNodeValue("billQRCode")
    strEInvUrl_Out = mobjJson.GetJsonNodeValue("pictureUrl")
    If gBs_Type.支持版本 > BS_Version.V2_0_3 And bytInvoiceType <> 2 Then
        strEInvUrl1_Out = mobjJson.GetJsonNodeValue("pictureNetUrl")
    Else
        strEInvUrl1_Out = ""
    End If
        
    strExtend = strExtend & "" & GetJsonNodeString("totalamt", dbl票据金额, Json_num)
    If bytInvoiceType = 2 Then '预交
        strExtend = strExtend & "," & GetJsonNodeString("voucherbatchcode", mobjJson.GetJsonNodeValue("voucherBatchCode"), Json_Text)
        strExtend = strExtend & "," & GetJsonNodeString("voucherno", mobjJson.GetJsonNodeValue("voucherNo"), Json_Text)
        strExtend = strExtend & "," & GetJsonNodeString("voucherrandom", mobjJson.GetJsonNodeValue("voucherRandom"), Json_Text)
    End If
    strExtend = "{" & strExtend & "}"
    
    strSysSouceName_Out = gstrProductName
    clsEInvoiceModule_zlCreateEInvoice = True
    Exit Function
ErrHand:
    strMsg = Err.Description
End Function
Private Function clsEInvoiceModule_zlCheckCancelEInvoiceValied(ByVal frmMain As Object, ByVal lngID As Long, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查作废电子票据的合法性
    '入参:
    '    frmMain-调用主窗口
    '    lngID-需要作废的电子票据记录ID:电子票据使用记录.ID
    '    strExtend-扩展参数,暂留,现传为空
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数,暂留,现传为空
    '返回:作废合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.电子票据作废前的一些合法性检查
    '调用者:
    '   1.门诊退费、结帐作废、退号、 体检结帐作废、预交退款等
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnResult As Boolean, blnCheckAcc As Boolean
    Dim strJson_In As String, strJsonAcc_In As String, strJson_Out As String
    Dim strResult As String, strEInvoiceNO As String
    
    If gBs_Type.支持版本 > BS_Version.V2_0_3 Then blnCheckAcc = True
    If mdlEInvoice_BS.zlGetJson_CheckCancelEInvoice(lngID, strEInvoiceNO, strJson_In, strErrMsg_Out, blnCheckAcc, strJsonAcc_In) = False Then Exit Function
    If Not EInvoiceRequest("getEBillStatesByBillInfo", strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    If mobjJson.SetJsonString(strJson_Out) = False Then
        strErrMsg_Out = "博思接口返回的Json解析错误，请检查！"
        Exit Function
    End If
    
    If blnCheckAcc Then
        strJson_Out = ""
        If Not EInvoiceRequest("getEBillAccountStatus", strJsonAcc_In, blnResult, strJson_Out) Then
            strErrMsg_Out = strJson_Out
            Exit Function
        End If
        If mobjJson.SetJsonString(strJson_Out) = False Then
            strErrMsg_Out = "博思接口返回的Json解析错误，请检查！"
            Exit Function
        End If
        strResult = mobjJson.GetJsonNodeValue("state")
        If strResult = "1" Then
            strErrMsg_Out = "电子票据【" & strEInvoiceNO & "】已入账，不能作废！"
            Exit Function
        End If
    End If
    
    clsEInvoiceModule_zlCheckCancelEInvoiceValied = True
End Function



Private Function clsEInvoiceModule_zlCancelEInvoice(ByVal frmMain As Object, ByVal lngID As Long, ByRef strEInvoiceClientCode As String, ByRef strSysSouceName_Out As String, _
    ByRef strEInvoiceCode_out As String, ByRef strEInvoiceNo_Out As String, _
    ByRef strCheckCode_out As String, ByRef strCreateTime_Out As String, ByRef strEInvQRCode_Out As String, ByRef strEInvUrl_Out As String, ByRef strEInvUrl1_Out As String, _
    ByRef strEinvRemark_Out As String, ByRef strErrMsg_Out As String, Optional ByRef strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废电子票据接口
    '入参:
    '    frmMain-调用主窗口
    '    lngID-需要作废的电子票据记录ID: 电子票据使用记录.ID
    '    strEInvoiceClientCode-电子票据开票点
    '    strEInvoiceNode_Out-开票点
    '    strExtend-扩展参数，JSON格式。当his没有票据信息，但平台有，需要传入平台的票据信息
    '       bustype '业务标识
    '       billbatchcode '电子票据代码
    '       billno '电子票据号码
    '出参:
    '    strEInvoiceClientCode-电子票据开票点
    '    strSysSouceName_Out-系统来源名称
    '    strEInvoiceCode_out-电子票据代码(红票)
    '    strEInvoiceNO_out-电子票据号码(红票)
    '    strCheckCode_out-电子票据校验码(红票)
    '    strCreateTime_Out-电子票据生成时间(红票),格式:yyyy-mm-dd HH:MM:SS
    '    strEInvQRCode_Out-电子票据二维码图片数据(红票),base64加密存储,使用时，需解码
    '    strEInvUrl_Out-电子票据URL地址(红票)
    '    strEInvUrl1_Out-电子票据外网URL地址(红票)
    '    strEinvRemark_Out-电子票据的其他信息(红票)
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数，JSON格式返回
    '        voucherbatchcode    String  50  红票预交金凭证代码
    '        voucherno   String  20  红票预交金凭证号码
    '        voucherrandom   String  20  红票预交金凭证校验码
    '返回:作废成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.对已经成生的电子票据进行作废
    '   2.返回的电子票据信息存储在“电子票据使用记录”中
    '调用者:
    '   1.门诊退费、结帐作废、退号、 体检结帐作废、预交退款等
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnResult As Boolean, blnNoInputReason As Boolean
    Dim strJson_In As String, strJson_Out As String, strServiceCode As String
    Dim strReason  As String
    Dim strEInvoiceType As String, strEInvoiceCode As String, strEInvoiceNO As String
    
    If mlngCurEInvoiceID = lngID Then
        clsEInvoiceModule_zlCancelEInvoice = True '作废纸质票据同时作废了电子票据
        Exit Function
        
        blnNoInputReason = True
        strReason = mstrCurReason
    End If
    
    If mobjJson.SetJsonString(strExtend) Then
        strEInvoiceType = mobjJson.GetJsonNodeValue("bustype")
        strEInvoiceCode = mobjJson.GetJsonNodeValue("billbatchcode")
        strEInvoiceNO = mobjJson.GetJsonNodeValue("billno")
    End If
    If mdlEInvoice_BS.zlGetJson_CancelEInvoice(frmMain, lngID, strEInvoiceClientCode, blnNoInputReason, strServiceCode, _
        strJson_In, strReason, strErrMsg_Out, strEInvoiceType, strEInvoiceCode, strEInvoiceNO) = False Then Exit Function
    If Not EInvoiceRequest(strServiceCode, strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    
    strExtend = ""
    If mobjJson.SetJsonString(strJson_Out) = False Then
        strErrMsg_Out = "博思接口返回的Json解析错误，请检查！"
        Exit Function
    End If
    strSysSouceName_Out = gstrProductName '暂定
    strEInvoiceCode_out = mobjJson.GetJsonNodeValue("eScarletBillBatchCode")
    strEInvoiceNo_Out = mobjJson.GetJsonNodeValue("eScarletBillNo")
    strCheckCode_out = mobjJson.GetJsonNodeValue("eScarletRandom")
    strCreateTime_Out = mobjJson.GetJsonNodeValue("createTime")
    strEInvQRCode_Out = mobjJson.GetJsonNodeValue("billQRCode")
    strEInvUrl_Out = mobjJson.GetJsonNodeValue("pictureUrl")
    If gBs_Type.支持版本 > BS_Version.V2_0_3 Then
        strEInvUrl1_Out = mobjJson.GetJsonNodeValue("pictureNetUrl")
    Else
        strEInvUrl1_Out = ""
    End If
    If strServiceCode = "cancelPayMentVoucherBalance" Then
        strExtend = strExtend & "" & GetJsonNodeString("voucherbatchcode", mobjJson.GetJsonNodeValue("voucherBatchCode"), Json_Text)
        strExtend = strExtend & "," & GetJsonNodeString("voucherno", mobjJson.GetJsonNodeValue("voucherNo"), Json_Text)
        strExtend = strExtend & "," & GetJsonNodeString("voucherrandom", mobjJson.GetJsonNodeValue("voucherRandom"), Json_Text)
        strExtend = "{" & strExtend & "}"
    End If
    
    strSysSouceName_Out = gstrProductName
    clsEInvoiceModule_zlCancelEInvoice = True
End Function

Private Function clsEInvoiceModule_zlSendEinvoiceMsg(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发送电子票据通知信息
    '入参:
    '    frmMain-调用主窗口
    '    lngEInvoiceID-电子票据记录ID
    '    strExtend-扩展参数，暂留，现传为空
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数，暂留，现传为空
    '返回:发送成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.主要是通知患者电子票据的相关信息:比如 打印告知单、在微信、电子邮件等平台通知患者查看电子票据
    '调用者:
    '   1.门诊收费（或部分退费）、预交、结帐、体检结帐等开具电子票据成功后，需要调用本接口通知患者
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson_In As String, strJson_Out As String
    Dim blnResult As Boolean
    
    '组织业务数据
    If mdlEInvoice_BS.zlGetJson_SendNotice(lngEInvoiceID, strJson_In, strErrMsg_Out) = False Then Exit Function
    If Not EInvoiceRequest("sendBillNotice", strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    clsEInvoiceModule_zlSendEinvoiceMsg = True
End Function


Private Function clsEInvoiceModule_zlGetNextInvoiceNo(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, strInvoiceNO_Out As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取下一张有效发票号
    '入参:
    '    frmMain-调用主窗口
    '    bytInvoiceType-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '    strExtend-扩展参数，暂留，现传为空
    '出参:
    '    strInvoiceNO_Out-下一张有效发票号
    '    strErrMsg_Out-获取失败时，返回的提示信息
    '    strExtend-扩展参数，暂留，现传为
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.在换开纸质发票时，需要先读取下一张有效的发票号,以便显示在主界面上
    '调用者:
    '   1.门诊收费(部分退费)、结帐、挂号、 体检结帐、缴预交等进入界面时，需要读取显示在界面上
    '   2.换开、重新换开
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson_In As String, strJson_Out As String, strPaperCode As String
    Dim blnResult As Boolean
    
    '组织业务数据
    If mdlEInvoice_BS.zlGetJson_GetNextInvoiceNo(bytInvoiceType, mstrEInvoiceClientCode, strPaperCode, strJson_In, strErrMsg_Out) = False Then Exit Function
    If Not EInvoiceRequest("getPaperBillNo", strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    If mobjJson.SetJsonString(strJson_Out) = False Then
        strErrMsg_Out = "博思接口返回的Json解析错误，请检查！"
        Exit Function
    End If
    '返回纸质票据代码+号码
    strInvoiceNO_Out = strPaperCode & mobjJson.GetJsonNodeValue("pBillNo")
    clsEInvoiceModule_zlGetNextInvoiceNo = True
End Function
Private Function clsEInvoiceModule_zlTurnPaperInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, ByVal lngEInvoiceID As Long, ByVal strEInvoiceCode As String, ByVal strEInvoiceNO As String, _
    ByVal strCreateTime As String, ByVal strOperatorCode As String, ByVal strOperatorName As String, ByRef strInvoiceNO_Out As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:换开纸质发票接口
    '入参:
    '    frmMain-调用主窗口
    '    bytInvoiceType-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '    strInvoiceNO-发票号
    '    lngEInvoiceID-电子票据使用记录ID
    '    strEInvoiceCode-电子票据代码
    '    strEInvoiceNO-电子票据号码
    '    strCreateTime-电子票据生成时间,格式:yyyy-mm-dd HH:MM:SS
    '    strOperatorCode-操作员编号
    '    strOperatorName-操作员姓名
    '    strExtend-扩展参数，暂留，现传为空
    '出参:
    '    strInvoiceNO_Out:返回本次的发票号信息,多个用逗号分离;返回空时，将以传入的发票号为准
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数，暂留，现传为空
    '返回:换开成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.主要是针对电子票据换开纸质票据业务
    '调用者:
    '   1.门诊收费(部分退费)、结帐、挂号、 体检结帐、缴预交
    '   2.换开、重新换开
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson_In As String, strJson_Out As String, strServiceCode As String
    Dim blnResult As Boolean
    
    '组织业务数据
    '内部处理换开或重新换开
    If mdlEInvoice_BS.zlGetJson_TurnPaper(bytInvoiceType, mstrEInvoiceClientCode, strInvoiceNO, _
                    lngEInvoiceID, strEInvoiceCode, strEInvoiceNO, strCreateTime, strOperatorCode, _
                    strOperatorName, strServiceCode, strJson_In, strErrMsg_Out) = False Then Exit Function
    If Not EInvoiceRequest(strServiceCode, strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    
    strInvoiceNO_Out = strInvoiceNO
    clsEInvoiceModule_zlTurnPaperInvoice = True
End Function
Private Function clsEInvoiceModule_zlPrintPaperInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, ByVal lngEInvoiceID As Long, ByVal strOperatorCode As String, ByVal strOperatorName As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印纸质发票接口
    '入参:
    '    frmMain-调用主窗口
    '    bytInvoiceType-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '    strInvoiceNO-发票号
    '    lngBalanceID-bytInvoiceType=2(预交)时：原预交ID,bytInvoiceType<>2(预交)时：原结帐ID
    '    lngEInvoiceID-电子票据使用记录ID
    '    strOperatorCode-操作员编号
    '    strOperatorName-操作员姓名
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend -扩展参数，暂留，现传为空
    '返回:打印返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.   打印分两种：一种是HIS自己打印;一种是通过三方提供的票据打印   不管哪种方式，都是在此接口中编制打印
    '调用者:
    '   1.门诊收费(部分退费)、结帐、挂号、 体检结帐、缴预交
    '   2.换开、重新换开
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson_In As String, strJson_Out As String, strServiceCode As String
    Dim blnResult As Boolean
    
    '组织业务数据
    '内部处理打印或重新打印
    If mdlEInvoice_BS.zlGetJson_TurnPaperPrint(bytInvoiceType, strInvoiceNO, strServiceCode, strJson_In, strErrMsg_Out) = False Then Exit Function
    If Not PrintEInvoiceRequest(strServiceCode, strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    clsEInvoiceModule_zlPrintPaperInvoice = True
End Function
Private Function clsEInvoiceModule_zlCancelPaperInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, ByVal lngEInvoiceID As Long, _
                ByVal strOperatorCode As String, ByVal strOperatorName As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废纸质票据接口
    '入参:
    '    frmMain-调用主窗口
    '    bytInvoiceType-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '    strInvoiceNO-发票号
    '    lngBalanceID-"1.bytInvoiceType=2(预交)时：原预交ID"
    '    2.bytInvoiceType<>2(预交)时：原结帐ID"
    '    lngEInvoiceID-电子票据使用记录ID
    '    strOperatorCode-操作员编号
    '    strOperatorName-操作员姓名
    '    strExtend -扩展参数，暂留，现传为空
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend -扩展参数，JSON格式，格式如下
    '            einvoiceinfo    String      电子票据信息：当作废纸质票据同时作废了电子票据时通过strExpend返回
    '                einvoicecode    String  20  电子红票代码
    '                einvoiceno      String  20  电子红票号码
    '                einvoicerandom  String  20  电子红票校验码
    '                createtime      String  17  电子红票生成时间
    '                billqrcode      String  不限    电子红票二维码图片数据
    '                pictureurl      String  不限    电子票据H5页面URL
    '                pictureneturl   String  不限    电子票据外网H5页面URL地址
    '                syssoucename    String  不限    开票系统名称
    '                voucherbatchcode    String  50  红票预交金凭证代码,作废预交票据时才有节点
    '                voucherno       String  20  红票预交金凭证号码,同上
    '                voucherrandom   String  10  红票预交金凭证校验码,同上
    '返回:作废纸质票据成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.针对打印机打印卡纸、撕纸等造成的纸质票据损坏，而平台又换纸成功的，需要作废换卡的纸质票据
    '调用者:
    '   1.门诊收费、结帐、挂号、 体检结帐、缴预交等的作废功能
    '   2.重新换开及作废
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng冲销ID As Long
    Dim strBusDateTime As String, strReason As String
    Dim blnResult As Boolean
    Dim strJson_In As String, strJson_Out As String, strServiceCode As String
    Dim strEInvoiceCode As String, strEInvoiceNO As String, strCheckCode As String
    Dim strEInvUrl As String, strEInvUrl1 As String, strEinvRemark As String
    Dim strSysSouceName As String, strEInvQRCode As String
    Dim strVoucherCode As String, strVoucherNo As String, strVoucherRandom As String
    
    strExtend = ""
    '组织业务数据
    If mdlEInvoice_BS.zlGetJson_CancelPaper(frmMain, bytInvoiceType, strInvoiceNO, _
                    mstrEInvoiceClientCode, strOperatorName, _
                    strServiceCode, strBusDateTime, strJson_In, strReason, strErrMsg_Out) = False Then Exit Function
    
    If Not EInvoiceRequest(strServiceCode, strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    
    If mobjJson.SetJsonString(strJson_Out) = False Then
        strErrMsg_Out = "博思接口返回的Json解析错误，请检查！"
        Exit Function
    End If
    
    strExtend = strExtend & "" & GetJsonNodeString("einvoicecode", mobjJson.GetJsonNodeValue("eScarletBillBatchCode"), Json_Text)
    strExtend = strExtend & "," & GetJsonNodeString("einvoiceno", mobjJson.GetJsonNodeValue("eScarletBillNo"), Json_Text)
    strExtend = strExtend & "," & GetJsonNodeString("einvoicerandom", mobjJson.GetJsonNodeValue("eScarletRandom"), Json_Text)
    strExtend = strExtend & "," & GetJsonNodeString("createtime", mobjJson.GetJsonNodeValue("createTime"), Json_Text)
    strExtend = strExtend & "," & GetJsonNodeString("billqrcode", mobjJson.GetJsonNodeValue("billQRCode"), Json_Text)
    strExtend = strExtend & "," & GetJsonNodeString("pictureurl", mobjJson.GetJsonNodeValue("pictureUrl"), Json_Text)
    strExtend = strExtend & "," & GetJsonNodeString("syssoucename", gstrProductName, Json_Text)
    If gBs_Type.支持版本 > BS_Version.V2_0_3 Then
        strExtend = strExtend & "," & GetJsonNodeString("pictureneturl", mobjJson.GetJsonNodeValue("pictureUrl"), Json_Text)
    Else
        strExtend = strExtend & "," & GetJsonNodeString("pictureneturl", "", Json_Text)
    End If
    If bytInvoiceType = 2 Then
        strExtend = strExtend & "," & GetJsonNodeString("voucherbatchcode", mobjJson.GetJsonNodeValue("voucherBatchCode"), Json_Text)
        strExtend = strExtend & "," & GetJsonNodeString("voucherno", mobjJson.GetJsonNodeValue("voucherNo"), Json_Text)
        strExtend = strExtend & "," & GetJsonNodeString("voucherrandom", mobjJson.GetJsonNodeValue("voucherRandom"), Json_Text)
    End If
    '返回完整Json
    strExtend = "{" & Chr(34) & "einvoiceinfo" & Chr(34) & ":{" & strExtend & "}}"
    mlngCurEInvoiceID = lngEInvoiceID
    mstrCurReason = strReason
    clsEInvoiceModule_zlCancelPaperInvoice = True
End Function

Private Sub clsEInvoiceModule_zlEInvoiceManager(ByVal frmMain As Object, Optional strExtend As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:电子票据管理接口
    '入参:
    '    frmMain-调用主窗口
    '    strExtend -扩展参数，暂留，现传为空
    '出参:
    '    strExtend -扩展参数，暂留，现传为空
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '    1.主要是在导航台管理中调用电子票据管理,类似于医保部件中的“保险账户管理”模块
    '    2.该接口功能：对电子票据的对账、作废、查询等
    '调用者:
    '   1.新增模块："1144-电子票据管理"
    '   2.入口程序：由部件zL9CashBill调用本接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmEInvoiceManage Is Nothing Then Set mfrmEInvoiceManage = New frmEInvoiceManage
    Call mfrmEInvoiceManage.ShowMe(frmMain, mlngSys, mlngModule, mstrDBUser, Me)
End Sub

Private Sub clsEInvoiceModule_zlBHEInvoiceManager(ByVal lngMain As Long, Optional ByRef strExtend As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:电子票据管理接口
    '入参:
    '    lngMain:主窗体句柄
    '    strExtend -扩展参数，暂留，现传为空
    '出参:
    '    strExtend -扩展参数，暂留，现传为空
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '    1.主要是在导航台管理中调用电子票据管理,类似于医保部件中的“保险账户管理”模块
    '    2.该接口功能：对电子票据的对账、作废、查询等
    '调用者:
    '   1.新增模块："1144-电子票据管理"
    '   2.入口程序：由部件zL9CashBill调用本接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmEInvoiceManage Is Nothing Then Set mfrmEInvoiceManage = New frmEInvoiceManage
    Call mfrmEInvoiceManage.BHShowMe(lngMain, mlngSys, mlngModule, mstrDBUser, Me)
End Sub

Private Sub Class_Initialize()
    Set mobjJson = New clsJson
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mcnOracle = Nothing
    Set mfrmMain = Nothing
    Set mobjJson = Nothing
    If Not mfrmParaForm Is Nothing Then
        Unload mfrmParaForm
        Set mfrmParaForm = Nothing
    End If
    If Not mfrmEInvoiceManage Is Nothing Then Unload mfrmEInvoiceManage: Set mfrmEInvoiceManage = Nothing
End Sub

Private Function EInvoiceRequest(ByVal strQuestCode As String, ByVal strData As String, blnResult As Boolean, strMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 调用博思服务
    ' 入参 : strQuestCode-调用的服务接口
    '        strData-业务数据
    ' 出参 : blnResult-博思接口调用是否成功
    '        strMsg-博思返回接口Json数据
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/21 16:45
    '---------------------------------------------------------------------------------------
    Dim objHttpRequest As WinHttp.WinHttpRequest
    Dim strGUID As String, strCharSet As String, strNoise As String, strSign As String, strMethod As String
    Dim strContentType As String, strBody As String, strResponse As String, strVersion As String
    Dim strUrl As String
    Dim arrCodeBytes() As Byte
    Dim objMD5 As New clsMD5, arrTmp() As Byte
    
    On Error GoTo ErrHand
    strGUID = GetGUID()
    blnResult = False: strMsg = ""
    strVersion = "1.0"

    '字符转码
    strCharSet = gBs_Type.字符编码
    If strCharSet = "Unicode" Then
        arrCodeBytes = strData 'zlStr.StringToUnicode(strData)
    Else
        arrCodeBytes = zlStr.StringToUTF8Bytes(strData)
    End If
    strData = zlStr.EncodeBase64(arrCodeBytes)
    
    '响应方式
    strMethod = "POST"
    
    '签名-MD5加密
    strSign = "appid=" & gBs_Type.应用帐号 & "&data=" & strData & "&noise=" & strGUID & "&key=" & gBs_Type.签名私钥 & "&version=" & strVersion
    arrTmp = strSign
    strSign = objMD5.DigestByteToHexStr(objMD5.Unicode_To_UTF8(arrTmp))
    
    '内容格式
    strContentType = gBs_Type.数据传输方式
    If strContentType = "" Then strContentType = "application/json"
    
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "开始", "CharSet=" & strCharSet & ",Method=" & strMethod & ",ContentType=" & strContentType & ",Version=" & strVersion)
    
    '服务调用内容
    If strContentType = "application/json" Then
        strBody = "{""appid"":""" & gBs_Type.应用帐号 & """,""data"":""" & strData & """,""noise"":""" & strGUID & """,""version"":""" & strVersion & """,""sign"":""" & strSign & """}"
    Else
        Exit Function '暂不支持其他格式
    End If
    
    '初始化HTTP
    Set objHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    strUrl = gBs_Type.URL_Type & "://" & gBs_Type.URL_Address & "/" & strQuestCode
    objHttpRequest.Open strMethod, strUrl, False '同步发送接收
    
    '设置参数
    'WinHttp.SetTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
    'resolveTimeout = 10000 '解析 DNS 名字的超时时间，10000 毫秒。
    'connectTimeout = 10000 '建立 Winsock 连接的超时时间，10000 毫秒。
    'sendTimeout = 120000 '发送数据的超时时间，120000 毫秒。
    'receiveTimeout = 60000 '接收 response 的超时时间，60000 毫秒。
    objHttpRequest.SetTimeouts 60000, 60000, 120000, 60000 '设置操作超时时间
    ' 非常重要(忽略错误)
    objHttpRequest.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    ' 其它请求头设置
    objHttpRequest.SetRequestHeader "Content-Type", "application/json"
    objHttpRequest.SetRequestHeader "Content-Length", Len(strBody)
    
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "执行", "服务地址=" & strUrl & ",入参=" & vbCrLf & strBody)
    
    '发送请求
    objHttpRequest.Send strBody
    '获取返回结果
    strResponse = objHttpRequest.ResponseText
    Set objHttpRequest = Nothing
    
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "执行", "出参=" & vbCrLf & strResponse)
    
    Dim strData_Out As String, strResult_Out As String
    '检查data节点是否正常返回
    If InStr(strResponse, "data") < 1 Then
        strMsg = "博思data节点返回空"
        Exit Function
    End If
    
    '将返回结果MD5加密后比对
    strData = zlStr.JSONParse("data", strResponse)
    strNoise = zlStr.JSONParse("noise", strResponse)
    strSign = zlStr.JSONParse("sign", strResponse)
    
    strBody = "appid=" & gBs_Type.应用帐号 & "&data=" & strData & "&noise=" & strNoise & "&key=" & gBs_Type.签名私钥 & "&version=" & strVersion
    arrTmp = strBody
    
    If objMD5.DigestByteToHexStr(objMD5.Unicode_To_UTF8(arrTmp)) <> strSign Then
        strMsg = "博思电子票据数据丢失，请重试！"
        Exit Function
    End If
    
    arrCodeBytes = zlStr.DecodeBase64(zlStr.JSONParse("data", strResponse), True)
    If strCharSet = "Unicode" Then
        strData_Out = zlStr.UnicodeToString(arrCodeBytes)
    Else
        strData_Out = zlStr.UTF8BytesToString(arrCodeBytes)
    End If
    
    If InStr(strData_Out, "message") < 1 Then
        '检查message节点是否正常返回
        strMsg = "博思message节点返回空"
        Exit Function
    End If
    
    strResult_Out = zlStr.JSONParse("result", strData_Out)
    blnResult = (strResult_Out = "S0000")
    strMsg = zlStr.DecodeBase64(zlStr.JSONParse("message", strData_Out))
    
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "结束", "返回=true")
    
    Set objMD5 = Nothing
    EInvoiceRequest = blnResult
    Exit Function
ErrHand:
    strMsg = Err.Description
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "结束", "返回=false,错误信息=" & strMsg)
End Function

Private Function PrintEInvoiceRequest(ByVal strQuestCode As String, ByVal strData As String, blnResult As Boolean, strMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 调用博思服务
    ' 入参 : strQuestCode-调用的服务接口
    '        strData-业务数据
    ' 出参 : blnResult-博思接口调用是否成功
    '        strMsg-博思返回接口Json数据
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/21 16:45
    '---------------------------------------------------------------------------------------
    Dim objHttpRequest As WinHttp.WinHttpRequest
    Dim strGUID As String, strCharSet As String, strNoise As String, strSign As String, strMethod As String
    Dim strContentType As String, strBody As String, strResponse As String, strVersion As String
    Dim strUrl As String
    Dim arrCodeBytes() As Byte
    Dim objMD5 As New clsMD5, arrTmp() As Byte
    
    On Error GoTo ErrHand
    strGUID = GetGUID()
    blnResult = False: strMsg = ""
    strVersion = "1.0"

    '字符转码
    strCharSet = gBs_Type.字符编码
    If strCharSet = "Unicode" Then
        arrCodeBytes = strData 'zlStr.StringToUnicode(strData)
    Else
        arrCodeBytes = zlStr.StringToUTF8Bytes(strData)
    End If
    strData = zlStr.EncodeBase64(arrCodeBytes)
    
    '响应方式
    strMethod = "POST"
    
    '签名-MD5加密
    strSign = "appid=" & gBs_Type.应用帐号 & "&data=" & strData & "&noise=" & strGUID & "&key=" & gBs_Type.签名私钥 & "&version=" & strVersion
    arrTmp = strSign
    strSign = objMD5.DigestByteToHexStr(objMD5.Unicode_To_UTF8(arrTmp))
    
    '内容格式
    strContentType = gBs_Type.数据传输方式
    If strContentType = "" Then strContentType = "application/json"
    
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "开始", "CharSet=" & strCharSet & ",Method=" & strMethod & ",ContentType=" & strContentType & ",Version=" & strVersion)
    
    '服务调用内容
    If strContentType = "application/json" Then
        strBody = "{""appid"":""" & gBs_Type.应用帐号 & """,""data"":""" & strData & """,""noise"":""" & strGUID & """,""version"":""" & strVersion & """,""sign"":""" & strSign & """}"
    Else
        Exit Function '暂不支持其他格式
    End If
    
    strBody = "{""method"":""" & strQuestCode & """,""params"":" & strBody & "}"
    If strCharSet = "Unicode" Then
        arrCodeBytes = zlStr.StringToUnicode(strBody)
    Else
        arrCodeBytes = zlStr.StringToUTF8Bytes(strBody)
    End If
    strBody = zlStr.EncodeBase64(arrCodeBytes)
    
    '初始化HTTP
    strUrl = gBs_Type.URL_Type & "://" & mstrPrintIp & ":" & mstrPrintPort & "/extend?dllName=NontaxIndustry&Func=CallNontaxIndustry&payload=" & strBody
    
    Set objHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHttpRequest.Open strMethod, strUrl, False '同步发送接收
    
    '设置参数
    'WinHttp.SetTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
    'resolveTimeout = 10000 '解析 DNS 名字的超时时间，10000 毫秒。
    'connectTimeout = 10000 '建立 Winsock 连接的超时时间，10000 毫秒。
    'sendTimeout = 120000 '发送数据的超时时间，120000 毫秒。
    'receiveTimeout = 60000 '接收 response 的超时时间，60000 毫秒。
    objHttpRequest.SetTimeouts 60000, 60000, 120000, 60000 '设置操作超时时间

    ' 非常重要(忽略错误)
    objHttpRequest.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    ' 其它请求头设置
    objHttpRequest.SetRequestHeader "Content-Type", "application/json"
    objHttpRequest.SetRequestHeader "Content-Length", Len(strBody)
    
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "执行", "服务地址=" & strUrl & ",入参=" & vbCrLf & strBody)
    
    '发送请求
    objHttpRequest.Send strBody
    '获取返回结果
    strResponse = objHttpRequest.ResponseText
    Set objHttpRequest = Nothing
    
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "执行", "出参=" & vbCrLf & strResponse)
    
    If strQuestCode = "InitParams" Then
        Select Case strResponse
            Case "2001"
                strMsg = "参数缺失"
            Case "2002"
                strMsg = "指定路径下未找到该dll"
            Case "2002"
                strMsg = "方法调用异常"
            Case "T2000"
                PrintEInvoiceRequest = True
            Case "T2001"
                strMsg = "调用初始化方法，参数缺失"
            Case "T2002"
                strMsg = "调用初始化方法，执行异常"
            Case "T2003"
                strMsg = "调用业务方法前，未调用初始化方法"
            Case "T2004"
                strMsg = "调用内部方法，参数缺失"
            Case "T2005"
                strMsg = "调用业务方法，执行异常"
            Case Else
                strMsg = strResponse
        End Select
        Set objMD5 = Nothing
        Exit Function
    End If
    
    Dim strData_Out As String, strResult_Out As String
    '检查data节点是否正常返回
    If InStr(strResponse, "data") < 1 Then
        strMsg = "博思data节点返回空"
        Exit Function
    End If
    
    '将返回结果MD5加密后比对
    strData = zlStr.JSONParse("data", strResponse)
    strNoise = zlStr.JSONParse("noise", strResponse)
    strSign = zlStr.JSONParse("sign", strResponse)
    
    strBody = "appid=" & gBs_Type.应用帐号 & "&data=" & strData & "&noise=" & strNoise & "&key=" & gBs_Type.签名私钥 & "&version=" & strVersion
    arrTmp = strBody
    
    If objMD5.DigestByteToHexStr(objMD5.Unicode_To_UTF8(arrTmp)) <> strSign Then
        strMsg = "博思电子票据数据丢失，请重试！"
        Exit Function
    End If
    
    arrCodeBytes = zlStr.DecodeBase64(strData, True)
    If strCharSet = "Unicode" Then
        strData_Out = zlStr.UnicodeToString(arrCodeBytes)
    Else
        strData_Out = zlStr.UTF8BytesToString(arrCodeBytes)
    End If
    
    If InStr(strData_Out, "message") < 1 Then
        '检查message节点是否正常返回
        strMsg = "博思message节点返回空"
        Exit Function
    End If
    
    strResult_Out = zlStr.JSONParse("result", strData_Out)
    blnResult = (strResult_Out = "S0000")
    strMsg = zlStr.DecodeBase64(zlStr.JSONParse("message", strData_Out))
    
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "结束", "返回=true")
    
    Set objMD5 = Nothing
    PrintEInvoiceRequest = blnResult
    Exit Function
ErrHand:
    strMsg = Err.Description
    Call zlWritLogEx(Me, mlngModule, "EInvoiceRequest", "结束", "返回=false,错误信息=" & strMsg)
End Function

Public Function clsEInvoiceModule_ZlGetTotalCheckData(ByVal dtBegin As Date, ByVal dtEnd As Date, ByRef cllDatas_Out As Collection, _
    Optional ByVal bytMode As Byte = 1, Optional ByVal str开票点 As String, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取平台电子票据汇总数据
    '入参:
    '   dtBegin、dtEnd 业务开始日期、结束日期
    '   bytMode 数据类型：1-开票和退票，2-仅退票
    '   str开票点 平台开票点编码
    '出参:
    '   cllDatas 汇总数据，集合元素：(业务日期,开票数,开票金额,总笔数,返回结果,错误原因),Key=_业务日期
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJsonIn As String, strJsonOut As String, blnResult As Boolean
    Dim cllTemp As Collection, dtCurrent As Date
    Dim blnSuccess As Boolean, strErrMsg As String
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    Set cllDatas_Out = New Collection
    
    
    dtCurrent = dtBegin
    Do While dtCurrent <= dtEnd
        '1-费用业务发生时间
        If clsEInvoiceModule_ZLCheckTimeMode = 1 Then
            Select Case bytMode
            Case 2
                '服务标识：checkWriteOffData
                '入参：
                '   busDate 业务日期  String  10  是 格式:yyyy-MM-dd
                '   busType 业务标识    String  20  是  直接填写业务系统内部编码值，由医疗平台配置对照  V3_1_0
                '   placeCode   开票点编码  String  50  是  直接填写业务系统内部编码值，由医疗平台配置对照 V3_1_0
                '   isReturnList  是否返回列表  String  1 否 0:不返回，1:返回，空:返回
                '   pageNo  页码  Integer 10  是 传入值为1代表第一页,传入值为2代表第二页,依此类推.默认返回的数据是从第一页开始；第二次传入的页码计算由返回值中的总条数/每页条数
                '   pageSize  每页条数  Integer 10  是 每页返回最多返回200条
                '出参：
                '   busDate 业务日期  String  10  是
                '   copyNum 退费总笔数 Integer   是
                '   totalAmt  当前页退费总金额  Number  18,2  是
                '   totalNum  当前页退费总开票数 Integer 10  是
                '   allTotalAmt 所有页总金额  Number  18,2  是
                '   allTotalNum 所有页总开票数 Integer 10  是
                '   billNoList  退费票号段明细 String  不限  是 详见B-1-1,JSON格式
                strJsonIn = ""
                strJsonIn = strJsonIn & "" & GetJsonNodeString("busDate", Format(dtCurrent, "yyyy-MM-dd"), Json_Text)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("isReturnList", 0, Json_num)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("pageNo", 1, Json_num)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("pageSize", 200, Json_num)
                strJsonIn = "{" & strJsonIn & "}"
                
                 If EInvoiceRequest("checkWriteOffData", strJsonIn, blnResult, strJsonOut) Then
                        blnSuccess = mobjJson.SetJsonString(strJsonOut)
                        If Not blnSuccess Then strErrMsg = "服务 checkWriteOffData 返回的Json解析错误！JSON=" & strJsonOut
                    Else
                        strErrMsg = strJsonOut
                    End If
                     
                    If blnSuccess Then
                        Set cllTemp = New Collection
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("busDate")), "业务日期"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalNum")), "开票数"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalAmt")), "开票金额"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("copyNum")), "总笔数"
                        cllTemp.Add "成功", "返回结果"
                        cllTemp.Add "", "错误原因"
                        cllDatas_Out.Add cllTemp, "_" & cllTemp("业务日期")
                    End If
            
            Case Else
                If str开票点 = "" Then
                    '服务标识：checkTotalData
                    '入参：
                    '   busDate 业务日期  String  10  是 格式:yyyy-MM-dd
                    '   busType 业务标识    String  20  否  直接填写业务系统内部编码值，由医疗平台配置对照  V3_1_0
                    '   placeCode   开票点编码  String  50  否  直接填写业务系统内部编码值，由医疗平台配置对照 V3_1_0
                    '   isReturnList  是否返回列表  String  1 否 0:不返回，1:返回，空:返回
                    '   pageNo  页码  Integer 10  是 传入值为1代表第一页,传入值为2代表第二页,依此类推.默认返回的数据是从第一页开始；第二次传入的页码计算由返回值中的总条数/每页条数
                    '   pageSize  每页条数  Integer 10  是 每页返回最多返回200条
                    '出参：
                    '   busDate 业务日期  String  10  是
                    '   copyNum 总笔数 Integer   是
                    '   totalAmt  当前页总金额  Number  18,2  是
                    '   totalNum  当前页总开票数 Integer   是
                    '   allTotalAmt 所有页总金额  Number  18,2  是
                    '   allTotalNum 所有页总开票数 Integer 10  是
                    '   billNoList  票号段明细 String  不限  是 详见B-1-1,JSON格式
                    strJsonIn = ""
                    strJsonIn = strJsonIn & "" & GetJsonNodeString("busDate", Format(dtCurrent, "yyyy-MM-dd"), Json_Text)
                    strJsonIn = strJsonIn & "," & GetJsonNodeString("isReturnList", 0, Json_num)
                    strJsonIn = strJsonIn & "," & GetJsonNodeString("pageNo", 1, Json_num)
                    strJsonIn = strJsonIn & "," & GetJsonNodeString("pageSize", 200, Json_num)
                    strJsonIn = "{" & strJsonIn & "}"
                    
                    If EInvoiceRequest("checkTotalData", strJsonIn, blnResult, strJsonOut) Then
                        blnSuccess = mobjJson.SetJsonString(strJsonOut)
                        If Not blnSuccess Then strErrMsg = "服务 checkTotalData 返回的Json解析错误！JSON=" & strJsonOut
                    Else
                        strErrMsg = strJsonOut
                    End If
                     
                    If blnSuccess Then
                        Set cllTemp = New Collection
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("busDate")), "业务日期"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalNum")), "开票数"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalAmt")), "开票金额"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("copyNum")), "总笔数"
                        cllTemp.Add "成功", "返回结果"
                        cllTemp.Add "", "错误原因"
                        cllDatas_Out.Add cllTemp, "_" & cllTemp("业务日期")
                    End If
                
                Else
                    '服务标识：checkDataByPlace
                    '入参：
                    '   busDate 业务日期  String  10  是 格式:yyyy-MM-dd
                    '   placeCode 开票点编码 String  50  是 直接填写业务系统内部编码值，由医疗平台配置对照
                    '   busType 业务标识    String  20  是  直接填写业务系统内部编码值，由医疗平台配置对照  V3_1_0
                    '   isReturnList  是否返回列表  String  1 否 0:不返回，1:返回，空:返回
                    '   pageNo  页码  Integer 10  是 传入值为1代表第一页,传入值为2代表第二页,依此类推.默认返回的数据是从第一页开始；第二次传入的页码计算由返回值中的总条数/每页条数
                    '   pageSize  每页条数  Integer 10  是 每页返回最多返回200条
                    '出参：
                    '   busDate 业务日期  String  10  是
                    '   copyNum 总笔数 Integer   是
                    '   totalAmt  当前页总金额  Number  18,2  是
                    '   totalNum  当前页总开票数 Integer   是
                    '   allTotalAmt 所有页总金额  Number  18,2  是
                    '   allTotalNum 所有页总开票数 Integer 10  是
                    '   billNoList  票号段明细 String  不限  是 详见B-1-1,JSON格式
                    strJsonIn = ""
                    strJsonIn = strJsonIn & "" & GetJsonNodeString("busDate", Format(dtCurrent, "yyyy-MM-dd"), Json_Text)
                    strJsonIn = strJsonIn & "," & GetJsonNodeString("placeCode", str开票点, Json_Text)
                    strJsonIn = strJsonIn & "," & GetJsonNodeString("isReturnList", 0, Json_num)
                    strJsonIn = strJsonIn & "," & GetJsonNodeString("pageNo", 1, Json_num)
                    strJsonIn = strJsonIn & "," & GetJsonNodeString("pageSize", 200, Json_num)
                    strJsonIn = "{" & strJsonIn & "}"
                    
                    blnSuccess = False
                    If EInvoiceRequest("checkDataByPlace", strJsonIn, blnResult, strJsonOut) Then
                        blnSuccess = mobjJson.SetJsonString(strJsonOut)
                        If Not blnSuccess Then strErrMsg = "服务 checkDataByPlace 返回的Json解析错误！JSON=" & strJsonOut
                    Else
                        strErrMsg = strJsonOut
                    End If
                     
                    If blnSuccess Then
                        Set cllTemp = New Collection
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("busDate")), "业务日期"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalNum")), "开票数"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalAmt")), "开票金额"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("copyNum")), "总笔数"
                        cllTemp.Add "成功", "返回结果"
                        cllTemp.Add "", "错误原因"
                        cllDatas_Out.Add cllTemp, "_" & cllTemp("业务日期")
                    End If
                End If
            End Select
            
        Else '0-票据开具时间
            Select Case bytMode
            Case 2
                '服务标识：checkWriteOffDataByIvcDate
                '入参：
                '   ivcDate 开票日期  String  10  是 格式:yyyy-MM-dd
                '   placeCode   开票点编码  String  50  是  直接填写业务系统内部编码值，由医疗平台配置对照
                '   busType 业务标识    String  20  是  直接填写业务系统内部编码值，由医疗平台配置对照
                '   isReturnList  是否返回列表  String  1 否 0:不返回，1:返回，空:返回
                '   pageNo  页码  Integer 10  是 传入值为1代表第一页,传入值为2代表第二页,依此类推.默认返回的数据是从第一页开始；第二次传入的页码计算由返回值中的总条数/每页条数
                '   pageSize  每页条数  Integer 10  是 每页返回最多返回200条
                '出参：
                '   ivcDate 开票日期  String  10  是
                '   copyNum 退费总笔数 Integer   是
                '   totalAmt  当前页退费总金额  Number  18,2  是
                '   totalNum  当前页退费总开票数 Integer 10  是
                '   allTotalAmt 所有页总金额  Number  18,2  是
                '   allTotalNum 所有页总开票数 Integer 10  是
                '   billNoList  退费票号段明细 String  不限  是 详见B-1-1,JSON格式
                strJsonIn = ""
                strJsonIn = strJsonIn & "" & GetJsonNodeString("ivcDate", Format(dtCurrent, "yyyy-MM-dd"), Json_Text)
                
                strJsonIn = strJsonIn & "," & GetJsonNodeString("isReturnList", 0, Json_num)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("pageNo", 1, Json_num)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("pageSize", 200, Json_num)
                strJsonIn = "{" & strJsonIn & "}"
                
                 If EInvoiceRequest("checkWriteOffDataByIvcDate", strJsonIn, blnResult, strJsonOut) Then
                        blnSuccess = mobjJson.SetJsonString(strJsonOut)
                        If Not blnSuccess Then strErrMsg = "服务 checkWriteOffData 返回的Json解析错误！JSON=" & strJsonOut
                    Else
                        strErrMsg = strJsonOut
                    End If
                     
                    If blnSuccess Then
                        Set cllTemp = New Collection
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("ivcDate")), "业务日期"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalNum")), "开票数"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalAmt")), "开票金额"
                        cllTemp.Add NVL(mobjJson.GetJsonNodeValue("copyNum")), "总笔数"
                        cllTemp.Add "成功", "返回结果"
                        cllTemp.Add "", "错误原因"
                        cllDatas_Out.Add cllTemp, "_" & cllTemp("业务日期")
                    End If
            
            Case Else
                '服务标识：checkTotalDataByIvcDate
                '入参：
                '   ivcDate 开票日期  String  10  是 格式:yyyy-MM-dd
                '   busType 业务标识    String  20  否  直接填写业务系统内部编码值，由医疗平台配置对照
                '   placeCode   开票点编码  String  50  否  直接填写业务系统内部编码值，由医疗平台配置对照
                '   isReturnList  是否返回列表  String  1 否 0:不返回，1:返回，空:返回
                '   pageNo  页码  Integer 10  是 传入值为1代表第一页,传入值为2代表第二页,依此类推.默认返回的数据是从第一页开始；第二次传入的页码计算由返回值中的总条数/每页条数
                '   pageSize  每页条数  Integer 10  是 每页返回最多返回200条
                '出参：
                '   ivcDate 开票日期  String  10  是
                '   copyNum 总笔数 Integer   是
                '   totalAmt  当前页总金额  Number  18,2  是
                '   totalNum  当前页总开票数 Integer   是
                '   allTotalAmt 所有页总金额  Number  18,2  是
                '   allTotalNum 所有页总开票数 Integer 10  是
                '   billNoList  票号段明细 String  不限  是 详见B-1-1,JSON格式
                strJsonIn = ""
                strJsonIn = strJsonIn & "" & GetJsonNodeString("ivcDate", Format(dtCurrent, "yyyy-MM-dd"), Json_Text)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("busType", "", Json_Text)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("placeCode", str开票点, Json_Text)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("isReturnList", 0, Json_num)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("pageNo", 1, Json_num)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("pageSize", 200, Json_num)
                strJsonIn = "{" & strJsonIn & "}"
                
                 If EInvoiceRequest("checkTotalDataByIvcDate", strJsonIn, blnResult, strJsonOut) Then
                    blnSuccess = mobjJson.SetJsonString(strJsonOut)
                    If Not blnSuccess Then strErrMsg = "服务 checkTotalData 返回的Json解析错误！JSON=" & strJsonOut
                Else
                    strErrMsg = strJsonOut
                End If
                 
                If blnSuccess Then
                    Set cllTemp = New Collection
                    cllTemp.Add NVL(mobjJson.GetJsonNodeValue("ivcDate")), "业务日期"
                    cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalNum")), "开票数"
                    cllTemp.Add NVL(mobjJson.GetJsonNodeValue("allTotalAmt")), "开票金额"
                    cllTemp.Add NVL(mobjJson.GetJsonNodeValue("copyNum")), "总笔数"
                    cllTemp.Add "成功", "返回结果"
                    cllTemp.Add "", "错误原因"
                    cllDatas_Out.Add cllTemp, "_" & cllTemp("业务日期")
                End If
            End Select
        End If
        
        If Not blnSuccess Then
            Set cllTemp = New Collection
            cllTemp.Add Format(dtCurrent, "yyyyMMdd"), "业务日期"
            cllTemp.Add 0, "开票数"
            cllTemp.Add 0, "开票金额"
            cllTemp.Add 0, "总笔数"
            cllTemp.Add "失败", "返回结果"
            cllTemp.Add strErrMsg, "错误原因"
            cllDatas_Out.Add cllTemp, "_" & cllTemp("业务日期")
        End If
        
        dtCurrent = DateAdd("d", 1, dtCurrent)
    Loop
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Public Function clsEInvoiceModule_ZlGetDetailCheckData(ByVal dtBegin As Date, ByVal dtEnd As Date, ByRef cllDatas_Out As Collection, _
    Optional ByVal bytMode As Byte = 1, Optional ByVal str开票点 As String, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取平台电子票据明细数据
    '入参:
    '   dtBegin、dtEnd 业务开始时间、结束时间
    '   bytMode 数据类型：1-开票和退票，2-仅退票
    '   str开票点 平台开票点编码
    '出参:
    '   cllDatas 汇总数据，集合元素(业务日期,业务流水号,开票点,票据种类名称,票据代码,票据号码,开票金额,开票时间,数据类型,关联票据代码,关联票据号码),Key=_业务流水号
    '       其中，数据类型：1-正常电子, 2-电子红票, 3-换开纸质, 4-换开纸质红票, 5-空白纸质
   '                  票据状态：1-正常, 2-冲红, 3-作废
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJsonIn As String, strJsonOut As String, strJsonMain As String
    Dim blnResult As Boolean, lngCount As Long
    Dim lngPageCount As Long, lngPage As Long, lngPageSize As Long
    Dim cllDatas As Collection, clldata As Collection, cllTemp As Collection
    Dim varDataType As Variant, i As Long
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    Set cllDatas_Out = New Collection
    
    '1-费用业务发生时间
    If clsEInvoiceModule_ZLCheckTimeMode = 1 Then
        '服务标识：getBillByBusDate
        '入参：
        '   busBgnDate  业务起始日期  String  17  是 格式：yyyyMMddHHmmssSSS
        '   busEndDate  业务截止日期  String  17  是 格式：yyyyMMddHHmmssSSS
        '   placeCode 开票点编码 String  50  否 直接填写业务系统内部编码值，由医疗平台配置对照，不填写，则查询单位下所有开票据点
        '   dataType  数据类型  String  1 否 1 正常电子、2 电子红票、3 换开纸质、4 换开纸质红票、5空白纸质
        '   pageNo  页码  Integer 10  是 传入值为1代表第一页,传入值为2代表第二页,依此类推.默认返回的数据是从第一页开始；第二次传入的页码计算由返回值中的总条数/每页条数
        '   pageSize  每页条数  Integer 10  是 每页返回最多返回200条
        '出参：
        '   total 总条数 Integer 10  是
        '   pageNo  当前页面  Integer 10  是
        '   billList  开票明细列表  String  不限  是
        '       busDate 业务日期  String  17  是 格式:yyyyMMddHHmmssSSS
        '       busNo 业务流水号 String  50  是
        '       placeCode 开票点编码 String  50  是
        '       billName  票据种类名称  String  100 是
        '       billBatchCode 票据代码  String  50  是
        '       billNo  票据号码  String  20  是
        '       random  校验码 String  20  是
        '       totalAmt  总金额 Number  18,2  是
        '       ivcDateTime 开票时间  String  17  是 yyyyMMddHHmmssSSS
        '       dataType  数据类型  String  1 是 1 正常电子、2 电子红票、3 换开纸质、4 换开纸质红票、5空白纸质
        '       state 状态  String  1 是 1正常 、2作废
        '       relateBillBatchCode 关联电子票据代码  String  50  否
        '       relateBillNo  关联电子票据号码  String  20  否
        lngPageSize = 200
        
        strJsonMain = ""
        strJsonMain = strJsonMain & "" & GetJsonNodeString("busBgnDate", Format(dtBegin, "yyyyMMddHHmmss001"), Json_Text)
        strJsonMain = strJsonMain & "," & GetJsonNodeString("busEndDate", Format(dtEnd, "yyyyMMddHHmmss001"), Json_Text)
        strJsonMain = strJsonMain & "," & GetJsonNodeString("placeCode", str开票点, Json_Text)
        strJsonMain = strJsonMain & "," & GetJsonNodeString("pageSize", lngPageSize, Json_num)
        
        Select Case bytMode
        Case 2
            varDataType = Array(2) '仅取 2-电子红票
        Case Else
            varDataType = Array(1, 2) '仅取 1-正常电子, 2-电子红票
        End Select
        
        For i = 0 To UBound(varDataType)
            lngPage = 1: lngPageCount = 1
            Do While lngPage <= lngPageCount
                strJsonIn = strJsonMain
                strJsonIn = strJsonIn & "," & GetJsonNodeString("dataType", varDataType(i), Json_num)
                strJsonIn = strJsonIn & "," & GetJsonNodeString("pageNo", lngPage, Json_num)
                strJsonIn = "{" & strJsonIn & "}"
                 If Not EInvoiceRequest("getBillByBusDate", strJsonIn, blnResult, strJsonOut) Then
                    strErrMsg_Out = strJsonOut
                    Exit Function
                End If
                If mobjJson.SetJsonString(strJsonOut) = False Then
                    strErrMsg_Out = "服务 getBillByBusDate 返回的Json解析错误！JSON=" & strJsonOut
                    Exit Function
                End If
                
                lngCount = Val(mobjJson.GetJsonNodeValue("total"))
                lngPageCount = lngCount \ lngPageSize + IIf(lngCount Mod lngPageSize = 0, 0, 1)
                lngPage = lngPage + 1
                
                Set cllDatas = mobjJson.GetJsonListValue("billList")
                If Not cllDatas Is Nothing Then
                    For Each clldata In cllDatas
                        Set cllTemp = New Collection
                        cllTemp.Add NVL(clldata("_busDate")), "业务日期"
                        cllTemp.Add NVL(clldata("_busNo")), "业务流水号"
                        cllTemp.Add NVL(clldata("_busType")), "业务标识"
                        cllTemp.Add NVL(clldata("_placeCode")), "开票点"
                        cllTemp.Add NVL(clldata("_billName")), "票据种类名称"
                        cllTemp.Add NVL(clldata("_billBatchCode")), "票据代码"
                        cllTemp.Add NVL(clldata("_billNo")), "票据号码"
                        cllTemp.Add NVL(clldata("_totalAmt")), "开票金额"
                        cllTemp.Add NVL(clldata("_ivcDateTime")), "开票时间"
                        cllTemp.Add NVL(clldata("_dataType")), "数据类型"
                        cllTemp.Add NVL(clldata("_state")), "票据状态"
                        cllTemp.Add NVL(clldata("_relateBillBatchCode")), "关联票据代码"
                        cllTemp.Add NVL(clldata("_relateBillNo")), "关联票据号码"
                        cllDatas_Out.Add cllTemp, varDataType(i) & "_" & clldata("_busNo") '票据类型_业务流水号
                    Next
                End If
            Loop
        Next
        clsEInvoiceModule_ZlGetDetailCheckData = True
        Exit Function
    End If
    
    If gBs_Type.支持版本 = BS_Version.V2_0_3 Then Exit Function
    '0-票据开具时间
    '服务标识: getBillByIvcDate
    '入参:
    '   busBgnDate  起始日期  String  17  是 格式：yyyyMMddHHmmssSSS
    '   busEndDate  截止日期  String  17  是 格式：yyyyMMddHHmmssSSS
    '   placeCode 开票点编码 String  50  否 直接填写业务系统内部编码值，由医疗平台配置对照，不填写，则查询单位下所有开票据点
    '   dataType  数据类型  String  1 否 1 正常电子、2 电子红票、3 换开纸质、4 换开纸质红票、5空白纸质
    '   pageNo  页码  Integer 10  是 传入值为1代表第一页,传入值为2代表第二页,依此类推.默认返回的数据是从第一页开始；第二次传入的页码计算由返回值中的总条数/每页条数
    '   pageSize  每页条数  Integer 10  是 每页返回最多返回200条
    '出参:
    '   total 总条数 Integer 10  是 用于分页
    '   pageNo  当前页码  Integer 10  是
    '   billList  开票明细列表  String  不限  是 详见B-1-1,JSON格式列表
    '       busDate 业务日期  String  17  是 格式:yyyyMMddHHmmssSSS
    '       busNo 业务流水号 String  50  是
    '       placeCode 开票点编码 String  50  是
    '       billName  票据种类名称  String  100 是
    '       billBatchCode 票据代码  String  50  是
    '       billNo  票据号码  String  20  是
    '       random  校验码 String  20  是
    '       totalAmt  总金额 Number  18,2  是
    '       ivcDateTime 开票时间  String  17  是 yyyyMMddHHmmssSSS
    '       dataType  数据类型  String  1 是 1 正常电子、2 电子红票、3 换开纸质、4 换开纸质红票、5空白纸质
    '       state 状态  String  1 是 1正常 、2作废
    '       relateBillBatchCode 关联电子票据代码  String  50  否
    '       relateBillNo  关联电子票据号码  String  20  否
    lngPageSize = 200
    
    strJsonMain = ""
    strJsonMain = strJsonMain & "" & GetJsonNodeString("busBgnDate", Format(dtBegin, "yyyyMMddHHmmss000"), Json_Text)
    strJsonMain = strJsonMain & "," & GetJsonNodeString("busEndDate", Format(dtEnd, "yyyyMMddHHmmss000"), Json_Text)
    strJsonMain = strJsonMain & "," & GetJsonNodeString("placeCode", str开票点, Json_Text)
    strJsonMain = strJsonMain & "," & GetJsonNodeString("pageSize", lngPageSize, Json_num)
    
    Select Case bytMode
    Case 2
        varDataType = Array(2) '仅取 2-电子红票
    Case Else
        varDataType = Array(1, 2) '仅取 1-正常电子, 2-电子红票
    End Select
    
    For i = 0 To UBound(varDataType)
        lngPage = 1: lngPageCount = 1
        Do While lngPage <= lngPageCount
            strJsonIn = strJsonMain
            strJsonIn = strJsonIn & "," & GetJsonNodeString("dataType", varDataType(i), Json_num)
            strJsonIn = strJsonIn & "," & GetJsonNodeString("pageNo", lngPage, Json_num)
            strJsonIn = "{" & strJsonIn & "}"
             If Not EInvoiceRequest("getBillByIvcDate", strJsonIn, blnResult, strJsonOut) Then
                strErrMsg_Out = strJsonOut
                Exit Function
            End If
            If mobjJson.SetJsonString(strJsonOut) = False Then
                strErrMsg_Out = "服务 getBillByBusDate 返回的Json解析错误！JSON=" & strJsonOut
                Exit Function
            End If
            
            lngCount = Val(mobjJson.GetJsonNodeValue("total"))
            lngPageCount = lngCount \ lngPageSize + IIf(lngCount Mod lngPageSize = 0, 0, 1)
            lngPage = lngPage + 1
            
            Set cllDatas = mobjJson.GetJsonListValue("billList")
            If Not cllDatas Is Nothing Then
                For Each clldata In cllDatas
                    Set cllTemp = New Collection
                    cllTemp.Add NVL(clldata("_busDate")), "业务日期"
                    cllTemp.Add NVL(clldata("_busNo")), "业务流水号"
                    cllTemp.Add NVL(clldata("_busType")), "业务标识"
                    cllTemp.Add NVL(clldata("_placeCode")), "开票点"
                    cllTemp.Add NVL(clldata("_billName")), "票据种类名称"
                    cllTemp.Add NVL(clldata("_billBatchCode")), "票据代码"
                    cllTemp.Add NVL(clldata("_billNo")), "票据号码"
                    cllTemp.Add NVL(clldata("_totalAmt")), "开票金额"
                    cllTemp.Add NVL(clldata("_ivcDateTime")), "开票时间"
                    cllTemp.Add NVL(clldata("_dataType")), "数据类型"
                    cllTemp.Add NVL(clldata("_state")), "票据状态"
                    cllTemp.Add NVL(clldata("_relateBillBatchCode")), "关联票据代码"
                    cllTemp.Add NVL(clldata("_relateBillNo")), "关联票据号码"
                    cllDatas_Out.Add cllTemp, varDataType(i) & "_" & clldata("_busNo") '票据类型_业务流水号
                Next
            End If
        Loop
    Next
    clsEInvoiceModule_ZlGetDetailCheckData = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Public Property Get clsEInvoiceModule_ZLCheckTimeMode() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：获取电子票据核对的时间类型
    '说明：0-票据开具时间，1-费用业务发生时间
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_ZLCheckTimeMode = Decode(gBs_Type.支持版本, BS_Version.V2_0_3, 1, BS_Version.V3_1_0, 0, 0)
End Property

Private Function InitPrintSet(ByRef strErrMsg_Out As String) As Boolean
    Dim blnResult As Boolean
    Dim strJson_In As String, strJson_Out As String
    Dim strUrl As String
    
    If mblnInitPrintSet Then InitPrintSet = True: Exit Function '
    strUrl = gBs_Type.URL_Type & "://" & Left(gBs_Type.URL_Address, InStr(1, gBs_Type.URL_Address, "api") - 1)
    
    strJson_In = GetJsonNodeString("url", strUrl, Json_Text)
    strJson_In = strJson_In & "," & GetJsonNodeString("appkey", gBs_Type.签名私钥, Json_Text)
    strJson_In = strJson_In & "," & GetJsonNodeString("type", "medical", Json_Text)
    strJson_In = strJson_In & "," & GetJsonNodeString("testValue", "测试连接", Json_Text)
    strJson_In = strJson_In & "," & GetJsonNodeString("printPreview", "0", Json_Text)
    strJson_In = "{" & strJson_In & "}"
    
    If Not PrintEInvoiceRequest("InitParams", strJson_In, blnResult, strJson_Out) Then
        strErrMsg_Out = strJson_Out
        Exit Function
    End If
    mblnInitPrintSet = True
    InitPrintSet = True
End Function
'
'Public Function HTTP博思(ByVal strUrl As String, ByVal appid As String, ByVal key As String, ByVal noise As String, ByVal strData As String)
'    Dim aHttpRequest        As WinHttp.WinHttpRequest
'    Dim sMethod             As String
'    Dim sBody               As String
'    Dim sResponse           As String
'    Dim strTmp              As String
'    Dim strData_Out As String
'    Dim blnResult As Boolean
'    Dim strGUID As String, strCharSet As String, strNoise As String, strSign As String, strMethod As String
'    Dim strContentType As String, strBody As String, strResponse As String, strVersion As String
'    Dim bytCodeBytes() As Byte
'
'    Dim md5s As New clsMD5
'    Dim arrTmp() As Byte
'    Dim JsonObj As Object               '返回Json对象
'
'    'md5签名,data取BASE64编码后的值
'    strTmp = "appid=" & appid & "&data=" & Encode(strData) & "&noise=" & noise & "&key=" & key & "&version=1.0"
'    arrTmp = strTmp
'
'    strTmp = md5s.DigestByteToHexStr(md5s.Unicode_To_UTF8(arrTmp))
'
'    '组织完整入参
'    sBody = "{""appid"":""" & appid & """,""data"":""" & Encode(strData) & """,""noise"":""" & noise & """,""version"":""1.0"",""sign"":""" & strTmp & """}"
'
'    sMethod = "POST"        '或者(GET)
'
'    ''创建WinHttp.WinHttpRequest
'    Set aHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
'
'    '' 同步接收数据
'    aHttpRequest.Open sMethod, strUrl, False
'    '' 非常重要(忽略错误)
'    aHttpRequest.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
'    '' 其它请求头设置
'    aHttpRequest.SetRequestHeader "Content-Type", "application/json"
'    aHttpRequest.SetRequestHeader "Content-Length", Len(sBody)
'
'    '' 发送
'    aHttpRequest.Send sBody
'
'    '' 得到返回文本(或者是其它)
'    sResponse = aHttpRequest.ResponseText
'    Set aHttpRequest = Nothing
'
'    If InStr(sResponse, "data") < 1 Then
'    '检查data节点是否正常返回
'        HTTP博思 = "{""result"":""9999"",""message"":""data节点返回空""}"
'        Exit Function
'    End If
'
'    bytCodeBytes = zlStr.DecodeBase64(zlStr.JSONParse("data", sResponse), True)
'    If strCharSet = "Unicode" Then
'        strData_Out = zlStr.UnicodeToString(bytCodeBytes)
'    Else
'        strData_Out = zlStr.UTF8BytesToString(bytCodeBytes)
'    End If
'
'    If InStr(strData_Out, "message") < 1 Then
'        '检查message节点是否正常返回
'        'strMsg = "博思message节点返回空"
'        Exit Function
'    End If
'
'    strTmp = zlStr.JSONParse("result", strData_Out)
'    blnResult = (strTmp = "S0000")
'    strTmp = zlStr.DecodeBase64(zlStr.JSONParse("message", strData_Out))
'
'End Function
