VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEInvoice_Test"
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
'二、票据库存管理相关接口
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
'编制:刘兴洪
'日期:2020-03-03 14:11:42
'*********************************************************************************************************************************************
Private mfrmMain As Object, mlngSys As Long, mlngModule As Long
Private mstrOperatorCode As String, mstrOperatorName As String, mstrDBUser As String
Private mstrEInvoiceNodeCode As String
Private mcnOracle As ADODB.Connection
Private Declare Function GetCurrentDirectory Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private mfrmEInvoiceManage As frmEInvoiceManage

Private Function clsEInvoiceModule_zlGetEInvoiceProductName(ByVal frmMain As Object, strProductName As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取电子票据系统名称
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
    strProductName = "中联测试库"
    gstrProductName = strProductName
    clsEInvoiceModule_zlGetEInvoiceProductName = True
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
    Set clsEInvoiceModule_zlGetParaFrom = Nothing
End Function

Private Function clsEInvoiceModule_zlInitialize(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModule As Long, ByVal strOperatorCode As String, ByVal strOperatorName As String, ByVal strDBUser As String, Optional ByVal blnDeviceSet As Boolean, Optional ByRef strEInvoiceNodeCode As String, Optional strExtend As String) As Boolean
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
    mstrEInvoiceNodeCode = strEInvoiceNodeCode
    clsEInvoiceModule_zlInitialize = True
End Function

Private Function clsEInvoiceModule_zlPrintEInvoice(ByVal frmMain As Object, ByVal lngID As Long, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印电子票据接口
    '入参:
    '    frmMain-调用主窗口
    '    lngID-需要作废的电子票据记录ID: 电子票据使用记录.ID
    '    strExtend-扩展参数，暂留，现传为空
    '出参:
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数，暂留，现传为空
    '返回:打印成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '   1.主要是打印电子票据(主要就A4纸票据)
    '调用者:
    '    1.门诊收费（或部分退费）、预交、结帐、体检结帐等开具电子票据成功后，需要调用本接口通知患者"
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If MsgBox("是否打印电子票据接口成功?", vbQuestion + vbYesNo + vbDefaultButton1, "电子票据") = vbNo Then
        strErrMsg_Out = "打印电子票据接口时，选择了“否”"
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

Private Function clsEInvoiceModule_zlCancelBlankInvoice(ByVal frmMain As Object, ByVal lngID As Long, ByVal bytInvoiceType As Byte, ByVal strAuthorName As String, ByVal strReason As String, ByVal strBatchNo As String, ByVal strStartInvoice As String, ByVal strEndInvoice As String, ByVal strHappenTime As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
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
    '    strCreateTime_Out-电子票据生成时间,格式:yyyy-mm-dd HH:MM:SS
    '    strEInvQRCode_Out-电子票据二维码图片数据,base64加密存储,使用时,需解码
    '    strEInvUrl_Out-电子票据URL地址
    '    strEInvUrl1_Out-电子票据URL外网地址
    '    strEinvRemark_Out-电子票据的其他信息
    '    strErrMsg_Out-失败原因
    '    strExtend-扩展参数,JSON格式
    '    strExtend-扩展参数,可为空,返回预交凭证/Json格式
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
    Dim strAppPath As String
    Dim str凭证代码 As String, str凭证号码 As String, str凭证检验码 As String
    
    On Error GoTo errHandle
    
    strAppPath = App.Path & "\文档资料\QRCode.jpg"
    zlStr.PrefixNO
    strSysSouceName_Out = "中联测试库"
    strEInvoiceCode_out = "C" & Format(Now, "yyyymmddHHMMSS")
    strEInvoiceNo_Out = "E" & Format(Now, "yyyymmddHHMMSS") & zlStr.LPAD(Int(Rnd * 1000), 4, "0")
    strCheckCode_out = "X" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    strCreateTime_Out = Format(zlDatabase.Currentdate, "yyyyMMddHHmmss000")
    'base64转码
    strEInvQRCode_Out = zlStr.EncodeBase64_File(strAppPath)
    
    '解码二维码
'    strAppPath = App.Path & "\文档资料\QRCode1.jpg"
'    strAppPath = zlStr.DecodeBase64_File(strEInvQRCode_Out, strAppPath)
    
    strEInvUrl_Out = "http://192.168.0.1/h"
    strEInvUrl1_Out = "http://192.168.0.1/h1"
    strEinvRemark_Out = "测试"
    strErrMsg_Out = ""
    str凭证代码 = "A" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    str凭证号码 = "B" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    str凭证检验码 = "C" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    strExtend = "{""voucherbatchcode"":""" & str凭证代码 & """,""voucherNo"":""" & str凭证号码 & """,""voucherrandom"":""" & str凭证检验码 & """}"
    If MsgBox("是否开票成功?", vbQuestion + vbYesNo + vbDefaultButton1, "电子票据") = vbNo Then
        strErrMsg_Out = "开票提醒时，选择了“否”"
        Exit Function
    End If
    
    clsEInvoiceModule_zlCreateEInvoice = True
     
    Exit Function
errHandle:
     strErrMsg_Out = Err.Description
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
    '    strExtend-扩展参数，暂留，现传为空
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
    Dim strAppPath As String
    On Error GoTo errHandle
    
    strAppPath = App.Path & "\文档资料\QRCode.jpg"
 
    strEInvoiceCode_out = "CT" & Format(Now, "yyyymmddHHMMSS")
    strEInvoiceNo_Out = "ET" & Format(Now, "yyyymmddHHMMSS") & zlStr.LPAD(Int(Rnd * 1000), 4, "0")
    strCheckCode_out = "XT" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    strCreateTime_Out = Format(zlDatabase.Currentdate, "yyyyMMddHHmmss000")
    
    'base64转码
    strEInvQRCode_Out = zlStr.EncodeBase64_File(strAppPath)
    strSysSouceName_Out = "中联测试库"
    '解码二维码
    'strAppPath = App.Path & "\文档资料\QRCode1.jpg"
    'strAppPath = zlStr.DecodeBase64_File(strEInvQRCode_Out, strAppPath)

    
    strEInvUrl_Out = "http://192.168.0.1/h"
    strEInvUrl1_Out = "http://192.168.0.1/h1"
    strEinvRemark_Out = "测试"
    strErrMsg_Out = ""
    strExtend = ""
    If MsgBox("是否作废电子票据成功?", vbQuestion + vbYesNo + vbDefaultButton1, "电子票据") = vbNo Then
        strErrMsg_Out = "作废电子票据提醒时，选择了“否”"
        Exit Function
    End If
    clsEInvoiceModule_zlCancelEInvoice = True
    Exit Function
errHandle:
     strErrMsg_Out = Err.Description
End Function
Private Function clsEInvoiceModule_zlSendEinvoiceMsg(ByVal frmMain As Object, ByVal lngID As Long, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发送电子票据通知信息
    '入参:
    '    frmMain-调用主窗口
    '    lngID-电子票据记录ID
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
    
    If MsgBox("是否发送电子票据通知成功?", vbQuestion + vbYesNo + vbDefaultButton1, "电子票据") = vbNo Then
        strErrMsg_Out = "发送电子票据通知时，选择了“否”"
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
    strInvoiceNO_Out = "F" & Format(Now, "yyyymmddHHMMSS") & zlStr.LPAD(Int(Rnd * 10000), 5, "0")
    If MsgBox("获取下一张有效发票号是否成功?" & vbCrLf & "返回的发票号为" & strInvoiceNO_Out, vbQuestion + vbYesNo + vbDefaultButton1, "电子票据") = vbNo Then
        strErrMsg_Out = "获取下一张有效发票号时，选择了“否”"
        strInvoiceNO_Out = ""
        Exit Function
    End If
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
    
    If MsgBox("是否换开纸质发票接口成功?" & vbCrLf & "当前发票号：" & strInvoiceNO, vbQuestion + vbYesNo + vbDefaultButton1, "电子票据") = vbNo Then
        strErrMsg_Out = "打印电子票据接口时，选择了“否”"
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
    If MsgBox("打印纸质发票接口成功?" & vbCrLf & "当前发票号：" & strInvoiceNO, vbQuestion + vbYesNo + vbDefaultButton1, "电子票据") = vbNo Then
        strErrMsg_Out = "打印纸质发票接口时，选择了“否”"
        Exit Function
    End If
    clsEInvoiceModule_zlPrintPaperInvoice = True
End Function

Private Function clsEInvoiceModule_zlCancelPaperInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, ByVal lngEInvoiceID As Long, ByVal strOperatorCode As String, ByVal strOperatorName As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
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

    If MsgBox("作废纸质票据接口成功?" & vbCrLf & "当前发票号：" & strInvoiceNO, vbQuestion + vbYesNo + vbDefaultButton1, "电子票据") = vbNo Then
        strErrMsg_Out = "作废纸质票据接口时，选择了“否”"
        Exit Function
    End If
    
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

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mcnOracle = Nothing
    Set mfrmMain = Nothing
    If Not mfrmEInvoiceManage Is Nothing Then Unload mfrmEInvoiceManage: Set mfrmEInvoiceManage = Nothing
End Sub

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
    '   cllDatas 汇总数据，集合元素(业务日期,业务类型,业务流水号,开票点,票据种类名称,票据代码,票据号码,开票金额,开票时间,数据类型,关联票据代码,关联票据号码),Key=_业务流水号
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
End Function

Public Property Get clsEInvoiceModule_ZLCheckTimeMode() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：获取电子票据核对的时间类型
    '说明：0-票据开具时间，1-费用业务发生时间
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_ZLCheckTimeMode = 0
End Property

