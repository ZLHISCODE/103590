VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Version As String
Private mfrmMain As Object
Private mlngModul As Long
Private mstrDBUser As String   '当前数据库用户
Public mblnYLMgr As Boolean '当前模块是否是医疗卡管理
Private mobjPubThreeInterface As clsThreeInterface
Private mcnOracle As ADODB.Connection    '本地模块用
Private mlngfrmMainHwnd As Long

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, strDBUser As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据主程序指定功能，调用执行相关程序
    '入参:lngSys : 系统编号
    '     lngModul:需要执行的功能序号
    '     cnMain:主程序的数据库连接
    '     frmMain:主窗体
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-11-17 16:35:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmList As frmSquareCardManager, frmTemp As Form, strTitle As String
    Dim frmPatiCard As frmPatiCureCardMgr
    
    Set gfrmMain = frmMain
    If Not CheckValid Then Exit Sub
    Set gcnOracle = cnMain: gstrDBUser = strDBUser
    Set mcnOracle = cnMain: mstrDBUser = strDBUser: mlngModul = lngModul   '52689
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrUnitName = GetUnitName()
    glngModul = lngModul: glngSys = lngSys
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"

    Call mobjPubThreeInterface.InitCommon(cnMain)
    '初始系统参数值
    Call zlinitSystemPara
    Call CreateSquareCardObject(mfrmMain, lngModul)
    '-------------------------------------------------
    '1503 -消费卡管理
    '1505-医疗卡类别管理
    '1107-医疗卡发放管理
    '1152-一卡通管理工具
    On Error Resume Next
    Select Case lngModul
    Case 1107  '医疗卡发放管理
'        Call zlInitPatiCards '冉俊明,2014-5-9
        strTitle = "医疗卡发放管理"
        Call CreatePublicExpenseObject(lngModul)
         For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmPatiCard = New frmPatiCureCardMgr
        frmPatiCard.ShowList lngModul, strTitle, frmMain
    Case 1503 '消费卡管理
        Call zlInitCards
        strTitle = "消费卡管理"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmSquareCardManager
        frmList.ShowList lngModul, strTitle, frmMain
        
    Case 1505 '医疗卡类别管理
        frmPatiCureCardTypeMgr.Show , frmMain
    Case 1152 '一卡通管理工具
        Call zlThreeInterFaceMgrTool(frmMain, lngModul)
    Case Else
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前数据库登录用户名
    Dim frmList As frmSquareCardManager, frmTemp As Form, strTitle As String
    Dim frmPatiCard As frmPatiCureCardMgr
    
    Set gfrmMain = Nothing
    If Not CheckValid Then Exit Sub
    Set gcnOracle = cnMain: gstrDBUser = strDBUser
    Set mcnOracle = cnMain: mstrDBUser = strDBUser: mlngModul = lngModul   '52689
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrUnitName = GetUnitName()
    glngModul = lngModul: glngSys = lngSys
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    
    Call mobjPubThreeInterface.InitCommon(cnMain)
    '初始系统参数值
    Call zlinitSystemPara
    Call CreateSquareCardObject(mfrmMain, lngModul)
    '-------------------------------------------------
    '1503 -消费卡管理
    '1505-医疗卡类别管理
    '1107-医疗卡发放管理
    Select Case lngModul
        Case 1107 '医疗卡发放管理
'         Call zlInitPatiCards
         strTitle = "医疗卡发放管理"
        Call CreatePublicExpenseObject(lngModul)
         For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmPatiCard = New frmPatiCureCardMgr
        frmPatiCard.ShowList lngModul, strTitle, lngMain
    Case 1503 '消费卡管理
        Call zlInitCards
        strTitle = "消费卡管理"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmSquareCardManager
        frmList.ShowList lngModul, strTitle, lngMain
    Case 1505 '医疗卡类别管理
        os.ShowChildWindow frmPatiCureCardTypeMgr.hWnd, lngMain
    End Select
End Sub

Private Sub Class_Initialize()
    Dim intDebug As Integer
    glngInstanceCount = glngInstanceCount + 1
    Set mobjPubThreeInterface = New clsThreeInterface
    
    Version = App.Major & "." & App.Minor & "." & App.Revision
    intDebug = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard", "TestCardNO", 0))
    gblnTestCardNo = IsDesinMode Or intDebug = 1
    mblnYLMgr = False
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    If Not gfrmCardMgr Is Nothing Then CloseWindows = True: Exit Function
    If gblnNotCloseWindows Then CloseWindows = True: Exit Function
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
End Function
Private Sub Class_Terminate()
    Dim i As Long
    
    Err = 0: On Error Resume Next
    Set mobjPubThreeInterface = Nothing
    Set mcnOracle = Nothing
    Set mfrmMain = Nothing

    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
     
    If mblnYLMgr Then Exit Sub '如果是医疗卡管理本身则不销毁数据
    Call zlReleaseResources
  
    
End Sub
Public Function zlInitComponents(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    ByVal cnOracle As ADODB.Connection, _
    Optional blnDeviceSet As Boolean = False, _
    Optional strExpand As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参: frmMain-调用的主窗体
    '        lngModule-HIS调用模块号
    '       lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       blnDeviceSet-设备设置调用初始化
    '       strExpand-扩展信息(可选传入:卡类别ID-不传时,表示全部初始化,传入时,只初始化指定的接口)
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 11:49:39
    '说明:如果是第三方接口主动读取数据,则在渠道更改的部件中传入连接.
    '    1.如果第三接口存在签到的情况，可以在此接口中进行签道.
    '    2.如果第三方接口读取数据,需要先初始化.
    '调用者:
    '    1.  进入门诊收费、挂号时调用本接口
    '    2.  进入医生站、护士站、医技站等时调用本接口
    '    3.  第三方接口需要主动读取数据时,必须先初始化.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    glngSys = lngSys: gstrDBUser = strDBUser: Set mfrmMain = frmMain
    mstrDBUser = strDBUser: mlngModul = lngModule     '52689
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    
    Set mcnOracle = cnOracle
    If gcnOracle Is Nothing Then Set gcnOracle = cnOracle
    Call mobjPubThreeInterface.InitCommon(cnOracle)
    Call zl_GetUserInfo(cnOracle): Call zlinitSystemPara(cnOracle)
    zlInitComponents = mobjPubThreeInterface.zlInitComponents(frmMain, lngModule, glngSys, strDBUser, cnOracle, blnDeviceSet, strExpand)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCardDevSet(ByVal frmMain As Object, Optional bytCardType As Byte = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlCardDevSet(设备参数设置接口):主要是配置相关的设备的相关参数
    '入参:frmMain Form    HIS传入 调用的主窗体
    '       bytCardType- 1-消费卡;2-医疗卡;0-不区分消费卡和医疗卡,统一设置
    '出参:
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    'HIS调用说明.
    '    1.  在门诊收费管理的参数设置'设备配置' zlCardDevSet
    '    2.  在住院结帐界面的参数设置'设备配置' zlCardDevSet
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCardDevSet = mobjPubThreeInterface.zlCardDevSet(frmMain, bytCardType)
End Function

Public Function zlGetIDKindStr(Optional strIDKindStr As String = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取读卡的种类
    '入参:strIDKindStr    String  IN
    '有两种格式:
    '一种是缺省的:短名1|全名1|读卡标志1;…. ;短名n|全名n|读卡标志n
    '另一种是该接口返回的格式:短名|全名|读卡标志|卡类别ID|卡号长度;…
    '出参:
    '返回: 短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '           是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
    '        其中:卡类别ID|长度是本次增加的,由调用者根据情况来确认.
    '       比如:身|身份证号|0|0|18;IC|IC卡号|1|0|8;门|门诊号|0|0|0;就|就诊卡|0|0|0;建|建行卡|0|0|10
    '      出现错误时,返回空
    '编制:刘兴洪
    '日期:2011-06-14 14:43:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetIDKindStr = GetIDKindStr(strIDKindStr, , Me)
End Function

Public Function zlAdviceSendBulidBarCode(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal intRecNature As Integer, ByVal strNos As String, Optional ByVal strExpand As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医嘱发送生成条码接口
    '入参:frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  医疗卡类别ID
    '    intRecNature    Integer In  记录性质
    '    strNos  String  In  单据号，多个用逗号分隔
    '    strExpand   String  In/Out  扩展参数,暂留，现为空
    '返回:只要一个接口调用返回True:否则返回False
    '编制:刘兴洪
    '日期:2016-03-22 13:35:10
    '说明：
    '１. 在医嘱发送并生成费用后打印指引单或处方单前调用此接口
    '２. 接口内部根据记录性质和单据号，记录"二维码"或"条码"信息。
    '３. 操作形式是：医嘱发送后根据接口生成的二维码打印在指引单或处方上，然后通过移动平台扫描二维码进行支付。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlAdviceSendBulidBarCode = mobjPubThreeInterface.zlAdviceSendBulidBarCode(frmMain, lngModule, lngCardTypeID, intRecNature, strNos, strExpand)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetAvailabilityCardType() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前工作站有效的支付卡,比如银行卡;消费卡等
    '返回:格式:短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式;…
    '编制:刘兴洪
    '日期:2011-06-14 15:16:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetAvailabilityCardType = GetAvailabilityCardType()
End Function

Public Function zlGetCards(ByVal bytType As Byte) As Cards
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的卡对象
    '入参:bytType-0-所有医疗卡;
    '             1-启用的医疗卡,
    '             2-所有存在三方账户的三方卡
    '             3-启用的三方账户的医疗卡
    '出参:
    '返回:成功返回本地支持的卡对象
    '编制:刘兴洪
    '日期:2013-10-23 17:43:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetCards = mobjPubThreeInterface.zlGetCards(bytType)
End Function
 
Public Function zlGetIDKindCardTypeID(ByVal strIDKindStr As String, ByVal strIDKind As String, _
    ByRef lngCardTypeID As Long, ByRef lngCardLen As Long, _
    Optional ByRef strKindName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:主要是根据上IDKind部件中的IDKind来获取相应卡类别ID,以便查找相关的病人信息:
    '入参:strIDKindStr  -缺省的StrIDKindStr:格式:短名|全名|读卡标志|卡类别ID|卡号长度;…如
    '       身|身份证号|0|0|18;IC|IC卡号|1|0|8;门|门诊号|0|0|0;就|就诊卡|0|0|0;建|建行卡|0|0|10
    '       strIDKind-可以为缺省的名称;也可以为索引(索引从0…N):用名称,重复时,指向第一个.
    '出参:lngCardTypeID- 卡类别ID
    '       lngCardLen- 卡号长度
    '       strKindName-名称
    '返回:Boolean 返回    成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-06-14 16:11:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetIDKindCardTypeID = GetIDKindCardTypeID(strIDKindStr, strIDKind, lngCardTypeID, lngCardLen, strKindName)
End Function

Public Function zlGetPatiID(ByVal strCardType As String, _
    ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, _
    Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, _
    Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, _
    Optional objCtl As Object = Nothing, _
    Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, _
    Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, _
    Optional ByRef blnUserCancel As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的医疗类别和卡号,获取对应的病人ID
    '入参:strCardType-卡类别,如果为数字,这为卡类别ID,如果为字符,则为类别名称
    '        strCardNo-卡号
    '        blnNotShowErrMsg-不显示错误的提示信息
    '        blnOnlyContractPati-仅读取签约病人
    '        blnblnCertificate-根据证件类型医疗卡读取病人信息
    '        blnUserCancel-选择器中，用户选择了取消
    '出参:strErrMsg-返回的错误信息
    '        lng病人ID-返回的病人ID
    '        strCardPass-返回的解密后的密码
    '        lngCardTypeID-返回指定的卡类别ID:47522
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiID = GetPatiID(strCardType, strCardNo, blnNotShowErrMsg, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, mcnOracle, blnCertificate, blnUserCancel)
    
End Function

Public Function zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNos As String, _
    Optional ByVal strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       strCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       dblMoney-支付金额(退款时为负数)
    '       strNos-本次支付所涉及的单据
    '       strXMLExpend-(XML串:验证密码:自助机用)
    '出参:
    '   strXMLExpend-(XML串:错误信息)
    '返回:扣款合法,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 16:42:43
    '说明:
    '   在调用扣款前，由于存在Oracle事务问题， 所以再调用扣款交易前， _
    '   先进行数据的合法性检查,以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlPaymentCheck = mobjPubThreeInterface.zlPaymentCheck(frmMain, lngModule, strCardTypeID, bln消费卡, strCardNo, dblMoney, strNos, strXMLExpend)
End Function

Public Function zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, ByVal strPrepayNos As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:13:48
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlPaymentMoney = mobjPubThreeInterface.zlPaymentMoney(frmMain, lngModule, lngCardTypeID, bln消费卡, strCardNo, strBalanceIDs, strPrepayNos, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor)
End Function

Public Function zlBrushCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByRef dbl金额 As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varSquareBalance As Variant, _
    Optional ByVal bln转预交 As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str费用来源 As String, _
    Optional ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定支付类别,弹出刷卡窗口
    '入参:rsClassMoney:收费类别,金额
    '        lngCardTypeID-为零时,为老一卡通刷卡
    '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
    '       dblBrushTotaled-消费有效,表示已经刷消费卡总额(主要用于多次刷卡)
    '       str上次限制类别-上次刷消费时的限制类别(同次多次刷消费卡时,需要检查本次刷卡类别与上次类别是否一致,不一致不允许刷卡消费)
    '       varSquareBalance- Collection类型,当前已经刷卡的信息(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文 ))
    '       bln预交-是否转预交
    '       blnAllPay-是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
    '       strXmlIn-XML入参,目前格式如下:
    '       <IN>
    '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
    '       </IN>
    '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '       lng病人ID - 病人ID(使用消费卡支付时传入)
    '出参:str限制类别-限制类别(消费卡返回)
    '        lng消费卡ID-消费卡信息.ID(消费卡返回)
    '       strCardNO-返回刷卡的卡号
    '       strPassWord-返回刷卡所对应的密码
    '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:18:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBrushCard = mobjPubThreeInterface.zlBrushCard(frmMain, lngModule, rsClassMoney, lngCardTypeID, bln消费卡, _
        strPatiName, strSex, strOld, dbl金额, strCardNo, strPassWord, bln退费, blnShowPatiInfor, bln退现, _
        bln余额不足禁止, varSquareBalance, bln转预交, blnAllPay, strXmlIn, str费用来源, lng病人ID)
End Function

Public Function zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户回退交易前的检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额(正数)
    '       strSwapNo-交易流水号(退款时检查)
    '       strSwapMemo-交易说明(退款时传入)
    '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
    '返回:退款合法,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:24:55
    '说明:
    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
    '    以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlReturncheck = mobjPubThreeInterface.zlReturncheck(frmMain, lngModule, lngCardTypeID, bln消费卡, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strXMLExpend)
End Function

Public Function zlReturnCashCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退现交易检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(退款时检查)
    '       strSwapMemo-交易说明(退款时传入)
    '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
    '返回:退现合法,返回true,否则返回Flase
    '编制:刘尔旋
    '日期:2017-7-3
    '说明:
    '    1.使用模块：病人结帐管理
    '    2.触发时机:在结帐列表中点“X"时，调用，如果返回true,则允许退现，否则禁止
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlReturnCashCheck = mobjPubThreeInterface.zlReturnCashCheck(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strXMLExpend)
End Function

Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
    ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额(正数)
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExtendInfor-本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '出参: strSwapNo-交易流水号(退款交易流水号)
    '      strSwapMemo-交易说明(退款交易说明)
    '      strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:31:32
    '说明:
    '       在所有需要对已经消费的地方进行退款的情况下，调用该接口,目前规划在:收费室；挂号等。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlReturnMoney = mobjPubThreeInterface.zlReturnMoney(frmMain, lngModule, lngCardTypeID, bln消费卡, strCardNo, strBalanceIDs, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor)

End Function

Public Function zlReturnMultiMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strInXML As String, _
    ByVal lng冲销ID As Long, ByRef strOutXml As String, ByRef strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易(多笔回退)
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strInXML-XML串:
    '       <JSLIST>
    '           <JS>
    '               <KH>卡号</KH>
    '               <JYLSH>交易流水号</JYLX>
    '               <JYSM>交易说明</JYSM>
    '               <ZFJE>作废金额</ZFJE>
    '               <JSLX>类型</JSLX>  //1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '               <ID></ID>    //类型=1时,预交ID;类型=2时，为原结帐ID
    '           </JS>
    '       </JSLIST>
    '       lng冲销ID-作废时的冲销ID(作废时或退费时有效，否则为0）
    '       strExpend-无（暂留，待以后扩展)
    '出参:
    '     strOutXML-返回XML串
    '       <JSLIST>
    '           <JS>
    '               <KH>卡号</KH>
    '               <TKLSH>退款交易流水号</TKLSH>
    '               <TKSM>退款交易说明</TKSM>
    '               <ID></ID>
    '           </JS>
    '       </JSLIST>
    '      strExpend-交易的扩展信息
    '       <Expends>
    '           <Expend>
    '               <XMMC>项目名称1</XMMC>
    '               <XMNR>项目内容2</XMNR>
    '           </Expend>
    '       </Expends>
    '返回:函数返回    True:调用成功,False:调用失败
    '日期:2015-11-10
    '说明:
    '   目前只有结帐程序时有效（结帐退款),用于一次性处理同一卡类别的多笔三方交易作废
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlReturnMultiMoney = mobjPubThreeInterface.zlReturnMultiMoney(frmMain, lngModule, lngCardTypeID, bln消费卡, strInXML, lng冲销ID, strOutXml, strExpend)
End Function

Public Function zlSquareAffirm(ByVal frmMain As Object, _
    ByVal lngModule As Long, strPrivs As String, _
    Optional ByVal lngPatiID As Long = 0, _
    Optional ByVal lngCardTypeID As Long = 0, _
    Optional ByVal bln消费卡 As Boolean = False, _
    Optional ByVal bytBillType As Byte, _
    Optional ByVal strNos As String = "", _
    Optional ByVal str医嘱IDs As String = "", _
    Optional ByRef strExpand As String = "", _
    Optional ByRef lng结帐ID As Long = 0, _
    Optional ByVal bln使用预交 As Boolean = True) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 消费确认接口 , 主要是应用于病人在各消费环境进行消费确认
    '入参:frmMain-传入调用对象
    '       lngModule:调用的模块号
    '       strPrivs:权限串
    '       lngPatiID :病人ID,可以不传,在本接口窗体中刷卡!
    '       bytBillType:单据类别: 1-收费记录;2-记帐记录；4-挂号记录;5-卡费;10-缴预交
    '       strNOs:格式为( 单据1,单据2),配合BytBillType单据类型使用.一次只能使用一种性质
    '                   如:  A0001,A002,A003…;
    '       str医嘱IDs:格式为:ID1,ID2,...
    '       strCardNO-主界面中刷的卡号
    '       strExpand-扩展参数,参数用|线分隔:
    '                       第一位:表示是否诊间支付(1-表示诊间支付;0-非诊间支付), _
    '                       第二位后暂无
    '       bln使用预交-是否允许使用预交：Ture，允许使用预交款，且存在预交款时缺省使用预交款；False，不允许使用预交款，必须要有启用的三方帐户
    '出参:
    '返回:Boolean 返回    成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-06-15 09:53:37
    '说明:
    '      如果strNos和str医嘱IDs都没传,只是对指定病人的门诊收费划价单收费和门诊记帐划价进行审核.
    '      如果病人ID不传入,则需要在窗体中先进行刷卡找到病人后,再进行消费确认.
    '      如果是诊间支付的,则需要检查zl_GetThreeCardtypeID来获取卡类别ID,>0表示不用再刷卡(弹出界在),直接调用三方接口进行支付
    '      经咨询曾超,肯定是HIS这边在用,所以不用区分连接
    '调用者:
    '    1.  检查;检验;药房等.
    '    2.  其他所有需要进行消费确认的地方都应该调用该接口.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivsTemp As String, strSQL As String
    Dim varData As Variant, blnCliniqueRoomPay As Boolean
    Dim frmAffirm  As New frmSquareAffirm
    On Error GoTo errHandle
    
    If lngPatiID = 0 Then Exit Function
    
    '第一位:诊间支付
    varData = Split(strExpand & "|||||", "|")
    glngSys = zlGetSystemNo(glngSys)
    strPrivsTemp = GetPrivFunc(glngSys, 1151)
    If strPrivsTemp = "" Then
        MsgBox "一卡通消费操作的权限不足,请与系统管理员联系!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    blnCliniqueRoomPay = Val(varData(0)) = 1    '69192
    strPrivsTemp = ";" & strPrivsTemp & ";"
    zlSquareAffirm = frmAffirm.zlSquareAffirm(frmMain, lngModule, _
        strPrivsTemp, lngPatiID, lngCardTypeID, bln消费卡, _
        blnCliniqueRoomPay, bytBillType, strNos, str医嘱IDs, strExpand, lng结帐ID, bln使用预交)
    If Not frmAffirm Is Nothing Then
        Unload frmAffirm
    End If
    Set frmAffirm = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlCliniqueRoomPayPrintSet(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:诊间支付打印设置
    '入参:frmMain-调用的主窗体
    '编制:刘兴洪
    '日期:2014-02-25 13:42:09
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If frmSquareAffirmParaSet.SetPara(frmMain) = False Then Exit Sub
End Sub


Public Function zlIsAllowCliniqueRoomPay(ByVal lngModule As Long, _
    ByVal lng病人ID As Long, ByVal lng医嘱ID As Long, _
    ByRef lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:当前病人是否允许支持诊间支付
    '入参:lngModule-调用的模块号(待用)
    '       lng病人ID-指定的病人ID
    '       lng医嘱ID:医嘱ID
    '出参:lngCardTypeID-返回三方账户支付的卡类别ID,函数返回True时有效
    '返回:返回true表示允许诊间支付,否则返回False不允许支付，lngCardTypeID为零
    '编制:刘兴洪
    '日期:2014-01-14 17:14:14
    '说明:
    '   当接口返回true时，需要将lngCardTypeID传给zlSquareAffirm接口的lngCardTypeID
    '   一般医生站用,不区分连接
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strPrivsTemp As String
    
    On Error GoTo errHandle
    lngCardTypeID = 0
    glngSys = zlGetSystemNo(glngSys)
    strPrivsTemp = GetPrivFunc(glngSys, 1151)
    '无一卡通消费权限，直接返回false
    If strPrivsTemp = "" Then Exit Function
    strSQL = "Select  Zl_Get_Threecardtypeid([1],[2],[3]) as 卡类别ID from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取诊间支付的卡类别ID", lngModule, lng病人ID, lng医嘱ID)
    If rsTemp.EOF Then Exit Function
    lngCardTypeID = Val(Nvl(rsTemp!卡类别id))
    If lngCardTypeID = 0 Then Exit Function
    zlIsAllowCliniqueRoomPay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlReadCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal blnOlnyCardNO As Boolean, _
    ByVal strExpand As String, _
    ByRef strOutCardNo As String, _
    ByRef strOutPatiInforXml As String, _
    Optional ByRef strPhotoFile As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读卡接口
    '入参:frmMain-调用的父窗口
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       blnOlnyCardNO-仅仅读取卡号
    '       strExpand-扩展参数
    '出参:strOutCardNO-返回的卡号
    '       strOutPatiInforXML-(病人信息返回.XML串)
    '       strPhotoFile-当从卡中读取取像片后,传给HIS的像片文件(含完整路径)
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 11:56:41
    '说明:
    '   １. 在所有绑定卡时,需要建立病人信息档案时，读取该接口
    '   ２. 在所有需要刷卡进行就诊的地方，都需要调用该接口
    'strOutPatiInforXML参数说明
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    '    姓名    Varchar2    100
    '    性别    Varchar2    4
    '    年龄    Varchar2    10
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '    出生地点    Varchar2    50
    '    身份证号    VARCHAR2    18
    '    其他证件    Varchar2    20
    '    职业    Varchar2    80
    '    民族    Varchar2    20
    '    国籍    Varchar2    30
    '    学历    Varchar2    10
    '    婚姻状况    Varchar2    4
    '    区域    Varchar2    30
    '    家庭地址    Varchar2    50
    '    家庭电话    Varchar2    20
    '    家庭地址邮编    Varchar2    6
    '    监护人  Varchar2    64
    '    联系人姓名  Varchar2    64
    '    联系人关系  Varchar2    30
    '    联系人地址  Varchar2    50
    '    联系人电话  Varchar2    20
    '    工作单位    Varchar2    100
    '    单位电话    Varchar2    20
    '    单位邮编    Varchar2    6
    '    单位开户行  Varchar2    50
    '    单位帐号    Varchar2    20
    '   ABO血型 Varchar2    10
    '   RH  Varchar2    10
    '   哮喘标志    Number  2
    '   心脏病标志  Number  2
    '   心脑血管病标志  Number  2
    '   癫痫病标志  Number  2
    '   凝血紊乱标志    Number  2
    '   糖尿病标志  Number  2
    '   青光眼标志  Number  2
    '   透析标志    Number  2
    '   器官移植标志    Number  2
    '   器官缺失标志    Number  2
    '   可装卸义肢标志  Number  2
    '   心脏起搏器标志  Number  2
    '   其他医学警示    Varchar2    100
    '   联系信息
    '       姓名    Varchar2    30
    '       关系    Varchar2    30
    '       电话    Varchar2    30
    '       身份证号    Varchar2    30
    '   健康档案编号    Varchar2    18
    '   新农合证号  Varchar2    18
    '   其他证件
    '       信息名  Varchar2    15
    '       信息值  Varchar2    100
    '   其他信息 可以返回上面接点不存在的数据
    '       信息名  Varchar2    20      例: 工作单位, 血压等
    '       信息值  Varchar2    50      比如:信息名=血压;信息值='XXmmHg'
    '   过敏情况        30      以下信息存放于病人过敏药物表中
    '       药物名称    Varchar2    100
    '       药物反应    Varchar2    101
    '   免疫记录
    '       疫苗名称    Varchar2    100
    '       接种时间    Varchar2    8
    '   医疗卡属性
    '       信息名  Varchar2    20
    '       信息值  Varchar2    50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlReadCard = mobjPubThreeInterface.zlReadCard(frmMain, lngModule, lngCardTypeID, blnOlnyCardNO, strExpand, strOutCardNo, strOutPatiInforXml, strPhotoFile)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlMakeCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal stInputPatiInforXml As String, _
    ByRef strPhotoFile As String, _
    ByRef strOutPatiInforXml As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:制卡接口
    '入参:
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  卡类别ID
    '    stInputPatiInforXml String(XML) In  HIS传入已经健档的病人信息
    '    strPhotoFile    String  In  表示像片文件(含完整路径)
    '    strExpend   String  In   XML,暂无,待以后扩展
    '出参:
    '    strOutPatiInforXml  String(XML) Out 调用接口后,传给HIS系统已经制卡后的病人信息,以便建立病人档案
    '    strExpend   String  Out  XML,暂无,待以后扩展
    '    strPhotoFile    String  In  表示像片文件(含完整路径)
    '返回 :True:调用成功,False:调用失败
    '调用者:
    '    目前只在医疗卡发放管理中调用此接口(主要是考虑输入的信息较多,其他窗口不适应)，相关的调用流程医疗卡发放管理的制卡
    '  strInputPatiInforXML及strOutPatiInforXML格式
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    '    姓名    Varchar2    100
    '    性别    Varchar2    4
    '    年龄    Varchar2    10
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '    出生地点    Varchar2    50
    '    身份证号    VARCHAR2    18
    '    其他证件    Varchar2    20
    '    职业    Varchar2    80
    '    民族    Varchar2    20
    '    国籍    Varchar2    30
    '    学历    Varchar2    10
    '    婚姻状况    Varchar2    4
    '    区域    Varchar2    30
    '    家庭地址    Varchar2    50
    '    家庭电话    Varchar2    20
    '    户口邮编    Varchar2    6
    '    监护人  Varchar2    64
    '    联系人姓名  Varchar2    64
    '    联系人关系  Varchar2    30
    '    联系人地址  Varchar2    50
    '    联系人电话  Varchar2    20
    '    工作单位    Varchar2    100
    '    单位电话    Varchar2    20
    '    单位邮编    Varchar2    6
    '    单位开户行  Varchar2    50
    '    单位帐号    Varchar2    20
    '    病人ID  Number  18      可以为空
    '    ABO血型 Varchar2    10      信息名='ABO'
    '    信息值=(A型;B型;O型;AB型;不详)
    '    RH  Varchar2    10      信息名='RH'
    '    信息值=(阴;阳;不详;未查)
    '    哮喘标志    Number  2       信息名='医学警示'
    '    信息值=' 哮喘;心脏病; 心脑血管病;…'
    '    心脏病标志  Number  2
    '    心脑血管病标志  Number  2
    '    癫痫病标志  Number  2
    '    凝血紊乱标志    Number  2
    '    糖尿病标志  Number  2
    '    青光眼标志  Number  2
    '    透析标志    Number  2
    '    器官移植标志    Number  2
    '    器官缺失标志    Number  2
    '    可装卸义肢标志  Number  2
    '    心脏起搏器标志  Number  2
    '    其他医学警示    Varchar2    100
    '    联系信息
    '        姓名    Varchar2    30
    '        关系    Varchar2    30
    '        电话    Varchar2    30
    '        身份证号    Varchar2    30
    '    健康档案编号    Varchar2    18
    '    新农合证号  Varchar2    18
    '    其他证件
    '        信息名  Varchar2    15
    '        信息值  Varchar2    100
    '    其他信息 可以返回上面接点不存在的数据
    '        信息名  Varchar2    20      例: 工作单位, 血压等
    '        信息值  Varchar2    50      比如:信息名=血压;信息值='XXmmHg'
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
 
    zlMakeCard = mobjPubThreeInterface.zlMakeCard(frmMain, lngModule, lngCardTypeID, stInputPatiInforXml, _
        strPhotoFile, strOutPatiInforXml, strExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlBandCardArfter(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  卡类别ID
    '    lng病人ID   Long    IN   病人ID
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用者:
    '    1.  医疗卡发放管理中绑定卡后调用
    '    2.  病人入院登记;病人挂号管理;病人信息登记中进行卡绑定时调用
    '其他说明:
    '    相关的调用流程见后续各模块的绑定卡的详细说明
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle

    zlBandCardArfter = mobjPubThreeInterface.zlBandCardArfter(frmMain, lngModule, lngCardTypeID, _
     lng病人ID, strExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlMzInforWriteToCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    ByVal lngBalanceID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写门诊信息接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  传入写卡类别ID:
    '           1)传入刷卡的类别ID
    '           2)传入零时,需要选择某个卡类别ID
    '    lng病人ID   Long    In  病人ID
    '    lngBalanceID    Long    In  结算序号(某次结算的序号)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用时机:
    '         医疗卡类别.是否写卡=1才调用
    '调用者:
    '    1.  目前收费管理完成后,调用该接口
    '    2.  退费完成后,调用该接口
    '说明:
    '        门诊相关信息的写入 , 主要是渠道编制的接口内部写入
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlMzInforWriteToCard = mobjPubThreeInterface.zlMzInforWriteToCard(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, lngBalanceID, strExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlZyInforWriteToCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    ByVal lngBalanceID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写住院信息接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  传入写卡类别ID:
    '           1)传入刷卡的类别ID
    '           2)传入零时,需要选择某个卡类别ID
    '    lng病人ID   Long    In  病人ID
    '    lngBalanceID    Long    In  结帐ID(可以不传入)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用时机:
    '        医疗卡类别.是否写卡=1才调用
    '调用者:
    '    1.  出院时
    '    2.  结帐时
    '    3.  结帐作废时
    '说明:
    '       住院相关信息的写入 , 主要是渠道编制的接口内部写入
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlZyInforWriteToCard = mobjPubThreeInterface.zlZyInforWriteToCard(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, lngBalanceID, strExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlHealthArchiveIsSHow(frmMain As Object, _
    ByVal lngModule As Long, _
    ByRef strFunName As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:居民健档档案
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    strFunName  String  Out 返回功能名, 即在医生工作站的菜单中显示的名称(菜单位置放在"工具"菜单下)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    '        Boolean 函数返回    True:表示应该加入档案显示菜单,False:不加入档案菜单
    '调用者:
    '    1.  门诊(住院)医生工作站调用,医生站
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlHealthArchiveIsSHow = mobjPubThreeInterface.zlHealthArchiveIsSHow(frmMain, lngModule, strFunName, strExpend)
End Function

Public Function zlHealthArchivesShow(frmMain As Object, _
    ByVal lngModule As Long, ByVal lng病人ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:居民健档档案
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lng病人ID-病人ID
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    医生工作站不做返回判断,渠道编制的接口部件中可以返回值
    ' 调用者:门诊(住院)医生工作站后调用
    '说明:     该接口分两层:
    '    1.  一层是在zl9CardSquare部件中该接口不返回值,由医生站直接调用,该接口调用渠道编制的接口界面,否则按标准的健康卡档案界面显示(见下图)
    '    2.  二层是渠道编制的接口部件:该接口允许返回值为TRUE时,表示利用渠道自己编制的界面展示,否则将按HIS固定的健康档案信息显示(见下图)
    '调用者:
    '    1.  门诊(住院)医生工作站调用,医生站
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlHealthArchivesShow = mobjPubThreeInterface.zlHealthArchivesShow(frmMain, lngModule, lng病人ID, strExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlInitInterFacel(ByVal lngCardTypeID As Long, _
    Optional bln消费卡 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化指定卡接口
    '入参:lngCardTypeID-指定卡类别
    '       bln消费卡-是否消费卡
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 15:29:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlInitInterFacel = mobjPubThreeInterface.zlInitInterFacel(lngCardTypeID, bln消费卡)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSendCard(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, Optional lng病人ID As Long = 0, Optional ByVal strPrivs As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医疗卡发卡编辑
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        lngCardTypeID-医疗卡类别
    '        lng病人ID-传入指定的病人ID,对指定病人进行发卡
    '        strPrivs-可以传入空，传入空时取医疗卡管理模块的相关权限
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 12:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static blnNoFirst As Boolean  '是否第一次调用,初始化全局变量
    Dim rsTemp As ADODB.Recordset
    
    '不使用本地连接,因为这个肯定在HIS中应用
    gstrPrivs = strPrivs
    '问题48416
    If strPrivs = "" Then gstrPrivs = ";" & GetPrivFunc(glngSys, 1107) & ";"
    If Not blnNoFirst Then
        '55564 2012-10-13 lgf
        'If Not CheckValid Then Exit Function
        '权限设置(改为不可见)
        If InStr(gstrPrivs, ";发卡;") = 0 And InStr(gstrPrivs, ";绑定卡;") = 0 Then
            MsgBox "你没有被授予医疗卡发放和绑定的权限!", vbInformation, gstrSysName
            Exit Function
        End If
         '用户信息
        Call GetUserInfo
        Call zlinitSystemPara
        Call CreateSquareCardObject(frmMain, lngModule)
        Call CreatePublicExpenseObject(lngModule)
        blnNoFirst = True
    End If
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    If lngCardTypeID = 0 Then lngCardTypeID = Val(zlDatabase.GetPara("上次发卡类别", glngSys, 1107, 0))
    If lngCardTypeID = 0 Then
        Set rsTemp = zlGet医疗卡类别
         '问题48416
        rsTemp.Filter = "名称='就诊卡' and 是否固定=1 and 是否启用=1"
        If rsTemp.RecordCount = 0 Then
            rsTemp.Filter = 0
            MsgBox "未找到具体的医疗卡类别,不能继续操作!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        lngCardTypeID = Val(Nvl(rsTemp!id))
    End If
     '问题48416
    zlSendCard = frmPatiCureCardEdit.zlShowCard(frmMain, 1107, gstrPrivs, Cr_发卡, lngCardTypeID, "", lng病人ID)
End Function

Public Function zlModifyPass(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional lng病人ID As Long, Optional strCardNo As String, _
    Optional blnCheckOldPass As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:入口
    '入参:frmMain-调用的主窗体
    '       lngModule -模块号
    '       lngCardTypeId-卡类别ID
    '       lng病人ID-病人ID
    '       strCardNo-卡号
    '       blnCheckOldPass-是否验证原密码输入是否正确.
    '返回:修改成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-07-29 11:08:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmNew As frmModiPatiPass
    Set frmNew = New frmModiPatiPass
    zlModifyPass = frmNew.zlModifyPass(frmMain, lngModule, lngCardTypeID, lng病人ID, strCardNo, blnCheckOldPass)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
End Function

Public Function zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, strExpand As String, dblMoney As Double, Optional bln消费卡 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取帐户余额
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '       strCardNo-卡号
    '       strExpand-预留，为空,以后扩展
    '       bln消费卡-是否为消费卡
    '出参:dblMoney-返回帐户余额
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 16:29:48
    '说明:
    '       在所有需要扣款的地方，都要检查帐户余额是否充足,帐户不充足时不允许扣款.
    '       如果某些第三方接口不存在余额接口，可以固定返回一定的金额。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlGetAccountMoney = mobjPubThreeInterface.zlGetAccountMoney(frmMain, lngModule, lngCardTypeID, strCardNo, strExpand, dblMoney, bln消费卡)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardNODencode(ByVal strCardNo As String, ByVal lng卡类别ID As Long, bln消费卡 As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡号的加密数据
    '返回:返回被加密的数据(*****123)
    '编制:刘兴洪
    '日期:2011-08-25 09:07:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlGetCardNODencode = GetCardNODencode(strCardNo, lng卡类别ID, "", bln消费卡)
End Function

Private Function ThreePayRegisterMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
     ByVal lng挂号结帐ID As Long, ByRef cllPro As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If lngCardTypeID = 0 Then ThreePayRegisterMoney = False: Exit Function
     'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle:
    If mobjPubThreeInterface.zlPaymentMoney(frmMain, lngModule, lngCardTypeID, False, strCardNo, CStr(lng挂号结帐ID), "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
    Call zlAddUpdateSwapSQL(False, lng挂号结帐ID, lngCardTypeID, False, strCardNo, strSwapGlideNO, strSwapMemo, cllPro)
    Call zlAddThreeSwapSQLToCollection(False, lng挂号结帐ID, lngCardTypeID, False, strCardNo, strSwapExtendInfor, cllThreeSwap)
    ThreePayRegisterMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
       Resume
    End If
End Function
Public Function zlRegisterIncept(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal strNO As String, ByVal strInceptRoom As String, _
    ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, _
    Optional ByRef strExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:挂号接收接口
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strNO-预约单号
    '        strInceptRoom-接诊诊室
    '        lngCardTypeID-刷卡类别ID(可以付入零),表示只扣预存款
    '        strCardNo-卡号
    '       strExtendInfor-扩展信息(暂留,待以后扩展)
    '出参:
    '       strExtendInfor-交易扩展信息(暂留,待以后扩展)
    '返回:接收成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:13:48
    '说明:
    '    A.扣款流程:先检查有无门诊预存款,有,则扣门诊预存款;无则扣传入进来和银行存款)
    '    B.该接口不支持刷卡验证,直接从预存款中扣除或直接从银行帐户扣除
    '    C.本接口不支持用消费卡支付
    '调用者:
    '   医生工作站(刷卡接收)
    ' 问题:44543
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strProIncept As String, strSumUpdate As String
    Dim strSQL As String, rsTemp As ADODB.Recordset, rs时段 As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean
    Dim lng安排ID As Long, int预约有效时间 As Integer  '提前分钟数
    Dim dtDate As Date, dtNow As Date, lng结帐ID As Long, blnNotThree As Boolean
    Dim blnCardTypePay As Boolean '医疗卡支付
    Dim dblMoney As Double, cllPro As Collection, cllThreeSwap As Collection
    Dim lng病人ID As Long, rsTest As ADODB.Recordset
    Err = 0: On Error GoTo errHandle
    
    '挂号一般是使用的医生站的,肯定为HIS的,所以不使用连接来区分
    Set mfrmMain = frmMain
        
     strSQL = "" & _
    "   Select B.ID as 安排ID,A.ID, A.NO, A.病人id, A.姓名, A.登记时间, A.发生时间, A.预约时间, A.号别, A.号序, A.执行人," & _
    "                Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6','周五', '7', '周六', Null)  as 星期 " & _
    "   From 病人挂号记录 A,挂号安排 B" & _
    "   Where A.记录性质 = 2 and A.号别=B.号码 And A.记录状态 = 1 And A.NO = [1]  " & _
    "               And Trunc(Nvl(A.预约时间, A.发生时间)) = Trunc(Sysdate)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取预约挂号单", strNO)
    If rsTemp.EOF Then
        MsgBox "未找到预约单号为" & strNO & "的预约病人,可能该病人未预约今天的号!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng安排ID = Val(Nvl(rsTemp!安排ID))
    lng病人ID = Val(Nvl(rsTemp!病人ID))
    
    dtNow = zlDatabase.Currentdate
    'zlDatabase.ExecuteProcedure "Zl_挂号安排_Autoupdate(" & lng安排ID & ")", "挂号自动执行"
    '检查预约时间段是否正确
   ' gstrSQL = "" & _
    "Select A.星期,A.开始时间,A.结束时间,A.限制数量,A.是否预约 " & _
    " From 挂号安排时段 A,挂号安排 B" & _
    " Where A.安排ID=B.ID and A.安排id = [1]  And A.星期 = [2] "
    'Set rs时段 = zlDatabase.OpenSQLRecord(strSQL, "获取有效时段", lng安排ID, CStr(Nvl(rsTemp!星期)))
     int预约有效时间 = Val(zlDatabase.GetPara("预约有效时间", glngSys, 1111, "0"))
    dtDate = CDate(Format(rsTemp!预约时间, "yyyy-mm-dd HH:MM:SS")) - 1 / 24 / 60 * int预约有效时间
    If dtNow > dtDate Then
        MsgBox "注意:" & vbCrLf & "预约单号为:" & strNO & "的病人已经过预约时间,不能接诊!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    Set cllPro = New Collection: Set cllThreeSwap = New Collection
     strSQL = "" & _
    "   Select A.病人id, A.标识号, max(Nvl(B.预交余额, 0) - Nvl(B.费用余额, 0)) As 余额,  " & _
    "       nvl(Sum(A.实收金额),0) As 挂号金额 " & _
    "   From 门诊费用记录 A, 病人余额 B " & _
    "   Where A.病人id = B.病人id(+) And B.性质(+) = 1 and b.类型(+)=1 And A.NO = [1]  " & _
    "           And A.记录性质 = 4 And A.记录状态 = 0 " & _
    "   Group by A.病人ID,A.标识号"
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取门诊挂号费用", strNO)
     blnCardTypePay = False
     dblMoney = Round(Val(Nvl(rsTmp!挂号金额)), 2)
    
     If dblMoney > Round(Val(Nvl(rsTmp!余额)), 2) Or Val(Nvl(rsTmp!余额)) < 0 Then
        If lngCardTypeID <> 0 Then  '55449:如果传进来的是自制卡的,余额不足时,需按银行卡进行处理
            gstrSQL = "Select 是否自制,是否存在帐户 From 医疗卡类别 Where  ID=[1]"
            Set rsTest = zlDatabase.OpenSQLRecord(gstrSQL, "获取卡类别", lngCardTypeID)
            If Not rsTest.EOF Then
                If Val(Nvl(rsTest!是否自制)) = 1 Then lngCardTypeID = 0
                If Val(Nvl(rsTest!是否存在帐户)) = 0 Then lngCardTypeID = 0
            Else
                lngCardTypeID = 0
            End If
        End If
        
        If lngCardTypeID = 0 Then
            blnNotThree = False
            If zlSelectPayType(frmMain, lngCardTypeID, blnNotThree) = False Then
                If blnNotThree Then
                    MsgBox "病人预存款余额不足,不能进行挂号接收!", vbOKOnly + vbInformation, gstrSysName
                End If
                 '问题:52702
                Exit Function
            End If
        End If
        blnCardTypePay = True
     End If
     
    If lngCardTypeID <= 0 And Not blnCardTypePay Then
        '需要刷卡验证进行接收(lngCardTypeID-主界面传入,表示已经挂号)
        '47216
        If zlDatabase.PatiIdentify(frmMain, glngSys, lng病人ID, CCur(dblMoney), lngModule, 0, lngCardTypeID, False) = False Then Exit Function
    End If
    
    If blnCardTypePay Then
        '刷第三方卡
        If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
        If ThreeSwapCheck(frmMain, lngModule, lng病人ID, strNO, lngCardTypeID, dblMoney, strCardNo, "") = False Then Exit Function
    End If
    
    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    'zl_病人预约挂号_接收
    strSQL = "zl_病人预约挂号_接收("
    '  No_In   病人挂号记录.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  诊室_In 病人挂号记录.诊室%Type
    strSQL = strSQL & "'" & strInceptRoom & "',"
    '    结帐id_In     门诊费用记录.结帐id%Type := Null,
    strSQL = strSQL & lng结帐ID & ","
    '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & IIf(blnCardTypePay, lngCardTypeID, "NULL") & ","
    '    卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & IIf(blnCardTypePay, "'" & strCardNo & "'", "NULL") & ","
    '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & IIf(blnCardTypePay, "NULL", "NULL") & ","
    '    交易说明_In   病人预交记录.交易说明%Type := Null
    strSQL = strSQL & IIf(blnCardTypePay, "NULL", "NULL") & ")"

    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, "挂号接收保存"
    If blnCardTypePay Then
        If ThreePayRegisterMoney(frmMain, lngModule, lngCardTypeID, strCardNo, lng结帐ID, cllPro, cllThreeSwap, dblMoney) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        Call zlExecuteProcedureArrAy(cllPro, "三方交易更新", True, True)
    End If
    gcnOracle.CommitTrans: blnTrans = False
    If blnCardTypePay Then
        Err = 0: On Error GoTo ErrThreeHand
        zlExecuteProcedureArrAy cllThreeSwap, "保存第三方交易数据"
    End If
    zlRegisterIncept = True: blnTrans = False
    
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        If blnTrans Then gcnOracle.BeginTrans
        Resume
    End If
    Call SaveErrLog
    Exit Function
ErrThreeHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ThreeSwapCheck(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lng病人ID As Long, ByVal strNos As String, _
    ByVal lngCardTypeID As Long, ByVal dblMoney As Double, _
    ByRef strCardNo As String, ByRef strCardPassWord As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡验证
    '返回:返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim rsInfor As ADODB.Recordset
    Dim strSQL As String
    Dim objDatabase As New clsDataBase
    
    On Error GoTo errHandle
    
    Call objDatabase.InitCommon(mcnOracle)
    strSQL = "Select 姓名,性别,年龄 From 病人信息 where 病人ID=[1]"
    Set rsInfor = objDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病人ID)
    Set objDatabase = Nothing
    
    If rsInfor.EOF Then Exit Function
    '    zlBrushCard(frmMain As Object, _
    '        ByVal lngModule As Long, _
    '        ByVal lngCardTypeID As Long, _
    '        ByVal strPatiName As String, ByVal strSex As String, _
    '        ByVal strOld As String, ByVal dbl金额 As Double, _
    '        Optional ByRef strCardNo As String, _
    '        Optional ByRef strPassWord As String
    If mobjPubThreeInterface.zlBrushCard(frmMain, lngModule, Nothing, lngCardTypeID, False, _
     Nvl(rsInfor!姓名), Nvl(rsInfor!性别), Nvl(rsInfor!年龄), dblMoney, strCardNo, strCardPassWord) = False Then Exit Function
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If mobjPubThreeInterface.zlPaymentCheck(frmMain, lngModule, lngCardTypeID, False, _
            strCardNo, dblMoney, strNos, strXMLExpend) = False Then Exit Function
    ThreeSwapCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThreeInterFaceMgrTool(ByVal frmMain As Object, ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方接口管理工具
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '编制:刘兴洪
    '日期:2012-02-13 10:17:26
    '问题:46203
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlThreeInterFaceMgrTool = mobjPubThreeInterface.zlThreeInterFaceMgrTool(frmMain, lngModule)
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub SetEnabled(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置Enabled属性
    '入参:blnEnabled-true,表示自动获取;False表示不自动获取
    '编制:刘兴洪
    '日期:2012-08-22 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle
    If mobjPubThreeInterface Is Nothing Then Exit Sub
    Call mobjPubThreeInterface.SetEnabled(blnEnabled)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub zlInitEvents(ByVal hWnd As Long, ByVal objComEvents As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置公共事件
    '入参:objComEvents-公共事件
    '编制:刘兴洪
    '日期:2012-08-22 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call mobjPubThreeInterface.zlInitEvents(hWnd, objComEvents)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlGetYLCards() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别集
    '返回: 医疗卡类别的集合
    '编制:刘兴洪
    '日期:2012-08-24 14:37:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetYLCards = zlGet医疗卡类别(mcnOracle)
End Function
Public Function zlGetAvailabilityWriteCardType() As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:存在写卡类别
    '出参:返回写卡类别,多个用逗号分离
    '返回:存在写卡类别的ID,如:123,232,...
    '编制:刘兴洪
    '日期:2013-06-07 10:40:59
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strWriteCardIDs As String
    Dim intAutoRead As Integer, intAutoSplitTime As Integer, blnStartCardType As Boolean, str部件 As String
    On Error GoTo errHandle
    
    Set rsTemp = zlGetYLCards
    
    rsTemp.Filter = "是否写卡=1 And 是否启用=1"
    If rsTemp.EOF Then zlGetAvailabilityWriteCardType = "": Exit Function
    strWriteCardIDs = ""
     With rsTemp
        '自制卡(即消费卡)
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            intAutoRead = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & Nvl(!编码), "自动读取", "0"))
            intAutoSplitTime = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & Nvl(!编码), "自动读取间隔", "300"))
            If Val(Nvl(rsTemp!是否自制)) = 1 Then   '自制卡,都启用
                blnStartCardType = True
            Else
                blnStartCardType = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\医疗卡\" & Nvl(!编码), "启用", "0")) = 1
            End If
            If blnStartCardType Then
                strWriteCardIDs = strWriteCardIDs & "," & Val(Nvl(rsTemp!id))
            End If
            .MoveNext
        Loop
     End With
     Set rsTemp = Nothing
    If strWriteCardIDs <> "" Then strWriteCardIDs = Mid(strWriteCardIDs, 2)
    zlGetAvailabilityWriteCardType = strWriteCardIDs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, _
    ByRef objCard As Card) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别给卡对象(此对象
    '入参:lngCardTypeID-卡类别ID
    '       bln消费卡-是否消费卡
    '出参:objCard-返回卡对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-25 10:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, str部件 As String
    Dim int自动读取 As Integer, int自动读取间隔 As Integer, bln启用 As Boolean
    Dim objDatabase As New clsDataBase, str读卡性质 As String
    
    On Error GoTo errHandle
    Call objDatabase.InitCommon(mcnOracle)
    
    Set objCard = New Card
    '77872,李南春,2014/10/28:是否支持转帐及代扣
    '85565,李南春,2015/7/10:读卡性质
    '90875,李南春,2016/1/21:是否证件类型
    '103310,李南春,2016/12/7:卡号后增加回车符位
    If Not bln消费卡 Then
        strSQL = "" & _
         "   Select A.Id,A.编码,A.名称,A.短名,A.前缀文本,A.卡号长度,A.缺省标志,A.是否固定,A.是否严格控制," & _
         "           nvl(A.是否自制,0) as 是否自制,nvl(A.是否存在帐户,0) as 是否存在帐户," & _
         "           nvl(A.是否全退,0) as 是否全退,nvl(A.是否重复使用,0) as 是否重复使用 ," & _
         "           nvl(A.密码长度,10) as 密码长度,nvl(A.密码长度限制,0) as 密码长度限制,nvl(A.密码规则,0) as 密码规则," & _
         "           nvl(A.是否退现,0) as 是否退现,A.部件,A.备注,A.特定项目,A.结算方式,A.是否启用,A.卡号密文,Nvl(A.密码输入限制,0) as 密码输入限制,Nvl(A.是否缺省密码,0) as 是否缺省密码," & _
         "           nvl(A.是否模糊查找,0) as 是否模糊查找,nvl(A.是否制卡,0) as 是否制卡, decode(nvl(A.是否自制,0),1,1,nvl(A.是否发卡,0)) as 是否发卡,nvl(A.是否写卡,0) as 是否写卡,B.性质 as 结算性质," & _
         "           nvl(A.是否转帐及代扣,0) as 是否转帐及代扣, nvl(A.读卡性质,'1000') as 读卡性质,nvl(是否证件,0) as 是否证件, " & _
         "           nvl(A.是否持卡消费,0) as 是否持卡消费,nvl(A.发送调用接口,0) as 发送调用接口, " & _
         "           Nvl(a.是否退款验卡,0) As 是否退款验卡,A.设备是否启用回车,A.是否缺省退现 " & _
         "    From 医疗卡类别 A,结算方式 B " & _
         "    Where  A.ID=[1] And A.结算方式=B.名称(+)" & _
         "    "
         Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "读取指定医疗卡", lngCardTypeID)
         If rsTemp.EOF Then Set objDatabase = Nothing: Exit Function
        If Val(Nvl(rsTemp!是否启用)) = 1 Then
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & Nvl(rsTemp!编码), "自动读取", "0"))
            int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & Nvl(rsTemp!编码), "自动读取间隔", "300"))
            If Val(Nvl(rsTemp!是否自制)) = 1 Then   '自制卡,都启用
                bln启用 = True
            Else
                '问题号:54098
                If (Nvl(rsTemp!名称) Like "*身份证*" Or Nvl(rsTemp!名称) Like "*IC卡*") And Val(Nvl(rsTemp!是否固定)) = 1 And Nvl(rsTemp!部件) = "" Then
                    bln启用 = True
                Else
                    bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\医疗卡\" & Nvl(rsTemp!编码), "启用", "0")) = 1
                End If
            End If
        Else
            bln启用 = False
        End If
        
        str部件 = Trim(Nvl(rsTemp!部件))
        'ID,编码,名称,短名,前缀文本,卡号长度,缺省标志,是否固定,是否严格控制,是否刷卡,是否自制,是否存在帐户,是否全退,部件,备注,特定项目,结算方式,是否启用
        Set objCard = New Card
        With objCard
            .接口序号 = Nvl(rsTemp!id)
            .接口编码 = Nvl(rsTemp!编码)
            .短名 = Nvl(rsTemp!短名)
            .名称 = Nvl(rsTemp!名称)
            .前缀文本 = Nvl(rsTemp!前缀文本)
            .卡号长度 = Val(Nvl(rsTemp!卡号长度)) + Val(Nvl(rsTemp!设备是否启用回车))
            .缺省标志 = Val(Nvl(rsTemp!缺省标志)) = 1
            .系统 = Val(Nvl(rsTemp!是否固定)) = 1
            .是否严格控制 = Val(Nvl(rsTemp!是否严格控制)) = 1
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动读取间隔
            .自制卡 = Val(Nvl(rsTemp!是否自制)) = 1
            .是否存在帐户 = Val(Nvl(rsTemp!是否存在帐户)) = 1
            .是否全退 = Val(Nvl(rsTemp!是否全退)) = 1
            .卡号重复使用 = Val(Nvl(rsTemp!是否重复使用)) = 1
            .结算方式 = Nvl(rsTemp!结算方式)
            .接口程序名 = Nvl(rsTemp!部件)
            .特定项目 = Nvl(rsTemp!特定项目)
            .启用 = bln启用
            .备注 = Nvl(rsTemp!备注)
            .卡号密文规则 = Nvl(rsTemp!卡号密文)
            .是否退现 = Val(Nvl(rsTemp!是否退现)) = 1
            .密码长度 = Val(Nvl(rsTemp!密码长度))
            .密码长度限制 = Val(Nvl(rsTemp!密码长度限制))
            .密码规则 = Val(Nvl(rsTemp!密码规则))
            .密码输入限制 = Val(Nvl(rsTemp!密码输入限制))
            .是否缺省密码 = Val(Nvl(rsTemp!是否缺省密码)) = 1
            .是否制卡 = Val(Nvl(rsTemp!是否制卡)) = 1   '56615
            .是否发卡 = Val(Nvl(rsTemp!是否发卡)) = 1 Or .自制卡
            .是否写卡 = Val(Nvl(rsTemp!是否写卡)) = 1
            .结算性质 = Val(Nvl(rsTemp!结算性质))
            .是否转帐及代扣 = Val(Nvl(rsTemp!是否转帐及代扣)) = 1
            str读卡性质 = Nvl(rsTemp!读卡性质, "1000")
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            .是否持卡消费 = Val(Nvl(rsTemp!是否持卡消费)) = 1
            .是否退款验卡 = Val(Nvl(rsTemp!是否退款验卡)) = 1
            .是否证件 = Val(Nvl(rsTemp!是否证件)) = 1
            .设备是否启用回车 = Val(Nvl(rsTemp!设备是否启用回车)) = 1
            .是否缺省退现 = Val(Nvl(rsTemp!是否缺省退现)) = 1
        End With
        Set objDatabase = Nothing: zlGetCard = True
        Exit Function
    End If
    
    strSQL = "" & _
    "   Select A.编号 as ID,A.编号 ,A.名称,substr(A.名称,1,1) as 短名,A.结算方式,nvl(A.自制卡,0)  as 自制卡,A.前缀文本,A.卡号长度, " & _
    "           nvl(A.是否退现,0) as 是否退现,nvl(A.是否全退,0) as 是否全退," & _
    "           nvl(A.密码长度,10) as 密码长度,nvl(A.密码长度限制,0) as 密码长度限制,nvl(A.密码规则,0) as 密码规则," & _
    "           A.部件,A.系统,A.是否密文,0 as 密码输入限制,0 as 是否缺省密码," & _
    "           0 as 是否模糊查找,0 as 是否制卡, 1 as 是否发卡, 0 as 是否写卡,nvl(A.启用,0) as 启用,b.性质 as 结算性质," & _
    "           nvl(A.读卡性质,'1000') as 读卡性质 " & _
    "   From 消费卡类别目录 A,结算方式 B " & _
    "   Where A.编号=[1] and A.结算方式=B.名称(+)"
    Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "读取指定消费卡", lngCardTypeID)
    If rsTemp.EOF Then Set objDatabase = Nothing: Exit Function
    
    With rsTemp
        '自制卡(即消费卡)
        If .RecordCount <> 0 Then .MoveFirst
        bln启用 = Val(Nvl(!启用)) = 1
        If bln启用 Then
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & Nvl(!编号), "自动读取", "0"))
            int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & Nvl(!编号), "自动读取间隔", "300"))
            bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "启用", "0")) = 1
        End If
        '编号,名称,结算方式,nvl(自制卡,0)  as 自制卡,前缀文本,卡号长度,部件,系统,是否密文
        str部件 = Trim(Nvl(rsTemp!部件))
        Set objCard = New Card
        With objCard
            .接口序号 = Nvl(rsTemp!编号)
            .接口编码 = Nvl(rsTemp!编号)
            .短名 = Left(Nvl(rsTemp!名称), 1)   '默认取第一个
            .名称 = Nvl(rsTemp!名称)
            .前缀文本 = Nvl(rsTemp!前缀文本)
            .卡号长度 = Val(Nvl(rsTemp!卡号长度))
            .系统 = Val(Nvl(rsTemp!系统)) = 1
            .是否严格控制 = False
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动读取间隔
            .自制卡 = Val(Nvl(rsTemp!自制卡)) = 1
            .是否存在帐户 = True 'Not (Val(Nvl(rsTemp!自制卡)) = 1)
            .是否全退 = Val(Nvl(rsTemp!是否全退)) = 1
            .结算方式 = Nvl(rsTemp!结算方式)
            .接口程序名 = Nvl(rsTemp!部件)
            .特定项目 = ""
            .启用 = bln启用
            .卡号重复使用 = True
            .备注 = ""
            .卡号密文规则 = Nvl(rsTemp!是否密文)
            .消费卡 = True
            .是否退现 = Val(Nvl(rsTemp!是否退现)) = 1
            .密码长度 = Val(Nvl(rsTemp!密码长度))
            .密码长度限制 = Val(Nvl(rsTemp!密码长度限制))
            .密码规则 = Val(Nvl(rsTemp!密码规则))
            .密码输入限制 = Val(Nvl(rsTemp!密码输入限制))
            .是否缺省密码 = Val(Nvl(rsTemp!是否缺省密码)) = 1
            .是否制卡 = Val(Nvl(rsTemp!是否制卡)) = 1   '56615
            .是否发卡 = Val(Nvl(rsTemp!是否发卡)) = 1 Or .自制卡
            .是否写卡 = Val(Nvl(rsTemp!是否写卡)) = 1
            .结算性质 = Val(Nvl(rsTemp!结算性质))
            str读卡性质 = Nvl(rsTemp!读卡性质, "1000")
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
        End With
    End With
    zlGetCard = True
    Set objDatabase = Nothing
    Exit Function
errHandle:
    If Not objDatabase Is Nothing Then
        If objDatabase.ErrCenter = 1 Then Resume
    Else
        If ErrCenter() = 1 Then Resume
    End If
    Set objDatabase = Nothing
End Function

Public Function zlSetBrushCardObject(ByVal lngCardTypeID As Long, _
    ByRef objBrushCard As Object, ByRef strExpend As String, Optional ByVal bln消费卡 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡对象
    '入参:lngCardTypeID-医疗卡类别ID
    '    objBrushCard -指定的刷卡对象(文本框)
    '    strExpend-扩展信息(暂无,待以后扩展用)
    '出参:strExpend-扩展信息(暂无,待以后扩展用)
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-21 14:04:13
    '说明:主要针对刷卡方式的,比如:有些银行卡加密的,需要刷卡后进行解密
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlSetBrushCardObject = mobjPubThreeInterface.zlSetBrushCardObject(lngCardTypeID, objBrushCard, strExpend, bln消费卡)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlTransferAccountsCheck(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, ByVal dblMoney As Double, ByVal strBalanceID As String, Optional ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:转帐检查
    '入参:
    '   frmMain-调用的主窗体
    '   lngModule-HIS调用模块号
    '   lngCardTypeID-卡类别ID
    '   strCardNo-卡号
    '   dblMoney-转帐金额(代扣时为负数)
    '   strBalanceID-本次支付结算ID，4-门诊退费业务为原结帐ID
    '   strXMLExpend-XML串:
    '       <IN>
    '           <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务
    '       </IN>
    '出参:strXMLExpend-XML串:
    '        <OUT>
    '           <ERRMSG>错误信息</ERRMSG >
    '        </OUT>
    '返回:检查的数据合法,返回True:否则返回False
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    On Error GoTo errHandle
    zlTransferAccountsCheck = mobjPubThreeInterface.zlTransferAccountsCheck(frmMain, lngModule, lngCardTypeID, strCardNo, dblMoney, strBalanceID, strXMLExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlTransferAccountsMoney(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, ByVal strBalanceID As String, ByVal dblMoney As Double, Optional ByRef strSwapGlideNO As String, _
    Optional ByRef strSwapMemo As String, Optional ByRef strSwapExtendInfor As String, Optional ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方帐户转帐
    '入参:
    '   frmMain-调用的主窗体
    '   lngModule-HIS调用模块号
    '   lngCardTypeID-卡类别ID
    '   strCardNo-卡号
    '   strBalanceID-结算ID，4-门诊退费业务为原结帐ID
    '   dblMoney-转帐金额
    '    strSwapExtendInfor-退费业务时，传入本次退费的冲销ID:
    '                        格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                        收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '   strXMLExpend-XML串:
    '       <IN>
    '           <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务
    '       </IN>
    '出参:
    '   strSwapGlideNO-交易流水号
    '   strSwapMemo -交易说明
    '   strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '   strXMLExpend-XML串:
    '        <OUT>
    '           <ERRMSG>错误信息</ERRMSG >
    '        </OUT>
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain
    zlTransferAccountsMoney = mobjPubThreeInterface.zlTransferAccountsMoney(frmMain, lngModule, _
     lngCardTypeID, strCardNo, strBalanceID, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function SetParentHwnd(ByVal lngfrmMainHwnd As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置父窗口句配
    '入参:lngfrmMainHwnd-父窗口句柄
    '返回:设置成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-19 14:52:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If lngfrmMainHwnd = 0 Then Exit Function
    mlngfrmMainHwnd = lngfrmMainHwnd
End Function

Public Function ZlGetParaConfig(ByVal frmMain As Object, _
    ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, ByVal intPara As Integer, _
    Optional strErrMsg As String, Optional strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取接口参数
    '入参: frmMain-调用的主窗体
    '       intPara: 包含如下值
    '                1-刷卡和支付在同一页面:true-新模式；False-旧模式
    '                2-作废时是否单独调用退费接口
    '       strExpend-扩展参数，暂留，现传为空
    '出参:strErrMsg-返回的错误信息
    '       strExpend-扩展参数，暂留，现传为空
    '返回:函数返回True:调用成功,False:调用失败
   '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    ZlGetParaConfig = mobjPubThreeInterface.ZlGetParaConfig(frmMain, lng卡类别ID, bln消费卡, intPara, strErrMsg, strExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlGetCardNo(ByRef strCardNo As String, Optional strExpend As String) As Boolean
    '获取刷卡卡号
    '出参:strCardNo-刷卡卡号
    '     strExpend-扩展参数，暂留，现传为空
    On Error GoTo errHandle
    ZlGetCardNo = mobjPubThreeInterface.ZlGetCardNo(strCardNo, strExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

