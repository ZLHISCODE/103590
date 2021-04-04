VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicOneCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'功能:一卡通涉及的相关数据对象(包含获取一卡通的方法，数据集等)
'涉及的方法或函数:
'  1.公共方法
'  1.1  zlInitComponents-初始化对象(第一步必须先初始化，否则可能使用中会发生异常)
'  1.2  ErrCenter-错误处理中心
'  1.3  SaveErrLog-保存错误日志
'  1.3  WritDebugLog-写调试日志
'  1.4  SetPara-参数设置
'  1.5  GetPara-获取参数
'  1.7  GetPrivFunc-获取权限串
'  2.病人相关数据
'  2.1 zlGetPatiIDFromBedNumber:根据床位号，获取病人ID
'  2.2 zlGetPatiInforFromPatiID:根据病人ID获取病人信息对象
'  2.3 zlGetPatiInforFromXML:根据XML串，获取病人信息对象
'  3. 卡对象相关
'  3.1 zlGetCards-获取有效的卡类别
'涉及的公共属性:
'  5.1 objThirdSwap-获取三方接口对象
'  5.2 objOneCardObject-一卡通对象
'  5.3 objDataBase-涉及数据库操作对象
'编制:刘兴洪
'日期:2018-12-03 11:06:14
'*********************************************************************************************************************************************
Private mcnOracle As ADODB.Connection
Private mstrDBUser As String, mlngModul As Long
Private mfrmMain As Object
Private mobjDataBase As clsDataBase
Private mobjOneCardObject As clsOneCardDataObject
Private mobjThirdSwap As clsThirdSwapInterface  '三方交易接口
Private mblnIsObjRegisterAlone As Boolean

Public Property Get objThirdSwap() As clsThirdSwapInterface
    If mobjThirdSwap Is Nothing Then
        Set mobjThirdSwap = New clsThirdSwapInterface
        Call mobjThirdSwap.InitCommon(mcnOracle)
    End If
    Set objThirdSwap = mobjThirdSwap
End Property
Public Property Get objOneCardObject() As clsOneCardDataObject
    Call zlGetOneCardDataObject(mcnOracle, mobjOneCardObject)
    Set objOneCardObject = mobjOneCardObject
End Property
Public Function zlIsExistOraConnect() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否存在Oracle连接
    '编制:刘兴洪
    '日期:2018-12-06 18:17:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    zlIsExistOraConnect = mobjDataBase.zlIsExistOraConnect
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlInitComponents(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    ByVal cnOracle As ADODB.Connection, Optional blnDeviceSet As Boolean = False, Optional strExpand As String, _
    Optional ByVal blnIsObjRegisterAlone As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参: frmMain-调用的主窗体
    '        lngModule-HIS调用模块号
    '       lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       blnDeviceSet-设备设置调用初始化
    '       strExpand-扩展信息(可选传入:卡类别ID-不传时,表示全部初始化,传入时,只初始化指定的接口)
    '       blnIsObjRegisterAlone-是否使用独立的注册部件(True:使用:zlRegisterAlone.DLL,否则使用zlRegister.dll)
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
    mstrDBUser = strDBUser: mlngModul = lngModule: glngModul = lngModule
     
    Set mcnOracle = cnOracle
    If gcnOracle Is Nothing Then Set gcnOracle = cnOracle
    
    If mblnIsObjRegisterAlone <> blnIsObjRegisterAlone Then
        Set gobjOneDataBase = Nothing
    End If
    
    mblnIsObjRegisterAlone = blnIsObjRegisterAlone
    Call zlInitPublicVar '初始化全局变量
    If zlGetOneDataBase(cnOracle, mobjDataBase, blnIsObjRegisterAlone) = False Then Exit Function
    If zlGetOneCardDataObject(cnOracle, mobjOneCardObject) = False Then Exit Function
    
    If mobjThirdSwap Is Nothing Then Set mobjThirdSwap = New clsThirdSwapInterface
    Call mobjThirdSwap.InitCommon(cnOracle, mblnIsObjRegisterAlone)   '初始化三方接口变量
    gstrUnitName = mobjDataBase.GetUnitName
    zlInitComponents = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function zlIsMobileNo(ByVal strInput As String, Optional ByRef strRutType As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '功能:判断传入的是否为手机号
    '入参:strInput-手机号
    '出参:strRutType-查询结果:0-成功;1-不是有效号段;2-号码长度不对
    '返回:True-传入号码为手机号;False-传入号码不为手机号
    '编制:刘尔旋
    '日期:2017-1-25
    '---------------------------------------------------------------------------------------------
    zlIsMobileNo = mobjOneCardObject.zlIsMobileNo(strInput, strRutType)
End Function

Public Function zlGetPatiIDFromBedNumber(ByVal lng病区ID As Long, _
    ByVal str床号 As String, Optional ByRef lng主页ID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号获取病人ID
    '出参:lng主页ID-返回当前床号的主页ID
    '返回:成功返回病人ID,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlGetPatiIDFromBedNumber = mobjOneCardObject.zlGetPatiIDFromBedNumber(lng病区ID, str床号, lng主页ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlGetPatiID(ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的医疗类别和卡号,获取对应的病人ID
    '入参:strCardType-卡类别,如果为数字,这为卡类别ID,如果为字符,则为类别名称
    '       strCardNo-卡号
    '       blnNotShowErrMsg-不显示错误的提示信息
    '       frmMain-调用的主窗体
    '       objCtl-调用的控件
    '       blnShowMergePati-当出现多个满足条件的病人时,是否显示合并功能按钮
    '       blnOnlyContractPati-签约病人
    '       blnUserCancel-选择器中，用户选择了取消
    '       lngShowCardNoTypeID-过滤出多条病信息时，弹出选择器中显示的卡号的卡类别ID,0-表示不显示卡号；>0表示显示指定卡号类别的ID
    '       blnNotCheckValidDate-是否对卡终止使用时间进行检查,true-不检查终止使用时间,false-检查
    '出参:strErrMsg-返回的错误信息
    '       lng病人ID-返回的病人ID
    '       strCardPass-返回卡号的密码
    '       lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    zlGetPatiID = mobjOneCardObject.zlGetPatiID(strCardType, strCardNo, blnNotShowErrMsg, lng病人ID, _
      strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, _
      blnCertificate, blnUserCancel, lngShowCardNoTypeID, blnNotCheckValidDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    Set zlGetCards = mobjThirdSwap.zlGetCards(bytType)
End Function


Public Function zlGetPatiInforFromPatiID(ByVal lng病人ID As Long, ByRef objPati As clsPatiInfor, _
    ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String) As Boolean
    
    zlGetPatiInforFromPatiID = mobjOneCardObject.zlGetPatiInforFromPatiID(lng病人ID, objPati, strErrMsg, strOtherName, strOtherValue)
End Function

Public Function zlGetPatiInforFromXML(ByVal strPatiXml As String, ByRef strErrMsg_Out As String, ByRef objPatiInfor_Out As clsPatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从XML中获取病人信息
    '入参:strPatiXml-病人信息XML
    '
    '出参:objPatiInfor_Out-返回病人信息对象集
    '      strErrMsg_Out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 14:29:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiInforFromXML = mobjOneCardObject.zlGetPatiInforFromXML(strPatiXml, strErrMsg_Out, objPatiInfor_Out)
End Function
    

Public Function ErrCenter() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:错误处理中心
    '编制:刘兴洪
    '日期:2018-12-05 11:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
   ErrCenter = mobjDataBase.ErrCenter
End Function

Public Sub SaveErrLog()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写日志
    '编制:刘兴洪
    '日期:2018-12-05 11:33:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Sub
    End If
    Call mobjDataBase.SaveErrLog
End Sub
Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True, Optional ByVal lngDeptID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngSys=使用该参数的系统编号，如100
    '      lngModual=使用该参数的模块号，如1230
    '      blnSetup=调用模块是否有参数设置权限
    '      lngDeptID=部门级参数
    '返回：设置是否成功
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    SetPara = mobjDataBase.SetPara(varPara, strValue, glngSys, lngModual, blnSetup, lngDeptID)
End Function
Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer, Optional ByVal lngDeptID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：读取指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      lngSys=使用该参数的系统编号，如100
    '      lngModual=使用该参数的模块号，如1230
    '      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
    '      blnNotCache=是否不从缓存中读取
    '      arrControl=控件数组，如Array(Me.Text1, Me.CheckBox1)，用于函数内部自动处理对应控件的显示颜色，是否禁止设置。
    '      blnSetup=调用模块是否有参数设置权限
    '      intType=返回参数，返回参数类型
    '      lngDeptID=部门级参数，根据部门读取
    '返回：参数值，字符串形式
    '编制:刘兴洪
    '日期:2018-12-05 14:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    GetPara = mobjDataBase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType, lngDeptID)
End Function

Public Function GetPrivFunc(lngSys As Long, lngProgId As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：返回当前用户具有的指定程序的功能串
    '参数：lngSys     如果是固定模块，则为0
    '      lngProgId  程序序号
    '返回：分号间隔的功能串,为空表示没有权限
    '编制:刘兴洪
    '日期:2018-12-05 15:01:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    GetPrivFunc = mobjDataBase.GetPrivFunc(lngSys, lngProgId)
End Function

Public Sub WritDebugLog(ByVal strDev As String, strInput As String, strOutPut As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记录日志
    '编制:刘兴洪
    '日期:2018-12-05 11:35:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Sub
    End If
    Call mobjDataBase.WritLog(strDev, strInput, strOutPut)
End Sub


Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Call zlInitCommLib  '初始化CommLib对象
End Sub

Private Sub Class_Terminate()

    Err = 0: On Error Resume Next
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mobjOneCardObject Is Nothing Then Set mobjOneCardObject = Nothing
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If Not mobjThirdSwap Is Nothing Then Set mobjThirdSwap = Nothing
    Call zlReleaseResources '释放资源
End Sub

 
