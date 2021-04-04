VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendCardAndDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'预交款及发卡
'公共接口:
'    1.zlInitCommon-初始化接口
'    2.zlSetFontSize-设置界面字体大小
'    3.zlGetForm -获取主窗体对象
'    4.zlRecalcCardFee-新计算卡费用：在性别/年龄/医疗付款方式/费别/自份证号改变后调用
'    5.zlSetCardNo-在病人项进行刷卡后，如果是新卡,则需要将该卡号作为发卡卡号，需要设置该卡号
'    6.zlGetSendCard-获取当前的发卡卡对象
'    7.zlSetUnitInfo-设置缴款单位信息(工作单位，账号等数据完成后，需要设置)
'    8.zlSetInsureInfo:设置医保信息(医保身份验证后调用)
'    7.zlClearControlInfo -清除界面控件信息
'数据保存接口
'    1.zlSaveDataBeforCheckIsValid-检查数据的合法性:在保存前需要调用
'    2.zlSaveData-执行数据保存操作
'    3.zlSaveDataAfter-数据保存后执行(主要是重新刷新票据,清空控件信息等)
'公共事件:
'    1.RequestRefreshPatiInf-请求重新根据XML格式的输出内容，刷新病人信息
'    2.InputOver-输入完成事件(表示最后一项输入完成，以便光标跳转到下一项输入内容
'    3.Activate-子窗体激活
'    4.ExcuteQRCodePayment-要求执行扫码付操作(需要在外部直接调用确定按钮0
'    5.zlSetFocus-光标定位
'    6.ControlGotFocus-焦点进入控件时触发
'公共属性
'    1.RealName-设置当前病人是否进行了实名认证(实名认证后，需要赋值)
'    2.Width-窗体宽度
'    3.Height-窗体高度
'编制:刘兴洪
'日期:2019-11-27 16:52:50
'*********************************************************************************************************************************************
Private mlngfrmMainHwnd As Long '父窗口对象
Private WithEvents mfrmSendCardAndDeposit As frmSendCardAndDeposit
Attribute mfrmSendCardAndDeposit.VB_VarHelpID = -1
Private mlngModule As Long
'------------------------------------------------------------------------------------------------------
'公共事件
Public Event Activate() '子窗体激活
Public Event RequestRefreshPatiInf(ByVal strCardNo As String, ByVal strPatiInfoXML As String)
Public Event InputOver()    '输入完成
Public Event ExcuteQRCodePayment() '执行扫码付
Public Event ExcuteReadQRCode()
Public Event ControlGotFocus(objControl As Object)

Public Function zlInitCommon(ByVal frmMain As Object, ByVal lngSys As Long, ByVal cnOracle As ADODB.Connection, _
    ByVal lngModule As Long, ByVal bln门诊预交 As Boolean, ByVal bln住院预交 As Boolean, _
    ByVal lngCardTypeID As Long, blnAllowSendCard As Boolean, ByVal blnAllowBoundCard As Boolean, ByVal blnAllowInsureAccDeposit As Boolean, _
    Optional btQRCodePay As Object, Optional objShowTotalMoneyControl As Object, Optional blnView As Boolean = False, _
    Optional ByVal blnCancel As Boolean, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '     frmMain-调用的主窗体
    '     lngModule-模块号
    '     btQRCodePay-扫码付按钮
    '     objShowTotalMoneyControl-显示的总额控件:lable或Text
    '     bln门诊预交-是否缴门诊预交
    '     bln住院预交-是否缴住院预交
    '     lngSendCardTypeID-当前发卡类别ID:传入0时，则参数：blnAllowSendCard及blnAllowBoundCard-无效
    '     blnAllowSendCard-允许发卡
    '     blnAllowBoundCard-允许绑定卡
    '     blnView-是否查看
    '     strPrivs-当前操作模块权限
    '     blnAllowInsureAccDeposit-是否允许医保账户缴预交
    '     blnCancel-当前是否作废操作
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser: mlngModule = lngModule
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    Call gclsInsure.InitOracle(gcnOracle)
    zlInitCommon = mfrmSendCardAndDeposit.zlInit(frmMain, lngModule, bln门诊预交, bln住院预交, lngCardTypeID, blnAllowSendCard, blnAllowBoundCard, blnAllowInsureAccDeposit, _
        btQRCodePay, objShowTotalMoneyControl, blnView, , blnCancel)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlSetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘兴洪
    '日期:2014-04-09 11:46:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmSendCardAndDeposit Is Nothing Then Exit Sub
     mfrmSendCardAndDeposit.zlSetFontSize (bytSize)
End Sub

Public Sub zlSetFocus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:光标定位
    '入参
    '编制:刘兴洪
    '日期:2020-01-13 17:53:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmSendCardAndDeposit Is Nothing Then Exit Sub
    mfrmSendCardAndDeposit.zlSetFocus
End Sub


Public Function zlRecalcCardFee(ByVal objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算卡费信息
    '入参:objPati-病人对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-23 17:44:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlRecalcCardFee = mfrmSendCardAndDeposit.zlRecalcCardFee(objPati)
End Function

Public Function zlSetCardNo(ByVal strCardNo As String, objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置卡号给卡号文本框
    '入参:objPati-病人信息集
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 18:53:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSetCardNo = mfrmSendCardAndDeposit.zlSetCardNo(strCardNo, objPati)
End Function
Public Sub zlSetUnitInfo(ByVal str单位帐号 As String, ByVal str缴款单位 As String, ByVal str单位开户行 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置单位账号
    '编制:刘兴洪
    '日期:2019-11-26 13:37:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmSendCardAndDeposit.zlSetUnitInfo(str单位帐号, str缴款单位, str单位开户行)
End Sub
Public Sub zlSetInsueInfo(ByVal int险类 As Integer, ByVal cur账户余额 As Currency, ByVal str医保号 As String, ByVal str密码 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置医保信息(医保身份验证后调用)
    '编制:刘兴洪
    '日期:2019-11-26 13:37:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmSendCardAndDeposit.zlSetInsueInfo(int险类, cur账户余额, str医保号, str密码)
End Sub
Public Sub zlClearControlInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除界面控件信息
    '编制:刘兴洪
    '日期:2019-11-26 13:37:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmSendCardAndDeposit.zlClearControlInfo
End Sub


Public Function zlGetSendCard(ByRef objSendCard_Out As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前的发卡对象
    '入参:
    '出参:objSendCard_Out-返回当前发卡的对象
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 15:13:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetSendCard = mfrmSendCardAndDeposit.zlGetSendCard(objSendCard_Out)
End Function

Public Function zlSaveDataBeforCheckIsValid(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo, _
    Optional ByVal bln自动识别的身份证 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '入参:objPati-病人信息集
    '     blnNewPati-是否新病人
    '     bln自动识别的身份证-是否自动识别的身份证号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 13:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSaveDataBeforCheckIsValid = mfrmSendCardAndDeposit.zlSaveDataBeforCheckIsValid(blnNewPati, objPati, bln自动识别的身份证)
End Function
 

Public Function zlSaveData(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '入参:objPati-病人信息集
    '     blnNewPati-是否新病人
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 13:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSaveData = mfrmSendCardAndDeposit.zlSaveData(blnNewPati, objPati)
End Function

Public Function zlSaveDataAfter() As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存后执行(主要是重新刷新票据,清空控件信息等)
    '入参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 15:08:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlSaveDataAfter = mfrmSendCardAndDeposit.zlSaveDataAfter
End Function

Private Sub Class_Initialize()
    Set mfrmSendCardAndDeposit = New frmSendCardAndDeposit
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmSendCardAndDeposit Is Nothing Then Unload mfrmSendCardAndDeposit
    Set mfrmSendCardAndDeposit = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
Private Sub mfrmSendCardAndDeposit_Activate()
    RaiseEvent Activate
End Sub
Public Function zlGetForm() As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取子业务窗口,由主程序初始化时只调用一次获取
    '返回:获取发卡及预交窗体对象,用于主界面绑定
    '编制:刘兴洪
    '日期:2014-04-09 12:00:14
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetForm = mfrmSendCardAndDeposit
End Function
Public Property Get RealName() As Boolean
       RealName = mfrmSendCardAndDeposit.RealName
End Property
Public Property Let RealName(ByVal vNewValue As Boolean)
     mfrmSendCardAndDeposit.RealName = vNewValue
End Property

Public Property Get Width() As Long
       Width = mfrmSendCardAndDeposit.GetWidth
End Property

Public Property Get Height() As Long
       Height = mfrmSendCardAndDeposit.GetHeight
End Property

Private Sub mfrmSendCardAndDeposit_ControlGotFocus(objControl As Object)
    RaiseEvent ControlGotFocus(objControl)
End Sub

Private Sub mfrmSendCardAndDeposit_ExcuteQRCodePayment()
    RaiseEvent ExcuteQRCodePayment
End Sub

Private Sub mfrmSendCardAndDeposit_ExcuteReadQRCode()
   RaiseEvent ExcuteReadQRCode
End Sub

Private Sub mfrmSendCardAndDeposit_InputOver()
    RaiseEvent InputOver
End Sub

Private Sub mfrmSendCardAndDeposit_RequestRefreshPatiInf(ByVal strCardNo As String, ByVal strPatiInfoXML As String)
    RaiseEvent RequestRefreshPatiInf(strCardNo, strPatiInfoXML)
End Sub


