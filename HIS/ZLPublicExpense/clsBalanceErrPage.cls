VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceErrPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'预交款及发卡
'公共接口:
'    1.zlInitCommon-初始化接口
'    2.zlRefreshData-重新刷新数据
'    3.zlGetForm -获取主窗体对象
'编制:刘兴洪
'日期:2019-11-27 16:52:50
'*********************************************************************************************************************************************
Private mlngfrmMainHwnd As Long '父窗口对象
Private mfrmSendCardAndDepositErrPage As frmSendCardAndDepositErrPage
Attribute mfrmSendCardAndDepositErrPage.VB_VarHelpID = -1
Private mlngModule As Long

Public Function zlInitCommon(ByVal frmMain As Object, ByVal lngSys As Long, ByVal cnOracle As ADODB.Connection, _
    ByVal lngModule As Long, int应用场合 As Byte, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '     frmMain-调用的主窗体
    '     lngModule-模块号
    '     int应用场合 -1-医疗卡发卡;2-病人信息登记;3-病人入院 登记;4-预约挂号接收
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-01-04 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser: mlngModule = lngModule
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    zlInitCommon = mfrmSendCardAndDepositErrPage.zlInit(frmMain, int应用场合, lngModule)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetForm() As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取子业务窗口,由主程序初始化时只调用一次获取
    '返回:获取发卡及预交窗体对象,用于主界面绑定
    '编制:刘兴洪
    '日期:2020-01-04 12:00:14
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetForm = mfrmSendCardAndDepositErrPage
End Function
Private Sub Class_Initialize()
    Set mfrmSendCardAndDepositErrPage = New frmSendCardAndDepositErrPage
    glngInstanceCount = glngInstanceCount + 1
End Sub
Public Sub zlRefreshData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新数据
    '入参:objPati-病人信息集
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-01-04 16:03:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmSendCardAndDepositErrPage.zlRefreshData
     
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmSendCardAndDepositErrPage Is Nothing Then Unload mfrmSendCardAndDepositErrPage
    Set mfrmSendCardAndDepositErrPage = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
 
