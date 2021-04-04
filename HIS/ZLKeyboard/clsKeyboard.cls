VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmCom As frmCom
Public Function OpenPassKeyoardInput(ByVal frmMain As Object, _
    ByVal objPassCtl As Object, Optional blnAffirmPass As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码建盘输入
    '入参:frmMain-调用的主窗体
    '       objPassCtl-输入的密码控件
    '       blnAffirmPass-False:请输入密码;true:请输入确认密码
    '出参:
    '返回:打开成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:30:54
    '--------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmCom Is Nothing Then Exit Function
    If Not gblnStartKeyboard Then OpenPassKeyoardInput = True: Exit Function
    On Error GoTo errHandle
    OpenPassKeyoardInput = mfrmCom.OpenPassKeyoardInput(frmMain, objPassCtl, blnAffirmPass)
    Exit Function
errHandle:
End Function

Public Function ColsePassKeyoardInput(ByVal frmMain As Object, ByVal objPassCtl As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭密码建盘输入
    '入参:frmMain-调用的主窗体
    '       objPassCtl-输入的密码控件
    '出参:
    '返回:关闭成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmCom Is Nothing Then Exit Function
    If Not gblnStartKeyboard Then ColsePassKeyoardInput = True: Exit Function
    On Error GoTo errHandle
    ColsePassKeyoardInput = mfrmCom.ColsePassKeyoardInput(frmMain, objPassCtl)
    Exit Function
errHandle:
End Function
Public Function zlCardDevSet(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlCardDevSet(设备参数设置接口):主要是配置相关的设备的相关参数
    '入参:frmMain Form    HIS传入 调用的主窗体
    '出参:
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    frmDeviceSet.Show 1, frmMain
End Function

Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    gblnStartKeyboard = Val(GetSetting("ZLSOFT", "公共模块\zlKeyboard", "启用", "0")) = 1
    Set mfrmCom = New frmCom
    Load mfrmCom
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmCom Is Nothing Then Unload mfrmCom
    Set mfrmCom = Nothing
End Sub
