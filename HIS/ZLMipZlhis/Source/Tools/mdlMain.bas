Attribute VB_Name = "mdlMain"
Option Explicit

'Public gcnOracle As ADODB.Connection
Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrSysName As String               '当前用户简码
Public gobjRegister As Object               '注册授权部件zlRegister
Public gobjComLib As Object

Public Sub Main()
        
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    '创建注册部件(用于登录时获取连接对象)
    On Error Resume Next
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    If gobjRegister Is Nothing Then
        Err.Clear
        MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
        Exit Sub
    End If
    If frmLogin.ShowLogin() = False Then Exit Sub
    Call gobjComLib.InitCommon(gcnOracle)
    If gcnOracle.State <> adStateOpen Then
        Exit Sub
    End If
    frmMain.Show
    
End Sub

