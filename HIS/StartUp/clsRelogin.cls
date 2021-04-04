VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mobjLogin       As Object   '登录对象
'属性，获取输入的用户名
Public Property Get InputUser() As String
    InputUser = mobjLogin.InputUser
End Property
'属性，获取输入的密码
Public Property Get InputPwd() As String
    InputPwd = mobjLogin.InputPwd
End Property
'属性，获取输入的服务器
Public Property Get ServerName() As String
    ServerName = mobjLogin.ServerName
End Property
'属性，获取是否密码需要转换
Public Property Get IsTransPwd() As Boolean
    IsTransPwd = mobjLogin.IsTransPwd
End Property
'属性，获取输入的用户对应的数据库用户
Public Property Get DBUser() As String
    DBUser = mobjLogin.DBUser
End Property
'属性，获取是否系统所有者
Public Property Get IsSysOwner() As Boolean
    IsSysOwner = mobjLogin.IsSysOwner
End Property
'属性，获取多帐套选择的Systems
Public Property Get Systems() As String
    Systems = mobjLogin.Systems
End Property
'属性，获取菜单组
Public Property Get MenuGroup() As String
    MenuGroup = mobjLogin.MenuGroup
End Property
'属性，获取菜单组
Public Property Get DeptName() As String
    DeptName = mobjLogin.DeptName
End Property
'属性，获取EMR对象
Public Property Get EMR() As Object
    Set EMR = mobjLogin.EMR
End Property
'属性，设置EMR对象
Public Property Set EMR(ByRef objEMR As Object)
    Set mobjLogin.EMR = objEMR
End Property
'属性，EMR的验证用户
Public Property Get EMRUser() As String
    EMRUser = mobjLogin.EMRUser
End Property

'属性，EMR的验证密码，若是数据库密码增加[DBPASSWORD]前缀
Public Property Get EMRPwd() As String
    EMRPwd = mobjLogin.EMRPwd
End Property

'属性，EMRUser、EMRPwd是否属于代理用户。如当LIS单独库时，使用ZLHIS账户来初始化EMR,此时该属性为TRUE
Public Property Get IsEMRProxy() As Boolean
    IsEMRProxy = mobjLogin.IsEMRProxy
End Property

Public Sub Relogin(ByVal FrmMainObj As Object)
    Call Me.Dispose
    Unload FrmMainObj
    Call Main
End Sub

Public Function UpdateClient(Optional ByVal blnForceUpdate As Boolean) As Boolean
    UpdateClient = mobjLogin.UpdateClient(blnForceUpdate)
End Function

Public Function Login(Optional ByVal intType As Integer, Optional strCommand As String) As ADODB.Connection
    On Error Resume Next
    If mobjLogin Is Nothing Then
        Set mobjLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If mobjLogin Is Nothing Then
        MsgBox "创建ZLLogin部件对象失败,请检查文件是否存在并且正确注册。"
        Exit Function
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    Set Login = mobjLogin.Login(intType, strCommand, , , , App.hInstance)
End Function

Public Sub Dispose()
'销毁对象
    '已经设置了销毁对象时不销毁变量，故手工销毁变量
    If Not mobjLogin Is Nothing Then
        mobjLogin.ClearVars (True)
    End If
    Set mobjLogin = Nothing
    If Not gobjWait Is Nothing Then Unload gobjWait
    Set gobjWait = Nothing
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    Set mobjLogin = CreateObject("ZLLogin.clsLogin")
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Class_Terminate()
    Set mobjLogin = Nothing
    Set gobjWait = Nothing
End Sub
