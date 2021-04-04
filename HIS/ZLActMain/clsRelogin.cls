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
Private mobjCurEmr As Object 'EMR新版电子病历
Private mstrUserName As String 'InputUser属性
Private mstrInputPwd As String 'InputPwd属性
Private mstrServerName As String 'ServerName属性
Private mstrDBUser As String 'DBUser属性
Private mblnTransPwd As Boolean 'blnTransPwd属性
Private mstrMenuGroup As String '菜单组
Private mblnSysOwner As Boolean '是否系统所有者
Private mstrSystems As String '多帐套选择的系统
Private mstrDeptName As String '登录部门
Private mblnEMRProxy As Boolean
Private mstrEMRPwd   As String
Private mstrEMRUser  As String
'属性，获取输入的用户名
Public Property Get InputUser() As String
    InputUser = mstrUserName
End Property

Friend Property Let InputUser(ByVal strUserName As String)
    mstrUserName = strUserName
End Property
'属性，获取输入的密码
Public Property Get InputPwd() As String
    InputPwd = mstrInputPwd
End Property

Friend Property Let InputPwd(ByVal strInputPwd As String)
    mstrInputPwd = strInputPwd
End Property
'属性，获取输入的服务器
Public Property Get ServerName() As String
    ServerName = mstrServerName
End Property

Friend Property Let ServerName(ByVal strServerName As String)
    mstrServerName = strServerName
End Property
'属性，获取输入的用户对应的数据库用户
Public Property Get DBUser() As String
    DBUser = mstrDBUser
End Property

Friend Property Let DBUser(ByVal strDBUser As String)
    mstrDBUser = strDBUser
End Property
'属性，获取是否密码需要转换
Public Property Get IsTransPwd() As Boolean
    IsTransPwd = mblnTransPwd
End Property

Friend Property Let IsTransPwd(ByVal blnTransPwd As Boolean)
    mblnTransPwd = blnTransPwd
End Property
'属性，获取是否系统所有者
Public Property Get IsSysOwner() As Boolean
    IsSysOwner = mblnSysOwner
End Property

Friend Property Let IsSysOwner(ByVal blnSysOwner As Boolean)
    mblnSysOwner = blnSysOwner
End Property
'属性，获取多帐套选择的Systems
Public Property Get Systems() As String
    Systems = mstrSystems
End Property

Friend Property Let Systems(ByVal strSystems As String)
    mstrSystems = strSystems
End Property

'属性，获取登录部门
Public Property Get DeptName() As String
    DeptName = mstrDeptName
End Property

Friend Property Let DeptName(ByVal strDeptName As String)
    mstrDeptName = strDeptName
End Property
'属性，登录站点编号
Public Property Get NodeNo() As String
    NodeNo = gstrNodeNo
End Property

Friend Property Let NodeNo(ByRef strNodeNo As String)
    gstrNodeNo = strNodeNo
End Property
'属性，获取EMR对象
Public Property Get mobjEmr() As Object
    Set mobjEmr = mobjCurEmr
End Property
'属性，设置EMR对象
Public Property Set mobjEmr(ByRef objEMR As Object)
    Set mobjCurEmr = objEMR
End Property
'属性，EMR的验证用户
Public Property Get EMRUser() As String
    EMRUser = mstrEMRUser
End Property

Friend Property Let EMRUser(ByRef strEMRUser As String)
    mstrEMRUser = strEMRUser
End Property
'属性，EMR的验证密码，若是数据库密码增加[DBPASSWORD]前缀
Public Property Get EMRPwd() As String
    EMRPwd = mstrEMRPwd
End Property

Friend Property Let EMRPwd(ByRef strEMRPwd As String)
    mstrEMRPwd = strEMRPwd
End Property

'属性，EMRUser、EMRPwd是否属于代理用户。如当LIS单独库时，使用ZLHIS账户来初始化EMR,此时该属性为TRUE
Public Property Get IsEMRProxy() As Boolean
    IsEMRProxy = mblnEMRProxy
End Property

Friend Property Let IsEMRProxy(ByRef blnEMRProxy As Boolean)
    mblnEMRProxy = blnEMRProxy
End Property

Public Function UpdateClient(Optional ByVal blnForceUpdate As Boolean) As Boolean
    gblnTimer = True
    UpdateClient = UpdateZLHIS(OS.ComputerName, True, blnForceUpdate)
    gblnTimer = False
End Function

Private Sub Class_Terminate()
    If Not mobjEmr Is Nothing Then
        mobjEmr.Dispose
        Set mobjEmr = Nothing
    End If
End Sub

