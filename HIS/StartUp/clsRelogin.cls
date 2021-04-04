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
Private mobjLogin       As Object   '��¼����
'���ԣ���ȡ������û���
Public Property Get InputUser() As String
    InputUser = mobjLogin.InputUser
End Property
'���ԣ���ȡ���������
Public Property Get InputPwd() As String
    InputPwd = mobjLogin.InputPwd
End Property
'���ԣ���ȡ����ķ�����
Public Property Get ServerName() As String
    ServerName = mobjLogin.ServerName
End Property
'���ԣ���ȡ�Ƿ�������Ҫת��
Public Property Get IsTransPwd() As Boolean
    IsTransPwd = mobjLogin.IsTransPwd
End Property
'���ԣ���ȡ������û���Ӧ�����ݿ��û�
Public Property Get DBUser() As String
    DBUser = mobjLogin.DBUser
End Property
'���ԣ���ȡ�Ƿ�ϵͳ������
Public Property Get IsSysOwner() As Boolean
    IsSysOwner = mobjLogin.IsSysOwner
End Property
'���ԣ���ȡ������ѡ���Systems
Public Property Get Systems() As String
    Systems = mobjLogin.Systems
End Property
'���ԣ���ȡ�˵���
Public Property Get MenuGroup() As String
    MenuGroup = mobjLogin.MenuGroup
End Property
'���ԣ���ȡ�˵���
Public Property Get DeptName() As String
    DeptName = mobjLogin.DeptName
End Property
'���ԣ���ȡEMR����
Public Property Get EMR() As Object
    Set EMR = mobjLogin.EMR
End Property
'���ԣ�����EMR����
Public Property Set EMR(ByRef objEMR As Object)
    Set mobjLogin.EMR = objEMR
End Property
'���ԣ�EMR����֤�û�
Public Property Get EMRUser() As String
    EMRUser = mobjLogin.EMRUser
End Property

'���ԣ�EMR����֤���룬�������ݿ���������[DBPASSWORD]ǰ׺
Public Property Get EMRPwd() As String
    EMRPwd = mobjLogin.EMRPwd
End Property

'���ԣ�EMRUser��EMRPwd�Ƿ����ڴ����û����統LIS������ʱ��ʹ��ZLHIS�˻�����ʼ��EMR,��ʱ������ΪTRUE
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
        MsgBox "����ZLLogin��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ"
        Exit Function
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    Set Login = mobjLogin.Login(intType, strCommand, , , , App.hInstance)
End Function

Public Sub Dispose()
'���ٶ���
    '�Ѿ����������ٶ���ʱ�����ٱ��������ֹ����ٱ���
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