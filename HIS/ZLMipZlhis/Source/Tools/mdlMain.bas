Attribute VB_Name = "mdlMain"
Option Explicit

'Public gcnOracle As ADODB.Connection
Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrSysName As String               '��ǰ�û�����
Public gobjRegister As Object               'ע����Ȩ����zlRegister
Public gobjComLib As Object

Public Sub Main()
        
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    '����ע�Ჿ��(���ڵ�¼ʱ��ȡ���Ӷ���)
    On Error Resume Next
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    If gobjRegister Is Nothing Then
        Err.Clear
        MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
        Exit Sub
    End If
    If frmLogin.ShowLogin() = False Then Exit Sub
    Call gobjComLib.InitCommon(gcnOracle)
    If gcnOracle.State <> adStateOpen Then
        Exit Sub
    End If
    frmMain.Show
    
End Sub

