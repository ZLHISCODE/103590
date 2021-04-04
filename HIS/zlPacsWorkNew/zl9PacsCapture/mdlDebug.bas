Attribute VB_Name = "mdlDebug"
Option Explicit

#Const DebugVer = True

Public Const M_STR_MODULE_MENU_TAG As String = "�ɼ�"
Public Const G_STR_HINT_TITLE As String = "��ʾ"

Public Const G_STR_REG_PATH_PUBLIC As String = "����ģ��\zl9PacsCapture"
Public Const G_STR_REG_PATH_PRIVATE As String = "˽��ģ��\zl9PacsCapture\"


Public Enum TDockState                      '��������״̬
    dsClosed = 0    '�ر�
    dsOpen = 1      '��
    dsClosing = 2   '�ر���
End Enum


Public gcnVideoOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��

Public gobjOwner As Object
Public glngRootHandle As Long
Public glngSys As Long
Public gstrPrivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModule As Long                   'ģ���
Public glngDepartId As Long                 '��ǰ����ID


Public gobjCapturePar As clsCaptureParameter    '��Ƶ�ɼ���ز�������

Public gobjNotifyEvent As clsNotifyEvent         '��Ϣ֪ͨ��������
Public gobjVideo As frmWork_Video                '��Ƶ�ɼ�����
'Public gobjGlobal As clsGlobal

Public glngCurVideoContainerHwnd As Long        '��ǰ��Ƶ���ڵ��������ھ��
'Public glngNextVideoContainerHwnd As Long       '��ǰ��Ƶ���ڴ���Z���е���һ���д��ھ��
Public gblnDockingState As TDockState           '�Ƿ��ڵ�������״̬

Public glngInstanceCount As Long
Public gblnOpenDebug As Boolean
Public gobjZOrder As Scripting.Dictionary
Public gblnIsQuitModule As Boolean

'debug property
Public gstrHotKeyTest As String

Private gcnOracle As ADODB.Connection
Private objTest As clsPacsCapture




Private Function IsDebugMode() As Boolean

    IsDebugMode = False

    On Error Resume Next

    Debug.Print 1 / 0

    If err.Number <> 0 Then

        IsDebugMode = True

    End If

End Function


Public Sub Main()
BUGEX "Main 1", True
    If UCase(Command()) = "DEBUG" Or IsDebugMode Then
BUGEX "Main Enter Debug", True
        frmTestLogin.Show
    End If

BUGEX "Main End", True




'    If Not IsDebugMode Then Exit Sub
'
'    Set gcnOracle = New ADODB.Connection
'
'    OraDataOpen "", "zlhis", "HIS"
'
'    Set gcnVideoOracle = gcnOracle
'
'    frmTestWindow.Show

End Sub

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String

    On Error Resume Next
    err = 0
    DoEvents
    With gcnOracle
        If .State = 1 Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '���������Ϣ
            MsgboxEx Nothing, err.Description, vbInformation, G_STR_HINT_TITLE

            OraDataOpen = False
            Exit Function
        End If
    End With

    err = 0
    On Error GoTo errhand

    OraDataOpen = True
    Exit Function

errhand:
    MsgboxEx Nothing, err.Description, vbOKOnly, G_STR_HINT_TITLE
    OraDataOpen = False
    err = 0
End Function


Public Sub OutputDebug(ByVal strMethob As String, objErr As ErrObject)
    If gblnOpenDebug Then
        OutputDebugString "[" & App.ProductName & "]" & strMethob & "��" & objErr.Description
    End If
End Sub


'Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
'    If gblnOpenDebug Or blnIsForce Then
'        OutputDebugString Now & " |---> " & strDebug
'    End If
'End Sub

Public Sub RaiseErr(objErr As ErrObject)
    Call err.Raise(objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext)
End Sub
