Attribute VB_Name = "mdlSockServer"
Option Explicit

'**************************
'       OEM����
'
'����    B0AEC9FA
'ҽҵ    D2BDD2B5
'����    CDD0C6D5
'����    D6D0C8ED
'��̩  BDF0BFB5CCA9
'ҽԺ    D2BDD4BA
'**************************

Public Type POINTAPI
        x As Long
        Y As Long
End Type
'---------------------------------------------------------------
'- ע���ȫ��������...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Const GWL_WNDPROC = -4
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ



Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public gstrProductName As String
Public gstrSysName As String                'ϵͳ����
Public gstrUserName As String               '�û���
Public gstrUserPwd As String                  '����
Public gstrServer As String                 '��������
Public gstrSQL    As String                 'ͨ�õ�SQL������

Public gcnOracle As ADODB.Connection     '�������ݿ�����
Public gcnZltools As ADODB.Connection     'zltools���Ӷ���,�����޸�


Public Sub Main()
    Dim objLogin As Object
    
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    
    If App.PrevInstance Then
        MsgBox " ���ݱ䶯֪ͨ�����Ѿ������� ", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    On Error Resume Next
    If objLogin Is Nothing Then
        Set objLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If objLogin Is Nothing Then
        MsgBox "����ZLLogin��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ"
        Exit Sub
    Else
        Set gcnOracle = objLogin.Login(2, CStr(Command()), , True)
        If gcnOracle Is Nothing Then
            Exit Sub
        ElseIf gcnOracle.State <> adStateOpen Then
            Exit Sub
        End If
    End If
    
    If Not IsDBA Then
        MsgBox "��ǰ����Ҫ��ʹ��DBA��¼��"
        Exit Sub
    End If
    gstrServer = objLogin.ServerName
    gstrUserName = objLogin.InputUser
    
    If IsDesinMode Then '���뻷�� ֱ��ȡHIS
        gstrUserPwd = "HIS"
    Else
        gstrUserPwd = GetDBPassword
    End If

    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "") & "���"
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    frmMain.Show
End Sub

Private Function IsDBA() As Boolean
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "DBA�ж�")
    IsDBA = Not rsTemp.EOF
    
    Exit Function
errH:
    ErrCenter
End Function
Public Function IsDesinMode() As Boolean
'���ܣ� ȷ����ǰģʽΪ���ģʽ
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
 
 
Private Function GetDBPassword() As String
    '��ȡ���ݿ�����
    Dim objRegister  As Object
    
    On Error Resume Next
    Set objRegister = CreateObject("zlRegister.clsRegister")
    If objRegister Is Nothing Then
        MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
        Exit Function
    End If
    Call SaveSetting("ZLSOFT", "����ȫ��", "��������", UCase("zlHisCrust.exe")) '����ZLRegister�������ж�
    GetDBPassword = objRegister.GetPassword(App.hInstance)
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function


