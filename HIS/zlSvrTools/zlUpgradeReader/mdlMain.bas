Attribute VB_Name = "mdlMain"
Option Explicit
Public gstrDBUser As String
Public gcnOracle As ADODB.Connection
Public gstrSysname As String '��������

Public gstrSystems As String 'ϵͳ����
Public gstr�û���λ���� As String '�ѵ�¼ʱ��Ϊ��

Public mclsAppTool As New zl9AppTool.clsAppTool
Public rsMenu As ADODB.Recordset
Public rsMenuPEIS As ADODB.Recordset

'-------------------------------------------------------------
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Const WinStyle = &H40000

'---��дINI�ļ���API����
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#End If
'----------------------
Public gobjRegister As Object               'ע����Ȩ����zlRegister
Public Enum �����嵥
    ���������嵥 = 10
    �ֵ������ = 11
    ��Ϣ�շ����� = 12
    ϵͳѡ������ = 13
    EXCEL������ = 14
    ���ز������� = 15
End Enum

Public Sub Main()
    Dim objLogin As Object
    
    gstrDBUser = ""
    gstrSysname = "����˵���Ķ���"
    gstr�û���λ���� = ""
    On Error Resume Next
    If objLogin Is Nothing Then
        Set objLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If objLogin Is Nothing Then
        Set gcnOracle = New ADODB.Connection
    Else
        Set gcnOracle = objLogin.Login(0, CStr(Command()))
        If gcnOracle Is Nothing And Not objLogin.IsCancel Then
            Exit Sub
        ElseIf gcnOracle Is Nothing Then 'ȡ���˳����Էǵ�½ģʽ����
            Set gcnOracle = New ADODB.Connection
        End If
    End If
    
    If gcnOracle.State = adStateOpen Then
        gstrSystems = objLogin.Systems
        gstrDBUser = objLogin.DBUser
        Set rsMenu = MenuGranted(objLogin.MenuGroup)
        Set rsMenuPEIS = MenuGranted("PEIS")
        
        If rsMenu.EOF Then
            MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysname
            Exit Sub
        End If
        gstr�û���λ���� = zlRegInfo("��λ����", , -1)
        Call frmMain.Show_me(1) '0- δ��¼��ʽ 1���ѵ�¼��ʽ
    Else
        Call frmMain.Show_me(0) '0- δ��¼��ʽ 1���ѵ�¼��ʽ
    End If
End Sub

Private Function MenuGranted(ByVal strMenuGroup As String) As ADODB.Recordset
    '-------------------------------------------------------------
    '���ܣ�������Ȩʹ�ò���װ�Ĳ���������������Ȩʹ�õĲ˵�����
    '������ע����
    '-------------------------------------------------------------
    Dim ArrCommand
    Dim StrSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim IntCount As Integer
    Dim strSystems As String
    Dim gstrMenuSys As String
    Dim BlnOnlySys As Boolean 'ֻ�б���ϵͳ
    Dim strSYS As String
    
    BlnOnlySys = (gstrSystems = "REPORT")
    If BlnOnlySys Then
        strSystems = " '0'"
    Else
        strSystems = Replace(gstrSystems, "','", ",")
    End If
    
    '--����Ȩ�޲˵�--
    With rsTemp
        If strMenuGroup <> "" Then gstrMenuSys = strMenuGroup
        strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
        If strObjs = "" Then strObjs = "'Zl9Common'"
        strObjs = Replace(strObjs, "','", ",")

        StrSQL = "SELECT ���, Id AS ���, Nvl(�ϼ�id, 0) AS �ϼ�, ����, Decode(Nvl(�̱���,'��'),'��',����,�̱���) As �̱���, ���, ˵��, Nvl(ģ��, 0) AS ģ��, Nvl(ϵͳ, 0) AS ϵͳ, " & _
                 "        Nvl(ͼ��, 0) AS ͼ��, nvl(����,'0') as ����, Decode(Upper(Rtrim(����)), 'ZL9REPORT', 1, 0) AS ���� " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu('" & gstrMenuSys & "', " & strSystems & ", " & strObjs & ") As " & _
                 " Zltools.t_Menu_Rowset)) " & _
                 " ORDER BY ���, Id"

        If .State = adStateOpen Then .Close
        .Open StrSQL, gcnOracle, adOpenKeyset
    End With
    
    Set MenuGranted = rsTemp
    
End Function

Public Sub WriteToIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
''дINI�ļ�
    Dim buff As String * 128
    buff = Trim(Value) + Chr(0)
    WritePrivateProfileString Section, Key, buff, Filename

End Sub

Public Function ReadFromIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String) As String
''��INI�ļ�
    Dim i As Long
    Dim buff As String * 128
    GetPrivateProfileString Section, Key, "", buff, 128, Filename
    i = InStr(buff, Chr(0))
    ReadFromIni = Trim(Left(buff, i - 1))
End Function
