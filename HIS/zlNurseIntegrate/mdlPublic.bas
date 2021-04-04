Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gobjGrid As Object

Public gstrDBUser As String
Public gstrNodeNo As String          '��ǰվ���ţ����δ��������վ�㣬��Ϊ"-"
Public gstrSysName As String
Public gblnAlone As Boolean '�Ƿ��������

Public gstrIntergrateIP As String    '�ƶ�����ID��ַ
Public gobjScriptControl  As MSScriptControl.ScriptControl
Public gstrRelatedUserID As String  '���廤����ԱID
Public gstrRelatedUnitID As String  '���廤����ID
Public gstrRelatedPatientID As String  '���廤����ID

Public glngPid As Long '��ǰ����Ľ���ID
Public gcllHideFrmsEx As Collection '���д��弯��

Public Type TYPE_INTERGRATE_USER_INFO  '���廤������Ա��Ϣ(�ӿ�UserLogin����)
    id As String
    UserName As String
    Name As String
    Sex As String
    Cookie As String
End Type
Public IntergrateUserInfo As TYPE_INTERGRATE_USER_INFO

Public Type TYPE_USER_INFO
    id As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
    רҵ����ְ�� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public glnglpPrevWndProc As Long
Public glngSCMIZE As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4&
Public Const SC_MAXIMIZE = &HF030& '���
Public Const SC_MINIMIZE = &HF020& '��С��
Public Const SC_RESTORE = &HF120& '��ԭ

Public Const GW_CHILD = &H5
Public Const GW_OWNER = &H4
Public Const GW_HWNDNEXT = 2
'���һ�����ڵľ�����ô�����ĳԴ�������ض��Ĺ�ϵ
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'�ҳ�ĳ�����ڵĴ�����(�̻߳����)�����ش����ߵı�־����
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'�ú���ȷ�������Ĵ��ھ���Ƿ��ʶһ���Ѵ��ڵĴ��ڡ�
'����ֵ��������ھ����ʶ��һ���Ѵ��ڵĴ��ڣ�����ֵΪ���㣻������ھ��δ��ʶһ���Ѵ��ڴ��ڣ�����ֵΪ��
Public Declare Function isWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
'����ֵ�����ָ���Ĵ��ڼ��丸���ھ���WS_VISIBLE��񣬷���ֵΪ���㣻���ָ���Ĵ��ڼ��丸���ڲ�����WS_VISIBLE��񣬷���ֵΪ�㡣���ڷ���ֵ�����˴����Ƿ����Ws_VISIBLE�����ˣ���ʹ�ô��ڱ����������ڸǣ���������ֵҲΪ���㡣
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'�ú��������ж�ָ���Ĵ����Ƿ�������ܼ��̻��������?
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
'��ָ�����ڵĽṹ��ȡ����Ϣ
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'ȡ��һ������ı��⣨caption�����֣�����һ���ؼ������ݣ���vb��ʹ�ã�ʹ��vb�����ؼ���caption��text���ԣ�
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'�ڴ��ڽṹ��Ϊָ���Ĵ���������Ϣ
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE        As Long = (-20)
Private Const GWL_STYLE          As Long = (-16)
Private Const WS_EX_TOOLWINDOW   As Long = &H80
Private Const WS_EX_CONTEXTHELP  As Long = &H400
Private Const WS_MAXIMIZEBOX     As Long = &H10000
Private Const WS_MINIMIZEBOX     As Long = &H20000
Private Const WS_SYSMENU         As Long = &H80000
Private Const WS_THICKFRAME      As Long = &H40000
Private Const WS_CAPTION = &HC00000
'��ȡ��ǰ����һ��Ψһ�ı�ʶ��
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
' �ú���ö��������Ļ�ϵĶ��㴰�ڣ��������ھ�����͸�Ӧ�ó�����Ļص�����
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'�Ƿ���64λ���̣�Is64bit��
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long

'ע������**********************************
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const KEY_ALL_ACCESS = (&H20000 Or &H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H20) And (Not &H100000)
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
'*****************************************************************
'*****��������ע���������õ���API����****************************
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'������ý������ƻ�ȡ
Public gstrExeName As String 'ִ�г��������
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 1024
End Type
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     '������(SC_MAXIMIZE)������ (�����Լ���Ҫ���ĳ���С����ԭ�ȡ�)
    If wParam = glngSCMIZE Then Exit Function
    WindowProc = CallWindowProc(glnglpPrevWndProc, hwnd, uMsg, wParam, lParam)
End Function

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������ض���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    
    Err = 0: On Error Resume Next
    
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    Set gobjGrid = GetObject("", "zl9Comlib.clsGrid")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then
        If gblnAlone = True Then Call gobjComlib.InitCommon(gcnOracle)
        zlGetComLib = True: Exit Function
    End If
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    If gblnAlone = True Then
        Call gobjComlib.InitCommon(gcnOracle)
    End If
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    
    If Not gobjComlib Is Nothing Then
        zlGetComLib = True
        gstrNodeNo = gobjComlib.gstrNodeNo
    End If
    Err = 0: On Error GoTo 0
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    '����:������Ա����,����ö��ŷ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    If str���� <> "" Then
        strSql = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��Ա����", str����)
    Else
        strSql = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��Ա����", UserInfo.id)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrH
    Set rsTmp = gobjDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.id = rsTmp!id
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = "" & rsTmp!����
            UserInfo.���� = "" & rsTmp!����
            UserInfo.����ID = Val("" & rsTmp!����ID)
            UserInfo.������ = "" & rsTmp!������
            UserInfo.������ = "" & rsTmp!������
            UserInfo.���� = Get��Ա����
            UserInfo.רҵ����ְ�� = "" & rsTmp!רҵ����ְ��
            GetUserInfo = True
        End If
    End If
    gstrDBUser = UserInfo.�û���
    Exit Function
ErrH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
 '���:strInfor-ԭ��
 '         lngStart-ֱʼλ��
'         lngLen-����
'����:�Ӵ�
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo ErrH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
ErrH:
    Err.Clear
    SubB = ""
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SendPostUrl(ByVal strUrl As String, ByVal strParams As String, objHttp As XMLHTTP, strErrMsg As String, Optional blnCookie As Boolean = False, Optional ByVal strIP As String = "") As Boolean
'���ܣ�����URL��ַ������������URL
'������
'       strUrl-URL��ַ��strParams ������JSON��ʽ��
'       objHttp��XMLHTTP����
    Dim oXmlHttp As XMLHTTP
    Dim intPos As Integer  '���Ӵ�����������
    If strIP = "" Then strIP = gstrIntergrateIP
    If blnCookie = True Then
'        UserLogin
    End If
    Set oXmlHttp = New MSXML2.ServerXMLHTTP
    oXmlHttp.open "POST", strUrl, False   '��ʼ��HTTP����
    oXmlHttp.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    If blnCookie = True Then
        oXmlHttp.setRequestHeader "Cookie", IntergrateUserInfo.Cookie
    End If
    oXmlHttp.setTimeouts 5000, 10000, 10000, 10000
    '��һ����ֵ: ����DNS���ֵĳ�ʱʱ��
    '�ڶ�����ֵ: ����Winsock���ӵĳ�ʱʱ��
    '��������ֵ: �������ݵĳ�ʱʱ��
    '���ĸ���ֵ: ����response�ĳ�ʱʱ��
    
    On Error Resume Next
RestartSend:  '�������ӣ��������ӳ�ʱ������һ�ν������ӣ�����������Σ�
    Err.Clear
    Call oXmlHttp.send(strParams)
    If Err <> 0 Then
        If Err.Number = -2147012894 Then  'IP��ַ�����ڻ����ӳ�ʱ
            If intPos > 0 Then
                strErrMsg = "�������廤�������ʧ�ܣ�����������ԭ���£�" & vbCrLf & _
                    "1���������õ�IP��ַ(" & strIP & ")�޷����ӣ�����ϵ����Ա���½�������" & vbCrLf & _
                    "2���������������������ӷ�������ʱ��������ˢ�»��ٴβ���" & vbCrLf & "��ϸ��Ϣ��" & Err.Description
            Else
                intPos = intPos + 1
                GoTo RestartSend
            End If
        Else 'IP��ַ��ȷ�����ǲ��Ƿ�����
            strErrMsg = Err.Description & "���������õ�IP��ַ�Ƿ����ƶ���������ַ��" & vbCrLf & "IP��ַ��" & strIP
        End If
        Err.Clear
        Exit Function
    End If
    
    'ReadyState
    'HTTP �����״̬.��һ�� XMLHttpRequest ���δ���ʱ��������Ե�ֵ�� 0 ��ʼ��ֱ�����յ������� HTTP ��Ӧ�����ֵ���ӵ� 4��
    '0   Uninitialized   ��ʼ��״̬��XMLHttpRequest �����Ѵ������ѱ� abort() �������á�
    '1   Open    open() �����ѵ��ã����� send() ����δ���á�����û�б����͡�
    '2   Sent    Send() �����ѵ��ã�HTTP �����ѷ��͵� Web ��������δ���յ���Ӧ��
    '3   Receiving ������Ӧͷ�����Ѿ����յ�?��Ӧ�忪ʼ���յ�δ���?
    '4   Loaded  HTTP ��Ӧ�Ѿ���ȫ���ա�
    
    If oXmlHttp.readyState = 4 Then '���ݽ��ճɹ�
        If oXmlHttp.Status = "200" Then
            Set objHttp = oXmlHttp
        Else
            If oXmlHttp.Status = "404" Then 'IP��ַ��ȷ����ʱ����ĵ�ַ����ȷ
                strErrMsg = "Http��ַ����ȷ������ϵ��������̣�" & vbCrLf & "Http��ַ��" & strUrl
            Else
                strErrMsg = oXmlHttp.statusText & vbCrLf & "Http��ַ��" & strUrl
            End If
            Exit Function
        End If
    Else
        strErrMsg = "HTTP ��Ӧ����δ���������գ���������������������"
        Exit Function
    End If
    
    SendPostUrl = True
End Function


Public Function encodeURI(ByVal strValue As String) As String
'���ܣ��ַ���ת���� UTF-8 ���루URL�����ַ�ת����
    Dim strRetrun As String
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    gobjScriptControl.Language = "javascript"
    strRetrun = gobjScriptControl.Eval("encodeURI('" & strValue & "')")
    encodeURI = strRetrun
End Function

Public Function decodeURI(ByVal strValue As String) As String
'���ܣ��� UTF-8 ����ת��ΪURL�����ַ���URL�����ַ�ת����
    Dim strRetrun As String
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    gobjScriptControl.Language = "javascript"
    strRetrun = gobjScriptControl.Eval("decodeURI('" & strValue & "')")
    decodeURI = strRetrun
End Function

Public Function encodeURIComponent(ByVal strValue As String) As String
'���ܣ��ַ���ת���� UTF-8 ���루URL�����ַ�ת����
    Dim strRetrun As String
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    gobjScriptControl.Language = "javascript"
    strRetrun = gobjScriptControl.Eval("encodeURIComponent('" & strValue & "')")
    encodeURIComponent = strRetrun
End Function

Public Function decodeURIComponent(ByVal strValue As String) As String
'���ܣ��ַ���ת���� UTF-8 ���루URL�����ַ�ת����
    Dim strRetrun As String
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    gobjScriptControl.Language = "javascript"
    strRetrun = gobjScriptControl.Eval("decodeURIComponent('" & strValue & "')")
    decodeURIComponent = strRetrun
End Function

Public Function AnalysisJavaScriptEvent(ByVal strParam As String, objPopup As clsPopup) As Boolean
'���ܣ�
'strParam��ʽ��
'{
'  type: "CloseDialog" || "ShowDialog", // CloseDialog �رյ���  ShowDialog �򿪵���
'  moduleUrl: "/shiftReport", //����Url
'  title: "���౨��",
'  width: "100" || null,
'  height: "100" || null,
'  minimal: true,  //���
'  max: false,     //��С��
'  isRefresh: true  //�Ƿ�ˢ�¸�����
'  data: "xxxxxxxxxxxx"  //�򿪵�������Ҫ���ϵĲ���
'}
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl

    On Error GoTo ErrHand
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strParam & ";"
        objPopup.PopupParams = strParam
        objPopup.PopupType = "" & .Eval("json.type")
        objPopup.PopupModuleUrl = "" & .Eval("json.moduleUrl")
        objPopup.PopupTitle = "" & .Eval("json.title")
        objPopup.PopupWidth = Val("" & .Eval("json.width"))
        objPopup.PopupHeight = Val("" & .Eval("json.height"))
        objPopup.PopupMinimal = .Eval("json.minimal") 'IIf(UCase("" & .Eval("json.minimal")) = "TRUE", True, False)
        objPopup.PopupMax = .Eval("json.max") 'IIf(UCase("" & .Eval("json.max")) = "TRUE", True, False)
        objPopup.PopupIsRefresh = .Eval("json.isRefresh") ' IIf(UCase("" & .Eval("json.isRefresh")) = "TRUE", True, False)
        objPopup.PopupData = "" & encodeURIComponent(.Eval("json.data"))
        objPopup.PopupParentUrl = decodeURIComponent("" & .Eval("json.ParentUrl"))
        objPopup.PopupParentParam = "" & encodeURIComponent(.Eval("json.originParams"))
        objPopup.PopupPatientID = "" & .Eval("json.PatientID")
        objPopup.PopupUnitID = "" & .Eval("json.LessionID")
        objPopup.PopupUserID = "" & .Eval("json.UserID")
    End With
    AnalysisJavaScriptEvent = True
    Exit Function
ErrHand:
    MsgBox Err.Description & vbCrLf & "Json��" & strParam, vbInformation, gstrSysName
End Function

Public Sub WriteBusinessLOG(ByVal strFunc As String, ByVal strInput As String, ByVal strParams As String, Optional ByVal strOutput As String = "")
    '���ܣ���¼��־�ļ�����Ҫ���ڽӿڵ���
    '�������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim strLog As String, blnBeginWrite As Boolean
    
    '˵�����ò���Ϊ���廤��ӿ�ר�ýӿڣ�ͨ����ǿ���������־�ߵ��ù���������������ڵͰ汾�����������־
    strLog = "URL��" & strInput & Space(6) & "��Σ�" & strParams & Space(6) & IIf(strOutput <> "", "���Σ�" & strOutput, "")
    On Error Resume Next
    Call gobjComlib.LogWrite("�°滤ʿ����վ�����ƶ����廤��ӿڸ�����־", "�°滤ʿ����վ", strFunc, strLog)
    If Err <> 0 Then
        blnBeginWrite = True
        Err.Clear
    End If
    If blnBeginWrite Then
        If Not objFileSystem.FolderExists("C:\���廤����־") Then Call objFileSystem.CreateFolder("C:\���廤����־")
        strFileName = "C:\���廤����־\" & Format(Date, "yyyyMMdd") & ".LOG"
        If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
        Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
        strDate = Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine (String(50, "-"))
        objStream.WriteLine ("  ����:" & strFunc)
        objStream.WriteLine ("  URL:" & strInput)
        objStream.WriteLine ("  ���:" & strParams)
        objStream.WriteLine ("  ����:" & strOutput)
        objStream.WriteLine (String(50, "-"))
        objStream.Close
        Set objStream = Nothing
        If Err <> 0 Then Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
'���ܣ� ͨ��PIDö�������ľ��,������Ҫ�Ĵ��ڣ���ҳ�ؼ������ã�ֻҪˢ�º�ͻ�ȡ������������ˣ�
    Dim lngPid As Long
    Dim strText As String * 255
    '��ҳ����ÿ��ˢ��һ��ҳ�涼��һ���½���
    GetWindowThreadProcessId hwnd, lngPid
    If glngPid = lngPid Then
        If isWindow(hwnd) <> 0 Then
            If IsWindowVisible(hwnd) <> 0 Then
                If isNormalWindow(hwnd) Then
                Else
                    gcllHideFrmsEx.Add hwnd
                End If
            End If
        End If
    End If
    EnumWindowsProc = True
End Function

Public Function isNormalWindow(ByVal lngHwnd As Long) As Boolean
'�ų�����ؼ��Ĵ������
'�ų�DTPicker����������ѡ����棬�ý���ͨ��API�ж����б������ģ�ͨ��SPY++������û�еģ���ʱ�ų��ô���
    Dim strText As String * 256
    Dim strTmp As String
    On Error Resume Next
    If GetWindowLong(lngHwnd, GWL_STYLE) And WS_CAPTION Then
        Call GetWindowText(lngHwnd, strText, 255)
        strTmp = gobjComlib.zlStr.TruncZero(strText)
        isNormalWindow = strTmp <> ""
    Else
        isNormalWindow = False
    End If
End Function

Public Sub GetAllVisibleWindow(ByVal lngPid As Long)
    glngPid = lngPid
    Set gcllHideFrmsEx = New Collection
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub

 Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '���ܣ��Ƿ���64λϵͳ
    '���أ�
    '******************************************************************************************************************
    Dim handle As Long
    Dim bolFunc As Boolean
        
    bolFunc = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), bolFunc
    End If
    Is64bit = bolFunc
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'ע������
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'*****�¼�ע�����
Public Sub Createnewkey(ip As Long, snewkeyname As String)
    Dim hnewkey As Long
    Dim retval As Long
    retval = RegCreateKey(ip, snewkeyname, hnewkey)
    If retval = 0 Then
        RegCloseKey (hnewkey) '�ر����潨����򿪵���
    End If
End Sub
'ʵ������HKEY_CURRENT_USER�½�����"xiaopeng"
'����Ϊ createnewkey HKEY_CURRENT_USER ,"xiaopeng"
'******************************************************************
'*******ɾ��ע�����***********************************************
Public Function Deletekey(ip As Long, skeyname As String)
    Dim hKey As Long
    Dim retval As Long
    retval = RegOpenKeyEx(ip, skeyname, 0, KEY_ALL_ACCESS, hKey)
    If retval = 0 Then
        RegDeleteKey ip, skeyname
    End If
End Function
'ʵ����ɾ�����潨����HKEY_CURRENT_USER�µ���"xiaopeng"
'����Ϊ deletekey HKEY_CURRENT_USER ,"xiaopeng"
'******************************************************************
'********�½�,������ֵ����*****************************************
Public Sub Setkeyvalue(ByVal ip As Long, ByVal keyname As String, ByVal valuename As String, ByVal valuesetting As Variant, ByVal valuetype As Long)
    Dim retval As Long
    Dim hKey As Long
    If RegOpenKeyEx(ip, keyname, 0, KEY_ALL_ACCESS, hKey) > 0 Then Exit Sub
    Select Case valuetype
        Case REG_SZ
             RegSetValueExString hKey, valuename, 0&, REG_SZ, valuesetting, Len(valuesetting)
        Case REG_DWORD
             RegSetValueExLong hKey, valuename, 0, valuetype, valuesetting, 4
    End Select
    RegCloseKey (hKey)
End Sub
'ʵ������HKEY_CURRENT_USER�µ���"xiaopeng"�н�����Ϊ"redice",��ֵΪ"is xiaopeng",����ΪREG_SZ���¼�
'����Ϊ setkeyvalue HKEY_CURRENT_USER ,"xiaopeng" ,"redice","is xiaopeng",REG_SZ
'����:��HKEY_CURRENT_USER�µ���"xiaopeng"�н�����Ϊ"ceshi",��ֵΪ2,����ΪREG_DWORD���¼�
'����Ϊ"setkeyvalue HKEY_CURRENT_USER,"xiaopeng","ceshi",2,REG_DWORD
'********************************************************************************
'*********ɾ����ֵ����*********************************************************
Public Sub Deletevalue(ByVal ip As Long, ByVal keyname As String, ByVal valuename As String)
    Dim retval As Long
    Dim hKey As Long
    retval = RegOpenKeyEx(ip, keyname, 0, KEY_ALL_ACCESS, hKey)
    If retval > 0 Then
        Exit Sub
    End If
    RegDeleteValue hKey, valuename
    RegCloseKey hKey
End Sub
'ʵ����ɾ��HKEY_CURRENT_USER�µ���"xiaopeng"����Ϊ"redice"���¼�
'����Ϊ deletevalue HKEY_CURRENT_USER ,"xiaopeng","redice"
'******************************************************************
'**********��ѯ�Ѵ��ڵ���ֵ����************************************
Public Function getvalue(ByVal ip As Long, keyname As String, valuename As String, ByVal valuetype As Long) As String
    Dim retval As Long
    Dim hKey As Long
    Dim valuesetting As Variant
    Dim cddata As Long
    Dim lvalue As Long
    Dim svalue As String
    
    retval = RegOpenKeyEx(ip, keyname, 0, KEY_ALL_ACCESS, hKey)
    If retval > 0 Then
        getvalue = ""
        Exit Function
    End If
    retval = RegQueryValueEx(hKey, valuename, 0, valuetype, ByVal vbNullString, cddata)
    If retval <> 0 Then
        RegCloseKey hKey
        Exit Function
    End If
    Select Case valuetype
        Case REG_SZ
            svalue = String(cddata, Chr(0))
            RegQueryValueEx hKey, valuename, 0, valuetype, ByVal svalue, cddata
            valuesetting = Left$(svalue, cddata)
            getvalue = CStr(valuesetting)
        Case REG_DWORD
            RegQueryValueEx hKey, valuename, 0, valuetype, lvalue, cddata
            valuesetting = lvalue
            getvalue = CStr(valuesetting)
    End Select
End Function
'ʵ������ȡHKEY_CURRENT_USER�µ���"xiaopeng"����Ϊ"redice"���¼��ļ�ֵ
'����Ϊ getvalue HKEY_CURRENT_USER ,"xiaopeng","redice"

'--------------------------------------------------------------------------------------------------------------------------------------------------------
'���ܣ���ȡ������õĽ�����
'--------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
'szExeName ���ص������ļ���,û�ҵ� ���� ""
'szPathName ���ص�����·����,��"\"  ����,û�ҵ� ���� ""
    Dim my As PROCESSENTRY32
    Dim hProcessHandle As Long
    Dim success As Long
    Dim l As Long

    l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If l Then
        my.dwSize = 1060
        If (Process32First(l, my)) Then
            Do
            If my.th32ProcessID = processID Then
               CloseHandle l
               szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
               For l = Len(szExeName) To 1 Step -1
                   If Mid$(szExeName, l, 1) = "\" Then Exit For
               Next l
               szPathName = Left$(szExeName, l)
               Exit Sub
            End If
            Loop Until (Process32Next(l, my) < 1)
        End If
        CloseHandle l
    End If
End Sub

Public Function SetWBIEVerSion(ByVal strExeName As String, strMsg As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�����������汾IE11,ֻ���ڷ�IDE���������ò���������
    '���أ�
    '******************************************************************************************************************
    Dim strLocal As String
    If Is64bit Then
        strLocal = "SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
    Else
        strLocal = "SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
    End If
    If getvalue(HKEY_LOCAL_MACHINE, strLocal, strExeName, REG_DWORD) <> "11000" Then
        '�½�ע�����
        Call Setkeyvalue(HKEY_LOCAL_MACHINE, strLocal, strExeName, "11000", REG_DWORD)
        '���ע����Ƿ�д��ɹ�
        If getvalue(HKEY_LOCAL_MACHINE, strLocal, strExeName, REG_DWORD) <> "11000" Then
            strMsg = "�޸�WebBrowser��ҳ�ؼ�Ĭ��ʹ��IE11�����ʧ�ܣ����޷��������廤��ҳ�����ݣ����ֹ����ע���" & vbCrLf & _
                "32λϵͳ��HKEY_LOCAL_MACHINE" & strLocal & vbCrLf & _
                "64λϵͳ��HKEY_LOCAL_MACHINE" & strLocal & vbCrLf & _
                "�������ݣ�[����]-exe��������[����]-REG_DWORD    [��ֵ]-11000"
            Exit Function
        End If
    End If
    
    SetWBIEVerSion = True
End Function

