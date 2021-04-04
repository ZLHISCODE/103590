Attribute VB_Name = "mdlPublic"
Option Explicit
Public gobjFile As New FileSystemObject

Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '��¼Windows����ϵͳ�����������ļ��ĸ�ʽ�͹�����Ϣ����Ҫ��¼��ͬ�ļ����ļ�����׺����֮��Ӧ��Ӧ�ó��������Ӽ��ɷ�Ϊ���࣬һ�����Ѿ�ע��ĸ����ļ�����չ���������Ӽ�ǰ�涼��һ������������һ���Ǹ����ļ������й���Ϣ��
    HKEY_CURRENT_USER = &H80000001 '�˸��������˵�ǰ��¼�û����û������ļ���Ϣ����Щ��Ϣ��֤��ͬ���û���¼�����ʱ��ʹ���Լ��ĸ��Ի����ã������Լ������ǽֽ���Լ����ռ��䡢�Լ��İ�ȫ����Ȩ�޵ȡ�
    HKEY_LOCAL_MACHINE = &H80000002 '�˸��������˵�ǰ��������������ݣ���������װ��Ӳ���Լ����������á���Щ��Ϣ��Ϊ���е��û���¼ϵͳ����ġ���������ע��������Ӵ�Ҳ������Ҫ�ĸ�����
    HKEY_USERS = &H80000003 '�˸�������Ĭ���û�����Ϣ��Default�Ӽ�����������ǰ��¼�û�����Ϣ��
    HKEY_PERFORMANCE_DATA = &H80000004 '��Windows NT/2000/XPע�������Ȼû��HKEY_DYN_DATA����������ȴ������һ����Ϊ��HKEY_ PERFOR MANCE_DATA����������ϵͳ�еĶ�̬��Ϣ���Ǵ���ڴ��Ӽ��С�ϵͳ�Դ���ע����༭���޷������˼�
    HKEY_CURRENT_CONFIG = &H80000005  '�˸���ʵ������HKEY_LOCAL_MACHINE�е�һ���֣����д�ŵ��Ǽ������ǰ���ã�����ʾ������ӡ���������������Ϣ�ȡ������Ӽ���HKEY_LOCAL_ MACHINE\ Config\0001��֧�µ�������ȫһ����
    HKEY_DYN_DATA = &H80000006 '�˸����б���ÿ��ϵͳ����ʱ��������ϵͳ���ú͵�ǰ������Ϣ���������ֻ������Windows 98�С�
End Enum

' ע����ؼ��ְ�ȫѡ��...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                       
' ����ֵ...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0

'ע�����������
Private Enum REGValueType
    REG_NONE = 0                       ' No value type
    REG_SZ = 1 'Unicode���ս��ַ���
    REG_EXPAND_SZ = 2 'Unicode���ս��ַ���
    REG_BINARY = 3 '��������ֵ
    REG_DWORD = 4 '32-bit ����
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' ��������ֵ��
End Enum

Public Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'OpenFolder�����Ļص�����ʹ��
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const WM_USER = &H400
Public Const BFFM_SETSELECTION = (WM_USER + 102)
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_STATUSTEXT = &H4
Public Const MAX_PATH = 260

Public gstrAPIPath As String

Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function StrCSpn Lib "shlwapi.dll" Alias "StrCSpnW" (ByVal lpStr&, ByVal lpCharacters&) As Long
Public Declare Function StrCSpnI Lib "shlwapi.dll" Alias "StrCSpnIW" (ByVal lpStr&, ByVal lpCharacters&) As Long
Public Declare Function StrRStr Lib "shell32.dll" Alias "StrRStrW" (ByVal lpStart&, ByVal lpEnd&, ByVal lpSrch&) As Long
Public Declare Function StrRStrI Lib "shell32.dll" Alias "StrRStrIW" (ByVal lpStart&, ByVal lpEnd&, ByVal lpSrch&) As Long

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

Public Sub GetServerInfo(ByVal strServer As String, ByRef setServiceName As String, strServerIp As String, ByRef strServerPort As String)
    '����:����tnsname.ora�ļ���ȡ������IP���˿ڡ�ʵ����
    '�������: strServer=������
    '�������� setServiceName = ʵ����  strServerIp = ������IP   strServerPort = �������˿�
    Dim strTxt As String, strFile As String
    Dim lngTmp As Long, strTmp As String
    
    On Error Resume Next
    
    strFile = GetTNSFile
    If strFile = "" Then Exit Sub
    strTxt = gobjFile.OpenTextFile(strFile).ReadAll
    
    strServer = UCase(strServer): strTxt = ConvertStr(strTxt) '��ʽ���ַ�
    strTxt = Mid(strTxt, InStr(1, strTxt, strServer & "="))
    '��ȡIP
    lngTmp = InStr(1, strTxt, "HOST=")
    strTmp = Mid(strTxt, lngTmp + Len("HOST="))
    strServerIp = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
    
    '��ȡ�˿�
    lngTmp = InStr(1, strTxt, "PORT=")
    strTmp = Mid(strTxt, lngTmp + Len("PORT="))
    strServerPort = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
    
    '��ȡ������
    lngTmp = InStr(1, strTxt, "SERVICE_NAME=")
    strTmp = Mid(strTxt, lngTmp + Len("SERVICE_NAME="))
    setServiceName = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
    
End Sub


Private Function GetTNSFile() As String
    '����:��ȡtnsname.ora�ļ�
    Dim strPath As String, strFile As String
    Dim rsOraHome As ADODB.Recordset, arrTmp() As String
    Dim i As Integer
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer

    '��ȡ��������tnsadmin
    strPath = Environ("TNS_ADMIN")
    If strPath <> "" Then
        strFile = strPath & "\tnsnames.ora" 'Oracle 8i����
        
        If gobjFile.FileExists(strFile) = False Then
            strFile = strPath & "NET80\ADMIN\tnsnames.ora" 'Oracle 8
        End If
        If gobjFile.FileExists(strFile) = False Then strFile = ""
    End If
        
    '��ȡע���
    If strFile = "" Then
        Set rsOraHome = New ADODB.Recordset
        With rsOraHome
            .Fields.Append "Name", adVarChar, 256 'Name
            .Fields.Append "VerSion", adInteger  '�汾
            .Fields.Append "Times", adInteger '�ڼ��ΰ�װ
            .Fields.Append "Server", adInteger '1-������,2-�ͻ���
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
            '1:��ȡ64λ��32Ŀ¼���Զ���λ��SOFTWARE\Wow6432Node\Oracle 2����ȡ32λ��32λĿ¼
            arrTmp = GetAllSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle")
            If TypeName(arrTmp) = "Empty" Then
                Exit Function
            Else
                For i = LBound(arrTmp) To UBound(arrTmp)
                    If UCase(arrTmp(i)) Like "KEY_ORA*HOME*" Then
                        intVersion = 0: intTimes = 0:  intServer = 1
                        If GetOraInfoByRegKey(arrTmp(i), intVersion, intTimes, intServer) Then
                            .AddNew Array("Name", "VerSion", "Times", "Server"), Array("\" & arrTmp(i), intVersion, intTimes, intServer)
                            .Update
                        End If
                    End If
                Next
                If UBound(arrTmp) <> -1 Then ''����Ŀ¼������Oracle_Home��Ϣ��Ĭ�϶�ȡ���
                    .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
                End If
                .Sort = "VerSion Desc,Times Desc,Server"
                Do While Not .EOF
                    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle" & !Name, "ORACLE_HOME")
                    If strPath = "" And !Name & "" = "" Then
                        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle", "ORA_CRS_HOME")
                    End If
                    If strPath <> "" Then
                        strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i����
                        If gobjFile.FileExists(strFile) Then Exit Do
                        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
                        If gobjFile.FileExists(strFile) Then Exit Do
                    End If
                    strFile = ""
                    .MoveNext
                Loop
            End If
        End With
    End If
    If strFile = "" Then Exit Function
    
    GetTNSFile = strFile
End Function

Public Function GetAllSubKey(ByVal KeyRoot As Long, KeyName As String) As Variant
'����:��ȡĳ�����������
'���أ�=��������
    Dim lnghKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim strSubKey As Variant
    strSubKey = Array()
    lngIdx = 0: strName = String(256, Chr(0))
    lngRet = RegOpenKey(KeyRoot, KeyName, lnghKey)
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lnghKey, lngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve strSubKey(UBound(strSubKey) + 1)
                strSubKey(UBound(strSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> 0
    End If
    RegCloseKey lnghKey
    GetAllSubKey = strSubKey
End Function

Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'����:ͨ��OracleHome����ȡOracle��Ϣ
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*�汾Home_32Bit
    'Key_Ora*�汾_Home*
    arrTmp = Split(UCase(strOraHome), "_")
    For i = 1 To UBound(arrTmp)
        If arrTmp(i) Like "HOME*" Then
            intTimes = ValEx(arrTmp(2))
            blnRetrun = True
        ElseIf arrTmp(i) Like "*HOME*" Then
            intTimes = Val(Mid(arrTmp(1), InStr(UCase(arrTmp(1)), "HOME") + 4))
            blnRetrun = True
        End If
        If arrTmp(i) Like "ORADB*" Then
            intVer = ValEx(Mid(arrTmp(1), 6))
            intServer = 1
            blnRetrun = True
        ElseIf arrTmp(i) Like "ORACLIENT*" Then
            intVer = ValEx(Mid(arrTmp(1), 10))
            intServer = 2
            blnRetrun = True
        ElseIf arrTmp(i) Like "*CLIENT*" Then
            intServer = 2
            intVer = ValEx(arrTmp(i))
            blnRetrun = True
        End If
    Next
    GetOraInfoByRegKey = blnRetrun
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'���ܣ���ע���
    Dim i As Long                                           ' ѭ��������
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' �����򿪵�ע����ؼ���
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' ע����ؼ�����������
    Dim tmpVal As String                                    ' ע����ؼ��ֵ���ʱ�洢��
    Dim KeyValSize As Long                                  ' ע����ؼ��ֱ����ߴ�
    
    ' �� KeyRoot {HKEY_LOCAL_MACHINE...} �´�ע����ؼ���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע����ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ��������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע����ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' ���/�����ؼ��ֵ�ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' �����ؼ���ֵ��ת������...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ������������...
    Case REG_SZ, REG_EXPAND_SZ                              ' �ַ���ע����ؼ�����������
        sKeyVal = tmpVal                                     ' �����ַ�����ֵ
    Case REG_DWORD                                          ' ���ֽ�ע����ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ�������ֵ��
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' ת�����ֽ�Ϊ�ַ���
    End Select
    
    GetKeyValue = sKeyVal                                   ' ����ֵ
    rc = RegCloseKey(hKey)                                  ' �ر�ע����ؼ���
    Exit Function                                           ' �˳�
    
GetKeyError:    ' ����������������...
    GetKeyValue = vbNullString                              ' ���÷���ֵΪ���ַ���
    rc = RegCloseKey(hKey)                                  ' �ر�ע����ؼ���
End Function

Public Function ValEx(ByVal varInput As Variant) As Variant
'���ܣ�����Valֻ�������ֿ�ͷʶ��ValEx�Ե�һ�����ֽ���ʶ��
    Dim lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
End Function

Public Function ConvertStr(ByVal strSource As String) As String
    '����:ȥ���ַ����Ŀո�\���з�,��ת��Ϊ��д
    
    strSource = UCase(strSource)
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, vbNewLine, "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    strSource = Replace(strSource, vbTab, "")
    strSource = Replace(strSource, vbBack, "")
    ConvertStr = strSource
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function


Public Function OpenFolder(ByVal frmodtvOwner As Form, Optional strTitle As String, Optional ByVal strInitDir As String) As String
'    '----------------------------------------------------------------------------------------------------
'    '����:ѡ���ļ���
'    '����:frmodtvOwner-ѡ���ļ��еĸ�����
'    '       strFolderName-ָ�����ļ���
'    '       strTitle-����
'    '       strInitDir-Ĭ�ϴ�·��
'    '����:strFolderName-����ѡ����ļ���
'    '----------------------------------------------------------------------------------------------------
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    gstrAPIPath = strInitDir & Chr(0)
    With tBrowseInfo
        .hwndOwner = frmodtvOwner.hwnd
        .lpszTitle = lstrcat(strTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        .lpfnCallback = AddressOfFunction(AddressOf OpenDirCallbackProc)
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH * 2)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
       OpenFolder = sBuffer
    End If
End Function

Public Function OpenDirCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
 '���ܣ�OpenFolder�ص��������������ô򿪵��ļ��ĳ�ʼ·��
    Dim ret As Long
    Dim sBuffer As String
  
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hwnd, BFFM_SETSELECTION, 1, ByVal gstrAPIPath)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH * 2)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
            End If
    End Select
    
    OpenDirCallbackProc = 0
End Function

Private Function AddressOfFunction(Address As Long) As Long
'���ܣ�OpenFolder�Ӻ���
    AddressOfFunction = Address
End Function


Public Sub InitTable(vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    Dim arrTmp      As Variant
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            arrTmp = Split(arrHead(i), ",")
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = arrTmp(0)
            .ColKey(.FixedCols + i) = arrTmp(0)
    
            If UBound(arrTmp) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(arrTmp(1))
                .ColAlignment(.FixedCols + i) = Val(arrTmp(2))
                
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
    End With
End Sub

Public Function GetLogPath() As String
    '����:��ȡLog��־��·��
    Dim strlogPath As String
    
    strlogPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "Log"
     If Not gobjFile.FolderExists(strlogPath) Then gobjFile.CreateFolder strlogPath
     strlogPath = strlogPath & "\��־����"
     If Not gobjFile.FolderExists(strlogPath) Then gobjFile.CreateFolder strlogPath
     strlogPath = strlogPath & "\NoticeLog"
     If Not gobjFile.FolderExists(strlogPath) Then gobjFile.CreateFolder strlogPath

    GetLogPath = strlogPath
End Function


Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single = -1, Optional frmParent As Object, Optional blnPer As Boolean)
'���ܣ���ʾ�����صȴ�����ȴ���(strInfo)
'����:strInfo=�ȴ��������ʾ��Ϣ
'     sngPer=����
    Static blnShow As Boolean
    
    If sngPer > 1 Then sngPer = 1
    
    If strInfo = "" Then
        frmFlash.avi.Close
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            If sngPer = -1 Then
                '��ʾ�ȴ�
                frmFlash.avi.Open GetSetting("ZLSOFT", "ע����Ϣ", "gstrAviPath", "") & "\" & "Findfile.avi"
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                frmFlash.lbl.Caption = strInfo
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    Err.Clear
                    frmFlash.Show , frmParent
                    If Err.Number <> 0 Then
                        Err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.avi.Play
                frmFlash.Refresh
            Else
                '��ʾ����
                frmFlash.avi.Visible = False
                frmFlash.picDo.Visible = True
                frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
                frmFlash.lbl.Left = frmFlash.picDo.Left
                frmFlash.lblPer.Top = frmFlash.lbl.Top
                frmFlash.lbl.Caption = strInfo
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If blnPer Then
                    If sngPer > 0 Then
                        frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                    Else
                        frmFlash.lblPer.Caption = ""
                    End If
                    frmFlash.lblPer.Visible = True
                End If
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    Err.Clear
                    frmFlash.Show , frmParent
                    If Err.Number <> 0 Then
                        Err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.Refresh
            End If
            blnShow = True
        Else
            frmFlash.lbl.Caption = strInfo
            If sngPer >= 0 Then
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Sub OnlyIntCK(ByRef KeyAscii As Integer)
'���ܣ���������������
'��TEXTBOX��KEYPRESSʱ����ʹ�ã���KeyAscII��Ϊ�������뼴��

    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Public Sub OnlyStrCK(ByRef KeyAscii As Integer, ParamArray arrChr() As Variant)
'���ܣ������������ֺ���ĸ,��ָ���ַ�
'��Ҫָ�����ַ���KeyAscII ��ͨ��������ʽ���δ���
'֧��ճ�����ƣ���ݼ�KeyAscii�� CRTRL+C = 3 ,CTRL+V  =22
     Dim intIdx As Integer, intFlag As Integer
    
    intFlag = 1
    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) = 0 Then
            intFlag = 0
        End If
    End If
    
    For intIdx = LBound(arrChr) To UBound(arrChr)
        If Chr(KeyAscii) = arrChr(intIdx) Then
            intFlag = 1
        End If
    Next
    
    If intFlag = 0 Then
        KeyAscii = 0
    End If
    
End Sub