Attribute VB_Name = "mdlPDF"
Option Explicit

'marrFoxitSetEx˵��
'Private Enum enmFoxitPDF
'    ȱʡ�����ļ�·�� = 0
'    �����ӡ�������ļ�·��_F7 = 1
'    ȱʡ�����ļ�·��_����������_F7 = 2
'    �����ӡ�������ļ�·��_����������_F7 = 3
'    ȱʡ�����ļ�·��_����������_F7_64b = 4
'End Enum

'ע���ؼ��ָ�����
Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006 '�˸����б���ÿ��ϵͳ����ʱ��������ϵͳ���ú͵�ǰ������Ϣ���������ֻ������Windows 98�С�
End Enum

'ע�����������
Private Enum REGValueType
    REG_SZ = 1 'Unicode���ս��ַ���
    REG_EXPAND_SZ = 2 'Unicode���ս��ַ���
    REG_BINARY = 3 '��������ֵ
    REG_DWORD = 4 '32-bit ����
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' ��������ֵ��
End Enum

Private Type PROCESSENTRY32
    lSize                                     As Long
    lUsage                                    As Long
    lProcessId                                As Long
    lDefaultHeapId                            As Long
    lModuleId                                 As Long
    lThreads                                  As Long
    lParentProcessId                          As Long
    lPriClassBase                             As Long
    lFlags                                    As Long
    sExeFile                                  As String * 1024
End Type


'ע���ѡ��
Private Const REaD_CONTROL = &H20000
Private Const KEY_QUERY_VaLUE = &H1
Private Const KEY_SET_VaLUE = &H2
Private Const KEY_CREaTE_Sub_KEY = &H4
Private Const KEY_ENUMERaTE_Sub_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREaTE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VaLUE + KEY_ENUMERaTE_Sub_KEYS + KEY_NOTIFY + REaD_CONTROL
Private Const KEY_WRITE = KEY_SET_VaLUE + KEY_CREaTE_Sub_KEY + REaD_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VaLUE + KEY_SET_VaLUE + KEY_CREaTE_Sub_KEY + KEY_ENUMERaTE_Sub_KEYS + KEY_NOTIFY + KEY_CREaTE_LINK + REaD_CONTROL

'��������ֵ
Private Const ERROR_SUCCESS = 0
Private Const ERROR_BADKEY = 2
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_MORE_DATA = 234&

Private Const CSIDL_APPDATA As Long = &H1A                          '���û���\Ӧ�ó��������
Private Const TH32CS_SNAPPROCESS As Long = &H2

Private Const MSTR_FOXIT As String = "Foxit Reader PDF Printer"
Private Const MSTR_TINY As String = "TinyPDF"

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function RegQueryValueEx_String Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_Long Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_ValueType Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Private mblnReset As Boolean
Private marrReset() As Byte
Private mstrError As String
Private mblnAllow As Boolean
Private mobjFSO As FileSystemObject
Private mstrFoxitSet As String
Private mstrFoxitCach As String
Private mstrPDFDevice As String
Private mstrFileName As String
Private mstrFoxitCachCur As String
Private mblnFoxit As Boolean
Private mblnFoxit7 As Boolean
Private marrFoxitSetEx(4) As String

'######################################################################################################################
'����
Public Function PDFInitialize(ByVal objFmt As RPTFmt) As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ�����������Ƿ�����������PDF
    '���أ�����True��ʾ���������PDF�ļ���False��ʾ�����������PDF�ļ�
    '******************************************************************************************************************
    Dim strPDFFile As String
    Dim strPath As String * 255
    Dim lngPIDsplwow64 As Long
    Dim objFile As File
    
    On Error GoTo ErrHand
    
    mblnAllow = False
    If mobjFSO Is Nothing Then
        Set mobjFSO = New FileSystemObject
    End If
    
    '���PDF�����ӡ��
    If CheckPDFDevice = False Then Exit Function
                    
    Select Case mstrPDFDevice
    Case MSTR_TINY
        '�޸�ע�����Ϣ
        Call GetTempPath(255, strPath)
        strPDFFile = Trim(Left(strPath, InStr(strPath, Chr(0)) - 1)) & "Demo.pdf"
        If Dir(strPDFFile) <> "" Then
            Kill strPDFFile
            DoEvents
        End If
        'ע�����
        If ModifyRegist(strPDFFile, False, False, False, "", "") = False Then Exit Function
        'ģ�����
        If OutputDemo() = False Then Exit Function
    Case MSTR_FOXIT
        '����Foxit�����ӡ��Ŀ¼
        Call SetFoxitPrinter(mstrFoxitCachCur)
        'ģ������ļ���ͬʱ����Foxit��صı���������Ϣ
        If OutputDemo() = False Then Exit Function
        
        '��û��SPLWOW64.EXE���̣���ӡ֮��Ϳ��ܴ��ڣ���ʱ�ý��������ļ����ȱʡ�ļ����ø���
        If Is64bit And mblnFoxit7 And mblnFoxit And marrFoxitSetEx(4) = "" Then
            '64λĳЩ�����ʹ�øý��̵������ļ�
            lngPIDsplwow64 = GetProcessID("SPLWOW64.EXE")
            If lngPIDsplwow64 <> 0 And mobjFSO.FolderExists(mobjFSO.GetParentFolderName(marrFoxitSetEx(4))) Then
                For Each objFile In mobjFSO.GetFolder(mobjFSO.GetParentFolderName(marrFoxitSetEx(0))).Files
                    If UCase(objFile.name) Like "*_" & lngPIDsplwow64 & "__FOXITTEMP.XML" Then
                        marrFoxitSetEx(4) = objFile.Path
                        Exit For
                    End If
                Next
            End If
        End If
        
        '���ģ������ļ���Ŀ¼
        Call ClearFolder(mstrFoxitCachCur)
    Case Else
        Exit Function
    End Select
    
    '���ô�ӡ������
    With Printer
        On Error Resume Next
        .PaperSize = objFmt.ֽ��
        If Err.Number <> 0 Then
            If Not gblnSilentMode Then
                MsgBox mdlPublic.FormatString("��ӡ����[1]����֧�ָ��Զ���ֽ�ŵĳߴ磡", .DeviceName) _
                    , vbInformation, "ע��"
            End If
        End If
        .Width = objFmt.W
        .Height = objFmt.H
        .Orientation = objFmt.ֽ��
    End With
    
    PDFInitialize = True
    mblnAllow = True
    
    Exit Function
    
ErrHand:
    mstrError = Err.Description
End Function

Public Function PDFFile(ByVal strPDFFile As String, _
                        Optional ByVal blnCopyable As Boolean = False, _
                        Optional ByVal blnEditable As Boolean = False, _
                        Optional ByVal blnPrintable As Boolean = False, _
                        Optional ByVal strPassWord As String = "", _
                        Optional ByVal strAttachs As String = "") As Boolean
    '******************************************************************************************************************
    '���ܣ��������PDF�ļ��Ļ���
    '������strPDFFile=����ļ����������ļ�·�����ļ���չ��
    '                 �ļ�·��������ڣ��һ��Զ�����ͬ���ļ�
    '                 ���δָ�����򵯳��ļ�����Ի���
    '      blnCopyable=�����PDF�ļ��Ƿ�����������
    '      blnEditable=�����PDF�ļ��Ƿ�����༭����
    '      blnPrintable=�����PDF�ļ��Ƿ������ӡ���
    '      strPassword=�Ƿ�Ҫ����������
    '      strAttachs=Ҫ�ӵ�PDF�еĸ����ļ���(����·��),�����"|"�ָ�
    '���أ�
    'ע�⣺�ú�����Ҫ��Printer���κδ�ӡ����֮ǰ����(����API��ʽ����)
    '******************************************************************************************************************
        
    On Error GoTo ErrHand
    
    If mblnAllow = False Then Exit Function
            
    '�޸�ע�����Ϣ
    If strPDFFile = "" Then
        mstrError = "δָ��PDF�ļ����ƣ��������PDF��"
        Exit Function
    End If
    
    mstrFileName = strPDFFile
    
    Select Case mstrPDFDevice
    Case MSTR_FOXIT
        PDFFile = SetFoxitPrinter(mstrFoxitCachCur)
    Case MSTR_TINY
        PDFFile = ModifyRegist(strPDFFile, blnCopyable, blnEditable, blnPrintable, strPassWord, strAttachs)
    End Select
    
    Exit Function
    
ErrHand:
    mstrError = Err.Description
End Function

Public Function GetLastError() As String
    GetLastError = mstrError
End Function

'######################################################################################################################
'˽��
Private Function Is64bit() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '���أ�
    '******************************************************************************************************************
    Dim lngHandle As Long, lngFunc As Long
        
    lngHandle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If lngHandle > 0 Then
        IsWow64Process GetCurrentProcess(), lngFunc
    End If
    Is64bit = lngFunc <> 0
End Function

Private Function CheckPDFDevice() As Boolean
    '******************************************************************************************************************
    '���ܣ����PDF�����ӡ��
    '���أ�True-���������ӡ��False-�����������ӡ��
    '******************************************************************************************************************
    
    Dim intLoop As Integer
    Dim lngRetrun As Long, lngProcessID As Long, lngPos As Long
    Dim strDeviceName As String, strPath As String, strProcessName As String
    Dim strTemp As String * 100
    Dim objFile As File
    Dim objFolder As Folder
    Dim strShortSvrName As String, strProFile As String
    Dim arrTmp() As Byte
    Dim lngPIDsplwow64 As Long
    
    '����Ƿ����TinyPDF(32λϵͳ���ȼ��Foxit Reader PDF Printer���ټ��TinyPDF) Foxit Reader PDF Printer (64λϵͳ)��ӡ��
    strDeviceName = MSTR_FOXIT
    If UCase(Printer.DeviceName) <> UCase(strDeviceName) Then
        For intLoop = 0 To Printers.count - 1
            If Printers(intLoop).DeviceName = strDeviceName Then Set Printer = Printers(intLoop):  Exit For
        Next
        If intLoop >= Printers.count Then
            If Not Is64bit Then
                strDeviceName = MSTR_TINY
                For intLoop = 0 To Printers.count - 1
                    If Printers(intLoop).DeviceName = strDeviceName Then Set Printer = Printers(intLoop):  Exit For
                Next
                If intLoop >= Printers.count Then
                    mstrError = "û�м�⵽��װ��" & MSTR_TINY & "��" & MSTR_FOXIT & "�����ӡ�����������PDF��"
                    Exit Function
                End If
            Else
                mstrError = "û�м�⵽��װ��" & strDeviceName & "�����ӡ�����������PDF��"
                Exit Function
            End If
        End If
    End If
    
    mstrPDFDevice = strDeviceName
    If strDeviceName = MSTR_FOXIT Then
        If Not mobjFSO.FolderExists(mstrFoxitCachCur) Or marrFoxitSetEx(0) = "" Then
            mblnFoxit7 = False
            '��ȡ��ӡ�������ļ�Ŀ¼
            strTemp = String(50, " ")
            lngRetrun = GetSpecialFolderPath(0, strTemp, CSIDL_APPDATA, False)
            strPath = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
            strPath = strPath & "\Foxit Software\Foxit PDF Creator"
            If mobjFSO.FolderExists(strPath) = False Then
                mstrError = "δ�ҵ�Foxit Reader PDF��ӡ������Ŀ¼���ļ������������ָ����Ŀ¼!" & vbCrLf & "����Ŀ¼:" & strPath
                Exit Function
            End If
            
            For intLoop = LBound(marrFoxitSetEx) To UBound(marrFoxitSetEx)
                marrFoxitSetEx(intLoop) = ""
            Next
            
            If mobjFSO.FolderExists(strPath & "\Foxit Reader PDF Printer") Then
                'Foxit 7.0
                mblnFoxit7 = True
                marrFoxitSetEx(0) = strPath & "\Foxit Reader PDF Printer\FoxitPrinterProfile.xml"
                'VB��ѡ����Ϊȱʡ����ʱ�������ļ�
                If GetRegValue("HKEY_CURRENT_USER\Printers\DevModePerUser", "Foxit Reader PDF Printer", arrTmp) Then
                    strProFile = GetFoxitProfile(arrTmp)
                    If mobjFSO.FolderExists(mobjFSO.GetParentFolderName(strProFile)) Then
                        marrFoxitSetEx(0) = mobjFSO.GetParentFolderName(strProFile) & "\FoxitPrinterProfile.xml"
                        marrFoxitSetEx(1) = strProFile
                    ElseIf GetRegValue("HKEY_CURRENT_USER\Printers\DevModes2", "Foxit Reader PDF Printer", arrTmp) Then
                        strProFile = GetFoxitProfile(arrTmp)
                        If mobjFSO.FolderExists(mobjFSO.GetParentFolderName(strProFile)) Then
                            marrFoxitSetEx(0) = mobjFSO.GetParentFolderName(strProFile) & "\FoxitPrinterProfile.xml"
                            marrFoxitSetEx(1) = strProFile
                        End If
                    End If
                End If
                
                '��ȡsplwow64.exe��·��
                If Is64bit Then
                    '64λĳЩ�����ʹ�øý��̵������ļ�
                    lngPIDsplwow64 = GetProcessID("SPLWOW64.EXE")
                    If lngPIDsplwow64 <> 0 Then
                        For Each objFile In mobjFSO.GetFolder(strPath & "\Foxit Reader PDF Printer").Files
                            If UCase(objFile.name) Like "*_" & lngPIDsplwow64 & "__FOXITTEMP.XML" Then
                                marrFoxitSetEx(4) = objFile.Path
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                'VB����δ��ѡʹ�õ������ļ�
                If GetRegValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Printers\Foxit Reader PDF Printer\DsSpooler", "shortServerName", strShortSvrName) Then
                    If GetRegValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Printers\Foxit Reader PDF Printer", "Default DevMode", arrTmp) Then
                        strProFile = GetFoxitProfile(arrTmp)
                        If mobjFSO.FolderExists(mobjFSO.GetParentFolderName(strProFile)) Then
                            '��û�������ļ�ʱ��ȱʡ�����ļ�
                            marrFoxitSetEx(2) = mobjFSO.GetParentFolderName(strProFile) & "\FoxitPrinterProfile.xml"
                            marrFoxitSetEx(3) = strProFile
                        End If
                    End If
                End If
            Else
                marrFoxitSetEx(0) = strPath & "\FoxitReaderPrinterProfile.xml"
            End If
            
            '��ȡ·��
            If InDesign Then
                mstrFoxitCach = "C:\APPSOFT\FoxitPrinter"
            Else
                mstrFoxitCach = Mid(App.Path & "\", 1, InStr(4, App.Path & "\", "\")) & "FoxitPrinter"
            End If
            If Not mobjFSO.FolderExists(mobjFSO.GetParentFolderName(mstrFoxitCach)) Then
                Call mobjFSO.CreateFolder(mobjFSO.GetParentFolderName(mstrFoxitCach))
            End If
            If Not mobjFSO.FolderExists(mstrFoxitCach) Then
                Call mobjFSO.CreateFolder(mstrFoxitCach)
            End If
        
            '�жϽ����Ƿ����,���������ļ���
            For Each objFolder In mobjFSO.GetFolder(mstrFoxitCach).SubFolders
                lngPos = InStr(objFolder.name, "_")
                If lngPos > 0 Then
                    lngProcessID = Val(Mid(objFolder.name, 1, lngPos - 1))
                    strProcessName = Mid(objFolder.name, lngPos + 1)
                    If Not FindProcess(strProcessName, lngProcessID) Then
                        Call ClearFolder(objFolder.Path, True)
                    End If
                Else
                    Call ClearFolder(objFolder.Path, True)
                End If
            Next
            mstrFoxitCachCur = mstrFoxitCach & "\" & GetCurProcessInfo
            'ɾ��ͬ���ļ��������޷������ļ���
            If mobjFSO.FileExists(mstrFoxitCachCur) Then
                Call FileNormal(mstrFoxitCachCur)
                Call mobjFSO.DeleteFile(mstrFoxitCachCur, True)
            End If
            If mobjFSO.FolderExists(mstrFoxitCachCur) Then
                Call ClearFolder(mstrFoxitCachCur)
                If mobjFSO.GetFolder(mstrFoxitCachCur).Files.count <> 0 Then
                    mstrError = "�޷����Foxit Reader PDF��ӡ������Ŀ¼���ļ������������ָ����Ŀ¼!" & vbCrLf & "����Ŀ¼:" & mstrFoxitCachCur
                    Exit Function
                End If
            Else
                Call mobjFSO.CreateFolder(mstrFoxitCachCur)
            End If
        End If
        Call SetFoxitPrinter(mstrFoxitCachCur)
        mblnFoxit = True
    End If
        
    CheckPDFDevice = True
End Function

Private Function ModifyRegist(ByVal strPDFFile As String, Optional ByVal blnCopyable As Boolean, Optional ByVal blnEditable As Boolean, Optional ByVal blnPrintable As Boolean, Optional ByVal strPassWord As String, Optional ByVal strAttachs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�ָ��TinyPDF��ӡ���������
    '������strPDFFile=����ļ����������ļ�·�����ļ���չ��
    '                 �ļ�·��������ڣ��һ��Զ�����ͬ���ļ�
    '                 ���δָ�����򵯳��ļ�����Ի���
    '      blnCopyable=�����PDF�ļ��Ƿ�����������
    '      blnEditable=�����PDF�ļ��Ƿ�����༭����
    '      blnPrintable=�����PDF�ļ��Ƿ������ӡ���
    '      strPassword=�Ƿ�Ҫ����������
    '      strAttachs=Ҫ�ӵ�PDF�еĸ����ļ���(����·��),�����"|"�ָ�
    'ע�⣺�ú�����Ҫ��Printer���κδ�ӡ����֮ǰ����(����API��ʽ����)
    '******************************************************************************************************************
    Dim arrData() As Byte
    Dim intSect As Integer, intAdr As Integer
    Dim intTag As Integer, strFile As String
    Dim i As Integer, j As Integer
    Dim strWord As String
    Dim strRegister(92) As String
    Dim aryRegister As Variant
    Dim intLoop As Integer
            
            
    '��ȡ����
    GetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModePerUser", "TinyPDF", arrData
    
    On Error Resume Next
    Err = 0
    i = UBound(arrData)
    If Err <> 0 Then i = -1
    On Error GoTo ErrHand
    
    If i = -1 Then
        '��ע���

        strRegister(0) = "84,0,105,0,110,0,121,0,80,0,68,0,70,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(1) = "0,0,0,0,0,0,0,0,0,0,0,0,0,1,4,0,4,220,0,236,16,19,78,1,0,1,0,9,0,0,0,0,0,100,0,1,0,15,0,88,2,2,0,1,0,0,0,3,0,0"
        strRegister(2) = "0,65,117,116,111,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(3) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(4) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,63,0,0,0,1,0,0,0,3,0,0,0,44,1,0,0,194,1,0,0,2,80,0,0,3,0,0"
        strRegister(5) = "0,44,1,0,0,194,1,0,0,2,80,0,0,3,0,0,0,176,4,0,0,8,7,0,0,2,0,0,0,0,0,0,0,1,0,0,0,1,0,0,0,100,0,0,0,2,0,0,0,6"
        strRegister(6) = "0,0,0,1,3,0,0,26,1,0,0,44,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(7) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1"
        strRegister(8) = "0,0,0,0,0,0,0,70,1,0,0,72,1,0,0,0,0,0,0,74,1,0,0,76,1,0,0,78,1,0,0,80,1,0,0,82,1,0,0,84,1,0,0,86,1,0,0,88,1,0"
        strRegister(9) = "0,90,1,0,0,0,0,0,0,0,0,65,0,114,0,105,0,97,0,108,0,0,0,65,0,114,0,105,0,97,0,108,0,32,0,78,0,97,0,114,0,114,0,111,0,119,0,0,0,65"
        strRegister(10) = "0,114,0,105,0,97,0,108,0,32,0,85,0,110,0,105,0,99,0,111,0,100,0,101,0,32,0,77,0,83,0,0,0,67,0,101,0,110,0,116,0,117,0,114,0,121,0,32,0,71"
        strRegister(11) = "0,111,0,116,0,104,0,105,0,99,0,0,0,67,0,111,0,117,0,114,0,105,0,101,0,114,0,32,0,78,0,101,0,119,0,0,0,71,0,101,0,111,0,114,0,103,0,105,0,97"
        strRegister(12) = "0,0,0,73,0,109,0,112,0,97,0,99,0,116,0,0,0,76,0,117,0,99,0,105,0,100,0,97,0,32,0,67,0,111,0,110,0,115,0,111,0,108,0,101,0,0,0,84,0,97"
        strRegister(13) = "0,104,0,111,0,109,0,97,0,0,0,84,0,105,0,109,0,101,0,115,0,32,0,78,0,101,0,119,0,32,0,82,0,111,0,109,0,97,0,110,0,0,0,84,0,114,0,101,0,98"
        strRegister(14) = "0,117,0,99,0,104,0,101,0,116,0,32,0,77,0,83,0,0,0,86,0,101,0,114,0,100,0,97,0,110,0,97,0,0,0,0,0,115,82,71,66,32,73,69,67,54,49,57,54,54"
        strRegister(15) = "45,50,46,49,0,85,46,83,46,32,87,101,98,32,67,111,97,116,101,100,32,40,83,87,79,80,41,32,118,50,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(16) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(17) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(18) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(19) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(20) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(21) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(22) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(23) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(24) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(25) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(26) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(27) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(28) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(29) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(30) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(31) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(32) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(33) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(34) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(35) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(36) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(37) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(38) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(39) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(40) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(41) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(42) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(43) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(44) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(45) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(46) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(47) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(48) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(49) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(50) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(51) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(52) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(53) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(54) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(55) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(56) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(57) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(58) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(59) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(60) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(61) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(62) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(63) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(64) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(65) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(66) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(67) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(68) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(69) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(70) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(71) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(72) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(73) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(74) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(75) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(76) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(77) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(78) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(79) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(80) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(81) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(82) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(83) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(84) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(85) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(86) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(87) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(88) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(89) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(90) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(91) = "0"
        
        For i = 0 To 91
            aryRegister = Split(strRegister(i), ",")
            
            For j = 0 To UBound(aryRegister)
                ReDim Preserve arrData(intLoop)
                arrData(intLoop) = Val(aryRegister(j))
                intLoop = intLoop + 1
            Next
            
        Next
    End If
    
    If Not mblnReset Then
        marrReset = arrData
        mblnReset = True
    End If

    '��������
    arrData(Val("&H00E0")) = &H0 'ҳ�߾�
    arrData(Val("&H00E1")) = &H0 'ҳ�߾�
    arrData(Val("&H00E4")) = &H0 '���Զ���
    arrData(Val("&H011C")) = &H1 'Ƕ����������
    arrData(Val("&H0130")) = &H0 'RGB��ɫ(sRGB������)
    
    If strPassWord <> "" Then
        arrData(Val("&H013C")) = &H1 '���û�����
        For i = 1 To Len(strPassWord)
            arrData(Val("&H0140") + i - 1) = Asc(Mid(strPassWord, i, 1))
        Next
        arrData(Val("&H0140") + i - 1) = &H0
    Else
        arrData(Val("&H013C")) = &H0 '���û�����
        arrData(Val("&H0140")) = &H0
    End If
    
    arrData(Val("&H0164")) = &H1 '���а�ȫ����
    arrData(Val("&H0168")) = &H0  '��ȫ��������Ϊ��
    If blnPrintable Then
        arrData(Val("&H0189")) = &H2  '����߷ֱ��ʴ�ӡ
    Else
        arrData(Val("&H0189")) = &H0  '�������ӡ
    End If
    If blnEditable Then
        arrData(Val("&H018A")) = &H4  '����ȡҳ��֮����κ�����
    Else
        arrData(Val("&H018A")) = &H0  '���������
    End If
    If blnCopyable Then
        arrData(Val("&H018C")) = &H1  '����������
    Else
        arrData(Val("&H018C")) = &H0  '��������
    End If
    arrData(Val("&H0190")) = &H1  '��������ʱ��������Ļ�Ķ����豸�Ӿ�����ط����ı�
    If strPDFFile <> "" Then
        arrData(Val("&H0194")) = &H2  'ָ���ļ����(����·��)
    Else
        arrData(Val("&H0194")) = &H0  '��ʾ���
    End If
    arrData(Val("&H01A0")) = &H1  'ֱ�Ӹ����ļ�
    
    '���ݶΣ�����ļ��������ļ�
    arrData(Val("&H01C8")) = &H0
    arrData(Val("&H01C8") + 1) = &H0
    intAdr = Val("&H01CA")
    intSect = 1 '���ݶ����
    intTag = 1 '1-��������,2-�������
    Do While intAdr <= 4552
        If intSect = 1 Or intSect = 2 Then 'Ƕ��/��Ƕ�������
            If arrData(intAdr) = 0 And arrData(intAdr + 1) = 0 Then
                If intTag = 1 Then
                    intTag = 2
                ElseIf intTag = 2 Then
                    intTag = 1
                    intSect = intSect + 1
                End If
            Else
                intTag = 1
            End If
            intAdr = intAdr + 2
        ElseIf intSect = 3 Then '�м�����
            If arrData(intAdr) = 0 Then
                intAdr = intAdr + 1
            Else
                intSect = intSect + 1
            End If
        ElseIf intSect = 4 Or intSect = 5 Then 'RGB/CMYK�����ļ���
            If arrData(intAdr) = 0 Then
                intSect = intSect + 1
            End If
            intAdr = intAdr + 1
        ElseIf intSect = 6 Then '���Ŀ¼��
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H0198")) = Val("&H" & Right(strWord, 2)) '��λ�ֽ�
            arrData(Val("&H0198") + 1) = Val("&H" & Left(strWord, 2)) '��λ�ֽ�
            
            arrData(intAdr) = 0
            arrData(intAdr + 1) = 0
            intAdr = intAdr + 2
            intSect = intSect + 1
        ElseIf intSect = 7 Then '����ļ���
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H019C")) = Val("&H" & Right(strWord, 2)) '��λ�ֽ�
            arrData(Val("&H019C") + 1) = Val("&H" & Left(strWord, 2)) '��λ�ֽ�
            
            If strPDFFile = "" Then
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            Else
                For i = 1 To Len(strPDFFile)
                    strWord = Hex(AscW(Mid(strPDFFile, i, 1)))
                    If Len(strWord) = 2 Then
                        strWord = "00" & strWord
                    End If
                    
                    arrData(intAdr + i * 2 - 2) = Val("&H" & Right(strWord, 2)) '��λUnicode
                    arrData(intAdr + i * 2 - 1) = Val("&H" & Left(strWord, 2)) '��λUnicode
                Next
                intAdr = intAdr + Len(strPDFFile) * 2
                
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            End If
            intSect = intSect + 1
        ElseIf intSect = 8 Then '�м�����
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H01A4")) = Val("&H" & Right(strWord, 2)) '��λ�ֽ�
            arrData(Val("&H01A4") + 1) = Val("&H" & Left(strWord, 2)) '��λ�ֽ�
            
            For i = 1 To 16
                arrData(intAdr + i - 1) = 0
            Next
            intAdr = intAdr + 16
            intSect = intSect + 1
        ElseIf intSect = 9 Then '�����ļ�
            'Ŀǰ�������ü��ظ����ᵼ�����ɵ�PDF�򿪳���
            If strAttachs = "" Then
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            Else
                For i = 0 To UBound(Split(strAttachs, "|"))
                    strFile = Split(strAttachs, "|")(i)
                    For j = 1 To Len(strFile)
                        strWord = Hex(AscW(Mid(strFile, j, 1)))
                        If Len(strWord) = 2 Then
                            strWord = "00" & strWord
                        End If
                        
                        arrData(intAdr + j * 2 - 2) = Val("&H" & Right(strWord, 2)) '��λUnicode
                        arrData(intAdr + j * 2 - 1) = Val("&H" & Left(strWord, 2)) '��λUnicode
                    Next
                    intAdr = intAdr + Len(strFile) * 2
                    
                    arrData(intAdr) = 0
                    arrData(intAdr + 1) = 0
                    intAdr = intAdr + 2
                Next
            End If
            '�����˳�
            Exit Do
        End If
    Loop
    
    For i = Val("&H01A8") To Val("&H01C4") Step 4
        strWord = Hex(arrData(Val("&H01A4")) + arrData(Val("&H01A4") + 1) * 256 + (i - Val("&H01A8")) / 2 + 2)
        strWord = String(4 - Len(strWord), "0") & strWord
        
        arrData(i) = Val("&H" & Right(strWord, 2)) '��λ�ֽ�
        arrData(i + 1) = Val("&H" & Left(strWord, 2)) '��λ�ֽ�
        arrData(i + 2) = 0
        arrData(i + 3) = 0
    Next
    
    '��������
    SetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModePerUser", "TinyPDF", arrData
    SetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModes2", "TinyPDF", arrData
    
    ModifyRegist = True
    
    Exit Function
ErrHand:
    mstrError = Err.Description
End Function

Private Function OutputDemo() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '˵����
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    Call OutPut(Printer)
        
    OutputDemo = True
    Exit Function
ErrHand:
    mstrError = Err.Description
    Printer.EndDoc
End Function

Private Sub OutPut(objOut As Object)
    '------
    objOut.Font.name = "����"
    objOut.Font.Size = 18
    objOut.ForeColor = vbRed
    objOut.CurrentY = 300
    objOut.CurrentX = (objOut.ScaleWidth - objOut.TextWidth("PDF�ļ����ɲ���ʾ��")) / 2
    objOut.Print "PDF�ļ����ɲ���ʾ��"
    
    '------
    objOut.DrawWidth = 2 '�߿��ڴ�ӡ�����������Ǻ�����
    objOut.Line (100, 800)-(objOut.ScaleWidth - 100, 800), vbBlue
    
    '------
    objOut.Font.name = "����"
    objOut.Font.Size = 12
    objOut.ForeColor = vbBlack
    objOut.CurrentX = 300
    objOut.CurrentY = 1000 + 100
    objOut.Print "��ϲ��"
    
    objOut.CurrentX = 300
    objOut.Print "��������Զ�ȡ�����Ϣ����˵���ڱ����Ͽ�������PDF�ļ���"
    objOut.EndDoc
    
End Sub

Private Sub ResetPDF()
    '******************************************************************************************************************
    '���ܣ�����TinyPDF��ӡ�������������
    '˵�����ú����ڴ�ӡ�����ɺ����
    '******************************************************************************************************************
    If mblnReset Then
        SetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModePerUser", "TinyPDF", marrReset
        SetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModes2", "TinyPDF", marrReset
        Erase marrReset
        mblnReset = False
    Else
        DeleteRegValue HKEY_CURRENT_USER, "Printers\DevModePerUser", "TinyPDF"
        DeleteRegValue HKEY_CURRENT_USER, "Printers\DevModes2", "TinyPDF"
    End If
End Sub

'######################################################################################################################
Private Function GetRegValueBinary(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String, arrData() As Byte) As Boolean
    '���ܣ���ȡע�����ָ��λ�õĶ�����ֵ
    Dim lngKey As Long, lngReturn As Long
    Dim lngLength As Long

    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If

    lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, REG_BINARY, ByVal 0, lngLength)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If

    ReDim arrData(lngLength - 1)
    lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, REG_BINARY, arrData(0), lngLength)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If

    RegCloseKey lngKey
    GetRegValueBinary = True
End Function

Private Function SetRegValueBinary(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String, arrData() As Byte) As Boolean
    '******************************************************************************************************************
    '���ܣ�����ע�����ָ��λ�õĶ�����ֵ
    '˵����
    '  1.��ע��������ʱ���Զ�����
    '  2.���ע��������������ͻ��Ϊ����������
    '******************************************************************************************************************
    Dim lngKey As Long, lngReturn As Long

    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If

    lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, REG_BINARY, arrData(0), UBound(arrData) + 1)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If

    RegCloseKey lngKey
    SetRegValueBinary = True
End Function

Private Function DeleteRegValue(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String) As Boolean
    '���ܣ�ɾ��ע�����ָ��λ�õ���Ŀ
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long


    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If

    lngReturn = RegDeleteValue(lngKey, strValueName)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If

    RegCloseKey lngKey
    DeleteRegValue = True
End Function

Private Sub ClearFolder(ByVal strFolder As String, Optional ByVal blnDelFolder As Boolean)
'******************************************************************************************************************
'���ܣ�����ָ���ļ���
'������strFolder-������ļ��У�blnDelFolder-�Ƿ�ɾ�����ļ���
'˵��:
'******************************************************************************************************************
    Dim objFile          As File
    On Error Resume Next
    If mobjFSO.FolderExists(strFolder) Then
        For Each objFile In mobjFSO.GetFolder(strFolder).Files
            Call FileNormal(objFile.Path)
            Call mobjFSO.DeleteFile(objFile.Path, True)
        Next
        If blnDelFolder Then Call mobjFSO.DeleteFolder(strFolder, True)
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function FindProcess(ByVal strProcessName As String, Optional ByVal lngProcID As Long) As Boolean
'******************************************************************************************************************
'���ܣ� �ж�ָ�����ƺͽ���ID�Ľ����Ƿ����
'˵��:
'******************************************************************************************************************
    Dim uProcess As PROCESSENTRY32
    Dim lngSnapShot As Long, lngRet As Long
    Dim strFindName As String, lngPos As Long
    Dim lngPid As Long
    On Error GoTo errH
    FindProcess = False
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot <> 0 Then
        uProcess.lSize = 1060
        If (Process32First(lngSnapShot, uProcess)) Then
            Do
                lngPos = InStr(1, uProcess.sExeFile, Chr(0))
                strFindName = UCase(Left(uProcess.sExeFile, lngPos - 1))
                If strFindName = strProcessName Then
                    lngPid = uProcess.lProcessId
                    If lngProcID = lngPid Then
                        FindProcess = True
                        Exit Do
                    End If
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
        lngRet = CloseHandle(lngSnapShot)
    End If
    Exit Function
errH:
    Err.Clear
End Function

Private Function GetCurProcessInfo() As String
'******************************************************************************************************************
'���ܣ� ��ȡ��ǰ���̵Ľ������ƺͽ���ID
'���أ�����ID_����EXE����
'˵��:
'******************************************************************************************************************

    Dim lngCurProcID        As Long
    Dim uProcess            As PROCESSENTRY32
    Dim lngSnapShot         As Long, lngRet         As Long
    Dim strFindName         As String, lngPos       As Long
    Dim lngPid              As Long
    
    lngCurProcID = GetCurrentProcessId
    
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot <> 0 Then
        uProcess.lSize = 1060
        If (Process32First(lngSnapShot, uProcess)) Then
            Do
                lngPid = uProcess.lProcessId
                If lngCurProcID = lngPid Then
                    lngPos = InStr(1, uProcess.sExeFile, Chr(0))
                    strFindName = UCase(Left(uProcess.sExeFile, lngPos - 1))
                    GetCurProcessInfo = lngCurProcID & "_" & strFindName
                    Exit Do
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
        lngRet = CloseHandle(lngSnapShot)
    End If
End Function

Private Function FileNormal(ByVal strSource As String) As Boolean
'******************************************************************************************************************
'���ܣ� ����ļ���ֻ����־
'˵��:
'******************************************************************************************************************

    On Error Resume Next
    If mobjFSO.FileExists(strSource) Then
        If FileSystem.GetAttr(strSource) <> vbNormal Then
            FileSystem.SetAttr strSource, vbNormal
        End If
    End If
    
    FileNormal = Err.Number = 0
    If Err.Number <> 0 Then Err.Clear
End Function

Private Function SetFoxitPrinter(ByVal strFilePath As String) As Boolean
'���ܣ�����Foxit�����ӡ��Ŀ¼
    Dim objStream As TextStream
    Dim i As Integer
    Dim blnDo As Boolean, blnOK As Boolean
    
    On Error GoTo errH
    
    blnOK = False
    For i = LBound(marrFoxitSetEx) To UBound(marrFoxitSetEx)
        blnDo = False
        If marrFoxitSetEx(i) <> "" Then
            If mobjFSO.FileExists(marrFoxitSetEx(i)) Then
                If SetFoxitFolder(strFilePath, marrFoxitSetEx(i)) Then
                    blnDo = True
                End If
            End If
            If blnDo = False Then
                '����һ���ļ�
                Set objStream = mobjFSO.CreateTextFile(marrFoxitSetEx(i), True)
                If Not mblnFoxit7 Then
                    objStream.WriteLine "<FXCreatorData><General Folder="""" Overwrite=""1"" UseDefFileName=""0"" OpenFile=""0"" ImageCompress=""0"" IgonareBK=""0"" PDFA1B=""0"" PDFVersion=""17"" Gray=""0"" BlackAndWhite=""0"" DPI=""600""/>"
                    objStream.WriteLine "<Layout UOM=""0"" PaperSize=""9"" PaperWidth=""2100"" PaperLength=""2970"" Orientation=""1"" FormName=""A4""/>"
                    objStream.WriteLine "</FXCreatorData>"
                Else
                    objStream.WriteLine "<FXCreatorData><General DefaultFolder="""" Overwrite=""1"" UseDefFileName=""0"" OpenFile=""0"" TemplateName=""Standard"" IgonareBK=""0"" PDFVersion=""17"" ColorSpace=""2"" DPI=""600"" DeleteLogFile=""0""/>"
                    objStream.WriteLine "<Layout UOM=""0"" PaperSize=""9"" PaperWidth=""2100"" PaperLength=""2970"" Orientation=""1"" FormName=""A4""/>"
                    objStream.WriteLine "<DocumentInfo AddDocInfo=""0"" DocTitle="""" DocSubject="""" DocAuthor="""" DocKeyWords="""" DocCreator=""""/>"
                    objStream.WriteLine "</FXCreatorData>"
                End If
                objStream.Close
                Set objStream = Nothing
                If Not SetFoxitFolder(strFilePath, marrFoxitSetEx(i)) Then
                    mstrError = "��ӡ�������ļ���ʽ������Ч��XML�ļ������飡" & vbCrLf & "�ļ�·��:" & marrFoxitSetEx(i)
                    marrFoxitSetEx(i) = ""
                Else
                    blnDo = True
                End If
            End If
            If blnDo Then
                blnOK = True
                mstrError = ""
            ElseIf blnOK Then
                mstrError = ""
            End If
        End If
    Next
    SetFoxitPrinter = blnOK
    Exit Function
    
errH:
    mstrError = "(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function

Private Function SetFoxitFolder(ByVal strFilePath As String, ByVal strFoxitSet As String) As Boolean
    Dim objXML As Object, objNode As Object
    Dim strFolder As String
    
    On Error GoTo errH
    
    Set objXML = CreateObject("MSXML2.DOMDocument")
    If objXML.Load(strFoxitSet) = True Then
        Set objNode = objXML.selectSingleNode("FXCreatorData").selectSingleNode("General")
        If mblnFoxit7 Then
            strFolder = "DefaultFolder"
        Else
            strFolder = "Folder"
        End If
        If objNode.Attributes.getNamedItem(strFolder).Text <> strFilePath Then
            objNode.Attributes.getNamedItem(strFolder).Text = strFilePath
        End If
        
        If objNode.Attributes.getNamedItem("Overwrite").Text <> "1" Then
            objNode.Attributes.getNamedItem("Overwrite").Text = "1"
        End If
        If objNode.Attributes.getNamedItem("UseDefFileName").Text <> "1" Then
            objNode.Attributes.getNamedItem("UseDefFileName").Text = "1"
        End If
        If objNode.Attributes.getNamedItem("OpenFile").Text <> "0" Then
            objNode.Attributes.getNamedItem("OpenFile").Text = "0"
        End If
        Call objXML.Save(strFoxitSet)
        SetFoxitFolder = True
    End If
    Exit Function
    
errH:
    Err.Clear
End Function

Public Function PDFFileSuccess() As Boolean
'******************************************************************************************************************
'���ܣ�PDF����ļ�����
'���أ�True,�ļ����ɳɹ���False-�ļ�����ʧ��
'˵�����ú����ڴ�ӡ�����ɺ����
'******************************************************************************************************************
    Dim strFileName As String
    Dim objFile As File
    
    If mblnAllow = False Then Exit Function
    
    Select Case mstrPDFDevice
    Case MSTR_FOXIT
        For Each objFile In mobjFSO.GetFolder(mstrFoxitCachCur).Files
            strFileName = objFile.Path
            Exit For
        Next
        If strFileName <> "" Then
            '��strFileName�ļ����Ƶ�
            Call mobjFSO.CopyFile(strFileName, mstrFileName, True)
            PDFFileSuccess = mobjFSO.FileExists(mstrFileName)
            Call ClearFolder(mstrFoxitCachCur)
        End If
    Case MSTR_TINY
        PDFFileSuccess = mobjFSO.FileExists(mstrFileName)
    End Select
End Function

Private Function GetFoxitProfile(ByRef arrData As Variant) As String
    Dim i       As Long
    Dim strTmp  As String
    Dim arrTmp()    As Byte
    Dim LngLoop     As Long
    
    On Error GoTo errH
    Err.Clear
    i = UBound(arrData)
    For i = &HDC - 1 To UBound(arrData)
        If arrData(i) <> 0 Then
            ReDim Preserve arrTmp(LngLoop)
            arrTmp(LngLoop) = Val(arrData(i))
            LngLoop = LngLoop + 1
        End If
    Next
    strTmp = StrConv(arrTmp(), vbUnicode)
    strTmp = Replace(strTmp, "\\\", "\")
    GetFoxitProfile = strTmp
    Exit Function
    
errH:
    Err.Clear
End Function

Private Function GetRegValue(ByVal strKey As String, ByVal strValueName As String, ByRef varValue As Variant, Optional blnOneString As Boolean = False) As Boolean
'���ܣ���ȡע�����ָ��λ�õ�ֵ
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'          strValue=����ֵ
'          strValueType=�������ͣ�Ĭ��Ϊ�ַ���
'           blnOneString = ��REG_EXPAND_SZ��REG_MULTI_SZ,REG_BINARY��Ч��-  True �������ص�һ�ַ������Ҳ����κδ���ֻȥ���ַ���β��
'���أ��Ƿ��ȡ�ɹ�
'˵������ǰֻ��REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ��REG_DWORD��REG_BINARYʵ���˶�ȡ��û�в�ѯ�������Զ����Ҽ���
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '������Ч��ע����λ,��ȡ��������
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '�򿪱���
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '�ַ������Ͷ�ȡ
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '���ܳ��������������
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = Left(strBuf, InStr(strBuf, Chr(0)) - 1)
                Case REG_EXPAND_SZ ' ���价���ַ�������ѯ���������ͷ��ض���ֵ
                    If Not blnOneString Then
                        varValue = ExpandEnvStr(Left(strBuf, InStr(strBuf, Chr(0)) - 1))
                    Else
                        varValue = Left(strBuf, InStr(strBuf, Chr(0)) - 1)
                    End If
                Case REG_MULTI_SZ ' �����ַ���
                    If Not blnOneString Then
                        If Len(strBuf) <> 0 Then ' �������Ƿǿ��ַ��������Էָ
                            strBufVar = Split(Left$(strBuf, Len(strBuf) - 1), Chr$(0))
                        Else ' ���ǿ��ַ�����Ҫ����S(0) ���������
                            ReDim strBufVar(0) As String
                        End If
                        ' ��������ֵ������һ���ַ������飿��
                        varValue = strBufVar()
                    Else
                        varValue = Left(strBuf, InStr(strBuf, Chr(0)) - 1)
                    End If
            End Select
        Case REG_DWORD
            lngReturn = RegQueryValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, Len(lngBuf))
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): varValue = 0: Exit Function
            varValue = lngBuf
        Case REG_BINARY
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, ByVal 0, lngLength)
            If lngReturn <> ERROR_SUCCESS And lngReturn <> ERROR_MORE_DATA Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            ReDim bytBuf(lngLength - 1)
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
            If lngReturn <> ERROR_SUCCESS And lngReturn <> ERROR_MORE_DATA Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            If lngLength <> UBound(bytBuf) + 1 Then
               ReDim Preserve bytBuf(0 To lngLength - 1) As Byte
            End If
            ' �����ַ�����ע�⣺Ҫ���ֽ��������ת����
            If blnOneString Then
                'ѭ�����ݣ����ֽ�ת��Ϊ16�����ַ���
                For i = LBound(bytBuf) To UBound(bytBuf)
                   strTmp = CStr(Hex(bytBuf(i)))
                   If (Len(strTmp) = 1) Then strTmp = "0" & strTmp
                   strReturn = strReturn & " " & strTmp
                Next i
                varValue = Trim$(strReturn)
            Else
                varValue = bytBuf()
            End If
    End Select
    RegCloseKey lngKey
    GetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetKeyValueInfo(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef hRootKey As REGRoot, Optional ByRef strSubKey As String, Optional ByRef lngType As Long) As Boolean
'���ܣ����ݼ�λ��ȡ����ֵ���ӽ�,�Լ�ֵ����
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'���Σ�
'          hRootKey=����
'          strSubKey=�ӽ�
'          lngType=������
'���أ��Ƿ��ȡ�ɹ�
    Dim strRoot As String, lngPos As String, hKey As Long
    Dim lngReturn As Long, strName As String * 255
    
    On Error GoTo errH
    hRootKey = 0: strSubKey = "": lngType = 0
    lngPos = InStr(strKey, "\")
    If lngPos = 0 Then Exit Function
    strRoot = Mid(strKey, 1, lngPos - 1)
    strSubKey = Mid(strKey, lngPos + 1)
    
    hRootKey = Decode(UCase(strRoot), _
                    "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
                    "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
                    "HKEY_LOCAL_MACHINE", HKEY_LOCAL_MACHINE, _
                    "HKEY_USERS", HKEY_USERS, _
                    "HKEY_PERFORMANCE_DATA", HKEY_PERFORMANCE_DATA, _
                    "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
                    "HKEY_DYN_DATA", HKEY_DYN_DATA, 0)
    If hRootKey = 0 Then Exit Function
    If lngType <> -1 Then
        'ʹ�ò�ѯ��ʽ�򿪣����м������Ͳ�ѯ
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        If strValueName <> "" Then
            lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
            '�����ֶγ��������Ȳ��������Գ����˳�
            'If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (hKey): Exit Function
        End If
        RegCloseKey (hKey)
    End If
    GetKeyValueInfo = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Private Function ExpandEnvStr(ByVal strInput As String) As String
'���ܣ����ַ����еĻ��������滻Ϊ����ֵ
'         strInput=���������������ַ���
'���أ���ʵ�ʵ�ֵ�滻�ַ����еĻ�����������ַ���
    '// �磺 %PATH% �򷵻� "c:\;c:\windows;"
    Dim lngLen As Long, strBuf As String, strOld As String
    strOld = strInput & "  " ' ��֪ΪʲôҪ�������ַ������򷵻�ֵ������������ַ���
    strBuf = "" '// ��֧��Windows 95
    '// get the length
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, lngLen)
    '// չ���ַ���
    strBuf = String$(lngLen - 1, Chr$(0))
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, LenB(strBuf))
    '// ���ػ�������
    ExpandEnvStr = Left(strBuf, InStr(strBuf, Chr(0)) - 1)
End Function

Private Function Decode(ParamArray arrPar() As Variant) As Variant
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

Private Function GetProcessID(ByVal strExeName As String) As Long
'���ܣ���ȡָ���������Ľ���ID
'���أ�����ID
    
    Dim uProcess            As PROCESSENTRY32
    Dim lngSnapShot         As Long, lngRet         As Long
    Dim strFindName         As String, lngPos       As Long
    Dim lngPid              As Long
    
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot <> 0 Then
        uProcess.lSize = 1060
        If (Process32First(lngSnapShot, uProcess)) Then
            Do
                lngPos = InStr(1, uProcess.sExeFile, Chr(0))
                strFindName = UCase(Left(uProcess.sExeFile, lngPos - 1))
                If strFindName = strExeName Then
                    GetProcessID = uProcess.lProcessId
                    Exit Do
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
        lngRet = CloseHandle(lngSnapShot)
    End If
End Function


