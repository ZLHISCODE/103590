Attribute VB_Name = "mdlPDF"
Option Explicit
'Window�汾����
Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'�µĴ�ӡ��ʽʹ��-----------------------------------------------------------
'ע����dmFields��Long��,as Long��β����&��
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_COLLATE = &H8000&
Public Const DM_FORMNAME = &H10000
'Constants for DocumentProperties() call
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY
Public Const DM_PROMPT = 4
Public Const DM_IN_PROMPT = DM_PROMPT
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
'Constants for DocumentProperties() return
Public Const IDOK = 1
Public Const IDCANCEL = 2
'Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'ֽ�Ŵ�ӡ�߽����================================================================
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'��ͬ��ӡ���Ĵ�ӡ��Ԫ���Ȳ�ͬ
Public Const PHYSICALWIDTH = 110   'Physical Width in device units
Public Const PHYSICALHEIGHT = 111  'Physical Height in device units
Public Const PHYSICALOFFSETX = 112 'Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Public Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Public Const LOGPIXELSY = 90
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version

'ע���ؼ��ָ�����
Public Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
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

'ע���ѡ��
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Private Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
      
'��������ֵ
Private Const ERROR_SUCCESS = 0
Private Const ERROR_BADKEY = 2
Private Const ERROR_ACCESS_DENIED = 8
      
'��ע���
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'��ѯע���ֵ
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'����ע���ֵ
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'ɾ��ע���Ŀ¼/��Ŀ
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'�ر�ע���
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Function SetNTPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional ByVal blnPrompt As Boolean) As Boolean
'���ܣ�NT�����У����ô�ӡ�����Զ���ֽ�ųߴ�
'������lngWidth��lngHeight=mm(����)
'     intOrient=1-����,2-����
'     intCopys=��ӡ����(�����ӡ��֧��,1-9999,��֧��ʱ�������,Ҳ��Ӱ����������)
'˵��������Width,Height�⣬����ͨ�����������õ����Բ�ֱ�ӷ�ӳ��Printer�ϣ�
'      (ȡDevModeҲ��ӳ������������Ҫ��GetJob���ܻ�ȡ����Ĵ�ӡ�ĵ�����)
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte
    Dim lngSize As Long
    
    Dim lngPrtDC As Long
    Dim lngHandle As Long
    Dim strPrtName As String
    
    lngPrtDC = Printer.hDC
    strPrtName = Printer.DeviceName
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then
        'Retrieve the size of the DEVMODE:fMode=0
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        'Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        '���ô�ӡ�ĵ�����
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = intWidth * 10 'in tenths of a millimeter
        vDevMode.dmPaperLength = intHeight * 10 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '�߼���ӡ����(��ȡ��ʱ,Copiesֻ֧��1;����֪��ôȡ����)
        vDevMode.dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERLENGTH Or DM_PAPERWIDTH Or DM_COPIES 'Or DM_COLLATE
        
        'Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        If blnPrompt Then
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_IN_PROMPT Or DM_OUT_BUFFER)
        Else
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        End If
        If lngSize = IDOK Then SetNTPrinterPaper = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetNTPrinterPaper = False
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function IsWindowsNT() As Boolean
'���ܣ��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function
 
Private Function GetWinPlatform() As Long
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function


Public Function GetRegValue(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String, strValue As String) As Boolean
'���ܣ���ȡע�����ָ��λ�õ�ֵ
'������strSubKey=��"Software\WinRAR\Extraction"
'      strValueName=Ϊ��ʱ��ʾȡ"(Ĭ��)"ֵ
    Dim lngKey As Long, lngReturn As Long
    Dim strData As String, lngLength As Long
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_QUERY_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    lngLength = 2048
    strData = Space(lngLength)
    lngReturn = RegQueryValueEx(lngKey, strValueName, 0, REG_SZ, strData, lngLength)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    strValue = Mid(strData, 1, lngLength)
    
    If InStr(strValue, Chr(0)) > 0 Then
        strValue = Mid(strValue, 1, InStr(strValue, Chr(0)) - 1)
    End If
    RegCloseKey lngKey
    GetRegValue = True
End Function

Public Function GetRegValueBinary(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String, arrData() As Byte) As Boolean
'���ܣ���ȡע�����ָ��λ�õĶ�����ֵ
    Dim lngKey As Long, lngReturn As Long
    Dim lngLength As Long
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_QUERY_VALUE, lngKey)
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

Public Function SetRegValue(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String, ByVal strValue As String) As Boolean
'���ܣ�����ע�����ָ��λ�õ�ֵ
'˵����
'  1.��ע��������ʱ���Զ�����
'  2.���ע��������������ͻ��Ϊ�ַ�������

    Dim lngKey As Long
    Dim lngLength As Long, lngReturn As Long
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    lngLength = LenB(StrConv(strValue, vbFromUnicode))
    lngReturn = RegSetValueEx(lngKey, strValueName, 0, REG_SZ, strValue, lngLength)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    SetRegValue = True
End Function

Public Function SetRegValueBinary(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String, arrData() As Byte) As Boolean
'���ܣ�����ע�����ָ��λ�õĶ�����ֵ
'˵����
'  1.��ע��������ʱ���Զ�����
'  2.���ע��������������ͻ��Ϊ����������

    Dim lngKey As Long, lngReturn As Long
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
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

Public Function DeleteRegKey(ByVal hKey As REGRoot, ByVal strSubKey As String) As Boolean
'���ܣ�ɾ��ע�����ָ��λ�õ�Ŀ¼
'˵����
'  1.���������Ŀ¼�����ܼ���ɾ��
'  2.���Ŀ¼������Ŀ������ɾ��
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, "", 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    lngReturn = RegDeleteKey(lngKey, strSubKey)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    DeleteRegKey = True
End Function

Public Function DeleteRegValue(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String) As Boolean
'���ܣ�ɾ��ע�����ָ��λ�õ���Ŀ
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
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

