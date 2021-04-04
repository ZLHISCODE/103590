Attribute VB_Name = "mdlPublic"
Option Explicit

Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'---------------------------------------------------------------
'-ע��� API ����...
'---------------------------------------------------------------
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'---------------------------------------------------------------
'- ע��� Api ����...
'---------------------------------------------------------------
' Reg Data Types...
Public Const REG_SZ = 1                         ' Unicode���ս��ַ���
Public Const REG_EXPAND_SZ = 2                  ' Unicode���ս��ַ���
Public Const REG_DWORD = 4                      ' 32-bit ����
' ע���������ֵ...
Public Const REG_OPTION_NON_VOLATILE = 0       ' ��ϵͳ��������ʱ���ؼ��ֱ�����
' ע���ؼ��ְ�ȫѡ��...
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
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
' ע���ؼ��ָ�����...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
' ����ֵ...
Public Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- ע���ȫ��������...
'---------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'---------------------------------------------------------
'�������
'---------------------------------------------------------
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1


'�ж������Ƿ�Ϊ��
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

'��ȡ�����Ķ��IP
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)


Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'����DICOM������صĽṹ�������
Public Type AEconnection        '��¼������Ϣ��������DICOM�ؼ���DicomConnection������
    Association As Long         '��¼��ǰ���ӵ�id
    ServiceAE As String                '�����е�AE����
    DeviceIP As String                '�豸IP��ַ
    TimeStamp As String         'ʱ�������¼���ӽ�����ʱ��
    Deleted As Boolean          'ɾ����ǣ��Ƿ�ɾ��
End Type
Public AEconnections() As AEconnection  '�洢������Ϣ������

Public Type Service
    DeviceIP As String          '��¼�豸��IP��ַ
    DeviceAE As String          '��¼�豸��AE����
    DevicePort As String        '��¼�豸�Ķ˿�
    DeviceName As String        '��¼�豸����
    ServiceAE As String         '��¼PACS�����AE����
    ServicePort As String       '��¼PACS����Ķ˿ں�
    SOP As String               '��¼������
    Modality As String          '��¼�豸��Ӱ�����
    Started  As Boolean         '��¼��ǰ�����Ƿ�ɹ�����
End Type
Public Services() As Service    '�洢Ӧ���ڵ�ǰIP��ַ��DICOM�����

Public Type AEPara              '��¼��������ļ򵥲���
    AE As String                '��¼�����е�AE����
    IP As String                '��¼�豸IP��ַ
    ParaName As String          '��������
    ParaValue As String         '����ֵ
End Type
Public AEParas() As AEPara      '�洢Ӧ���ڵ�ǰIP��ַ�Ĳ���


Public Type FTPDevice           '��¼FTP�洢�豸
    No As String                '�洢�豸��
    IP As String                'IP��ַ
    User As String              '�û���
    Password As String          '����
    FTPDir As String            'FTPĿ¼
End Type
Public FTPDevices() As FTPDevice        '�洢Ӧ���ڵ�ǰIP��FTP�洢�豸

Public gstrLocalIP As String             '�洢����IP��ַ
Private iNet As New clsFtp      '��Ϊ����������Ŀ���ǣ��Ժ��޸ĳ�FTP�豸�Ų��ı��ʱ�򣬲�������FTP

Public Const ATTR_������� As String = "8:20"
Public Const ATTR_���ʱ�� As String = "8:30"
Public Const ATTR_�ɼ����� As String = "8:22"
Public Const ATTR_�ɼ�ʱ�� As String = "8:32"
Public Const ATTR_ͼ������ As String = "8:23"
Public Const ATTR_ͼ��ʱ�� As String = "8:33"


Public Const ATTR_Ӱ����� As String = "8:60"
Public Const ATTR_����豸 As String = "8:1090"
Public Const ATTR_����� As String = "28:34"
Public Const ATTR_���к� As String = "20:11"
Public Const ATTR_ͼ��� As String = "20:13"
Public Const ATTR_ͼ������ As String = "8:8"

Public Const ATTR_��� As String = "18:50"
Public Const ATTR_ͼ��λ�ò��� As String = "20:32"
Public Const ATTR_ͼ������ As String = "20:37"
Public Const ATTR_�ο�֡UID As String = "20:52"
Public Const ATTR_��Ƭλ�� As String = "20:1041"
Public Const ATTR_���� As String = "28:10"
Public Const ATTR_���� As String = "28:11"
Public Const ATTR_���ؾ��� As String = "28:30"

Public Const TS_JPEG����ѹ�� As String = "1.2.840.10008.1.2.4.70"
Public Const TS_RLE�г�ѹ�� As String = "1.2.840.10008.1.2.5"
Public Const TS_JPEG2000����ѹ�� As String = "1.2.840.10008.1.2.4.90"

Public gcnAccess As New ADODB.connection, strBeginDate As String
Public gstrAccessPath As String         'Access���ݿ���ļ�·�����ļ���ǰ׺���ޡ�.mdb��
Public gstrAccessName As String         'Access���ݿ���ļ�·�����ļ���

Public gstrSQL As String


Public Function funGetFTPDevice(strDeviceNO As String, strIP As String, strUser As String, strPsw As String, strFTPDir As String) As Boolean
    Dim i As Integer
    
    For i = 1 To UBound(FTPDevices)
        If FTPDevices(i).No = strDeviceNO Then
            strIP = FTPDevices(i).IP
            strUser = FTPDevices(i).User
            strPsw = FTPDevices(i).Password
            strFTPDir = FTPDevices(i).FTPDir
            Exit For
        End If
    Next i
    If i <= UBound(FTPDevices) Then
        funGetFTPDevice = True
    Else
        funGetFTPDevice = False
    End If
End Function

Public Function funGetQRParas(strServiceAE As String, strDeviceIP As String, blnCGet As Boolean, _
    intPatientIDMatch As Integer)
    Dim i As Integer
    
    '��ȡ��������
    intPatientIDMatch = 0
    blnCGet = False
    
    For i = 1 To UBound(AEParas)
        If UCase(AEParas(i).AE) = UCase(strServiceAE) And AEParas(i).IP = strDeviceIP Then
            Select Case AEParas(i).ParaName
            Case ZLPACS_QR����CGET
                blnCGet = AEParas(i).ParaValue
            Case ZLPACS_QR����IDƥ��
                intPatientIDMatch = AEParas(i).ParaValue
            End Select
        End If
    Next i
    funGetQRParas = True
End Function

Public Function funGetAEMWLParas(strServiceAE As String, strDeviceIP As String, intFilterModality As Integer, _
        intDayInterval As Integer, blnUseForceResult As Boolean, intBodypartType As Integer, _
        strBodypartSplitter As String, intResultFilter As Integer) As Boolean
    Dim i As Integer
    
    '��ʼ������
    intDayInterval = 3
    intFilterModality = 0
    intBodypartType = 0
    strBodypartSplitter = ""
    intResultFilter = 0
    
    '��ȡ��������
    If SafeArrayGetDim(AEParas) <> 0 Then
        '��¼������־
        Call WriteProcessLog("funGetAEMWLParas", "��ȡWorklist�Ĳ���", "UBound(AEParas) = " & UBound(AEParas) & " ,UCase(strServiceAE) = " & UCase(strServiceAE) & " , strDeviceIP = " & strDeviceIP)
    
        For i = 1 To UBound(AEParas)
            If UCase(AEParas(i).AE) = UCase(strServiceAE) And AEParas(i).IP = strDeviceIP Then
                Select Case AEParas(i).ParaName
                Case ZLPACS_MWL���˷�ʽ
                    intFilterModality = Val(AEParas(i).ParaValue)
                Case ZLPACS_MWL��������
                    intDayInterval = Val(AEParas(i).ParaValue)
                Case ZLPACS_MWL��ǿ�ƽ��
                    blnUseForceResult = AEParas(i).ParaValue
                Case ZLPACS_MWL�ಿλ��ʽ
                    intBodypartType = Val(AEParas(i).ParaValue)
                Case ZLPACS_MWL�ಿλ�ָ���
                    strBodypartSplitter = AEParas(i).ParaValue
                Case ZLPACS_MWL��ѯ��������
                    intResultFilter = Val(AEParas(i).ParaValue)
                End Select
            End If
        Next i
    End If
    
    funGetAEMWLParas = True
End Function
    
Private Function GetAEconnection(ByVal Association As Long, ByRef strServiceAE As String, ByRef strDeviceIP As String) As Boolean
    
    Dim i As Integer
    '���ҷ���AE��IP
    For i = 1 To UBound(AEconnections)
        If AEconnections(i).Association = Association Then
            strServiceAE = AEconnections(i).ServiceAE
            strDeviceIP = AEconnections(i).DeviceIP
            Exit For
        End If
    Next i
    
    If i <= UBound(AEconnections) Then
        GetAEconnection = True
    Else
        GetAEconnection = False
    End If
End Function

Private Function GetFilmStor(ByVal iService As Long, ByRef strServiceAE As String, ByRef strDeviceIP As String) As Boolean
    
    On Error GoTo err
    strServiceAE = Services(iService).ServiceAE
    strDeviceIP = Services(iService).DeviceIP
    
    GetFilmStor = True
    Exit Function
err:
    GetFilmStor = False
End Function


Public Function funGetAEStoreParas(ByVal Association As String, ByVal Modality As String, ByRef strIPAddress As String, ByRef blnSplitSeriesUID As Boolean, ByRef intImageMatchItem As Integer, _
    ByRef intDBMatchItem As Integer, ByRef blnMatchStudyUID As Boolean, ByRef strStoreDeviceNo As String, ByRef intEncode As Integer, _
    ByRef strAutoRoute As String, ByRef intFilterModality As Integer, ByRef strAutoRouteCompression, ByRef strAutoRouteDir, ByRef strPhysician) As Boolean
    
'    '�����������
    Dim i As Integer
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strServiceAE As String      'PACS�����AE����
    Dim strDeviceIP As String       '�豸IP��ַ
    Dim blnRet As Boolean
    
    blnRet = GetAEconnection(Val(Association), strServiceAE, strDeviceIP)
    
    '�Ҳ�����Ӧ�ķ���AE����¼����ʧ�ܣ�Ȼ���Services ���ҵ�һ��Ӱ�������ͬ���豸����ȡ����豸�Ĳ���
    If blnRet = False Then
        WriteLog 41, vbObjectError + 1, "ͨ��Association���Ҳ�����Ӧ�ķ���AE,Association = " & Association & vbCrLf _
                & " UBound(AEconnections) = " & UBound(AEconnections) & " Ӱ����� =" & Modality
                
        For i = 1 To UBound(Services)
            If UCase(Services(i).Modality) = UCase(Modality) And Services(i).Started = True Then
                strServiceAE = Services(i).ServiceAE
                strDeviceIP = Services(i).DeviceIP
                WriteLog 42, vbObjectError + 1, "����Ӱ�������ҵ���ͼ���Ӧ�ķ���AE���豸IP��ServiceAE = " & strServiceAE & vbCrLf _
                    & " DeviceIP = " & strDeviceIP
                Exit For
            End If
        Next i
        If strServiceAE = "" Or strDeviceIP = "" Then
            WriteLog 43, vbObjectError + 1, "�����Ҳ�����ͼ���Ӧ�ķ���AE��ͼ���޷����档"
            funGetAEStoreParas = False
            Exit Function
        End If
    End If
    
    '�����豸IP��ַ
    strIPAddress = strDeviceIP
    
    '��ʼ������
    blnSplitSeriesUID = False
    blnMatchStudyUID = True
    strStoreDeviceNo = ""
    intEncode = 0
    intImageMatchItem = 0
    intDBMatchItem = 0
    strAutoRoute = ""
    strAutoRouteCompression = ""
    strAutoRouteDir = ""
    intFilterModality = 0
    strPhysician = ""
    
    '��ȡ��������
    If SafeArrayGetDim(AEParas) <> 0 Then
        For i = 1 To UBound(AEParas)
            If UCase(AEParas(i).AE) = UCase(strServiceAE) And AEParas(i).IP = strDeviceIP Then
                Select Case AEParas(i).ParaName
                Case ZLPACS_��ͼ�����Ͳ������
                    blnSplitSeriesUID = AEParas(i).ParaValue
                Case ZLPACS_�洢�豸��
                    strStoreDeviceNo = AEParas(i).ParaValue
                Case ZLPACS_���ü��UIDƥ��
                    blnMatchStudyUID = AEParas(i).ParaValue
                Case ZLPACS_ѹ����ʽ
                    If AEParas(i).ParaValue = "JPEG����ѹ��" Then
                        intEncode = 0
                    ElseIf AEParas(i).ParaValue = "RLEѹ��" Then
                        intEncode = 1
                    ElseIf AEParas(i).ParaValue = "JPEG2000ѹ��" Then
                        intEncode = 2
                    Else    '��ѹ��
                        intEncode = 3
                    End If
                Case ZLPACS_���ݿ�ƥ����
                    intDBMatchItem = Val(AEParas(i).ParaValue)
                Case ZLPACS_ͼ��ƥ����
                    intImageMatchItem = Val(AEParas(i).ParaValue)
                Case ZLPACS_�Զ�·��
                    strAutoRoute = AEParas(i).ParaValue
                Case ZLPACS_�Զ�·��ѹ����ʽ
                    strAutoRouteCompression = AEParas(i).ParaValue
                Case ZLPACS_�Զ�·��Ŀ¼�ṹ
                    strAutoRouteDir = AEParas(i).ParaValue
                Case ZLPACS_�洢���˷�ʽ
                    intFilterModality = Val(AEParas(i).ParaValue)
                Case ZLPACS_��ȡ��鼼ʦ
                    If InStr(AEParas(i).ParaValue, ":") > 0 Then
                        strPhysician = AEParas(i).ParaValue
                    End If
                End Select
            End If
        Next i
    End If
    
    '���û�ж���洢�豸�ţ���ʹ�����ݿ��е�һ���洢�豸
    If strStoreDeviceNo = "" Then
        strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡAE�����������", CLng(1))
        
        If rsTmp.EOF Then
            WriteLog 4, vbObjectError + 1, "δ����Ӱ��洢�豸���뵽Ӱ���豸Ŀ¼�����ã�"
            funGetAEStoreParas = False
            Exit Function
        Else
            strStoreDeviceNo = rsTmp(0)
        End If
    End If
    
    funGetAEStoreParas = True
End Function

Private Function funGetStudyUID(ByVal strOldStudyUID As String) As String
'-----------------------------------------------------------------------------
'����:��ѯ���ݿ⣬�жϵ�ǰͼ��ļ��UID�Ƿ��Ѿ����������������ʱ���У�
'     ������ڣ����ڼ��UID�������Ӻ�׺����������ֱ�ӷ�������ļ��UID
'�޸���:�ƽ�
'�޸�����:2007-1-27
'-----------------------------------------------------------------------------
    Dim rsMatch As New ADODB.Recordset
    
    funGetStudyUID = strOldStudyUID
    gstrSQL = "select ���UID from Ӱ�����¼ where ���UID = [1]" & _
              " Union All Select ���UID from Ӱ����ʱ��¼ where ���UID = [1]"
    Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strOldStudyUID)
    If Not rsMatch.EOF Then
        '����һ���µļ��UID
        gstrSQL = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�")
        If Len(strOldStudyUID) <= 55 Then
            funGetStudyUID = strOldStudyUID & ".A" & rsMatch(0)
        Else
            funGetStudyUID = Left(strOldStudyUID, 55) & ".A" & rsMatch(0)
        End If
    End If
End Function

Public Function WriteToURL(ByRef ftpNet As clsFtp, ByVal SrcFileName As String, ByVal DestFileName As String) As Long
'���ܣ��������ļ����浽Զ��������
    Dim objFileSystem As New Scripting.FileSystemObject
    
    WriteToURL = 0  '��ȷ
    
    '����Զ��Ŀ·
    WriteToURL = ftpNet.FuncFtpMkDir("/", objFileSystem.GetParentFolderName(DestFileName))
    
    Call WriteProcessLog("WriteToURL", " ����Զ��Ŀ¼", "����Զ��Ŀ¼�����Ϊ�� " & IIf(WriteToURL = 1, "ʧ��", "�ɹ�") & "��Զ��Ŀ¼Ϊ�� " & objFileSystem.GetParentFolderName(DestFileName))
    
    'Ŀ¼�����ɹ����ϴ�ͼ��
    If WriteToURL = 1 Then Exit Function
    WriteToURL = ftpNet.FuncUploadFile(objFileSystem.GetParentFolderName(DestFileName), SrcFileName, objFileSystem.GetFileName(DestFileName))
    
    Call WriteProcessLog("WriteToURL", "�ϴ�ͼ��", "�ϴ�ͼ����Ϊ��" & IIf(WriteToURL = 0, "�ɹ�", IIf(WriteToURL = 1, "����ʧ��", "�ϴ�ʧ��")) & "�� �ļ���Ϊ��" & SrcFileName)
    
End Function

Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As String
'-----------------------------------------------------------------------------
'����:��ȡDICOM���Լ��е�ָ������ֵ,����VM�ж�ֵ��ά�ȣ�ʹ�á�\���Ѹ���ά�����ӳ�һ����
'������ objAttr ----���Լ���
'       AttrName ----Ҫ���ҵ���������
'����ֵ�����Ե�����
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    Dim i As Integer
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).VM = 1 Then
            GetImageAttribute = NVL(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).ValueByIndex(1))
        Else
            For i = 1 To objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).VM
                GetImageAttribute = GetImageAttribute & "\" & objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).ValueByIndex(i)
            Next i
        End If
    End If
End Function

Public Sub DeleteImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String)
'-----------------------------------------------------------------------------
'����:ɾ��DICOM���Լ��е�ָ������ֵ
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        Call objAttr.Remove("&h" & AttrTag(0), "&h" & AttrTag(1))
    End If
End Sub

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
'���ܣ�����DicomViewer��������
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub

Public Function ImageExist(Images As DicomImages, SeekImage As DicomImage) As Boolean
    Dim curImage As DicomImage
    
    ImageExist = False
    For Each curImage In Images
        If curImage.instanceUID = SeekImage.instanceUID Then ImageExist = True: Exit For
    Next
End Function

Private Sub WriteRecord(ByVal ImageType As String, ByVal strCheckNo As String, ByVal CheckDev As String, _
    ByVal PatientName As String, ByVal EnglishName As String, ByVal Sex As String, Age As Integer, _
    ByVal CheckUID As String, ByVal SeriesUID As String, ByVal ifTmp As Boolean)
'-----------------------------------------------------------------------------
'����:����Ӱ��������У����浽����Access�����ݿ��ļ���
'������ ImageType ----Ӱ�����
'       strCheckNo ----ͼ���е�ƥ��ID��������PatientID��PatientName��AccessionNumber
'       CheckDev ----����豸
'       PatientName ----����
'       EnglishName ----Ӣ����
'       Sex ----�Ա�
'       Age ----����
'       CheckUID ----���UID
'       SeriesUID ----����UID
'       ifTmp ----�Ƿ���ʱ��¼
'����ֵ��ֱ�Ӳ��롰Ӱ��������С���
'-----------------------------------------------------------------------------
    
    Dim rsTmp As ADODB.Recordset, strSQL As String
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Select id from Ӱ��������� Where ����UID='" & SeriesUID & "' And ����ʱ��>cDate('" & _
        strBeginDate & "')"
    Set rsTmp = gcnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        strSQL = "Insert Into Ӱ���������(Ӱ�����,����,����豸,����,Ӣ����,�Ա�,����,Ӱ����,����UID,���UID,��Ӧ���,����ʱ��)" & _
            " Values('" & ImageType & "'," & IIf(strCheckNo = "0", "Null", "'" & strCheckNo & "'") & ",'" & CheckDev & "','" & _
            funDeleteSpcialChar(PatientName) & "','" & funDeleteSpcialChar(EnglishName) & "','" & Sex & "'," & IIf(Age = -1, "Null", Age) & ",1,'" & _
            SeriesUID & "','" & CheckUID & "'," & CStr(Not ifTmp) & ",cDate('" & _
            Date & " " & Time() & "'))"
    Else
        strSQL = "Update Ӱ��������� Set Ӱ����=Ӱ����+1 Where ����UID='" & SeriesUID & "' And ����ʱ��>cDate('" & _
        strBeginDate & "')"
    End If
    gcnAccess.Execute strSQL
End Sub

Public Sub WriteLog(ByVal ErrorType As Integer, ErrorNum As Long, ErrorDesc As String)
'-----------------------------------------------------------------------------
'����:��д������־
'������ ErrorType ----�������ʹ��룬����ͼ�����100��WORKLIST��QR����200��FTP����300,funSplitSeriesUID����1001
'       ErrorNum ----�����
'       ErrorDesc ----��������
'����ֵ����
'-----------------------------------------------------------------------------
    Dim strSQL As String
    On Error Resume Next
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Insert Into ������־(����ʱ��,��������,�����,������Ϣ) " & _
        "Values(cDate('" & Date & " " & Time() & "')," & ErrorType & "," & ErrorNum & ",'" & Replace(ErrorDesc, "'", "''") & "')"
    
    gcnAccess.Execute strSQL
End Sub

Private Function funcAutoRouting(img As DicomImage, BufferDir As String, dtReceived As String, _
    strStudyUID As String, iEncode As Integer, strAutoRoute As String, strAutoRouteCompression As String, _
    strAutoRouteDir As String) As Long
'-----------------------------------------------------------------------------
'����:�Զ�·�ɣ���ͼ���͵�ָ���ĵط�
'������ img ----��Ҫ���͵�ͼ��
'       BufferDir---���ػ���·��
'       dtReceived---�������ڣ���Ϊͼ��·����һ����
'       strStudyUID---���UID����Ϊͼ��·����һ���֣������ֹ�������ͼ��·����һ����ͼ���еļ��UID��������Ҫ���ⲿ����
'       iEncode---ѹ����ʽ
'       strAutoRoute---·��Ŀ�ĵؼ��ϣ�ʹ�á�|���ָ������洢�豸��
'       strAutoRouteCompression---�Զ�·�ɵ�ѹ���������ϣ�ʹ�á�|���ָ�����ѹ����ʽ��0--���յ�ǰ��ʽѹ����1--��ѹ��
'       strAutoRouteDir---�Զ�·�ɵ�Ŀ¼�ṹ���ϣ�ʹ�á�|���ָ�����Ŀ¼�ṹ��0--��鼶��Ŀ¼��Ĭ�ϣ���1--���м���Ŀ¼��3D��
'����ֵ����
'-----------------------------------------------------------------------------
    Dim i As Integer            '����ѭ���ı���
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strDirURL As String         'FTP������Ŀ¼
    Dim strHost As String, strUser As String, strPwd As String
    Dim strRouteDest() As String    '��¼�Զ�·��Ŀ�ĵص��豸��
    Dim strRouteCompression() As String     '��¼�Զ�·�ɵ�ѹ����ʽ
    Dim strRouteDir() As String     '��¼�Զ�·�ɵ�Ŀ¼�ṹ
    Dim thisNet As New clsFtp       'FTP����
    Dim intCurRouteCompression As Integer
    Dim intCurRouteDir As Integer
    Dim strUploadDir As String      '���浽FTP�е�Ŀ¼����
    
    If strAutoRoute = "" Then Exit Function
    
    On Error GoTo ProcError
    
    '��ȡ�Զ�·�ɹ���
    strRouteDest = Split(strAutoRoute, "|")
    strRouteCompression = Split(strAutoRouteCompression, "|")
    strRouteDir = Split(strAutoRouteDir, "|")
    '����Զ�·�ɵ��豸�����Ͳ���������һ�£����¼������־��Ϊ����
    If UBound(strRouteDest) <> UBound(strRouteCompression) Or UBound(strRouteDest) <> UBound(strRouteDir) Then
        Call WriteLog(201, 100, "ͼ��ļ��UIDΪ " & strStudyUID & " ���Զ�·�ɵ��豸�����Ͳ���������һ�£����ܵ����Զ�·���޷���ȷ��ɣ��뵽��Ӱ���豸Ŀ¼���н������á�")
    End If
    
    '�Աȴ洢���򣬲�ƥ�����˳�
    For i = 0 To UBound(strRouteDest)
        '�����ݿ��в��Ҷ�Ӧ�Ĵ洢�豸IP��ַ���û�������
        strSQL = "Select IP��ַ,FTP�û���,FTP����,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As DirUrl From Ӱ���豸Ŀ¼ Where �豸��=  [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PACSͼ�񱣴�", strRouteDest(i))
        If rsTmp.EOF Then
            err.Raise vbObjectError + 1, "PACSͼ�񱣴�", "�Զ�·�� �豸�� " & strRouteDest(i) & " ���ô���"
        End If
        
        strHost = rsTmp!IP��ַ
        strUser = rsTmp!FTP�û���
        strPwd = NVL(rsTmp!FTP����)
        strDirURL = rsTmp!DirUrl
        
        '��ȡ�Զ�·�ɲ���
        intCurRouteCompression = 0
        intCurRouteDir = 0
        On Error Resume Next
        intCurRouteCompression = Val(strRouteCompression(i))
        intCurRouteDir = Val(strRouteDir(i))
        
        On Error GoTo ProcError
        '����ͼ��ָ��URL
        If intCurRouteCompression = 1 Then  '��ѹ��
            img.WriteFile BufferDir & img.instanceUID, True
        Else
            Select Case iEncode
                Case 0
                    img.WriteFile BufferDir & img.instanceUID, True, TS_JPEG����ѹ��
                Case 1
                    img.WriteFile BufferDir & img.instanceUID, True, TS_RLE�г�ѹ��
                Case 2
                    img.WriteFile BufferDir & img.instanceUID, True, TS_JPEG2000����ѹ��
                Case 3
                    img.WriteFile BufferDir & img.instanceUID, True
            End Select
        End If
        
        '��ʼFtp����,FTP ���ӳɹ������ϴ�ͼ��
        thisNet.FuncFtpDisConnect
        If thisNet.FuncFtpConnect(strHost, strUser, strPwd) <> 0 Then
            '����Ŀ¼�ɹ������ϴ�ͼ��
            If intCurRouteDir = 1 Then      '���м����Ŀ¼��3D��
                strUploadDir = strDirURL & dtReceived & "/" & strStudyUID & "/" & img.SeriesUID
            Else            '��鼶���Ŀ¼��Ĭ�ϣ�
                strUploadDir = strDirURL & dtReceived & "/" & strStudyUID
            End If
            If thisNet.FuncFtpMkDir("/", strUploadDir) <> 1 Then
                Call thisNet.FuncUploadFile(strUploadDir, BufferDir & img.instanceUID, img.instanceUID)
            End If
        End If
        Kill BufferDir & img.instanceUID
    Next
    
    thisNet.FuncFtpDisConnect
    Exit Function
ProcError:
    Call WriteLog(2, err.Number, err.Description)
    thisNet.FuncFtpDisConnect
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''WorkList���ֳ���''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub AddResultItem(DataSet As DicomDataSet, request As DicomDataSet, group As Long, element As Long, value As Variant)
    'ֻ������Ҫ����Ŀ
    If request.Attributes(group, element).Exists Then
        If IsNull(value) Then
            value = ""
        End If
        DataSet.Attributes.Add group, element, value
    End If
End Sub

Public Sub AddLinkedDateTimeCondition(ByRef query As String, datecondition As DicomAttribute, timecondition As DicomAttribute, dbname As String)
    Dim startdatetime As Date, enddatetime As Date
    If datecondition.Exists And timecondition.Exists Then
        startdatetime = datecondition.DateTimeFrom("1/1/1800") + timecondition.DateTimeFrom("0")
        enddatetime = datecondition.DateTimeTo("1/1/2999") + timecondition.DateTimeTo("0.9999")
        AddSingleDateCondition query, startdatetime, ">=", dbname
        AddSingleDateCondition query, enddatetime, "<=", dbname
        
    Else
        AddDateCondition query, datecondition, "DateValue(" & dbname & ")"
        AddDateCondition query, timecondition, "TimeValue(" & dbname & ")"
    End If
End Sub

Private Sub AddSingleDateCondition(ByRef query As String, Condition As Date, operator As String, dbname As String)
    ' all date formating goes through here to make it easy to change for different databases or locales
    query = query & " AND " & dbname & operator & "to_Date('" & Condition & "', 'yyyy-mm-dd hh24:mi:ss')"
End Sub

Public Sub AddDateCondition(ByRef query As String, Condition As DicomAttribute, dbname As String)
'-----------------------------------------------------------------------------
'����:������ڲ�ѯ����,��Ϊ*���������
'������ query ----��ѯ��SQL
'       Condition ---- �������ڵ����ݼ�
'       dbname ----���ݿ��ֶ�����
'����ֵ���ޣ�ֱ���޸�query
'-----------------------------------------------------------------------------
    If Condition.Exists Then
        If Condition.value <> "" And Condition.value <> "*" Then
            AddSingleDateCondition query, Condition.DateTimeFrom("1/1/1900") & " " & "00:00:01", ">=", dbname
            AddSingleDateCondition query, Condition.DateTimeTo("1/1/2900") & " " & "23:59:59", "<=", dbname
        End If
    End If
End Sub

Public Sub AddDateTimeCondition(ByRef query As String, ConditionDate As DicomAttribute, ConditionTime As DicomAttribute, dbname As String)
'-----------------------------------------------------------------------------
'����:�������ʱ���ѯ����,��Ϊ*�����������ֻ��ʱ��û�����ڵĲ����
'������ query ----��ѯ��SQL
'       ConditionDate ---- �����������ڵ����ݼ�
'       ConditionTime ---- ʱ���������ڵ����ݼ�
'       dbname ----���ݿ��ֶ�����
'����ֵ���ޣ�ֱ���޸�query
'-----------------------------------------------------------------------------
    '
    If ConditionDate.Exists And ConditionTime.Exists Then
        If ConditionTime.value <> "" And ConditionTime.value <> "*" Then
            If ConditionDate.value <> "" And ConditionDate.value <> "*" Then
                '����ʱ����������
                AddSingleDateCondition query, ConditionDate.DateTimeFrom("1/1/1900") & " " & ConditionTime.DateTimeFrom("0:0:1"), ">=", dbname
                AddSingleDateCondition query, ConditionDate.DateTimeTo("1/1/2900") & " " & ConditionTime.DateTimeTo("23:59:59"), "<=", dbname
            Else
                '��ʱ�䣬û��������������ѯ���ˣ�������
            End If
        Else
            'ʱ������Ϊ�գ���ֻ������������
            AddDateCondition query, ConditionDate, dbname
        End If
    End If
End Sub

Public Sub AddIDCondition(ByRef query As String, Condition As DicomAttribute, dbID As String, dbSendNum As String, Optional ByVal blnAndConnect As Boolean = True)
'���ID��ѯ����,��Ϊ*���������
'����˵����
'           query---��ѯ��������
'           Condition ----����Ĳ�ѯ����DicomAttribute�����ֵ�и�ʽ�ǡ�����1_����2��,��ʹ��dbID��dbSendNum����������֯���
'           dbID ----���ݿ��е�ID�ֶ���
'           dbSendNum ----���ݿ��еķ��ͺ��ֶ���
'           blnAndConnect ----��And ����Or������True---And������False---Or����

    Dim strAdviceID As String, strSendNum As String
    Dim strID As String
    If Condition.Exists And Not IsNull(Condition.value) Then
        strID = Condition.value
        If strID <> "*" Then        '��Ϊ*���������
            strAdviceID = Split(strID, "_")(0)
            AddStringCondition query, Val(strAdviceID), dbID, blnAndConnect
            If InStr(strID, "_") > 0 And Len(Trim(dbSendNum)) > 0 Then
                strSendNum = Split(strID, "_")(1)
                AddStringCondition query, Val(strSendNum), dbSendNum, blnAndConnect
            End If
        End If
    End If
End Sub

Public Sub AddCondition(ByRef query As String, Condition As DicomAttribute, dbname As String, Optional ByVal blnAndConnect As Boolean = True)
'����ۺϲ�ѯ����,��Ϊ*���������

    Dim values As Variant
    Dim i As Integer
    
    '�ж������Ƿ�����Ҳ�Ϊ��
    If Condition.Exists And Not IsNull(Condition.value) Then
        If Condition.Multiple Then
            query = query & IIf(blnAndConnect = True, " AND ", " OR ") & " (FALSE "
            values = Condition.value
            For i = 1 To UBound(values, 1)
                If values(i) <> "*" Then
                    query = query & "OR " & dbname & "='" & values(i) & "'"
                End If
            Next
            query = query & ")"
        Else
            AddStringCondition query, Condition.value, dbname, blnAndConnect
        End If
    End If
End Sub

Public Sub AddStringCondition(ByRef query As String, Condition As String, dbname As String, Optional ByVal blnAndConnect As Boolean = True)
'����ַ�����ѯ����,��Ϊ*���������
    If Condition <> "" And Condition <> "*" And Condition <> "*=*" Then
        If InStr(Condition, "*") Then
            query = query & IIf(blnAndConnect, " AND (", " OR (") & dbname & " like '" & StarToPercent(Condition) & "')"
        Else
            query = query & IIf(blnAndConnect, " AND (", " OR (") & dbname & "= '" & Condition & "')"
        End If
    End If
End Sub

Private Function StarToPercent(s As String) As String
    Dim z As Integer
    While InStr(s, "*")
       z = InStr(s, "*")
       s = Left(s, z - 1) & "%" & Mid(s, z + 1)
    Wend
    StarToPercent = s
End Function

Public Function NewResultItem(request As DicomDataSet) As DicomDataSet
    Dim d As DicomDataSet, a As DicomAttribute
    Set d = New DicomDataSet
    For Each a In request.Attributes
        d.Attributes.Add a.group, a.element, a.value
    Next
    Set NewResultItem = d
End Function

Public Sub AddCountItem(DataSet As DicomDataSet, request As DicomDataSet, group As Long, element As Long, _
                SourceName As String, SourceValue As String, TargetName As String)
'-----------------------------------------------------------------------------
'����:  ���ݴ�������󣬲�ѯ��Ӧ�������������������ͼ����������Query/Retrieve��ʹ�ã�
'       ���ֲ�ѯ���ٶȺ����������ܲ�ʹ��,����ֻʹ���˲�ѯͼ�������Ĳ���
'������ DataSet ----���ص����ݼ�
'       request ----Ҫ���ҵ���������
'       group ----Ҫ���ҵ���������
'       element ----Ҫ���ҵ������Ԫ�غ�
'       SourceName ----���ҵ�Դ���𣬰�����PATIENTID��StudyUID��SERIESUID����ʵ��������ֵ����Ӧ��������
'       SourceValue ----���ҵ�����ֵ
'       TargetName ----Ҫ���ص����ݼ��𣬰�����STUDYUID��SERIESUID��INSTANCEUID
'����ֵ���ޣ�ֱ����DataSet��д���ص�����
'-----------------------------------------------------------------------------
    Dim rsTemp As Recordset
    Dim strSQL As String
    
    '���������û�������Ŀ���򲻽��в�ѯ��ֱ���˳�
    If Not request.Attributes(group, element).Exists Then Exit Sub
    
    If UCase(SourceName) = "PATIENTID" And UCase(TargetName) = "STUDYUID" Then
        strSQL = "select count(*) as count from " _
                & "(select c.���� from Ӱ�����¼ c , " _
                & "(select a.����id,b.ҽ��id,b.���ͺ� from ����ҽ����¼ a,����ҽ������ b " _
                & "where a.����id=[1] AND A.���ID IS NULL and a.id=b.ҽ��id) d " _
                & "where c.ҽ��id = d.ҽ��id and c.���ͺ� = d.���ͺ�)"
    ElseIf UCase(SourceName) = "PATIENTID" And UCase(TargetName) = "SERIESUID" Then
        strSQL = "select count(*) as count from " _
                & "(select e.����uid from Ӱ�����¼ c , Ӱ�������� e , " _
                & "(select a.����id,b.ҽ��id,b.���ͺ� from ����ҽ����¼ a,����ҽ������ b " _
                & "where a.����id=[1] AND A.���ID IS NULL and a.id=b.ҽ��id) d " _
                & "where c.ҽ��id = d.ҽ��id and c.���ͺ� = d.���ͺ� and c.���uid = e.���uid)"
    ElseIf UCase(SourceName) = "PATIENTID" And UCase(TargetName) = "INSTANCEUID" Then
        strSQL = " select count(*) as count from " _
                & "(select f.ͼ��uid from Ӱ�����¼ c , Ӱ�������� e , Ӱ����ͼ�� f , " _
                & "(select a.����id,b.ҽ��id,b.���ͺ� from ����ҽ����¼ a,����ҽ������ b " _
                & "where a.����id=[1] AND A.���ID IS NULL and a.id=b.ҽ��id) d " _
                & "Where c.ҽ��id = d.ҽ��id And c.���ͺ� = d.���ͺ� " _
                & "and c.���uid = e.���uid and e.����uid = f.����uid) "
    ElseIf UCase(SourceName) = "STUDYUID" And UCase(TargetName) = "SERIESUID" Then
        strSQL = " select count(*) as count from " _
                & "(select b.����uid from Ӱ�����¼ a , Ӱ�������� b " _
                & "where a.���uid = [1] and a.���uid = b.���uid) "
    ElseIf UCase(SourceName) = "STUDYUID" And UCase(TargetName) = "INSTANCEUID" Then
        strSQL = " select count(*) as count from " _
                & "(select d.ͼ��uid from Ӱ����ͼ�� d , " _
                & "(select b.����uid from Ӱ�����¼ a , Ӱ�������� b " _
                & "where a.���uid =[1] and a.���uid = b.���uid) c " _
                & "where d.����uid = c.����uid)"
    ElseIf UCase(SourceName) = "SERIESUID" And UCase(TargetName) = "INSTANCEUID" Then
        strSQL = "select count(*) as count from " _
                & "(select b.ͼ��uid from Ӱ�������� a , Ӱ����ͼ�� b " _
                & "where a.����uid = [1] and a.����uid = b.����uid)"
    End If
    If UCase(SourceName) = "PATIENTID" Then
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ؼ�¼������", CLng(SourceValue))
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ؼ�¼������", SourceValue)
    End If
    DataSet.Attributes.Add group, element, rsTemp!Count
End Sub

Private Function funcGetSeriesUID(strOldSeriesUID As String, strImageType As String) As String
'-----------------------------------------------------------------------------
'����:������������UID��ѯ��������Ӱ�����Ͳ�ֺ��������UID
'�޸���:�ƽ�
'�޸�����:2007-4-18
'-----------------------------------------------------------------------------
    Dim rsMatch As New ADODB.Recordset
    Dim intMax As Integer
    Dim intCur As Integer
    Dim blnMatch As Boolean
    
    funcGetSeriesUID = strOldSeriesUID
    gstrSQL = "select 0 as ��ʱ,����UID,�������� from Ӱ�������� where ����UID like  [1]" & _
              " Union All Select 1 as ��ʱ,����UID,�������� from Ӱ����ʱ���� where ����UID like [1]"
    Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strOldSeriesUID & "%")
    
    While Not rsMatch.EOF
        If rsMatch("����UID") = strOldSeriesUID Then
            intCur = 0
        Else
            intCur = Val(Right(rsMatch("����UID"), Len(rsMatch("����UID")) - InStrRev(rsMatch("����UID"), ".")))
        End If
        
        If intMax < intCur Then intMax = intCur
        If rsMatch("��������") = strImageType Then
            funcGetSeriesUID = rsMatch("����UID")
            blnMatch = True
            rsMatch.MoveLast
        End If
        rsMatch.MoveNext
    Wend
    
    If blnMatch = False Then
        '�����µ�UID
        funcGetSeriesUID = strOldSeriesUID & "." & intMax + 1
    End If
End Function


Public Sub SaveImages(Images As DicomImages, ByVal BufferDir As String)
'���ܣ�����ͼ��
    Dim curImage As DicomImage
    Dim i As Integer, iCount As Integer     '�����ͼ����
    Dim rsTmp As New ADODB.Recordset
    Dim blnTmp As Boolean                   '�Ƿ񱻱������ʱ��¼
    Dim dtReceived As String
    
    Dim PatientName As String, EnglishName As String, Sex As String, Age As Integer
    
    Dim strBirth As String
    Dim strAdviceID As String   'ͼ���е�ƥ��ID��ID���ܳ����������String��������PatientID��PatientName��AccessionNumber��ͳ��Ϊҽ��ID
    
    Dim lngSeriesNo As Long
    Dim lngImageNo As Long
    Dim strStudyDateTime As String  '�洢�����ݿ��еġ�����ʱ�䡱��ͼ���еļ�����ں�ʱ��
    Dim strStudyUID As String       '�洢���α���ͼ��ʱʹ�õļ��UID
    Dim strSeriesUID As String      '�洢���α���ͼ��ʱʹ�õ�����UID
    
    Dim strSeriesDesp As String     '��������
    Dim strSQLbak As String
    '�����������
    Dim blnSplitSeriesUID As Boolean    '����ͼ�����Ͳ������UID
    Dim intImageMatchItem As Integer    'ͼ��ƥ���� ��0--PatientID��1--AccessionNumber��2--PatientName
    Dim intDBMatchItem As Integer       '���ݿ�ƥ��� 0--����ƥ��;1--���˱�ʶƥ��;2--����ʶƥ��
    Dim blnMatchStudyUID As Boolean     '���ü��UIDƥ��
    Dim strStoreDeviceNo As String      '�洢�豸��
    Dim intEncode As Integer            'ѹ����ʽ
    Dim strOldStoreDeviceNo As String   '������һ��ͼ���FTP�豸��
    Dim strAutoRoute As String          '�����Զ�·��Ŀ�ĵؼ��ϣ�ʹ�á�|���ָ������洢�豸��
    Dim strAutoRouteCompression As String '�����Զ�·�ɵ�ѹ���������ϣ�ʹ�á�|���ָ�����ѹ����ʽ��0--���յ�ǰ��ʽѹ����1--��ѹ��
    Dim strAutoRouteDir As String       '�����Զ�·�ɵ�Ŀ¼�ṹ���ϣ�ʹ�á�|���ָ�����Ŀ¼�ṹ��0--��鼶��Ŀ¼��Ĭ�ϣ���1--���м���Ŀ¼��3D��
    Dim intFilterModality As Integer    '���˷�ʽ 0--��Ӱ�������ˣ�1--��IP��ַ����
    Dim strPhysician As String          '��ȡ��鼼ʦ
    Dim strPhysicianName As String      'ͼ���м�鼼ʦ�ֶε�����
    'FTP�洢����
    Dim strFTPDir As String
    '��ʱʹ�õ�FTP�洢����
    Dim strNewDeviceID As String
        
    'AE���Ӳ���
    Dim strServiceAE As String
    Dim strDeviceIP As String
    
    Dim lngResult As Long           '����FTP�������صĴ���
    Dim blnNewStudy As Boolean      '��¼�Ƿ��µļ��
    
    Dim blnInDBTrans As Boolean     '��¼�Ƿ������ݿ�����֮��
    Dim arrSQL() As Variant         '��¼��Ҫִ�еĴ洢���̵�����
    Dim strModality As String       '��¼ͼ���Ӱ�����
    Dim str����豸 As String       '��¼ͼ���еļ���豸�����ƥ��ɹ����������ݿ��еļ���豸�ֶε�����
    
    On Error GoTo DBError
    
    iCount = 0
    For Each curImage In Images
        '�ȼ�����ͼ���Ƿ��Ѿ��������ݿ�����
        gstrSQL = "Select ͼ��UID From Ӱ����ͼ�� Where ͼ��UID= [1] " & _
            " Union All Select ͼ��UID From Ӱ����ʱͼ�� Where ͼ��UID= [1] "
        strSQLbak = gstrSQL
        strSQLbak = Replace(strSQLbak, "Ӱ����ͼ��", "HӰ����ͼ��")
        gstrSQL = gstrSQL & " Union ALL " & strSQLbak
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS����ͼ��", curImage.instanceUID)
        
        Call WriteProcessLog("SaveImages", "���ͼ���Ƿ��Ѿ��������ݿ���", "������ǣ����ݿ���" & IIf(rsTmp.EOF = True, "û��", "��") & "���ͼ��")
        
        '����ͼ���򱣴�ͼ��,��������һ��ͼ��
        If rsTmp.EOF Then
            '��¼ԭ���Ĵ洢�豸�ţ����Ҷ�ȡ��ǰͼ���Ӧ�Ĵ洢����
            strOldStoreDeviceNo = strStoreDeviceNo
            strModality = GetImageAttribute(curImage.Attributes, ATTR_Ӱ�����)
            str����豸 = GetImageAttribute(curImage.Attributes, ATTR_����豸)
            
            Call WriteProcessLog("SaveImages", "��ȡ��ǰͼ��Ĵ洢��������", "Ӱ����� = " & strModality & "������豸 = " & str����豸)
            
            '��ȡ��ǰͼ��Ĵ洢�����������Զ�ƥ�����
            If funGetAEStoreParas(curImage.Tag, strModality, strDeviceIP, blnSplitSeriesUID, intImageMatchItem, intDBMatchItem, blnMatchStudyUID, _
                strStoreDeviceNo, intEncode, strAutoRoute, intFilterModality, strAutoRouteCompression, strAutoRouteDir, strPhysician) = True Then
                
                'ȷ������ͼ�񣬶����ж�Ӧ���մ洢������׼������ͼ�����ȴ���ͼ���ļ���Ȼ�󱣴��FTP�ļ�
                '����Ӱ������
                DeleteImageAttribute curImage.Attributes, ATTR_����� 'ɾ��������
                '��ȡͼ����Ϣ,���ȶ�ȡ�ɼ�ʱ������ڣ������ڲ����ڣ��ٶ�ȡ�������
                'ԭ���ǿ´���豸��ͨ��Worklist��ȡ��Ϣʱ����ѡ�����ʱ�䡱��д��������ڣ�
                '�����ǰ�������ᵼ�¼�����ڸ��ɼ��������죬Ӱ�����ʱ���ѯ
                dtReceived = Format(GetImageAttribute(curImage.Attributes, ATTR_�ɼ�����), "yyyyMMdd")  '����ͼ���еĲɼ����ڸ�dtReceived����ֵ
                strStudyDateTime = Format(GetImageAttribute(curImage.Attributes, ATTR_�ɼ�����), "yyyy-MM-dd") & _
                    " " & Format(GetImageAttribute(curImage.Attributes, ATTR_�ɼ�ʱ��), "HH:MM")
                If dtReceived = "" Then
                    '��ȡ���ʱ�������
                    dtReceived = Format(GetImageAttribute(curImage.Attributes, ATTR_�������), "yyyyMMdd")  '����ͼ���еļ�����ڸ�dtReceived����ֵ
                    strStudyDateTime = Format(GetImageAttribute(curImage.Attributes, ATTR_�������), "yyyy-MM-dd") & _
                        " " & Format(GetImageAttribute(curImage.Attributes, ATTR_���ʱ��), "HH:MM")
                End If
                
                strStudyUID = curImage.StudyUID             '����ͼ���ڵļ��UID��strStudyUID����ֵ
                PatientName = Replace(curImage.Name, "'", "��")
                EnglishName = Replace(curImage.Name, "'", "��")
                Sex = Replace(curImage.Sex, "'", "��")
                
                '����Ƕ�֡ͼ���򴴽��µ�����UID
                strSeriesUID = curImage.SeriesUID
                If curImage.FrameCount > 1 Then
                    strSeriesUID = funcGetSeriesUID(strSeriesUID, "MultiFrame")
                End If
                strSeriesDesp = Replace(curImage.SeriesDescription, "'", "��")
                '��ȡͼ���е���ƥ��ID
                strAdviceID = funGetMatchIDInImg(curImage, intImageMatchItem, intDBMatchItem)
                '����ͼ�����Ͳ������UID
                If blnSplitSeriesUID = True Then
                    If funSplitSeriesUID(curImage, strSeriesUID, strSeriesDesp) <> 0 Then
                        err.Raise vbObjectError + 4, "funSplitSeriesUID", "�������Ͳ������UID����,���ִ����ͼ���ǣ�" & curImage.Name
                    End If
                End If
                
                '�жϵ�ǰͼ��洢�豸���Ƿ�ı䣬����ı䣬��������ȡFTP�洢�豸��������������FTP
                If strStoreDeviceNo <> strOldStoreDeviceNo Then
                    '��������FTP,lngResult=0�ɹ���lngResult=1����ʧ�ܣ�lngResult=2�޷���ȡ�û���������
                    lngResult = funReConnectFTP(strStoreDeviceNo, iNet, strFTPDir, 1)
                    If lngResult = 1 Then
                        err.Raise vbObjectError + 1, "PACSͼ�񱣴�", "FTP ����ʧ�ܣ�"
                    ElseIf lngResult = 2 Then
                        err.Raise vbObjectError + 1, "PACSͼ�񱣴�", "FTP �޷���ȡFTPĿ¼���û��������룡"
                    End If
                End If
                
                '��ѯ�Ƿ����Ѿ�ƥ��ɹ��ļ�¼
                lngResult = funIsPreMatched(blnMatchStudyUID, intDBMatchItem, strStudyUID, strAdviceID, strDeviceIP, _
                                 strSeriesUID, strModality, dtReceived, intFilterModality, strNewDeviceID, strStoreDeviceNo, _
                                 blnTmp, str����豸, PatientName, EnglishName, Age, Sex, strStudyDateTime)
                If lngResult = 0 Then   'ƥ��ɹ�
                    blnNewStudy = False 'ƥ��ɹ��������µļ��
                    '����豸�Ÿı䣬����������FTP
                    If strNewDeviceID <> strStoreDeviceNo Then
                        strStoreDeviceNo = strNewDeviceID
                        lngResult = funReConnectFTP(strStoreDeviceNo, iNet, strFTPDir, 2)
                        If lngResult = 1 Then
                            err.Raise vbObjectError + 1, "PACSͼ�񱣴�", "FTP ����ʧ�ܣ�"
                        ElseIf lngResult = 2 Then
                            err.Raise vbObjectError + 1, "PACSͼ�񱣴�", "FTP �޷���ȡFTPĿ¼���û��������룡"
                        End If
                    End If
                Else    'ƥ�䲻�ɹ�
                    If blnMatchStudyUID = False Then  '��ѯ���UID�Ƿ��ظ������ظ��򴴽��µļ��UID
                        strStudyUID = funGetStudyUID(strStudyUID)
                    End If
                    blnNewStudy = True
                End If
                
                '����FTPͼ���ļ�������Ŀ¼
                lngResult = funUploadImage(curImage, iNet, intEncode, BufferDir, strFTPDir, strStudyUID, dtReceived)
                If lngResult = 1 Then
                    err.Raise vbObjectError + 2, "PACSͼ�񱣴�", "FTP ��" & Val(curImage.BorderWidth) & "�δ洢ʧ�ܣ�" _
                        & " ����������" & curImage.Name & " ͼ��UID �� " & curImage.instanceUID _
                        & " ����豸�� " & str����豸
                ElseIf lngResult = 2 Then
                    err.Raise vbObjectError + 3, "PACSͼ�񱣴�", "ͼ�񱻷�����FTP ��" & Val(curImage.BorderWidth) & "�δ洢ʧ�ܣ�" _
                        & " ����������" & curImage.Name & " ͼ��UID �� " & curImage.instanceUID _
                        & " ����豸�� " & str����豸 & " ͼ�񸱱�����Ϊ��" & BufferDir & "ͼ��UID.bak"
                ElseIf lngResult = 3 Then
                    err.Raise vbObjectError, "�ϴ�����", "funUploadImage �ϴ�ͼ����ִ���"
                End If
                
                '׼����ʼ��֯����ͼ��Ĵ洢��������
                arrSQL = Array()
                
                '���û��Ԥ��ƥ��ɹ��ļ�¼����˵�����ͼ����ĳ�����ĵ�һ��ͼ�񣬲��������鲢����ƥ��
                '������Ҳ��������飬��˵��ƥ�䲻�ɹ���ͼ��ᱻ�������ʱ�����һ����¼
                If blnNewStudy = True Then      'û���Ѿ�ƥ��ɹ��ļ�¼���򰴲���ID��Ӣ��������
                    Select Case intDBMatchItem
                        Case 0 '����ƥ��
                            gstrSQL = "Select Distinct A.����,A.Ӣ����,A.�Ա�,A.����,A.����豸,A.ҽ��ID,A.���ͺ�,B.�״�ʱ��,abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.�״�ʱ��) as tInterval,b.ִ�й��� " & _
                                " From Ӱ�����¼ A,����ҽ������ B,Ӱ���豸Ŀ¼ C " & _
                                " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And A.����豸 =C.�豸�� And b.ִ��״̬=3 And b.ִ�й���>=2 " & _
                                " And " & IIf(intFilterModality = 0, " UPPER(C.Ӱ�����)=[3] ", " C.IP��ַ=[2] ") & " And A.����= [1] And A.���UID Is Null Order By tInterval"
                        Case 1 '���˱�ʶƥ��
                            gstrSQL = "Select Distinct A.����,A.Ӣ����,A.�Ա�,A.����,A.����豸,A.ҽ��ID,A.���ͺ�,B.�״�ʱ��,abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.�״�ʱ��) as tInterval,b.ִ�й��� " & _
                                " From Ӱ�����¼ A,����ҽ������ B,����ҽ����¼ C,������Ϣ D,Ӱ���豸Ŀ¼ E " & _
                                " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And A.����豸 =E.�豸�� And C.���ID IS NULL And A.ҽ��ID=C.ID And C.����ID=D.����ID" & _
                                " And " & IIf(intFilterModality = 0, " UPPER(E.Ӱ�����)=[3] ", " E.IP��ַ=[2] ") & " And b.ִ��״̬=3 And b.ִ�й���>=2 " & _
                                " And ((D.סԺ��=[1] AND C.������Դ=2) OR (D.�����= [1] AND C.������Դ<>2)) And A.���UID Is Null Order By tInterval"
                        Case 2 '����ʶƥ��
                            gstrSQL = "Select Distinct A.����,A.Ӣ����,A.�Ա�,A.����,A.����豸,A.ҽ��ID,A.���ͺ�,B.�״�ʱ��,abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.�״�ʱ��) as tInterval,b.ִ�й��� " & _
                                " From Ӱ�����¼ A,����ҽ������ B,����ҽ����¼ C" & _
                                " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And B.ҽ��ID=C.ID And C.���ID IS NULL And b.ִ��״̬=3 And b.ִ�й���>=2 " & _
                                " And A.ҽ��ID= [1] And A.���UID Is Null Order By tInterval"
                    End Select
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strAdviceID, strDeviceIP, UCase(strModality))
                        
                    '���ҵ���ƥ��ļ�¼������HIS��д�ļ���¼��Ӧ
                    If rsTmp.EOF = False Then
                        '��¼��ǰ�ļ���豸
                        str����豸 = NVL(rsTmp("����豸"))
                        PatientName = NVL(rsTmp("����"))
                        EnglishName = NVL(rsTmp("Ӣ����"))
                        Age = Val(NVL(rsTmp("����"), 0))
                        Sex = NVL(rsTmp("�Ա�"))
                        
                        '����ƥ���¼
                        gstrSQL = "ZL_Ӱ�����¼_SET(" & rsTmp("ҽ��ID") & "," & rsTmp("���ͺ�") & ",'" & _
                            strStudyUID & "',null," & _
                            "to_Date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS'),'" & strStoreDeviceNo & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = gstrSQL
                        
                        '��Ϊִ�����
                        '���жϵ�ǰ"����ҽ������"�е�"ִ�й���"�Ƿ�С��3,�����,����Ҫ�޸�ִ�й���
                        If rsTmp!ִ�й��� < 3 Then
                            gstrSQL = "ZL_Ӱ����_STATE(" & rsTmp("ҽ��ID") & "," & rsTmp("���ͺ�") & ",3)"
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = gstrSQL
                        End If
                        
                        '��¼��鼼ʦ
                        If strPhysician <> "" Then
                            strPhysicianName = GetImageAttribute(curImage.Attributes, strPhysician)
                            If strPhysicianName <> "" Then
                                gstrSQL = "Zl_Ӱ��ʦִ��('" & strPhysicianName & "'," & rsTmp("ҽ��ID") & ",2)"
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = gstrSQL
                            End If
                        End If
                        blnTmp = False
                    Else        'û���ҵ�ƥ��ļ�¼���������ʱ����¼
                        '������������
                        If IsDate(curImage.DateOfBirthAsDate) Then
                            If curImage.DateOfBirthAsDate <> "0:00:00" Then
                                strBirth = Format(curImage.DateOfBirthAsDate, "YYYY-MM-DD")
                            Else
                                strBirth = ""
                            End If
                            
                            If curImage.Attributes(&H10, &H1010).Exists And Not IsNull(curImage.Attributes(&H10, &H1010)) Then
                                Age = Val(curImage.Attributes(&H10, &H1010).value)
                            Else
                                If strBirth = "" Then
                                    Age = 0
                                Else
                                    Age = CStr(Year(Date) - Year(strBirth))
                                End If
                            End If
                        Else
                            Age = 0: strBirth = ""
                        End If
                        '���������Ҫ�ֶ�
                        PatientName = Replace(curImage.Name, "'", "��")
                        EnglishName = Replace(curImage.Name, "'", "��")
                        Sex = Replace(curImage.Sex, "'", "��")
                        
                        '�������ƥ���ͼ���еĺ���Ϊ-1 �� ���ͼ���PatientID�����¶�ȡһ����롣
                        If strAdviceID = "-1" Then
                            '����ǰ��ռ���ƥ�䣬ֱ����ȡpatientID���ǲ��˱�ʶ��ҽ��ID����ȡ��ֵ
                            strAdviceID = IIf(intDBMatchItem = 0, curImage.PatientID, getStrNumber(curImage.PatientID))
                        End If
                        gstrSQL = "ZL_Ӱ����ʱ���_INSERT('" & strModality & "','" & strAdviceID & "','" & _
                            PatientName & "','" & EnglishName & "','" & Sex & "','" & Age & "'," & _
                            IIf(Len(strBirth) = 0, "Null", "to_Date('" & strBirth & "','YYYY-MM-DD')") & ",Null,Null,'" & _
                            str����豸 & "','" & strStudyUID & "'," & _
                            "to_Date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS'),'" & strStoreDeviceNo & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = gstrSQL
                        blnTmp = True
                    End If
                End If
                
                '�ж��Ƿ���Ҫ�����µ�����
                gstrSQL = "Select ����UID From " & IIf(blnTmp, "Ӱ����ʱ����", "Ӱ��������") & _
                    " Where ����UID= [1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strSeriesUID)
                
                If rsTmp.EOF Then
                    '�����µļ������
                    lngSeriesNo = IIf(GetImageAttribute(curImage.Attributes, ATTR_���к�) = "", -1, GetImageAttribute(curImage.Attributes, ATTR_���к�))
                    If lngSeriesNo <> -1 Then
                        gstrSQL = "select ���к� from " & IIf(blnTmp, "Ӱ����ʱ����", "Ӱ��������") & _
                            " where ���UID=[1] AND ���к� =[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strStudyUID, lngSeriesNo)
                        
                        If Not rsTmp.EOF Then
                            gstrSQL = "select max(���к�) from " & IIf(blnTmp, "Ӱ����ʱ����", "Ӱ��������") & _
                                " where ���UID=[1] "
                            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strStudyUID)
                            If Not rsTmp.EOF Then lngSeriesNo = NVL(rsTmp(0), 0) + 1
                        End If
                    Else
                        gstrSQL = "select max(���к�) from " & IIf(blnTmp, "Ӱ����ʱ����", "Ӱ��������") & _
                            " where ���UID=[1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strStudyUID)
                        If rsTmp.EOF = False Then
                            lngSeriesNo = NVL(rsTmp(0), 0) + 1
                        Else
                            lngSeriesNo = 1
                        End If
                    End If
                    '�����µ�����
                    gstrSQL = "ZL_Ӱ������_INSERT('" & strStudyUID & "','" & strSeriesUID & "','" & _
                        strSeriesDesp & "'," & IIf(blnTmp, 1, 0) & "," & lngSeriesNo & ")"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
                
                '��������ظ���ͼ���
                lngImageNo = IIf(GetImageAttribute(curImage.Attributes, ATTR_ͼ���) = "", -1, GetImageAttribute(curImage.Attributes, ATTR_ͼ���))
                If lngImageNo <> -1 Then
                    gstrSQL = "select ͼ��� from " & IIf(blnTmp, "Ӱ����ʱͼ��", "Ӱ����ͼ��") & _
                        " where ����UID = [1] and ͼ��� = [2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strSeriesUID, lngImageNo)
                    
                    If rsTmp.EOF = False Then
                        gstrSQL = "select max(ͼ���) from " & IIf(blnTmp, "Ӱ����ʱͼ��", "Ӱ����ͼ��") & _
                            " where ����UID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strSeriesUID)
                        lngImageNo = NVL(rsTmp(0), 0) + 1
                    End If
                Else
                    gstrSQL = "select max(ͼ���) from " & IIf(blnTmp, "Ӱ����ʱͼ��", "Ӱ����ͼ��") & _
                        " where ����UID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strSeriesUID)
                    If rsTmp.EOF = False Then
                        lngImageNo = NVL(rsTmp(0), 0) + 1
                    Else
                        lngImageNo = 1
                    End If
                End If
                '�����µ�ͼ��
                gstrSQL = "ZL_Ӱ��ͼ��_INSERT('" & curImage.instanceUID & "','" & strSeriesUID & "','" _
                    & strSeriesDesp & "'," & IIf(blnTmp, 1, 0) & "," & lngImageNo & "," _
                    & "to_Date('" & Format(GetDateAttribute(curImage.Attributes, ATTR_�ɼ�����, 1) & " " & GetDateAttribute(curImage.Attributes, ATTR_�ɼ�ʱ��, 2), "yyyy-MM-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," _
                    & "to_Date('" & Format(GetDateAttribute(curImage.Attributes, ATTR_ͼ������, 1) & " " & GetDateAttribute(curImage.Attributes, ATTR_ͼ��ʱ��, 2), "yyyy-MM-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'),'" _
                    & GetImageAttribute(curImage.Attributes, ATTR_���) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_ͼ��λ�ò���) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_ͼ������) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_�ο�֡UID) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_��Ƭλ��) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_����) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_����) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_���ؾ���) & "'," _
                    & IIf(curImage.FrameCount = 1, 0, 1) & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
                
                '��B��ͼ�񱣴�ɱ���ͼ��
                If UCase(strModality) = "US" Then
                    gstrSQL = "ZL_Ӱ���鱨��_ADD('" & strStudyUID & "','" & curImage.instanceUID & ".jpg')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
                
                '�������ݿ�����������ͼ��
                gcnOracle.BeginTrans
                blnInDBTrans = True
                For i = 0 To UBound(arrSQL)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ͼ��")
                Next i
                gcnOracle.CommitTrans
                blnInDBTrans = False
                
                
                '���汾����־������Ӱ���������
                WriteRecord strModality, strAdviceID, str����豸, PatientName, EnglishName, Sex, Age, strStudyUID, strSeriesUID, blnTmp
                
                '�Զ�·��
                '--------------------------��û�д���
                Call funcAutoRouting(curImage, BufferDir, dtReceived, strStudyUID, intEncode, strAutoRoute, strAutoRouteCompression, strAutoRouteDir)
            Else        'funGetAEStoreParas�Ľ���
                '��ȡ�����洢�Ĳ��������¼������־,��������һ��ͼ��
                'ƥ�����δ֪������ϵͳ����ķ��������͵�ͼ�񣬲�����
                Call GetAEconnection(Val(curImage.Tag), strServiceAE, strDeviceIP)
                WriteLog 3, vbObjectError + 1, "�� IP= " & strDeviceIP & " ���͸� AE= " & strServiceAE & " ��ͼ���Ѿ����յ������Ǳ������Ӳ���ϵͳ������ķ���ԣ�ͼ���޷����档"
                If strDeviceIP = "" Or strServiceAE = "" Then
                    '���ҷ���AE��IP
                    For i = 1 To UBound(AEconnections)
                        WriteLog 200, 201, " Association = " & Val(curImage.Tag) & " i = " & i & " UBound(AEconnections) = " & UBound(AEconnections) & vbCrLf _
                            & " AEconnections(i).Association = " & AEconnections(i).Association & " AEconnections(i).ServiceAE = " & AEconnections(i).ServiceAE & vbCrLf _
                            & " AEconnections(i).DeviceIP = " & AEconnections(i).DeviceIP & " AEconnections(i).TimeStamp  = " & AEconnections(i).TimeStamp
                    Next i
                End If
            End If
        Else    'end of ���ͼ���Ƿ������ݿ���
            'ͼ���Ѿ����������ݿ��е�ĳ����������ͼ�����¼������־����������һ��ͼ��
            WriteLog 3, vbObjectError + 1, "Ӱ��" & curImage.instanceUID & "�Ѵ��ڣ�"
        End If
        iCount = iCount + 1
        If iCount >= 20 Then Exit For
    Next
    
    For i = 1 To iCount
        Images.Remove 1
    Next
    iNet.FuncFtpDisConnect
    Exit Sub
    
DBError:
    
    If blnInDBTrans = True Then
        gcnOracle.RollbackTrans
    End If
    
    Dim lngerrNumber As Long
    
    lngerrNumber = err.Number   '�ȼ�¼����ţ�������д��־�Ⱥ������ڲ��������˴�������ı�����
    
    '�ȼ�¼������־���ٴ�������
    Call WriteLog(4, err.Number, "����ͼ��ʱ���ִ��󣬴�������Ϊ��" & err.Description)
    
    '�����ض�����
    If lngerrNumber = vbObjectError + 2 Then  '��X���ϴ�ʧ��
        For i = 1 To iCount
            Images.Remove 1
        Next
    ElseIf lngerrNumber = vbObjectError + 3 Or lngerrNumber = vbObjectError + 4 Then  '3�ϴ�ʧ�ܴ����ﵽ���ޣ�4��funSplitSeriesUID�� ����ͼ�񣬲���ͼ�񱣴浽����
        For i = 1 To iCount + 1
            If i = iCount + 1 Then
                '����һ����Ҫ��������ͼ�񣬱��浽����������
                Images(1).WriteFile BufferDir & Images(1).instanceUID & ".bak", True
            End If
            Images.Remove 1
        Next
    End If
    
    iNet.FuncFtpDisConnect
End Sub

Public Sub subSaveAssociation(connection As DicomConnection)
    Dim lngCount  As Long

    '������������
    ReDim Preserve AEconnections(UBound(AEconnections) + 1) As AEconnection
    lngCount = UBound(AEconnections)

    AEconnections(lngCount).ServiceAE = connection.CalledAET
    AEconnections(lngCount).Association = connection.Association
    AEconnections(lngCount).DeviceIP = connection.RemoteIP
    AEconnections(lngCount).TimeStamp = Now
    AEconnections(lngCount).Deleted = False
End Sub

Public Function GetDateAttribute(objAttr As DicomAttributes, ByVal AttrName As String, iType As Integer) As String
'-----------------------------------------------------------------------------
'����:��ȡ�������͵�����ֵ��������ֿ�ֵ�����Զ�ʹ�õ�ǰ����
'������ objAttr ----���Լ���
'       AttrName ----Ҫ���ҵ���������
'       iType ----���� 1--���ڣ�2--ʱ��
'����ֵ�����Ե�����
'-----------------------------------------------------------------------------
    Dim strDateValue As String
    
    strDateValue = GetImageAttribute(objAttr, AttrName)
    
    If iType = 1 Then   '����
        If strDateValue = "" Or IsDate(strDateValue) = False Then
            strDateValue = Date
        End If
        strDateValue = Format(strDateValue, "yyyy-mm-dd")
    Else    'ʱ��
        If strDateValue = "" Or IsDate(strDateValue) = False Then
            strDateValue = Time
        End If
        strDateValue = Format(strDateValue, "hh:mm:ss")
    End If
    
    GetDateAttribute = strDateValue
End Function

Private Function funReConnectFTP(strStoreDeviceNo As String, ByRef ftpNet As clsFtp, strFTPDir As String, intType As Integer) As Long
'-----------------------------------------------------------------------------
'����:��������Ĳ�������������FTP
'������ strStoreDeviceNo ----FTP���ӵ��豸��
'       ftpNet ---- FTP����
'       strFTPDir ----���ص�FTPĿ¼
'       intType ----��ȡ���Ӳ����ķ��� 1--��FTPDevices�����ж�ȡ��2--�����ݿ��в�ѯ
'����ֵ��0--�ɹ���1--����ʧ�ܣ�2--��ȡ�û���������ʧ��
'-----------------------------------------------------------------------------
    Dim strIP As String
    Dim strUser As String
    Dim strPassWord As String
    Dim blnRet As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngResult As Long
    
    On Error GoTo err
    
    '��ȡ��ǰͼ��Ĵ洢�豸
    If intType = 1 Then     '��FTPDevices�����ж�ȡ
        blnRet = funGetFTPDevice(strStoreDeviceNo, strIP, strUser, strPassWord, strFTPDir)
    Else        '�����ݿ��в�ѯ
        strSQL = "select IP��ַ,FTPĿ¼,FTP�û���,FTP���� from Ӱ���豸Ŀ¼  Where �豸��  = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�洢�豸", strStoreDeviceNo)
        If rsTemp.RecordCount = 1 Then
            strIP = NVL(rsTemp("IP��ַ"))
            strUser = NVL(rsTemp("FTP�û���"))
            strPassWord = NVL(rsTemp("FTP����"))
            strFTPDir = NVL(rsTemp("FTPĿ¼")) & "/"
            blnRet = True
        End If
    End If
    
    '��������FTP
    If blnRet = True Then
        ftpNet.FuncFtpDisConnect
        lngResult = ftpNet.FuncFtpConnect(strIP, strUser, strPassWord)
        If lngResult = 0 Then
            'FTP���Ӵ���
            funReConnectFTP = 1
            WriteLog 300, vbObjectError + 1, "funReConnectFTP FTP���Ӵ��󣬸�ͼ���޷����棬�豸�� = " & strStoreDeviceNo
        End If
    Else
        '�����豸�ţ��޷���ȡFTPĿ¼���û���������
        funReConnectFTP = 2
        WriteLog 301, vbObjectError + 1, "funReConnectFTP �޷���ȡFTPĿ¼���û��������룬��ͼ���޷����棬�豸�� = " & strStoreDeviceNo
    End If

    Exit Function
err:
    Call WriteLog(302, err.Number, "funReConnectFTP: " & err.Description)
    ftpNet.FuncFtpDisConnect
End Function

Private Function funSplitSeriesUID(ByRef img As DicomImage, ByRef strSeriesUID As String, ByRef strSeriesDesp As String) As Long
'-----------------------------------------------------------------------------
'����:����ͼ�����Ͳ������UID
'������ img ----��Ҫ��ֵ�ͼ��
'       strSeriesUID ---- ���ص�����UID
'       strSeriesDesp ----���ص���������
'����ֵ��0--�ɹ���1--ʧ��
'-----------------------------------------------------------------------------
    Dim strImageType As String      'ͼ�����LOCALIZER,AXIAL
    Dim vImageType() As String      'ͼ�����
    
    On Error GoTo err
    
    '��ȡͼ�����
    strImageType = GetImageAttribute(img.Attributes, ATTR_ͼ������)
    vImageType = Split(strImageType, "\")
    strImageType = vImageType(3)
    '����ͼ�����Ͳ������
    strSeriesUID = funcGetSeriesUID(strSeriesUID, strImageType)
    strSeriesDesp = strImageType
    img.SeriesUID = strSeriesUID
    
    Exit Function
err:
    Call WriteLog(1001, err.Number, "funSplitSeriesUID: " & err.Description)
    funSplitSeriesUID = 1
End Function

Private Function funUploadImage(ByRef img As DicomImage, ByRef ftpNet As clsFtp, ByVal intEncode As Integer, _
    ByVal strBufferDir As String, ByVal strFTPDir As String, ByVal strStudyUID As String, ByVal strDtReceived As String) As Long
'-----------------------------------------------------------------------------
'����:����ͼ��FTP��
'������ img ----��Ҫ�����ͼ��
'       ftpNet ---- FTP����
'       intEncode ---- ѹ����ʽ
'       strBufferDir ---- ���ػ���·��
'       strFTPDir ---- FTP�Ĵ洢Ŀ¼
'       strStudyUID ---- ���UID
'       strDtReceived --- �������ڣ�FTPĿ¼����ɲ���
'����ֵ��0--�ɹ���1--��X�γ����ϴ�ʧ�ܣ�2--�ϴ�ʧ�ܴ����ﵽ���ޣ�����ͼ��3--��������
'-----------------------------------------------------------------------------
    Dim blnNoCompress As Boolean    '��¼��ǰͼ���Ƿ���Ҫѹ��
    Dim lngResult As Long           '��¼����ֵ

    On Error GoTo err
    
    '�����ж�ͼ���Ƿ����ڲ���ѹ���ģ�����Philips��3D�ؽ�Ч��ͼ�Ͳ���ѹ����ѹ����ͼ����ɺڰ�
    blnNoCompress = False
    If Not IsNull(img.Attributes(&H28, &H2)) And img.Attributes(&H28, &H2).Exists _
        And Not IsNull(img.Attributes(&H28, &H4)) And img.Attributes(&H28, &H4).Exists _
        And Not IsNull(img.Attributes(&H28, &H6)) And img.Attributes(&H28, &H6).Exists Then
        
        If img.Attributes(&H28, &H2).value = 3 And img.Attributes(&H28, &H4).value = "RGB" _
            And img.Attributes(&H28, &H6).value = 1 Then
            
            blnNoCompress = True
        End If
    End If
    
    '�ж�ͼ���Ƿ���PNMS�豸������CTͼ������ǣ������ӿհ״������ض���
    If Not IsNull(img.Attributes(&H8, &H70)) And img.Attributes(&H8, &H70).Exists Then
        If UCase(img.Attributes(&H8, &H70)) = "PNMS" Then
            If UCase(img.Attributes(&H8, &H60)) = "CT" Then
                img.Attributes.Add &H28, &H120, -2000
            End If
        End If
    End If
    
    Call WriteProcessLog("funUploadImage", "�ϴ�ͼ��FTP", "ͼ��ѹ��=" & blnNoCompress & "�� ͼ��ѹ����ʽ=" & intEncode)
    
    If blnNoCompress = True Then
        img.WriteFile strBufferDir & img.instanceUID, True
    Else
        Select Case intEncode
            Case 0
                img.WriteFile strBufferDir & img.instanceUID, True, TS_JPEG����ѹ��
            Case 1
                img.WriteFile strBufferDir & img.instanceUID, True, TS_RLE�г�ѹ��
            Case 2
                img.WriteFile strBufferDir & img.instanceUID, True, TS_JPEG2000����ѹ��
            Case 3
                img.WriteFile strBufferDir & img.instanceUID, True
        End Select
    End If
    
    Call WriteProcessLog("funUploadImage", "���汾�ػ����ļ�", "���ػ����ļ���Ϊ��" & strBufferDir & img.instanceUID)
    
    '�ϴ�FTPͼ���ļ�
    lngResult = WriteToURL(ftpNet, strBufferDir & img.instanceUID, strFTPDir & "/" & _
        strDtReceived & "/" & strStudyUID & "/" & img.instanceUID)
    
    '����ϴ�ʧ�ܣ�����ж�Ӧ�Ĵ���ʹ��BorderWidth����ʱ����ͼ�񱻳����ϴ��Ĵ���
    '�����ϴ�10�ζ�ʧ�ܣ����������ͼ��
    If lngResult <> 0 Then
        If NVL(img.BorderWidth, 0) = 0 Then
            img.BorderWidth = 1
        Else
            img.BorderWidth = img.BorderWidth + 1
        End If
        If img.BorderWidth < 10 Then
            funUploadImage = 1
            
            'FTP �� img.BorderWidth �δ洢ʧ�ܣ�ɾ����ʱͼ��
            Kill strBufferDir & img.instanceUID
            Exit Function
        Else
            funUploadImage = 2
            
            'ͼ�񱻷�����FTP �� img.BorderWidth �δ洢ʧ�ܣ�ɾ����ʱͼ��
            Kill strBufferDir & img.instanceUID
            Exit Function
        End If
    End If
    
    '���ͨ��DICOM��ʽ����B��ͼ��������Զ���B��ͼ�񱣴�ɱ���ͼ��
    If UCase(GetImageAttribute(img.Attributes, ATTR_Ӱ�����)) = "US" Then
        img.FileExport strBufferDir & img.instanceUID & ".jpg", "JPG", 80
        WriteToURL ftpNet, strBufferDir & img.instanceUID & ".jpg", strFTPDir & "/" & _
            strDtReceived & "/" & strStudyUID & "/" & img.instanceUID & ".jpg"
    End If
    
    'ɾ����ʱͼ��
    Kill strBufferDir & img.instanceUID
    Exit Function
err:
    Call WriteLog(1001, err.Number, "funUploadImage: " & err.Description)
    funUploadImage = 3
End Function

Private Function funGetMatchIDInImg(img As DicomImage, intImgMatchItem As Integer, intDatabaseMatchItem As Integer) As String
'-----------------------------------------------------------------------------
'����:������������ȡͼ���е�ƥ��ID
'������ img ----��Ҫƥ���ͼ��
'       intImgMatchItem ---- ƥ�����Ŀ��0--PatientID��1--AccessionNumber��2--PatientName
'       intDatabaseMatchItem ---���ݿ�ƥ����Ŀ�� 0--����ƥ��;1--���˱�ʶƥ��;2--����ʶƥ��
'����ֵ��ƥ��ID�����ͼ��������Ϊ�գ�����-1
'-----------------------------------------------------------------------------
    Dim aPatientID() As String
    Dim strTemp As String
    
    On Error GoTo err

    If intDatabaseMatchItem = 0 Then    '����ƥ�䣬֧���ַ��ͣ�����Ҫ���ָ���ֱ����ȡ�������ŵ�����
        Select Case intImgMatchItem
            Case 0 'Patient ID
                funGetMatchIDInImg = NVL(img.PatientID)
            Case 1 'Accession Number
                funGetMatchIDInImg = NVL(img.Attributes(&H8, &H50).value)
            Case 2 'Patient Name
                funGetMatchIDInImg = NVL(img.Name)
        End Select
        
        'ȥ��ƥ�����ǰ������0
        While Left(funGetMatchIDInImg, 1) = "0"
            funGetMatchIDInImg = Mid(funGetMatchIDInImg, 2)
        Wend
        
        Exit Function
    End If
    
    '����ID��ҽ��IDƥ�䣬��Ҫ��ͼ������ȡ��ֵ�͵�ID����ƥ��
    Select Case intImgMatchItem
        Case 0 'Patient ID
            aPatientID = Split(Replace(NVL(img.PatientID), "-", "_"), "_")
        Case 1 'Accession Number
            aPatientID = Split(Replace(NVL(img.Attributes(&H8, &H50).value), "-", "_"), "_")
        Case 2 'Patient Name
            aPatientID = Split(Replace(NVL(img.Name), "-", "_"), "_")
    End Select
    
    If UBound(aPatientID) >= 0 Then
        If UBound(aPatientID) > 0 Then
            strTemp = aPatientID(1)
        Else
            strTemp = aPatientID(0)
        End If
        
        '������ֵ���ַ��������һ������
        strTemp = getStrNumber(strTemp)
        funGetMatchIDInImg = strTemp
    Else
        funGetMatchIDInImg = -1
    End If
    Exit Function
err:
    funGetMatchIDInImg = 0
End Function

Private Function funIsPreMatched(ByVal blnMatchStudyUID As Boolean, ByVal intDBMatchItem As Integer, ByRef strStudyUID As String, _
    ByVal strAdviceID As String, ByVal strDeviceIP As String, ByVal strSeriesUID As String, ByVal strModality As String, _
    ByRef dtReceived As String, ByVal intFilterModality As Integer, ByRef strNewDeviceID As String, _
    ByVal strStoreDeviceNo As String, ByRef blnTmp As Boolean, ByRef str����豸 As String, ByRef strPatientName As String, _
    ByRef strEnglishName As String, ByRef intAge As Integer, ByRef strSex As String, ByVal strStudyDateTime As String) As Long
'-----------------------------------------------------------------------------
'����:�ж��Ƿ��Ѿ���ƥ��ɹ��ļ�¼
'������ blnMatchStudyUID ----�Ƿ�ƥ����UID
'       intDBMatchItem ---- ƥ������ݿ���Ŀ��0--����ƥ�䣬1--���˱�ʶƥ�䣬2--����ʶƥ��
'       strStudyUID ---- [IN][OUT]���UID����ѯ������鵽�����޸ļ��UID
'       strAdviceID ---- ͼ���е�ƥ��ID�������������PatientID��PatientName��AccessionNumber��ͳ��Ϊҽ��ID
'       strDeviceIP ---- �洢�豸IP
'       strSeriesUID ---- ͼ������UID
'       strModality ---- Ӱ�����
'       dtReceived ---[OUT] ���Ԥƥ��ɹ����򷵻����ݿ��С�Ӱ�����¼���ġ��������ڡ�
'       intFilterModality ---- �Ƿ���Ӱ��������
'       strNewDeviceID ---- ��ѯ�����´洢�豸ID
'       strStoreDeviceNo ---- ԭ���Ĵ洢�豸��
'       blnTmp ---- �Ƿ�ƥ�����ʱ��¼
'       str����豸 ---- [IN][OUT]ͼ���еļ���豸�����ƥ��ɹ������޸ĳ����ݿ��еļ���豸
'       strPatientName ----[OUT] ���ƥ��ɹ����������ݿ��е�������
'       strEnglishName ----[OUT] ���ƥ��ɹ����������ݿ��е�Ӣ����
'       intAge ----[OUT] ���ƥ��ɹ����������ݿ��е�����
'       strSex ----[OUT] ���ƥ��ɹ����������ݿ��е��Ա�
'       strStudyDateTime ---- ������Ԥƥ���ѯ��ʱ��������������ռ��UIDƥ�䣬ʹ�õ��ǿ����ظ��ļ��ţ���ʶ��ʱ��ʹ�ô˲�����ʱ����бȶԣ��������ʱ��ļ���¼��
'����ֵ��0-ƥ��ɹ���1-��ƥ���¼
'-----------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    If blnMatchStudyUID Then    '���ռ��UIDƥ��
        Select Case intDBMatchItem
            Case 0, 1 '����ƥ��,'���˱�ʶ(����ţ�סԺ��)ƥ�� ��ֱ���жϼ��UID�Ƿ���ͬ
                strSQL = "Select 0 As ��ʱ,���UID,��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ�����¼ A" & _
                    " Where A.���UID= [1]"
            Case 2 '����ʶ��ҽ��ID��ƥ��
                strSQL = "Select 0 As ��ʱ,���UID,��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ�����¼ Where ���UID= [1]" & _
                    " Or (ҽ��ID= [2] And ���UID Is Not Null)"
        End Select
        strSQL = strSQL & " Union All Select 1 As ��ʱ,���UID,��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ����ʱ��¼ Where ���UID= [1]"
        
        strSQL = strSQL & " Union All Select 0 As ��ʱ,C.���UID,C.��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ�����¼ C, Ӱ�������� D Where C.���UID = D.���UID And D.����UID = [4] " & _
            " Union All Select 1 As ��ʱ,E.���UID,E.��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ����ʱ��¼ E, Ӱ����ʱ���� F Where E.���UID = F.���UID And F.����UID = [4]"
    Else    '�����ռ��UIDƥ��
        Select Case intDBMatchItem
            Case 0 '����ƥ��
                strSQL = "Select 0 As ��ʱ,���UID,��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ�����¼ A,����ҽ������ B,Ӱ���豸Ŀ¼ E " & _
                    " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� AND A.����豸 =E.�豸�� AND (B.ִ��״̬=3 " & _
                    " And B.ִ�й���>2 And " & _
                    IIf(intFilterModality = 0, " UPPER(E.Ӱ�����)=[5] ", " E.IP��ַ=[3] ") & " And A.����= [2] And A.���UID Is Not Null" & _
                    " And abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.�״�ʱ��) = (Select min(abs(to_date('" & _
                    strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-D.�״�ʱ��)) from Ӱ�����¼ C, ����ҽ������ D,Ӱ���豸Ŀ¼ F Where C.ҽ��ID=D.ҽ��ID" & _
                    " And C.���ͺ�=D.���ͺ� AND C.����豸 =F.�豸�� AND (D.ִ��״̬=3 And D.ִ�й���>=2 And " & _
                    IIf(intFilterModality = 0, " UPPER(F.Ӱ�����)=[5] ", " F.IP��ַ=[3] ") & " And C.����= [2])))"
            Case 1 '���˱�ʶ������ţ�סԺ�ţ�ƥ��
                strSQL = "Select 0 As ��ʱ,���UID,��������,λ��һ,λ�ö�,����豸,a.����,a.Ӣ����,a.�Ա�,a.���� From Ӱ�����¼ A,����ҽ������ B,����ҽ����¼ C,������Ϣ D,Ӱ���豸Ŀ¼ I " & _
                    " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And A.ҽ��ID=C.ID And C.����ID=D.����ID And A.����豸 =I.�豸�� " & _
                    " And B.ִ��״̬=3 And B.ִ�й���>2 And " & _
                    IIf(intFilterModality = 0, " UPPER(I.Ӱ�����)=[5] ", " I.IP��ַ=[3] ") & _
                    " And ((D.סԺ��=[2] AND C.������Դ=2) OR (D.�����= [2] AND C.������Դ<>2))" & _
                    " And A.���UID Is Not Null  AND C.���ID IS NULL  " & _
                    " And abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.�״�ʱ��) = (Select min(abs(to_date('" & _
                    strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-F.�״�ʱ��)) from Ӱ�����¼ E,����ҽ������ F,����ҽ����¼ G,������Ϣ H,Ӱ���豸Ŀ¼ J " & _
                    " Where E.ҽ��ID=F.ҽ��ID And E.���ͺ�=F.���ͺ� And E.ҽ��ID=G.ID And G.����ID=H.����ID AND E.����豸 =J.�豸�� AND G.���ID IS NULL " & _
                    " And F.ִ��״̬=3 And F.ִ�й���>=2 And " & _
                    IIf(intFilterModality = 0, " UPPER(J.Ӱ�����)=[5] ", " J.IP��ַ=[3] ") & _
                    " And ((H.סԺ��=[2] AND G.������Դ=2) OR (H.�����= [2] AND G.������Դ<>2)))"
                    
            Case 2 '����ʶ��ҽ��ID��ƥ��
                strSQL = "Select 0 As ��ʱ,���UID,��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ�����¼ A Where  A.ҽ��ID= [2] And A.���UID Is Not Null"
        End Select
        strSQL = strSQL & " Union All Select 0 As ��ʱ,C.���UID,C.��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ�����¼ C, Ӱ�������� D Where C.���UID = D.���UID And D.����UID = [4] " & _
            " Union All Select 1 As ��ʱ,E.���UID,E.��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ����ʱ��¼ E, Ӱ����ʱ���� F Where E.���UID = F.���UID And F.����UID = [4]"
        '����ѯ�� lngAdviceID ���벻Ϊ-1ʱ���ٲ�ѯ��ʱ���Ƿ���ƥ��ɹ��ġ�
        If strAdviceID <> "-1" Then
            strSQL = strSQL & " Union All Select 1 As ��ʱ,���UID,��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ����ʱ��¼ Where ����= [2] and UPPER(Ӱ�����) =[5] and ���UID= [1] "
        Else
            '���ƥ��ID��-1��˵��ƥ��IDΪ�գ���ʱ��ѯ��ʱ��ʱ��Ӧ�ð��ռ��UID����ѯ����Ϊ���汣������ͼ��ʱ��������ȡ��һ��ƥ��ID
            strSQL = strSQL & " Union All Select 1 As ��ʱ,���UID,��������,λ��һ,λ�ö�,����豸,����,Ӣ����,�Ա�,���� From Ӱ����ʱ��¼ Where ���UID= [1]"
        End If
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "PACSͼ�񱣴�", strStudyUID, strAdviceID, strDeviceIP, strSeriesUID, strModality)
               
    '��¼SQL��־
    WriteProcessLog "funIsPreMatched", "�ж��Ƿ��Ѿ���ƥ��ɹ��ļ�¼", "��ѯSQLΪ�� " & vbCrLf & Replace(strSQL, "'", "��") & vbCrLf _
                    & "����Ϊ��1=" & strStudyUID & " ,2=" & strAdviceID & " ,3=" & strDeviceIP & " ,4=" & strSeriesUID & " ,5=" & strModality
    
    '���ƥ��ɹ������¼ƥ��ҽ���ļ��UID������ʱ�䡢�Ƿ���ʱ��¼������豸��
    If rsTemp.EOF = False Then
        strStudyUID = rsTemp("���UID")
        dtReceived = Format(rsTemp("��������"), "yyyyMMdd")
        blnTmp = IIf(rsTemp("��ʱ") = 1, True, False)    '���к�ͼ���Ƿ������ʱ��¼��
        str����豸 = NVL(rsTemp("����豸"))
        strPatientName = NVL(rsTemp("����"))
        strEnglishName = NVL(rsTemp("Ӣ����"))
        intAge = Val(NVL(rsTemp("����"), 0))
        strSex = NVL(rsTemp("�Ա�"))
        
        '�жϸ�ͼ�����ڵļ�¼�У��洢�豸�Ƿ���ڵ�ǰ���õĴ洢�豸
        If NVL(rsTemp("λ��һ")) <> "" Then
            strNewDeviceID = NVL(rsTemp("λ��һ"))
        ElseIf NVL(rsTemp("λ�ö�")) <> "" Then
            strNewDeviceID = NVL(rsTemp("λ�ö�"))
        Else    'λ��һ��λ�ö���û�д洢�豸��
            '��¼������־��Ȼ��ʹ�õ�ǰ���õĴ洢�豸��
            WriteLog 11, 100, "�Ӳ��˵�Ӱ�����¼���޷��ҵ��洢�豸��ʹ���������õĴ洢�豸����ͼ��" & " ����������" & strPatientName
            strNewDeviceID = strStoreDeviceNo
        End If
        funIsPreMatched = 0
    Else
        funIsPreMatched = 1
    End If
    Exit Function
err:
    Call WriteLog(1002, err.Number, "funIsPreMatched: " & err.Description)
    funIsPreMatched = 1
End Function

Public Function funcGetLocalIP() As String
'���ص�ǰ�������IP��ַ�����ö��ŷָ�
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim strLocalIPs As String

    '����Socket
    Call SocketsInitialize

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If

    hostent_addr = gethostbyname(hostname)

    If hostent_addr = 0 Then
        MsgBox "Winsock.dll is not responding."
        Exit Function
    End If

    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4

    ''''''''''''''''get all of the IP address if machine is  multi-homed

    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

        For i = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

        strLocalIPs = IIf(strLocalIPs = "", ip_address, strLocalIPs & "," & ip_address)

        ip_address = ""
        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
     Loop While (hostip_addr <> 0)

    '���Socket
    Call SocketsCleanup
    
    funcGetLocalIP = strLocalIPs
End Function


Private Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        MsgBox "Winsock.dll is not responding."
        End
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        End
    End If

    ''''''''''''''''iMaxSockets is not used in winsock 2. So the following check is only
    ''''''''''''''''necessary for winsock 1. If winsock 2 is requested,
    ''''''''''''''''the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        End
    End If
End Sub

Private Sub SocketsCleanup()
Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
        End
    End If
End Sub

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Private Function getStrNumber(ByVal strNumber As String) As String
    '������ֵ���ַ��������һ������������ֱ��ʹ��VAL��
    '��Ϊ������ܳ���16λ��ʹ��VAL���Զ�ת���ɿ�ѧ���������������λ����������
    
    On Error GoTo err
    If IsNumeric(strNumber) Then
        getStrNumber = strNumber
    Else
        If Val(strNumber) > 999999999999999# Then
            getStrNumber = 0
        Else
            getStrNumber = Val(strNumber)
        End If
    End If
    Exit Function
err:
      getStrNumber = 0
End Function

Public Sub WriteProcessLog(logSubName As String, logTitle As String, logDesc As String)
'���ܣ� ��¼������־
'������ logSubName -- ִ�в����Ĺ�������
'       logTitle  --  ��������
'       logDesc --    ��־����

    Dim strSQL As String
    
    On Error Resume Next
    If gblnProcessLog Then
        If gcnAccess.State = adStateClosed Then Exit Sub
        
        strSQL = "Insert into DICOMͨѶ��־ (ͨѶʱ��,ͨѶ����,��¼����,��¼����) " & _
            "Values( cDate('" & Date & " " & Time() & "'),'" & logSubName & "','" & logTitle & _
            "','" & logDesc & "')"
        gcnAccess.Execute strSQL
    End If
End Sub

Public Sub subNewLogFile()
'���ܣ� �����µ���־�ļ�
'������ ��
    
    Dim strDate As String
    
    On Error GoTo err
    
    '������ǰ������ʱ����
    strDate = Date & "-" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time)
    
    '������־�ļ�֮ǰ���ȹر���־�ļ�
    If gcnAccess.State <> adStateClosed Then gcnAccess.Close
    FileCopy gstrAccessName, gstrAccessPath & "-" & strDate & ".mdb"
        
    '�����������ݿ�
    gcnAccess.Open
    '��յ�ǰ��־�е�����
    gcnAccess.Execute "delete from DICOMͨѶ��־"
    gcnAccess.Execute "delete from ������־"
    gcnAccess.Execute "delete from Ӱ���������"
    
    
    'ѹ�����ݿ��ļ�
    gcnAccess.Close
    DBEngine.CompactDatabase gstrAccessName, gstrAccessPath & "-zip.mdb"
    Kill gstrAccessName
    FileCopy gstrAccessPath & "-zip.mdb", gstrAccessName
    Kill gstrAccessPath & "-zip.mdb"
    gcnAccess.Open
    
    Exit Sub
err:
    Call WriteLog(1013, err.Number, "��������־���ִ��󣬴��������ǣ�" & err.Description)
End Sub

Private Function funDeleteSpcialChar(strString As String) As String
'ɾ�����ɼ��������ַ���ascii����0��32֮��������ַ�,32Ϊ�ո�
    Dim i As Integer
    Dim strText As String
    
    On Error GoTo err
    
    strText = strString
    For i = 0 To 32
        strText = Replace(strText, Chr(i), "")
    Next i
    
    funDeleteSpcialChar = strText
    Exit Function
err:
    funDeleteSpcialChar = strString
    
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, "��ʾ"
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function

Public Function DynamicGet(ByVal strclass As String, ByVal strCaption As String) As Object
'��̬��������
    On Error Resume Next
    Set DynamicGet = GetObject("", strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, "��ʾ"
        Set DynamicGet = Nothing
    End If
    err.Clear
End Function
