Attribute VB_Name = "mdlLis"
Option Explicit

Public gblnInited   As Boolean
Public gcolPlugIn   As Collection '��չ�������ü��Ϸ�ʽ�ݴ���չ�������ʵ��
Public grsSample As ADODB.Recordset '��Ҫ��ӡ�ı걾�ļ�¼��
Public gblnDoctor As Boolean        '�Ƿ�ҽ��վ��ӡ
Public gblnPrintAll As Boolean      '�Ƿ��ӡ���б��棬True=��ӡ���б��棬False=��ӡΪ��ӡ�ı���

Public gobjCommFun      As Object
Public gobjControl      As Object
Public gobjPrintMode    As Object
Public gobjSystem       As Object


' ע�����������...
Public Enum ValueType
    REG_SZ = 1                         ' �ַ���ֵ
    REG_EXPAND_SZ = 2                  ' �������ַ���ֵ
    REG_BINARY = 3                     ' ������ֵ
    REG_DWORD = 4                      ' DWORDֵ
    REG_MULTI_SZ = 7                   ' ���ַ���ֵ
End Enum


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private marrName As Variant

Public Function ChkRsState(rs As ADODB.Recordset) As Boolean
On Error GoTo ErrH:
    With rs
        If rs Is Nothing Then
            ChkRsState = True
            Exit Function
        Else
            ChkRsState = False
        End If
        If rs.State = 0 Then
            ChkRsState = True
            Exit Function
        Else
            ChkRsState = False
        End If
        If .RecordCount < 1 Then
            ChkRsState = True
        Else
            ChkRsState = False
        End If
        If .EOF Or .BOF Then
            ChkRsState = True
        Else
            ChkRsState = False
        End If
    End With
    Exit Function
ErrH:
    err.Clear
End Function

Public Function ReadIni(strItem As String, strKey As String, strPath As String) As String
    Dim GetStr As String
    On Error GoTo ErrH

    GetStr = VBA.String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    ReadIni = GetStr
    Exit Function
ErrH:
    err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
    On Error GoTo ErrH
    WriteIni = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
ErrH:
    err.Clear
    WriteIni = False
End Function

Public Sub WriteLog(ByVal strFunc As String, ByVal strInput As String, ByVal strOutput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------
    
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate         As String
    Dim strFilename     As String
    Dim objStream       As TextStream
    Dim objFso          As New FileSystemObject
    
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    If Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv", "��ս�����־", 1)) = 1 Then
        If Dir(App.Path & "\����.TXT") = "" Then Exit Sub
    End If
    strFilename = App.Path & "\LisDev_" & Format(Date, "yyyyMMdd") & ".LOG"
    
    If Not objFso.FileExists(strFilename) Then Call objFso.CreateTextFile(strFilename)
    Set objStream = objFso.OpenTextFile(strFilename, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (String(50, "��"))
    objStream.WriteLine ("ִ��ʱ��:" & strDate & "�汾:" & App.Major & "." & App.Minor & "." & App.Revision)
    objStream.WriteLine ("����:" & strFunc)
    objStream.WriteLine ("  :" & strInput)
    objStream.WriteLine ("  :" & strOutput)
    'objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objFso = Nothing
    Set objStream = Nothing
End Sub


