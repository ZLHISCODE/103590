Attribute VB_Name = "mdlLis"
Option Explicit

Public gblnInited   As Boolean
Public gcolPlugIn   As Collection '扩展部件，用集合方式暂存扩展部件类的实例
Public grsSample As ADODB.Recordset '需要打印的标本的记录集
Public gblnDoctor As Boolean        '是否医生站打印
Public gblnPrintAll As Boolean      '是否打印所有报告，True=打印所有报告，False=打印为打印的报告

Public gobjCommFun      As Object
Public gobjControl      As Object
Public gobjPrintMode    As Object
Public gobjSystem       As Object


' 注册表数据类型...
Public Enum ValueType
    REG_SZ = 1                         ' 字符串值
    REG_EXPAND_SZ = 2                  ' 可扩充字符串值
    REG_BINARY = 3                     ' 二进制值
    REG_DWORD = 4                      ' DWORD值
    REG_MULTI_SZ = 7                   ' 多字符串值
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
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------
    
    '以下变量用于记录调用接口的入参
    Dim strDate         As String
    Dim strFilename     As String
    Dim objStream       As TextStream
    Dim objFso          As New FileSystemObject
    
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv", "清空接收日志", 1)) = 1 Then
        If Dir(App.Path & "\调试.TXT") = "" Then Exit Sub
    End If
    strFilename = App.Path & "\LisDev_" & Format(Date, "yyyyMMdd") & ".LOG"
    
    If Not objFso.FileExists(strFilename) Then Call objFso.CreateTextFile(strFilename)
    Set objStream = objFso.OpenTextFile(strFilename, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (String(50, "≡"))
    objStream.WriteLine ("执行时间:" & strDate & "版本:" & App.Major & "." & App.Minor & "." & App.Revision)
    objStream.WriteLine ("驱动:" & strFunc)
    objStream.WriteLine ("  :" & strInput)
    objStream.WriteLine ("  :" & strOutput)
    'objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objFso = Nothing
    Set objStream = Nothing
End Sub


