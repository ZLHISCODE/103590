Attribute VB_Name = "mdlComm"
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'全局通讯参数
Public gstrIP As String         'IP
Public glngPort As Integer      '端口
Public gintStart As Integer     '是否启用通讯

Public Function GetIniKeyValue(ByVal strPathAndFileName As String, ByVal strItem As String, ByVal strKey As String, Optional ByVal strDefault As String) As String
        '读取Ini文件中的值
        '配置文件不存在则创建，并写入默认值
        Dim objFile As New FileSystemObject


1         On Error GoTo GetIniKeyValue_Error

2       If Not objFile.FileExists(strPathAndFileName) Then
3           Call WriteIni(strItem, strKey, strDefault, strPathAndFileName)
4           GetIniKeyValue = strDefault
5       Else
6           GetIniKeyValue = ReadIni(strItem, strKey, strPathAndFileName)
7           If GetIniKeyValue = "" And strDefault <> "" Then
8               Call WriteIni(strItem, strKey, strDefault, strPathAndFileName)
9               GetIniKeyValue = strDefault
10          End If
11      End If



12        Exit Function
GetIniKeyValue_Error:
13        Call writeErrLog("zlPublicLIS", "mdlComm", "执行(GetIniKeyValue)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
14        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/24
'功    能:向外部提供获取参数的接口
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function funGetIniValue(ByVal strKey As String, Optional ByVal strDefault As String) As String
    funGetIniValue = GetIniKeyValue(App.Path & "\LisCommLocal.ini", "SET", strKey, strDefault)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/24
'功    能:写文件
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function WriteIni(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
    On Error GoTo errH
    WriteIni = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    WriteIni = False
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/24
'功    能:读文件
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function ReadIni(strItem As String, strKey As String, strPath As String, Optional strDefault As String = "") As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = VBA.String(128, 0)
    GetPrivateProfileString strItem, strKey, strDefault, GetStr, 256, strPath
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/10
'功    能:判断输入的内容是否是IP
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function IsIP(ByRef strIP As String) As Boolean
          Dim strPart() As String
          Dim i As Integer
          Dim j As Integer

1         On Error GoTo IsIP_Error

2         IsIP = True
3         strPart = Split(strIP, ".")
4         strIP = ""
5         If (UBound(strPart) <> 3) Then
6             IsIP = False
7         Else
8             For i = LBound(strPart) To UBound(strPart)
9                 j = Int(strPart(i))
10                If ((j < 0) And (j > 255)) Then
11                    IsIP = False
12                    Exit Function
13                End If
14                strIP = strIP & j & "."
15            Next i
16            If strIP <> "" Then
17                strIP = Mid(strIP, 1, Len(strIP) - 1)
18            End If
19        End If

20        Exit Function
IsIP_Error:
21        IsIP = False
22        Call writeErrLog("zlPublicLIS", "mdlComm", "执行(IsIP)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
23        Err.Clear
End Function

Public Function funInit() As Boolean

          Dim strPara As String
          '通讯参数

1         On Error GoTo funInit_Error

2         strPara = ComGetPara("LIS远程通讯参数", 2500, 2500, "0|127.0.0.1|8888")
3         gintStart = Val(Split(strPara, "|")(0))
4         gstrIP = Split(strPara, "|")(1)
5         glngPort = Split(strPara, "|")(2)


6         Exit Function
funInit_Error:
7         Call writeErrLog("zlPublicLIS", "mdlComm", "执行(funInit)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
8         Err.Clear
End Function
