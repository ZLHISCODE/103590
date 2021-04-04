Attribute VB_Name = "mdlComm"
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'ȫ��ͨѶ����
Public gstrIP As String         'IP
Public glngPort As Integer      '�˿�
Public gintStart As Integer     '�Ƿ�����ͨѶ

Public Function GetIniKeyValue(ByVal strPathAndFileName As String, ByVal strItem As String, ByVal strKey As String, Optional ByVal strDefault As String) As String
        '��ȡIni�ļ��е�ֵ
        '�����ļ��������򴴽�����д��Ĭ��ֵ
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
13        Call writeErrLog("zlPublicLIS", "mdlComm", "ִ��(GetIniKeyValue)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
14        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/24
'��    ��:���ⲿ�ṩ��ȡ�����Ľӿ�
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Function funGetIniValue(ByVal strKey As String, Optional ByVal strDefault As String) As String
    funGetIniValue = GetIniKeyValue(App.Path & "\LisCommLocal.ini", "SET", strKey, strDefault)
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/24
'��    ��:д�ļ�
'��    ��:
'��    ��:
'��    ��:
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
'��    ��:������
'����ʱ��:2017/4/24
'��    ��:���ļ�
'��    ��:
'��    ��:
'��    ��:
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
'��    ��:������
'����ʱ��:2017/5/10
'��    ��:�ж�����������Ƿ���IP
'��    ��:
'��    ��:
'��    ��:
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
22        Call writeErrLog("zlPublicLIS", "mdlComm", "ִ��(IsIP)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear
End Function

Public Function funInit() As Boolean

          Dim strPara As String
          'ͨѶ����

1         On Error GoTo funInit_Error

2         strPara = ComGetPara("LISԶ��ͨѶ����", 2500, 2500, "0|127.0.0.1|8888")
3         gintStart = Val(Split(strPara, "|")(0))
4         gstrIP = Split(strPara, "|")(1)
5         glngPort = Split(strPara, "|")(2)


6         Exit Function
funInit_Error:
7         Call writeErrLog("zlPublicLIS", "mdlComm", "ִ��(funInit)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
8         Err.Clear
End Function
