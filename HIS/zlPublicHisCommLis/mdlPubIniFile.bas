Attribute VB_Name = "mdlPubIniFile"
Option Explicit

Public gcollLog As New Collection   '日志对象

'读写ini 文件的API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'创建目录
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp" (ByVal PathName As String) As Long

Public Function ReadIni(strItem As String, strKey As String, strPath As String) As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = Replace(GetStr, Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
    On Error GoTo errH
    WriteIni = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    WriteIni = False
End Function


Public Sub SaveDBLog(intSource As Integer, intPriority As Integer, lngSampleID As Long, strProperties As String, strExplain As String, Optional intNO As Integer, Optional strProgram As String, Optional lngMachineID As Long)
          '功能     保存日志
          '参数     intSource       来源
          '         intPriority    优先级
          '         lngSampleID    标本ID
          '         strProperties  性质
          '         strExplain     说明
          '         intNo          模块号
          '         strProgram     模块名称
          Dim strSQL As String
          Dim strLocalIP As String

1     On Error GoTo SaveDBLog_Error

'2         strLocalIP = frmCommSetup.Winsock.LocalIP
          strLocalIP = gobjHisSystem.IP
3         strSQL = "Zl_检验操作日志_insert(" & intSource & "," & intPriority & ",'" & gUserInfo.Name & "'," & IIf(lngSampleID = 0, "null", lngSampleID) & ",'" & strProperties & "','" & strExplain & "'," & intNO & ",'" & strProgram & "'," & IIf(lngMachineID = 0, "null", lngMachineID) & ",'" & strLocalIP & "')"
4         Call ComExecuteProc(Sel_Lis_DB, strSQL, "日志保存")

5         Exit Sub
SaveDBLog_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlPubIniFile", "执行(SaveDBLog)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/7/12
'功    能:错误日志
'入    参:
'           strObject       出错的工程名
'           strModule       出错的模块
'           strTxt          错误信息
'           [blnMsg         是否弹出提示窗]
'           [intType        0=错误日志，1=通讯日志]
'出    参:
'返    回:  错误提示
'---------------------------------------------------------------------------------------
Public Function WriteErrLog(ByVal strObject As String, ByVal strModule As String, ByVal strErrTxt As String, _
                            Optional ByVal blnMsg As Boolean, Optional ByVal intType As Integer) As String
          Dim objFSO As New FileSystemObject
          Dim objFolder As Folder
          Dim objFiles As Files
          Dim objFile As File
          Dim strFileName As String
          Dim dFileCreatDate As Date
          Dim objTxt As TextStream
          Dim strFolder As String         '错误日志目录
          Dim strFilePath As String       '错误日志路径
          Dim dTimeNow As Date            '当前系统时间
          Dim strDrive As String          '磁盘名称
          Dim DrDrive As Drive            '磁盘驱动器
          Dim strOrcMsg As String

1         On Error GoTo Errhand

2         strFolder = Mid(App.Path, 1, InStrRev(App.Path, "\")) & "PUBLIC"
3         If intType = 0 Then
4             strFolder = strFolder & "\Log\错误日志\"
5             strFilePath = strFolder & "错误日志.log"
6         Else
7             strFolder = strFolder & "\Log\通讯日志\"
8             strFilePath = strFolder & "通讯日志.log"
9         End If
10        dTimeNow = Now    ' 使用ComCurrDate时，如果在ComCurrDate内部出现错误，在ComCurrDate内部调用错误日志（writeErrLog），则会出现死循环

          '检查目录是否存在，不存在则创建目录
11        If Not objFSO.FolderExists(strFolder) Then
12            Call MakeSureDirectoryPathExists(strFolder)
13        End If

          '检查文件是否存在，不存在则创建
14        If objFSO.FileExists(strFilePath) = False Then
              '删除创建超过30天的备份文件
15            Set objFolder = objFSO.GetFolder(strFolder)
16            Set objFiles = objFolder.Files
17            For Each objFile In objFiles
18                strFileName = objFile.Name
19                If IsDate(Format(Mid(strFileName, 4, 4) & "-" & Mid(strFileName, 8, 2) & "-" & Mid(strFileName, 10, 2), "yyyy-mm-dd")) Then
20                    dFileCreatDate = CDate(Format(Mid(strFileName, 4, 4) & "-" & Mid(strFileName, 8, 2) & "-" & Mid(strFileName, 10, 2), "yyyy-mm-dd"))
21                    If DateDiff("d", dFileCreatDate, dTimeNow) > 30 Then
22                        objFSO.DeleteFile (objFile.Path)
23                    End If
24                End If
25            Next

              '创建日志文件之前先检查是否有足够的磁盘空间
26            strDrive = objFSO.GetDriveName(objFSO.GetAbsolutePathName(strFolder))    '获取日志文件所在的磁盘名称
27            Set DrDrive = objFSO.GetDrive(strDrive)    '获取磁盘对象
28            If DrDrive.IsReady Then    '检查该磁盘是否可用
29                If DrDrive.FreeSpace < 52428800 * 2 Then    '如果磁盘剩余空间小于100M，则禁止创建日志文件
30                    MsgBox Mid(strDrive, 1, 1) & "盘空间不足，无法保存日志", vbInformation, gSysInfo.AppName
31                    Exit Function
32                End If
33            End If
34            objFSO.CreateTextFile (strFilePath)    '创建日志文件
35        Else
              '如果文件已经存在则检查文件大小，文件太大影响读取效率，这里限定为50M，大于50M则生成备份文件
36            Set objFile = objFSO.GetFile(strFilePath)
37            If objFile.Size > 52428800 Then    '判断文件是否大于50M
38                Name strFilePath As strFolder & "\LIS" & Format(dTimeNow, "yyyymmddhhmm") & ".bak"   '修改文件名
                  '当文件被重命名之后，需要重新检查并创建文件
39                Call WriteErrLog(strObject, strModule, strErrTxt)
40                Exit Function
41            End If
42        End If

          '写入错误日志
43        On Error Resume Next
44        Set objTxt = objFSO.OpenTextFile(strFilePath, ForAppending, False)
45        objTxt.WriteLine "====================" & strObject & ":" & strModule & " " & dTimeNow & "=================="    '标记是哪个部件那个模块出的错
46        objTxt.WriteLine strErrTxt
47        objTxt.Close
48        Set objTxt = Nothing

49        On Error GoTo Errhand

50        WriteErrLog = strErrTxt

          '存储过程中弹出的提示
51        If InStr(UCase(strErrTxt), "[ZLSOFT]") > 0 Then
52            strOrcMsg = Mid(UCase(strErrTxt), InStr(UCase(strErrTxt), "[ZLSOFT]") + 8, InStrRev(UCase(strErrTxt), "[ZLSOFT]") - InStr(UCase(strErrTxt), "[ZLSOFT]") - 8)
53            WriteErrLog = strOrcMsg
54        End If

55        If blnMsg And intType = 0 Then
56            If strOrcMsg <> "" Then
                  '存储过程中弹出的提示
57                MsgBox strOrcMsg, vbInformation, gSysInfo.AppName
58            Else
59                MsgBox strObject & ":" & strModule & ":" & strErrTxt, vbInformation, gSysInfo.AppName
60            End If
61        End If


62        Exit Function
Errhand:
63        MsgBox "写入日志出错,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, vbInformation, "提示"
64        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2020-02-12
'功    能:  通用日志输出
'入    参:
'           Oracle          连接对象
'           blnCallEnd      是否是结束时调用，True=是
'           strProcName     调用方法名
'           strTitle        标题
'           arrProcParas    参数
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Sub ExportLog(ByVal intSelDB As Integer, ByVal blnCallEnd As Boolean, ByVal strProcName As String, _
                     ByVal strTitle As String, ByVal strSQL As String, ParamArray arrProcParas() As Variant)
                     
          Const STR_CATEGORY As String = "公共日志"
          Dim strServerName As String
          Dim objLog As Object
          Dim arrPars() As Variant
          Dim Oracle As ADODB.Connection

1         On Error GoTo ExportLog_Error

2         If intSelDB = Sel_Lis_DB Then
3             Set Oracle = gcnLisOracle
4         ElseIf intSelDB = Sel_His_DB Then
5             Set Oracle = gcnHisOracle
6         End If

          '创建日志部件
7         Call CreateLog(Oracle)

8         strServerName = GetServerName(Oracle)

9         On Error Resume Next
10        Set objLog = gcollLog(strServerName)
11        Err.Clear: On Error GoTo ExportLog_Error
12        If objLog Is Nothing Then
13            Exit Sub
14        End If

15        arrPars = arrProcParas
16        If strSQL = "" Then
17            Call objLog.LogCall(STR_CATEGORY, App.ProductName, "clsDatabase", strProcName, strTitle _
                  , IIf(blnCallEnd, Val("1-结束"), Val("0-开始")), arrPars)
18        Else
19            Call objLog.LogCall(STR_CATEGORY, App.ProductName, "clsDatabase", strProcName, strTitle _
                  , IIf(blnCallEnd, Val("1-结束"), Val("0-开始")), strSQL, arrPars)
20        End If


21        Exit Sub
ExportLog_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "mdlPubIniFile", "执行(ExportLog)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
23        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2020-02-12
'功    能:  创建日志对象
'入    参:
'           Oracle      连接对象
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function CreateLog(Oracle As ADODB.Connection) As Boolean
          Dim objLog As Object
          Dim strServerName As String
          Dim i As Integer

1         On Error GoTo CreateLog_Error

2         strServerName = GetServerName(Oracle)

          '创建日志对象
3         If gcollLog.Count = 0 Then
              '第一次进来，为空，自动创建
4             On Error Resume Next
5             Set objLog = CreateObject("zlLog.clsLog")
6             Err.Clear: On Error GoTo CreateLog_Error
7             Call objLog.SetBusinessDB(Oracle)
8             gcollLog.Add objLog, strServerName
9         Else
              '检查当前连接对象是否创建了对应的日志部件，如果没有，则创建
10            On Error Resume Next
11            Set objLog = gcollLog(strServerName)
12            Err.Clear: On Error GoTo CreateLog_Error
13            If objLog Is Nothing Then
14                On Error Resume Next
15                Set objLog = CreateObject("zlLog.clsLog")
16                Err.Clear: On Error GoTo CreateLog_Error
17                Call objLog.SetBusinessDB(Oracle)
18                gcollLog.Add objLog, strServerName
19            End If
20        End If


21        Exit Function
CreateLog_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "mdlPubIniFile", "执行(CreateLog)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
23        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2020-02-12
'功    能:  通过连接对象获取服务器名
'入    参:
'           Oracle      连接对象
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function GetServerName(Oracle As ADODB.Connection) As String
          Dim strServerName As String
          Dim lngS As Long
          Dim lngE As Long

1         On Error GoTo GetServerName_Error

2         lngS = InStr(UCase(Oracle), "SOURCE=") + Len("SOURCE=")
3         lngE = InStr(lngS + 1, Oracle, ";")
4         strServerName = Mid(Oracle, lngS, lngE - lngS)
5         If InStr(strServerName, """(") > 0 Then
6             lngS = InStr(UCase(Oracle), "SERVICE_NAME=") + Len("SERVICE_NAME=")
7             lngE = InStr(lngS + 1, Oracle, ")")
8             strServerName = Mid(Oracle, lngS, lngE - lngS)
9         End If

10        GetServerName = strServerName


11        Exit Function
GetServerName_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlPubIniFile", "执行(GetServerName)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear

End Function
