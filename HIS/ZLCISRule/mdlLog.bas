Attribute VB_Name = "mdlLog"
Option Explicit

'读
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'写
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
'返回值:非零表示成功，零表示失败。会设置GetLastError

Private Const CON_SPLIT As String = ";"
Private mobjFso As New FileSystemObject         '文件对象

Public Function ReadIni(ByVal strNodeName As String, ByVal strKeyName As String, strFilePath As String) As String
    Dim strBuff As String
    Dim strReadStr As String
    Dim lngPos As Long
    
    On Error GoTo ErrH

    strBuff = VBA.String(255, 0)
    GetPrivateProfileString strNodeName, strKeyName, "", strBuff, 256, strFilePath
    strReadStr = VBA.Replace(strBuff, VBA.Chr(0), "")
    
    lngPos = InStr(1, strReadStr, CON_SPLIT, vbTextCompare)     '找到 ;的位置(结束标志)
    If lngPos >= 1 Then
        ReadIni = Trim(Left(strReadStr, lngPos - 1))
    Else
       '如果没有找到 有注释的标志
       ReadIni = strReadStr
    End If
    
    Exit Function
ErrH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(ByVal strNodeName As String, ByVal strKeyName As String, ByVal strValue As String, ByVal strFilePath As String) As Long
    Dim strBuff As String
    Dim strComment As String
    Dim strReadStr As String
    
    Dim lngRet As Long
    Dim lngPos As Long
    On Error GoTo ErrH
   strBuff = String(255, 0)
   lngRet = GetPrivateProfileString(strNodeName, strKeyName, "", strBuff, 256, strFilePath)
   strReadStr = VBA.Replace(strBuff, VBA.Chr(0), "")
   lngPos = InStr(1, strReadStr, CON_SPLIT, vbTextCompare)    '找到 ;的位置(结束标志)
   '如果有;取出其后的注释
   If lngPos >= 1 Then
      strComment = Trim(Right(strReadStr, lngRet - lngPos))
      strValue = strValue & strComment
   End If
   
    WriteIni = WritePrivateProfileString(strNodeName, strKeyName, strValue, strFilePath)
        
    Exit Function
ErrH:
    Err.Clear
End Function

Public Sub WriteLog(ByVal strLogTxt As String)
    '写一行日志，如果内容中有回车,换行符，替换为<CR><LF>
    '日志保存在当前目录下的[应用程序名称]Log目录下，文件名为日期.txt,默认保存7天的日志。

    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '日志路径，文件名，配置文件名
    Dim strLogSaveDays As String '日志保留天数
    Dim dblFreeSpace As Double   '剩余空间
    Dim strDelOldFile As String  '过期文件
    Dim objFile As File
    
    '是否开启日志
    If Not gblnLog Then Exit Sub
     
    '始终保存日志
    '2、清除过期日志
    strLogSaveDays = "7"  '保留7天的日志
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\" & App.EXEName & "*.log")
    Do While strDelOldFile <> ""
        Set objFile = mobjFso.GetFile(strLogPath & "\" & strDelOldFile)
        If DateDiff("d", objFile.DateLastModified, Now) > Val(strLogSaveDays) Then
            mobjFso.DeleteFile strLogPath & "\" & strDelOldFile, True
        End If
        strDelOldFile = Dir
    Loop
    
    '3、空间是否足够
    dblFreeSpace = GetFreeSpace(strLogPath)
    If dblFreeSpace >= 1024 And dblFreeSpace <= 10240 Then
        '空间不足，不写日志,产生一个警告文件
        If Not mobjFso.FileExists(strLogPath & "\空间不足.txt") Then Call mobjFso.CreateTextFile(strLogPath & "\空间不足.txt", True)
        Exit Sub
    Else
        '清除警告文件
        If mobjFso.FileExists(strLogPath & "\空间不足.txt") Then Call mobjFso.DeleteFile(strLogPath & "\空间不足.txt", True)
    End If
    '4、写入日志行
    strLogFile = strLogPath & "\" & App.EXEName & Format(Now, "yyyyMMdd") & ".log"
    Call SaveLog(strLogFile, strLogTxt)

End Sub

Private Sub SaveLog(ByVal strFilename As String, ByVal strInput As String, Optional ByVal strDate As String)
 
    Dim objStream As TextStream
    Dim strWritLing As String
    
    strWritLing = Replace$(strInput, Chr(&HD), "<CR>")
    strWritLing = Replace$(strInput, Chr(&HA), "<LF>")

    If strInput <> "" Then
        If Not mobjFso.FileExists(strFilename) Then Call mobjFso.CreateTextFile(strFilename)
        Set objStream = mobjFso.OpenTextFile(strFilename, ForAppending)
        If strDate = "" Then strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine (strDate & Chr(&H9) & strInput)
        objStream.Close
        Set objStream = Nothing
    End If
    
End Sub

Private Function GetFreeSpace(ByVal strPath As String) As Double
    '获取剩余空间
    Dim strDriv As String, Drv As Drive
    Dim strDir As String
    
    If mobjFso.FolderExists(strPath) Then
        strDriv = mobjFso.GetDriveName(mobjFso.GetAbsolutePathName(strPath))
        Set Drv = mobjFso.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function
