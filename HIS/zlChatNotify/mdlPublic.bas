Attribute VB_Name = "mdlPublic"
Option Explicit

Public gstrSysName As String
Public gstrUser As String
Public gstrChatURL As String        '发起讨论地址
Public gstrMyChatUrl As String        '我参与的讨论地址
Public gobjMain As Object           '导航台对象  通过此对象操作数据库
 
Public gfrmMain As frmMain

Public grsList  As ADODB.Recordset  '待读消息
Public gcolChat As Collection       '记录打开的讨论
Public gblnLog  As Boolean              '
Public gblnShow As Boolean          '用于记录等待窗体是否打开

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum E_Notify_Type   '0-初始化  1-消息 2-闪烁 3-还原
    E_初始化 = 0
    E_消息 = 1
    E_闪烁 = 2
    E_还原 = 3
End Enum
'----------------------------------------------------------------------------------------------------
'-----系统托盘相关声明
'----------------------------------------------------------------------------------------------------
Public Const MAX_TOOLTIP As Integer = 64
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEWHEEL = &H20A          '鼠标滚动
Public Const SW_RESTORE = 9
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOACTIVATE = &H10 '不激活窗体
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE  As Long = (-20)
Public Const conCOLOR_BULELIGHT As Long = &HE4B440
Public Const conCOLOR_BULE As Long = &HD48A00
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
'----------------------------------------------------------
'-------颜色-常量
'---------------------------------------------------------------
Public Const conCOLOR_TITLE_BAR As Long = 16298544 '16298544 rgb(48,178,248); 14392064 'RGB(0, 155, 219)


Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type


Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'获得鼠标指针在屏幕坐标上的位置
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'获得窗口在屏幕坐标中的位置
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'判断指定的点是否在指定的矩形内部
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
'用来使窗体始终在最前面
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'获取窗体状态
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'读
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'写
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
'返回值:非零表示成功，零表示失败。会设置GetLastError

Private Const CP_UTF8 = 65001
'Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
    
    
Private Const CON_SPLIT As String = ";"
Private mobjFso As New FileSystemObject         '文件对象
    
Public Sub InitRsList()
    Set grsList = New ADODB.Recordset
    With grsList
        .Fields.Append "ID", adBigInt
        .Fields.Append "Url", adVarChar, 500
        .Fields.Append "Sys_Code", adVarChar, 20
        .Fields.Append "Main_Code", adVarChar, 20
        .Fields.Append "Main_ID", adVarChar, 36
        .Fields.Append "Subject", adVarChar, 50
        
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub
 
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '先计算需求字节数
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    '然后转换
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function

Public Function URLEncode(ByVal strParameter As String, Optional strEncodeType As String = "utf8") As String
          Dim strTemp As String
          Dim strRet As String
          Dim strInput As String
          
          Dim i As Long
          Dim lngValue As Long
          Dim lngLen As Long
          Dim lngMax As Long
          
          Dim bytData() As Byte

10        On Error GoTo ErrH
20        lngLen = 32767
30        Do While Len(strParameter) > 0
40            lngMax = Len(strParameter)
50            If lngMax > lngLen Then
60                strInput = Mid(strParameter, 1, lngLen)
70                strParameter = Mid(strParameter, lngLen + 1, lngMax - lngLen)
80            Else
90                strInput = strParameter
100               strParameter = ""
110           End If
120           strTemp = ""
130           If "UTF8" = UCase(strEncodeType) Then
140               bytData = StringToUTF8Bytes(strInput)
150           Else
160               bytData = StrConv(strInput, vbFromUnicode)
170           End If
              
180           For i = 0 To UBound(bytData)
190               lngValue = bytData(i)
200               If (lngValue >= 48 And lngValue <= 57) Or _
                      (lngValue >= 65 And lngValue <= 90) Or _
                      (lngValue >= 97 And lngValue <= 122) Or _
                       InStr("$-_.+*'()", Chr(lngValue)) > 0 Then
                       '特殊字符不转"$-_.+*'()"
210                   strTemp = strTemp & Chr(lngValue)
220               ElseIf lngValue = 32 Then
                      '空格
230                   strTemp = strTemp & "+"
240               Else
250                   If lngValue <= 15 Then
260                       strTemp = strTemp & "%0" & UCase(Hex(lngValue))
270                   Else
280                       strTemp = strTemp & "%" & UCase(Hex(lngValue))
290                   End If
300               End If
310           Next
320           strRet = strRet & strTemp
330       Loop
340       URLEncode = strRet
350       Exit Function
ErrH:
360
     WriteLog "在zlChatNotify.mdlPublic.URLEncode的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description & vbNewLine
End Function

Public Function OpenChatRoom(ByVal strUrl As String, ByVal strSubject As String, Optional ByVal strSysCode As String, _
    Optional ByVal strMainCode As String, Optional ByVal dblMainId As Double, Optional ByVal strSender As String, _
    Optional ByVal strReceivers As String, Optional ByRef strMsg As String) As Boolean
    '参数:
    'strSubject         -讨论标题
    'strSysCode         -系统编码
    'strMainCode        -主体编码
    'dblMainId          -主体ID
    'strSender          -发起人
    'strReceivers       -参与人(多个参与人用分隔符";"分开)
    'strMsg             -返回错误信息(避免弹出模态提示引起主进程挂起,将提示信息返回给主进程处理。)
    '                    返回格式:提示类型[,]提示语句
          Dim strKey As String
          Dim objChat As frmChat
          
1         On Error GoTo ErrH
2         WriteLog "函数：OpenChatRoom 开始" & vbNewLine & _
                         "入参：url=" & strUrl & vbNewLine & _
                         "讨论标题:" & strSubject & vbNewLine & _
                         "系统编码:" & strSysCode & vbNewLine & _
                         "主体编码:" & strMainCode & vbNewLine & _
                         "主体ID:" & dblMainId & vbNewLine & _
                         "发起人:" & strSender & vbNewLine & _
                         "接收者:" & strReceivers & vbNewLine
3         strKey = strSysCode & "_" & strMainCode & "_" & dblMainId
4         On Error Resume Next
5         Set objChat = gcolChat(strKey)
6         On Error GoTo ErrH
7         If objChat Is Nothing Then
8             Set objChat = New frmChat
9             gcolChat.Add objChat, strKey
10        End If
11        OpenChatRoom = objChat.OpenChatRoom(strUrl, strSubject, strSysCode, strMainCode, dblMainId, strSender, strReceivers, strMsg)
12        WriteLog "函数：OpenChatRoom 结束"
13        Exit Function
ErrH:
14        strMsg = vbExclamation & "[,]" & "在zlChatNotify.mdlPublic.OpenChatRoom的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description
15        WriteLog strMsg & vbNewLine
End Function

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

    Dim strLogPath As String, strLogFile  As String    '日志路径，文件名，配置文件名
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
    
    If mobjFso.FolderExists(strPath) Then
        strDriv = mobjFso.GetDriveName(mobjFso.GetAbsolutePathName(strPath))
        Set Drv = mobjFso.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function

Public Sub SetFormTranslucency(hWnd As Long, crKey As Long, bAlpha As Byte, dwFlags As Long) '实现半透明窗体
'功能:设置窗体透明度
'hwnd,  窗口句柄
'crKey:指定需要透明的背景颜色值，可用RGB()宏
'bAlpha:设置透明度，0表示完全透明，255表示不透明
'dwFlags: 透明方式dwFlags参数可取以下值：
'       LWA_ALPHA=&H2时：crKey参数无效，bAlpha参数有效；
'       LWA_COLORKEY=&H1：窗体中的所有颜色为crKey的地方将变为透明，bAlpha参数无效。其常量值为1。
'       LWA_ALPHA | LWA_COLORKEY：crKey的地方将变为全透明，而其它地方根据bAlpha参数确定透明度。
   Dim lngRet As Long
   
    lngRet = GetWindowLong(hWnd, GWL_EXSTYLE)
    lngRet = lngRet Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, lngRet
    SetLayeredWindowAttributes hWnd, crKey, bAlpha, dwFlags
 End Sub

Public Sub SetWindowsInTaskBar(ByVal lnghwnd As Long, ByVal blnShow As Boolean)
'功能：设置窗体是否在任务条上显示
    Dim lngStyle As Long
    
    lngStyle = GetWindowLong(lnghwnd, GWL_EXSTYLE)
    If blnShow Then
        lngStyle = lngStyle Or &H40000
    Else
        lngStyle = lngStyle And Not &H40000
    End If
    Call SetWindowLong(lnghwnd, GWL_EXSTYLE, lngStyle)
End Sub

