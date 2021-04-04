Attribute VB_Name = "mdlCMDControl"
Option Explicit

'**************************
'功能:管道执行CMD命令并截取输出
'编写整理:祝庆
'**************************

Function GetCmdTxt(Command As String, Optional opo As Boolean)
    Dim Proc As PROCESS_INFORMATION '进程信息
    Dim Start As STARTUPINFO '启动信息
    Dim SecAttr As SECURITY_ATTRIBUTES '安全属性
    Dim hReadPipe As Long '读取管道句柄
    Dim hWritePipe As Long '写入管道句柄
    Dim lngBytesRead As Long '读出数据的字节数
    Dim strBuffer As String * 256 '读取管道的字符串buffer
    Dim i As Integer
    On Error Resume Next
    Dim ret As Long 'API函数返回值
    Dim retPro As Long
    Dim lpOutputs As String '读出的最终结果
    
    '设置安全属性
    With SecAttr
        .nLength = LenB(SecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With

    '创建管道
    ret = CreatePipe(hReadPipe, hWritePipe, SecAttr, 0)
    If ret = 0 Then
        MsgBox "无法创建管道", vbExclamation, "错误"
        Exit Function
    End If

    '设置进程启动前的信息
    With Start
        .cb = LenB(Start)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .hStdOutput = hWritePipe '设置输出管道
        .hStdError = hWritePipe '设置错误管道
    End With

    '启动进程
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS进程以ipconfig.exe为例
    'Command = "Rasdial adsl CD02887573165 87573165"
    
    retPro = CreateProcess(vbNullString, Command, SecAttr, SecAttr, True, NORMAL_PRIORITY_CLASS, ByVal 0, vbNullString, Start, Proc)
    If ret = 0 Then
        MsgBox "无法启动新进程", vbExclamation, "错误"
        ret = CloseHandle(hWritePipe)
        ret = CloseHandle(hReadPipe)
        Exit Function
    End If
    
    '因为无需写入数据，所以先关闭写入管道。而且这里必须关闭此管道，否则将无法读取数据
    ret = CloseHandle(hWritePipe)

    '从输出管道读取数据，每次最多读取256字节
    Do
        ret = ReadFile(hReadPipe, strBuffer, 256, lngBytesRead, ByVal 0)
        lpOutputs = lpOutputs & Left(strBuffer, lngBytesRead)
        DoEvents
    Loop While (ret <> 0) '当ret=0时说明ReadFile执行失败，已经没有数据可读了

    '读取操作完成，关闭各句柄
    ret = CloseHandle(retPro)
    ret = CloseHandle(Proc.hProcess)
    ret = CloseHandle(Proc.hThread)
    ret = CloseHandle(hReadPipe)
End Function
