VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessID As Long
    dwThreadID  As Long
End Type

Private Type STARTUPINFO
    Cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const NORMAL_PRIORITY_CLASS             As Long = &H20&
Private Const STARTF_USESTDHANDLES              As Long = &H100&
Private Const STARTF_USESHOWWINDOW              As Long = &H1
Private Const SW_HIDE                           As Integer = 0 '隐藏窗口，激活另一个窗口
Private Const INFINITE                          As Long = &HFFFF&

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'当等待仍在挂起状态时，句柄被关闭，那么函数行为是未定义的。该句柄必须具有 SYNCHRONIZE 访问权限。
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Public Function Run(ByVal strCommand As String, Optional ByRef strOutput As String, Optional ByRef strErr As String, Optional ByVal lngWait As Long = INFINITE) As Boolean
'功能：执行命令行，并获取命令行输出
    Dim piProc          As PROCESS_INFORMATION '进程信息
    Dim stStart         As STARTUPINFO '启动信息
    Dim saSecAttr       As SECURITY_ATTRIBUTES '安全属性
    Dim lnghReadPipe    As Long '读取管道句柄
    Dim lnghWritePipe   As Long '写入管道句柄
    Dim lngBytesRead    As Long '读出数据的字节数
    Dim strBuffer       As String * 256 '读取管道的字符串buffer
    Dim lngRet          As Long 'API函数返回值
    Dim lngRetPro       As Long
    Dim strlpOutputs    As String '读出的最终结果
    
    strErr = "": strOutput = ""
    On Error Resume Next
    '设置安全属性
    With saSecAttr
        .nLength = LenB(saSecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With
    
    '创建管道
    lngRet = CreatePipe(lnghReadPipe, lnghWritePipe, saSecAttr, 0)
    If lngRet = 0 Then
        strErr = "无法创建管道。" & GetLastDllErr()
        Exit Function
    End If
    '设置进程启动前的信息
    With stStart
        .Cb = LenB(stStart)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = lnghWritePipe '设置输出管道
        .hStdError = lnghWritePipe '设置错误管道
    End With
    '启动进程
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS进程以ipconfig.exe为例
    lngRetPro = CreateProcess(vbNullString, strCommand & vbNullChar, saSecAttr, saSecAttr, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, stStart, piProc)
    If lngRetPro = 0 Then
        strErr = "无法启动进程。" & GetLastDllErr()
        lngRet = CloseHandle(lnghWritePipe)
        lngRet = CloseHandle(lnghReadPipe)
        Exit Function
    Else
        '因为无需写入数据，所以先关闭写入管道。而且这里必须关闭此管道，否则将无法读取数据
        lngRet = CloseHandle(lnghWritePipe)
        WaitForSingleObject piProc.hProcess, lngWait
        Do
            lngRet = ReadFile(lnghReadPipe, strBuffer, 256, lngBytesRead, ByVal 0)
            If lngRet <> 0 Then
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            Else
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            End If
            DoEvents
        Loop While (lngRet <> 0) '当ret=0时说明ReadFile执行失败，已经没有数据可读了
        '读取操作完成，关闭各句柄
        lngRet = CloseHandle(lngRetPro)
        lngRet = CloseHandle(piProc.hProcess)
        lngRet = CloseHandle(piProc.hThread)
        lngRet = CloseHandle(lnghReadPipe)
    End If
    strOutput = Replace(strlpOutputs, vbNullChar, "")
    Run = True
End Function

