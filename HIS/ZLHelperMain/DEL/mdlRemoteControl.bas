Attribute VB_Name = "mdlRemoteControl"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/5/9
'模块           mdlRemoteControl
'说明
'==================================================================================================
'管道获取CMD输出
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'用来创建一个新的进程和它的主线程，这个新进程运行指定的可执行文件。如果函数执行成功，返回非零值
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'创建一个匿名管道，并从中得到读写管道的句柄。如果函数执行成功，返回非零值
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
'从文件指针指向的位置开始将数据读出到一个文件中， 且支持同步和异步操作
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'当等待仍在挂起状态时，句柄被关闭，那么函数行为是未定义的。该句柄必须具有 SYNCHRONIZE 访问权限。
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Type STARTUPINFO
    cb                              As Long
    lpReserved                      As String
    lpDesktop                       As String
    lpTitle                         As String
    dwX                             As Long
    dwY                             As Long
    dwXSize                         As Long
    dwYSize                         As Long
    dwXCountChars                   As Long
    dwYCountChars                   As Long
    dwFillAttribute                 As Long
    dwFlags                         As Long
    wShowWindow                     As Integer
    cbReserved2                     As Integer
    lpReserved2                     As Long
    hStdInput                       As Long
    hStdOutput                      As Long
    hStdError                       As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess                        As Long
    hThread                         As Long
    dwProcessId                     As Long
    dwThreadId                      As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength                         As Long
    lpSecurityDescriptor            As Long
    bInheritHandle                  As Long
End Type
Private Const NORMAL_PRIORITY_CLASS  As Long = &H20&
Private Const STARTF_USESTDHANDLES   As Long = &H100&
Private Const STARTF_USESHOWWINDOW   As Long = &H1&
Private Const INFINITE               As Long = &HFFFF&
Public Const SW_HIDE = 0 '隐藏窗口，激活另一个窗口
Public Function RunCommand(commandline As String) As String
    Dim si As STARTUPINFO                                                       'used to send info the CreateProcess
    Dim pi As PROCESS_INFORMATION                                               'used to receive info about the created process
    Dim retval As Long                                                          'return value
    Dim hRead As Long                                                           'the handle to the read end of the pipe
    Dim hWrite As Long                                                          'the handle to the write end of the pipe
    Dim sBuffer(0 To 63) As Byte                                                'the buffer to store data as we read it from the pipe
    Dim lgSize As Long                                                          'returned number of bytes read by readfile
    Dim sa As SECURITY_ATTRIBUTES
    Dim strResult As String                                                     'returned results of the command line
    
    'set up security attributes structure
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1&                                                    'inherit, needed for this to work
        .lpSecurityDescriptor = 0&
    End With
    'create our anonymous pipe an check for success
    ' note we use the default buffer size
    ' this could cause problems if the process tries to write more than this buffer size
    retval = CreatePipe(hRead, hWrite, sa, 0&)
    If retval = 0 Then
'        MsgBox "错误提示:创建管道失败!"
        RunCommand = ""
        Exit Function
    End If
    'set up startup info
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW                 'tell it to use (not ignore) the values below
        .wShowWindow = SW_HIDE
        .hStdOutput = hWrite                                                    'pass the write end of the pipe as the processes standard output
    End With
    'run the command line and check for success
    retval = CreateProcess(vbNullString, commandline & vbNullChar, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, si, pi)
    If retval Then
        'wait until the command line finishes
        ' trouble if the app doesn't end, or waits for user input, etc
        WaitForSingleObject pi.hProcess, INFINITE
        'read from the pipe until there's no more (bytes actually read is less than what we told it to)
        Do While ReadFile(hRead, sBuffer(0), 64, lgSize, ByVal 0&)
            'convert byte array to string and append to our result
            strResult = strResult & StrConv(sBuffer(), vbUnicode)
            'TODO = what's in the tail end of the byte array when lgSize is less than 64???
            Erase sBuffer()
            If lgSize <> 64 Then Exit Do
            DoEvents
        Loop
        'close the handles of the process
        CloseHandle pi.hProcess
        CloseHandle pi.hThread
    Else
'        MsgBox "错误提示:创建进程失败!" & vbCrLf
        RunCommand = ""
        Exit Function
    End If
    'close pipe handles
    CloseHandle hRead
    CloseHandle hWrite
    'return the command line output
    RunCommand = Replace(strResult, vbNullChar, "")
End Function
