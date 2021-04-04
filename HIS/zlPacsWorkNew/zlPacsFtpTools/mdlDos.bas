Attribute VB_Name = "mdlDos"
Option Explicit

'私有的数据结构申明

'延时函数部分
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type STARTUPINFO '(createprocess)
    cb As Long
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

Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Public Declare Function OpenWaitableTimer Lib "kernel32" Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Public Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
'

Public Type PROCESS_INFORMATION '(creteprocess)
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Public Type SECURITY_ATTRIBUTES '(createprocess)
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'常数声明
Public Const CREATE_NO_WINDOW = &H8000000
Public Const CREATE_NEW_CONSOLE = &H10

Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const STARTF_USESTDHANDLES = &H100&
Public Const STARTF_USESHOWWINDOW = &H1
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_QUERY_INFORMATION = &H400

Public Const ERROR_ALREADY_EXISTS = 183&
Public Const UNITS = 4294967296#
Public Const MAX_LONG = -2147483648#
Public Const INFINITE = &HFFFF

Public Const WAIT_ABANDONED& = &H80&
Public Const WAIT_ABANDONED_0& = &H80&
Public Const WAIT_FAILED& = -1&
Public Const WAIT_IO_COMPLETION& = &HC0&
Public Const WAIT_OBJECT_0& = 0
Public Const WAIT_OBJECT_1& = 1
Public Const WAIT_TIMEOUT& = &H102&

Public Const QS_HOTKEY& = &H80
Public Const QS_KEY& = &H1
Public Const QS_MOUSEBUTTON& = &H4
Public Const QS_MOUSEMOVE& = &H2
Public Const QS_PAINT& = &H20
Public Const QS_POSTMESSAGE& = &H8
Public Const QS_SENDMESSAGE& = &H40
Public Const QS_TIMER& = &H10
Public Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Public Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Public Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Public Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
'

'函数声明
Public Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, _
    lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, _
                           ByVal lpBuffer As Long, _
                           ByVal nBufferSize As Long, _
                           ByRef lpBytesRead As Long, _
                           ByRef lpTotalBytesAvail As Long, _
                           ByRef lpBytesLeftThisMessage As Long _
                           ) As Long

Public Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As Long, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Public Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hHandle As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, _
                                                   ByVal lpBuffer As Long, _
                                                   ByVal nNumberOfBytesToWrite As Long, _
                                                   ByRef lpNumberOfBytesWritten As Long, _
                                                   lpOverlapped As Any) As Long

Public Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, _
                ByVal hSourceHandle As Long, _
                ByVal hTargetProcessHandle As Long, _
                lpTargetHandle As Long, _
                ByVal dwDesiredAccess As Long, _
                ByVal bInheritHandle As Long, _
                ByVal dwOptions As Long) As Long



Public Const DUPLICATE_SAME_ACCESS = &H2



Public Enum InitResult
  ERROR_OK = 0
  ERROR_INIT_INPUT_HANDLE = 1
  ERROR_INIT_OUTPUT_HANDLE = 2
  ERROR_DUP_READ_HANDLE = 3
  ERROR_DUP_WRITE_HANDLE = 4
  ERROR_CREATE_CHILD_PROCESS = 5
  ERROR_QUERY_INFO_SIZE = 6
End Enum

Public Enum TermResult
 ERROR_OK = 0
End Enum

Public Enum InputResult
 ERROR_OK = 0
 ERROR_QUERY_WRITE_INFO_SIZE = 1
 ERROR_DATA_TO_LARGE = 2
 ERROR_WRITE_INFO = 3
 ERROR_WRITE_UNEXPECTED = 5
End Enum

Public Enum OutputResult
 ERROR_OK = 0
 ERROR_QUERY_READ_INFO_SIZE = 1
 ERROR_ZERO_INFO_SIZE = 2
 ERROR_READ_INFO = 3
 ERROR_UNEQUAL_INFO_SIZE = 4
 ERROR_READ_UNEXPECTED = 5
End Enum


Public Function Unicode8Encode(bTemp As String) As Byte()
'编码UNICODE UTF-8
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strTotal() As Byte
    Dim strTmp As String
    Dim Code As Long
    Dim Code1 As Long
    Dim Code2 As Long
    Dim Code3 As Long
    Dim Code4 As Long
    Dim Code5 As Long
    Dim Code6 As Long  '已生成的字节数
    Dim bNo As Long
    
    k = Len(bTemp)
    bNo = 0
    
    ReDim strTotal(k * 3)
    For i = 1 To k
        Code = CLng("&H" + Hex(AscW(Mid(bTemp, i, 1))))
        If Code < 128& Then
            strTotal(bNo) = Code
            bNo = bNo + 1
            If bNo > 422386 Then
                Debug.Print Code
            End If
        ElseIf Code < 2048& Then
            Code1 = ((Code And 1984&) \ 32&) + 192
            Code2 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            bNo = bNo + 2
        ElseIf Code < 65536 Then
            Code1 = ((Code And 61440) \ 4096&) + 224
            Code2 = ((Code And 4032&) \ 64&) + 128
            Code3 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            bNo = bNo + 3
        ElseIf Code < 2097152 Then
            Code1 = ((Code And 1835008) \ 262144) + 240
            Code2 = ((Code And 258048) \ 4096&) + 128
            Code3 = ((Code And 4032&) \ 64&) + 128
            Code4 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            bNo = bNo + 4
        ElseIf Code < 67108864 Then
            Code1 = ((Code And 50331648) \ 16777216) + 248
            Code2 = ((Code And 16515072) \ 262144) + 128
            Code3 = ((Code And 258048) \ 4096&) + 128
            Code4 = ((Code And 4032&) \ 64&) + 128
            Code5 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            strTotal(bNo + 4) = Code5
            bNo = bNo + 5
        Else
            Code1 = IIf(Code And 1073741824 = 1073741824, 253&, 252&)
            Code2 = ((Code And 1056964608) \ 16777216) + 128
            Code3 = ((Code And 16515072) \ 262144) + 128
            Code4 = ((Code And 258048) \ 4096&) + 128
            Code5 = ((Code And 4032&) \ 64&) + 128
            Code6 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            strTotal(bNo + 4) = Code5
            strTotal(bNo + 5) = Code6
            bNo = bNo + 6
        End If
    Next
    
    ReDim Preserve strTotal(bNo - 1)
    Unicode8Encode = strTotal
End Function

