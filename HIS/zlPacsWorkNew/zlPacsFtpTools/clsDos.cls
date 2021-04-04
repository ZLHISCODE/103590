VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlTimer As Long

Private PipeR4InputChannel As Long, PipeW4InputChannel As Long, hInputHandle As Long
Private PipeR4OutputChannel As Long, PipeW4OutputChannel As Long, hOutputHandle As Long


Private Declare Function timeGetTime Lib "winmm.dll" () As Long


Private Proc As PROCESS_INFORMATION
Private OutPutData As String, ReadData As Boolean, OutData As String
Private mblnIsAbort As Boolean

Public Event Events()


Property Get IsAbort() As Boolean
    IsAbort = mblnIsAbort
End Property

Public Sub Wait(MilliSeconds As Long) '延时函数
    On Error GoTo errHandler
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double

    mlTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer" & Format$(Now(), "NNSS"))

    If Err.LastDllError <> ERROR_ALREADY_EXISTS Then
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, 0)
    End If

    dblDelay = CDbl(MilliSeconds) * 10000#

    ft.dwHighDateTime = -CLng(dblDelay / UNITS) - 1
    dblDelayLow = -UNITS * (dblDelay / UNITS - Fix(CStr(dblDelay / UNITS)))

    If dblDelayLow < MAX_LONG Then dblDelayLow = UNITS + dblDelayLow

    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, False)

    Do
        If mblnIsAbort Then Exit Sub
        
        lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
        DoEvents
    Loop Until lBusy = WAIT_OBJECT_0

    CloseHandle mlTimer
    
    mlTimer = 0
    Exit Sub

errHandler:
    Err.Raise Err.Number, Err.Source, "[clsWaitableTimer.Wait]" & Err.Description
End Sub

'========================================================================================

Private Function InitDosIO() As InitResult
   Dim Sa As SECURITY_ATTRIBUTES, Ret As Long

   With Sa
      .nLength = Len(Sa)
      .bInheritHandle = 1&
      .lpSecurityDescriptor = 0&
   End With

    Ret = CreatePipe(PipeR4InputChannel, PipeW4InputChannel, Sa, 1024&)
    If Ret = 0 Then '建立进程输入管道
        InitDosIO = ERROR_INIT_INPUT_HANDLE
        Exit Function
    End If

    Ret = CreatePipe(PipeR4OutputChannel, PipeW4OutputChannel, Sa, 4096&) '建立输出通道,若建立失败，则关闭管道，退出
    If Ret = 0 Then '建立进程的输出管道
        CloseHandle PipeR4InputChannel
        CloseHandle PipeW4InputChannel
        InitDosIO = ERROR_INIT_OUTPUT_HANDLE
        Exit Function
    End If


   Ret = DuplicateHandle(GetCurrentProcess(), PipeW4InputChannel, GetCurrentProcess(), hInputHandle, 0, True, DUPLICATE_SAME_ACCESS)
   If Ret = 0 Then '转换写句柄
     CloseHandle PipeR4InputChannel
     CloseHandle PipeW4InputChannel
     CloseHandle PipeR4OutputChannel
     CloseHandle PipeW4OutputChannel
     InitDosIO = ERROR_DUP_WRITE_HANDLE
     Exit Function
   End If
   Ret = CloseHandle(PipeW4InputChannel)
   If Ret = 0 Then
    'MsgBox "close handle eerr"
   End If
  Ret = DuplicateHandle(GetCurrentProcess(), PipeR4OutputChannel, GetCurrentProcess(), hOutputHandle, 0, True, DUPLICATE_SAME_ACCESS)
   If Ret = 0 Then '转换读句柄
     CloseHandle PipeR4InputChannel
     CloseHandle PipeW4InputChannel
     CloseHandle PipeR4OutputChannel
     CloseHandle PipeW4OutputChannel
     InitDosIO = ERROR_DUP_READ_HANDLE
     Exit Function
   End If
 Ret = CloseHandle(PipeR4OutputChannel)
 If Ret = 0 Then
  'MsgBox "close handle 2 er"
 End If


    Dim Start As STARTUPINFO, CmdStr As String
    Start.cb = Len(Start)
    Start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    Start.hStdOutput = PipeW4OutputChannel
    Start.hStdError = PipeW4OutputChannel
    Start.hStdInput = PipeR4InputChannel
    CmdStr = "CMD" '需要打开的命令行外壳(CMD.EXE)
    Ret& = CreateProcessA(0&, CmdStr, Sa, Sa, True, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, Proc)

    If Ret <> 1 Then '建立控制进程
        CloseHandle PipeR4InputChannel
        CloseHandle PipeW4InputChannel
        CloseHandle PipeR4OutputChannel
        CloseHandle PipeW4OutputChannel
        InitDosIO = ERROR_CREATE_CHILD_PROCESS
        Exit Function
    End If
    ReadData = True
End Function

Public Function DosInput(ByVal Str As String) As InputResult
 Dim Buflen As Long, BtWritten As Long, Rtn As Long
 Dim BtTest() As Byte
 
' If mblnIsAbort Then Exit Function
 
 ReadData = True

 BtTest = StrConv(Str & vbCrLf, vbFromUnicode) 'Unicode8Encode(str & vbCrLf) '
 Buflen = UBound(BtTest) + 1

 Rtn = WriteFile(hInputHandle, StrPtr(BtTest), Buflen, BtWritten, ByVal 0&)

 If BtWritten = 0 Then
   DosInput = ERROR_WRITE_INFO
   Exit Function
 End If
 DosInput = 0
End Function

Public Function DosOutPutEx(Optional ByVal lngTimeOut As Long = 30000, Optional ByVal lngMinTime As Long = 100) As String '默认命令超时10秒
    Dim OutPt As String, OldTime As Long, NewTime As Long, Tmp() As String
    
    If ReadData = False Then Exit Function
    If mblnIsAbort Then
        DosOutPutEx = "强制结束!"
        Exit Function
    End If
    
    OldTime = timeGetTime
    OutPutData = ""
    
    Do
        If mblnIsAbort Then
            DosOutPutEx = "强制结束!"
            Exit Function
        End If
        
        DoEvents
        If DosOutput(OutPt) = 0 Then
            OutData = Left$(OutPt, Len(OutPt) - 1)
            OutPutData = OutPutData & Left$(OutPt, Len(OutPt) - 1)
            Call NowEvents
            Tmp = Split(OutPt, Chr$(13))
            If Len(Tmp(UBound(Tmp))) <> 0 Then
                If InStr(Tmp(UBound(Tmp)), ":\") = 3 And Right(Tmp(UBound(Tmp)), 2) = ">" & Chr$(0) Then Exit Do
            End If
            
            If Trim(Replace(OutPutData, vbCrLf, "")) <> "" Then
                lngTimeOut = lngMinTime
            Else
                OutPutData = ""
            End If
        End If
        
        NewTime = timeGetTime
        
        Call Wait(10) '如需要更高实时性可以改为1
        
        If NewTime - OldTime >= lngTimeOut Then
            If lngTimeOut = 30000 Then OutPutData = "********** 超时，未获取到有效的反馈信息！ **********"
            Exit Do
        End If
    Loop
    
    DosOutPutEx = OutPutData
    ReadData = False
End Function

Public Function GetOut() As String '实时数据
    GetOut = OutData
End Function

Friend Sub NowEvents() '接口
    RaiseEvent Events
End Sub

Private Function DosOutput(ByRef StrOutput As String) As OutputResult
  Dim Ret As Long, TmpBuf As String * 128, BtRead As Long, BtTotal As Long, BtLeft As Long
  Dim Rtn As Long, lngbytesread As Long

  Rtn = PeekNamedPipe(hOutputHandle, StrPtr(TmpBuf), 128, BtRead, BtTotal, BtLeft)
  If Rtn = 0 Then '查询信息量
    DosOutput = ERROR_QUERY_INFO_SIZE
    Exit Function
  End If

  If BtTotal = 0 Then '若信息为空，则退出
    DosOutput = ERROR_ZERO_INFO_SIZE
    Exit Function
  End If

  Dim Btbuf() As Byte, BtReaded As Long
  ReDim Btbuf(BtTotal)
  Ret = ReadFile(hOutputHandle, VarPtr(Btbuf(0)), BtTotal, lngbytesread, 0&)

  If Ret = 0 Then
    DosOutput = ERROR_READ_INFO
    Exit Function
  End If
  If BtTotal <> lngbytesread Then
   DosOutput = ERROR_UNEQUAL_INFO_SIZE
  End If

 Dim strBuf As String
 strBuf = StrConv(Btbuf, vbUnicode)
 'Debug.Print strBuf
 StrOutput = strBuf

End Function

Private Function EndDosIo() As Long
 Dim Ret As Long
 
 CloseHandle PipeR4InputChannel
 CloseHandle PipeW4InputChannel
 CloseHandle PipeR4OutputChannel
 CloseHandle PipeW4OutputChannel
 CloseHandle Proc.hThread
 CloseHandle Proc.hProcess
 
If EndProcess(Proc.dwProcessId) = False Then
'   MsgBox "主服务程序[CMD.EXE]没有关闭，请您手动关闭 ", vbInformation, "不好意思"
End If
End Function

Private Function EndProcess(ByVal ProcessID As Long) As Boolean
  Dim hProcess As Long, ExitCode As Long, Rst As Long
  hProcess = OpenProcess(PROCESS_TERMINATE Or PROCESS_QUERY_INFORMATION, True, ProcessID)
   If hProcess <> 0 Then
     GetExitCodeProcess hProcess, ExitCode
     If ExitCode <> 0 Then
       Rst = TerminateProcess(hProcess, 0)
       CloseHandle hProcess
        If Rst = 0 Then

         EndProcess = False
        Else
         EndProcess = True
        End If
     Else
       EndProcess = False
     End If
   Else
     EndProcess = False
   End If

End Function

Public Sub Abort()
    mblnIsAbort = True
    
    Call EndDosIo
    
    If mlTimer <> 0 Then
        CloseHandle mlTimer
    End If
End Sub

'===========================================================================================
Private Sub Class_Initialize()
    If InitDosIO <> 0 Then MsgBox "DOS通道创建失败！"
End Sub


Private Sub Class_Terminate()
On Error Resume Next
    Call Abort
    
    Err.Clear
End Sub

