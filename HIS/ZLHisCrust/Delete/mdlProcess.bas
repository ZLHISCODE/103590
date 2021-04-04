Attribute VB_Name = "mdlProcess"
'///////////////////////////////////////////////////////////////////////////////
'
'       模块：进程句柄操作
'       功能：进程句柄操作获得指定进程的Hwnd
'       编写：祝庆
'       日期：2010年11月24日
'
'///////////////////////////////////////////////////////////////////////////////

Option Explicit
Private mlngPid As Long
Private mlngHwnd As Long
'==============================================================================
'=功能： 通过PID枚举所属的句柄,查找需要的窗口
'==============================================================================
Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim Pid1 As Long
    Dim wText As String * 255
    GetWindowThreadProcessId hwnd, Pid1
    If mlngPid = Pid1 Then
        GetWindowText hwnd, wText, 100
        If InStrRev(wText, "%", -1) > 0 Then
            mlngHwnd = hwnd
        End If
'        frmPidHwnd.List1.AddItem "句柄:" & hwnd & "  标题:" & wText
    End If
    EnumWindowsProc = True
End Function

Public Sub Find_Window(ByVal lngPid As Long)
    mlngPid = lngPid
    mlngHwnd = 0
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub


'查找进程的函数
Public Sub fun_KillProcess(ByVal ProcessName As String)
    Dim strdata As String
    Dim my As PROCESSENTRY32
    Dim l As Long
    Dim l1 As Long
    Dim mName As String
    Dim i As Integer, Pid As Long
    Dim mProcID As Long
    l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If l Then
        my.dwSize = 1060
        If (Process32First(l, my)) Then
            Do
                i = InStr(1, my.szExeFile, Chr(0))
                mName = LCase(Left(my.szExeFile, i - 1))
                If mName = LCase(ProcessName) Then
                    Pid = my.th32ProcessID
                    mProcID = OpenProcess(1&, -1&, Pid)

                    TerminateProcess mProcID, 0&
                End If
            Loop Until (Process32Next(l, my) < 1)
        End If
        l1 = CloseHandle(l)
    End If
End Sub

'查找进程是否存在
Public Function fun_ExitsProcess(ByVal ProcessName As String) As Long
    Dim strdata As String
    Dim my As PROCESSENTRY32
    Dim l As Long
    Dim l1 As Long
    Dim mName As String
    Dim i As Integer, Pid As Long
    Dim mProcID As Long
    l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If l Then
        my.dwSize = 1060
        If (Process32First(l, my)) Then
            Do
                i = InStr(1, my.szExeFile, Chr(0))
                mName = LCase(Left(my.szExeFile, i - 1))
                If mName = LCase(ProcessName) Then
                    Pid = my.th32ProcessID
                    mProcID = OpenProcess(1&, -1&, Pid)
                    fun_ExitsProcess = mProcID
                End If
            Loop Until (Process32Next(l, my) < 1)
        End If
        l1 = CloseHandle(l)
    End If
End Function

Public Sub KillProcess(ByVal mProcID As Long)
    On Error Resume Next
    Call TerminateProcess(mProcID, 0&)
End Sub

Public Function TerminatePID(ByVal lngPid As Long) As Boolean

    '功能:结束指定的进程
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16

    Dim lngProcess As Long, pHandle As Long, ret As Long
    
    TerminatePID = False
    
    On Error GoTo Errhand:
    pHandle = OpenProcess(SYNCHRONIZE, False, lngPid)
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
    Call TerminateProcess(lngProcess, 1&)
    ret = WaitForSingleObject(pHandle, INFINITE)
    ret = CloseHandle(pHandle)
    TerminatePID = True
Errhand:

End Function

Public Function zlGetFileProcess(ByVal strFile As String, ByRef cllOutProcess As Collection) As Boolean

    '功能:获取指定文件的相关进程
    '入参:strFile-指定的DLL文件
    '出参:cllOutProcess-返回被引用的进程值
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-01-20 13:59:35

    Dim uProcess As PROCESSENTRY32, uMdlInfor As MODULEENTRY32
    Dim lngMdlProcess As Long, strExeName As String, lngSnapShot As Long, strDLLName As String
    
    On Error GoTo Errhand:
    '创建进程快照
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
      uProcess.dwSize = Len(uProcess)
      If Process32First(lngSnapShot, uProcess) Then
        Do
          '获得进程的标识符
          strExeName = UCase(Left(Trim(uProcess.szExeFile), InStr(1, Trim(uProcess.szExeFile), vbNullChar) - 1))
          If strExeName Like "*" & UCase(strFile) & "*" Then
             '一般来说只有Exe文件才会存在
            On Error Resume Next
            cllOutProcess.Add Array(uProcess.th32ProcessID, strExeName, uProcess.th32ProcessID), "B" & uProcess.th32ProcessID
            If Err <> 0 Then
                cllOutProcess.Remove "B" & uMdlInfor.th32ProcessID
                cllOutProcess.Add Array(uProcess.th32ProcessID, strExeName, uProcess.th32ProcessID), "B" & uProcess.th32ProcessID
            End If
            On Error GoTo Errhand:
          Else
                lngMdlProcess = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, uProcess.th32ProcessID)
                If lngMdlProcess > 0 Then
                    uMdlInfor.dwSize = Len(uMdlInfor)
                    If Module32First(lngMdlProcess, uMdlInfor) Then
                          Do
                                strDLLName = UCase(Left(Trim(uMdlInfor.szExePath), InStr(1, Trim(uMdlInfor.szExePath), vbNullChar) - 1))
                                If uProcess.th32ProcessID = uMdlInfor.th32ProcessID Then
                                    If strDLLName Like "*" & UCase(strFile) & "*" Then
                                        On Error Resume Next
                                        cllOutProcess.Add Array(uProcess.th32ProcessID, strExeName, uMdlInfor.th32ProcessID), "K" & uMdlInfor.th32ProcessID
                                        If Err <> 0 Then
                                            cllOutProcess.Remove "K" & uMdlInfor.th32ProcessID
                                            cllOutProcess.Add Array(uProcess.th32ProcessID, strExeName, uMdlInfor.th32ProcessID), "K" & uMdlInfor.th32ProcessID
                                        End If
                                        On Error GoTo Errhand:
                                    End If
                                End If
                          Loop Until (Module32Next(lngMdlProcess, uMdlInfor) < 1)
                    End If
                    CloseHandle (lngMdlProcess)
                End If
            End If
        Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
      End If
      CloseHandle (lngSnapShot)
    End If
    zlGetFileProcess = True
    Exit Function
Errhand:
End Function
