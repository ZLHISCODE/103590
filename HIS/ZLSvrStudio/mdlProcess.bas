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
Public gHwnd As Long
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
            gHwnd = hwnd
        End If
'        frmPidHwnd.List1.AddItem "句柄:" & hwnd & "  标题:" & wText
    End If
    EnumWindowsProc = True
End Function

Public Sub Find_Window(ByVal lngPid As Long)
    mlngPid = lngPid
    gHwnd = 0
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub


'查找进程的函数
Public Sub fun_KillProcess(ByVal ProcessName As String)
    Dim strData As String
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
