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

'进度结构
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 1024
End Type

Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000

'''进程处理API
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
        ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

'''窗体处理API
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim mlngPid As Long
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
