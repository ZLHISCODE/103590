Attribute VB_Name = "mdlProcess"
'///////////////////////////////////////////////////////////////////////////////
'
'       ģ�飺���̾������
'       ���ܣ����̾���������ָ�����̵�Hwnd
'       ��д��ף��
'       ���ڣ�2010��11��24��
'
'///////////////////////////////////////////////////////////////////////////////

Option Explicit

Private mlngPid As Long
Public gHwnd As Long
'==============================================================================
'=���ܣ� ͨ��PIDö�������ľ��,������Ҫ�Ĵ���
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
'        frmPidHwnd.List1.AddItem "���:" & hwnd & "  ����:" & wText
    End If
    EnumWindowsProc = True
End Function

Public Sub Find_Window(ByVal lngPid As Long)
    mlngPid = lngPid
    gHwnd = 0
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub


'���ҽ��̵ĺ���
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
