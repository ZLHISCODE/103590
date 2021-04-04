Attribute VB_Name = "mdlAPI"
Option Explicit
'--------------------------------------------------------
'功  能：本模块用于存储API调用的各种函数
'编制人：赵彤宇
'编制日期：2004.6
'过程函数清单：
'       ShowTitle() 设置窗体是否显示标题栏
'修改记录：
'
'-------------------------------------------------------
Public frmMain As frmViewer
''建立子目录
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''处理鼠标滚轮
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROC = (-4)

Public preWinProc As Long
Public plngFilmPreWndProc As Long       'Film窗体原来的消息处理程序
Public plngFilmViewPreWndProc As Long       'Film窗体原来的消息处理程序

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''给Pic画凹凸使用
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''[放大镜使用]''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'判断数组是否为空
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'使用API函数修改MsgBox，使其可以在调用的时候，指定父窗体
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const MB_OK = &H0&

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'使用主板播放声音
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Const BEEP_Do0 = 264
Public Const BEEP_Re = 297
Public Const BEEP_Mi = 330
Public Const BEEP_Fa = 352
Public Const BEEP_Sol = 396
Public Const BEEP_la = 440
Public Const BEEP_Ti = 495
Public Const BEEP_Do1 = 528

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'得到Mouse坐标,用于移动窗体
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'**********************************另存文件API调用*****************************************
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'***********************************************************************************


Public Function Wndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pt As POINTAPI
    Dim wzDelta As Integer
    On Error Resume Next
    wzDelta = OS.HIWORD(wParam)
    
    Select Case Msg
        Case WM_MOUSEWHEEL
            If Sgn(wzDelta) = 1 Then    '鼠标上滚
                Call frmMain.MouseWheel(1)
            Else                        '鼠标下滚
                Call frmMain.MouseWheel(0)
            End If
    End Select
    Wndproc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
''为了处理双屏时对话框的正确显示位置，用API函数改写了一下MsgBox函数
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = MB_OK, _
    Optional Title As String = "", Optional frmParent As Object = Nothing) As Long
    If Not frmParent Is Nothing Then
        MsgBox = MessageBox(frmParent.hwnd, Prompt, Title, Buttons)
    ElseIf frmMain Is Nothing Then
        MsgBox = VBA.Interaction.MsgBox(Prompt, Buttons, Title)
    Else
        MsgBox = MessageBox(frmMain.hwnd, Prompt, Title, Buttons)
    End If

End Function

Public Function FilmHook(ByVal hwnd As Long) As Long
    '返回并保存原来默认的窗口过程指针
    If App.LogMode <> 0 Then
        FilmHook = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf FilmWindowProc)
    End If
End Function

Public Sub FilmUnhook(ByVal hwnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
    If App.LogMode <> 0 Then
        temp = SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc)
    End If
End Sub

Function FilmWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'------------------------------------------------
'功能：胶片打印窗口的windows消息处理程序，专门处理鼠标滚轮 消息
'参数：
'返回：
'------------------------------------------------
    Dim pt As POINTAPI
    Dim wzDelta As Integer

    wzDelta = OS.HIWORD(wParam)

    If uMsg = WM_MOUSEWHEEL Then
        If Not frmMain.mfrmFilm Is Nothing Then
            If Sgn(wzDelta) = 1 Then    '鼠标上滚
                Call frmMain.mfrmFilm.MouseWheel(1)
            Else                        '鼠标下滚
                Call frmMain.mfrmFilm.MouseWheel(0)
            End If
        End If
    End If
  
    '调用原来的窗口过程
    FilmWindowProc = CallWindowProc(plngFilmPreWndProc, hw, uMsg, wParam, lParam)
End Function

Public Function FilmViewHook(ByVal hwnd As Long) As Long
    '返回并保存原来默认的窗口过程指针
    If App.LogMode <> 0 Then
        FilmViewHook = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf FilmViewWindowProc)
    End If
End Function

Public Sub FilmViewUnhook(ByVal hwnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
    If App.LogMode <> 0 Then
        temp = SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc)
    End If
End Sub

Function FilmViewWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'------------------------------------------------
'功能：胶片打印窗口的windows消息处理程序，专门处理鼠标滚轮 消息
'参数：
'返回：
'------------------------------------------------
    Dim pt As POINTAPI
    Dim wzDelta As Integer

    wzDelta = OS.HIWORD(wParam)

    If uMsg = WM_MOUSEWHEEL Then
        If Not frmMain.mfrmFilm Is Nothing Then
            If Not frmMain.mfrmFilm.mfrmFilmView Is Nothing Then
                If Sgn(wzDelta) = 1 Then    '鼠标上滚
                    Call frmMain.mfrmFilm.mfrmFilmView.MouseWheel(1)
                Else                        '鼠标下滚
                    Call frmMain.mfrmFilm.mfrmFilmView.MouseWheel(0)
                End If
            End If
        End If
    End If
  
    '调用原来的窗口过程
    FilmViewWindowProc = CallWindowProc(plngFilmViewPreWndProc, hw, uMsg, wParam, lParam)
End Function
