Attribute VB_Name = "MdlMdi"
Option Explicit
'--菜单函数--
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long

'返回窗体的菜单句柄
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'返回指定位置的弹出菜单的句柄
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'获取指定菜单的句柄(弹出菜单返回-1;分隔菜单返回0)
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'获取指定菜单的菜单项数
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'取指定菜单项的字串
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
'参数 类型及说明
'hMenu Long，                   菜单的句柄
'nPosition Long，               定义了新条目插入点的一个现有菜单条目的标志符。如果在wFlags中指定了MF_BYCOMMAND标志，
'                           这个参数就代表欲改变的菜单条目的命令ID。如设置的是MF_BYPOSITION标志，这个参数就代表
'                           菜单条目在菜单中的位置，第一个条目的位置为零
'wFlags Long，                  一系列常数标志的组合。参考ModifyMenu
'wIDNewItem Long，              指定菜单条目的新菜单ID。如果在wFlags中指定了MF_POPUP标志，就应该指定弹出式菜单的一个句柄
'lpNewItem                      如果在wFlags参数中设置了MF_STRING标志，就代表要设置到菜单中的字串（String）。
'                           如设置的是MF_BITMAP标志，就代表一个Long型变量，其中包含了一个位图句柄
'常数列表
Public Const MF_BYPOSITION = &H400&
Public Const MF_STRING = &H0&               '在指定的条目处放置一个字串。不与vb的caption属性兼容
Public Const MF_POPUP = &H10&               '将一个弹出式菜单置于指定的条目.可用于创建子菜单及弹出式菜单
Public Const MF_SEPARATOR = &H800&          '在指定的条目处显示一条分隔线
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'为指定窗体设置新的菜单
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const WM_COMMAND = &H111

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CascadeWindows% Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpKids As Long)
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type POINTAPI
     x As Long
     y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Const WM_GETMINMAXINFO = &H24

'公用
Public LngAddFunc As Long
Public CollMenu As New Collection                       '菜单集合
Public CollOpenWindowHdl As New Collection              '已运行的窗体句柄
Public Const Menu_Hdl As Integer = 0                    '菜单句柄
Public Const Menu_Code As Integer = 1                   '菜单编号
Public Const Menu_Modul As Integer = 2                  '菜单模块
Public Const Menu_Component As Integer = 3              '对应部件名称
Public Const Menu_UpperHdl As Integer = 4               '其上级菜单句柄
Public Const Menu_Caption As Integer = 5                '标题及快捷键
Public Const Menu_ID As Integer = 6                     '菜单ID
Public Const Menu_Sys As Integer = 7                    '系统编号

Public gLngMinH As Double
Public gLngMinW As Double
Public gLngMaxH As Double
Public gLngMaxW As Double

Public Function MenuProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim LngFind As Long, BlnFind As Boolean, BlnRun As Boolean
    Dim StrComponent As String, lngModul As Long, StrCaption As String, lngSys As Long
    Dim LngTargetHdl As Long
    
    '处理菜单事件
    If uMsg = WM_COMMAND Then
        '查找对应的集合
        If wParam >= 菜单基准.其它功能菜单 - 1 Then
            Select Case wParam
            Case 99999901   '启动自定义报表
                Call ExecuteFunc(0, "ZL9REPORT", 99999901)
            End Select
        ElseIf wParam >= 菜单基准.窗口菜单 - 1 Then '窗口列表菜单
            For LngFind = 0 To CollOpenWindowHdl.Count - 1
                If wParam = CollOpenWindowHdl("K_" & LngFind)(2) Then
                    LngTargetHdl = CollOpenWindowHdl("K_" & LngFind)(0)
                    Exit For
                End If
            Next
            
            If IsIconic(LngTargetHdl) Then
                Call ShowWindow(LngTargetHdl, 9)    '还原指定窗体为原大小
            End If
            Call SetActiveWindow(LngTargetHdl)
            MenuProc = 1
        ElseIf wParam > 菜单基准.功能菜单 - 1 Then  '程序模块列表菜单
            BlnFind = False
            For LngFind = 0 To CollMenu.Count - 1
                If CollMenu("K_" & LngFind)(Menu_ID) = wParam Then
                    BlnFind = True
                    StrCaption = CollMenu("K_" & LngFind)(Menu_Caption)
                    If InStr(1, StrCaption, "(") <> 0 And InStr(1, StrCaption, ")") <> 0 Then StrCaption = Mid(StrCaption, 1, InStr(1, StrCaption, "("))
                    StrComponent = CollMenu("K_" & LngFind)(Menu_Component)
                    lngSys = CollMenu("K_" & LngFind)(Menu_Sys)
                    lngModul = CollMenu("K_" & LngFind)(Menu_Modul)
                    Exit For
                End If
            Next
            '找到则执行
            If BlnFind Then
                Call AddHistory(lngSys & "," & lngModul)
                Call frmMdi.LoadHistory
                
                '查找该模块是否已运行,是则设为活动窗体
                BlnRun = False
                For LngFind = 0 To CollOpenWindowHdl.Count - 1
                    If StrCaption = CollOpenWindowHdl("K_" & LngFind)(1) Then
                        BlnRun = True
                        LngTargetHdl = CollOpenWindowHdl("K_" & LngFind)(0)
                        Exit For
                    End If
                Next
                If BlnRun Then
                    If IsIconic(LngTargetHdl) Then
                        Call ShowWindow(LngTargetHdl, 9)            '还原指定窗体为原大小
                    End If
                    Call SetActiveWindow(LngTargetHdl)
                Else
                    ExecuteFunc lngSys, StrComponent, lngModul
                End If
                MenuProc = 1
            Else
                MenuProc = CallWindowProc(LngAddFunc, FrmMainface.hwnd, uMsg, wParam, lParam)
            End If
        Else
            MenuProc = CallWindowProc(LngAddFunc, FrmMainface.hwnd, uMsg, wParam, lParam)
        End If
    ElseIf uMsg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lParam, Len(MinMax)
        MinMax.ptMinTrackSize.x = gLngMinW \ 15
        MinMax.ptMinTrackSize.y = gLngMinH \ 15
        MinMax.ptMaxTrackSize.x = gLngMaxW \ 15
        MinMax.ptMaxTrackSize.y = gLngMaxH \ 15
        CopyMemory ByVal lParam, MinMax, Len(MinMax)
        MenuProc = 1
    Else
        MenuProc = CallWindowProc(LngAddFunc, FrmMainface.hwnd, uMsg, wParam, lParam)
    End If
End Function


