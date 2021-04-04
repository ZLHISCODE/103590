Attribute VB_Name = "MdlMdi"
Option Explicit
'--�˵�����--
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long

'���ش���Ĳ˵����
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'����ָ��λ�õĵ����˵��ľ��
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'��ȡָ���˵��ľ��(�����˵�����-1;�ָ��˵�����0)
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'��ȡָ���˵��Ĳ˵�����
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'ȡָ���˵�����ִ�
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
'���� ���ͼ�˵��
'hMenu Long��                   �˵��ľ��
'nPosition Long��               ����������Ŀ������һ�����в˵���Ŀ�ı�־���������wFlags��ָ����MF_BYCOMMAND��־��
'                           ��������ʹ������ı�Ĳ˵���Ŀ������ID�������õ���MF_BYPOSITION��־����������ʹ���
'                           �˵���Ŀ�ڲ˵��е�λ�ã���һ����Ŀ��λ��Ϊ��
'wFlags Long��                  һϵ�г�����־����ϡ��ο�ModifyMenu
'wIDNewItem Long��              ָ���˵���Ŀ���²˵�ID�������wFlags��ָ����MF_POPUP��־����Ӧ��ָ������ʽ�˵���һ�����
'lpNewItem                      �����wFlags������������MF_STRING��־���ʹ���Ҫ���õ��˵��е��ִ���String����
'                           �����õ���MF_BITMAP��־���ʹ���һ��Long�ͱ��������а�����һ��λͼ���
'�����б�
Public Const MF_BYPOSITION = &H400&
Public Const MF_STRING = &H0&               '��ָ������Ŀ������һ���ִ�������vb��caption���Լ���
Public Const MF_POPUP = &H10&               '��һ������ʽ�˵�����ָ������Ŀ.�����ڴ����Ӳ˵�������ʽ�˵�
Public Const MF_SEPARATOR = &H800&          '��ָ������Ŀ����ʾһ���ָ���
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Ϊָ�����������µĲ˵�
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

'����
Public LngAddFunc As Long
Public CollMenu As New Collection                       '�˵�����
Public CollOpenWindowHdl As New Collection              '�����еĴ�����
Public Const Menu_Hdl As Integer = 0                    '�˵����
Public Const Menu_Code As Integer = 1                   '�˵����
Public Const Menu_Modul As Integer = 2                  '�˵�ģ��
Public Const Menu_Component As Integer = 3              '��Ӧ��������
Public Const Menu_UpperHdl As Integer = 4               '���ϼ��˵����
Public Const Menu_Caption As Integer = 5                '���⼰��ݼ�
Public Const Menu_ID As Integer = 6                     '�˵�ID
Public Const Menu_Sys As Integer = 7                    'ϵͳ���

Public gLngMinH As Double
Public gLngMinW As Double
Public gLngMaxH As Double
Public gLngMaxW As Double

Public Function MenuProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim LngFind As Long, BlnFind As Boolean, BlnRun As Boolean
    Dim StrComponent As String, lngModul As Long, StrCaption As String, lngSys As Long
    Dim LngTargetHdl As Long
    
    '����˵��¼�
    If uMsg = WM_COMMAND Then
        '���Ҷ�Ӧ�ļ���
        If wParam >= �˵���׼.�������ܲ˵� - 1 Then
            Select Case wParam
            Case 99999901   '�����Զ��屨��
                Call ExecuteFunc(0, "ZL9REPORT", 99999901)
            End Select
        ElseIf wParam >= �˵���׼.���ڲ˵� - 1 Then '�����б�˵�
            For LngFind = 0 To CollOpenWindowHdl.Count - 1
                If wParam = CollOpenWindowHdl("K_" & LngFind)(2) Then
                    LngTargetHdl = CollOpenWindowHdl("K_" & LngFind)(0)
                    Exit For
                End If
            Next
            
            If IsIconic(LngTargetHdl) Then
                Call ShowWindow(LngTargetHdl, 9)    '��ԭָ������Ϊԭ��С
            End If
            Call SetActiveWindow(LngTargetHdl)
            MenuProc = 1
        ElseIf wParam > �˵���׼.���ܲ˵� - 1 Then  '����ģ���б�˵�
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
            '�ҵ���ִ��
            If BlnFind Then
                Call AddHistory(lngSys & "," & lngModul)
                Call frmMdi.LoadHistory
                
                '���Ҹ�ģ���Ƿ�������,������Ϊ�����
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
                        Call ShowWindow(LngTargetHdl, 9)            '��ԭָ������Ϊԭ��С
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


