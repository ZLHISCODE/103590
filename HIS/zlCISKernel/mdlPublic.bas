Attribute VB_Name = "mdlPublic"
Option Explicit
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public glngPreHWnd As Long '����֧�������ֹ���
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long�����ֵ
Public Type PointAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type MINMAXINFO
        ptReserved As PointAPI
        ptMaxSize As PointAPI
        ptMaxPosition As PointAPI
        ptMinTrackSize As PointAPI
        ptMaxTrackSize As PointAPI
End Type
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_GETMINMAXINFO = &H24
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Const LVM_SETCOLUMNWIDTH = &H101E
Public Const LVSCW_AUTOSIZE = -1
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Const MK_LBUTTON = &H1 '��ȡ������״̬
Public Const WM_MOUSEWHEEL = &H20A

'Shell------------------------------------------------
Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'��������ComboBox����
Public Const CB_ADDSTRING = &H143
Public Const CB_SETITEMDATA = &H151
Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetComboData Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'����API
'Public Const OPAQUE = 2
Public Const ETO_OPAQUE = 2
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'��ͼAPI
Public Const PS_SOLID = 0
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Const HWND_TOPMOST = -1

Public Const COLEditBackColor = &HE1FFE1    'ǳ��
Public Const COLSelBackColor = &HFAEADA          'ǳ��

Public Function MousePressButton(lngTbr As Long, objButton As Button) As Boolean
'���ܣ��жϵ�ǰ��Ļ����Ƿ���ָ�����߰�ť��ʾ�����ڰ���
    Dim vRect As RECT, vPos As PointAPI
        
    '���жϵ�ǰ�Ƿ��ڰ���״̬
    If (GetKeyState(MK_LBUTTON) And &H80) <> 0 Then
        '���жϵ�ǰ�����������Χ
        GetCursorPos vPos
        
        GetWindowRect lngTbr, vRect
        With objButton
            vRect.Left = vRect.Left + .Left / Screen.TwipsPerPixelX
            vRect.Top = vRect.Top + .Top / Screen.TwipsPerPixelY
            vRect.Right = vRect.Left + .Width / Screen.TwipsPerPixelX
            vRect.Bottom = vRect.Top + .Height / Screen.TwipsPerPixelY
        End With
        
        If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
            And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
            MousePressButton = True
        End If
    End If
End Function

Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'���ܣ��жϵ�ǰ��Ļ����Ƿ���ָ�����ڵ���ʾ������
    Dim vRect As RECT, vPos As PointAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Function MoveObj(lngHwnd As Long) As RECT
'���ܣ��ڶ����MouseDown�¼��е���,����������Hwnd����
'���أ������Ļ������ֵ
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'���ܣ�������ʽ���߰�ť�е���һ���˵�
    Dim vRect As RECT, vDot1 As PointAPI, vDot2 As PointAPI
    
    Call GetWindowRect(ToolBar.hwnd, vRect)
    vDot1.X = vRect.Left: vDot1.Y = vRect.Top
    vDot2.X = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot2)
    
    vDot1.X = vDot1.X * 15: vDot1.Y = vDot1.Y * 15
    vDot2.X = vDot2.X * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.X + Button.Left, vDot2.Y
End Sub

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hwnd, Msg, wp, lp)
End Function

Public Function GetColFormat(vsGrid As Object) As String
'���ܣ���ȡָ�������и�ʽ���Դ�
    Dim strTmp As String, i As Long
    
    For i = 0 To vsGrid.Cols - 1
        '�п�,�пɼ�,�ж���,������
        strTmp = strTmp & ";" & vsGrid.ColWidth(i) & "," & IIF(vsGrid.ColHidden(i), 0, 1) & "," & vsGrid.ColAlignment(i) & "," & Val(vsGrid.ColData(i))
    Next
    GetColFormat = Mid(strTmp, 2)
End Function

Public Sub SetColFormat(vsGrid As Object, ByVal strFormat As String)
'���ܣ��ָ�ָ�������и�ʽ
    Dim arrCol As Variant, i As Long
    If strFormat = "" Then Exit Sub
    
    arrCol = Split(strFormat, ";")
    For i = 0 To UBound(arrCol)
        '�п�,�пɼ�,�ж���,������
        vsGrid.ColWidth(i) = Val(Split(arrCol(i), ",")(0))
        vsGrid.ColHidden(i) = Val(Split(arrCol(i), ",")(1)) = 0
        vsGrid.ColAlignment(i) = Val(Split(arrCol(i), ",")(2))
        vsGrid.Cell(2, vsGrid.FixedRows, i, vsGrid.Rows - 1, i) = Val(Split(arrCol(i), ",")(2))
        
        vsGrid.ColData(i) = Val(Split(arrCol(i), ",")(3))
    Next
    vsGrid.Cell(2, 0, 0, vsGrid.FixedRows - 1, vsGrid.Cols - 1) = 4
End Sub

Public Function Custom_WndMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hwnd, Msg, wp, lp)
End Function

Public Function CheckAdviceWindow(ByVal strTitle As String) As Boolean
'���ܣ����ҽ���༭�����Ƿ��Ѿ���
    Dim lngHwnd As Long
    
    '�������ڴ���
    lngHwnd = FindWindow("ThunderFormDC", strTitle)
    If lngHwnd = 0 Then
        lngHwnd = FindWindow("ThunderRT6FormDC", strTitle)
    End If
    If lngHwnd <> 0 Then
        MsgBox "ҽ���༭�����Ѿ��򿪣�������ɵ�ǰ��������ִ�С�", vbInformation, gstrSysName
        Call ShowWindow(lngHwnd, SW_RESTORE)
        Call BringWindowToTop(lngHwnd)
        Exit Function
    End If
    CheckAdviceWindow = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function BeSureMode(ByVal strMode1 As String, strMode2 As String, strMode3 As String) As Integer
'���ܣ�ʶ�����ֱ���ģʽ�����ȼ�
    If strMode1 = "-" Then BeSureMode = 1: Exit Function
    If strMode2 = "-" Then BeSureMode = 2: Exit Function
    If strMode3 = "-" Then BeSureMode = 3: Exit Function
    If strMode1 <> "" And strMode2 = "" And strMode3 = "" Then BeSureMode = 1: Exit Function
    If strMode1 = "" And strMode2 <> "" And strMode3 = "" Then BeSureMode = 2: Exit Function
    If strMode1 = "" And strMode2 = "" And strMode3 <> "" Then BeSureMode = 3: Exit Function
    If strMode3 <> "" Then BeSureMode = 3: Exit Function
    If strMode2 <> "" Then BeSureMode = 2: Exit Function
    If strMode1 <> "" Then BeSureMode = 1: Exit Function
End Function

Public Function VsfOnlyOneRow(ByRef vsf As VSFlexGrid) As Boolean
'���ܣ��ж��Ƿ����һ�пɼ���
    Dim i As Long, k As Long
    
    With vsf
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then
                k = k + 1
                If k > 1 Then Exit For
            End If
        Next
        VsfOnlyOneRow = Not (k > 1)
    End With
End Function

Public Function GetPriceType(ByVal lng����ID As Long, ByVal lng�շ�ϸĿID As Long, ByVal int���� As Integer, ByVal bln���� As Boolean) As String
'���ܣ���ȡ��Ӧҽ����������
    Dim str���� As String
    
    'ȡҽ���ķ�������
    If int���� <> 0 Then
        On Error Resume Next
        str���� = gclsInsure.GetItemInsure(lng����ID, lng�շ�ϸĿID, 0, bln����, int����)
        If str���� <> "" Then
            If UBound(Split(str����, ";")) >= 5 Then
                str���� = Split(str����, ";")(5)
            Else
                str���� = ""
            End If
        End If
    End If
    GetPriceType = str����
End Function

Public Function DynamicCreate(ByVal strClass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strClass)
    
    If err <> 0 Then
        If blnMsg Then MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function

Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'���ܣ��ж��������Ƿ���ԭ�ۺ��ִ��޶��ķ�Χ��
'������varL=ԭ��,varR=�ּ�,varI=������
'���أ�������ڷ�Χ��,��Ϊ��ʾ��Ϣ,����Ϊ�մ�
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '�����ֵ������ͬ,���þ���ֵ�ж�
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "����ļ۸����ֵ���ڷ�Χ(" & FormatEx(Abs(varL), 5) & "-" & FormatEx(Abs(varR), 5) & ")��."
        End If
    Else
        '������Ų���ͬ,����ԭʼ��Χ�ж�
        If varI < varL Or varI > varR Then
            CheckScope = "����ļ۸�ֵ���ڷ�Χ(" & FormatEx(varL, 5) & "-" & FormatEx(varR, 5) & ")��."
        End If
    End If
End Function
