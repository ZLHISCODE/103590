Attribute VB_Name = "mdlCommon"
Option Explicit
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
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
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_GETMINMAXINFO = &H24
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const GWL_STYLE = (-16)              'Set the window style
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000        '��߿�
Public Const WS_SYSMENU = &H80000           '�ڱ������Ƿ�߱�ϵͳ�˵�
Public Const WS_MINIMIZEBOX = &H20000       '�߱���С����ť
Public Const WS_MAXIMIZEBOX = &H10000       '�߱���󻯰�ť
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Const MK_LBUTTON = &H1 '��ȡ������״̬
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_RESETCONTENT = &H14B
Public Const CB_ADDSTRING = &H143
Public gstrMatchMethod As String

Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Public Function VsScroll(vsf As VSFlexGrid) As Boolean '�ж�ˮƽ�������Ŀɼ���
    VsScroll = gobjComlib.Grid.VScrollVisible(vsf)
End Function
    
Public Function HeScroll(vsf As VSFlexGrid) As Boolean '�жϴ�ֱ�������Ŀɼ���
    HeScroll = gobjComlib.Grid.VScrollVisible(vsf)
End Function

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    If gstrMatchMethod = "" Then
        gstrMatchMethod = Val(gobjDatabase.GetPara("����ƥ��"))
    End If
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

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
    MouseInRect = gobjControl.MouseInRect(lngHwnd)
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Call gobjControl.FormSetCaption(objForm, blnCaption, blnBorder)
End Sub

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
    
    Call GetWindowRect(ToolBar.hWnd, vRect)
    vDot1.X = vRect.Left: vDot1.Y = vRect.Top
    vDot2.X = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hWnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hWnd, vDot2)
    
    vDot1.X = vDot1.X * 15: vDot1.Y = vDot1.Y * 15
    vDot2.X = vDot2.X * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.X + Button.Left, vDot2.Y
End Sub

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As PointAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As PointAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    SysColor2RGB = gobjComlib.OS.SysColor2RGB(lngColor)
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function



Public Sub FindCboIndex(objCbo As Object, lngData As Long, Optional Keep As Boolean)
'���ܣ�����Ŀֵ����ComboBox����Ŀ����
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Call gobjComlib.Cbo.FindIndex(objCbo, lngData, Keep)
End Sub

Public Sub GetCboIndex(objCbo As Object, strFind As String, Optional Keep As Boolean)
'���ܣ����ַ�����ComboBox�в�������
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Call gobjComlib.Cbo.FindIndex(objCbo, strFind, Keep)
End Sub

Public Function SeekCboIndex(objCbo As Object, varData As Variant) As Long
'���ܣ���ItemData��Text����ComboBox������ֵ
    SeekCboIndex = gobjComlib.Cbo.FindIndex(objCbo, varData)
End Function
Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
    IsType = gobjComlib.Rec.IsType(varType, varBase)
End Function
Public Function InDesign() As Boolean
    InDesign = gobjComlib.OS.IsDesinMode
End Function

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
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
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, Msg, wp, lp)
End Function
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnTrans As Boolean = True, Optional blnCommit As Boolean = True)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnTrans-�Ƿ��������
    '     blnCommit-ִ������̺�,�ύ����(ǰ��:blnTrans=true)
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnTrans Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call gobjDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnCommit And blnTrans Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 'ʹComboBox����ĵ���ƥ�书��ʧЧ
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '������Իس���������
    End If
End Function

Public Function zlCboFindItem(ByVal cboObj As Object, ByVal lngFindID As Long, _
    Optional strItem As String = "", Optional blnOnlyFind As Boolean = True, Optional blnFindLocal As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Combox��ItemData���ݽ��ж�λ
    '��Σ�cboObj-Combox����
    '         lngFindID-��Ҫ���ҵ�ID
    '         strItem-��Ҫ���ҵĻ����ӵ�����(��blnOnlyFind=false)ʱ
    '         blnOnlyFind-�Ƿ����.
    '        blnFindLocal-�ҵ���,��λ��
    '���Σ�
    '���أ��ҵ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-04-06 17:28:17
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngLocate As Long
    zlCboFindItem = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If cboObj.ItemData(lngLocate) = lngFindID Then
            If blnFindLocal Then cboObj.ListIndex = lngLocate
            zlCboFindItem = True
            Exit Function
        End If
    Next
    If blnOnlyFind Then Exit Function
    cboObj.AddItem strItem
    cboObj.ItemData(cboObj.NewIndex) = lngFindID
    If blnFindLocal Then cboObj.ListIndex = cboObj.NewIndex
    zlCboFindItem = True
End Function
Public Function zlCheckPrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '����:���ָ����Ȩ���Ƿ����
    '����:strPrivs-Ȩ�޴�
    '     strMyPriv-����Ȩ��
    '����,����Ȩ��,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    zlCheckPrivs = gobjComlib.zlStr.IsHavePrivs(strPrivs, strMyPriv)
End Function


Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(gobjDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo Errhand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
Errhand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(gobjDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(gobjDataBase.GetPara("������������", , , True)) = 1
    Err = 0: On Error GoTo Errhand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
Errhand:
End Function
Public Function Between(X, a, B) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    Between = gobjComlib.Between(X, a, B)
End Function
Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim StrPass As String, strReturn As String, strSource As String, strTarget As String
    
    StrPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(StrPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function zlInitMEPIPati(ByRef rsPati As ADODB.Recordset) As Boolean
    Set rsPati = New ADODB.Recordset
    With rsPati
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "����ID", adBigInt, , adFldIsNullable
            .Append "��ҳID", adBigInt, , adFldIsNullable
            .Append "�Һ�ID", adBigInt, , adFldIsNullable
            .Append "�����", adVarChar, 18, adFldIsNullable
            .Append "סԺ��", adVarChar, 18, adFldIsNullable
            .Append "ҽ����", adVarChar, 30, adFldIsNullable
            .Append "���֤��", adVarChar, 18, adFldIsNullable
            .Append "����֤��", adVarChar, 20, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "�Ա�", adVarChar, 4, adFldIsNullable
            .Append "��������", adVarChar, 20, adFldIsNullable
            .Append "�����ص�", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 20, adFldIsNullable
            .Append "ѧ��", adVarChar, 10, adFldIsNullable
            .Append "ְҵ", adVarChar, 80, adFldIsNullable
            .Append "������λ", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����״��", adVarChar, 4, adFldIsNullable
            .Append "��ͥ�绰", adVarChar, 20, adFldIsNullable
            .Append "��ϵ�˵绰", adVarChar, 20, adFldIsNullable
            .Append "��λ�绰", adVarChar, 20, adFldIsNullable
            .Append "��ͥ��ַ", adVarChar, 100, adFldIsNullable
            .Append "��ͥ��ַ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "���ڵ�ַ", adVarChar, 100, adFldIsNullable
            .Append "���ڵ�ַ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "��λ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "��ϵ�˵�ַ", adVarChar, 100, adFldIsNullable
            .Append "��ϵ�˹�ϵ", adVarChar, 30, adFldIsNullable
            .Append "��ϵ������", adVarChar, 64, adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    zlInitMEPIPati = True
End Function
