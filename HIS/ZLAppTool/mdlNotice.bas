Attribute VB_Name = "mdlNotice"
Option Explicit

Public AlertCount As Integer

Public Const PI    As Double = 3.14159265358979
Public Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

Public Type POINTAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ

Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public glngTXTProc As Long


'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Sub DrawAngle(picDraw As PictureBox, ByVal fAngle As Single)
    '---------------------------------------------------------------------------------------------------------------
    '����:
    '---------------------------------------------------------------------------------------------------------------
    Dim iSize       As Integer
    Dim iFillStyle  As Integer
    Dim lFillColor  As Long
    Dim lForeColor  As Long
    Dim lRet        As Long
    Dim uaPts(3)    As POINTAPI

    'Size arrow to best fit picDraw at any angle
    iSize = IIf(picDraw.ScaleHeight < picDraw.ScaleWidth, Int(picDraw.ScaleHeight / PI), Int(picDraw.ScaleWidth / PI))
    
    'Setup the 4 points of the arrow using the first point
    'as the center and the other points offset from the center.
    uaPts(0).X = picDraw.ScaleWidth / 2
    uaPts(0).Y = picDraw.ScaleHeight / 2
    uaPts(1).X = uaPts(0).X - iSize
    uaPts(1).Y = uaPts(0).Y - iSize
    uaPts(2).X = uaPts(0).X + iSize
    uaPts(2).Y = uaPts(0).Y
    uaPts(3).X = uaPts(0).X - iSize
    uaPts(3).Y = uaPts(0).Y + iSize
    
    'Rotate the arrow to the correct angle
    Call RotatePoints(uaPts(0), uaPts, fAngle)
    
    'Save picDraw settings
    iFillStyle = picDraw.FillStyle
    lFillColor = picDraw.FillColor
    lForeColor = picDraw.ForeColor
    
    'Setup picDraw to fill the arrow
    picDraw.FillStyle = vbFSSolid   'Solid Fill
    picDraw.FillColor = &HFFFFFF    'Inside = White
    picDraw.ForeColor = &H0&        'Border = Black
    
    'Draw the filled arrow
    lRet = Polygon(picDraw.hDC, uaPts(0), 4)
    
    'Restore picDraw settings
    picDraw.FillStyle = iFillStyle
    picDraw.FillColor = lFillColor
    picDraw.ForeColor = lForeColor

    'Free the memory
    Erase uaPts
    
End Sub


Private Sub RotatePoints(uAxisPt As POINTAPI, uRotatePts() As POINTAPI, fDegrees As Single)
    '---------------------------------------------------------------------------------------------------------------
    '����:
    '---------------------------------------------------------------------------------------------------------------
    
    'Rotates an array of PointAPI points around a center point by fDegrees
    
    Dim lIdx        As Long
    Dim fDX         As Single
    Dim fDY         As Single
    Dim fRadians    As Single

    fRadians = fDegrees * RADS
    
    For lIdx = 0 To UBound(uRotatePts)
        fDX = uRotatePts(lIdx).X - uAxisPt.X
        fDY = uRotatePts(lIdx).Y - uAxisPt.Y
        uRotatePts(lIdx).X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
        uRotatePts(lIdx).Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    Next lIdx
    
End Sub

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture)
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = "˫����������Ϣ�б�" & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub ModifyIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture)
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = "˫����������ϸ���" & Chr$(0)

    Shell_NotifyIcon NIM_MODIFY, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '���ܣ�����������ɾ��ͼ��
    
    Dim t As NOTIFYICONDATA
    On Error Resume Next
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '�¼�����������
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
    
End Sub


Public Sub DrawColorToColor(picDraw As Object, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal blnVertical As Boolean = True, Optional ByVal blnBorder As Boolean = False)
'������һ����ɫ����һ����ɫ�Ľ���
    Dim VR, VG, VB As Single
    Dim R, G, b, R2, G2, B2 As Integer
    Dim temp As Long, Y As Long, X As Long
    Dim tmpMode As Long
    Dim blnAutoRedraw As Boolean
    
    'ֻ�д����ͼƬ���Ի�
    If Not (TypeOf picDraw Is PictureBox Or TypeOf picDraw Is Form) Then Exit Sub
    tmpMode = picDraw.ScaleMode
    blnAutoRedraw = picDraw.AutoRedraw
    
    picDraw.ScaleMode = 3
    picDraw.AutoRedraw = True
    
    temp = (Color1 And 255)
    R = temp And 255
    temp = Int(Color1 / 256)
    G = temp And 255
    temp = Int(Color1 / 65536)
    b = temp And 255
    temp = (Color2 And 255)
    R2 = temp And 255
    temp = Int(Color2 / 256)
    G2 = temp And 255
    temp = Int(Color2 / 65536)
    B2 = temp And 255

    If blnVertical Then
        VR = Abs(R - R2) / picDraw.ScaleHeight
        VG = Abs(G - G2) / picDraw.ScaleHeight
        VB = Abs(b - B2) / picDraw.ScaleHeight
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB
        For Y = 0 To picDraw.ScaleHeight
            R2 = R + VR * Y
            G2 = G + VG * Y
            B2 = b + VB * Y
            picDraw.Line (0, Y)-(picDraw.ScaleWidth, Y), RGB(R2, G2, B2)
        Next Y
    Else
        VR = Abs(R - R2) / picDraw.ScaleWidth
        VG = Abs(G - G2) / picDraw.ScaleWidth
        VB = Abs(b - B2) / picDraw.ScaleWidth
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB
        For X = 0 To picDraw.ScaleWidth
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X
            picDraw.Line (X, 0)-(X, picDraw.ScaleHeight), RGB(R2, G2, B2)
        Next X
    End If
    
    If blnBorder Then
        picDraw.DrawWidth = 2
        picDraw.Line (1, 1)-(picDraw.ScaleWidth - 1, picDraw.ScaleHeight - 1), &HC000&, B
        picDraw.DrawWidth = 1
    End If
    
    picDraw.Refresh
    picDraw.ScaleMode = tmpMode
    picDraw.AutoRedraw = blnAutoRedraw
End Sub

Public Function AppendSapceRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:������ؼ��Ŀ���
    '����:objVsf Ҫ�����еı��ؼ�����
    '����:���ɹ�����True,���򷵻� False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    lngTop = objVsf.Cell(flexcpTop, objVsf.Rows - 1, 0) + objVsf.RowHeight(objVsf.Rows - 1)
    '1.�������е���
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    '2.���¼�����Ҫ������
    For lngLoop = 1 To objVsf.Cols - 1
        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)
        With objLineY(lngLoop)
            .ZOrder
            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height
            .BorderColor = objVsf.GridColor
            .Visible = True
        End With
    Next
    
    '3.���¼�����Ҫ�ĺ���
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height
        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)
        With objLineX(lngIndex)
            .ZOrder
            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0)
            .Y2 = .Y1
            .BorderColor = objVsf.GridColor
            .Visible = True
            lngTop = .Y1
        End With
    Loop
        
    AppendSapceRows = True
    Exit Function
    
errHand:

End Function
