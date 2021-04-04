Attribute VB_Name = "mdlCISCore"
Option Explicit

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gobjCISCore As clsCISCore

Public gstrSysName As String                'ϵͳ����
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrHelpPath As String
Public gstrDBUser As String
Public gblnOK As Boolean

Public glngSys As Long                      '������¼ϵͳ��
Public gstrPrivs As String                  '������¼Ȩ��

Public gstrMatch As String                  '���ݱ��ز�����ƥ��ģʽ��ȷ������ƥ�����
Public gstrSql As String

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO
Public grsDuty As ADODB.Recordset '���ҽ��ְ��
Public grsSysPars As ADODB.Recordset

Public glngPen As Long '��ǰ���ʶ���
Public glngBrush As Long '��ǰˢ�Ӷ���

Public gcurPenColor As Long '��ǰʹ�õ�����ɫ
Public gcurPenStyle As Byte '��ǰʹ�õ�����
Public gcurPenWidth As Byte '��ǰʹ�õ��߿�
Public gcurFillColor As Long '��ǰʹ�õ����ɫ
Public gcurFillStyle As Integer '��ǰʹ�õ������ʽ

Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ

Public Const LONG_MAX = 2147483647 'Long�����ֵ
'======================================================================================================================
'API���岿��
Public Const WS_CHILD = &H40000000
Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
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
Public Const WM_UNDO = &H304
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'�������
Public Const EM_LINESCROLL = &HB6 'lngW=��������,lngL=��������
Public Const EM_SCROLL = &HB5 '������������
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'lngR(>=0)
Public Const EM_GETLINECOUNT = &HBA 'lngR(>=1,�����Զ��۵���)
Public Const EM_LINELENGTH = &HC1 '��һ��δ����ǰ��Ч
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB

Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
'API��ͼ
'---------------------------------------------------------------------------------------------------------------------
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
'Style
Public Const BS_HATCHED = 2
Public Const BS_NULL = 1
Public Const BS_SOLID = 0
'Hatch
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_VERTICAL = 1                '  |||||
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Const PS_NULL = 5
Public Const PS_SOLID = 0
Public Const PS_DOT = 2
Public Const PS_DASH = 1
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_INSIDEFRAME = 6
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
'======================================================================================================================
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000


Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        " Select A.ID,C.����ID,A.���,A.����,A.����,B.�û���" & _
        " From ��Ա�� A,�ϻ���Ա�� B,������Ա C" & _
        " Where A.ID = B.��ԱID And A.ID = C.��ԱID And C.ȱʡ = 1 And Upper(B.�û���) = USER"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�û��� = IIf(IsNull(rsTmp!�û���), "", rsTmp!�û���)
        gstrDBUser = UserInfo.�û���
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ArrayIcons(objLvw As ListView, Optional intBegin As Integer = 1, Optional blnShow As Boolean)
'���ܣ����ݵ�һ��ͼ���λ��������������ͼ��
    Dim i As Integer, t As Long
    Dim r As RECT

    Call GetClientRect(objLvw.hWnd, r)
    
    If blnShow Then
        If objLvw.ListItems(intBegin).Top < 30 Then
           objLvw.ListItems(intBegin).Top = 30
        ElseIf objLvw.ListItems(intBegin).Top + objLvw.ListItems(intBegin).Height > (r.Bottom - r.Top) * Screen.TwipsPerPixelY Then
            objLvw.ListItems(intBegin).Top = (r.Bottom - r.Top) * Screen.TwipsPerPixelY - objLvw.ListItems(intBegin).Height
        End If
    End If
    
    '�����ͼ��
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To objLvw.ListItems.Count
        With objLvw.ListItems(i)
            'Item��Width�������ֲ���,Left��ָͼ��
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t + .Height
        End With
    Next
    
    '�����ͼ��
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To 1 Step -1
        With objLvw.ListItems(i)
            'Item��Width�������ֲ���,Left��ָͼ��
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t - .Height
        End With
    Next
End Sub

Public Function OpenRecord(rsTmp As ADODB.Recordset, strSQL As String, ByVal strTitle As String, _
    Optional CursorType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, strTitle, strSQL)
    rsTmp.Open strSQL, gcnOracle, CursorType, LockType
    Call SQLTest
    
    Set OpenRecord = rsTmp
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        Case "Variant" '����ȷ����
            strLog = Replace(strLog, "[" & i & "]", "?")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '�����µĲ���
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "String" '�ַ�
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, 500, varValue)
        Case "Variant" '����ȷ����
        End Select
    Next

    'ִ�з��ؼ�¼��
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function

Public Sub ResetDrawStyle()
'���ܣ�ɾ����ǰ���õĻ��ʺͻ�ˢ
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen
End Sub

Public Sub SetDrawStyleFromValue(lngHDc As Long, PenColor As Long, PenStyle As Byte, PenWidth As Byte, FillColor As Long, FillStyle As Integer)
'���ܣ�����ָ��ֵ���õ�ǰ�Ļ��ʵĻ�ˢ
    Dim vBrush As LOGBRUSH
    Dim lngPen As Long, lngBrush As Long
    
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen
    
    '����
    lngPen = CreatePen(PenStyle, IIf(PenWidth < 1, 1, PenWidth), PenColor)
    glngPen = SelectObject(lngHDc, lngPen)
    
    '��ˢ
    vBrush.lbColor = FillColor
    If FillStyle = -1 Then
        vBrush.lbStyle = BS_NULL
    ElseIf FillStyle = -2 Then
        vBrush.lbStyle = BS_SOLID
    Else
        vBrush.lbStyle = BS_HATCHED
        vBrush.lbHatch = FillStyle
    End If
    lngBrush = CreateBrushIndirect(vBrush)
    glngBrush = SelectObject(lngHDc, lngBrush)
End Sub

Public Sub TextOut(objOut As Object, ByVal strOut As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal strFont As String, ByVal sngScale As Single)
'���ܣ���ָ���豸��ָ����Χ���������
'������strFont="����,�ֺ�,��ɫ,0000",sngScale=�������
'˵�����Զ�����,֧�ֻس�
    Dim arrFont() As String, arrLine() As String
    Dim lngWidth As Long, lngW As Long, i As Integer
    
    If Trim(Replace(strOut, vbCrLf, "")) = "" Then Exit Sub
    If strFont = "" Then Exit Sub
    
    arrFont = Split(strFont, ",")
    objOut.FontName = arrFont(0)
    objOut.FontSize = CSng(arrFont(1)) * sngScale
    objOut.ForeColor = CLng(arrFont(2))
    objOut.FontBold = Mid(arrFont(3), 1, 1) = "1"
    objOut.FontItalic = Mid(arrFont(3), 2, 1) = "1"
    objOut.FontUnderline = Mid(arrFont(3), 3, 1) = "1"
    objOut.FontStrikethru = Mid(arrFont(3), 4, 1) = "1"
    
    X1 = X1 * sngScale: Y1 = Y1 * sngScale
    X2 = X2 * sngScale: Y2 = Y2 * sngScale
        
    strOut = Replace(strOut, vbCrLf, "'")
    lngWidth = X2 - X1
    ReDim arrLine(0)
    For i = 1 To Len(strOut)
        If Mid(strOut, i, 1) = "'" Then
            lngW = 0
            Do While Mid(strOut, i, 1) = "'"
                ReDim Preserve arrLine(UBound(arrLine) + 1)
                i = i + 1
            Loop
        End If
        If i <= Len(strOut) Then
            lngW = lngW + objOut.TextWidth(Mid(strOut, i, 1))
            If lngW > lngWidth Then
                ReDim Preserve arrLine(UBound(arrLine) + 1)
                lngW = 0
            End If
            arrLine(UBound(arrLine)) = arrLine(UBound(arrLine)) & Mid(strOut, i, 1)
        End If
    Next
    objOut.CurrentY = Y1 + 2
    For i = 0 To UBound(arrLine)
        objOut.CurrentX = X1 + 2
        objOut.Print arrLine(i)
    Next
End Sub

Public Function ReadCaseMap(lngID As Long) As StdPicture
'���ܣ����ݱ��ͼID����ͼ�ζ���
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ͼ�� From �������ͼ Where Ԫ��ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISCore", lngID)
        
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!ͼ��) Then Exit Function
    
    On Error GoTo 0
    
    intFile = FreeFile
    strFile = CurDir & "\zlNewPicture" & Timer & ".pic"
    
    Open strFile For Binary As intFile
    
    lngFileSize = rsTmp.Fields("ͼ��").ActualSize
    ReDim arrBin(lngFileSize - 1) As Byte
    arrBin() = rsTmp.Fields("ͼ��").GetChunk(lngFileSize)
    Put intFile, , arrBin()
    Close intFile
    
    Set ReadCaseMap = VB.LoadPicture(strFile)
    Kill strFile
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadVoiceToFile(lngID As Long) As String
'���ܣ����ݲ�����¼ID���������ļ�
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ���˲���¼�� Where ������¼ID=" & lngID
    OpenRecord rsTmp, strSQL, "mdlCISCore"
        
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!����) Then Exit Function
    
    intFile = FreeFile
    strFile = CurDir & "\zlNewPicture" & Timer & ".Mp3"
    
    Open strFile For Binary As intFile
    
    lngFileSize = rsTmp.Fields("����").ActualSize
    ReDim arrBin(lngFileSize - 1) As Byte
    arrBin() = rsTmp.Fields("����").GetChunk(lngFileSize)
    Put intFile, , arrBin()
    Close intFile
    
    ReadVoiceToFile = strFile
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    If Len(strFile) > 0 Then Close intFile: Kill strFile
    Call SaveErrLog
End Function

Public Sub ShowCaseMap(objCaseMap As StdPicture, objMapItems As MapItems, objDraw As Object, _
    Optional x As Long, Optional y As Long, Optional W As Long, Optional H As Long)
'���ܣ���ʾ�������ͼ����
'������objDraw=��ʾ��Ŀ�����,����ScaleMode����ΪPixel
'      objMapItems=�����е�ǰ��Ŀ�ı��ͼ����
'      X,Y,W,H=��ʾ��Ŀ�귶Χ,���Բ�ָ��,��λΪPixel
    Dim arrTmp() As String, arrXY() As POINTAPI
    Dim i As Integer, j As Integer
    Dim lngW As Long, lngH As Long 'ͼƬ�ߴ�
    Dim sngScale As Single
        
    Screen.MousePointer = 11
    LockWindowUpdate objDraw.hWnd
    
    'ȷ��ͼƬ�ߴ缰��ʾ����
    objDraw.ScaleMode = vbPixels
    lngW = objDraw.ScaleX(objCaseMap.Width, vbHimetric, vbPixels) '��HiMetricת��ΪPixel
    lngH = objDraw.ScaleY(objCaseMap.Height, vbHimetric, vbPixels)
    If W = 0 Then W = objDraw.ScaleWidth
    If H = 0 Then H = objDraw.ScaleHeight
    If W / lngW < H / lngH Then
        sngScale = W / lngW
    Else
        sngScale = H / lngH
    End If
    
    objDraw.Cls
    objDraw.PaintPicture objCaseMap, x, y, lngW * sngScale, lngH * sngScale
            
    '������Ԫ��
    For i = 1 To objMapItems.Count
        With objMapItems(i)
            If .���� <> 0 Then
                Call SetDrawStyleFromValue(objDraw.hDC, .����ɫ, .����, .�߿� * sngScale, .���ɫ, .��䷽ʽ)
            End If
            Select Case .����
                Case 0 '�ı�
                    Call TextOut(objDraw, .����, (.X1 * sngScale + x) / sngScale, (.Y1 * sngScale + y) / sngScale, (.X2 * sngScale + x) / sngScale, (.Y2 * sngScale + y) / sngScale, .����, sngScale)
                Case 1 '����
                    MoveToEx objDraw.hDC, .X1 * sngScale + x, .Y1 * sngScale + y, 0
                    LineTo objDraw.hDC, .X2 * sngScale + x, .Y2 * sngScale + y
                Case 2 '����
                    arrTmp = Split(.�㼯, ";")
                    For j = 0 To UBound(arrTmp)
                        ReDim Preserve arrXY(j)
                        arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) * sngScale + x
                        arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) * sngScale + y
                    Next
                    Polyline objDraw.hDC, arrXY(0), UBound(arrXY) + 1
                Case 3 '����
                    Rectangle objDraw.hDC, .X1 * sngScale + x, .Y1 * sngScale + y, .X2 * sngScale + x, .Y2 * sngScale + y
                Case 4 '�����
                    arrTmp = Split(.�㼯, ";")
                    For j = 0 To UBound(arrTmp)
                        ReDim Preserve arrXY(j)
                        arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) * sngScale + x
                        arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) * sngScale + y
                    Next
                    Polygon objDraw.hDC, arrXY(0), UBound(arrXY) + 1
                Case 5 'Բ
                    Ellipse objDraw.hDC, .X1 * sngScale + x, .Y1 * sngScale + y, .X2 * sngScale + x, .Y2 * sngScale + y
            End Select
        End With
    Next
    objDraw.Refresh
    
    Call ResetDrawStyle
    
    LockWindowUpdate 0
    Screen.MousePointer = 0
End Sub

Public Function EditFlag(frmParent As Object, varԪ�� As Variant, Optional Flags As Variant, Optional blnViewOnly As Boolean) As MapItems
'���ܣ��ڵ�����ģ̬�����б༭��鿴ָ���Ĳ������ͼ
'������frmParnet=���ø�����
'      varԪ��=���ͼԪ�صı���(�ַ���)��ID(������)
'      Flags=Long�ͣ�Ҫ�޸ĵ�"���˲�������"�б��ͼԪ�ض�Ӧ��ID��
'            MapItems��Ҫ��ʾ�ı�ע��
'            ������������ʾ������ע
'      blnViewOnly=�Ƿ�ֻ�鿴�����ܱ༭
'���أ�Mapitems
'      ȡ���༭��鿴ģʽ����Empty(Not isArray)��
    Dim frmNew As frmMapEdit
    Dim rsTmp As New ADODB.Recordset
    Dim arrSQL() As Variant, strSQL As String
    
    Dim objCaseMap As StdPicture, i As Long
    Dim objMapItems As New MapItems, objMapItem As MapItem
    Dim lngMapID As Long, strMapName As String
    
    Dim iMin As Long, iMax As Long, aItems() As String
    Dim strFont As String, strContent As String, strDots As String
    
    On Error GoTo errH
        
    '��ȡ���ͼԪ�ص�����
    If TypeName(varԪ��) = "String" Then
        strSQL = "Select * From ����Ԫ��Ŀ¼ Where ����=[1]"
    Else
        strSQL = "Select * From ����Ԫ��Ŀ¼ Where ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", CStr(varԪ��))
    If rsTmp.EOF Then Exit Function '����Ҫ��ͼ�α���
    
    lngMapID = rsTmp!ID
    strMapName = rsTmp!���� & IIf(IsNull(rsTmp!˵��), "", "(" & rsTmp!˵�� & ")")
    
    Set objCaseMap = ReadCaseMap(rsTmp!ID)
    If objCaseMap Is Nothing Then Exit Function '����Ҫ��ͼ�α���
    
    '��ȡ���ͼ�ı�ע����
    If IsEmpty(Flags) Then Flags = 0
    
    If TypeName(Flags) = "Long" Then
        If Flags <> 0 Then
            strSQL = "Select * From ���˲������ͼ Where ����ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", Flags)
            Do While Not rsTmp.EOF
                With rsTmp
                    objMapItems.Add !����, zlCommFun.NVL(!����), _
                        IIf(IsNull(!����), IIf(!���� = 0, "����,9,0,0000", ""), !����), _
                        zlCommFun.NVL(!�㼯), zlCommFun.NVL(!X1, 0), _
                        zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                        zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!���ɫ, &HFFFFFF), _
                        zlCommFun.NVL(!��䷽ʽ, -1), zlCommFun.NVL(!����ɫ, 0), _
                        zlCommFun.NVL(!����, 0), zlCommFun.NVL(!�߿�, 1), "_" & objMapItems.Count + 1
                End With
                rsTmp.MoveNext
            Loop
        End If
    Else
        For i = 1 To Flags.Count
            Set objMapItem = Flags(i)
            '"����,'����','����','�㼯',X1,Y1,X2,Y2,���ɫ,��䷽ʽ,����ɫ,����,�߿�"
            With objMapItem
                objMapItems.Add .����, .����, .����, .�㼯, _
                    .X1, .Y1, .X2, .Y2, .���ɫ, .��䷽ʽ, _
                    .����ɫ, .����, .�߿�, "_" & objMapItems.Count + 1
            End With
        Next
    End If
    
    On Error GoTo 0
    
    Set frmNew = New frmMapEdit
    frmNew.mblnModi = Not blnViewOnly
    frmNew.mlngMapID = lngMapID
    frmNew.mstrMapName = strMapName
    Set frmNew.mobjCaseMap = objCaseMap
    Set frmNew.mobjMapItems = objMapItems
    frmNew.Show 1, frmParent
    
    If gblnOK Then
        Set EditFlag = frmNew.mobjMapItems
'        For i = 1 To frmNew.mobjMapItems.Count
'            Set objMapItem = frmNew.mobjMapItems(i)
'            '"����,'����','����','�㼯',X1,Y1,X2,Y2,���ɫ,��䷽ʽ,����ɫ,����,�߿�"
'            With objMapItem
'                EditFlag.Add .����, .����, .����, .�㼯, _
'                    .X1, .Y1, .X2, .Y2, .���ɫ, .��䷽ʽ, _
'                    .����ɫ, .����, .�߿�
'            End With
'        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowFlagInOjbect(objDraw As Object, varԪ�� As Variant, Optional Flags As Variant, Optional x As Long, Optional y As Long, Optional W As Long, Optional H As Long, Optional blnMoved As Boolean = False)
'���ܣ���ָ���Ķ���(PictureBox��Form)����ʾ���ͼ
'������objDraw=PictureBox�������,����ScaleMode����ΪPixel
'      varԪ��=���ͼԪ�صı���(�ַ���)��ID(������)
'      Flags=Long�ͣ�"���˲�������"�б��ͼԪ�ض�Ӧ��ID��
'            MapItems��Ҫ��ʾ�ı�ע��
'            �������,����ʾ���ͼ����
'      X,Y,W,H=��ʾ��Ŀ��ͻ��˷�Χ,���Բ�ָ��,��λΪPixel
'˵�����������øú������д�ӡ���(��Ϊ��API��ͼ,��˲���ֱ�ӽ�objDrawָ��Ϊ��ӡ��,������PictureBox�ϰ�һ�����������,ȡPictureBox.Image�������ӡ��)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objCaseMap As StdPicture, objMapItems As New MapItems
    
    Dim i As Long, iMin As Long, iMax As Long, aItems() As String
    Dim strFont As String, strContent As String, strDots As String
    
    On Error GoTo errH
        
    '��ȡ���ͼԪ�ص�����
    If TypeName(varԪ��) = "String" Then
        strSQL = "Select * From ����Ԫ��Ŀ¼ Where ����=[1]"
    Else
        strSQL = "Select * From ����Ԫ��Ŀ¼ Where ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", CStr(varԪ��))
    If rsTmp.EOF Then Exit Sub '����Ҫ��ͼ�α���
    
    Set objCaseMap = ReadCaseMap(rsTmp!ID)
    If objCaseMap Is Nothing Then Exit Sub '����Ҫ��ͼ�α���
    
    '��ȡ���ͼ�ı�ע����
    If IsEmpty(Flags) Then Flags = 0
    
    If TypeName(Flags) = "Long" Then
        If Flags <> 0 Then
            strSQL = "Select * From ���˲������ͼ Where ����ID=[1]"
            If blnMoved Then
                strSQL = Replace(strSQL, "���˲������ͼ", "H���˲������ͼ")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", Flags)
            Do While Not rsTmp.EOF
                With rsTmp
                    objMapItems.Add !����, zlCommFun.NVL(!����), _
                        IIf(IsNull(!����), IIf(!���� = 0, "����,9,0,0000", ""), !����), _
                        zlCommFun.NVL(!�㼯), zlCommFun.NVL(!X1, 0), _
                        zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                        zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!���ɫ, &HFFFFFF), _
                        zlCommFun.NVL(!��䷽ʽ, -1), zlCommFun.NVL(!����ɫ, 0), _
                        zlCommFun.NVL(!����, 0), zlCommFun.NVL(!�߿�, 1)
                End With
                rsTmp.MoveNext
            Loop
        End If
    Else
        Set objMapItems = Flags
    End If
    
    On Error GoTo 0
    
    Call ShowCaseMap(objCaseMap, objMapItems, objDraw, x, y, W, H)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'���没�˱��ͼ
Public Sub SaveFlag(ByVal ContentID As Long, Flags As Variant, DataConn As Connection)
    Dim i As Long, iMin As Long, iMax As Long
    Dim strSQL As String
    Dim objMapItem As MapItem
    
    If IsEmpty(Flags) Then Exit Sub
    
    If TypeName(Flags) = "MapItems" Then
        For i = 1 To Flags.Count
            Set objMapItem = Flags(i)
            '"����,'����','����','�㼯',X1,Y1,X2,Y2,���ɫ,��䷽ʽ,����ɫ,����,�߿�"
            With objMapItem
                strSQL = .���� & ",'" & .���� & "','" & .���� & "','" & .�㼯 & "'," & _
                    .X1 & "," & .Y1 & "," & .X2 & "," & .Y2 & "," & .���ɫ & "," & .��䷽ʽ & "," & _
                    .����ɫ & "," & .���� & "," & .�߿�
            End With
            DataConn.Execute "ZL_���˲������ͼ_SAVE(" & ContentID & "," & strSQL & ")", , adCmdStoredProc
        Next
    Else
        If UBound(Flags) = -1 Then Exit Sub
    
        iMin = LBound(Flags): iMax = UBound(Flags)
        For i = iMin To iMax
            DataConn.Execute "ZL_���˲������ͼ_SAVE(" & ContentID & "," & Flags(i) & ")", , adCmdStoredProc
        Next
    End If
End Sub

Public Function GetMap(ByVal lng����ID As Long, ByVal picDraw As PictureBox, Optional blnMoved As Boolean = False) As StdPicture
    Dim rsTmp As New ADODB.Recordset
    Dim objFlags As MapItems
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select a.Ԫ�ر���,b.ID From ���˲������� a,����Ԫ��Ŀ¼ b Where a.ID=[1] And a.Ԫ�ر���=b.����"
    If blnMoved Then
        strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISCore", lng����ID)
    If Not rsTmp.EOF Then
        Set objFlags = GetMapItems(lng����ID, blnMoved)
        With picDraw
            .AutoRedraw = True: .ScaleMode = vbPixels: .Cls: .BackColor = RGB(255, 255, 255)
            Set .Picture = ReadCaseMap(rsTmp(1))
            .Width = .ScaleX(.Picture.Width, vbHimetric, vbTwips): .Height = .ScaleY(IIf(.Picture.Height = 0, 1, .Picture.Height), vbHimetric, vbTwips)
            .Width = IIf(.Width > 10000, 10000, .Width): .Height = .Height * .Width / .ScaleX(IIf(.Picture.Width = 0, 1, .Picture.Width), vbHimetric, vbTwips)
            .Cls: Set .Picture = Nothing
        End With
        ShowFlagInOjbect picDraw, CStr(rsTmp(0)), objFlags, blnMoved:=blnMoved
        Set GetMap = picDraw.Image
    Else
        Set GetMap = New StdPicture
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetMapItems(lngItemID As Long, Optional blnMoved As Boolean = False) As MapItems
'���ܣ���ȡ��Ƕ���
'������lngItemID��"���˲�������"�б��ͼԪ�ض�Ӧ��ID��
'���أ�Mapitems
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    Set GetMapItems = New MapItems
    
    On Error GoTo DBError
    strSQL = "Select * From ���˲������ͼ Where ����ID=[1]"
    If blnMoved Then
        strSQL = Replace(strSQL, "���˲������ͼ", "H���˲������ͼ")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", lngItemID)
    Do While Not rsTmp.EOF
        With rsTmp
            GetMapItems.Add !����, zlCommFun.NVL(!����), _
                IIf(IsNull(!����), IIf(!���� = 0, "����,9,0,0000", ""), !����), _
                zlCommFun.NVL(!�㼯), zlCommFun.NVL(!X1, 0), _
                zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!���ɫ, &HFFFFFF), _
                zlCommFun.NVL(!��䷽ʽ, -1), zlCommFun.NVL(!����ɫ, 0), _
                zlCommFun.NVL(!����, 0), zlCommFun.NVL(!�߿�, 1)
        End With
        rsTmp.MoveNext
    Loop
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

'Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
''����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
'    If InStr(strInput, "'") > 0 Or InStr(strInput, ";") > 0 Or InStr(strInput, ",") > 0 Or InStr(strInput, "`") > 0 Or InStr(strInput, """") > 0 Then
'        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
'        Exit Function
'    End If
'    If intMax > 0 Then
'        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
'            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
'            Exit Function
'        End If
'    End If
'
'    StrIsValid = True
'End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "���ַ���", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function Check�Ƿ����(strSource As String, strTarge As String) As Boolean
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    Check�Ƿ���� = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    Check�Ƿ���� = True
End Function

Public Sub SelectRow(mshObject As Object, Optional ByVal BackColor As Long = &H8000000D, Optional ByVal ForeColor As Long = &H8000000E)
    Dim i As Integer
    Dim blnPre As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshObject
        blnPre = .Redraw
        intRow = .Row
        intCol = .Col
        .Redraw = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = BackColor
            .CellForeColor = ForeColor
        Next
        
        .Row = intRow
        .Col = intCol
        .Redraw = blnPre
    End With
End Sub

Public Sub UnSelectRow(mshObject As Object, Optional lngColorSave As Long = 0)
    Dim i As Integer
    Dim blnPre As Boolean
    
    With mshObject
        blnPre = .Redraw
        .Redraw = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = .BackColor
            .CellForeColor = lngColorSave
        Next
        .Redraw = blnPre
    End With
End Sub
'�滻�����������Ԫ��
Public Function GetSpecValue(ItemName As String, sPatientID As String, sPageID As String, iPatientType As Integer) As String
    'sPatientID������ID
    'sPageID����ҳID��Һŵ���
    'iPatientType��0=���1=סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTemp As String
    
    If Len(Trim(sPatientID)) = 0 Then GetSpecValue = "": Exit Function
    strSQL = ""
    Err = 0: On Error GoTo DBError
    
    Select Case ItemName
    Case "��дǩ��", "-1", "SXR"
        GetSpecValue = UserInfo.����: Exit Function
    Case "��ǰ����", "-2", "DQRQ"
        strSQL = "Select to_Char(SysDate,'YYYY-MM-DD') From Dual"
    Case "��ǰʱ��", "-3", "DQSJ"
        strSQL = "Select to_Char(SysDate,'YYYY-MM-DD HH24:MI:SS') From Dual"
    Case "����", "XM"
        strSQL = "Select ���� From ������Ϣ Where ����ID=" & sPatientID
    Case "�Ա�", "XB"
        strSQL = "Select �Ա� From ������Ϣ Where ����ID=" & sPatientID
    Case "����", "NL"
        strSQL = "Select ���� From ������Ϣ Where ����ID=" & sPatientID
    Case "ְҵ", "ZY"
        strSQL = "Select ְҵ From ������Ϣ Where ����ID=" & sPatientID
    Case "����", "MZ"
        strSQL = "Select ���� From ������Ϣ Where ����ID=" & sPatientID
    Case "����", "GJ"
        strSQL = "Select ���� From ������Ϣ Where ����ID=" & sPatientID
    Case "����״��", "HYZK"
        strSQL = "Select ����״�� From ������Ϣ Where ����ID=" & sPatientID
    Case "��������", "CSRQ"
        strSQL = "Select to_char(��������,'YYYY-MM-DD') From ������Ϣ Where ����ID=" & sPatientID
    Case "�����ص�", "CSDD"
        strSQL = "Select �����ص� From ������Ϣ Where ����ID=" & sPatientID
    Case "���֤��", "SFZH"
        strSQL = "Select ���֤�� From ������Ϣ Where ����ID=" & sPatientID
    Case "���", "SF"
        strSQL = "Select ��� From ������Ϣ Where ����ID=" & sPatientID
    Case "ѧ��", "XL"
        strSQL = "Select ѧ�� From ������Ϣ Where ����ID=" & sPatientID
    Case "��ͥ��ַ", "JTDZ"
        strSQL = "Select ��ͥ��ַ From ������Ϣ Where ����ID=" & sPatientID
    Case "��ͥ�绰", "JTDH"
        strSQL = "Select ��ͥ�绰 From ������Ϣ Where ����ID=" & sPatientID
    Case "������λ", "GZDW"
        strSQL = "Select ������λ From ������Ϣ Where ����ID=" & sPatientID
    Case "��λ�绰", "DWDH"
        strSQL = "Select ��λ�绰 From ������Ϣ Where ����ID=" & sPatientID
    Case "�����", "MZH"
        strSQL = "Select ����� From ������Ϣ Where ����ID=" & sPatientID
    Case "���￨��", "JZKH"
        strSQL = "Select ���￨�� From ������Ϣ Where ����ID=" & sPatientID
    Case "�������", "JZKS"
        strSQL = "Select D.����" & _
                " From ���ű� D," & _
                "      (Select Distinct ���˿���ID" & _
                "      From ���˷��ü�¼" & _
                "      Where ����id=" & sPatientID & " And No='" & sPageID & "'" & _
                "            And ��¼����=4 And ��¼״̬=1 And �շ����='1') R" & _
                " Where D.Id=R.���˿���ID"
    Case "����ʱ��", "JZSJ"
        strSQL = "Select Distinct to_char(����ʱ��,'YYYY-MM-DD HH24:MI:SS')" & _
                " From ���˷��ü�¼" & _
                " Where ����id=" & sPatientID & " And No='" & sPageID & "'" & _
                "       And ��¼����=4 And ��¼״̬=1 And �շ����='1'"
    Case "�Ƿ���", "SFJZ"
        strSQL = "Select Distinct nvl(�Ӱ��־,0)" & _
                " From ���˷��ü�¼" & _
                " Where ����id=" & sPatientID & " And No='" & sPageID & "'" & _
                "       And ��¼����=4 And ��¼״̬=1 And �շ����='1'"
    Case "סԺ��", "ZYH"
        strSQL = "Select סԺ�� From ������Ϣ Where ����ID=" & sPatientID
    Case "סԺ����", "ZYCS"
        strSQL = "Select סԺ���� From ������Ϣ Where ����ID=" & sPatientID
    Case "��Ժ����", "RYRQ"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select to_char(��Ժ����,'YYYY-MM-DD')" & _
                " From ������ҳ" & _
                " Where ����id=" & sPatientID & " And ��ҳid=" & sPageID
        End If
    Case "��Ժ����", "CYRQ"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select to_char(��Ժ����,'YYYY-MM-DD')" & _
                " From ������ҳ" & _
                " Where ����id=" & sPatientID & " And ��ҳid=" & sPageID
        End If
    Case "סԺĿ��", "ZYMD"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select סԺĿ��" & _
                " From ������ҳ" & _
                " Where ����id=" & sPatientID & " And ��ҳid=" & sPageID
        End If
    Case "��Ժ����", "RYKS"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select D.����" & _
                " From ���ű� D," & _
                "      (Select Distinct ��Ժ����ID" & _
                "       From ������ҳ" & _
                "       Where ����id=" & sPatientID & " And ��ҳid=" & sPageID & ") P" & _
                " Where D.Id=P.��Ժ����ID"
        End If
    Case "��Ժ����", "RYBQ"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select D.����" & _
                " From ���ű� D," & _
                "      (Select Distinct ��Ժ����ID" & _
                "       From ������ҳ" & _
                "       Where ����id=" & sPatientID & " And ��ҳid=" & sPageID & ") P" & _
                " Where D.Id=P.��Ժ����ID"
        End If
    Case "��ǰ����", "DQCH"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select ��Ժ����" & _
                " From ������ҳ" & _
                " Where ����id=" & sPatientID & " And ��ҳid=" & sPageID
        End If
    Case "��ǰ����", "DQBQ"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select D.����" & _
                " From ���ű� D," & _
                "      (Select Distinct ��ǰ����ID" & _
                "       From ������ҳ" & _
                "       Where ����id=" & sPatientID & " And ��ҳid=" & sPageID & ") P" & _
                " Where D.Id=P.��ǰ����ID"
        End If
    Case "��ǰ����", "DQKS"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select D.����" & _
                " From ���ű� D," & _
                "      (Select Distinct ��Ժ����ID" & _
                "       From ������ҳ" & _
                "       Where ����id=" & sPatientID & " And ��ҳid=" & sPageID & ") P" & _
                " Where D.Id=P.��Ժ����ID"
        End If
    Case "��ǰ����", "DQBK"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select ��ǰ����" & _
                " From ������ҳ" & _
                " Where ����id=" & sPatientID & " And ��ҳid=" & sPageID
        End If
    Case "סԺҽʦ", "ZYYS"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select סԺҽʦ" & _
                " From ������ҳ" & _
                " Where ����id=" & sPatientID & " And ��ҳid=" & sPageID
        End If
    Case "���λ�ʿ", "ZRHS"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select ���λ�ʿ" & _
                " From ������ҳ" & _
                " Where ����id=" & sPatientID & " And ��ҳid=" & sPageID
        End If
    Case "������", "ZHZD"
        strSQL = "Select �������" & _
                " From ������ϼ�¼" & _
                " Where ����id=" & sPatientID & " And ȡ��ʱ�� Is Null" & _
                "       And ��¼���� In (" & _
                "           Select Max(��¼����)" & _
                "           From ������ϼ�¼" & _
                " Where ����id=" & sPatientID & " And ȡ��ʱ�� Is Null)"
        With rsTmp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "��ȡ�滻��", strSQL): rsTmp.Open strSQL, gcnOracle, adOpenKeyset: Call SQLTest
            strTemp = ""
            Do While Not .EOF
                strTemp = strTemp & vbCrLf & .AbsolutePosition & "." & .Fields(0).Value
                .MoveNext
            Loop
            If strTemp <> "" Then strTemp = Mid(strTemp, 3)
            GetSpecValue = strTemp: Exit Function
        End With
    Case "����ҩ��", "GMYW"
        strSQL = "Select ҩ����" & _
                " From ���˹�����¼" & _
                " Where ����id=" & sPatientID & " And ���=1" & _
                "       And ��¼ʱ�� In (" & _
                "           Select Max(��¼ʱ��)" & _
                "           From ���˹�����¼" & _
                " Where ����id=" & sPatientID & " And ���=1)"
        With rsTmp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "��ȡ�滻��", strSQL): rsTmp.Open strSQL, gcnOracle, adOpenKeyset: Call SQLTest
            strTemp = ""
            Do While Not .EOF
                strTemp = strTemp & vbCrLf & .AbsolutePosition & "." & .Fields(0).Value
                .MoveNext
            Loop
            If strTemp <> "" Then strTemp = Mid(strTemp, 3)
            GetSpecValue = strTemp: Exit Function
        End With
    End Select
    
    If strSQL = "" Then GetSpecValue = "": Exit Function
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "��ȡ�滻��", strSQL): rsTmp.Open strSQL, gcnOracle, adOpenKeyset: Call SQLTest
        If .EOF Or .BOF Then
            GetSpecValue = ""
        Else
            GetSpecValue = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
        End If
    End With
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

'�滻�����ַ����еı���
Public Function ReplaceString(strSource As String, sPatientID As String, sPageID As String, iPatientType As Integer, _
    Optional strVariableBegin As String = "{{", Optional strVariableEnd As String = "}}") As String
    Dim iLen1 As Integer, iLen2 As Integer
    Dim iStrPoint As Long, iStrLength As Long
    Dim iVariableBeginPos As Long, iVariableEndPos As Long
    Dim strVariable As String, strReturn As String
    iLen1 = Len(strVariableBegin): iLen2 = Len(strVariableEnd)
    
    ReplaceString = strSource
    iStrPoint = 1: iStrLength = Len(ReplaceString)
    Do While iStrPoint <= iStrLength
        If iVariableBeginPos > 0 Then
            iVariableEndPos = InStr(iVariableBeginPos + iLen1, ReplaceString, strVariableEnd)
            If iVariableEndPos = 0 Then
                Exit Do
            Else
                strVariable = Mid(ReplaceString, _
                    iVariableBeginPos + iLen1, iVariableEndPos - (iVariableBeginPos + iLen1))
                strReturn = GetSpecValue(strVariable, sPatientID, sPageID, iPatientType)
                
                ReplaceString = Replace(ReplaceString, strVariableBegin + strVariable + strVariableEnd, strReturn)
                iStrPoint = iVariableBeginPos + Len(strReturn)
                iStrLength = Len(ReplaceString)
                
                iVariableBeginPos = 0
            End If
        Else
            iVariableBeginPos = InStr(iStrPoint, ReplaceString, strVariableBegin)
            If iVariableBeginPos = 0 Then Exit Do
        End If
    Loop
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

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

Public Function GetFileId(ByVal lngWritId As Long) As Long
'���ܣ���ȡ�û�ָ���ļ��Ĳ��˲�����¼���Ա���������Ȳ���
'������lngWritId-��Ҫ���ҵĲ����ļ�
    With frmWritImp
        .lblWrit.Tag = lngWritId
        .Show 1
        GetFileId = .lngFileId
        Unload frmWritImp
    End With
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
'����ȡΪ��������
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

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

Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'���ܣ�ȥ��TextBox��Ĭ���Ҽ��˵�
    If msg <> WM_CONTEXTMENU Then
        WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
    End If
End Function

Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strName As String = "")
'���ܣ���PictureBoxģ���3Dƽ�水ť
'������intStyle:0=ƽ��,-1=����,1=͹��
    
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Sub PrintDiagReport(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, objParent As Object, Optional ByVal PrtMode As Integer = 2, Optional ObjPic As Object = Nothing, Optional blnMoved As Boolean = False)
'��ӡ���ﱨ��
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsImages As ADODB.Recordset
    Dim strRptName As String
    Dim aImages(1, 8) As Variant, aFlagImages(1, 8) As Variant, i As Integer
    Dim strTempPath As String, lngBuffSize As Long
    Dim intReportFormatItem As Integer
    
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strNO As String, lng��¼���� As Long
    Dim iTmpFileCount As Integer, iFlagCount As Integer
    Dim objImages As New DicomImages, intRows As Integer, intCols As Integer, objAssembleImage As New DicomImage
    
    On Error GoTo DBError
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    strTempPath = Mid(strTempPath, 1, InStrRev(strTempPath, "\"))
    
    strSQL = "Select A.NO,A.��¼����,'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2'" & _
        " From ����ҽ������ A,���˲�����¼ B,�����ļ�Ŀ¼ C" & _
        " Where A.����ID=B.ID And B.�ļ�ID=C.ID And A.ҽ��ID=[1] And A.���ͺ�=[2]"
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "��鱨��", lngҽ��ID, lng���ͺ�)
    If rsTmp.EOF Then
        MsgBox "������δ��д���棬���ܴ�ӡ��", vbInformation, gstrSysName
    Else
        strRptName = rsTmp(2): strNO = rsTmp(0): lng��¼���� = rsTmp(1)
        If PrtMode = 2 Then
            If Not ReportPrintSet(gcnOracle, glngSys, strRptName, objParent) Then Exit Sub
            intReportFormatItem = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\zl9Report\LocalSet\" & strRptName, "Format", 1)
        Else
            intReportFormatItem = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\zl9Report\frmReport" & strRptName, "��ʽ", 1)
        End If
        'PACS��Ӱ��ͼƬ
        strSQL = "Select A.�û���1,A.����1,A.Host1,A.Root1,A.URL1,A.�û���2,A.����2,A.Host2,A.Root2,A.URL2," & _
            "a.�豸��1,a.�豸��2,A.NO,A.��¼���� From" & _
            " (Select E.IP��ַ As Host1,'/'||E.FtpĿ¼||'/' as Root1,e.�豸�� as �豸��1," & _
            "Decode(D.��������,Null,'',to_Char(D.��������,'YYYYMMDD')||'/')" & _
            "||D.���UID||'/'||A.ͼ���ļ� As URL1," & _
            "F.IP��ַ As Host2,'/'||f.FtpĿ¼||'/' as Root2," & _
            "Decode(D.��������,Null,'',to_Char(D.��������,'YYYYMMDD')||'/')" & _
            "||D.���UID||'/'||A.ͼ���ļ� As URL2,f.�豸�� as �豸��2," & _
            "C.NO,C.��¼����,E.�û��� as �û���1,E.���� as ����1,F.�û��� as �û���2,F.���� as ����2, Rownum As Seq " & _
            " From ���˲����ⲿͼ A,���˲������� B,����ҽ������ C,Ӱ�����¼ D,Ӱ���豸Ŀ¼ E,Ӱ���豸Ŀ¼ F" & _
            " Where A.����ID=B.ID And B.������¼ID=C.����ID And C.ҽ��ID=D.ҽ��ID" & _
            " And C.���ͺ�=D.���ͺ� And D.λ��һ=E.�豸��(+) and d.λ�ö�=F.�豸��(+)" & _
            " And C.ҽ��ID=[1] And C.���ͺ�=[2]" & _
            " Order By A.���) A"
        If blnMoved Then
            strSQL = Replace(strSQL, "���˲����ⲿͼ", "H���˲����ⲿͼ")
            strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "��鱨��", lngҽ��ID, lng���ͺ�, intReportFormatItem)
        iTmpFileCount = rsTmp.RecordCount
        strSQL = "Select A.���,B.����,B.W,B.H" & _
            " From zlReports A,zlRPTItems B,�����ļ�Ŀ¼ C,����ҽ����¼ D,���Ƶ���Ӧ�� E" & _
            " Where A.ID=B.����ID And A.���='ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2'" & _
            " And C.ID=E.�����ļ�ID And D.������ĿID=E.������ĿID And Nvl(B.����,0)=1 And B.����=11" & _
            " And E.Ӧ�ó���=D.������Դ And D.ID=[1]" & _
            " And B.���� Not Like '���%' and b.��ʽ��=[3]" & _
            " Order BY b.����"
        If blnMoved Then
            strSQL = Replace(strSQL, "���˲����ⲿͼ", "H���˲����ⲿͼ")
            strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rsImages = OpenSQLRecord(strSQL, "��鱨��", lngҽ��ID, lng���ͺ�, intReportFormatItem)
        If rsImages.RecordCount = 1 Then
            'ͼ���Ű�
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & objFileSystem.GetParentFolderName(rsTmp("URL1"))
                strTmpFile = Replace(strTmpFile, "/", "\")
                MkLocalDir strTmpFile
                strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(rsTmp("URL1"))
                
                If Dir(strTmpFile, vbDirectory) = "" Then
                    If strDeviceNO1 <> rsTmp("�豸��1") Then
                        strDeviceNO1 = rsTmp("�豸��1")
                        Inet1.FuncFtpConnect NVL(rsTmp("Host1")), NVL(rsTmp("�û���1")), NVL(rsTmp("����1"))
                    End If
                    
                    If strDeviceNO2 <> rsTmp("�豸��2") Then
                        strDeviceNO2 = rsTmp("�豸��2")
                        Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("�û���2")), NVL(rsTmp("����2"))
                    End If
                    
                    If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
                        Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
                    End If
                End If
'                objAssembleImage.FileImport strTmpFile, "JPEG"
'                objImages.Add objAssembleImage
                
                objImages.AddNew
                objImages(objImages.Count).FileImport strTmpFile, "JPEG"
                
                rsTmp.MoveNext
            Next
            If objImages.Count > 0 Then
                ResizeRegion i, rsImages("W"), rsImages("H"), intRows, intCols
                Set objAssembleImage = funAssembleImage(objImages, intRows, intCols, rsImages("H"), rsImages("W"))
                strTmpFile = objFileSystem.GetParentFolderName(strTmpFile) & "\" & objFileSystem.GetTempName
                objAssembleImage.FileExport strTmpFile, "JPEG"
                    
                aImages(0, 0) = rsImages("����")
                aImages(1, 0) = strTmpFile
            End If
            For i = 1 To 8
                aImages(0, i) = "1"
                aImages(1, i) = "1"
            Next
        Else
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                If rsImages.EOF Then Exit For
                
    '            strTmpFile = strTempPath & objFileSystem.GetFileName(rsTmp(3))
                
                strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & objFileSystem.GetParentFolderName(rsTmp("URL1"))
                strTmpFile = Replace(strTmpFile, "/", "\")
                MkLocalDir strTmpFile
                strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(rsTmp("URL1"))
                
                If Dir(strTmpFile, vbDirectory) = "" Then
                    If strDeviceNO1 <> rsTmp("�豸��1") Then
                        strDeviceNO1 = rsTmp("�豸��1")
                        Inet1.FuncFtpConnect NVL(rsTmp("Host1")), NVL(rsTmp("�û���1")), NVL(rsTmp("����1"))
                    End If
                    
                    If strDeviceNO2 <> rsTmp("�豸��2") Then
                        strDeviceNO2 = rsTmp("�豸��2")
                        Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("�û���2")), NVL(rsTmp("����2"))
                    End If
                    
                    'Inet.strIPAddress = Nvl(rsTmp(2)): Inet.strUser = Nvl(rsTmp(6)): Inet.strPsw = Nvl(rsTmp(7))
                    If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
    '                    strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
    '                        Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
                    End If
                End If
                    
                aImages(0, i) = rsImages("����")
                aImages(1, i) = strTmpFile
                rsImages.MoveNext
                rsTmp.MoveNext
            Next
            For i = rsTmp.RecordCount To 8
                aImages(0, i) = "1"
                aImages(1, i) = "1"
            Next
        End If
        
        If Not ObjPic Is Nothing Then
            '���ͼ������
            strSQL = "Select B.���,B.����,A.Ԫ��ID,A.����ID,B.W,B.H From" & _
                " (Select B.ID As Ԫ��ID,A.ID ����ID,Rownum As Seq From ���˲������� A,����Ԫ��Ŀ¼ B,����ҽ������ C" & _
                " Where C.����ID=A.������¼ID AND A.Ԫ�ر���=B.���� And" & _
                " C.ҽ��ID=[1] And C.���ͺ�=[2] And A.Ԫ������=3) A," & _
                " (Select A.���,B.����,B.W,B.H,Rownum As Seq" & _
                " From zlReports A,zlRPTItems B,�����ļ�Ŀ¼ C,����ҽ����¼ D,���Ƶ���Ӧ�� E" & _
                " Where A.ID=B.����ID And A.���='ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2'" & _
                " And C.ID=E.�����ļ�ID And D.������ĿID=E.������ĿID And Nvl(B.����,0)=1 And B.����=11" & _
                " And E.Ӧ�ó���=D.������Դ And D.ID=[1]" & _
                " And B.���� Like '���%'" & _
                " Order BY Trunc(Y/567),Trunc(X/567)) B Where A.Seq=B.Seq"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            If blnMoved Then
                strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "��鱨��", lngҽ��ID, lng���ͺ�)
            iFlagCount = rsTmp.RecordCount
            ObjPic.Cls
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = strTempPath & objFileSystem.GetTempName
                
                '���������ߴ�
                On Error Resume Next
                Set ObjPic.Picture = ReadCaseMap(rsTmp(2))
                ObjPic.Width = ObjPic.ScaleX(ObjPic.Picture.Width, vbHimetric, vbTwips): ObjPic.Height = ObjPic.ScaleY(ObjPic.Picture.Height, vbHimetric, vbTwips)
                If ObjPic.Width / ObjPic.Height > rsTmp(4) / rsTmp(5) Then
                    ObjPic.Width = ObjPic.Height * rsTmp(4) / rsTmp(5)
                Else
                    ObjPic.Height = ObjPic.Width / (rsTmp(4) / rsTmp(5))
                End If
                ObjPic.Cls: Set ObjPic.Picture = Nothing
                On Error GoTo DBError
                Call ShowMapInOjbect_1(ObjPic, rsTmp(2), rsTmp(3), blnMoved:=blnMoved)
                SavePicture ObjPic.Image, strTmpFile
                ObjPic.Cls
            
                aFlagImages(0, i) = rsTmp(1)
                aFlagImages(1, i) = strTmpFile
                rsTmp.MoveNext
            Next
            For i = rsTmp.RecordCount To 8
                aFlagImages(0, i) = "1"
                aFlagImages(1, i) = "1"
            Next
            Call ReportOpen(gcnOracle, glngSys, strRptName, objParent, _
                "NO=" & strNO, _
                "����=" & lng��¼����, "ҽ��ID=" & lngҽ��ID, _
                aImages(0, 0) & "=" & aImages(1, 0), _
                aImages(0, 1) & "=" & aImages(1, 1), _
                aImages(0, 2) & "=" & aImages(1, 2), _
                aImages(0, 3) & "=" & aImages(1, 3), _
                aImages(0, 4) & "=" & aImages(1, 4), _
                aImages(0, 5) & "=" & aImages(1, 5), _
                aImages(0, 6) & "=" & aImages(1, 6), _
                aImages(0, 7) & "=" & aImages(1, 7), _
                aImages(0, 8) & "=" & aImages(1, 8), _
                aFlagImages(0, 0) & "=" & aFlagImages(1, 0), _
                aFlagImages(0, 1) & "=" & aFlagImages(1, 1), _
                aFlagImages(0, 2) & "=" & aFlagImages(1, 2), _
                aFlagImages(0, 3) & "=" & aFlagImages(1, 3), _
                aFlagImages(0, 4) & "=" & aFlagImages(1, 4), _
                aFlagImages(0, 5) & "=" & aFlagImages(1, 5), _
                aFlagImages(0, 6) & "=" & aFlagImages(1, 6), _
                aFlagImages(0, 7) & "=" & aFlagImages(1, 7), _
                aFlagImages(0, 8) & "=" & aFlagImages(1, 8), PrtMode)
        Else
            Call ReportOpen(gcnOracle, glngSys, strRptName, objParent, _
                "NO=" & strNO, _
                "����=" & lng��¼����, "ҽ��ID=" & lngҽ��ID, _
                aImages(0, 0) & "=" & aImages(1, 0), _
                aImages(0, 1) & "=" & aImages(1, 1), _
                aImages(0, 2) & "=" & aImages(1, 2), _
                aImages(0, 3) & "=" & aImages(1, 3), _
                aImages(0, 4) & "=" & aImages(1, 4), _
                aImages(0, 5) & "=" & aImages(1, 5), _
                aImages(0, 6) & "=" & aImages(1, 6), _
                aImages(0, 7) & "=" & aImages(1, 7), _
                aImages(0, 8) & "=" & aImages(1, 8), PrtMode)
        End If
        'ɾ����ʱ�ļ�
'        For i = 0 To iTmpFileCount - 1
'            objFileSystem.DeleteFile aImages(1, i), True
'        Next
        For i = 0 To iFlagCount - 1
            If Dir(aFlagImages(1, i), vbDirectory) <> "" Then
                objFileSystem.DeleteFile aFlagImages(1, i), True
            End If
        Next
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMapInOjbect_1(objDraw As Object, varԪ�� As Variant, Optional lng����ID As Long, Optional x As Long, Optional y As Long, Optional W As Long, Optional H As Long, Optional blnMoved As Boolean = False)
'���ܣ���ָ���Ķ���(PictureBox��Form)����ʾ���ͼ
'������objDraw=PictureBox�������,����ScaleMode����ΪPixel
'      varԪ��=���ͼԪ�صı���(�ַ���)��ID(������)
'      lng����ID="���˲�������"�б��ͼԪ�ض�Ӧ��ID,�������,����ʾ���ͼ����
'      X,Y,W,H=��ʾ��Ŀ��ͻ��˷�Χ,���Բ�ָ��,��λΪPixel
'˵�����������øú������д�ӡ���(��Ϊ��API��ͼ,��˲���ֱ�ӽ�objDrawָ��Ϊ��ӡ��,������PictureBox�ϰ�һ�����������,ȡPictureBox.Image�������ӡ��)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objCaseMap As StdPicture, objMapItems As New MapItems
    
    On Error GoTo errH
        
    '��ȡ���ͼԪ�ص�����
    If TypeName(varԪ��) = "String" Then
        strSQL = "Select * From ����Ԫ��Ŀ¼ Where ����=[1]"
    Else
        strSQL = "Select * From ����Ԫ��Ŀ¼ Where ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", CStr(varԪ��))
    If rsTmp.EOF Then Exit Sub '����Ҫ��ͼ�α���
    
    Set objCaseMap = ReadCaseMap(rsTmp!ID)
    If objCaseMap Is Nothing Then Exit Sub '����Ҫ��ͼ�α���
    
    '��ȡ���ͼ�ı�ע����
    If lng����ID <> 0 Then
        strSQL = "Select * From ���˲������ͼ Where ����ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "���˲������ͼ", "H���˲������ͼ")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", lng����ID)
        Do While Not rsTmp.EOF
            With rsTmp
                objMapItems.Add !����, zlCommFun.NVL(!����), _
                    IIf(IsNull(!����), IIf(!���� = 0, "����,9,0,0000", ""), !����), _
                    zlCommFun.NVL(!�㼯), zlCommFun.NVL(!X1, 0), _
                    zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                    zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!���ɫ, &HFFFFFF), _
                    zlCommFun.NVL(!��䷽ʽ, -1), zlCommFun.NVL(!����ɫ, 0), _
                    zlCommFun.NVL(!����, 0), zlCommFun.NVL(!�߿�, 1)
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    On Error GoTo 0
    
    Call ShowCaseMap(objCaseMap, objMapItems, objDraw, x, y, W, H)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CheckOneDuty(ByVal strҽ�� As String, ByVal strְ�� As String, ByVal strҽ�� As String, ByVal blnҽ�� As Boolean) As String
'���ܣ���鵱ǰָ��ҩƷ����ְ���Ƿ����
'������strҽ��=ҩƷҽ����ʾ����
'      strְ��=ҩƷ����ְ��
'      strҽ��=����ҽ��
'      blnҽ��=�Ƿ񹫷ѻ�ҽ������
'      grsDuty=��¼ҽ��ְ�񻺴�
'���أ�ְ���������ʾ��Ϣ����������򷵻ؿա�
    Const STR_ְ�� = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim intְ��A As Integer, intְ��B As Integer
    
    If Len(strְ��) <> 2 Or strҽ�� = "" Then Exit Function
    
    'ȡҩƷ����ְ��
    If blnҽ�� Then
        intְ��B = Val(Right(strְ��, 1))
    Else
        intְ��B = Val(Left(strְ��, 1))
    End If
    If intְ��B = 0 Then Exit Function '������
    
    'ȡҽ��ְ��
    If grsDuty Is Nothing Then
        Set grsDuty = New ADODB.Recordset
        grsDuty.Fields.Append "ҽ��", adVarChar, 50
        grsDuty.Fields.Append "ְ��", adInteger
        grsDuty.CursorLocation = adUseClient
        grsDuty.LockType = adLockOptimistic
        grsDuty.CursorType = adOpenStatic
        grsDuty.Open
    End If
    grsDuty.Filter = "ҽ��='" & strҽ�� & "'"
    If grsDuty.EOF Then
        On Error GoTo errH
        strSQL = "Select ����,Nvl(Ƹ�μ���ְ��,0) as ְ�� From ��Ա�� Where ����='" & strҽ�� & "'"
        Set rsTmp = New ADODB.Recordset
        Call OpenRecord(rsTmp, strSQL, "mdlCISWork")
        On Error GoTo 0
        If Not rsTmp.EOF Then
            grsDuty.AddNew
            grsDuty!ҽ�� = rsTmp!����
            grsDuty!ְ�� = rsTmp!ְ��
            grsDuty.Update
        End If
    End If
    If Not grsDuty.EOF Then
        intְ��A = grsDuty!ְ��
    End If
        
    '���ְ��Ҫ��
    If intְ��A = 0 Then
        'ҽ��δ����ְ������
        strMsg = """" & strҽ�� & """Ҫ��Ĵ���ְ�����㣺" & vbCrLf & vbCrLf & IIf(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """δ����ְ��"
    ElseIf intְ��B < intְ��A Then
        '��ֵԽСְ��Խ��
        strMsg = """" & strҽ�� & """Ҫ��Ĵ���ְ�����㣺" & vbCrLf & vbCrLf & IIf(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """��ְ��Ϊ""" & Split(STR_ְ��, ",")(intְ��A - 1) & """��"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub MkLocalDir(ByVal strDir As String)
'���ܣ���������Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function GetSysParVal(Optional ByVal int������ As Integer = -9999, Optional ByVal strDefault As String) As String
'���ܣ���ȡָ��ϵͳ������ֵ
'������int������=Ϊ-9999ʱ����ʼ��������
'      strDefault=���û��ֵ��Ϊ�յ�ȱʡֵ
    Dim blnDo As Boolean, strSQL As String
    
    On Error GoTo errH
    
    blnDo = True
    If int������ <> -9999 Then
        If Not grsSysPars Is Nothing Then
            If grsSysPars.State = 1 Then blnDo = False
        End If
    End If
    If blnDo Then
        strSQL = "Select ������,������,����ֵ From ϵͳ������"
        Set grsSysPars = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(grsSysPars, strSQL, "GetSysParVal")
    End If
    
    If int������ <> -9999 Then
        grsSysPars.Filter = "������=" & int������
        If Not grsSysPars.EOF Then
            GetSysParVal = NVL(grsSysPars!����ֵ, strDefault)
        Else
            GetSysParVal = strDefault
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function funAssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'���viewer�е���ʾ������ͼ���һ��ͼ��

    Dim Image As New DicomImage '��ͼ��
    Dim imgs As New DicomImages '��ʱ�洢��Ļ�ɼ���ͼ��
    Dim intWidth As Integer     '��ͼ��Ŀ��
    Dim intHeight As Integer    '��ͼ��ĸ߶�
    Dim Simg As New DicomImage
    Dim intLeft As Integer
    Dim intRight As Integer
    Dim intTop As Integer
    Dim intBottom As Integer
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '����ͼ���ռ�õ�������
    Dim intImgRectHeight As Integer '����ͼ���ռ�õ�����߶�
    Dim i As Integer
    Dim intMaxWidth As Integer      'ƴ�Ӻ�ͼ��������
    Dim intMaxHeight As Integer     'ƴ�Ӻ�ͼ������߶�
    Dim intBorder As Integer        'ͼ��֮��ı߾�
    Dim intImgX As Integer          'X�����ͼ������
    Dim intImgY As Integer          'Y�����ͼ������
    Dim intActualSizex As Integer   'ͼ����ת�任��X��������ص���
    Dim intActualSizey As Integer   'ͼ����ת�任��Y��������ص���
    Dim intOffsetX As Integer       'ƴ��ʱX�����λ��
    Dim intOffsetY As Integer       'ƴ��ʱY�����λ��
    Dim dlImgLabel As DicomLabel    'ͼ��ı�ע
    Dim lngWhiteX As Long           '��ͼ���ɫ�ĳɰ�ɫ��X���
    Dim lngWhiteY As Long           '��ͼ���ɫ�ĳɰ�ɫ��Y�߶�
    
    If AssembleViewer.Count <= 0 Then
        '����һ����ͼ**************
        Exit Function
    End If

    '������ͼ��Ŀ�Ⱥ͸߶�

    '��ͼ��Ŀ�Ⱥ͸߶Ȳ��ܹ�����intMaxWidth��intMaxHeight����ȡ��߶ȣ�
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '������ͼ��Ŀ�Ⱥ͸߶�

    'ʹ��ԭͼ��Ŀ�Ⱥ͸߶Ⱥͣ�����Viewer�ı�����������

    '����ͼ����¿��

    For i = 1 To AssembleViewer.Count
        sZoom = (lngWidth / intCols) / (AssembleViewer(i).SizeX * Screen.TwipsPerPixelX)
        If sZoom > (lngHeight / intRows) / (AssembleViewer(i).SizeY * Screen.TwipsPerPixelY) Then
            sZoom = (lngHeight / intRows) / (AssembleViewer(i).SizeY * Screen.TwipsPerPixelY)
        End If
        AssembleViewer(i).Zoom = sZoom
        '�ɼ�ͼ��
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '��ȷ������ͼ��Ŀ�Ⱥ͸߶�
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '������ͼ��
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT����MONOCHROME2,CR����MONOCHROME1��
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    'ƴ����ͼ��
    For i = 1 To imgs.Count
        '����ͼ����λ��
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set funAssembleImage = Image
End Function

Private Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1

    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub
