Attribute VB_Name = "mdlPublic"
Option Explicit 'Ҫ���������
'ϵͳ������ʱ����
Public gfrmMain As Object
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrSQL As String

Public glngSys As Long
Public glngModul As Long
Public gblnShowInTaskBar As Boolean

Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public gstrProductName As String

Public gstrUnitName As String '�û���λ����
Public gstrDBUser As String '��ǰ���ݿ��û���

Public gstrIme As String '�Զ��Ŀ������뷨
Public gblnOK As Boolean
Public Const LONG_MAX = 2147483647 'Long�����ֵ
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'�����������------------------------------------
Public Const conLineWide As Integer = 30 '������ռ���(��λΪ�)ռ�����߿��
Public Const conLineHigh As Integer = 30 '������ռ�߶�(��λΪ�)ռ�����߸߶�
Public gobjOutTo As Object

Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    pסԺ���� = 1133
    p���˽��� = 1137
    p���ò�ѯ = 1139
    pһ���嵥 = 1141
    p���ʲ��� = 1150
    pסԺҽ���´� = 1253
    pԤ���� = 1103
End Enum

'API��Ϣ-----------------------------------------
Public lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    �������� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Const WM_VSCROLL = &H115
Public Const SB_TOP = 6

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'Windows���----------------------------------
Public Const ETO_OPAQUE = 2
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'���뷨����API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
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
 
'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'��ͼAPI
Public Const PS_SOLID = 0
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Public Enum Em_Appearance
    Show_3D = 1     '3D��ʾ
    Show_Flat = 0   'ƽ��
End Enum
Public Enum Em_BorderStyle
    Show_Fixed_Single = 1
    Show_None = 0   '�ޱ߿���
End Enum

Public Enum gAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum
Public Enum EM_DrawStyle
    DW_Flat = 0  '= ƽ��
    Dw_SubKen = -1 '= ����
    Dw_Heave = 1  '= ͹��
    Dw_Deepen_Subken = -2 '= ���,
    Dw_Deepen_Heave = 2 ' = ��͹��
End Enum
Private Type TY_System_para_Balance
    blnˢ���������� As Boolean  '�Ƿ�ˢ����������
    bln��Ժ��׼���� As Boolean '1-��Ժ��׼����,0-��Ժ�������
    bytAuditing As Byte  '����δ��˵��ݵĽ��ʴ���:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    byt���δִ�� As Byte    '��Ժ�ͽ��ʳ�Ժʱ����Ƿ���δִ����Ŀ��δ��ҩƷ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    byt���δ��ҩ As Byte   '�ڳ�Ժ���ʼ�������������г�Ժʱ�Ƿ��鲡�˵�δ��ҩƷ��Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    blnҽ��������ܳ�Ժ As Boolean 'ҽ���´��Ժҽ���������˳�Ժ

End Type

Public Enum Em_InputMode
    InPut_Chars = 0 '�ַ�
    InPut_Numbers = 1   '����
    Input_Moneys = 2    '���(����������)
    Input_NegativeMoneys = 4 '���������
End Enum


'�ؼ���λ
Public Type ty_ctlObject_Locale
    '�ؼ���λ��
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    '�����б����С�߶ȺͿ��
    minWidth As Single
    minHeight As Single
    
    '�½��б��ʵ��λ��
    DownLeft As Single
    DownTop As Single
    DownWidth As Single
    DownHeight As Single
 
    
    '��ģ���
    ScreenWidth As Single
    ScreenHeight As Single
    
End Type

Public Function OpenIme(Optional strIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If strIme = "���Զ�����" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

Public Function SetWidth(cboHwnd As Long, NewWidthPixel As Long) As Boolean
'���ܣ����� Combo �����Ŀ��,��λΪ pixels
    '��twipsΪ��λ����
    Call cbo.SetListWidth(cboHwnd, NewWidthPixel * Screen.TwipsPerPixelX)
    
    SetWidth = True
End Function

Public Function GetWidth(cboHwnd As Long) As Long
'���ܣ� ȡ�� Combo �����Ŀ��,��λΪ pixels
    Dim lRetVal As Long
    lRetVal = cbo.ListWidth(cboHwnd)
    If lRetVal <> -1 Then
        GetWidth = lRetVal / Screen.TwipsPerPixelX
    Else
        GetWidth = 0
    End If
End Function

Public Function GetBillTotal(objBill As ExpenseBill) As Currency
'���ܣ���ȡ���ݷ�Ŀ�ϼƽ��
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    
    For Each objBillDetail In objBill.Details
        For Each objBillIncome In objBillDetail.InComes
            GetBillTotal = GetBillTotal + objBillIncome.ʵ�ս��
        Next
    Next
End Function

Public Function GetBillRowTotal(objBillInComes As BillInComes) As Currency
'���ܣ���ȡ���ݷ�Ŀ�ϼƽ��
    Dim objBillIncome As New BillInCome
    For Each objBillIncome In objBillInComes
        GetBillRowTotal = GetBillRowTotal + objBillIncome.ʵ�ս��
    Next
End Function

Public Function GetFirstRow(curBill As ExpenseBill, Optional strClass As String) As Integer
'���ܣ���ȡ��ǰ�����е�һ��ΪҩƷ���շ��к�
'������strClass=ȡ��һ��ҩ����ҩ��,��ΪҩƷ
'���أ�0=û��ҩƷ�շ���
    Dim i As Long
    If curBill.Details.Count = 0 Then GetFirstRow = 0
    For i = 1 To curBill.Details.Count
        If strClass = "" Then
            If InStr(",5,6,7,", curBill.Details(i).�շ����) > 0 Then
                GetFirstRow = i: Exit Function
            End If
        Else
            If curBill.Details(i).�շ���� = strClass Then
                GetFirstRow = i: Exit Function
            End If
        End If
    Next
End Function

Public Function GetFirstClass(curBill As ExpenseBill) As String
'���ܣ���ȡ��ǰ�����е�һ��ΪҩƷ���շ��к�
'���أ�0=û��ҩƷ�շ���
    Dim i As Long
    If curBill.Details.Count = 0 Then GetFirstClass = ""
    For i = 1 To curBill.Details.Count
        If InStr(",5,6,7,", curBill.Details(i).�շ����) > 0 Then
            GetFirstClass = curBill.Details(i).�շ����: Exit Function
        End If
    Next
End Function

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If msg = WM_GETMINMAXINFO Then
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
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, msg, wp, lp)
End Function

Public Function strPad(ByVal strPre As String, ByVal intLen As Integer, ByVal strFill As String, ByVal bytAlign As Byte, Optional ByVal blnTrim As Boolean) As String
'���ܣ�����ַ���
'������
'     strPre=Ҫ�����ַ���
'     intLen=����ĳ���
'     strFill=Ҫ�����ַ�
'     bytAlign=1,2/��,�Ҷ��룬�����ʱ����ԭ�ַ����ұ����
'     blnTrim=���ַ�������ʱ���Ƿ�ǿ�а�ָ�����Ƚ�ȡ��
'���أ��������ַ���
'˵����һ�����ֵ��������ַ����ȴ���
    Dim i As Long
    
    If LenB(StrConv(strPre, vbFromUnicode)) >= intLen Then
        If blnTrim Then
            For i = 1 To Len(strPre)
                strPad = strPad & Mid(strPre, i, 1)
                If LenB(StrConv(strPad, vbFromUnicode)) >= intLen Then Exit For
            Next
        Else
            strPad = strPre
        End If
    Else
        If Len(strFill) > 1 Then strFill = Left(strFill, 1)
        If bytAlign = 1 Then
            strPad = strPre
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
        ElseIf bytAlign = 2 Then
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
            strPad = strPad & strPre
        End If
    End If
End Function

Public Sub PrintCell(ByVal Text As String, _
    ByVal X As Single, ByVal Y As Single, _
    Optional ByVal Wide, _
    Optional ByVal High, _
    Optional Alignment As Byte = 0, _
    Optional ForeColor As Long = 0, _
    Optional GridColor As Long = 0, _
    Optional FillColor As Long = 0, _
    Optional LineStyle As String = "1111", _
    Optional FontName, Optional FontSize, _
    Optional FontBold, Optional FontItalic)
    '------------------------------------------------
    '���ܣ� ��ָ�������ӡһ�����ݵ�Ԫ,������ǰ�����ƶ�����Ԫ���Ͻ�λ��
    '������
    '   Text:    ������ַ���,���в������س����з�
    '   X:       ���Ͻ�X����
    '   Y:       ���Ͻ�Y����
    '   Wide:    ������
    '   High:    ����߶�
    '   Alignment:    ����ģʽ��0-�����(ȱʡ),1-�Ҷ���,2-����
    '   ForeColorǰ��ɫ,ȱʡΪ��ɫ
    '   GridColor����ɫ,ȱʡΪ��ɫ
    '   FillColor���ɫ,ȱʡΪ�豸����ɫ,����ϵͳ�����˺�ɫ��ɫ�룬���Խ�����������ɫ
    '   LineStyle:����ֱ�Ϊ�������µ��������
    '           0-���ߣ�1-9����Ӵ֣�1Ϊȱʡ
    '   FontName,FontSize,FontBold,FontItalic:��������
    '���أ�
    '------------------------------------------------
    Dim aryString() As String       '�س��ָ���ַ���
    Dim lngOldForeColor As Long     '����豸ȱʡǰ��ɫ
    Dim intRow As Integer, intAllRow As Integer
    Dim strRest As String, sngYMove As Single
    Dim oldFontName, oldFontSize, oldFontBold, oldFontItalic
    lngOldForeColor = gobjOutTo.ForeColor
    
    On Error Resume Next
    With gobjOutTo
        If Not IsMissing(FontName) Then
            oldFontName = gobjOutTo.FontName
            .FontName = FontName
        End If
        If Not IsMissing(FontSize) Then
            .FontSize = FontSize
            oldFontSize = gobjOutTo.FontSize
        End If
        If Not IsMissing(FontBold) Then
            .FontBold = FontBold
            oldFontBold = gobjOutTo.FontBold
        End If
        If Not IsMissing(FontItalic) Then
            .FontItalic = FontItalic
            oldFontItalic = gobjOutTo.FontItalic
        End If
    End With
    
    If IsMissing(Wide) Then Wide = gobjOutTo.TextWidth(Text) + 2 * conLineWide
    If IsMissing(High) Then High = gobjOutTo.TextHeight(Text) + 2 * conLineHigh
'    Wide = CLng(Wide)
'    High = CLng(High)
    If Wide * High = 0 Then Exit Sub
    
    If UCase(TypeName(LineStyle)) <> "STRING" Then LineStyle = CStr(LineStyle)
    If Len(LineStyle) < 4 Then
        LineStyle = Left(LineStyle & "1111", 4)
    End If
    
    '------------------------------------------
    '   ���ߴ�ӡ
    '------------------------------------------
    If Mid(LineStyle, 1, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 1, 1)
        gobjOutTo.Line (X, Y)-(X + Wide, Y), GridColor
    End If
    
    If Mid(LineStyle, 2, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 2, 1)
        gobjOutTo.Line (X, Y)-(X, Y + High), GridColor
    End If
    
    If Mid(LineStyle, 3, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 3, 1)
        gobjOutTo.Line (X + Wide, Y)-(X + Wide, Y + High), GridColor
    End If
    
    If Mid(LineStyle, 4, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 4, 1)
        gobjOutTo.Line (X, Y + High)-(X + Wide, Y + High), GridColor
    End If
    
    If Wide > conLineWide And High > conLineHigh Then
        '------------------------------------------
        '   ��ɫ���
        '------------------------------------------
'        If FillColor <> 0 Then
'            Printer.FillStyle = 1
'            gobjOutTo.Line (X + conLineWide / 2, Y + conLineHigh / 2)- _
'                (X + Wide - conLineWide / 2, Y + High - conLineHigh / 2), _
'                FillColor, BF
'        End If
        
        '------------------------------------------
        '   ���ִ�ӡ
        '------------------------------------------
        gobjOutTo.ForeColor = ForeColor
    
        If InStr(Text, vbCrLf) = 0 And InStr(Text, Chr(13)) = 0 Then
            If Wide - conLineWide < gobjOutTo.TextWidth("1") Then    'С��һ���ַ�
                intAllRow = 1
            Else
                If gobjOutTo.TextWidth(Text) Mod (Wide - conLineWide) = 0 Then
                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide)
                Else
                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide) + 1
                End If
            End If
            For intRow = intAllRow To 1 Step -1
                If High >= gobjOutTo.TextHeight(Text) * intRow Then
                    Exit For
                End If
            Next
            intAllRow = intRow
            sngYMove = (High - conLineHigh - gobjOutTo.TextHeight(Text) * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow - 1
                Do While gobjOutTo.TextWidth(Text) > Wide - conLineWide
                    If Len(Trim(Text)) <= 1 Then Exit Do
                    Text = Left(Text, Len(Text) - 1)
                Loop
                strRest = Mid(strRest, Len(Text) + 1)
                Select Case Alignment
                Case 2
                    gobjOutTo.CurrentX = X + (Wide - gobjOutTo.TextWidth(Text)) / 2
                Case 1
                    gobjOutTo.CurrentX = X - conLineWide / 2 + Wide - gobjOutTo.TextWidth(Text)
                Case Else
                    gobjOutTo.CurrentX = X + conLineWide / 2
                End Select
                gobjOutTo.CurrentY = Y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(Text)
                gobjOutTo.Print Text
                Text = strRest
            Next
        Else
            If InStr(Text, vbCrLf) > 0 Then
                aryString = Split(Trim(Text), vbCrLf)
            Else
                aryString = Split(Trim(Text), Chr(13))
            End If

            intAllRow = UBound(aryString)
            sngYMove = (High - conLineHigh - gobjOutTo.TextHeight("ZYL") * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow
                strRest = aryString(intRow)
                Select Case Alignment
                Case 2
                    Dim blnLR As Boolean
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        blnLR = Not blnLR
                        strRest = IIf(blnLR, Left(strRest, Len(strRest) - 1), Right(strRest, Len(strRest) - 1))
                    Loop
                    gobjOutTo.CurrentX = X + (Wide - gobjOutTo.TextWidth(strRest)) / 2
                Case 1
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        strRest = Right(strRest, Len(strRest) - 1)
                    Loop
                    gobjOutTo.CurrentX = X - conLineWide / 2 + Wide - gobjOutTo.TextWidth(strRest)
                Case Else
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        strRest = Left(strRest, Len(strRest) - 1)
                    Loop
                    gobjOutTo.CurrentX = X + conLineWide / 2
                End Select
                
                gobjOutTo.CurrentY = Y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(strRest)
                If gobjOutTo.CurrentY + gobjOutTo.TextHeight(strRest) > Y + High Then Exit For
                If gobjOutTo.CurrentY >= Y Then gobjOutTo.Print strRest
            
            Next
        End If
    End If
    gobjOutTo.CurrentX = X + Wide
    gobjOutTo.CurrentY = Y
    gobjOutTo.DrawStyle = 0
    gobjOutTo.DrawWidth = 1
    gobjOutTo.ForeColor = lngOldForeColor

    If Not IsMissing(FontName) Then gobjOutTo.FontName = oldFontName
    If Not IsMissing(FontSize) Then gobjOutTo.FontSize = oldFontSize
    If Not IsMissing(FontBold) Then gobjOutTo.FontBold = oldFontBold
    If Not IsMissing(FontItalic) Then gobjOutTo.FontItalic = oldFontItalic
End Sub

Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'���ܣ��ж��������Ƿ���ԭ�ۺ��ִ��޶��ķ�Χ��
'������varL=ԭ��,varR=�ּ�,varI=������
'���أ�������ڷ�Χ��,��Ϊ��ʾ��Ϣ,����Ϊ�մ�
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '�����ֵ������ͬ,���þ���ֵ�ж�
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "����ļ۸����ֵ���ڷ�Χ(" & FormatEx(Abs(varL), gintFeePrecision) & "-" & FormatEx(Abs(varR), gintFeePrecision) & ")��."
        End If
    Else
        '������Ų���ͬ,����ԭʼ��Χ�ж�
        If varI < varL Or varI > varR Then
            CheckScope = "����ļ۸�ֵ���ڷ�Χ(" & FormatEx(varL, gintFeePrecision) & "-" & FormatEx(varR, gintFeePrecision) & ")��."
        End If
    End If
End Function

Public Sub SetGridWidth(msh As Control, frmParent As Object)
'���ܣ��Զ���������п�,����С�ʺ�Ϊ׼
    Dim blnRedraw As Boolean
    Dim blnDo As Boolean, i As Long, j As Long, strText As String
    Dim lngStart As Long, lngEnd As Long, lngMaxLen As Long, lngCurLen As Long, lngMWRow As Long
        
    On Local Error Resume Next
    
    blnRedraw = msh.Redraw
    msh.Redraw = False
    lngStart = IIf(msh.FixedRows = 0, 0, msh.FixedRows - 1)
    lngEnd = msh.Rows - 1
    
    For i = 0 To msh.Cols - 1
        lngMaxLen = LenB(StrConv(msh.TextMatrix(0, i), vbFromUnicode))  '����Ϊ�������
        lngMWRow = 0
        For j = lngStart To lngEnd
            blnDo = True
            strText = msh.TextMatrix(j, i)
            
            If msh.MergeRow(j) Then
                If i > 0 Then If strText = msh.TextMatrix(j, i - 1) Then blnDo = False
                If blnDo Then
                    If i < msh.Cols - 1 Then If strText = msh.TextMatrix(j, i + 1) Then blnDo = False
                End If
            End If
            If blnDo Then
                lngCurLen = LenB(StrConv(strText, vbFromUnicode))
                If lngCurLen > lngMaxLen Then
                    lngMaxLen = lngCurLen
                    lngMWRow = j
                End If
            End If
        Next
        msh.ColWidth(i) = frmParent.TextWidth(msh.TextMatrix(lngMWRow, i)) + 100
        If msh.ColWidth(i) > 3090 Then msh.ColWidth(i) = 3000
    Next
    
    msh.Redraw = blnRedraw
End Sub

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hWnd, msg, wp, lp)
End Function

Public Sub SaveRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

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

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'���ܣ���ָ�����ֱҴ��������д���,���ش����Ľ��
'������curMoney=Ҫ���зֱҴ���Ľ��(ΪӦ�ɽ��,2λС��)
'      gBytMoney=
'         0.������
'         1.��ȡ�������뷨,eg:0.51=0.50;0.56=0.60
'         2.�����շ�,eg:0.51=0.60,0.56=0.60
'         3.����շ�,eg:0.51=0.50,0.56=0.50
'         4.�����������˫,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           �����������˫,����ҹ���ѧ����ίԱ����ʽ�䲼�ġ�������Լ����,������vb��Round����,�������������ְ�����λ����ʱ�����Ը����ֽ���������Լ
'           �����м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ
'         5.�������塢�������,�Խǽ��д�������Ҫ�ȶԷֱҽ�������,��0.29(��)���¶�����ǣ�0.80(��)���϶����ǣ�0.3-0.79����Ϊ0.5��
'         6.��������:eg:0.15=0.10:0.16=0.2:   ���˺� ����:34519  ����:2010-12-06 09:58:02
'91385,������5.�������塢������롱�����ȶԷֱҽ����������룬��0.24(��)���¶�����ǣ�0.75(��)���϶����ǣ�0.25-0.74������Ϊ0.5
'       �ֱ����������룬��ô0.00��0.24=0��0.25��0.5=0.50, 0.50��0.74=0.50��0.75��1.00=1������������ռ50%�ı���

    Dim intSign As Integer, curTmp As Currency

    If gBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf gBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '��ȡ��λ���,�ٴ���ֱ�,��:0.248 ��0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf gBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf gBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf gBytMoney = 4 Then
        CentMoney = Format(FormatEx(curMoney, 1), "0.00")
    ElseIf gBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf gBytMoney = 6 Then
         '���˺� ����:34519 ��������:eg:0.15=0.10:0.16=0.2:    ����:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function

Public Sub ExChangeLocate(objA As Object, objB As Object)
'���ܣ�����ҽ���Ϳ������ҵ�����λ��
    Dim X1 As Long, Y1 As Long, w1 As Long, t1 As Integer
    Dim X2 As Long, Y2 As Long, w2 As Long, t2 As Integer
    Dim obj1 As Object, obj2 As Object
    
    X1 = objA.Left
    Y1 = objA.Top
    w1 = objA.Width
    t1 = objA.TabIndex
    Set obj1 = objA.Container

    X2 = objB.Left
    Y2 = objB.Top
    w2 = objB.Width
    t2 = objB.TabIndex
    Set obj2 = objB.Container
    
    Set objB.Container = obj1
    If TypeName(objB) = "Label" Then
        objB.Left = X1 + w1 - objB.Width
    Else
        objB.Left = X1
        objB.Width = w1
    End If
    objB.Top = Y1
    objB.TabIndex = t1
    
    Set objA.Container = obj2
    If TypeName(objA) = "Label" Then
        objA.Left = X2 + w2 - objA.Width
    Else
        objA.Left = X2
        objA.Width = w2
    End If
    objA.Top = Y2
    objA.TabIndex = t2
End Sub

Public Function MoneyOverFlow(objBill As ExpenseBill) As Boolean
'���ܣ���鵥�ݺϼƽ���Ƿ����
'˵������Currency����922337203685477Ϊ׼
    Dim dblӦ�� As Double, dblʵ�� As Double
    Dim i As Integer, j As Integer
    
    'Ҫ��VALתΪDouble��������
    For i = 1 To objBill.Details.Count
        For j = 1 To objBill.Details(i).InComes.Count
            If Abs(dblӦ�� + Val(objBill.Details(i).InComes(j).Ӧ�ս��)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            If Abs(dblʵ�� + Val(objBill.Details(i).InComes(j).ʵ�ս��)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            dblӦ�� = dblӦ�� + Val(objBill.Details(i).InComes(j).Ӧ�ս��)
            dblʵ�� = dblʵ�� + Val(objBill.Details(i).InComes(j).ʵ�ս��)
        Next
    Next
End Function

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�������߶�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    GetTaskbarHeight = OS.TaskbarHeight
End Function


Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln������� As Boolean = True, Optional bln���� As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     bln�������     �Ƿ���и������
    '     bln����         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    zlDblIsValid = zlCommFun.DblIsValid(strInput, intMax, bln�������, bln����, hWnd, str��Ŀ)
End Function

Public Function Where����ʱ��(Optional strAlias As String) As String
    If strAlias = "" Then
        Where����ʱ�� = " (����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null) "
    Else
        Where����ʱ�� = " (" & strAlias & ".����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".����ʱ�� is null) "
    End If
End Function

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub


Public Sub zlRaisEffect(picBox As Object, Optional intStyle As EM_DrawStyle, _
    Optional strName As String = "", Optional TxtAlignment As gAlignment = 1)
    '���ܣ���PictureBoxģ���3Dƽ�水ť
    'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
    Dim PicRect As RECT
    Dim lngTmp As Long
    If picBox Is Nothing Then Exit Sub
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            If intStyle = 2 Then
                    DrawEdge .hDC, PicRect, EDGE_RAISED Or BF_SOFT, BF_RECT
            ElseIf intStyle = -2 Then
                    DrawEdge .hDC, PicRect, EDGE_SUNKEN Or BF_SOFT, BF_RECT
            Else
                DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
            End If
        End If
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If TxtAlignment = mCenterAgnmt Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            ElseIf TxtAlignment = mLeftAgnmt Then
                .CurrentX = .ScaleLeft
            Else
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) '-10
            End If
            picBox.Print strName
        End If
        .ScaleMode = lngTmp
        .Refresh
    End With
End Sub
 
'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function NotRightMenuMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then NotRightMenuMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function


