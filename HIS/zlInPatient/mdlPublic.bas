Attribute VB_Name = "mdlPublic"
Option Explicit 'Ҫ���������

'ϵͳ���ñ���

Public gfrmMain As Object                   '����̨����
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gMainPrivs As String                 '���������������е�Ȩ��,ע����ڲ�ģ��Ȩ��
Public gstrPrivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String            'OEM��Ʒ����
Public glngSys As Long
Public glngModul As Long

Public gstrSQL As String
Public gblnOK As Boolean

Public gblnLED As Boolean       '��Ԥ����ʱ�Ƿ�ʹ��LED��������
Public gblnLedWelcome As Boolean '�Ƿ���Ԥ�����겡�˺���ʾ��ӭ��Ϣ

Public gobjSquare As SquareCard  '�����㲿��

'�û���Ϣ------------------------
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

'---------------------------------------
'����27554 by lesfeng 2010-01-19
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const HTCAPTION = 2
Public Const GWL_WNDPROC = -4
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const SRCCOPY = &HCC0020
Public Const SM_CYCAPTION = 4
Public Const CB_GETDROPPEDSTATE = &H157

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function MoveObj(lngHwnd As Long) As RECT
'���ܣ��ڶ����MouseDown�¼��е���,����������Hwnd����
'���أ������Ļ������ֵ
   
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Function GetColNum(lvwTemp As ListView, strHead As String) As Integer
    Dim i As Integer
    For i = 1 To lvwTemp.ColumnHeaders.Count
        If lvwTemp.ColumnHeaders(i).Text = strHead Then GetColNum = i: Exit Function
    Next
End Function

Public Sub SetCenter(frm As Form)
'���ܣ������嶨λ����Ļ����
    frm.Left = (Screen.width - frm.width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
End Sub

Public Function CheckLen(txt As TextBox, intLen As Integer) As Boolean
'���ܣ���鹤�������ʵ�����Ƿ���ָ�����Ƴ�����
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox Mid(txt.Name, 4) & "ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�", vbInformation, gstrSysName
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function

Public Function CaptionHeight() As Long
'����:����ϵͳ����������߶�(������Ϊ��λ)
    CaptionHeight = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
End Function

Public Sub SetItemInfo(lvw As Object, pan As Object)
'���ܣ�����Listview��ǰѡ���У���ʾ��״̬����
    Dim i As Integer, strInfo As String
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If lvw.SelectedItem.Text <> "" Then
        strInfo = "/" & lvw.ColumnHeaders(1).Text & ":" & lvw.SelectedItem.Text
    End If
    
    For i = 2 To lvw.ColumnHeaders.Count
        If lvw.SelectedItem.SubItems(i - 1) <> "" Then
            strInfo = strInfo & "/" & lvw.ColumnHeaders(i).Text & ":" & lvw.SelectedItem.SubItems(i - 1)
        End If
    Next
    If strInfo <> "" Then pan.Text = Mid(strInfo, 2)
End Sub

Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Sub SetGridWidth(msh As Control, frmParent As Object)
'���ܣ��Զ���������п�,����С�ʺ�Ϊ׼
    Dim blnRedraw As Boolean
    Dim blnDo As Boolean, i As Long, j As Long
    Dim lngStart As Long, lngEnd As Long, lngMaxWidth As Long
        
    blnRedraw = msh.Redraw
    msh.Redraw = False
    lngStart = IIf(msh.FixedRows = 0, 0, msh.FixedRows - 1)
    lngEnd = msh.Rows - 1
    
    For i = 0 To msh.Cols - 1
        lngMaxWidth = 0
        For j = lngStart To lngEnd
            blnDo = True
            If msh.MergeRow(j) Then
                If i > 0 Then If msh.TextMatrix(j, i) = msh.TextMatrix(j, i - 1) Then blnDo = False
                If i < msh.Cols - 1 Then If msh.TextMatrix(j, i) = msh.TextMatrix(j, i + 1) Then blnDo = False
            End If
            If blnDo Then
                If Len(msh.TextMatrix(j, i)) > Len(msh.TextMatrix(lngMaxWidth, i)) Then
                    lngMaxWidth = j
                End If
            End If
        Next
        msh.ColWidth(i) = IIf(frmParent.TextWidth(msh.TextMatrix(lngMaxWidth, i)) > 3000, 3000, frmParent.TextWidth(msh.TextMatrix(lngMaxWidth, i)) + 90)
    Next
    
    msh.Redraw = blnRedraw
End Sub

Public Function CheckFormInput(objForm As Object, Optional ByVal strIgnore As String, Optional ByVal strToNumText As String = "") As Boolean
    '����:strIgnore-�����Ŀؼ���,�����ж��,����,�ŵȷָ�
    '����:strToNumText--��Ҫ���н�ǧ��λ��ʽ�Ľ��ת����������ʽ���ı��ؼ�����,�����ж��,����,�ŵȷָ�
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled And Not obj.Locked Then
                strText = ""
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                    If InStr(1, "," & UCase(strToNumText) & ",", "," & UCase(obj.Name) & ",") > 0 Then
                        strText = StrToNum(strText)
                    End If
                Case "ComboBox"
                    If obj.Style = 0 Then strText = obj.Text
                End Select
                If InStr(UCase(strIgnore), UCase(obj.Name)) = 0 Then
                    If InStr(strText, "'") > 0 _
                        Or InStr(strText, "|") > 0 _
                        Or InStr(strText, "~") > 0 _
                        Or InStr(strText, "^") > 0 Then
                        MsgBox "���������а����Ƿ��ַ���", vbInformation, gstrSysName
                        obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                        obj.SetFocus: Exit Function
                    End If
                End If
            End If
        End If
    Next
    CheckFormInput = True
End Function

Public Function GetIDDate(ID As String) As String
'���ܣ��������֤�ŷ��س�������,��ʽ"yyyy-MM-dd"
'������ID=���֤��,Ӧ��Ϊ15λ��18λ
    Dim strTmp As String
    
    If Len(ID) = 15 Then
        strTmp = Mid(ID, 7, 6)
        If Len(strTmp) = 6 And IsNumeric(strTmp) Then
            strTmp = "19" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2)
        End If
    ElseIf Len(ID) = 18 Then
        strTmp = Mid(ID, 7, 8)
        If Len(strTmp) = 8 And IsNumeric(strTmp) Then
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2)
        End If
    End If
    If IsDate(strTmp) Then GetIDDate = strTmp
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

Public Sub CboLoadData(ByRef cbo As ComboBox, ByRef rsTmp As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
    '����:װ��������ָ�������������������е���������
    '����:cbo   Ҫװ�ؼ�¼����������ؼ�
    '     rsTmp     ��¼������,Ҫ������������������,Id,���룬����
    '     blnClear    װ��ʱ�Ƿ����ԭ�е���������,ȱʡΪTrue
    
    If rsTmp.Fields.Count < 3 Then Exit Sub
    If blnClear = True Then cbo.Clear
    
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        While Not rsTmp.EOF
            cbo.AddItem rsTmp.Fields(1).Value & "-" & rsTmp.Fields(2).Value
            cbo.ItemData(cbo.NewIndex) = Val(rsTmp.Fields(0).Value)
            rsTmp.MoveNext
        Wend
        rsTmp.MoveFirst
    End If
End Sub

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



