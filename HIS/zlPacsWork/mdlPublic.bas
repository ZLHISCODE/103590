Attribute VB_Name = "mdlPublic"
Option Explicit
Public lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long�����ֵ

Public Const ETO_OPAQUE = 2
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_GETMINMAXINFO = &H24
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'''''''''''''''''''''''''''''''''''''''''''''����ȫ���ȼ�'''''''''''''''''''''''''''''''''''''''''''
Public Declare Function RegisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long) As Long
Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4

Public preWinProc As Long
Public Modifiers As Long, uVirtKey As Long, idHotKey As Long

Private Type taLong
    ll As Long
End Type

Private Type t2Int
    lWord As Integer
    hWord As Integer
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'��������ComboBox����
Public Const CB_ADDSTRING = &H143
Public Const CB_SETITEMDATA = &H151
Public Const CB_SETCURSEL = &H14E

Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetComboData Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindComboStr Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

'DICOMͼ�����
Public Const ATTR_������� As String = "8:20"
Public Const ATTR_���ʱ�� As String = "8:30"
Public Const ATTR_Ӱ����� As String = "8:60"
Public Const ATTR_����豸 As String = "8:1090"

Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
    Dim vRect As RECT, vPos As POINTAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.x >= vRect.Left And vPos.x <= vRect.Right _
        And vPos.y >= vRect.Top And vPos.y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.Hwnd, vRect)
    lngStyle = GetWindowLong(objForm.Hwnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.Hwnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.Hwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
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
    Dim vRect As RECT, vDot1 As POINTAPI, vDot2 As POINTAPI
    
    Call GetWindowRect(ToolBar.Hwnd, vRect)
    vDot1.x = vRect.Left: vDot1.y = vRect.Top
    vDot2.x = vRect.Right: vDot2.y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.Hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.Hwnd, vDot2)
    
    vDot1.x = vDot1.x * 15: vDot1.y = vDot1.y * 15
    vDot2.x = vDot2.x * 15: vDot2.y = vDot2.y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.x + Button.Left, vDot2.y
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

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal LngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = LngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal Hwnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, Hwnd, Msg, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
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
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, 500, varValue)
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, 500, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
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

Public Function OpenRecord(rsTmp As ADODB.Recordset, strSQL As String, ByVal strTitle As String, _
    Optional CursorType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, strTitle, strSQL)
    rsTmp.Open strSQL, gcnOracle, CursorType, LockType
    Call SQLTest
    
    Set OpenRecord = rsTmp
End Function

Public Sub ExecuteProc(ByVal strSQL As String, ByVal strCaption As String)
'���ܣ�ִ��SQL���
'    Call SQLTest(App.ProductName, strCaption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    Call zl9comlib.SQLTest(App.ProductName, strCaption, strSQL)
    Call zl9comlib.zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Call zl9comlib.SQLTest
End Sub

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
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

Public Function GetFullDate(ByVal strText As String) As String
'���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd HH:mm)
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '���봮�а������ڷָ���
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                'ֻ���������ڲ���
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                'ֻ������ʱ�䲿��
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '����Ƿ�����,����ԭ����
            strTmp = strText
        End If
    Else
        '���������ڷָ���
        If Len(strTmp) <= 2 Then
            '��������dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '��������MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '��������yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '��������MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '��������yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '��������yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    GetFullDate = strTmp
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Public Function StringMask(ByVal strText As String, ByVal strMask As String) As Boolean
'���ܣ�����ַ����Ƿ�ֻ����ָ�����ַ�
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strMask, Mid(strText, i, 1)) = 0 Then Exit Function
    Next
    StringMask = True
End Function

Public Function ExeTimeValid(ByVal strTime As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As Boolean
'���ܣ����ָ����ִ��ʱ���Ƿ�Ϸ�
    Dim arrTime() As String, strTmp As String, i As Integer
    Dim strPreTime As String, intPreDay As Long, intCurDay As Long
    
    If strTime = "" Then Exit Function
    
    If str�����λ = "��" Then
        '1/8:00-3/15:00-5/9:00��1/8:00-3/15-5/9:00
        If Not StringMask(strTime, "0123456789:-/") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
        
        For i = 0 To UBound(arrTime)
            If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
            '���ڲ���
            strTmp = Split(arrTime(i), "/")(0)
            If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
            intCurDay = Val(strTmp)
            If intCurDay < 1 Or intCurDay > 7 Then Exit Function
            If intPreDay <> 0 Then
                If intCurDay < intPreDay Then Exit Function
            End If
            
            '����ʱ�䲿��
            strTmp = Split(arrTime(i), "/")(1)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Then Exit Function
            If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            
            strPreTime = Format(strTmp, "HH:mm")
            intPreDay = intCurDay
        Next
    ElseIf str�����λ = "��" Then
        If intƵ�ʼ�� = 1 Then
            '8:00-12:00-14:00��8:00-12-14:00
            If Not StringMask(strTime, "0123456789:-") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
            
            For i = 0 To UBound(arrTime)
                strTmp = arrTime(i)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Then Exit Function
                If strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                strPreTime = Format(strTmp, "HH:mm")
            Next
        Else
            '1/8:00-1/15:00-2/9:00��1/8:00-1/15-2/9:00
            If Not StringMask(strTime, "0123456789:-/") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
            
            For i = 0 To UBound(arrTime)
                If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
                '�����������
                strTmp = Split(arrTime(i), "/")(0)
                If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
                intCurDay = Val(strTmp)
                If intCurDay < 1 Or intCurDay > intƵ�ʼ�� Then Exit Function
                If intPreDay <> 0 Then
                    If intCurDay < intPreDay Then Exit Function
                End If
                
                '����ʱ�䲿��
                strTmp = Split(arrTime(i), "/")(1)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Then Exit Function
                If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                
                strPreTime = Format(strTmp, "HH:mm")
                intPreDay = intCurDay
            Next
        End If
    ElseIf str�����λ = "Сʱ" Then
        '1:30-2-3:30
        If Not StringMask(strTime, "0123456789:-") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
        
        For i = 0 To UBound(arrTime)
            strTmp = arrTime(i)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) < 1 Or Val(Split(strTmp, ":")(0)) > intƵ�ʼ�� Or Split(strTmp, ":")(0) = "" Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Then Exit Function
            If strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            strPreTime = Format(strTmp, "HH:mm")
        Next
    End If
    
    ExeTimeValid = True
End Function

Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean)
'���ܣ���ComboBox�в��Ҳ���λ
'������blnEvent=��λʱ�Ƿ񴥷�Click�¼�
'˵����δ�ܶ�λʱ,����ListIndex=-1
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If NeedName(objCbo.List(i)) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.Hwnd, i)
            End If
            Exit Sub
        End If
    Next
    If blnEvent Then
        objCbo.ListIndex = -1
    Else
        Call zlControl.CboSetIndex(objCbo.Hwnd, -1)
    End If
End Sub

Public Function GetColFormat(vsGrid As Object) As String
'���ܣ���ȡָ�������и�ʽ���Դ�
    Dim strTmp As String, i As Long
    
    For i = 0 To vsGrid.Cols - 1
        '�п�,�пɼ�,�ж���
        strTmp = strTmp & ";" & vsGrid.ColWidth(i) & "," & IIf(vsGrid.ColHidden(i), 0, 1) & "," & vsGrid.ColAlignment(i)
    Next
    GetColFormat = Mid(strTmp, 2)
End Function

Public Sub SetColFormat(vsGrid As Object, ByVal strFormat As String)
'���ܣ��ָ�ָ�������и�ʽ
    Dim arrCol As Variant, i As Long
    If strFormat = "" Then Exit Sub
    
    arrCol = Split(strFormat, ";")
    For i = 0 To UBound(arrCol)
        vsGrid.ColWidth(i) = Val(Split(arrCol(i), ",")(0))
        vsGrid.ColHidden(i) = Val(Split(arrCol(i), ",")(1)) = 0
        vsGrid.ColAlignment(i) = Val(Split(arrCol(i), ",")(2))
        vsGrid.Cell(2, vsGrid.FixedRows, i, vsGrid.Rows - 1, i) = Val(Split(arrCol(i), ",")(2))
    Next
    vsGrid.Cell(2, 0, 0, vsGrid.FixedRows - 1, vsGrid.Cols - 1) = 4
End Sub

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * vNumber)
End Function

Public Function Between(x, a, B) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < B Then
        Between = x >= a And x <= B
    Else
        Between = x >= B And x <= a
    End If
End Function

Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ��ʱ���Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '������δ���û���ͣ��ʱ��ֹͣ
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function

Public Function DateIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ�������Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
'˵��������ʱ���ж�,����ͣ���ڰ���ʼ����ֹ�����ж�
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Format(Split(arrPause(i), ",")(0), "yyyy-MM-dd")
        strEnd = Format(Split(arrPause(i), ",")(1), "yyyy-MM-dd")
        If strEnd = "" Then strEnd = "3000-01-01" '������δ���û���ͣ��ʱ��ֹͣ
        If strEnd > strBegin Then
            If Between(Format(vDate, "yyyy-MM-dd"), strBegin, _
                Format(DateAdd("d", -1, CDate(strEnd)), "yyyy-MM-dd")) Then
                DateIsPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function TimeisLastPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ��ʱ���Ƿ������һ����ͣ��ʱ����,�����һ����ͣû������
'˵������Ϊ���������,�������û����ֹʱ��,ĳЩ�������ѭ��
    Dim arrPause() As String
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    
    strBegin = Split(arrPause(UBound(arrPause)), ",")(0)
    strEnd = Split(arrPause(UBound(arrPause)), ",")(1)
    If strEnd = "" Then
        strEnd = "3000-01-01 00:00:00"
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeisLastPause = True: Exit Function
        End If
    End If
End Function

Public Function Calc�����ֽ�ʱ��(lng���� As Long, ByVal dat��ʼʱ�� As Date, dat��ֹʱ�� As Date, strPause As String, _
    ByVal strִ��ʱ�� As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As String
'���ܣ�������������εķֽ�ִ��ʱ��,Ҫ��<=��ֹʱ�估������ͣʱ�����
'������dat��ʼʱ��=ҽ���Ŀ�ʼִ��ʱ��
'      dat��ֹʱ��=ҽ����ִ����ֹʱ��,û��ʱ����"3000-01-01"
'      strPause=ҽ������ͣʱ���
'���أ�1."ʱ��1,ʱ��2,...."(yyyy-MM-dd HH:mm:ss)
'      2.lng����=ʵ���ܹ��ֽ�Ĵ���
'˵����1.��Ϊ��ֹʱ�������,��˷ֽ������ʱ���������С��Ҫ�ֽ�Ĵ���
'      2.�������Ǽٶ���ִ��ʱ�估Ƶ��������ȫ��ȷ������¼��㡣
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime() As String, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    vCurTime = dat��ʼʱ��
    arrTime = Split(strִ��ʱ��, "-")
    
    If str�����λ = "��" Then
        vCurTime = GetWeekBase(dat��ʼʱ��)
        Do While lng���� > 0
            '1/8:00-3/15:00-5/9:00
            For i = 1 To intƵ�ʴ���
                vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                    strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                Else
                    strTmp = Split(arrTime(i - 1), "/")(1)
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                If vTmpTime > dat��ֹʱ�� Then
                    Exit Do
                ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                    Exit Do
                ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    lng���� = lng���� - 1
                    If lng���� = 0 Then Exit Do
                End If
            Next
            vCurTime = vCurTime + 7
        Loop
    ElseIf str�����λ = "��" Then
        Do While lng���� > 0
            If intƵ�ʼ�� = 1 Then
                '8:00-12:00-14:00��8-12-14
                For i = 1 To intƵ�ʴ���
                    If InStr(arrTime(i - 1), ":") = 0 Then
                        strTmp = arrTime(i - 1) & ":00"
                    Else
                        strTmp = arrTime(i - 1)
                    End If
                    vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    
                    If vTmpTime > dat��ֹʱ�� Then
                        Exit Do
                    ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                        Exit Do
                    ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        lng���� = lng���� - 1
                        If lng���� = 0 Then Exit Do
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To intƵ�ʴ���
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime > dat��ֹʱ�� Then
                        Exit Do
                    ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                        Exit Do
                    ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        lng���� = lng���� - 1
                        If lng���� = 0 Then Exit Do
                    End If
                Next
            End If
            vCurTime = vCurTime + intƵ�ʼ��
        Loop
    ElseIf str�����λ = "Сʱ" Then
        '10:00-20:00-40:00��10-20-40��02:30
        Do While lng���� > 0
            For i = 1 To intƵ�ʴ���
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                If vTmpTime > dat��ֹʱ�� Then
                    Exit Do
                ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                    Exit Do
                ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    lng���� = lng���� - 1
                    If lng���� = 0 Then Exit Do
                End If
            Next
            vCurTime = vCurTime + intƵ�ʼ�� / 24
        Loop
    End If
    lng���� = UBound(Split(Mid(strDetailTime, 2), ",")) + 1
    Calc�����ֽ�ʱ�� = Mid(strDetailTime, 2)
End Function

Public Function Calc���ڷֽ�ʱ��(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal strִ��ʱ�� As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As String
'���ܣ���ʱ��μ�����εķֽ�ִ��ʱ�估����
'������datBegin-datEnd=Ҫ�����ʱ���,����datBeginӦΪÿ�����ڵĿ�ʼ��׼ʱ��
'      strPause=��ͣ��ʱ���
'���أ�"ʱ��1,ʱ��2,...."(yyyy-MM-dd HH:mm:ss),ʱ�������Ϊ����
'˵����1.ʱ�����Ҫ�ų���ͣ��ʱ���,����������˶�����
'      2.�������Ǽٶ���ִ��ʱ�估Ƶ��������ȫ��ȷ������¼��㡣
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime() As String, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    vCurTime = datBegin
    arrTime = Split(strִ��ʱ��, "-")
    
    If str�����λ = "��" Then
        vCurTime = GetWeekBase(datBegin)
        Do While vCurTime <= datEnd
            '1/8:00-3/15:00-5/9:00
            For i = 1 To intƵ�ʴ���
                vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                    strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                Else
                    strTmp = Split(arrTime(i - 1), "/")(1)
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "��" Then
        Do While vCurTime <= datEnd
            If intƵ�ʼ�� = 1 Then
                '8:00-12:00-14:00��8-12-14
                For i = 1 To intƵ�ʴ���
                    If InStr(arrTime(i - 1), ":") = 0 Then
                        strTmp = arrTime(i - 1) & ":00"
                    Else
                        strTmp = arrTime(i - 1)
                    End If
                    vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To intƵ�ʴ���
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + intƵ�ʼ��, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "Сʱ" Then
        '10:00-20:00-40:00��10-20-40��02:30
        Do While vCurTime <= datEnd
            For i = 1 To intƵ�ʴ���
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = vCurTime + intƵ�ʼ�� / 24
        Loop
    End If
    Calc���ڷֽ�ʱ�� = Mid(strDetailTime, 2)
End Function

Public Function CalcȱʡҩƷ����(ByVal dbl���� As Double, ByVal int�Ƴ� As Integer, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, Optional ByVal strִ��ʱ�� As String, _
    Optional ByVal dbl����ϵ�� As Double, Optional ByVal dbl��װϵ�� As Double, Optional ByVal int���� As Integer) As Double
'���ܣ����Ƴ̼��������Լ���ҩƷ������ȱʡ����(���䷽ȱʡ����)
'������dbl����=��������λ��һ������
'      int�Ƴ�=һ���Ƴ̵�����
'      int����=0-�ɷ���,1-������,2-һ����(��ʱʧЧ),-N-N���ڷ���ʹ����Ч
'      dbl��װϵ��=�����װ��סԺ��װ
'���أ���סԺ��λ�����ҩƷ����
'˵����
'     1.ҩƷ������������������סԺ��װ���Եġ�
'     2.dbl����ϵ��,dbl��װϵ��,int����=��ҩ������,ֻ���㸶��
    Dim dbl��� As Double, dbl���� As Double
    Dim dblʣ�� As Double, dblOne As Double
    Dim intStep As Integer, dblEnd As Double
    Dim arrTime() As String, strBegin As String
    Dim strTime As String, i As Integer, j As Integer
    
    '�Ƴ̲���һ��Ƶ������ʱ�Ͳ����Ƴ�
    If str�����λ = "��" Then
        If int�Ƴ� < 7 Then int�Ƴ� = 1
    ElseIf str�����λ = "��" Then
        If int�Ƴ� < intƵ�ʼ�� Then int�Ƴ� = 1
    ElseIf str�����λ = "Сʱ" Then
        If int�Ƴ� < intƵ�ʼ�� / 24 Then int�Ƴ� = 1
    End If
    
    'һ��Ƶ�����ڵĴ���(����)
    If str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / 7
    ElseIf str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / intƵ�ʼ��
    ElseIf str�����λ = "Сʱ" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * 24
    End If
    
    If dbl����ϵ�� = 0 And dbl��װϵ�� = 0 Then
        '��ҩ����(����) = ����*�Ƴ�*(Ƶ�ʴ���/Ƶ�ʼ��)
        dbl���� = IntEx(int�Ƴ� * dbl���)
    Else
        'ҩƷ�������� = ����/סԺ��װ(����*�Ƴ�*(Ƶ�ʴ���/Ƶ�ʼ��))
        If int���� = 0 Then
            '�ɷ���
            dbl���� = dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��
        ElseIf int���� = 1 Then
            '������
            dbl���� = IntEx(dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��)
        ElseIf int���� = 2 Then
            'һ����(��ʱʧЧ)
            dbl���� = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * IntEx(int�Ƴ� * dbl���)
        ElseIf int���� < 0 Then
            'ABS(int����)���ڷ���ʹ����Ч(�����������)
            If strִ��ʱ�� <> "" Then
                'һ������/סԺ��װ�ļ���
                dblOne = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��)
                'ȱʡִ�еĴ�����ʱ��ֽ�
                strTime = Calc�����ֽ�ʱ��(IntEx(int�Ƴ� * dbl���), Date, CDate("3000-01-01"), "", strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                If strTime <> "" Then
                    arrTime = Split(strTime, ",")
                    dblʣ�� = dblOne: dbl���� = 1
                    strBegin = arrTime(0)
                    
                    '��������
                    For i = 0 To UBound(arrTime)
                        If dblʣ�� < dbl���� Or CDate(arrTime(i)) - CDate(strBegin) > Abs(int����) Then
                            If CDate(arrTime(i)) - CDate(strBegin) > Abs(int����) Then
                                dblʣ�� = dblOne
                            Else
                                dblʣ�� = dblʣ�� + dblOne
                            End If
                            dbl���� = dbl���� + 1
                            strBegin = arrTime(i)
                        End If
                        dblʣ�� = dblʣ�� - dbl����
                    Next
                End If
            End If
        End If
    End If
    CalcȱʡҩƷ���� = dbl����
End Function

Public Function Calc����ҩƷ����(ByVal dat��ʼִ��ʱ�� As Date, lng���� As Long, str�ֽ�ʱ�� As String, _
    ByVal dbl���� As Double, ByVal dbl����ϵ�� As Double, ByVal dbl��װϵ�� As Double, _
    ByVal int���� As Integer, ByVal dat��ֹʱ�� As Date, ByVal strPause As String, ByVal strִ��ʱ�� As String, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As Double
'���ܣ������ʹ������������Լ����ҩ����
'������dat��ʼִ��ʱ��=ҽ���Ŀ�ʼִ��ʱ��,���ڼ�����һִ�����ڿ�ʼ��׼ʱ��
'      lng����=���μƻ�Ҫ���͵Ĵ���
'      dbl����=��������λ��һ������
'      int����=0-�ɷ���,1-������,2-һ����(��ʱʧЧ),-N-N���ڷ���ʹ����Ч(��24Сʱ����)
'      dbl��װϵ��=�����װ��סԺ��װ
'���в������ڲ�����ҩƷ����(����-N��)��
'      str�ֽ�ʱ��=���η��ͼƻ�ִ�еķֽ�ʱ��,�������Ӧ
'      strPause=ҽ������ͣʱ���
'      dat��ֹʱ��=ҽ����ִ����ֹʱ��,û��ʱ����"3000-01-01"
'���أ�1.������/סԺ��λ�����ҩƷ����
'      2.lng����=������ҩƷ(����-N�ͷ���ҩƷ)������ʵ��ִ�д���(����)
'      3.str�ֽ�ʱ��=������ҩƷ(����-N�ͷ���ҩƷ)�����ķֽ�ʱ��(����)
'˵����ҩƷ������������������סԺ��װ���Եġ�
    Dim dbl���� As Double, dblʣ�� As Double
    Dim arrTime() As String, dblOne As Double
    Dim strBegin As String, datBase As Date
    Dim strTmp As String, i As Long
    
    If int���� = 0 Then
        '�ɷ���
        dbl���� = dbl���� * lng���� / dbl����ϵ�� / dbl��װϵ��
    ElseIf int���� = 1 Then
        '������
        dbl���� = IntEx(dbl���� * lng���� / dbl����ϵ�� / dbl��װϵ��)
        '�����������ʱ,����ľ�����ʹ��,�Ӷ�ʹ���ʹ�������
        dblʣ�� = dbl���� * dbl��װϵ�� * dbl����ϵ�� - dbl���� * lng����
        If dblʣ�� >= dbl���� And dbl���� <> 0 Then
            'ʣ�����ۿ���ִ�еĴ���
            i = Int(dblʣ�� / dbl����)
            'ʣ��ʵ�ʿ���ִ�еĴ�����ʱ��ֽ�(����ֹʱ������)
            arrTime = Split(str�ֽ�ʱ��, ",")
            datBase = Calc�����ڿ�ʼʱ��(dat��ʼִ��ʱ��, CDate(arrTime(UBound(arrTime))), intƵ�ʼ��, str�����λ)
            
            '��������չʱ��ʱ,���һ����������ִ�е�ʱ�䲻�ټ���,����ͣ����
            strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
            If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
            
            strTmp = Calc�����ֽ�ʱ��(i, datBase, dat��ֹʱ��, strPause, strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
            If strTmp <> "" Then
                lng���� = lng���� + i
                str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & strTmp
            End If
        End If
    ElseIf int���� = 2 Then
        'һ����(��ʱʧЧ)
        dbl���� = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * lng����
    ElseIf int���� < 0 Then
        'ABS(int����)���ڷ���ʹ����Ч(�����������)
        arrTime = Split(str�ֽ�ʱ��, ",")
        strBegin = arrTime(0)
        
        'һ������/סԺ��װ�ļ���(������λ)
        dblOne = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��)
        'һ������/סԺ��װ�ļ���(��װ��λ)
        dbl���� = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��)
        
        '��������
        dblʣ�� = dblOne
        For i = 0 To UBound(arrTime)
            '��һ��ѭ���϶���,���Բ���������
            If dblʣ�� < dbl���� Or CDate(arrTime(i)) - CDate(strBegin) > Abs(int����) Then
                If CDate(arrTime(i)) - CDate(strBegin) > Abs(int����) Then
                    dblʣ�� = dblOne
                    dbl���� = dbl���� + IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��)
                Else
                    If dblʣ�� + dbl����ϵ�� * dbl��װϵ�� >= dbl���� Then
                        'ֻ��ʣ���һ����װ��λ����
                        dblʣ�� = dblʣ�� + dbl����ϵ�� * dbl��װϵ��
                        dbl���� = dbl���� + 1
                    Else
                        '��Ҫʣ���һ�ΰ�װ��λ�Ź�
                        dblʣ�� = dblʣ�� + dblOne
                        dbl���� = dbl���� + IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��)
                    End If
                End If
                strBegin = arrTime(i)
            End If
            dblʣ�� = dblʣ�� - dbl����
        Next
        
        'ʣ�ಿ�ּ�������Ч���ڰ����������,�Ӷ�ʹ���ʹ�������
        If dblʣ�� >= dbl���� And dbl���� <> 0 Then
            'ʣ�����ۿ���ִ�еĴ���
            i = Int(dblʣ�� / dbl����)
            'ʣ��ʵ�ʿ���ִ�еĴ�����ʱ��ֽ�(����ֹʱ������)
            datBase = Calc�����ڿ�ʼʱ��(dat��ʼִ��ʱ��, CDate(arrTime(UBound(arrTime))), intƵ�ʼ��, str�����λ)
            
            '��������չʱ��ʱ,���һ����������ִ�е�ʱ�䲻�ټ���,����ͣ����
            strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
            If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
            
            strTmp = Calc�����ֽ�ʱ��(i, datBase, dat��ֹʱ��, strPause, strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
            If strTmp <> "" Then
                arrTime = Split(strTmp, ",")
                For i = 0 To UBound(arrTime)
                    If dblʣ�� < dbl���� Or CDate(arrTime(i)) - CDate(strBegin) > Abs(int����) Then
                        Exit For
                    End If
                    lng���� = lng���� + 1
                    str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & arrTime(i)
                    dblʣ�� = dblʣ�� - dbl����
                Next
            End If
        End If
    End If
    Calc����ҩƷ���� = dbl����
End Function

Public Function Calc�����ڿ�ʼʱ��(ByVal dat��ʼִ��ʱ�� As Date, ByVal datĳ��ִ��ʱ�� As Date, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As Date
'���ܣ����ݳ�����ĳ��ִ��ʱ�䣬�õ����ڸ������ڵĿ�ʼ��׼ʱ��
    Dim datBegin As Date, datCurr As Date
    
    datCurr = dat��ʼִ��ʱ��
    datBegin = datCurr
    If str�����λ = "��" Then datCurr = GetWeekBase(datCurr)
    
    Do While datCurr <= datĳ��ִ��ʱ��
        datBegin = datCurr
        If str�����λ = "��" Then
            datCurr = datCurr + 7
        ElseIf str�����λ = "��" Then
            datCurr = datCurr + intƵ�ʼ��
        ElseIf str�����λ = "Сʱ" Then
            datCurr = DateAdd("h", intƵ�ʼ��, datCurr)
        End If
    Loop
    Calc�����ڿ�ʼʱ�� = datBegin
End Function

Public Function Trim�ֽ�ʱ��(ByVal lng���� As Long, ByVal str�ֽ�ʱ�� As String) As String
'���ܣ���ҽ��ִ�еķֽ�ʱ�䰴�������нض�
    Dim arrTime() As String, strTmp As String, i As Long
    
    arrTime = Split(str�ֽ�ʱ��, ",")
    For i = 0 To lng���� - 1
        strTmp = strTmp & "," & arrTime(i)
    Next
    Trim�ֽ�ʱ�� = Mid(strTmp, 2)
End Function

Public Function Calc�����Գ�������(ByVal datBegin As Date, ByVal datEnd As Date, _
    ByVal str�ϴ�ִ��ʱ�� As String, ByVal strִ����ֹʱ�� As String, _
    ByVal strPause As String, Optional str�״�ʱ�� As String, _
    Optional strĩ��ʱ�� As String, Optional str�ֽ�ʱ�� As String) As Long
'���ܣ��Գ����Է�ҩ��������������Ӧ�÷��͵Ĵ���,����ĩʱ��
'������str�ϴ�ִ��ʱ��=��һ�����ڱ��η��͵Ŀ�ʼʱ��
'      strִ����ֹʱ��=��ֹ���첻����
'���أ����θ�ҽ�����͵Ĵ���
'      str�״�ʱ��,strĩ��ʱ��=����yyyy-MM-dd HH:mm:ss
'˵���������Գ���������ÿ�췢��һ�δ���,��������봲λ������(��ͣʱ����ʼ����ֹ;��ֹ���첻����)
    Dim curDate As Date, lng���� As Long, blnSend As Boolean
    
    str�״�ʱ�� = "": strĩ��ʱ�� = "": str�ֽ�ʱ�� = ""
    curDate = CDate(Format(datBegin, "yyyy-MM-dd"))
    Do While curDate <= CDate(Format(datEnd, "yyyy-MM-dd"))
        If Not DateIsPause(curDate, strPause) Then
            blnSend = True
            If str�ϴ�ִ��ʱ�� <> "" Then
                If Format(curDate, "yyyy-MM-dd") <= Format(str�ϴ�ִ��ʱ��, "yyyy-MM-dd") Then
                    blnSend = False 'Ӧ�����ϴ�ִ��ʱ���ִ��
                End If
            End If
            If strִ����ֹʱ�� <> "" Then
                If Format(curDate, "yyyy-MM-dd") >= Format(strִ����ֹʱ��, "yyyy-MM-dd") Then
                    blnSend = False 'ӦС��ִ����ֹʱ���ִ��
                End If
            End If
            If blnSend Then
                lng���� = lng���� + 1
                If str�״�ʱ�� = "" Then
                    str�״�ʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00") '��Ϊ���ִ��
                End If
                strĩ��ʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00")
                str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & strĩ��ʱ��
            End If
        End If
        curDate = curDate + 1
    Loop
    str�ֽ�ʱ�� = Mid(str�ֽ�ʱ��, 2)
    Calc�����Գ������� = lng����
End Function

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

Public Sub GetCboIndex(objCbo As Object, strFind As String, Optional Keep As Boolean)
'���ܣ����ַ�����ComboBox�в�������
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Dim i As Integer
    
    '�Ⱦ�ȷ����
    For i = 0 To objCbo.ListCount - 1
        If objCbo.List(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.List(i)) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '���ģ������
    If strFind <> "" Then
        For i = 0 To objCbo.ListCount - 1
            If InStr(objCbo.List(i), strFind) > 0 Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Sub FindCboIndex(objCbo As Object, lngData As Long, Optional Keep As Boolean)
'���ܣ�����Ŀֵ����ComboBox����Ŀ����
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Dim i As Integer
    
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Function SeekCboIndex(objCbo As Object, lngData As Long) As Long
'���ܣ���ItemData����ComboBox������ֵ
    Dim i As Integer
    
    SeekCboIndex = -1
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                SeekCboIndex = i: Exit Function
            End If
        Next
    End If
End Function

Public Function GetNextControl(ByVal intTab As Integer, ByVal frmForm As Object) As Object
'���ܣ���ȡ��һ�����˳��Ŀؼ�
    Dim objNext As Object, i As Long
    
    '���ұȵ�ǰ�ؼ�TabIndex���
    For i = 0 To frmForm.Controls.Count - 1
        If InStr("TextBox,ComboBox,VSFlexGrid", TypeName(frmForm.Controls(i))) > 0 Then
            If frmForm.Controls(i).Enabled And frmForm.Controls(i).Visible And frmForm.Controls(i).TabStop Then
                If frmForm.Controls(i).TabIndex > intTab Then
                    If objNext Is Nothing Then
                        Set objNext = frmForm.Controls(i)
                    ElseIf frmForm.Controls(i).TabIndex < objNext.TabIndex Then
                        Set objNext = frmForm.Controls(i)
                    End If
                End If
            End If
        End If
    Next
    If objNext Is Nothing Then
        'û�����ұȵ�ǰ�ؼ�TabIndexС��
        For i = 0 To frmForm.Controls.Count - 1
            If InStr("TextBox,ComboBox,VSFlexGrid", TypeName(frmForm.Controls(i))) > 0 Then
                If frmForm.Controls(i).Enabled And frmForm.Controls(i).Visible And frmForm.Controls(i).TabStop Then
                    If frmForm.Controls(i).TabIndex < intTab Then
                        If objNext Is Nothing Then
                            Set objNext = frmForm.Controls(i)
                        ElseIf frmForm.Controls(i).TabIndex < objNext.TabIndex Then
                            Set objNext = frmForm.Controls(i)
                        End If
                    End If
                End If
            End If
        Next
    End If
    Set GetNextControl = objNext
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function Custom_WndMessage(ByVal Hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.x = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.x = 1600
        MinMax.ptMaxTrackSize.y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, Hwnd, Msg, wp, lp)
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
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

Public Function CheckAdviceWindow(ByVal strTitle As String) As Boolean
'���ܣ����ҽ���༭�����Ƿ��Ѿ���
    Dim lngHwnd As Long
    
    '�������ڴ���
    lngHwnd = FindWindow("ThunderFormDC", strTitle)
    If lngHwnd <> 0 Then
        MsgBox "ҽ���༭�����Ѿ��򿪣�������ɵ�ǰ��������ִ�С�", vbInformation, gstrSysName
        Call ShowWindow(lngHwnd, SW_RESTORE)
        Call BringWindowToTop(lngHwnd)
        Exit Function
    End If
    CheckAdviceWindow = True
End Function

Public Function GetWeekBase(ByVal datDate As Date) As Date
'���ܣ���ȡָ��ʱ���������ڵ�����һ��ʱ��
    GetWeekBase = Format(datDate - (Weekday(datDate, vbMonday) - 1), "yyyy-MM-dd 00:00:00")
End Function
Public Function Wndproc(ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'����windows��Ϣ
    Dim hw As Long             '����"ZLPACS Viewer"���
    Dim lngX As Long           '������Ŀ�
    Dim LngY As Long           '������ĸ�
    Dim objPacsCore As Object
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim GetCurr As POINTAPI
    Dim SetCurr As POINTAPI
    Dim FrmReprot As Form
    On Error Resume Next
    If Msg = WM_HOTKEY Then
        If wParam = idHotKey Then
            Dim lp As taLong, i2 As t2Int
            lp.ll = lParam
            LSet i2 = lp
            If (i2.lWord = Modifiers) And i2.hWord = uVirtKey Then
                hw = FindWindow(vbNullString, "ZLPACS Viewer")
                If hw <> 0 Then
                    lngX = Screen.Width / Screen.TwipsPerPixelX
                    LngY = Screen.Height / Screen.TwipsPerPixelY
                    GetCursorPos GetCurr
                    If GetCurr.x < 0 Or GetCurr.x > lngX Then
                        SetCurr.x = (Screen.Width / Screen.TwipsPerPixelX) / 2
                        SetCurr.y = (Screen.Height / Screen.TwipsPerPixelY) / 2
                        SetCursorPos SetCurr.x, SetCurr.y
                        Set FrmReprot = frmPACStation.GetReprotFrm()
                        If Not FrmReprot Is Nothing Then
                            FrmReprot.SetFocus
                        Else
                            frmPACStation.SetFocus
                        End If
                    Else
                        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
                        lngLeft = objPacsCore.GetLeft / Screen.TwipsPerPixelX
                        lngTop = objPacsCore.GetTop / Screen.TwipsPerPixelY
                        lngWidth = objPacsCore.Getwidth / Screen.TwipsPerPixelX
                        lngHeight = objPacsCore.GetHeight / Screen.TwipsPerPixelY
                        Debug.Print lngLeft & " " & lngWidth
                        If lngLeft < 0 Or lngLeft + lngWidth > lngX Then
                            SetCurr.x = (lngLeft + lngWidth) - (lngWidth / 2)
                            SetCurr.y = (lngHeight + lngTop) - (lngHeight / 2)
                            SetCursorPos SetCurr.x, SetCurr.y
                            objPacsCore.SetViewerFocus
                        End If
                    End If
                End If
            End If
        End If
    End If
    '��������ȼ���Ϣ�����ԭ���ĳ���
    Wndproc = CallWindowProc(preWinProc, Hwnd, Msg, wParam, lParam)
End Function

