Attribute VB_Name = "mdlCommon"
Option Explicit
Public glngHook As Long
Public gdtBegin As Date
Public glngOld As Long
Public gblnOk As Boolean

'------------------------------------------------------------------------------------------------------------------------------------
'ö������
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum


Public Sub UnHookKBD()
    If glngHook <> 0 Then
    UnhookWindowsHookEx glngHook
    glngHook = 0
    End If
End Sub

Public Function EnableKBDHook()
    If glngHook <> 0 Then
        gdtBegin = Time
        Exit Function
    End If
    gdtBegin = Time
    glngHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf MyKBHFunc, App.hInstance, App.ThreadID)
End Function

Public Function MyKBHFunc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (Time - gdtBegin) * 60 * 60 * 24 < 0.3 Then
        MyKBHFunc = 1 '��ʾҪ�������ѶϢIf wParam = vbKeySnapshot Then '��� ��û�а���PrintScreen��MyKBHFunc = 1 '�����Hook��Ե����ѶϢEnd If
    Else
        MyKBHFunc = 0
    End If
    Call CallNextHookEx(glngHook, iCode, wParam, lParam) '������һ��HookEnd Function
End Function


Public Function NeedName(strList As String, Optional ByVal strSplit As String) As String
'���ܣ��ӱ���������ϴ��з��������
'������strList=����������ϴ�,��"012-�ڿ�","(012)�ڿ�","[012]�ڿ�"
'          strSplit=ָ���ı������Ʒָ����û��ָ������Ĭ�����ȼ����н���,������ֻ�������»������������м�ָ��
'˵��:1-strList��()��[]�ָ����������ʱ��������[����]��(����)��ͷ,�������Ϊ���ֻ���ĸ
'     2-�ָ��������ȼ����س���(Chr(13)��>����(Chr(10))> - > [] > ()
    NeedName = GetNeedName(strList, strSplit)
End Function

Public Function GetNeedName(strList As String, Optional ByVal strSplit As String) As String
'���ܣ��ӱ���������ϴ��з��������
'������strList=����������ϴ�,��"012-�ڿ�","(012)�ڿ�","[012]�ڿ�"
'          strSplit=ָ���ı������Ʒָ����û��ָ������Ĭ�����ȼ����н���,������ֻ�������»������������м�ָ��
'˵��:1-strList��()��[]�ָ����������ʱ��������[����]��(����)��ͷ,�������Ϊ���ֻ���ĸ
'     2-�ָ��������ȼ����س���(Chr(13)��>����(Chr(10))> - > [] > ()
    Dim intType As Integer
    
    intType = gobjComLib.Decode(strSplit, "", 0, Chr(13), 1, Chr(10), 2, "-", 3, "[]", 4, "()", 5, 6)
    
    If intType = 0 Or intType = 1 Then
        '�����ж��Իس����ָ�
        If InStr(strList, Chr(13)) > 0 Then
            GetNeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
            Exit Function
        End If
    End If
    
    If intType = 0 Or intType = 2 Then
        '�Ի��з��ָ�
        If InStr(strList, Chr(10)) > 0 Then
            GetNeedName = LTrim(Mid(strList, InStr(strList, Chr(10)) + 1))
            Exit Function
        End If
    End If
    
    If intType = 0 Or intType = 4 Then
        '��[]�ָ�
        If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
            If IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
                GetNeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
                Exit Function
            End If
        End If
    End If
    
    If intType = 0 Or intType = 5 Then
        '��()�ָ�
        If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
            If IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
                GetNeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
                Exit Function
            End If
        End If
    End If
    If intType = 0 Or intType = 3 Then
        '��-�ָ�
        GetNeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    Else
        GetNeedName = LTrim(Mid(strList, InStr(strList, strSplit) + IIf(InStr(strList, strSplit) = 0, 1, Len(strSplit))))
    End If
End Function

Public Function SetWindowResizeWndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = gWinRect.MinW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = gWinRect.MinH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = gWinRect.MaxW \ Screen.TwipsPerPixelX
        MinMax.ptMaxTrackSize.Y = gWinRect.MaxH \ Screen.TwipsPerPixelY
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SetWindowResizeWndMessage = 1
        Exit Function
    End If
    SetWindowResizeWndMessage = CallWindowProc(glngOld, hWnd, Msg, wp, lp)
End Function


'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    '��ȡָ������ĸ�����
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function



Public Sub TxtSelAll(objTxt As Object)
'���ܣ����༭��ĵ��ı�ȫ��ѡ��
'������objTxt=��Ҫȫѡ�ı༭�ؼ�,�ÿؼ�����SelStart,SelLength����
    
    If Trim(objTxt.Text) = "" Then Exit Sub
    
    If TypeName(objTxt) = "TextBox" Or TypeName(objTxt) = "ComboBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        
        If TypeName(objTxt) = "TextBox" Then
            If objTxt.MultiLine Then
                SendMessage objTxt.hWnd, WM_VSCROLL, SB_TOP, 0
            End If
        End If
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub


Public Function InputIsCard(ByRef txtInput As Object, ByVal KeyAscii As Integer, _
    ByVal blnBrushPassShow As Boolean, Optional blnNumberIsCarded As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ı����е�ǰ�����Ƿ���ˢ��(�Ƿ�ﵽ���ų��ȣ��ڵ��ó������ж�),������ϵͳ���������Ƿ�������ʾ
    '���:KeyAscii=��KeyPress�¼��е��õĲ���
    '       blnBrushPassShow-ˢ���Ƿ񿨺�������ʾ
    '       blnNumberIsCard-����Ĭ��Ϊ��ˢ��,ȱʡΪtrue,��ʾ����Ĭ��Ϊ����ˢ��,false-���ж��ǰ������ٶ����ж��Ƿ�ˢ��
    '����:��ˢ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-02 00:28:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
    
     'ˢ��ʱ����������ŵ��ɵ��÷�ȡ������
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then Exit Function
    
    '����ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    
    '�ж��Ƿ���ˢ��
    '55456:blnNumberIsCard
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) And blnNumberIsCarded Then  '�����������������ȫ���֣���Ϊ��ˢ��
        blnCard = True
    ElseIf KeyAscii > 32 Then
        sngNow = timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '��һ̨�ʼǱ����ԣ�һ����0.014����
        End If
    End If
    
    'ˢ��ʱ�����Ƿ�������ʾ
    If blnCard Then
        txtInput.PasswordChar = IIf(blnBrushPassShow, "*", "")
    Else
        txtInput.PasswordChar = ""
    End If
    InputIsCard = blnCard
End Function


Public Function RPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'���ܣ���ͬOracle��RPAD����
'����:��ָ���������ƿո�
 '������
 '       strText:����ַ���
 '       intCount:����ĳ���
 '       StrPAD:�����ַ�
 '       blnAutoSub:�ַ����������Զ���ȡ
'����:�����ִ�
   
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill Then
        RPAD = strText & String((intCount - lngTmp) \ lngFill, StrPAD)
    ElseIf lngTmp > intCount And blnAutoSub Then
        RPAD = SubB(strText, 1, intCount)
    Else
        RPAD = strText
    End If
End Function

Public Function IsDesinMode() As Boolean
'���ܣ� ȷ����ǰģʽΪ���ģʽ
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
 
Public Function LPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'���ܣ���ͬOracle��LPAD����
 '����:��ָ���������ƿո�
 '������
 '  strText:����ַ���
 '  intCount:����ĳ���
 '  StrPAD:�����ַ�
 '  blnAutoSub:�ַ����������Զ���ȡ
 '����:�����ִ�
 
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill Then
        LPAD = String((intCount - lngTmp) \ lngFill, StrPAD) & strText
    ElseIf lngTmp > intCount And blnAutoSub Then
        LPAD = SubB(strText, 1, intCount)
    Else
        LPAD = strText
    End If
End Function


Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, I As Integer
    
    I = 1
    varValue = arrPar(0)
    Do While I <= UBound(arrPar)
        If I = UBound(arrPar) Then
            Decode = arrPar(I): Exit Function
        ElseIf varValue = arrPar(I) Then
            Decode = arrPar(I + 1): Exit Function
        Else
            I = I + 2
        End If
    Loop
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
    Err = 0
    On Error GoTo errHand:
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
errHand:
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
    On Error GoTo errHand:
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
errHand:
End Sub

Public Function GetControlRect(ByVal lnghwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip/Pixel)
'���أ�blnTwip=True-����Twip��λ��False-�������ص�λ
    Dim vRect As RECT
    Call GetWindowRect(lnghwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function
Public Function GetIDCardDate(strCardID As String) As String
    '���ܣ��������֤�ŷ��س�������
    '������ID=���֤��,Ӧ��Ϊ15λ��18λ
    '���أ���ʽ"yyyy-MM-dd"
    Dim strTmp As String
    
    If gobjCommFun Is Nothing Then Call zlInitCommLib
    If Not gobjCommFun Is Nothing Then
       GetIDCardDate = gobjCommFun.GetIDCardDate(strCardID): Exit Function
    End If
    
    If Len(strCardID) = 15 Then
        strTmp = Mid(strCardID, 7, 6)
        If Len(strTmp) = 6 And IsNumeric(strTmp) Then
            strTmp = "19" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2)
        End If
    ElseIf Len(strCardID) = 18 Then
        strTmp = Mid(strCardID, 7, 8)
        If Len(strTmp) = 8 And IsNumeric(strTmp) Then
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2)
        End If
    End If
    If IsDate(strTmp) Then GetIDCardDate = strTmp
End Function
Public Sub FormSetCaption(ByVal objForm As Variant, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������objForm=��������󣬿��Դ���������������ʱblnCaption=false)
'         blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    Dim lnghwnd As Long
    If IsObject(objForm) Then
        lnghwnd = objForm.hWnd
    Else
        lnghwnd = objForm
    End If
    
    Call GetWindowRect(lnghwnd, vRect)
    lngStyle = GetWindowLong(lnghwnd, GWL_STYLE)
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
    SetWindowLong lnghwnd, GWL_STYLE, lngStyle
    SetWindowPos lnghwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub


Public Function IsNumOrChar(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж�ָ���ַ����Ƿ�ȫ�������ֺ�Ӣ����ĸ���ɣ�������������
    '       ����ĸ�������������ַ�������µļ�⣬isnumbericֻ���ж����֡�
    '��������SSC���ƣ�
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    Dim I As Integer, J As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For I = 1 To Len(Trim(strAsk))
            J = Asc(Mid(Trim(strAsk), I, 1))
            If Not ((J > 47 And J < 58) Or (J > 64 And J < 91) Or (J > 96 And J < 123)) Then
                IsNumOrChar = False
                Exit Function
            End If
        Next
    End If
    IsNumOrChar = True

End Function

Public Function IsCharAlpha(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж�ָ���ַ����Ƿ�ȫ����Ӣ����ĸ����    '
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    Dim I As Integer, J As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For I = 1 To Len(Trim(strAsk))
            J = Asc(Mid(Trim(strAsk), I, 1))
            If Not ((J > 64 And J < 91) Or (J > 96 And J < 123)) Then
                IsCharAlpha = False
                Exit Function
            End If
        Next
    End If
    IsCharAlpha = True
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж�ָ���ַ����Ƿ��к���
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    Dim I As Integer, J As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For I = 1 To Len(Trim(strAsk))
            J = Asc(Mid(Trim(strAsk), I, 1))
            If J < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub


Public Function OpenIme(Optional blnOpen As Boolean = False, Optional strImeName As String) As Boolean
    '����:���������뷨����ر����뷨
    '������strImeName-��ָ�������뷨��û��ָ��ʱ��ϵͳѡ�����õ�ȱʡ���뷨
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    
    If strImeName = "���Զ�����" Then OpenIme = True: Exit Function
    If gobjCommFun Is Nothing Then zlInitCommLib
    If Not gobjCommFun Is Nothing Then
        OpenIme = gobjCommFun.OpenIme(blnOpen, strImeName)
        Exit Function
    End If
    
    
    '�û�û�������ã��Ͳ�����
    If blnOpen Then
        If strImeName <> "" Then
            strIme = Trim(strImeName)
        Else
            strIme = Trim(gobjComLib.zlDatabase.GetPara("���뷨"))
        End If
        If strIme = "" Or strIme = "���Զ�����" Then Exit Function                'Ҫ������뷨��������û������
    End If
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))

    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '��Ҫ�����뷨�������ж��Ƿ�ָ�����뷨
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then
                        OpenIme = True
                        Exit Function
                    End If
                End If
            End If
        ElseIf blnOpen = False Then
            '�����������뷨��������Ӧ�˹ر����뷨������
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnOpen = False Then
        '����windows Vistaϵͳ��Ӣ�����뷨��ImmIsIME���Գ���1�����뷨,���,��Ҫ��������.
        '���˺�:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function
Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

