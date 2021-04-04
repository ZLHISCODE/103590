Attribute VB_Name = "mdlCommon"
Option Explicit


'API����
Public Const GWL_HWNDPARENT = (-8)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
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

Public Function ExistsColObject(Col, index) As Boolean
    '�жϼ������Ƿ����ָ������(�ؼ���)�ĳ�Ա
    On Error GoTo ErrorHandler
    
    Dim v As Variant
    
    If TypeName(Col(index)) = "Collection" Then
        '������Ӧ�ĳ�Ա�Ǽ���ʱ
        ExistsColObject = True
        Exit Function
    Else
        '������Ӧ�ĳ�Ա�ǷǼ���ʱ
        v = Col(index)
        ExistsColObject = True
        Exit Function
    End If
ErrorHandler:
    '�쳣ʱ��ʾ��������Ӧ�ĳ�Ա
    ExistsColObject = False
End Function
Public Function GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '���ݴ�����ַ������зֽ⣬����ָ���ַ����Ⱦ���Ҫ���зֽ⣬������浽������
    '��Σ�strInput-������ַ�����strSplitChar-�ַ��������ݵķָ���
    '���أ����飬���������Ա���ַ����Ȳ�����ָ������
    Dim strArray As Variant
    Dim ArrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '����ָ���ַ�ʱ����Ҫ�ֽ�
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '�޷ָ���ʱ
            strTmp = strInput
            Do While Len(strTmp) > lngLength
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = Mid(strTmp, 1, lngLength)
                strTmp = Mid(strTmp, lngLength + 1)
            Loop
            
            If strTmp <> "" Then
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = strTmp
            End If
        Else
            '�зָ���ʱ
            ArrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(ArrTmp)
        
            For i = 0 To lngCount
                If ArrTmp(i) <> "" Then
                    '�зָ�������Ҫ���ַָ���֮���ַ��������ԣ����ܰѷָ���֮����ַ���
                    If Len(IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)) > lngLength Then
                        ReDim Preserve strArray(UBound(strArray) + 1)
                        strArray(UBound(strArray)) = strTmp
                        strTmp = ArrTmp(i)
                    Else
                        strTmp = IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)
                    End If
                End If
                       
                If i = lngCount Then
                    ReDim Preserve strArray(UBound(strArray) + 1)
                    strArray(UBound(strArray)) = strTmp
                End If
            Next
        End If
    Else
        ReDim Preserve strArray(UBound(strArray) + 1)
        strArray(UBound(strArray)) = strInput
    End If
    
    GetArrayByStr = strArray
End Function
'ȡָ����ͷ����λ��
Public Function GetCol(mshFlex As Object, ByVal ColName As String) As Integer
    Dim i As Integer
    
    GetCol = -1
    
    If TypeName(mshFlex) = "MSHFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    ElseIf TypeName(mshFlex) = "VSFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    End If
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

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
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

'��ȡָ������ĸ�����
Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function

'��ȡָ������ı���
Public Function GetText(ByVal hwndFrm As Long) As String
    Dim strCaption As String * 256
    
    On Error Resume Next
   
    Call GetWindowText(hwndFrm, strCaption, 255)
    GetText = zlStr.TruncZero(strCaption)
End Function

Public Function GetVSFlexRows(ByVal vsfVal As VSFlexGrid, Optional ByVal blnHidden = False) As Long
'--------------------------------------------------------------
'���ܣ���VSFlexGrid��������������ͷ��
'������
'  blnHidden��True��������ص�������False�������ص�������
'���أ�������
'--------------------------------------------------------------
    Dim i As Long, lngRows As Long
    For i = 0 To vsfVal.Rows - 1
        If blnHidden Then
            If vsfVal.RowHidden(i) Then lngRows = lngRows + 1
        Else
            If vsfVal.RowHidden(i) = False Then lngRows = lngRows + 1
        End If
    Next
    GetVSFlexRows = lngRows
End Function

Public Function MoveSpecialChar(ByVal strInputString As String, Optional ByVal blnMoveSpace As Boolean = True) As String
    '1 ȥ��һ���ַ�: " '_%?"����_%?ת��Ϊ��Ӧ��ȫ���ַ�
    '2 ȥ�������ַ�:�˸��Ʊ����С��س�
    '3 blnMoveSpace���Ƿ�ȥ���ַ��еĿո�Ture-ȥ���ո�ע��ͷβ�ո�Ĭ��ȥ��
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intAsc As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '����ת�����ַ�
    
    strText = Trim(strInputString)
    
    If strText = "" Then
        MoveSpecialChar = ""
        Exit Function
    End If
    
    intStrLen = Len(strText)
    
    For n = 1 To intStrLen
        If InStr(GCST_INVALIDCHAR & CST_SPECIALCHAR, Mid(strText, n, 1)) = 0 Then
            strTmp = strTmp & Mid(strText, n, 1)
        Else
            Select Case Mid(strText, n, 1)
                Case "?"
                    strTmp = strTmp & "��"
                Case "%"
                    strTmp = strTmp & "��"
                Case "_"
                    strTmp = strTmp & "��"
            End Select
        End If
    Next
    
    strText = strTmp
    strTmp = ""
    
    intStrLen = Len(strText)
    
    If intStrLen = 0 Then
        MoveSpecialChar = ""
        Exit Function
    End If
        
    For n = 1 To intStrLen
        intAsc = Asc(Mid(strText, n, 1))
        Select Case intAsc
            Case 8, 9, 10, 13
            Case 32
                '�ո���
                If blnMoveSpace = False Then
                    strTmp = strTmp & Mid(strText, n, 1)
                End If
            Case Else
                strTmp = strTmp & Mid(strText, n, 1)
        End Select
    Next
    
    MoveSpecialChar = strTmp
    
End Function

'ת����ֵΪ����
Public Function TranNumToDate(ByVal strNum As String) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
End Function

Public Function ��ͬ����(ByVal sinFirst As Single, ByVal sinSecond As Single) As Boolean
    Dim blnFirst_���� As Boolean, blnSecond_���� As Boolean
    
    ��ͬ���� = False
    
    If sinFirst = 0 Or sinSecond = 0 Then '0��������֮��
        ��ͬ���� = True
        Exit Function
    End If
    
    blnFirst_���� = (sinFirst <= 0)
    blnSecond_���� = (sinSecond <= 0)
    
    ��ͬ���� = (blnFirst_���� = blnSecond_����)
End Function
'------------------------------------------------
'���ܣ� ����ת������
'������
'   strOld��ԭ����
'���أ� �������ɵ�����
'------------------------------------------------
Private Function TranPasswd(strOld As String) As String
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
Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub
