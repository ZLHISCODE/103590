Attribute VB_Name = "mdlMedical"
Option Explicit

Public Enum COLOR
    
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E0E0
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
End Enum


Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long

'ҽ������
'Public gclsInsure As New clsInsure
Public gblnInsure As Boolean '�Ƿ�����ҽ��
Public gintInsure As Integer

Public gblnStrictCtrl As Boolean
Public glngShareUseID As Long

Public gblnBill���� As Boolean
Public glng����ID As Long
Public gbytBalanceRows As Byte '�����վ����д�

Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrUnitName As String               '�û���λ����
Public gfrmMain As Object

Public gstrSQL As String
Public gstrMatch As String                  '���ݱ��ز�����ƥ��ģʽ��ȷ������ƥ�����
Public gblnOK As Boolean

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ

'HISϵͳ����
Public glngOld As Long, glngFormW As Long, glngFormH As Long

Public Type SYS_PARAM_INFO
    ���ý��С��λ�� As Integer
    �շ�������Ŀƥ�� As String
    ����Ʊ�ݺų��� As Integer
    ���￨���볤�� As Integer
    ���￨��ĸǰ׺ As String
    ���￨������ʾ As Boolean
    ��Ŀ����ƥ�䷽ʽ As Integer '0-˫��;1-����
End Type

Public ParamInfo As SYS_PARAM_INFO

Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"


Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29      '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
End Enum

Public Enum CHECKFORMAT
    �����ʼ�
    ����
    ���֤��
    ��ֵ
    �Զ���
End Enum

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
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & strInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & strInput & "') from dual"
    End If
    On Error GoTo errHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function


Public Sub NewColumn(msf As Object, ByVal vText As String, Optional ByVal vWidth As Single = 1200, Optional ByVal vAlignment As Byte = 9)
    Dim i As Long
    
    msf.Cols = msf.Cols + 1
    i = msf.Cols - 1
    
    msf.TextMatrix(0, i) = vText
    msf.ColWidth(i) = vWidth
    msf.ColAlignment(i) = vAlignment
    
    On Error Resume Next
    msf.ColAlignmentFixed(i) = vAlignment
    
End Sub

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    X = objPoint.X * 15 + objBill.CellLeft
    Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    Dim blnForeColor As Boolean
    Dim blnBkColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("ǰ��ɫ").Name = "ǰ��ɫ")
    blnBkColor = (rsData("����ɫ").Name = "����ɫ")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
        lngRow = 0
    Else
        
        If Val(objMsf.RowData(objMsf.Rows - 1)) <= 0 Then
            lngRow = objMsf.Rows - 2
        Else
            lngRow = objMsf.Rows - 1
        End If
                
    End If
    
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            If Trim(objMsf.TextMatrix(0, lngLoop)) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If strMask <> "" Then
                    objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.NVL(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
                Else
                    objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.NVL(rsData(objMsf.TextMatrix(0, lngLoop)))
                End If
            End If
            
        Next
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("ǰ��ɫ").Value)
        If blnBkColor Then objMsf.Cell(flexcpBackColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("����ɫ").Value)
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '������
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '��С��
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    Dim lngLoop As Long
    Dim strChar As String
    
    strChar = "ZXCVBNMASDFGHJKLQWERTYUIOPzxcvbnmasdfghjklqwertyuiop"
    
    Select Case bytMode
    Case 1          'ȫ����
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 2          'ȫ��ĸ
    
        For lngLoop = 1 To Len(Text)
            If InStr(strChar, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
        
    Case 99
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    gstrSQL = "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1"
    
    gstrSQL = "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical")
    GetMaxLength = rs.Fields(0).DefinedSize

End Function

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'����: װ��������ָ�������������������е���������
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            
            If rsTemp1.Fields.Count > 2 Then
                If Val(rsTemp1.Fields(2).Value) = 1 Then
                    objSource.ListIndex = objSource.NewIndex
                End If
            End If
            
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Sub LocationObj(ByRef objTxt As Object)
    On Error Resume Next
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
End Sub

Public Sub LocationGrid(ByRef vsf As Object, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1)
    
    On Error Resume Next
    
    If lngRow <> -1 Then vsf.Row = lngRow
    If lngCol <> -1 Then vsf.Col = lngCol
    
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

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

Public Sub ResetVsf(objVsf As Object)
    '
    objVsf.Rows = 2
    objVsf.RowData(1) = ""
    objVsf.Cell(flexcpText, 1, 0, 1, objVsf.Cols - 1) = ""
    
    On Error Resume Next
    
    Set objVsf.Cell(flexcpPicture, 1, 0, 1, objVsf.Cols - 1) = Nothing
End Sub

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '���ܣ����¶�������
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '-----------------------------------------------------------------------------------------
    '����:��ȡ����ʱ��
    '����:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    
    Select Case strMode
    Case "��  ʱ"      '��ʱ
        GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function

Public Sub ClearGrid(vsf As Object, Optional ByVal Row As Long = 1)
    '--------------------------------------------------------------------------------------------------------
    '����:����������
    '--------------------------------------------------------------------------------------------------------
    vsf.Rows = Row + 1
    vsf.RowData(Row) = 0
    vsf.Cell(flexcpText, Row, 0, Row, vsf.Cols - 1) = ""
    
End Sub

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Sub DeleteRecord(rs As ADODB.Recordset)
    '-----------------------------------------------------------------------------------
    '����:ɾ����¼��
    '����:rs        Ҫɾ���ļ�¼��
    '����:��
    '-----------------------------------------------------------------------------------
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then rs.MoveFirst
    While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Wend
    
errHand:
End Sub

Public Sub CopyRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset)
    '-----------------------------------------------------------------------------------
    '����:ɾ����¼��
    '����:rs        Ҫɾ���ļ�¼��
    '����:��
    '-----------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    Set rsTo = New ADODB.Recordset
    For lngLoop = 0 To rsFrom.Fields.Count - 1
        rsTo.Fields.Append rsFrom.Fields(lngLoop).Name, rsFrom.Fields(lngLoop).Type, rsFrom.Fields(lngLoop).DefinedSize
    Next
    rsTo.Open
    
    If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.Count - 1
            rsTo.Fields(lngLoop).Value = rsFrom.Fields(lngLoop).Value
        Next
        rsFrom.MoveNext
    Wend
    
errHand:
    
End Sub

Public Sub SelectRow(objVsf As Object, ByVal OldRow As Long, ByVal NewRow As Long, Optional ByVal lngBackColor As Long = -1)
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    On Error Resume Next
    
    If lngBackColor = -1 Then
        lngColor = objVsf.BackColorSel
    Else
        lngColor = lngBackColor
    End If
    
    If OldRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, OldRow, objVsf.FixedCols, OldRow, objVsf.Cols - 1) = objVsf.BackColor
    End If
    
    If NewRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, NewRow, objVsf.FixedCols, NewRow, objVsf.Cols - 1) = lngColor
    End If
    
End Sub

Public Sub DrawLine(pic As PictureBox, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1)
    '��(X1,Y1),(X2,Y2)֮��ʹ��ForeColorɫ��һֱ��
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    
    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)
    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth
End Sub

Public Function FilterRecord(rsTmp As ADODB.Recordset, ByVal strFilter As String) As Boolean
    rsTmp.Filter = ""
    rsTmp.Filter = strFilter
    
    FilterRecord = True
End Function

'Public Sub zlDatabase.ExecuteProcedure(ByVal strSQL As String, ByVal strCaption As String)
''���ܣ�ִ��SQL���
'    Call SQLTest(App.ProductName, strCaption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
'End Sub

Public Function CreateVsf(ByRef objVsf As Object, ByVal strVsf As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim varArray As Variant
    Dim varItem As Variant
    Dim i As Integer
    
    On Error GoTo errHand
    
    objVsf.Cols = 0
    
    varArray = Split(strVsf, ";")
    For lngLoop = 0 To UBound(varArray)
        varItem = Split(varArray(lngLoop), ",")
                
        objVsf.Cols = objVsf.Cols + 1
        i = objVsf.Cols - 1
    
        objVsf.TextMatrix(0, i) = varItem(0)
        objVsf.ColWidth(i) = Val(varItem(1))
        objVsf.ColAlignment(i) = Val(varItem(2))
        objVsf.ColHidden(i) = (Val(varItem(4)) = 0)
        objVsf.Cell(flexcpData, 0, i) = IIf(varItem(5) = "", varItem(0), varItem(5))
        
    Next
    
    CreateVsf = True
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

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
    Do While (lngTop + objVsf.RowHeightMin) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeightMin + IIf(lngIndex = 1, 30, 0)
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

Public Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:������ؼ��Ŀ���
    '����:objVsf Ҫ�����еı��ؼ�����
    '����:���ɹ�����True,���򷵻� False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
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
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function

Public Function GetNextCode(ByVal strTable As String, Optional ByVal strField As String = "����", Optional ByVal strFilter As String = "") As String
    Dim rs As New ADODB.Recordset
    Dim strFormat As String
    
    GetNextCode = "1"
    strFormat = "00000000000000000000"
    gstrSQL = "select nvl(max(" & strField & "),0) as ���� from " & strTable & IIf(strFilter = "", "", " where " & strFilter)

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical")
    If rs.BOF = False Then
        strFormat = IIf(rs!���� = 0, "0000", Mid(strFormat, 1, Len(rs!����)))
        GetNextCode = Format(rs!���� + 1, strFormat)
    End If
End Function

Public Function FillLvw(ByRef objLvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '-------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    LockWindowUpdate objLvw.hWnd
    
    Do While Not rs.EOF
        
        Set objItem = objLvw.ListItems.Add(, "K" & rs("ID").Value, rs("����").Value, rs("ͼ��").Value, rs("ͼ��").Value)
        For lngLoop = 2 To objLvw.ColumnHeaders.Count
            objItem.SubItems(lngLoop - 1) = zlCommFun.NVL(rs(objLvw.ColumnHeaders(lngLoop).Text).Value)
        Next
                        
        rs.MoveNext
    Loop
    
    LockWindowUpdate 0
    
    FillLvw = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function RestoreRow(ByRef objVsf As Object, ByVal strKey As String) As Boolean
    
    Dim lngLoop As Long
        
    For lngLoop = 1 To objVsf.Rows - 1
        If objVsf.RowData(lngLoop) = strKey Then
            objVsf.Row = lngLoop
            Exit Function
        End If
    Next
End Function

Public Function LoadGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal objIls As Object, Optional ByVal blnCharge As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    Dim strField As String
    Dim strIcon As String
    Dim blnField As Boolean
    Dim blnForeColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("ǰ��ɫ").Name = "ǰ��ɫ")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            strField = objMsf.Cell(flexcpData, 0, lngLoop)
            
            If Trim(strField) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If Left(strField, 1) = "[" Then
                
                    strField = Mid(strField, 2, Len(strField) - 2)
                    strIcon = ""
                    
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                    If Not (objIls Is Nothing) Then
                        strIcon = zlCommFun.NVL(rsData(strField))
                        If strIcon <> "" Then
                            Set objMsf.Cell(flexcpPicture, lngRow, lngLoop) = objIls.ListImages(strIcon).Picture
                        End If
                    End If
                    
                    objMsf.Cell(flexcpData, lngRow, lngLoop) = strIcon
                    objMsf.TextMatrix(lngRow, lngLoop) = strIcon
                Else
                
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                     If strMask <> "" Then
                        objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.NVL(rsData(strField)), strMask)
                    Else
                        objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.NVL(rsData(strField))
                    End If
                
                    objMsf.Cell(flexcpData, lngRow, lngLoop, lngRow, lngLoop) = objMsf.TextMatrix(lngRow, lngLoop)
                End If
                
            End If
NextCol:
            '��һ��
        Next
        
pointNext:
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("ǰ��ɫ").Value)
        
        rsData.MoveNext
    Loop
    
    LoadGrid = True
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetCol(ByVal objVsf As Object, ByVal strData As String) As Long
    
    Dim lngLoop As Long
    
    GetCol = -1
    For lngLoop = 0 To objVsf.Cols - 1
        If objVsf.Cell(flexcpData, 0, lngLoop) = strData Then
            GetCol = lngLoop
            Exit Function
        End If
    Next
    
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function GetNextNo(intBillID As Integer) As Variant

    GetNextNo = zlDatabase.GetNextNo(intBillID)
    
End Function

Public Function CheckUsedBill(bytKind As Byte, ByVal lng����ID As Long, Optional ByVal strBill As String) As Long
'���ܣ���鵱ǰ����Ա�Ƿ��п���Ʊ������(���û���),�����ؿ��õ�����ID
'������bytKind=Ʊ��
'      lng����ID=��һ�μ��ʱΪ�������õĹ�������ID,�Ժ�Ϊ�ϴ�ʹ�õ�����ID
'      strBill=Ҫ��鷶Χ��Ʊ�ݺ�
'˵����
'    1.�ڼ�鷶Χʱ,��������ж�������Ʊ��,��ֻҪ������һ��֮�о�����
'    2.�ڼ�鷶Χʱ,����Ҳ�ڼ�鷶Χ֮�ڡ�
'    3.���ж�������ʱ,ȱʡ���ٵ�����,��������,"���ʹ�õ�����"ԭ��
'���أ�
'      ������Ʊ������ID>0
'      0=ʧ��
'      -1:û������(�����δ����)��Ҳû�й���(δ����)
'      -2:���õĹ���������
'      -3:ָ��Ʊ�ݺŲ��ڵ�ǰ���÷�Χ��(������������Ʊ�ݵ����)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    
    '����Ա��ʣ�������Ʊ�ݼ�
    strSQL = _
        "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
        "From Ʊ�����ü�¼" & vbNewLine & _
        "Where Ʊ�� = [1] And ʹ�÷�ʽ = 1 And ʣ������ > 0 And ������ = [2]" & vbNewLine & _
        "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, ��ʼ����"
    Set rsSelf = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, UserInfo.����)
    
    If lng����ID = 0 Then
        '�����е�һ�μ��,��û�����ñ��ع���
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function 'Ҳû������Ʊ��
        '������Ʊ��,������ԭ�򷵻�
        lngReturn = rsSelf!ID
    Else
        '�ϴ�ʹ�õ�����ID���һ�μ��Ĺ���ID,���ж�����
        strSQL = "Select ID,ʹ�÷�ʽ,ʣ������,ǰ׺�ı�,��ʼ����,��ֹ���� From Ʊ�����ü�¼ Where Ʊ��=[1] And ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, lng����ID)
        If rsTmp.BOF = False Then
            If rsTmp!ʹ�÷�ʽ = 2 Then '����,Ҫ�ȿ���û������
                If Not rsSelf.EOF Then
                    '�����õģ�����
                    lngReturn = rsSelf!ID
                Else
                    'û������ȡ����
                    If rsTmp!ʣ������ = 0 Then CheckUsedBill = -2: Exit Function '�����Ѿ�����
                    lngReturn = rsTmp!ID
                    blnTmp = True
                End If
            Else
                '����Ʊ��
                If rsTmp!ʣ������ > 0 Then
                    '��ʣ��
                    lngReturn = rsTmp!ID
                Else
                    '������ʣ�������
                    If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '��������Ҳû��ʣ��
                    lngReturn = rsSelf!ID
                End If
            End If
        End If
    End If
    
    '���Ʊ�ŷ�Χ�Ƿ���ȷ
    If strBill <> "" Then
        If blnTmp Then
            '�ڹ��÷�Χ�ڷ�Χ�ж�
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!��ʼ����) And UCase(strBill) <= UCase(rsTmp!��ֹ����) And Len(strBill) = Len(rsTmp!��ʼ����)) Then
                lngReturn = -3
            End If
        Else
            '�ڿ������÷�Χ���ж�
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '����������,�������������м��
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CheckUsedBill = 0
End Function

Public Function GetNextBill(lng����ID As Long) As String
    '���ܣ�������������ID,��ȡ��һ��ʵ��Ʊ�ݺ�
    '˵����1.��ȡ������Χ�ڵ���ЧƱ��ʱ,���ؿ����û�����
    '      2.�ſ��ѱ���ĺ���
    Dim rsMain As New ADODB.Recordset
    Dim rsDelete As New ADODB.Recordset
    Dim strSQL As String, strBill As String
    
    On Error GoTo errH
    
    strSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ����" & _
        " From Ʊ�����ü�¼ Where Nvl(ʣ������,0)>0 And ID=[1]"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lng����ID)
    
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!��ǰ����) Then
        strBill = UCase(rsMain!��ʼ����)
    Else
        strBill = UCase(IncStr(rsMain!��ǰ����))
    End If
    
    strSQL = "Select Upper(����) as ���� From Ʊ��ʹ����ϸ" & _
        " Where ����=1 And ԭ��=5 And ����>=[2] And ����ID=[1] Order by ����"

    Set rsDelete = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lng����ID, strBill)
    
    Do While True
        '��鷶Χ
        If Left(strBill, Len(zlCommFun.NVL(rsMain!ǰ׺�ı�))) <> UCase(zlCommFun.NVL(rsMain!ǰ׺�ı�)) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!��ʼ����) And strBill <= UCase(rsMain!��ֹ����)) Then
            Exit Function
        End If
                
        '�ſ������
        rsDelete.Filter = "����='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
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

Public Function IntEx(vNumber As Variant) As Variant
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ȡ����ָ����ֵ����С����
    '------------------------------------------------------------------------------------------------------------------
    IntEx = -1 * Int(-1 * vNumber)
End Function

Public Function RePrintBalance(strNo As String, frmParent As Object, lng����ID As Long, ByVal bytKind As Byte) As Boolean
    '���ܣ���ǰ�տ��¼���´�ӡһ��Ʊ��
    Dim strSQL As String
    Dim strInvoice As String
    Dim lng����ID As Long
    Dim blnValid As Boolean
    Dim blnDo As Boolean
    
    '����ϸ����Ʊ��ʹ��
    If gblnStrictCtrl Then
        lng����ID = GetInvoiceGroupID(bytKind, 1, 0, glngShareUseID)
        Select Case lng����ID
            Case -1
                MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
        End Select
        If lng����ID <= 0 Then Exit Function
    End If
    
    blnDo = ReportPrintSet(gcnOracle, glngSys, "ZL1_BILL_1862", frmParent)

    If blnDo Then
        'ȡ��һ��Ʊ�ݺ���
        If Not gblnStrictCtrl Then
            '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
            strInvoice = UCase(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ǰ����Ʊ�ݺ�", ""))
            If strInvoice = "" Then
                '�п����ǵ�һ��ʹ��
                Do
                    strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                    vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                        
                    '�û�ȡ������,�����ӡ
                    If strInvoice = "" Then
                        If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        blnValid = True
                    Else
                        '���������Ч��
                        If zlCommFun.ActualLen(strInvoice) <> ParamInfo.����Ʊ�ݺų��� Then
                            MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & ParamInfo.����Ʊ�ݺų��� & " λ��", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    End If
                Loop While Not blnValid
            Else
                strInvoice = IncStr(strInvoice)
            End If
        Else
            '����Ʊ�����ö�ȡ
            strInvoice = GetNextBill(lng����ID)
            If strInvoice = "" Then
                '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                Do
                    strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                    vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                        
                    '�û�ȡ������,����ӡ
                    If strInvoice = "" Then Exit Function
                    
                    '���������Ч��
                    If GetInvoiceGroupID(bytKind, 1, lng����ID, glngShareUseID, strInvoice) = -3 Then
                        MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Loop While Not blnValid
            End If
        End If
        
        Call frmPrint.ReportPrint(2, strNo, lng����ID, lng����ID, strInvoice, , , , bytKind)
       
        RePrintBalance = True
    End If
End Function

Public Function GetMaxFact(ByVal strNo As String, ByVal bytKind As Byte) As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ�����ʵ��ݷ��������Ʊ�ݺ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    
    'Ӧȡ���һ�δ�ӡ��������
    strSQL = "Select Max(ID) From Ʊ�ݴ�ӡ���� Where ��������=[1] And NO=[2]"
    strSQL = "Select Max(����) as ���� From Ʊ��ʹ����ϸ" & _
        " Where Ʊ��=[1] And ����=1 And ��ӡID=(" & strSQL & ")"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", bytKind, strNo)
    
    If Not rsTmp.EOF Then GetMaxFact = zlCommFun.NVL(rsTmp!����)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetSysParameter(ByVal lngNo As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '����;��ȡϵͳ����
    '����:lngNo     ������
    '����:����ֵ
    '------------------------------------------------------------------------------------------------------------------

    GetSysParameter = zlDatabase.GetPara(lngNo, glngSys, , "")

End Function

Public Function ShowTxtFilter(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����;��ʾ�ı�����ѡ���б�(ֻ�����ı���ؼ�)
    '------------------------------------------------------------------------------------------------------------------
    Dim objPoint As POINTAPI
    Dim strInput As String
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand

    If rsData.BOF Then
        If blnPrompt Then MsgBox "û���ҵ���ƥ��Ľ����", , gstrSysName
        Exit Function                            'û�н����ֱ�ӷ���
    End If
            
    If rsData.RecordCount = 1 And blnFilter Then GoTo Over                    '��Ϊ��������ң����ֻ��һ������ֱ�ӷ���
    
    '������ʼ��
    strInput = "'%" & UCase(objTxt.Text) & "%'"
    Call ClientToScreen(objTxt.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        

    
    If frmSelectDialog.ShowSelect(frmParent, 2, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False) Then GoTo Over
   
    
    Exit Function
    
Over:
    
    Set rsResult = rsData
    
    ShowTxtFilter = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowTxtSelect(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False, _
                                    Optional ByVal bytStyle As Byte = 3) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:������+�б�ṹ
    '����:������2;�ɹ�����1;ȡ������0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim objPoint As POINTAPI
    
    On Error GoTo errHand

    If rsData.BOF Then
        MsgBox "û�п�ѡ������ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objTxt.hWnd, objPoint)
                
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If frmSelectDialog.ShowSelect(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = rsData
        ShowTxtSelect = True
        
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Public Function ShowGrdFilter(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True, _
                                    Optional ByVal blnMuli As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����;��ʾ�ı�����ѡ���б�(ֻ���ڱ��ؼ�)
    '------------------------------------------------------------------------------------------------------------------

    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand


    If rsData.BOF Then
        If blnPrompt Then MsgBox "û���ҵ���ƥ��Ľ����", , gstrSysName
        Exit Function                            'û�н����ֱ�ӷ���
    End If
    If rsData.RecordCount = 1 And blnFilter Then GoTo Over                    '��Ϊ��������ң����ֻ��һ������ֱ�ӷ���
        
    Call ClientToScreen(objVsf.hWnd, objPoint)
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight

    If frmSelectDialog.ShowSelect(frmParent, 2, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, , strSavePath, , False, blnMuli) Then GoTo Over
    
    Exit Function
    
Over:
    
    Set rsResult = rsData
    
    ShowGrdFilter = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowGrdSelect(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional ByVal blnMuliSel As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:������+�б�ṹ,Ӧ���ڱ��ؼ�
    '����:������2;�ɹ�����1;ȡ������0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        MsgBox "û�п�ѡ������ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objVsf.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
    
    If frmSelectDialog.ShowSelect(frmParent, 3, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, , strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = rsData
        ShowGrdSelect = True
        
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Public Function FillTreeData(ByRef objTvw As Object, ByVal rs As ADODB.Recordset, Optional ByVal blnExpand As Boolean = False) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '--------------------------------------------------------------------------------------------------------
    Dim objNode As Node
    
    On Error GoTo errHand
    
    LockWindowUpdate objTvw.hWnd
    
    Do While Not rs.EOF
        
        If IsNull(rs("�ϼ�id").Value) Then
            Set objNode = objTvw.Nodes.Add(, , "K" & zlCommFun.NVL(rs("ID").Value, 0), zlCommFun.NVL(rs("����").Value), rs("ͼ��").Value)
        Else
            Set objNode = objTvw.Nodes.Add("K" & rs("�ϼ�id").Value, tvwChild, "K" & zlCommFun.NVL(rs("ID").Value, 0), zlCommFun.NVL(rs("����").Value), rs("ͼ��").Value)
        End If
        
        objNode.ExpandedImage = rs("��ͼ��").Value
        objNode.Expanded = blnExpand
        
        rs.MoveNext
    Loop
    
    LockWindowUpdate 0
    
    FillTreeData = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function MedicalItemsRecord(ByRef rs As ADODB.Recordset, Optional ByVal bytMode As Byte = 1) As Boolean
    '������¼��,���ڱ���ѡ��������Ŀ
    Set rs = New ADODB.Recordset
    
    With rs
        If bytMode = 1 Then
            .Fields.Append "���", adVarChar, 50
            .Fields.Append "ID", adVarChar, 18
            .Fields.Append "�嵥id", adVarChar, 18
            .Fields.Append "���", adVarChar, 30
            .Fields.Append "����", adVarChar, 50
            .Fields.Append "�����۸�", adVarChar, 50
            .Fields.Append "���۸�", adVarChar, 50
            .Fields.Append "�ۿ�", adVarChar, 50
            .Fields.Append "�������", adVarChar, 50
            .Fields.Append "���㷽ʽ", adVarChar, 50
            .Fields.Append "ִ�п���", adVarChar, 50
            .Fields.Append "�ɼ�����", adVarChar, 50
            .Fields.Append "ִ�п���id", adVarChar, 18
            .Fields.Append "�ɼ���ʽ", adVarChar, 50
            .Fields.Append "�ɼ���ʽid", adVarChar, 18
            .Fields.Append "�ɼ�����id", adVarChar, 18
            .Fields.Append "����걾", adVarChar, 50
            .Fields.Append "��鲿λ", adVarChar, 2000
            .Fields.Append "��鲿λid", adVarChar, 18
            .Fields.Append "�¼�", adVarChar, 1
            .Fields.Append "ǰ��ɫ", adVarChar, 20
            .Fields.Append "ɾ��", adVarChar, 1
            .Fields.Append "����", adVarChar, 1
            .Fields.Append "�Ʒ���ϸ", adVarChar, 4000
            .Fields.Append "ѡ��", adVarChar, 1
        Else
            .Fields.Append "���", adVarChar, 50
            .Fields.Append "ID", adVarChar, 18
            .Fields.Append "IC����", adVarChar, 18
            .Fields.Append "����id", adBigInt, 18, adFldKeyColumn
            .Fields.Append "����", adVarChar, 50
            .Fields.Append "�����", adBigInt, 18
            .Fields.Append "������", adVarChar, 20
            .Fields.Append "�Ա�", adVarChar, 50
            .Fields.Append "���֤��", adVarChar, 50
            .Fields.Append "����״��", adVarChar, 50
            .Fields.Append "��������", adVarChar, 18
            .Fields.Append "���֤", adVarChar, 30
            .Fields.Append "����", adVarChar, 50
            .Fields.Append "����", adVarChar, 50
            .Fields.Append "����", adVarChar, 50
            .Fields.Append "ѧ��", adVarChar, 50
            .Fields.Append "ְҵ", adVarChar, 50
            .Fields.Append "���", adVarChar, 50
            .Fields.Append "��ϵ������", adVarChar, 50
            .Fields.Append "��ϵ�˵绰", adVarChar, 50
            .Fields.Append "�Ǽ�ʱ��", adVarChar, 30
            .Fields.Append "�����ʼ�", adVarChar, 50
            .Fields.Append "��ϵ�˵�ַ", adVarChar, 100
            .Fields.Append "������λ", adVarChar, 100
            .Fields.Append "���￨��", adVarChar, 10
            .Fields.Append "ǰ��ɫ", adVarChar, 30
            .Fields.Append "����ɫ", adVarChar, 30
            .Fields.Append "ɾ��", adVarChar, 1
            .Fields.Append "�¼�", adVarChar, 1
        End If
        .Open
    End With
End Function

Public Sub CopyGrid(ByVal objFrom As Object, ByRef objTo As Object, Optional ByVal lngStartCol As Long = 0)
    
    Dim lngRow As Long
    Dim lngCol As Long
        
    objTo.Rows = objFrom.Rows
    objTo.Cols = objFrom.Cols - lngStartCol
    
    For lngCol = lngStartCol To objFrom.Cols - 1
        objTo.ColWidth(lngCol - lngStartCol) = objFrom.ColWidth(lngCol)
        objTo.MergeCol(lngCol - lngStartCol) = objFrom.MergeCol(lngCol)
    Next
    
    For lngRow = 0 To objFrom.Rows - 1
        objTo.MergeRow(lngRow) = objFrom.MergeRow(lngRow)
        For lngCol = lngStartCol To objFrom.Cols - 1
            objTo.TextMatrix(lngRow, lngCol - lngStartCol) = objFrom.TextMatrix(lngRow, lngCol)
        Next
    Next
    
End Sub

Public Function CheckAllowMedical(ByVal lngKey As Long) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '����Ƿ�������
    gstrSQL = "SELECT 0,�Ƿ�����,��Լ��λid FROM ���ǼǼ�¼ WHERE ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        
        If rs("�Ƿ�����").Value = 1 Then
            If zlCommFun.NVL(rs("��Լ��λid").Value, 0) = 0 Then
                CheckAllowMedical = 1
'                strPrompt = "��ǰ��컹û��ȷ��������Ϣ��"
                Exit Function
            End If
        End If
        
    End If
    
    '����Ƿ�����Ա
    gstrSQL = "SELECT NVL(COUNT(1),0) AS ���� FROM �����Ա���� WHERE ����id IS NOT NULL AND �Ǽ�id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        If rs("����").Value = 0 Then
            'strPrompt = "��ǰ��컹û��ȷ�������Ա��"
            CheckAllowMedical = 2
            Exit Function
        End If
    End If
    
    '����Ƿ��������Ŀ
    gstrSQL = "SELECT B.�������,Sum(Decode(a.�Ǽ�id,Null,0,1)) AS ���� FROM �����Ŀ�嵥 A,������ B WHERE A.�Ǽ�id(+)=B.�Ǽ�id AND A.�������(+)=B.������� AND B.�Ǽ�id=[1] GROUP BY B.�������"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            If rs("����").Value = 0 Then
                CheckAllowMedical = 3
    '            strPrompt = "��ǰ���ġ�" & rs("�������").Value & "�����û��ȷ�������Ŀ��"
                Exit Function
            End If
            rs.MoveNext
        Loop
    Else
        CheckAllowMedical = 3
'        strPrompt = "��ǰ��컹û��ȷ�������Ŀ��"
        Exit Function
    End If
    
    gstrSQL = "SELECT 1 FROM �����Ա���� WHERE ������� NOT IN (SELECT ������� FROM ������ WHERE �Ǽ�id=[1]) AND �Ǽ�id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        CheckAllowMedical = 4
        Exit Function
    End If
    
    CheckAllowMedical = 0
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function InitSysPara() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ������
    '
    '
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim str����С��λ As String
    Dim strTmp As String
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '���ý���λ��
    '��ʾ���ý����㵽С�����ڶ���λ?
    ParamInfo.���ý��С��λ�� = Val(zlDatabase.GetPara(9, glngSys, , "2"))
    If ParamInfo.���ý��С��λ�� > 0 Then
        gstrDec = "0." & String(ParamInfo.���ý��С��λ��, "0")
    Else
        gstrDec = "0"
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '�շ�������Ŀ����ƥ��
    '��1λ1-ȫ����ֻ�����,��2λ1-ȫ��ĸֻ�����
    ParamInfo.�շ�������Ŀƥ�� = zlDatabase.GetPara(44, glngSys, , "11")

    
    '------------------------------------------------------------------------------------------------------------------
    '���￨��ĸǰ׺
    ParamInfo.���￨��ĸǰ׺ = zlDatabase.GetPara(27, glngSys, , "")
    
    '------------------------------------------------------------------------------------------------------------------
    '���￨���볤��,����Ʊ�ݺų���
    strTmp = zlDatabase.GetPara(20, glngSys, , "")
    If strTmp <> "" Then
        If UBound(Split(strTmp, "|")) >= 4 Then ParamInfo.���￨���볤�� = Val(Split(strTmp, "|")(4))
        If UBound(Split(strTmp, "|")) >= 3 Then ParamInfo.����Ʊ�ݺų��� = Val(Split(strTmp, "|")(3))
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '���ز���
    ParamInfo.��Ŀ����ƥ�䷽ʽ = Val(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0"))
    
errHand:
    
End Function

Public Function WriteItems(ByVal rs As ADODB.Recordset, ByRef rsItem As ADODB.Recordset, Optional ByVal bytDo As Byte = 0, Optional ByVal bytMode As Byte = 1) As Boolean
    
    '��ȡ�����Ŀ
    On Error GoTo errHand
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If bytMode = 1 Then
                rsItem.AddNew
                rsItem("���").Value = zlCommFun.NVL(rs("�������").Value)
                rsItem("ID").Value = zlCommFun.NVL(rs("ID").Value)
                rsItem("�嵥id").Value = Val(zlCommFun.NVL(rs("�嵥id").Value))
                rsItem("���").Value = zlCommFun.NVL(rs("���").Value)
                rsItem("����").Value = zlCommFun.NVL(rs("����").Value)
                rsItem("ִ�п���").Value = zlCommFun.NVL(rs("ִ�п���").Value)
                rsItem("���㷽ʽ").Value = zlCommFun.NVL(rs("���㷽ʽ").Value, "1")
                rsItem("�������").Value = zlCommFun.NVL(rs("�������").Value)
                rsItem("�����۸�").Value = Format(zlCommFun.NVL(rs("�����۸�").Value), "0.00##")
                rsItem("���۸�").Value = Format(zlCommFun.NVL(rs("���۸�").Value), "0.00##")
                rsItem("�ۿ�").Value = zlCommFun.NVL(rs("�ۿ�").Value)
                rsItem("ִ�п���id").Value = zlCommFun.NVL(rs("ִ�п���id").Value)
                rsItem("�ɼ���ʽ").Value = zlCommFun.NVL(rs("�ɼ���ʽ").Value)
                rsItem("�ɼ���ʽid").Value = zlCommFun.NVL(rs("�ɼ���ʽid").Value)
                rsItem("�ɼ�����id").Value = zlCommFun.NVL(rs("�ɼ�����id").Value)
                rsItem("�ɼ�����").Value = zlCommFun.NVL(rs("�ɼ�����").Value)
                rsItem("����걾").Value = zlCommFun.NVL(rs("����걾").Value)
                rsItem("��鲿λ").Value = zlCommFun.NVL(rs("��鲿λ").Value)
                rsItem("��鲿λid").Value = zlCommFun.NVL(rs("��鲿λid").Value)
                rsItem("�Ʒ���ϸ").Value = GetPriceList(zlCommFun.NVL(rs("�嵥id").Value))
                
                If bytDo = 1 Then
                    rsItem("�¼�").Value = "1"
                    rsItem("ǰ��ɫ").Value = "16711680"
                    rsItem("ɾ��").Value = ""
                End If
                
                If bytDo = 2 Then
                    rsItem("�¼�").Value = "1"
                    rsItem("ǰ��ɫ").Value = IIf(Val(zlCommFun.NVL(rs("�����嵥id").Value)) = 0, "0", "255")
                    rsItem("ɾ��").Value = ""
                    rsItem("����").Value = zlCommFun.NVL(rs("����").Value)
                End If
            Else
                rsItem.AddNew
                
                rsItem("���").Value = zlCommFun.NVL(rs("���").Value)
                rsItem("����").Value = zlCommFun.NVL(rs("����").Value)
                rsItem("IC����").Value = zlCommFun.NVL(rs("IC����").Value)
                rsItem("������").Value = zlCommFun.NVL(rs("������").Value)
                rsItem("�����").Value = zlCommFun.NVL(rs("�����").Value)
                rsItem("���￨��").Value = zlCommFun.NVL(rs("���￨��").Value)
                rsItem("���֤").Value = zlCommFun.NVL(rs("���֤").Value)
                rsItem("�Ա�").Value = zlCommFun.NVL(rs("�Ա�").Value)
                rsItem("����").Value = zlCommFun.NVL(rs("����").Value)
                rsItem("��������").Value = zlCommFun.NVL(rs("��������").Value)
                rsItem("����״��").Value = zlCommFun.NVL(rs("����״��").Value)
                rsItem("����id").Value = zlCommFun.NVL(rs("����id").Value)
                rsItem("����").Value = zlCommFun.NVL(rs("����").Value)
                rsItem("����").Value = zlCommFun.NVL(rs("����").Value)
                rsItem("ѧ��").Value = zlCommFun.NVL(rs("ѧ��").Value)
                rsItem("ְҵ").Value = zlCommFun.NVL(rs("ְҵ").Value)
                rsItem("���").Value = zlCommFun.NVL(rs("���").Value)
                rsItem("��ϵ������").Value = zlCommFun.NVL(rs("��ϵ������").Value)
                rsItem("��ϵ�˵绰").Value = zlCommFun.NVL(rs("��ϵ�˵绰").Value)
                rsItem("�����ʼ�").Value = zlCommFun.NVL(rs("�����ʼ�").Value)
                rsItem("��ϵ�˵�ַ").Value = zlCommFun.NVL(rs("��ϵ�˵�ַ").Value)
                rsItem("������λ").Value = zlCommFun.NVL(rs("������λ").Value)
                rsItem("�Ǽ�ʱ��").Value = zlCommFun.NVL(rs("�Ǽ�ʱ��").Value)
                
                If bytDo = 1 Then
                    rsItem("�¼�").Value = "1"
                    rsItem("ǰ��ɫ").Value = "16711680"
                    rsItem("ɾ��").Value = ""
                End If
                
                If bytDo = 2 Then
                    rsItem("�¼�").Value = "1"
                    rsItem("ǰ��ɫ").Value = "8388736"
                    rsItem("ɾ��").Value = ""
                End If
            End If
            
            rs.MoveNext
        Loop
    End If
    
    WriteItems = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetPriceList(ByVal lngKey As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand
                
    strSQL = "Select x.*,y.����,y.���㵥λ,z.����,z.�Ƽ�����,z.ִ�п���id,t.���� As ִ�п���,y.���,Decode(x.��׼����,0,0,Null,0,10*x.����/x.��׼����) As �ۿ� " & _
            "From  " & _
                "(Select a.�嵥id,a.�շ�ϸĿid,Sum(a.��׼����) As ��׼����,Sum(a.����) As ���� " & _
                "From �����Ŀ�Ƽ� a " & _
                "Where a.�嵥id = [1] " & _
                "Group By a.�嵥id,a.�շ�ϸĿid " & _
                ") x, " & _
                "�շ���ĿĿ¼ y, " & _
                "�����Ŀ�Ƽ� z,���ű� t " & _
            "Where x.�嵥id = z.�嵥id and t.id(+)=z.ִ�п���id " & _
                  "and x.�շ�ϸĿid=y.id and x.�շ�ϸĿid=z.�շ�ϸĿid "
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If strTmp <> "" Then strTmp = strTmp & ";"
            strTmp = strTmp & zlCommFun.NVL(rs("����")) & ":" & _
                    zlCommFun.NVL(rs("���㵥λ")) & ":" & _
                    zlCommFun.NVL(rs("����")) & ":" & _
                    zlCommFun.NVL(rs("��׼����")) & ":" & _
                    zlCommFun.NVL(rs("����")) & ":" & _
                    zlCommFun.NVL(rs("�շ�ϸĿid")) & ":" & _
                    zlCommFun.NVL(rs("�Ƽ�����")) & ":" & _
                    zlCommFun.NVL(rs("ִ�п���")) & ":" & _
                    zlCommFun.NVL(rs("ִ�п���id")) & ":" & _
                    zlCommFun.NVL(rs("���")) & ":" & _
                    zlCommFun.NVL(rs("�ۿ�"))
            
            rs.MoveNext
        Loop
    End If
                    
    GetPriceList = strTmp
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetTypePriceList(ByVal lngNo As Long, ByVal lngKey As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strSQL = "Select z.*,y.����,y.���㵥λ,x.�ּ�,x.��쵥��,z.�Ƽ�����,z.�ۿ� " & _
                "From " & _
                "( Select a.���,a.������Ŀid,a.�շ�ϸĿid,Sum(c.�ּ�) As �ּ�,Sum(c.�ּ�*Nvl(a.�ۿ�,1)) As ��쵥�� " & _
                  "From �շѼ�Ŀ c, " & _
                       "������ͼƼ� a " & _
                  "Where a.�շ�ϸĿid = c.�շ�ϸĿid " & _
                        "and c.ִ������<=SYSDATE and (c.��ֹ���� IS NULL OR c.��ֹ����>SYSDATE) " & _
                        "and A.���=[1] " & _
                        "and A.������Ŀid=[2] " & _
                  "Group by a.���,a.������Ŀid,a.�շ�ϸĿid " & _
                ") x, " & _
                "�շ���ĿĿ¼ y, " & _
                "������ͼƼ� z " & _
                "Where x.�շ�ϸĿid = y.ID " & _
                      "and z.���=x.��� " & _
                      "and z.������Ŀid=x.������Ŀid " & _
                      "and z.�շ�ϸĿid=x.�շ�ϸĿid"

    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lngNo, lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If strTmp <> "" Then strTmp = strTmp & ";"
            strTmp = strTmp & zlCommFun.NVL(rs("����")) & ":" & _
                    zlCommFun.NVL(rs("���㵥λ")) & ":" & _
                    zlCommFun.NVL(rs("����")) & ":" & _
                    zlCommFun.NVL(rs("�ּ�")) & ":" & _
                    zlCommFun.NVL(rs("�շ�ϸĿid")) & ":" & _
                    zlCommFun.NVL(rs("�Ƽ�����")) & ":" & _
                    zlCommFun.NVL(rs("��쵥��")) & ":" & _
                    10 * zlCommFun.NVL(rs("�ۿ�"), 0)
            
            rs.MoveNext
        Loop
    End If
                    
    GetTypePriceList = strTmp
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetCombList(ByVal strSQL As String) As String
    
    Dim rs As New ADODB.Recordset
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical")
    If rs.BOF = False Then
        Do While Not rs.EOF
            GetCombList = GetCombList & "|" & zlCommFun.NVL(rs.Fields(0).Value)
            rs.MoveNext
        Loop
    End If
    If GetCombList = "" Then
        GetCombList = " |"
    Else
        GetCombList = Mid(GetCombList, 2)
    End If
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function GetBirth(ByVal intYear As Integer, ByRef strStart As String, ByRef strEnd As String) As Boolean
        
    strStart = Format(DateAdd("yyyy", 0 - intYear - 1, Now), "yyyy-MM-dd")
    strEnd = Format(DateAdd("yyyy", 0 - intYear, Now), "yyyy-MM-dd")
    
End Function

Public Function CheckStrValid(ByVal Text As String, ByVal bytMode As CHECKFORMAT, Optional ByVal KeyCustom As String, Optional ByVal intLen As Integer = 0, Optional ByVal intDec As Integer = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    Select Case bytMode
    Case CHECKFORMAT.�����ʼ�
        
        If Trim(Text) <> "" Then
            If InStr(Text, "@") = 0 Then Exit Function
            If InStr(Text, "@") = 1 Then Exit Function
            If InStr(Text, "@") = Len(Text) Then Exit Function
        End If
        
    Case CHECKFORMAT.����
    
        If Trim(Text) <> "" Then
            If IsDate(Trim(Text)) = False Then Exit Function
        End If
        
    Case CHECKFORMAT.���֤��
        
        'ֻ�ܰ��� 0,1,2,3,4,5,6,7,8,9,X �ַ�
        
        If Trim(Text) <> "" Then
            If Len(Text) <> 15 And Len(Text) <> 18 Then Exit Function
            
            For lngLoop = 1 To Len(Text)
                If InStr("0123456789X", UCase(Mid(Text, lngLoop, 1))) = 0 Then Exit Function
            Next
            
        End If
    Case CHECKFORMAT.��ֵ
        
    Case CHECKFORMAT.�Զ���
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then Exit Function
        Next
    End Select
    
    CheckStrValid = True
End Function

Public Function Lpad(ByVal strText As String, ByVal lngLen As Long, ByVal strReplace As String) As String
    Dim lngL As Long
    
    lngL = Len(strText)
    If lngL > lngLen Then
        Lpad = Left(strText, lngLen)
    ElseIf lngL < lngLen Then
        Lpad = String(lngLen - lngL, strReplace) & strText
    Else
        Lpad = strText
    End If
End Function

Public Function EnterFocus(obj As Object) As Boolean
    
    On Error Resume Next
    
    obj.SetFocus
    
End Function

Public Function HaveExcel() As Boolean
    '------------------------------------------------
    '���ܣ��жϱ�����װ��EXCELû��
    '������
    '���أ����򷵻�True
    '------------------------------------------------

    On Error GoTo errHandle
    
    Dim objTemp  As Object
    
    Set objTemp = CreateObject("Excel.Application") '��һ��EXCEL����
    
    Set objTemp = Nothing
    
    HaveExcel = True
    
    Exit Function

errHandle:
    Set objTemp = Nothing
    HaveExcel = False
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "EXECUTE", adTinyInt
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String) As Boolean
    
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Sub DrawPicture(pic As Object, objPic As StdPicture, ByVal W As Long, ByVal H As Long)
'���ܣ���PictureBox���밴�ʵ�������һ��ͼ
'������W,H=Ҫ��ͼ�ĳߴ�
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    If W <= pic.ScaleWidth And H <= pic.ScaleHeight Then
        lngW = W: lngH = H
    Else
        sngW = W / pic.ScaleWidth
        sngH = H / pic.ScaleHeight
        If sngW > sngH Then
            lngW = W / sngW: lngH = H / sngW
        Else
            lngW = W / sngH: lngH = H / sngH
        End If
    End If
    
    pic.Cls
    On Error Resume Next
    pic.PaintPicture objPic, (pic.ScaleWidth - lngW) / 2, (pic.ScaleHeight - lngH) / 2, lngW, lngH
    
End Sub

Public Function ReadPicture(rsTable As ADODB.Recordset, strField As String, Optional strFile As String) As String
'-------------------------------------------------------------
'���ܣ���ָ���ļ�¼��ͼ���ֶθ���Ϊͼ����ʱ�ļ�
'������
'       rsTable   ͼ�δ洢��¼��
'       strField  ͼ���ֶ�
'       strFile   �û�������ļ�������ѡ�
'���أ�
'-------------------------------------------------------------
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, FileNum, j
    Dim aryChunk() As Byte
    Dim strTempFile As String
    
    On Error GoTo errH
    lngFileSize = rsTable.Fields(strField).ActualSize
    If lngFileSize = 0 Then
        'δ��ȡ��Ч����
        Exit Function
    End If
    
    FileNum = FreeFile
    If strFile = "" Then
        '���û���û�����ļ���ʱ
'        j = 0
        
        strFile = CreateTmpFile
        
'        Do While True
'            strTempFile = CurDir & "\zlNewPicture" & CStr(j) & ".pic"
'            If Len(Dir(strTempFile)) = 0 Then Exit Do
'            j = j + 1
'        Loop
'        strFile = strTempFile
    End If
    Open strFile For Binary As FileNum
    
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    rsTable.Move 0
    For j = 0 To intBolcks
        If j = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim aryChunk(lngCurSize - 1) As Byte
        aryChunk() = rsTable.Fields(strField).GetChunk(lngCurSize)
        Put FileNum, , aryChunk()
    Next
    Close FileNum
    ReadPicture = strFile
    Exit Function

errH:
    Close FileNum
'    Kill strFile
    ReadPicture = ""

End Function

Public Function GetTmpPath() As String
    
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
End Function

Public Function CreateTmpFile(Optional ByVal strFileType As String = "tmp") As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '����:
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strFileTemp As String
       
    
    strFileTemp = GetTmpPath
    
    strFileTemp = strFileTemp & "zlNewPic" & Format(Now, "yyyymmdd") & Format(Timer, "0") & "." & strFileType
    
    CreateTmpFile = strFileTemp
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'���ܣ��ж��Ƿ����ָ�������������͵�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���ID From ҩƷ�������� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = zlCommFun.NVL(rsTmp!���ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatientID(ByVal strIC As String) As Long
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����id From ������Ϣ Where IC����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", strIC)
    If Not rsTmp.EOF Then GetPatientID = zlCommFun.NVL(rsTmp!����id, 0)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcCharge(ByVal rsSource As ADODB.Recordset, ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim dbTmpʵ�ս�� As Double
    Dim dbTmp�ѽ��� As Double
    
    Dim dbӦ�ս��_�� As Double
    Dim dbӦ�ս��_�� As Double
    
    Dim dbʵ�ս�� As Double
    Dim db���ʽ�� As Double
    Dim db�շѽ�� As Double
    Dim dbδ���� As Double
    Dim dbδ�ս�� As Double
    Dim dbδ����ϼ� As Double
    Dim db�ѽ��� As Double
    
    On Error GoTo errH
    
    If rsSource.BOF Then Exit Function
    
    Set rs = New ADODB.Recordset
    With rs
    
        .Fields.Append "Ӧ�ս��_��", adVarChar, 30
        .Fields.Append "Ӧ�ս��_��", adVarChar, 30
        
        .Fields.Append "ʵ�ս��", adVarChar, 30
        .Fields.Append "���ʽ��", adVarChar, 30
        .Fields.Append "�շѽ��", adVarChar, 30
        
        .Fields.Append "δ����ϼ�", adVarChar, 30
        .Fields.Append "δ����", adVarChar, 30
        .Fields.Append "δ�ս��", adVarChar, 30
        .Open
    End With
    
    Do While Not rsSource.EOF
        
        dbTmpʵ�ս�� = zlCommFun.NVL(rsSource("ʵ�ս��").Value, 0)
        dbTmp�ѽ��� = zlCommFun.NVL(rsSource("���ʽ��").Value, 0)
        
        dbʵ�ս�� = dbʵ�ս�� + dbTmpʵ�ս��
        db�ѽ��� = db�ѽ��� + dbTmp�ѽ���
        
        If zlCommFun.NVL(rsSource("���ʷ���").Value, 0) = 1 Then
            db���ʽ�� = db���ʽ�� + dbTmpʵ�ս��
            dbδ���� = dbδ���� + (dbTmpʵ�ս�� - dbTmp�ѽ���)
            
            dbӦ�ս��_�� = dbӦ�ս��_�� + zlCommFun.NVL(rsSource("Ӧ�ս��").Value, 0)
        Else
            dbӦ�ս��_�� = dbӦ�ս��_�� + zlCommFun.NVL(rsSource("Ӧ�ս��").Value, 0)
        End If
        
        rsSource.MoveNext
    Loop
    
    db�շѽ�� = dbʵ�ս�� - db���ʽ��
    dbδ����ϼ� = dbʵ�ս�� - db�ѽ���
    dbδ�ս�� = dbδ����ϼ� - dbδ����
    
    
    rs.AddNew
    
    rs("Ӧ�ս��_��").Value = dbӦ�ս��_��
    rs("Ӧ�ս��_��").Value = dbӦ�ս��_��
    
    rs("ʵ�ս��").Value = dbʵ�ս��
    rs("���ʽ��").Value = db���ʽ��
    rs("�շѽ��").Value = db�շѽ��
    
    rs("δ����ϼ�").Value = dbδ����ϼ�
    rs("δ����").Value = dbδ����
    rs("δ�ս��").Value = dbδ�ս��
    rs.Update
    
    CalcCharge = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InputIsCard(ByVal strText As String, KeyAscii As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ��ж�ָ���ı����е�ǰ�����Ƿ���ˢ��,���ݴ���������ʾ
    '������
    '���أ�
    '******************************************************************************************************************

'    Dim strText As String
    Dim blnCard As Boolean
    Dim arrMask As Variant
    Dim intLoop As Integer

    '��ǰ�������ʾ������(��δ��ʾ����)
'    strText = txtInput.Text
'    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = UCase(strText & Chr(KeyAscii))
    End If
'    Debug.Print strText
        
    '�ж��Ƿ���ˢ��
    blnCard = False
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then
        blnCard = True
    ElseIf ParamInfo.���￨��ĸǰ׺ <> "" Then
        arrMask = Split(ParamInfo.���￨��ĸǰ׺, "|")
        For intLoop = 0 To UBound(arrMask)
            If strText Like arrMask(intLoop) & "*" Then
                If IsNumeric(Mid(strText, Len(arrMask(intLoop)) + 1)) And IsNumeric(Mid(strText, Len(arrMask(intLoop)) + 1, 1)) Then
                    blnCard = True
                End If
            End If
        Next
    End If
    
    'ˢ��ʱ�����Ƿ�������ʾ
'    If blnCard Then
'        txtInput.PasswordChar = IIf(gblnShowCard, "", "*")
'    Else
'        txtInput.PasswordChar = ""
'    End If
    
    InputIsCard = blnCard
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, Optional ByVal strBill As String) As Long
'���ܣ���ȡ�������ò���ָ��Ʊ��������÷�Χ�ڵ�����ID
'������bytKind      =   Ʊ��
'      intNum       =   Ҫ��ӡ��Ʊ������
'      lngLastUseID =   �ϴ�ʹ�õ�����ID
'      lngShareUseID=   ���ز���ָ���Ĺ���ID
'      strBill      =   ��ǰƱ�ݺţ����ڼ���������ε�Ʊ�ݷ�Χ
'���أ�
'      >0   =   �ɹ������õ�����ID
'      =0   =   ʧ��
'      -1   =   û������(����򲻹�����δ����),δ���ù���
'      -2   =   û������(����򲻹�����δ����),���õĹ���������򲻹�
'      -3   =   ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
'      -4   =   ָ�����ε�Ʊ�ݲ�����
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo errH
    '1.�ϴε����������Ƿ���ò�����
    If lngLastUseID > 0 Then
        strSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
                 "From Ʊ�����ü�¼ Where Ʊ��=[1] And ʣ������>=[2] And ID=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, intNum, lngLastUseID)
        With rsTmp
            If .RecordCount > 0 Then    'Ŀǰ��Ʊ�ݺſ��ܺ��ϴβ�ͬ��������Ҫ��鷶Χ
                If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '����û�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
                
            ElseIf intNum > 1 Then  '����ȷ���������ε���ʱ,��ǰƱ�ݺ��������β�����
                GetInvoiceGroupID = -4: Exit Function
            End If
        End With
    End If
        
    '2.�ϴε��������β����û򲻿���ʱ,ȡ������Ĳ������õ�
    '  �ж��������ʹ�õ�����,�ٵ�����,��������
    strSQL = "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����" & vbNewLine & _
        "From Ʊ�����ü�¼" & vbNewLine & _
        "Where Ʊ�� = [1] And ʣ������ >= [2] And ������ = [3] And ʹ�÷�ʽ = 1" & vbNewLine & _
        "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, ʣ������, �Ǽ�ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, intNum, UserInfo.����)
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then GetInvoiceGroupID = !ID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
            blnTmp = False
            strPre = "" & !ǰ׺�ı�
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = !ID: Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.û�����õ�,ʹ�ñ��ز���ָ���Ĺ�������
    If lngShareUseID > 0 Then
        strSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
                 "From Ʊ�����ü�¼ Where Ʊ��=[1] And ʣ������>=[2] And ID=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, intNum, lngShareUseID)
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    
    GetInvoiceGroupID = lngReturn   '����δ�ҵ���ԭ�����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStorage(ByVal lngKey As Long, ByVal lngDeptKey As Long) As Single
    '----------------------------------------------------------------------
    '����:��ȡҩƷ���
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    GetStorage = 0
        
    If lngKey = 0 Then Exit Function
    
    On Error GoTo errHand
    
    strSQL = "SELECT I.�Ƿ���,S.ҩ������ AS ҩ����������,S.����ϵ�� FROM �շ���ĿĿ¼ I,ҩƷ��� S WHERE I.ID=S.ҩƷid AND S.ҩƷid=[1]"
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
                        
        GetStorage = CalcStorage(lngKey, lngDeptKey, IIf(zlCommFun.NVL(rs("�Ƿ���").Value, 0) = 0, False, True), IIf(zlCommFun.NVL(rs("ҩ����������").Value, 0) = 0, False, True))
    
    End If
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetStock(ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long) As Double
'���ܣ���ȡָ��ҩ��ָ��ҩƷ���(�����۵�λ)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '��ȡ���(�����������),ҩ��������(����=0,����Ϊҩ��)����Ч��
    strSQL = _
        " Select Nvl(Sum(A.��������),0) as ��� From ҩƷ��� A" & _
        " Where (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
        " And A.����=1 And A.ҩƷID=[1] And A.�ⷿID=[2]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngҩƷID, lngҩ��ID)
    If Not rsTmp.EOF Then GetStock = rsTmp!���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcStorage(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal vChangePrice As Boolean, ByVal vBatch As Boolean) As Single

    '���ܣ���ȡָ��ҩ��ָ��ҩƷ���(�����۵�λ)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '��ȡ���(�����������),ҩ��������(����=0,����Ϊҩ��)����Ч��
    strSQL = _
        " Select Nvl(Sum(A.��������),0) as ��� From ҩƷ��� A" & _
        " Where (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
        " And A.����=1 And A.ҩƷID=[1] And A.�ⷿID=[2]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lngҩƷID, lng�ⷿID)
    If Not rsTmp.EOF Then CalcStorage = rsTmp!���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
'    Dim rs As New ADODB.Recordset
'    Dim strSQL As String
'
'
'    If lngҩƷID = 0 Then Exit Function
'
'    If vChangePrice And vBatch = False Then
'        'ֻ��ʵ��ҩƷ
'
'        strSQL = "SELECT NVL(A.��������,0) AS �������� FROM ҩƷ��� A WHERE A.ҩƷid=[1] AND A.�ⷿID=[2]"
'
'    ElseIf vChangePrice = False And vBatch Then
'        'ֻ��ҩ����������ҩƷ
'
'        strSQL = "Select Sum(Nvl(��������,0)) as �������� From ҩƷ���" & _
'                    " Where ����=1 " & _
'                    " And (Ч�� Is NULL Or Ч��>Trunc(Sysdate)) " & _
'                    " And �ⷿID=[2]" & _
'                    " And ҩƷID=[1]"
'
'    ElseIf vChangePrice And vBatch Then
'        '����ʵ��ҩƷ����ҩ����������ҩƷ
'
'        strSQL = "Select Sum(Nvl(��������,0)) as �������� From ҩƷ���" & _
'                    " Where ����=1 " & _
'                    " And (Ч�� Is NULL Or Ч��>Trunc(Sysdate)) " & _
'                    " And �ⷿID=[2]" & _
'                    " And ҩƷID=[1]"
'
'    Else
'        '�Ȳ���ʵ��ҩƷ�ֲ���ҩ����������ҩƷ,��ֻ��ʵ��ҩƷһ����
'
'        strSQL = "SELECT NVL(A.��������,0) AS �������� FROM ҩƷ��� A WHERE A.ҩƷid=[1] AND A.�ⷿID=[2]"
'
'    End If
'
'    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lngҩƷID, lng�ⷿID)
'
'    If rs.BOF = False Then CalcStorage = zlCommFun.NVL(rs("��������").Value, 0)

End Function

Public Function PromptStorageWarn(ByVal dbInput As Double, _
                                    ByVal dbStorage As Double, _
                                    ByVal strDrugName As String, _
                                    ByVal strExecuteDept As String, _
                                    ByVal strUnit As String, _
                                    Optional ByVal bytWarn As Byte = 1, _
                                    Optional ByVal bytApply As Byte = 1) As Integer
    '******************************************************************************************************************
    '���ܣ�
    '������bytWarn��0-�����;1-���,��������;2-��飬�����
    '���أ�
    '******************************************************************************************************************

    If dbInput > 0 And dbInput > dbStorage Then
        
        If bytApply = 1 Then
            Call ShowSimpleMsg("ҩƷ��" & strDrugName & "���ڿⷿ��" & strExecuteDept & "��ֻ��" & dbStorage & strUnit & "��")
            bytWarn = 0
        Else
            Select Case bytWarn
            Case 0
                
            Case 1
                If MsgBox("ҩƷ��" & strDrugName & "���ڿⷿ��" & strExecuteDept & "��ֻ��" & dbStorage & strUnit & "���Ƿ������", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
                    bytWarn = 0
                Else
                    bytWarn = 1
                End If
            Case 2
                MsgBox "ҩƷ��" & strDrugName & "���ڿⷿ��" & strExecuteDept & "��ֻ��" & dbStorage & strUnit & "�������ֹ��", vbOKOnly + vbCritical, gstrSysName
                bytWarn = 1
            End Select
        End If
        
    End If
    
    PromptStorageWarn = bytWarn
    
End Function

Public Function MakeMedicalCharge(ByRef rsSQL As ADODB.Recordset, ByVal lng�Ǽ�id As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    Dim lngCount As Long
    Dim dbSum As Double
    Dim dbӦ�ս�� As Double
    Dim dbʵ�ս�� As Double
    Dim db������� As Double
    Dim int����� As Double
    Dim int��� As Integer
    Dim lng���id As Long
    Dim strNow As String
    Dim int���� As Integer
    Dim int�������� As Integer
    Dim obj As Object
    
    On Error GoTo errHand
    
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    strSQL = "Select x.*,y.���� As ִ�п��� From (SELECT d.��� As �շ����,d.���㵥λ,a.�շ�ϸĿid,Decode(a.ִ�п���id,b.ִ�п���id,c.ִ�п���id,a.ִ�п���id) As ִ�п���id," & _
                    "a.��׼����*A.���� As Ӧ�ս��,a.����*A.���� As ʵ�ս��,a.����,Nvl(a.����,1) As �շ�����,a.��׼����,a.�Ƽ�����, " & _
                    "e.ID As ҽ��id,e.�������,Decode(e.�������,'E',c.�ɼ�No,c.No) As No,b.����;��,f.����id,f.����,f.�Ա�,f.����,f.�ѱ�,f.�����,e.��������id,e.ҽ������,d.���� As �շ���Ŀ " & _
            "FROM �����Ŀ�Ƽ� a, " & _
                "�����Ŀ�嵥 b, " & _
                "�����Ŀҽ�� c, " & _
                "�շ���ĿĿ¼ d, " & _
                "����ҽ����¼ e, " & _
                "������Ϣ f " & _
            "Where b.�Ǽ�id = [1] " & _
             "And d.ID = a.�շ�ϸĿID " & _
             "And C.�嵥id=b.ID " & _
             "And c.��ʱ���=1 " & _
             "And c.ҽ��id In (e.id,e.���id) " & _
             "And ((e.�������='E' And a.�Ƽ�����=2) Or (e.�������='C' And a.�Ƽ�����<>2) Or (e.�������='D' And a.�Ƽ�����<>2 And e.���id Is Null)) " & _
             "And c.����id=f.����id " & _
             "And b.ID = a.�嵥id) x,���ű� y Where x.ִ�п���id=y.ID " & _
            "Order By x.�շ�ϸĿid"
    
    Set rsCharge = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lng�Ǽ�id)
    If rsCharge.BOF = False Then
        Do While Not rsCharge.EOF
            If Val(rsCharge("�Ƽ�����").Value) > 0 And zlCommFun.NVL(rsCharge("No").Value) <> "" Then
                
                int���� = 0
                lng���id = 0
                If rsCharge("�շ����").Value = "4" Then
                    int�������� = 0
                    
                    strSQL = "Select �������� From �������� Where ����ID=[1]"
                    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", Val(rsCharge("�շ�ϸĿid").Value))
                    If rs.BOF = False Then int�������� = zlCommFun.NVL(rs("��������").Value, 0)
                    If int�������� = 1 Then
                        int���� = IIf(zlCommFun.NVL(rsCharge("����;��").Value, 1) = 1, 41, 42)
                    End If
                ElseIf InStr("567", rsCharge("�շ����").Value) > 0 Then
                    int���� = IIf(zlCommFun.NVL(rsCharge("����;��").Value, 1) = 1, 9, 8)
                End If
                
                If int���� > 0 Then
                    strSQL = "Select ���id From ҩƷ�������� Where ����=[1]"
                    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", int����)
                    If rs.BOF = False Then lng���id = zlCommFun.NVL(rs("���id").Value, 0)
                    If lng���id = 0 Then
                        If rsCharge("�շ����").Value = "4" Then
                            ShowSimpleMsg "����ȷ�����ϴ������ݵ�������,���ȵ���������������ã�"
                        Else
                            ShowSimpleMsg "����ȷ��ҩƷ�������ݵ�������,���ȵ���������������ã�"
                        End If
                        Exit Function
                    End If
                End If
                
                '���ҩƷ����Ͽ��
                If InStr("4567", rsCharge("�շ����").Value) > 0 Then
                    
                    db������� = GetStorage(Val(rsCharge("�շ�ϸĿid").Value), Val(rsCharge("ִ�п���id").Value))
                    If Val(rsCharge("�շ�����").Value) > db������� Then
                    
                        int����� = 0
                        If rsCharge("�շ����").Value = "4" Then
                            strSQL = "Select ��鷽ʽ From ���ϳ����� Where �ⷿID=[1]"
                        Else
                            strSQL = "Select ��鷽ʽ From ҩƷ������ Where �ⷿID=[1]"
                        End If
                        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", Val(rsCharge("ִ�п���id").Value))
                        If rs.BOF = False Then int����� = zlCommFun.NVL(rs("��鷽ʽ").Value, 0)
                        
                        '0-�����;1-���,��������;2-��飬�����
                        Select Case int�����
                        Case 0
                            
                        Case 1
                            Set obj = frmWait.mfrmMain
                            
                            If Not (obj Is Nothing) Then Unload frmWait
                            If PromptStorageWarn(Val(rsCharge("�շ�����").Value), db�������, rsCharge("�շ���Ŀ").Value, rsCharge("ִ�п���").Value, rsCharge("���㵥λ").Value, int�����, 2) <> 0 Then
                                If Not (obj Is Nothing) Then Call frmWait.OpenWait(obj, "���Ե�...")
                                Exit Function
                            End If
                            If Not (obj Is Nothing) Then Call frmWait.OpenWait(obj, "���Ե�...")
                        Case 2
                            Set obj = frmWait.mfrmMain
                            If Not (obj Is Nothing) Then Unload frmWait
                            Call PromptStorageWarn(Val(rsCharge("�շ�����").Value), db�������, rsCharge("�շ���Ŀ").Value, rsCharge("ִ�п���").Value, rsCharge("���㵥λ").Value, int�����, 2)
                            If Not (obj Is Nothing) Then Call frmWait.OpenWait(obj, "���Ե�...")
                            Exit Function
                        End Select
                        
                    End If
                    
                End If
            
                strSQL = "Select y.�ּ�,Decode(x.�ܼ�,0,0,null,0,round(y.�ּ�/x.�ܼ�,2)) As ����,y.������Ŀid,z.�վݷ�Ŀ " & _
                            "From ( " & _
                            "Select a.�շ�ϸĿid,Sum(a.�ּ�) As �ܼ� " & _
                            "From �շѼ�Ŀ a " & _
                            "Where a.ִ������ <= SYSDATE " & _
                                "and (a.��ֹ���� IS NULL OR a.��ֹ����>SysDate) " & _
                                "and a.�շ�ϸĿid=[1] " & _
                            "Group By a.�շ�ϸĿid " & _
                            ") x, " & _
                            "�շѼ�Ŀ y, " & _
                            "������Ŀ z " & _
                            "Where Y.ִ������ <= SYSDATE " & _
                              "and (y.��ֹ���� IS NULL OR y.��ֹ����>SysDate) " & _
                              "and y.������Ŀid=z.id " & _
                              "and y.�շ�ϸĿid=[1]"
                              
                Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", Val(rsCharge("�շ�ϸĿid").Value))
                If rsPrice.BOF = False Then
                    lngCount = 0
                    dbSum = 0
                    Do While Not rsPrice.EOF
                        lngCount = lngCount + 1
                        
'                        dbӦ�ս�� = rsPrice("�ּ�").Value * rsCharge("�շ�����").Value
                        dbӦ�ս�� = rsCharge("Ӧ�ս��").Value
                        dbʵ�ս�� = rsPrice("����").Value * rsCharge("ʵ�ս��").Value
                        
                        If lngCount = rsPrice.RecordCount Then
                            dbʵ�ս�� = rsCharge("ʵ�ս��").Value - dbSum
                        Else
                            dbSum = dbSum + dbʵ�ս��
                        End If
                        
                        If zlCommFun.NVL(rsCharge("�ѱ�").Value) <> "" Then
                            strSQL = "Select Round(Round([3], 5) * ʵ�ձ��� / 100, [4]) As ʵ�ս�� From �ѱ���ϸ Where ������Ŀid = [1] And �ѱ� = [2] And (Round([3], 5) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ)"
                            Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", Val(rsPrice("������Ŀid").Value), CStr(zlCommFun.NVL(rsCharge("�ѱ�").Value)), dbʵ�ս��, ParamInfo.���ý��С��λ��)
                            If rs.BOF = False Then
                                dbʵ�ս�� = zlCommFun.NVL(rs("ʵ�ս��").Value, dbʵ�ս��)
                            End If
                            
                        End If
                        
                        If ParamInfo.���ý��С��λ�� > 0 Then
                            dbӦ�ս�� = Format(dbӦ�ս��, "0." & String(ParamInfo.���ý��С��λ��, "0"))
                            dbʵ�ս�� = Format(dbʵ�ս��, "0." & String(ParamInfo.���ý��С��λ��, "0"))
                        End If
                        
                        If zlCommFun.NVL(rsCharge("����;��").Value, 1) = 1 Then
                            
                            strSQL = "Select Nvl(Max(���),0)+1 As ��� From ���˷��ü�¼ Where No=[1] And ��¼����=[2]"
                            Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", CStr(rsCharge("No").Value), 2)
                            If rs.BOF = False Then int��� = rs("���").Value
        
                            strSQL = "zl_������ʼ�¼_Insert('" & rsCharge("No").Value & "'," & int��� & "," & _
                                                            rsCharge("����id").Value & "," & ZVal(rsCharge("�����").Value) & "," & _
                                                            "'" & rsCharge("����").Value & "','" & rsCharge("�Ա�").Value & "'," & _
                                                            "'" & rsCharge("����").Value & "','" & rsCharge("�ѱ�").Value & "'," & _
                                                            "Null,0," & _
                                                            rsCharge("��������id").Value & "," & rsCharge("��������id").Value & "," & _
                                                            rsCharge("��������id").Value & ",'" & UserInfo.���� & "'," & _
                                                            "Null," & rsCharge("�շ�ϸĿid").Value & "," & _
                                                            "'" & rsCharge("�շ����").Value & "','" & rsCharge("���㵥λ").Value & "'," & _
                                                            "1," & rsCharge("�շ�����").Value & "," & _
                                                            "0," & rsCharge("ִ�п���id").Value & "," & _
                                                            "Null," & rsPrice("������ĿID").Value & "," & _
                                                            "'" & rsPrice("�վݷ�Ŀ").Value & "'," & rsPrice("�ּ�").Value & "," & _
                                                            dbӦ�ս�� & "," & dbʵ�ս�� & "," & _
                                                            "To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            "Null,0," & _
                                                            "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                                                            ZVal(lng���id) & ",Null,'" & rsCharge("ҽ������").Value & "'," & rsCharge("ҽ��ID").Value & ",Null,Null,Null,1,0,4)"
                            Call zlDatabase.ExecuteProcedure(strSQL, "mdlMedical")
                        Else
                            strSQL = "Select Nvl(Max(���),0)+1 As ��� From ���˷��ü�¼ Where No=[1] And ��¼����=[2]"
                            Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", CStr(rsCharge("No").Value), 1)
                            If rs.BOF = False Then int��� = rs("���").Value
                            
                            strSQL = "zl_���ﻮ�ۼ�¼_Insert('" & rsCharge("No").Value & "'," & int��� & "," & _
                                                            rsCharge("����id").Value & ",Null," & ZVal(rsCharge("�����").Value) & ",Null," & _
                                                            "'" & rsCharge("����").Value & "','" & rsCharge("�Ա�").Value & "'," & _
                                                            "'" & rsCharge("����").Value & "','" & rsCharge("�ѱ�").Value & "'," & _
                                                            "Null," & _
                                                            rsCharge("��������id").Value & "," & rsCharge("��������id").Value & "," & _
                                                            rsCharge("��������id").Value & ",'" & UserInfo.���� & "'," & _
                                                            "Null," & rsCharge("�շ�ϸĿid").Value & "," & _
                                                            "'" & rsCharge("�շ����").Value & "','" & rsCharge("���㵥λ").Value & "',Null," & _
                                                            "1," & rsCharge("�շ�����").Value & "," & _
                                                            "0," & rsCharge("ִ�п���id").Value & "," & _
                                                            "Null," & rsPrice("������ĿID").Value & "," & _
                                                            "'" & rsPrice("�վݷ�Ŀ").Value & "'," & rsPrice("�ּ�").Value & "," & _
                                                            dbӦ�ս�� & "," & dbʵ�ս�� & "," & _
                                                            "To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            "'ҽ������','" & UserInfo.���� & "'," & _
                                                            ZVal(lng���id) & ",'" & rsCharge("ҽ������").Value & "'," & rsCharge("ҽ��ID").Value & ",Null,Null,Null,1,0,4)"
                            Call zlDatabase.ExecuteProcedure(strSQL, "mdlMedical")
                        End If
                        
                        rsPrice.MoveNext
                    Loop
                    
                End If
            End If
            rsCharge.MoveNext
        Loop
    End If
    
    MakeMedicalCharge = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function DataMove(ByVal strRec As String, Optional ByVal bytMode As Byte = 1) As Boolean
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    DataMove = False
    
    Select Case bytMode
    Case 1
        strSQL = "Select 1 From H���ǼǼ�¼ Where ID=[1]"
        strRec = Val(strRec)
    Case 2
        strSQL = "Select 1 From H�����Ա���� Where ID=[1]"
        strRec = Val(strRec)
    Case 3
        strSQL = "Select 1 From H���ǼǼ�¼ Where ����=[1]"
    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", strRec)
    If rs.BOF = False Then
        DataMove = True
    End If
    
errHand:

End Function

Public Function DeleteMedicalItems(ByRef strSQL() As String, ByVal rs As ADODB.Recordset, ByVal str���� As String, ByVal lng�Ǽ�id As Long, Optional ByVal lng����id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            '���ϴ������Ŀ��������ҽ��

            If lng����id > 0 Then
                strSQL(ReDimArray(strSQL)) = "ZL_�����Ŀ�嵥_DELETE(" & lng�Ǽ�id & ",NULL," & Val(rs("�嵥id").Value) & "," & lng����id & ")"
            Else
                strSQL(ReDimArray(strSQL)) = "ZL_�����Ŀ�嵥_DELETE(" & lng�Ǽ�id & ",'" & rs("���").Value & "'," & Val(rs("�嵥id").Value) & ",0)"
            End If
            
            rs.MoveNext
        Loop
    End If
    
    DeleteMedicalItems = True
    
End Function

Public Function InsertMedicalItems(ByRef strSQL() As String, ByVal rs As ADODB.Recordset, ByVal lng�Ǽ�id As Long, Optional ByVal lng����id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  �¼��������Ŀ
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngLoop As Long
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            strTmp = ""
            varRow = Split(rs("�Ʒ���ϸ").Value, ";")
            For lngLoop = 0 To UBound(varRow)
                
                varCol = Split(varRow(lngLoop), ":")
                
                If strTmp <> "" Then strTmp = strTmp & ";"
                strTmp = strTmp & varCol(5) & ":" & varCol(2) & ":" & varCol(3) & ":" & varCol(4) & ":" & Val(varCol(8)) & ":" & Val(varCol(6))
                
            Next
            
            '���������Ŀ����Ϊҽ��
            If lng����id > 0 Then
                
                strTmp = "ZL_�����Ŀ�嵥_INSERT(" & lng�Ǽ�id & "," & _
                                                    "NULL," & _
                                                    rs("ID").Value & ",'" & _
                                                    rs("�������").Value & "'," & _
                                                    Val(rs("�����۸�").Value) & "," & _
                                                    Val(rs("���۸�").Value) & "," & _
                                                    Val(rs("ִ�п���id").Value) & "," & _
                                                    IIf(rs("�ɼ���ʽid") = "", "NULL", rs("�ɼ���ʽid")) & "," & _
                                                    IIf(rs("�ɼ�����id") = "", "NULL", rs("�ɼ�����id")) & ",'" & _
                                                    zlCommFun.NVL(rs("����걾").Value) & "','" & _
                                                    rs("��鲿λ").Value & "','" & _
                                                    rs("��鲿λid").Value & "'," & lng����id & "," & IIf(rs("���㷽ʽ").Value = "����", "1", "2") & ",'" & strTmp & "')"
        
            Else
            
                strTmp = "ZL_�����Ŀ�嵥_INSERT(" & lng�Ǽ�id & ",'" & _
                                            rs("���").Value & "'," & _
                                            rs("ID").Value & ",'" & _
                                            rs("�������").Value & "'," & _
                                            Val(rs("�����۸�").Value) & "," & _
                                            Val(rs("���۸�").Value) & "," & _
                                            rs("ִ�п���id").Value & "," & _
                                            IIf(rs("�ɼ���ʽid") = "", "NULL", rs("�ɼ���ʽid")) & "," & _
                                            IIf(rs("�ɼ�����id") = "", "NULL", rs("�ɼ�����id")) & ",'" & _
                                            rs("����걾").Value & "','" & _
                                            rs("��鲿λ").Value & "','" & _
                                            rs("��鲿λid").Value & "',NULL," & IIf(rs("���㷽ʽ").Value = "����", "1", "2") & ",'" & strTmp & "')"
            End If
            
            strSQL(ReDimArray(strSQL)) = strTmp
            
            rs.MoveNext
        Loop
    End If
    
    InsertMedicalItems = True
    
End Function


Public Function DeleteItem(ByRef rsSQL As ADODB.Recordset, ByVal rs As ADODB.Recordset, ByVal str���� As String, ByVal lng�Ǽ�id As Long, Optional ByVal lng����id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            '���ϴ������Ŀ��������ҽ��

            If lng����id > 0 Then
                strTmp = "ZL_�����Ŀ�嵥_DELETE(" & lng�Ǽ�id & ",NULL," & Val(rs("�嵥id").Value) & "," & lng����id & ")"
            Else
                strTmp = "ZL_�����Ŀ�嵥_DELETE(" & lng�Ǽ�id & ",'" & rs("���").Value & "'," & Val(rs("�嵥id").Value) & ",0)"
            End If
            
            Call SQLRecordAdd(rsSQL, strTmp)
            
            rs.MoveNext
        Loop
    End If
    
    DeleteItem = True
    
End Function

Public Function NewItem(ByRef rsSQL As ADODB.Recordset, ByVal rs As ADODB.Recordset, ByVal lng�Ǽ�id As Long, Optional ByVal lng����id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  �¼��������Ŀ
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngLoop As Long
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            strTmp = ""
            varRow = Split(rs("�Ʒ���ϸ").Value, ";")
            For lngLoop = 0 To UBound(varRow)
                
                varCol = Split(varRow(lngLoop), ":")
                
                If strTmp <> "" Then strTmp = strTmp & ";"
                strTmp = strTmp & varCol(5) & ":" & varCol(2) & ":" & varCol(3) & ":" & varCol(4) & ":" & Val(varCol(8)) & ":" & Val(varCol(6))
                
            Next
            
            '���������Ŀ����Ϊҽ��
            If lng����id > 0 Then
                
                strTmp = "ZL_�����Ŀ�嵥_INSERT(" & lng�Ǽ�id & "," & _
                                                    "NULL," & _
                                                    rs("ID").Value & ",'" & _
                                                    rs("�������").Value & "'," & _
                                                    Val(rs("�����۸�").Value) & "," & _
                                                    Val(rs("���۸�").Value) & "," & _
                                                    Val(rs("ִ�п���id").Value) & "," & _
                                                    IIf(rs("�ɼ���ʽid") = "", "NULL", rs("�ɼ���ʽid")) & "," & _
                                                    IIf(rs("�ɼ�����id") = "", "NULL", rs("�ɼ�����id")) & ",'" & _
                                                    zlCommFun.NVL(rs("����걾").Value) & "','" & _
                                                    rs("��鲿λ").Value & "','" & _
                                                    rs("��鲿λid").Value & "'," & lng����id & "," & IIf(rs("���㷽ʽ").Value = "����", "1", "2") & ",'" & strTmp & "')"
        
            Else
            
                strTmp = "ZL_�����Ŀ�嵥_INSERT(" & lng�Ǽ�id & ",'" & _
                                            rs("���").Value & "'," & _
                                            rs("ID").Value & ",'" & _
                                            rs("�������").Value & "'," & _
                                            Val(rs("�����۸�").Value) & "," & _
                                            Val(rs("���۸�").Value) & "," & _
                                            rs("ִ�п���id").Value & "," & _
                                            IIf(rs("�ɼ���ʽid") = "", "NULL", rs("�ɼ���ʽid")) & "," & _
                                            IIf(rs("�ɼ�����id") = "", "NULL", rs("�ɼ�����id")) & ",'" & _
                                            rs("����걾").Value & "','" & _
                                            rs("��鲿λ").Value & "','" & _
                                            rs("��鲿λid").Value & "',NULL," & IIf(rs("���㷽ʽ").Value = "����", "1", "2") & ",'" & strTmp & "')"
            End If
            
            Call SQLRecordAdd(rsSQL, strTmp)
            
            rs.MoveNext
        Loop
    End If
    
    NewItem = True
    
End Function

Public Function OutPutQuestBill(ByVal frmMain As Object, ByVal lngKey As Long, ByVal lngPatientKey As Long, ByVal strDeptID As String, ByVal strSample As String, _
                                Optional ByVal blnVerfiy As Boolean, Optional ByVal blnCheck As Boolean, Optional ByVal bytMode As Byte = 1) As Boolean
                                
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    
    On Error GoTo errHand
    
    strSQL = "Select * From (Select Nvl(c.��������,'0') As ��������,a.ִ�п���id,Decode(a.����걾,Null,1,2) As ����걾 " & _
                "From �����Ŀ�嵥 a,�����Ŀҽ�� b,����ҽ������ c " & _
                "Where b.�嵥id=a.ID And a.�Ǽ�id=[1] And c.ҽ��id(+)=b.ҽ��id And b.����id+0=[2] And (c.�������� Is Null Or (c.�������� Is Not Null And Instr([3],''''||a.����걾||'''')>0)) Group By Decode(a.����걾,Null,1,2),a.ִ�п���id,Nvl(c.��������,'0')) Order by ����걾,ִ�п���id,��������"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, frmMain.Caption, lngKey, lngPatientKey, strSample)
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            If InStr(strDeptID, "," & zlCommFun.NVL(rs("ִ�п���id").Value, 0) & ",") > 0 Then
                If zlCommFun.NVL(rs("��������").Value, "0") <> "0" Then
    
                    If blnVerfiy Then
                    
                        '����Ƿ�����Ҫ�������ε�
                        gstrSQL = "Select  Nvl(Max(C.����˳��),1) As ���ô��� " & _
                                    "From �����Ŀ�嵥 A,������ĿĿ¼ B,�����Ŀҽ�� E,�����Ŀ���� C,����ҽ������ d " & _
                                    "Where E.�嵥ID = A.ID " & _
                                        "AND B.ID=A.������Ŀid And d.ҽ��id(+)=e.ҽ��id " & _
                                        "AND B.���='C' AND C.������Ŀid=B.ID AND C.��������=2 AND C.����˳��>1 " & _
                                        "AND A.�Ǽ�ID=[1] " & _
                                        "AND E.����ID+0=[2] " & _
                                        "AND Nvl(d.��������,'0')=[3] " & _
                                        "AND A.ִ�п���ID+0=[4] " & _
                                        "AND Instr([5],''''||A.����걾||'''')>0 "
                                        
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, frmMain.Caption, lngKey, lngPatientKey, zlCommFun.NVL(rs("��������").Value, "0"), Val(zlCommFun.NVL(rs("ִ�п���id").Value)), strSample)
                                        
                        For intCount = 1 To rsTmp("���ô���").Value
                            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_4", frmMain, "�Ǽ�id=" & lngKey, "����id=" & lngPatientKey, "��������=" & zlCommFun.NVL(rs("��������").Value, "0"), "ִ�п���id=" & zlCommFun.NVL(rs("ִ�п���id").Value, 0), "����걾=" & strSample, bytMode)
                        Next
                        
                    End If
                    
                ElseIf blnCheck Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_5", frmMain, "�Ǽ�id=" & lngKey, "����id=" & lngPatientKey, "ִ�п���id=" & zlCommFun.NVL(rs("ִ�п���id").Value, 0), bytMode)
                End If
            End If
            rs.MoveNext
        Loop
    End If
            
    OutPutQuestBill = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function WriteFile(ByVal strFile As String, ByVal strText As String) As Boolean

    '******************************************************************************************************************
    '���ܣ�д��Ϣ��ָ���ļ�
    '�������ļ���
    '���أ���Ϣ����
    '******************************************************************************************************************
    
    Dim fso As New FileSystemObject
    Dim objTxt As TextStream
    
    On Error GoTo errHand
    
    Set objTxt = fso.OpenTextFile(strFile, ForAppending, True)
    objTxt.WriteLine strText
    
errHand:
    
End Function

