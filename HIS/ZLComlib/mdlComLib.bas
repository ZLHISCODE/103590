Attribute VB_Name = "mdlComLib"
Option Explicit
'**************************
'       OEM����
'
'ҽҵ  D2BDD2B5
'����  CDD0C6D5
'����  D6D0C8ED
'����  B4B4D6C7
'��̩ BDF0BFB5CCA9
'����  B1A6D0C5
'**************************

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gcnOracleOLEDB  As ADODB.Connection  '�������ݿ�����OLEDB��ʽ������ȡLOB����ʱһ�ζ�ȡ
Public gobjComLib As clsComLib
Public gobjRegister As Object               'ע����Ȩ����zlRegister

Public g_AutoConnect    As Boolean          'ͨ���ñ�������ͬʵ����gblnAutoConnect��ֵ����
Public g_NodeNo As String                   'ͨ���ñ�������ͬʵ����gstrNodeNo��ֵ����
Public g_NodeName As String                 'ͨ���ñ�������ͬʵ����gstrNodeName��ֵ����
Public glngSessionID As Long
Public gstrComputerName As String
Public gstrSysName As String                'ϵͳ����
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrAppsoft As String                'APPSOFT·��
Public gstrHelpPath As String
Public gblnOK As Boolean
Public gstrDBUser As String
Public gfrmMain As Object '����̨����
Public gblnShow As Boolean

Public gobjLogFile As FileSystemObject
Public gobjLogText As TextStream
Public gobjPlanExFile As FileSystemObject
Public gobjPlanExText As TextStream

Public gblnSQLTest As Boolean
Public gblnSQLLog As Boolean
Public gblnSQLPlan As Boolean   '���ܼ��ģʽ

Public gstrSysUser As String
Public gcnSysConn As ADODB.Connection 'sys����
Public gcnSysOLEDB As ADODB.Connection 'sys����,OLEDB��ʽ
Public gblnSys As Boolean
Public gstrRecentSQL As String  '���ִ�е�SQL���

Public grsDiagConn As ADODB.Recordset '������뵥��Ϲ���

'ϵͳ����
Public gblnRunLog As Boolean '�Ƿ��¼ʹ����־
Public gblnErrLog As Boolean '�Ƿ��¼���д���

Public grsParas As ADODB.Recordset 'ϵͳ��������
Public grsUserParas As ADODB.Recordset 'ϵͳ��������
Public grsDeptParas As ADODB.Recordset    'ϵͳ�������Ż���
Public grsUserInfo As ADODB.Recordset  '��ǰ�û�����Ա�Ͳ�����Ϣ����
Public gcolMoveDate As Collection    '��ʷ���ݵ�ת������

Public gclsMipClient As clsMipClient
Public gcllComlibs  As Collection       '���е�Comlib����ʵ������

Public gcolWriteLog As Collection '�洢�������͵���־�Ķ���
Public gstrLastLogName As String              '������һ��ʹ�õ���־����
Public gobjLastLog As TextStream          '������һ��ʹ�õ���־����
Public gstrLastLogInfoHeader As String        '������һ��ʹ�õ���־ͷ
Public gcolLastLogInfoHeader As Collection    '����������͵���־��һ�δ洢����־ͷ

Public glngPatiTypeWinProc As Long               'ԭʼ��Ϣ���
Public gclsPDF          As clsPDF       'PDF�����ȫ�ֻ��棬�Ա�ͬһ�����̹���һ��ʵ��

Public Const MSTR_DBLINK_KEY As String = "zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325"  '���ܽ�����Կ
'���ӷ�ʽ
Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
    OriginalConnection = 9
End Enum

Public Function FlexScroll(ByVal hWnd As Long, ByVal wMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long
'֧�ֹ��ֵĹ���
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '���¹�
            gobjComLib.zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '���Ϲ�
            gobjComLib.zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPatiTypeWinProc, hWnd, wMsg, wParam, lParam)
End Function

Public Sub ChangeAllIntanceConn(cnMain As ADODB.Connection)
'���ܣ�ͬ��clsComlib����ʵ���е�mcnoracle�����������԰�ť�Ľ��á�
'������ͬ������ʵ��������
    Dim objComlib As clsComLib
    If Not gcllComlibs Is Nothing Then
        For Each objComlib In gcllComlibs
            Call objComlib.ChangeIntanceConn(cnMain)
        Next
    End If
End Sub

Public Function SQLObject(ByVal strSQL As String) As String
'���ܣ�����SQL������õ��Ķ�����
'������strSQL=Ҫ������ԭʼSQL���
'���أ�SQL��������ʵ��Ķ�����,��"���ű�,���˷��ü�¼,ZLHIS.��Ա��"
'˵����1.��Oracle SELECT������
'      2.���SQL����еĶ�����ǰ����������ǰ׺,���ǰ׺���ᱻ��ȡ
'      3.��Ҫ����TrimChar;TrueObject��֧��
    Dim intB As Integer, intE As Integer, intL As Integer, intR As Integer
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Integer, j As Integer
    
    On Error GoTo errh
    
    '��д����ȥ��������ַ�
    strAnal = UCase(TrimChar(strSQL))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
    '�ȷֽ⴦��Ƕ���Ӳ�ѯ
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB 'ƥ�����������λ��
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                If intE - intB - 1 <= 0 Then
                    '���ڷ��Ӳ�ѯ,�����Ż�����������,��ʹѭ������
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '�Ӳ�ѯ���
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '�����Ӳ�ѯ������ΪΪ���������
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "Ƕ�ײ�ѯ")
                    '�ݹ����
                    strObject = strObject & "," & SQLObject(strSub)
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '��ƥ��������
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '�ֽ����
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '�ӵ�һ��From���沿�ݿ�ʼ
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject, "," & strTrue) = 0 And strTrue <> "Ƕ�ײ�ѯ" Then
                strObject = strObject & "," & strTrue
            End If
        Next
    Next
    '���
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errh:
    Err.Clear
End Function

Private Function TrimChar(Str As String) As String
'����:ȥ���ַ����������Ŀո�ͻس�(����ͷ�Ŀո�,�س�),��ȥ��TAB�ַ�,������������
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(Str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(Str)
    i = InStr(strTmp, "  ")
    Do While i > 0
        strTmp = Left(strTmp, i) & Mid(strTmp, i + 2)
        i = InStr(strTmp, "  ")
    Loop
    
    i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Do While i > 0
        strTmp = Left(strTmp, i + 1) & Mid(strTmp, i + 4)
        i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Loop
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Private Function TrueObject(ByVal strObject As String) As String
'���ܣ�SQLObject�������Ӻ���,����ȥ���������е������ַ�
    Dim i As Integer
    'Ѱ�ҵ�һ�������ַ�λ��
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    'Ѱ�Һ����һ���������ַ�
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function

Public Function SeekCboIndex(objCbo As Object, varData As Variant) As Long
'���ܣ���ItemData��Text����ComboBox������ֵ
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '�Ⱦ�ȷ����
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf gobjComLib.zlCommFun.GetNeedName(objCbo.List(i)) = varData Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '��ģ������
            For i = 0 To objCbo.ListCount - 1
                If InStr(objCbo.List(i), varData) > 0 Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    Else
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = varData Then
                SeekCboIndex = i: Exit Function
            End If
        Next
    End If
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function

'--------------------------------------------------
'���ܣ�����Ƿ�Ϊ����Ͽ���ADO�Ͽ������Ĵ���!
'���أ�True:�ָ����ӳɹ� False�ָ�����ʧ��
'--------------------------------------------------
Public Function CheckAdoConnction(ByRef blnStatus As Boolean) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnAdoErr As Boolean
    Dim strError As String
    On Error GoTo Errhand
    blnAdoErr = False
    blnStatus = False

    On Error GoTo Errhand
    Err = 0
    DoEvents
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    gcnOracle.Open
    If blnAdoErr Then
        'True '��ORA-12560������ORACLE��������
        CheckAdoConnction = True
    Else
        'False '������������
        CheckAdoConnction = False
        On Error Resume Next
        '�������жϿͻ����Ƿ񱻽�ֹʹ�ã�������ֹ�����Զ��Ͽ�����
        strSQL = "Select NVL(��ֹʹ��,0)  ��ֹʹ�� From zlClients Where ����վ=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckAdoConnction", gstrComputerName)
        If Err.Number <> 0 Then Err.Clear
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                If rsTmp!��ֹʹ�� = 1 Then
                    If gcnOracle.State = adStateOpen Then gcnOracle.Close
                    CheckAdoConnction = True
                    Call SaveSetting("ZLSOFT", "����ȫ��\��������Զ�����", "AutoConnect", 0)
                    MsgBox "��ǰ����վ�Ѿ�������Ա���ã�����ϵ����Ա������ò����µ�¼��", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    Exit Function
Errhand:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        If InStr(Err.Description, "ORA-12560") > 0 Then
            blnAdoErr = True
            Resume Next
        ElseIf InStr(Err.Description, "ORA-12543") > 0 Then
            blnAdoErr = True
            Resume Next
        Else
            '����������������������
            CheckAdoConnction = True
            blnStatus = True
        End If
    Else
        CheckAdoConnction = False
    End If
End Function

Public Function CheckErrConnectInfo(ByVal strErrNum As String, ByVal strNote As String, ByVal strErrInfo As String, ByVal intType As Integer) As Boolean
    '------------------------------------------------
    '���ܣ� ��������IntType(1,2)���vb��oralce���صľ��������Ϣ�����ж��Ƿ�Ϊ����Ͽ������Ĵ�������������Ĵ�������
    '������ strNote������Ϣ,strErrInfo������ϸ��Ϣ,intType �������� 1��VB���� 2:ORACLE����
    '���أ� True:���������Ĵ��� False:��������
    '------------------------------------------------
    Dim strTemp As String
    Dim i As Integer
    If intType = 1 Then
        'VB�������
   
        If InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02396") > 0 Then '����������ʱ��, ���������� IDLE_TIME profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02399") > 0 Then '�����������ʱ��, ������ע�� connect_time profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-01012") > 0 Then 'û�е�¼
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-00028") > 0 Then '�Ự����ֹ
            CheckErrConnectInfo = True
        Else
            If strErrNum = "3709" Then '3709�����������޷�����ִ�д˲������ڴ����������������ѱ��رջ���Ч����������
                CheckErrConnectInfo = True
            Else
                If strNote = "��ȷ���Ĵ���" Then
                    CheckErrConnectInfo = True
                Else
                    CheckErrConnectInfo = False
                End If
            End If
        End If
    Else
        'ORACLE�������
        If InStr(strErrInfo, "SQLSetConnectAttr") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12543") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02396") > 0 Then '����������ʱ��, ���������� IDLE_TIME profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02399") > 0 Then '�����������ʱ��, ������ע�� connect_time profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-01012") > 0 Then 'û�е�¼
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-00028") > 0 Then '�Ự����ֹ
            CheckErrConnectInfo = True
        Else
            CheckErrConnectInfo = False
        End If
    End If
End Function

Public Function GetGUID() As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim udtGUID As GUID
    
    On Error GoTo Errhand
    
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
                String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
                String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
                IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
                IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
                IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
                IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
                IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
                IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
                IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
                IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
    
    Exit Function
Errhand:
    'MsgBox Err.Description
End Function

Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, ByVal dwData As Long) As Long
     Dim monitorInf As MONITORINFO
     Dim R As RECT
     
     ReDim Preserve gMonitors(UBound(gMonitors) + 1)
     


     'initialize   the   MONITORINFO   structure
     monitorInf.cbSize = Len(monitorInf)
     'Get   the   monitor   information   of   the   specified   monitor
     GetMonitorInfo hMonitor, monitorInf
     'write   some   information   on   teh   debug   window

    
     gMonitors(UBound(gMonitors) - 1).monitorHandle = hMonitor
     gMonitors(UBound(gMonitors) - 1).monitorInf = monitorInf
     
     '������뷵��1���Ա����ִ��
     MonitorEnumProc = 1
  End Function

Public Function GetMonitorIndex(ByVal windowHandle As Long) As Long
'    '******************************************************************************************************************
'    '���ܣ���ü�����ID
'    '������windowHandle
'    '���أ�������ID
'    '******************************************************************************************************************

    Dim i As Integer

    Dim monitorCount As Integer
    monitorCount = 0

    On Error GoTo GetMonitorInf
      monitorCount = UBound(gMonitors)
GetMonitorInf:
      If monitorCount <= 1 Then
        ReDim Preserve gMonitors(1)
        gMonitors(1).monitorHandle = -1

        EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
      End If


    For i = 1 To UBound(gMonitors)
      If MonitorFromWindow(windowHandle, MONITOR_DEFAULTTONEAREST) = gMonitors(i).monitorHandle Then
        GetMonitorIndex = i - 1
        Exit Function
      End If
    Next i

    GetMonitorIndex = -1

End Function

'���ܺ���
Public Function Decipher(ByVal password As String, ByVal from_text As String) As String
    '����
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    password = Base64Encode(password) & "WIZARDPAGE"
    
    Dim offset As Long
    Dim str_len As Integer
    Dim i As Integer
    Dim ch As Integer
    offset = NumericPassword(password)
    Rnd -1
    Randomize offset

    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            Decipher = Decipher & Chr$(ch)
        End If
    Next i
End Function


'�ӽ����ַ�������,��֧������
Private Function Base64Encode(InStr1 As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim mInByte(3)     As Byte, mOutByte(4)       As Byte
    Dim myByte     As Byte
    Dim i     As Integer, LenArray       As Integer, j       As Integer
    Dim myBArray()     As Byte
    Dim OutStr1     As String
    myBArray() = StrConv(InStr1, vbFromUnicode)
    LenArray = UBound(myBArray) + 1
    For i = 0 To LenArray Step 3
      If LenArray - i = 0 Then
        Exit For
      End If
      If LenArray - i = 2 Then
        mInByte(0) = myBArray(i)
        mInByte(1) = myBArray(i + 1)
        Base64EncodeByte mInByte, mOutByte, 2
      ElseIf LenArray - i = 1 Then
        mInByte(0) = myBArray(i)
        Base64EncodeByte mInByte, mOutByte, 1
      Else
        mInByte(0) = myBArray(i)
        mInByte(1) = myBArray(i + 1)
        mInByte(2) = myBArray(i + 2)
        Base64EncodeByte mInByte, mOutByte, 3
      End If
      For j = 0 To 3
        OutStr1 = OutStr1 & Chr(mOutByte(j))
      Next j
    Next i
    Base64Encode = OutStr1
    
End Function

Private Sub Base64EncodeByte(mInByte() As Byte, mOutByte() As Byte, Num As Integer)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim tByte     As Byte
    Dim i     As Integer
    If Num = 1 Then
      mInByte(1) = 0
      mInByte(2) = 0
    ElseIf Num = 2 Then
      mInByte(2) = 0
    End If
    tByte = mInByte(0) And &HFC
    mOutByte(0) = tByte / 4
    tByte = ((mInByte(0) And &H3) * 16) + (mInByte(1) And &HF0) / 16
    mOutByte(1) = tByte
    tByte = ((mInByte(1) And &HF) * 4) + ((mInByte(2) And &HC0) / 64)
    mOutByte(2) = tByte
    tByte = (mInByte(2) And &H3F)
    mOutByte(3) = tByte
    For i = 0 To 3
      If mOutByte(i) >= 0 And mOutByte(i) <= 25 Then
        mOutByte(i) = mOutByte(i) + Asc("A")
      ElseIf mOutByte(i) >= 26 And mOutByte(i) <= 51 Then
        mOutByte(i) = mOutByte(i) - 26 + Asc("a")
      ElseIf mOutByte(i) >= 52 And mOutByte(i) <= 61 Then
        mOutByte(i) = mOutByte(i) - 52 + Asc("0")
      ElseIf mOutByte(i) = 62 Then
        mOutByte(i) = Asc("+")
      Else
        mOutByte(i) = Asc("/")
      End If
    Next i
    If Num = 1 Then
      mOutByte(2) = Asc("=")
      mOutByte(3) = Asc("=")
    ElseIf Num = 2 Then
      mOutByte(3) = Asc("=")
    End If
End Sub

Private Function NumericPassword(ByVal password As String) As Long
    Dim value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    str_len = Len(password)
    For i = 1 To str_len
        ch = Asc(Mid$(password, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = value
End Function

Public Function IsOLEDBConnection(ByVal cnMain As ADODB.Connection) As Boolean
'���ܣ��жϵ�ǰ�����Ƿ���OraOLEDB����
'����Provider���жϣ��������ַ�ʽ
'��ʽһ��'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
'��ʽ����
'.Provider = "OraOLEDB.Oracle"
'.Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUserName, strPassWord
'�����ַ�ʽ�����Զ�����.Provider����
    'ʹ��Like����Ϊ���ܺ������Ӱ汾��OraOLEDB.Oracle.1
    If UCase(cnMain.Provider) Like "ORAOLEDB.ORACLE*" Then
        IsOLEDBConnection = True
    End If
End Function

