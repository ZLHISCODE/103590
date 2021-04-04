Attribute VB_Name = "mdlStuff"
Option Explicit

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long
Public gstrAviPath As String
Public gstrVersion As String
Public gstrMatchMethod As String
Public gbytSimpleCodeTrans As Byte          '��Ƭ�����Ƿ���������л�����

Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrUnitName As String '�û���λ����
Public gfrmMain As Object

Public gstrSQL As String
Public gblnOK As Boolean
Public gstrIme As String

Public gobjSquareCard As Object             'һ��ͨ�ӿ�
Public gstrCardType As String           '���п���𣬸�ʽ������|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������);��
Public gintCardCount As Integer  '������
Public gblnIncomeItem As Boolean            '��¼����Ŀ¼�������Ƿ�������������Ŀ

Public gstrPriceClass As String         '�۸�ȼ�
Public gobjPlugIn As Object             '��ҽӿ�

Public Const glngRowByFocus = &HFFE3C8
Public Const glngRowByNotFocus = &HF4F4EA
Public Const glngFixedForeColorByFocus = &HFF0000
Public Const glngFixedForeColorNotFocus = &H80000012

'ҩƷ���۸�������󾫶�
Public Type Type_Digits
    Digit_��� As Integer
    Digit_�ɱ��� As Integer
    Digit_���ۼ� As Integer
    Digit_���� As Integer
End Type
Public gtype_UserDrugDigits As Type_Digits

'���ѿ���ʽ
Public Enum gCardFormat
    ���� = 0
    ȫ�� = 1
    ˢ����־ = 2
    �����ID = 3
    ���ų��� = 4
    ȱʡ��־ = 5
    �Ƿ�����ʻ� = 6
    �������� = 7
End Enum

Public Type TYPE_USER_INFO
    Id As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public gOraFmt_Max As g_FmtString


Public UserInfo As TYPE_USER_INFO
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
'��ʼ���ڵı�־
Public Enum StartDayFlag
    FirstDayOfWeek = 0
    FirstDayOfMonth = 1
    FirstDayOfQuarter = 2
    FirstDayOfHalfYear = 3
    FirstDayOfyear = 4
End Enum
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Sub zlPlugIn_Ini(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object)
    '�����չ�ӿڳ�ʼ��
    If objPlugIn Is Nothing Then
        On Error Resume Next
        Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not objPlugIn Is Nothing Then
            Call objPlugIn.Initialize(gcnOracle, lngSys, lngModul)
            If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub

Public Sub zlPlugIn_Unload(objPlugIn As Object)
    'ж����ҽӿ�
    Set objPlugIn = Nothing
End Sub
'ȡҩƷ��λ����
Public Function GetDrugUnit(ByVal lng�ⷿID As Long, ByVal frmCaption As String, Optional ByVal bln���� As Boolean = True) As String
    Dim rsProperty As New Recordset
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim intUnit As Integer, strUnit As String
    Dim blnȱʡ As Boolean
    Dim lngModul As Long
    
    On Error GoTo ErrHand
    
    If frmCaption Like "ҩƷ�������*" Then
        lngModul = 1343
    ElseIf frmCaption Like "Э��ҩƷ���*" Then
        lngModul = 1344
    ElseIf frmCaption Like "ҩƷ�ƿ����*" Then
        lngModul = 1304
    End If
    
    intUnit = 0
    '��������쵥����ֱ�ӷ���ע����еĵ�λ
    If lngModul = 1343 Or lngModul = 1304 Or lngModul = 1344 Then
        intUnit = Val(zlDataBase.GetPara("ҩƷ��λ", glngSys, lngModul))
        '���ز������õĵ�λ˳�����£�0-ȱʡ;1-ҩ��;2-����;3-סԺ;4-�ۼۣ���Ҫת��Ϊ��ϵͳ������һ��
        If intUnit = 1 Then
            intUnit = 4
        ElseIf intUnit = 4 Then
            intUnit = 1
        End If
        strUnit = intUnit
    End If
    
    If intUnit = 0 Then
        gstrSQL = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
        Set rsProperty = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��λ", lng�ⷿID)
        
        'ȡ������󼰲�������
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        
        If InStr(strWorkTemp, "ҩ��") <> 0 Then
            'ҩ�ⵥλ
            intUnit = 1
            strUnit = 4
        ElseIf InStr(strobjTemp, "1") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '���ﵥλ
            intUnit = 2
            strUnit = 2
        ElseIf InStr(strobjTemp, "2") <> 0 Then
            'סԺ��λ
            intUnit = 3
            strUnit = 3
        Else
            '�ۼ۵�λ����Ҫ���Ƽ���
            intUnit = 4
            strUnit = 1
        End If
        
        'ȡ��ҩ��ȱʡ��ʹ�õĵ�λ
        GetDrugUnit = GetSpecUnit(lng�ⷿID, intUnit)
    Else
        GetDrugUnit = Switch(strUnit = 1, "�ۼ۵�λ", strUnit = 2, "���ﵥλ", strUnit = 3, "סԺ��λ", strUnit = 4, "ҩ�ⵥλ")
    End If
    
    'ת��Ϊ��ʵ�ĵ�λ���ظ�������
    
    If glngSys / 100 = 8 Then
        'ҩ��ֻ���ۼ۵�λ��ҩ�ⵥλ
        GetDrugUnit = IIf(strUnit = 1, "�ۼ۵�λ", "ҩ�ⵥλ")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDrugUnit = "�ۼ۵�λ"
End Function
Public Function GetSpecUnit(ByVal lng�ⷿID As Long, ByVal int��Χ As Integer) As String
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    Dim strsql As String
    
    '����ָ���ָⷿ�����÷�Χ�ĵ�λ
    On Error GoTo ErrHand
    
    gstrSQL = "Select Nvl(����,1) AS ��λ From ҩƷ�ⷿ��λ Where �ⷿID=[1] And ���÷�Χ=[2] "
    Set rsProperty = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ��λ", lng�ⷿID, int��Χ)
   
    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!��λ
    Else
        gstrSQL = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
        Set rsProperty = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��λ", lng�ⷿID)
    
        'ȡ������󼰲�������
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            'סԺ��λ
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '���ﵥλ
            strUnit = 2
        ElseIf InStr(strWorkTemp, "ҩ��") <> 0 Then
            'ҩ�ⵥλ
            strUnit = 4
        Else
            '�ۼ۵�λ����Ҫ���Ƽ���
            strUnit = 1
        End If
    End If
    
    'ת��Ϊ��ʵ�ĵ�λ���ظ�������
    GetSpecUnit = Switch(strUnit = 1, "�ۼ۵�λ", strUnit = 2, "���ﵥλ", strUnit = 3, "סԺ��λ", strUnit = 4, "ҩ�ⵥλ")
    If glngSys / 100 = 8 Then
        'ҩ��ֻ���ۼ۵�λ��ҩ�ⵥλ
        GetSpecUnit = IIf(strUnit = 1, "�ۼ۵�λ", "ҩ�ⵥλ")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function Get��������(ByVal lng�ⷿID As Long, ByVal lng����ID As Long) As Integer
    '����ָ���ⷿ��ָ�����ϵķ�������
    '���أ�0-��������1-����
    Dim rsCheck As New ADODB.Recordset
    Dim int���� As Integer
    Dim bln���ϲ��� As Boolean
    Dim strsql As String
        
    On Error GoTo errHandle
    
    '�ж��Ƿ��Ƿ��ϲ���
    strsql = "select ����ID from ��������˵�� where �������� =  '���ϲ���' And ����id=[1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(strsql, "Get��������", lng�ⷿID)

    bln���ϲ��� = (Not rsCheck.EOF)
        
    '�ж϶�Ӧ��ҩƷĿ¼�еķ�������
    strsql = " Select Nvl(�ⷿ����,0) As �ⷿ����,nvl(���÷���,0) As ���÷��� " & _
              " From �������� Where ����ID=[1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(strsql, "Get��������", lng����ID)
              
    If bln���ϲ��� Then
        int���� = rsCheck!���÷���
    Else
        int���� = rsCheck!�ⷿ����
    End If
    
    Get�������� = int����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function AviShow(FrmMain As Form, Optional ByVal blnShow As Boolean = True)
    '����Flash����
    DoEvents
    
    If blnShow Then
        FS.ShowFlash "���ڲ�������,���Ժ�...", FrmMain
    Else
        FS.StopFlash
    End If
    
    DoEvents
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
Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = sys.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.Id = rsTmp!Id
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�û��� = UserInfo.����
        gstrUserName = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        GetUserInfo = True
    End If
    Exit Function
errH:
    Call ErrCenter
    Call SaveErrLog
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '�������������ID������
    '����������ɹ����� �¼�������; ���߷��� 0
    Dim strsql As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strsql = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ�ID is null " & strWhere & " connect by prior id=�ϼ�id"
    Else
        strsql = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with ID=" & strID & strWhere & " connect by prior id=�ϼ�id"
    End If
    
    Call zlDataBase.OpenRecordset(rsTemp, strsql, "��ȡָ����ı����������󳤶�")
    
    If rsTemp.EOF Then
        GetDownCodeLength = 0
    Else
        GetDownCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetDownCodeLength = 0
End Function

Public Function GetLocalCodeLength(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '����������ϼ�ID������
    '����������ɹ����� ������; ���߷��� 0
    Dim strsql As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strsql = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID is null" & strWhere
    Else
        strsql = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
    End If
    
    Call zlDataBase.OpenRecordset(rsTemp, strsql, "mdlCureBase")
    
    If rsTemp.EOF Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Public Function GetParentCode(ByVal str�ϼ�ID As String, ByVal strTableName As String) As String
    '������������ȡ�ϼ�����
    '����������ϼ�ID,����
    '����������ɹ����� �ϼ�����; ���߷��� ��
    Dim strsql As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        GetParentCode = ""
        Exit Function
    End If
    
    strsql = "select ���� from " & strTableName & " where ID=[1]"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(strsql, "��ȡ�ϼ�����", str�ϼ�ID)
    
    If rsTemp.EOF Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("����").Value
    End If
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetMaxLocalCode(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '��������������ָ������ϼ�ID ��ȡ������������
    '����������ϼ�ID,����
    '����������ɹ����� ������; ���߷��� ��
    Dim strsql As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strCode As String, strAllCode As String
    Dim intLength   As Integer
    err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strsql = "select max(to_number(����))+1 as MaxCode from " & strTableName & " where �ϼ�ID is null" & strWhere
    Else
        strsql = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
    End If
    intCode = GetLocalCodeLength(str�ϼ�ID, strTableName, strWhere)
    
    Call zlDataBase.OpenRecordset(rsTemp, strsql, "����ָ������ϼ�ID ��ȡ������������")
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Sub SetGridFocus(ByVal objGrid As VSFlexGrid, ByVal blnGetFoucs As Boolean)
    With objGrid
        If blnGetFoucs Then
            .GridColorFixed = &H80000008
            .GridColor = &H80000008
            .ForeColorFixed = glngFixedForeColorByFocus
            .BackColorSel = glngRowByFocus
        Else
            .GridColorFixed = &H80000011
            .GridColor = &H80000011
            .ForeColorFixed = glngFixedForeColorNotFocus
            .BackColorSel = glngRowByNotFocus
        End If
    End With
End Sub
 
Public Function GetFormat(ByVal dblInput As Double, ByVal intDotBit As Integer) As String
    GetFormat = zlStr.FormatEx(dblInput, intDotBit, , True)
End Function

Public Function BinTOHex(sString As String) As String
    Dim lngLoop As Integer, lngTemp As Long, lngJLoop As Integer, lngTmp As Long
    lngTemp = 0
    For lngLoop = 1 To Len(sString)
        If Mid(sString, lngLoop, 1) = "1" Then
            lngTmp = 1
            For lngJLoop = 0 To lngLoop - 2
                lngTmp = lngTmp * 2
            Next
        Else
            lngTmp = 0
        End If
        lngTemp = lngTemp + lngTmp
    Next
    BinTOHex = CStr(lngTemp)
End Function

Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Sub zlChangeCode(ByVal strTableName As String, _
    ByVal lng�ϼ�id As Long, _
    ByVal txtUpCode As TextBox, _
    ByVal txtCode As TextBox, _
    Optional ByVal chkChangeCode As CheckBox = Nothing, _
    Optional ByVal strCaption As String = "")
    '------------------------------------------------------------------------------------
    '���ܣ�����ѡ����ϼ�ȷ����ǰ�ı��룬�����ϼ�����������ʾ����
    '������strTableName-���ڷ���ı���
    '      lng�ϼ�ID-ѡ����ϼ�
    '      TxtUpCode-��ʾ���ϼ��ı���
    '      TxtUpCode-��ʾ�ı����ı���
    '      chkChangeCode-�����Ƿ�ı�ԭ�����ݿ��е���ʷ����ѡ��ؼ�
    '      strCaption-���ô����Capiton
    'ע�⣺���б�����ID,�ϼ�id,����
    '------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intMaxCodeLen As Integer  'ȷ�������ʵ�ʳ���
    err = 0: On Error GoTo ErrHand
    
   chkChangeCode.Value = 0
   chkChangeCode.Enabled = True
   
    If lng�ϼ�id = 0 Then
        txtUpCode.Text = ""
        gstrSQL = "select max(����) as ���� From " & strTableName & " Where �ϼ�ID is null "
        zlDataBase.OpenRecordset rsTemp, gstrSQL, strCaption
            
        With rsTemp
            intMaxCodeLen = .Fields("����").DefinedSize
            If IsNull(!����) Then
                txtCode.Text = "01"
                txtCode.MaxLength = intMaxCodeLen
                txtCode.Tag = txtCode.MaxLength
                chkChangeCode.Value = 1
                chkChangeCode.Enabled = False
            Else
                txtCode.MaxLength = Len(Trim(!����))
                txtCode.Tag = txtCode.MaxLength
                If !���� = String(txtCode.MaxLength, "9") Then
                    If txtCode.MaxLength >= intMaxCodeLen Then
                        ShowMsgBox "������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������"
                        txtCode.Text = Space(txtCode.MaxLength)
                       chkChangeCode.Value = 0
                       chkChangeCode.Enabled = False
                    Else
                        ShowMsgBox "�������Ѿ��ﵽ�������ƣ������������볤����������Ҫ"
                        txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                        txtCode.MaxLength = txtCode.MaxLength + 1
                        txtCode.Tag = txtCode.MaxLength
                       chkChangeCode.Value = 1
                    End If
                Else
                    txtCode.Text = Format(Mid(!����, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
                End If
            End If
        End With
        Exit Sub
   End If
   'ȷ���ϼ�����
   
    gstrSQL = "Select ���� From " & strTableName & " where id=[1]"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, strCaption, lng�ϼ�id)
    
    If Not rsTemp.EOF Then
        txtUpCode.Text = zlStr.NVL(rsTemp!����)
    End If
    
    '��ȷ���Ƿ����¼�
    gstrSQL = "select nvl(max(����),'') as ����  From " & strTableName & " Where  �ϼ�ID =[1] "
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, strCaption, lng�ϼ�id)
    
    intMaxCodeLen = rsTemp.Fields("����").DefinedSize

    If zlStr.NVL(rsTemp!����) = "" Then
        '�������¼�
        '�����ϼ�IDȡ�ϼ�����
'        gstrSQL = "Select ���� From " & strTableName & " where id=" & lng�ϼ�id
'        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
'        txtUpCode.Text = zlStr.Nvl(rsTemp!����)
        txtCode.MaxLength = intMaxCodeLen - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If txtCode.MaxLength > 1 Then
            txtCode.Text = "01"
        Else
            txtCode.Text = "1"
        End If
        chkChangeCode.Value = 1
        chkChangeCode.Enabled = False
        Exit Sub
    End If
    
    With rsTemp
        txtCode.MaxLength = Len(!����) - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If Mid(!����, Len(txtUpCode.Text) + 1) = String(txtCode.MaxLength, "9") Then
            If Len(txtUpCode.Text) + txtCode.MaxLength >= intMaxCodeLen Then
                ShowMsgBox "�÷����¼�������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������"
                txtCode.Text = Space(txtCode.MaxLength)
               chkChangeCode.Value = 0
               chkChangeCode.Enabled = False
            Else
                ShowMsgBox "�÷����¼��������Ѿ��ﵽ�������ƣ������������볤����������Ҫ"
                txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                txtCode.MaxLength = txtCode.MaxLength + 1
                txtCode.Tag = txtCode.MaxLength
               chkChangeCode.Value = 1
            End If
        Else
            txtCode.Text = Format(Mid(!����, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ImeLanguage(ByVal blnOpen As Boolean)
    '-----------------------------------------------------------------------------------
    '����: ��/�ر����뷨
    '����: blnOpen-�Ǵ򿪻��ǹر�(trueΪ��,falseΪ�ر�)
    '���أ�
    '-----------------------------------------------------------------------------------
    If blnOpen Then
        OS.OpenIme (True)
    Else
        OS.OpenIme False
    End If
End Sub

Public Function DepotProperty(ByVal lng��Աid As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo errHandle
    '����ָ����Ա�Ƿ����ҩ������
    gstrSQL = "Select Distinct �������� From ������Ա B,��������˵�� A " & _
             " Where A.�������� = '���Ŀ�' And " & _
             " A.����id = B.����id And B.��Աid = [1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��������", lng��Աid)
    If rsCheck.RecordCount <> 0 Then
        DepotProperty = True
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowCostPrice() As Boolean
    Dim blnCostPrice As Boolean
    
    On Error GoTo errHandle
    '�Ƿ�������ҩ����Ա�鿴���ݵĳɱ���
    blnCostPrice = Val(zlDataBase.GetPara(190, 100, , 0))
    
    'ҩ����Ա���ܣ�ֻ��ҩ����Ա���Բ�������Ϊ׼
    If DepotProperty(UserInfo.Id) Then
        ShowCostPrice = True
    Else
        ShowCostPrice = blnCostPrice
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'�����룬���ƣ���������ĳһ��
Public Function FindRownew(ByVal mshBill As BillEdit, ByVal int�Ƚ��� As Integer, _
    ByVal str�Ƚ�ֵ As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo errHandle
    FindRownew = True
    With mshBill
        If .Rows = 2 Then Exit Function
        If str�Ƚ�ֵ = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                If InStr(1, UCase(strCode), UCase(str�Ƚ�ֵ)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int�Ƚ���
                    .MsfObj.TopRow = .Row
                    .SetRowColor CLng(.Row), &HFFCECE, True
                    Exit Function
                End If
            End If
        Next
        gstrSQL = "" & _
        " SELECT DISTINCT b.���� " & _
        " FROM (SELECT DISTINCT A.�շ�ϸĿid " & _
        "       FROM �շ���Ŀ���� A" & _
        "       Where A.���� LIKE upper([1]) " & _
        "      ) A, �շ���ĿĿ¼ B " & _
        " Where a.�շ�ϸĿid = b.ID And (b.վ��=[2] or b.վ�� is null) "
        
        Set rsCode = zlDataBase.OpenSQLRecord(gstrSQL, "����ָ����������", GetMatchingSting(str�Ƚ�ֵ, False), gstrNodeNo)
        If rsCode.EOF Then
            FindRownew = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!����)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int�Ƚ���
                        .MsfObj.TopRow = .Row
                        .SetRowColor CLng(.Row), &HFFCECE, True
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            End If
        Next
        rsCode.Close
    End With
    FindRownew = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    err = 0: On Error GoTo ErrHand:
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
    err = 0
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

Public Function �ж�ֻ�߱����ϲ���(ByVal lng����ID As Long) As Boolean
    '�ж�ֻ�߱����ϱ����ʵ�:����ȡ���Ŀ���Ƽ������Ƶ����о߱����ϲ������ʵĲ���
    'lng����id-����id
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    �ж�ֻ�߱����ϲ��� = False
    gstrSQL = "select ��������, ����id, ������� from ��������˵�� where ����id =[1] And ��������='���ϲ���'"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ���ϲ��ŵĹ�������", lng����ID)
    
    
    If rsTemp.RecordCount = 0 Then
        Exit Function
    End If
    gstrSQL = "select ��������, ����id, ������� from ��������˵�� where ����id =[1] And �������� in( '���Ŀ�','�Ƽ���','����ⷿ')"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ���ϲ��ŵĹ�������", lng����ID)
    
    If rsTemp.RecordCount <> 0 Then
        Exit Function
    End If
    �ж�ֻ�߱����ϲ��� = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckNOExists(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where NO=[2] And ����=[1] And Rownum<2"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���ڸõ���", int����, strNo)
    If rsTemp.RecordCount = 0 Then Exit Function
    ShowMsgBox "�Ѿ����ڸõ��ݺ�(" & strNo & ")"
    CheckNOExists = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '-------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strsql As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strsql = cllProcs(i)
        Call zlDataBase.ExecuteProcedure(strsql, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Sub AddArray(ByRef cllData As Collection, ByVal strsql As String)
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strsql, "K" & i
End Sub

Public Function Check�����ⰴ�����ۼ���() As Boolean
    '����:ȷ��ϵͳ�����ڸ�������µĳɱ����㷽ʽ
    Check�����ⰴ�����ۼ��� = Val(zlDataBase.GetPara(120, glngSys, 0)) = 1
End Function
Public Function ��֤�����ۼ���(ByVal lng�ⷿID As Long, ByVal lng����ID As Long, ByVal lng���� As Long, ByVal lng����ϵ�� As Long, _
                    ByVal dbl����� As Double, ByVal dbl����� As Double, _
                    ByVal dblָ������� As Double, ByVal dbl���� As Double, ByVal dbl���۽�� As Double, _
                    ByRef dblOut��� As Double, ByRef dblOut���� As Double, ByRef dblOut�ɱ���� As Double) As Boolean
    '------------------------------------------------------------------------------------------------------------
    ' ����:��ȡ���εĳɱ��ۺͲ��
    ' ���㹫ʽ:
    '       1.�����<=0��
    '         1) �����-ʵ�ʲ��<=0 Or dbl������� < 0
    '               a.���ĸ���������㷽ʽ=1:
    '                      a)�����ۣ�0��
    '                           ���=���۽��*ָ�������
    '                           �ɱ���=��������-�����ۣ�/����
    '                      b)������>0
    '                           �ɱ���=������
    '                           ��ۣ����۽��-����*�ɱ���
    '               b.���ĸ���������㷽ʽ<>1
    '                           ���=���۽��*ָ�������
    '                           �ɱ���=��������-�����ۣ�/����
    '          2)�����-ʵ�ʲ��>0
    '                �ɱ���= (�����-ʵ�ʲ��)/�������
    '                ��ۣ����۽��-����*�ɱ���
    '        2.�����>0
    '                   ������=������*��ʵ�ʲ��/ʵ�ʽ�
    '                  �ɱ���=��������-�����ۣ�/����
    '------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, dbl���� As Double, dbl������� As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If dbl���� = 0 Then Exit Function
    dbl���� = Get�ɱ���(lng����ID, lng�ⷿID, lng����) * lng����ϵ��
    dbl��� = dbl���۽�� - dbl���� * dbl����
    
'    If dbl����� <= 0 Then
'        If dbl����� - dbl����� > 0 Then
'            gstrSQL = "Select (ʵ�ʽ��-ʵ�ʲ��)/ʵ������ as �ɱ��� From ҩƷ��� where �ⷿid=[1] and ҩƷid=[2] and nvl(����,0)=[3] and nvl(ʵ������,0)>0"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ɱ���", lng�ⷿID, lng����ID, lng����)
'            If rsTemp.EOF = False Then
'                dbl���� = Val(NVL(rsTemp!�ɱ���)) * lng����ϵ��
'            End If
'        End If
'
'        If dbl����� - dbl����� <= 0 Or dbl���� <= 0 Then
'            If Check�����ⰴ�����ۼ��� = True Then
'                dbl���� = Get������(lng����ID) * lng����ϵ��
'                If dbl���� = 0 Then
'                    dbl��� = dbl���۽�� * dblָ�������
'                    dbl���� = (dbl���۽�� - dbl���) / Dbl����
'                Else
'                    dbl��� = dbl���۽�� - Dbl���� * dbl����
'                End If
'            Else
'                    dbl��� = dbl���۽�� * dblָ�������
'                    dbl���� = (dbl���۽�� - dbl���) / Dbl����
'            End If
'        Else
'            'dbl����� - dbl�����>0
'            dbl��� = dbl���۽�� - dbl���� * Dbl����
'        End If
'    Else
'                dbl��� = dbl���۽�� * (dbl����� / dbl�����)
'                dbl���� = (dbl���۽�� - dbl���) / Dbl����
'    End If
    
    dblOut�ɱ���� = Round(dbl���� * dbl����, 7)
    dblOut��� = Round(dbl���, 7)
    dblOut���� = Round(dbl����, 7)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get������(ByVal lng����ID As Long) As Double
    '����:��ȡ������
    '����:lng����ID
    Dim strsql As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select �ɱ��� From �������� where ����id=[1]"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ�ɱ���", lng����ID)
    
    If rsTemp.EOF Then
        Get������ = 0
    Else
        Get������ = Val(zlStr.NVL(rsTemp!�ɱ���))
    End If
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ISCHECK��ǿ�ƿ���ָ���۸�() As Boolean
    '����:�ж��Ƿ�ǿ��Ҫ��������ۼ��ۼ�
     ISCHECK��ǿ�ƿ���ָ���۸� = Val(zlDataBase.GetPara(123, glngSys, 0)) = 1
End Function

Public Function ISCHECK�⹺��ǰ����() As Boolean
    '����:�ж��Ƿ�ǿ��Ҫ��������ۼ��ۼ�
    ISCHECK�⹺��ǰ���� = Val(zlDataBase.GetPara(127, glngSys, 0)) = 1
End Function
 
Public Function Check��ͨ����() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���֤��ǰ��Ա����ͨ���ҵ������Ա
    '����:�Ƿ���true,���򷵻�false
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, bln���ϲ������� As Boolean, strStock As String
    
    On Error GoTo errHandle
    bln���ϲ������� = Val(zlDataBase.GetPara(132, glngSys, 0)) = 1

    If bln���ϲ������� = False Then
        strStock = "K,V,12"
    Else
        strStock = "K,V,W,12"
    End If
    
    Check��ͨ���� = False
    gstrSQL = "" & _
        "   SELECT /*+ Rule*/ DISTINCT a.id, a.���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "       , Table(cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
        "   Where c.�������� = b.���� And (a.վ��=[2] or a.վ�� is null) " & _
        "       And b.����=D.Column_value " & _
        "       AND a.id = c.����id " & _
        "       AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " & _
        "       And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1]) "
        
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ��Ա�ⷿ����", UserInfo.Id, gstrNodeNo, strStock)
    If rsTemp.EOF Then
        Check��ͨ���� = True
    Else
        Check��ͨ���� = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�ɱ���(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long) As Double
'���ܣ���ȡ��ǰҩƷ�ĳɱ��۸�
'������ҩƷid,�ⷿid,����
'����ֵ�� �ɱ��۸�
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo errHandle
    
    gstrSQL = "select ƽ���ɱ��� from ҩƷ��� where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and ����=1"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "�ɱ���", lng����ID, lng�ⷿID, lng����)
    
    If rsData.EOF Then
        blnNullPrice = True
    ElseIf IsNull(rsData!ƽ���ɱ���) = True Then
        blnNullPrice = True
    ElseIf Val(rsData!ƽ���ɱ���) < 0 Then
        blnNullPrice = True
    End If
    
    If Not blnNullPrice Then
        Get�ɱ��� = rsData!ƽ���ɱ���
    Else
        '����޷��ӿ����ȡ�ɱ��ۣ���Ӳ���������ȡ
        gstrSQL = "select �ɱ��� from �������� where ����id=[1]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "�ɱ���", lng����ID)
        If Not rsData.EOF Then
            If Val(NVL(rsData!�ɱ���, 0)) > 0 Then
                Get�ɱ��� = rsData!�ɱ���
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Get���ۼ�(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double) As Double
    '���ܣ���ȡʱ��ҩƷ��ǰҩƷ�����ۼ�
    '����:ҩƷid,�ⷿid,����
    '����ֵ�����ۼ�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo errHandle
    If lng���� <> 0 Then
        gstrSQL = "select ���ۼ� from ҩƷ��� where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and ����=1"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "���ۼ�", lng����ID, lng�ⷿID, lng����)
    Else
        gstrSQL = "Select ʵ�ʽ�� / ʵ������ As ���ۼ�" & vbNewLine & _
                "   From ҩƷ���" & vbNewLine & _
                "   Where �ⷿid = [2] And ҩƷid = [1] And ���� = 1 And ʵ������ > 0"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "���ۼ�", lng����ID, lng�ⷿID)
    End If
    
    If rsData.EOF Or IsNull(rsData!���ۼ�) = True Then
        'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
        '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
        '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
        gstrSQL = "Select �ϴ��ۼ�,ָ�����ۼ�,nvl(ָ�������,0) as ָ�������,nvl(�ӳ���,0) as �ӳ���,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "���ۼ�", lng����ID)
        
        If IsNull(rsData!�ϴ��ۼ�) Then
            dblָ�����ۼ� = rsData!ָ�����ۼ�
            dbl��������� = rsData!���������
            
            Get���ۼ� = 0
            dbl�ɱ��� = Get�ɱ���(lng����ID, lng�ⷿID, lng����)
            dbl�ӳ��� = rsData!�ӳ��� / 100
            dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
            dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
            Get���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
        Else
            Get���ۼ� = rsData!�ϴ��ۼ� * dbl����ϵ��
        End If
    Else
        Get���ۼ� = rsData!���ۼ� * dbl����ϵ��
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    If blnNoBill Then
        x = objPoint.x * 15 'objBill.Left +
        y = objPoint.y * 15 + objBill.Height '+ objBill.Top
    Else
        x = objPoint.x * 15 + objBill.CellLeft
        y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub


Public Function ReturnParaData(ByVal lngSys As Long, ByVal str������IN As String) As ADODB.Recordset
    '-------------------------------------------------------------------------------------------
    '����:��ȡָ���Ĳ���ֵ,����һ����¼��
    '����:lngSys-ϵͳ
    '     str������IN-������In,�Զ��ŷ���
    '
    '����:������¼��
    '����:���˺�
    '����:2007/12/17
    '-------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    strsql = "" & _
        "   Select  /*+ Rule*/ ������,nvl(����ֵ,ȱʡֵ) as ����ֵ,����˵�� " & _
        "   From zlParameters A,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) B" & _
        "   where A.������ = B.Column_Value and a.ϵͳ=[1] and nvl(A.˽��,0)=0 and nvl(a.ģ��,0)=0  " & _
        "   order by ������"
        
    Set rsTemp = zlDataBase.OpenSQLRecord(strsql, "��ȡ����ֵ", lngSys, str������IN)
    
    Set ReturnParaData = rsTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'ȡ�ܣ��£��������꣬��ĵ�һ��
Public Function GetFirstDate(ByVal intInteval As Integer, ByVal datCurrent As Date) As Date
    Dim datReturn As Date
    
    Select Case intInteval
        Case FirstDayOfWeek       '��ǰ�ܵĵ�һ��
            datReturn = DateAdd("d", -Weekday(datCurrent) + 1, Now)
        Case FirstDayOfMonth       '��ǰ�µĵ�һ��
            datReturn = DateAdd("d", -Day(datCurrent) + 1, datCurrent)
        Case FirstDayOfQuarter       '��ǰ���ĵ�һ��
            Select Case DatePart("q", datCurrent)
                Case 1
                    datReturn = DateSerial(Year(datCurrent), 1, 1)
                    
                Case 2
                    datReturn = DateSerial(Year(datCurrent), 4, 1)
                Case 3
                    datReturn = DateSerial(Year(datCurrent), 7, 1)
                Case 4
                    datReturn = DateSerial(Year(datCurrent), 10, 1)
            End Select
        Case FirstDayOfHalfYear       '��ǰ����ĵ�һ��
            If Month(datCurrent) > 6 Then
                datReturn = DateSerial(Year(datCurrent), 7, 1)
            Else
                datReturn = DateSerial(Year(datCurrent), 1, 1)
            End If
        Case FirstDayOfyear       '��ǰ��ĵ�һ��
            datReturn = DateSerial(Year(datCurrent), 1, 1)
    End Select
    GetFirstDate = datReturn
End Function

Public Function Check��������(ByVal lng�ⷿID As Long, ByVal lng����ID As Long, ByVal lng���� As Long, _
    ByVal dbl�������� As Double, ByVal int����� As Integer, Optional ByVal intType As Integer = 0) As Boolean
    '------------------------------------------------------------------------------
    '����:���������ʱ�Ŀ��������Ƿ��㹻
    '����:���㷵�ط���true,���򷵻�False
    '����:
    '    int�����:0-�����;1-��飬��������,2-��飬�����ֹ
    '    intType��0-�����ÿ��,1-���ʵ�ʿ��
    '����:���˺�
    '����:2008/02/15
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, dbl���� As Double
    
    err = 0: On Error GoTo ErrHand:
    '0-�����
    If int����� = 0 Then Check�������� = True: Exit Function
    
    gstrSQL = "Select A.��������,A.ʵ������,B.���� From ҩƷ��� A,�շ���ĿĿ¼ B where A.ҩƷid=B.id And A.ҩƷid=[1] and A.�ⷿid=[2] and nvl(A.����,0)=[3] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�����ÿɴ�", lng����ID, lng�ⷿID, lng����)
    
    If rsTemp.EOF Then
        dbl���� = 0
        gstrSQL = "Select 0 as ��������,B.���� From �շ���ĿĿ¼ B where B.id=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�����ÿɴ�", lng����ID, lng�ⷿID, lng����)
        If rsTemp.EOF Then ShowMsgBox "ָ�����������ϲ�����,����!": Exit Function
    Else
        If intType = 0 Then
            dbl���� = Round(Val(zlStr.NVL(rsTemp!��������, 0)), g_С��λ��.obj_���С��.����С��)
        Else
            dbl���� = Round(Val(zlStr.NVL(rsTemp!ʵ������, 0)), g_С��λ��.obj_���С��.����С��)
        End If
    End If
    
    If dbl���� < Round(dbl��������, g_С��λ��.obj_���С��.����С��) Then
        If intType = 0 Then
            If int����� = 1 Then
                '1-��飬��������
                If MsgBox("��" & zlStr.NVL(rsTemp!����) & "���Ŀ��ÿ�治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            Else
                '2-��飬�����ֹ
                ShowMsgBox "��" & zlStr.NVL(rsTemp!����) & "���Ŀ��ÿ�治�㣬���ܼ�����"
                Exit Function
            End If
        Else
            If int����� = 1 Then
                '1-��飬��������
                If MsgBox("��" & zlStr.NVL(rsTemp!����) & "����ʵ�ʿ�治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            Else
                '2-��飬�����ֹ
                ShowMsgBox "��" & zlStr.NVL(rsTemp!����) & "����ʵ�ʿ�治�㣬���ܼ�����"
                Exit Function
            End If
        End If
    End If
    Check�������� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function ȡ��������(ByVal int���� As Integer, _
    ByVal strNo As String, _
    lng����ID As Long, int��� As Integer, Optional lng���ϵ�� As Long = 1) As Long
    '------------------------------------------------------------------------------
    '����:��ȡ��������
    '����:����ָ���е�����
    '����:���˺�
    '����:2008/02/15
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And ��� = [3] And ҩƷid = [4] And ���ϵ�� = [5]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ�������", int����, strNo, int���, lng����ID, lng���ϵ��)
    If rsTemp.EOF Then
        ȡ�������� = 0
    Else
        ȡ�������� = Val(zlStr.NVL(rsTemp!����))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog

End Function

Public Function SelectItem(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional ByVal blnNotMsg As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '     blnNotMsg-����ʾ.
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a"
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   Where ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  " & _
        "    "
    End If
    gstrSQL = gstrSQL & _
    "   order by ����"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    Set rsTemp = zlDataBase.ShowSQLSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If blnNotMsg = False Then
            ShowMsgBox "û���ҵ���������������,����!"
        End If
        Call zlControl.ControlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = zlStr.NVL(rsTemp!����) & "-" & zlStr.NVL(rsTemp!����)
            .Cell(flexcpData, .Row, .Col) = zlStr.NVL(rsTemp!����)
        End With
    Else
        Call zlControl.ControlSetFocus(objCtl, True)
        objCtl.Text = zlStr.NVL(rsTemp!����)
        objCtl.Tag = zlStr.NVL(rsTemp!����)
        OS.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Select����ѡ����(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str�������� As String = "", _
    Optional bln����Ա As Boolean = False, _
    Optional strsql As String = "") As Boolean
    '------------------------------------------------------------------------------
    '����:����ѡ����
    '����:objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '     str��������-��������:��"V,W,K"
    '     bln����Ա-�Ƿ�Ӳ���Ա����
    '     strSQL-ֱ�Ӹ���SQL��ȡ����(�����ű�ı���һ��Ҫ��A)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    strTittle = "����ѡ����"
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    If strsql <> "" Then
    
        gstrSQL = strsql
    Else
        gstrSQL = "" & _
        "   Select /*+ Rule*/ distinct a.Id,a.�ϼ�id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
        "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��"
    
        If str�������� = "" And bln����Ա = False Then
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a" & _
            "   Where 1=1"
        Else
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a, �������ʷ��� b,��������˵�� c," & _
            IIf(str�������� = "", "", "       (Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.�������� = b.����" & IIf(str�������� = "", "(+)", " and B.����=J.column_value ") & _
            "         AND a.id = c.����id " & _
            IIf(bln����Ա = False, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') or a.����ʱ�� is null ) And (a.վ��=[4] or a.վ�� is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.���� like upper([3]) or a.���� like upper([3]) or a.���� like [3] )"
        If IsNumeric(strSearch) Then                         '���������,��ֻȡ����
            If Mid(gSystem_Para.Para_���뷽ʽ, 1, 1) = "1" Then strFind = " And (A.���� Like Upper([3]))"
        ElseIf zlStr.IsCharAlpha(strSearch) Then            '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            '.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ" ))
            If Mid(gSystem_Para.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And  (a.���� Like Upper([3]))"
        ElseIf zlStr.IsCharChinese(strSearch) Then   'ȫ����
            strFind = " And a.���� Like [3] "
        End If
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strsql = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.�ϼ�id Is Null Connect By Prior A.ID = A.�ϼ�id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.����"
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strsql = "" Then
        '�����¼�
        Set rsTemp = zlDataBase.ShowSQLSelect(FrmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
    Else
        Set rsTemp = zlDataBase.ShowSQLSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.Id, str��������, strKey, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "û�����������Ĳ���,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    Call zlControl.ControlSetFocus(objCtl, True)
    If UCase(TypeName(objCtl)) = UCase("ComboBox") Then
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!Id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgBox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = zlStr.NVL(rsTemp!����) & "-" & zlStr.NVL(rsTemp!����)
        objCtl.Tag = Val(rsTemp!Id)
    End If
    OS.PressKey vbKeyTab
    Select����ѡ���� = True
End Function

Public Function zlCheckIsDate(ByVal strKey As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '����:����Ƿ�Ϸ���������,����Ϊ:(20070101��2007-01-01)����(01-01��0101)����(01<01-31>)
    '����:strKey-��Ҫ���Ĺؽ���
    '����:�Ϸ�������,���ر�׼��ʽ(yyyy-mm-dd),���򷵻�""
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKey) = 4 And InStr(1, strKey, "-") = 0 Then
        '0101,��Ҫ��ǰ�����
        strKey = Year(Now) & strKey
    ElseIf Len(Replace(strKey, "-", "")) = 4 And InStr(1, strKey, "-") > 0 Then
        '01-01��ʽ,��Ҫ����
        strKey = Year(Now) & Replace(strKey, "-", "")
    ElseIf Len(strKey) <= 2 And IsNumeric(strKey) Then
        'ָ����
        strKey = Format(Now, "YYYYMM") & IIf(Len(strKey) = 2, strKey, "0" & strKey)
    End If
    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
        strKey = TranNumToDate(strKey)
        If strKey = "" Then
            ShowMsgBox strTittle & "����Ϊ������,���飡"
            Exit Function
        End If
    End If
    If Not IsDate(strKey) Then
        ShowMsgBox strTittle & "����Ϊ��������(2000-10-10) ��20001010��,���飡"
        Exit Function
    End If
    zlCheckIsDate = strKey
End Function

Public Function zl����δ��˵���(ByVal lng����ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����Ƿ����δ��˵ĵ���
    '���:
    '����:
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-07 15:33:14
    '-----------------------------------------------------------------------------------------------------------

    '���ҩƷ�Ƿ����δ��˵���
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ҩƷid = [1] And Rownum = 1 And ������� Is Null"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "������������Ƿ����δ��˵���", lng����ID)
    zl����δ��˵��� = rsTemp.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Select��Ӧ��(ByVal FrmMain As Form, ByVal objCtl As Control, Optional ByVal strSearch As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��Ӧ��ѡ��
    '���:frmMain-���õ�������
    '    objCtl-���õĿؼ�
    '    strSearch-��������(""��ʾ����ѡ��)
    '����:
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-10 10:38:26
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As Recordset, strKey As String
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim bytStyle As Byte, blnĩ�� As Boolean
    
    
    strKey = GetMatchingSting(strSearch, False)
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    
 
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    gstrSQL = "" & _
        "   Select id,�ϼ�ID,����, ����, ����, ĩ��, ���֤��, ���֤Ч��, ִ�պ�, ִ��Ч��, ˰��ǼǺ�, ��ַ, �绰, ��������," & _
        "           �ʺ�, ��ϵ��, ����, ������, ���ö�, ����ί����, to_char(����ί������,'yyyy-mm-dd') as ����ί������, ������֤��, to_char(������֤����,'yyyy-mm-dd') as ������֤����," & _
        "           ҩ��ֱ�����, to_char(ҩ��ֱ�������,'yyyy-mm-dd') as ҩ��ֱ�������, ��Ȩ��, ��Ȩ��, վ��," & _
        "           to_char(����ʱ��,'yyyy-mm-dd') as ����ʱ��, decode(To_Char(����ʱ��,'yyyy-MM-dd'),'3000-01-01','', to_char(����ʱ��,'yyyy-mm-dd')) as ����ʱ��" & _
        "   From ��Ӧ�� " & _
        "   Where  (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null)  "
    If strSearch = "" Then
        gstrSQL = gstrSQL & _
            "           And (substr(����,5,1)=1 And (վ��=[2] or վ�� is null) Or Nvl(ĩ��,0)=0) " & _
            "   Start with �ϼ�ID is null connect by prior ID =�ϼ�ID " & _
            "   Order by level,ID"
        blnĩ�� = True
        bytStyle = 2
    Else
        gstrSQL = gstrSQL & _
            "    And (վ��=[2] or վ�� is null) And ĩ��=1 And substr(����,5,1)=1 " & _
            "    And (���� like upper([1]) Or ���� like [1] or ���� like [1]) "
        bytStyle = 0
        blnĩ�� = False
    End If
    Set rsTemp = zlDataBase.ShowSQLSelect(FrmMain, gstrSQL, bytStyle, "��Ӧ��ѡ����", Not blnĩ��, "", "��ѡ������������ϵĹ�Ӧ��", False, True, Not blnĩ��, sngX, sngY, lngH, blnCancel, False, False, strKey, gstrNodeNo)
        
    If blnCancel Then
        Call zlControl.ControlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "û���ҵ����������Ĺ�Ӧ��,����!"
        Call zlControl.ControlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = zlStr.NVL(rsTemp!����) & "-" & zlStr.NVL(rsTemp!����)
            .Cell(flexcpData, .Row, .Col) = zlStr.NVL(rsTemp!Id)
        End With
    Else
        Call zlControl.ControlSetFocus(objCtl, True)
        objCtl.Text = zlStr.NVL(rsTemp!����)
        objCtl.Tag = zlStr.NVL(rsTemp!Id)
        OS.PressKey vbKeyTab
    End If
    Select��Ӧ�� = True
End Function

'�����룬���ƣ���������ĳһ��
Public Function FindVsRowNew(ByVal vsBill As VSFlexGrid, ByVal int�Ƚ��� As Integer, _
    ByVal str�Ƚ�ֵ As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo errHandle
    FindVsRowNew = True
    With vsBill
        If .Rows = 2 Then Exit Function
        If str�Ƚ�ֵ = "" Then Exit Function
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                If InStr(1, UCase(strCode), UCase(str�Ƚ�ֵ)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int�Ƚ���
                    .TopRow = .Row
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = "" & _
        " SELECT DISTINCT b.���� " & _
        " FROM (    SELECT DISTINCT A.�շ�ϸĿid " & _
        "           FROM �շ���Ŀ���� A" & _
        "           Where A.���� LIKE upper([1]) " & _
        "       ) a, �շ���ĿĿ¼ B " & _
        " Where a.�շ�ϸĿid = b.ID And (b.վ��=[2] or b.վ�� is null) "
        
        Set rsCode = zlDataBase.OpenSQLRecord(gstrSQL, "����ָ����������", GetMatchingSting(str�Ƚ�ֵ, False), gstrNodeNo)
        If rsCode.EOF Then
            FindVsRowNew = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!����)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int�Ƚ���
                        .TopRow = .Row
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            End If
        Next
        rsCode.Close
    End With
    FindVsRowNew = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

 
Public Function SelectAndNotAddItem(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional blnδ�ҵ����� As Boolean = False, Optional strOra���� As String, Optional strWhere As String) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str���� As String, str���� As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    str���� = strKey
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   And ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  " & _
        "    "
    End If
    gstrSQL = gstrSQL & strWhere & _
    "   order by ����"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDataBase.ShowSQLSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        
        If blnδ�ҵ����� Then
            If zlStr.IsCharChinese(str����) = False Then GoTo NOAdd::
            If MsgBox("ע��:" & vbCrLf & _
                   "     δ�ҵ���ص�" & strTable & ",�Ƿ����ӡ�" & str���� & "����", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                Exit Function
            End If
            
            If AutoAddBaseItem(strTable, str����, str����, strTable & "����", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str����, str���� & "-" & str����)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = IIf(blnOnlyName, str����, str���� & "-" & str����)
                    End If
                End With
            Else
                If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str����, str���� & "-" & str����)
                objCtl.Tag = str����
                OS.PressKey vbKeyTab
            End If
            SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgBox "û���ҵ�����������" & strTable & ",����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, zlStr.NVL(rsTemp!����), zlStr.NVL(rsTemp!����) & "-" & zlStr.NVL(rsTemp!����))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .EditText = zlStr.NVL(rsTemp!����)
                .Cell(flexcpData, .Row, .Col) = zlStr.NVL(rsTemp!����)
            End If
        End With
    Else
        If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = zlStr.NVL(rsTemp!����)
        objCtl.Tag = zlStr.NVL(rsTemp!����)
        OS.PressKey vbKeyTab
    End If
    SelectAndNotAddItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function




Public Function AutoAddBaseItem(ByVal strTable As String, str���� As String, str���� As String, _
    Optional strTittle As String = "������Ŀ", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Զ�������Ŀ��Ϣ(ֻ����б���,���Ƶ���Ϣ����(ֻ���ӣ����������,����)
    '--�����:
    '--������:
    '--��  ��:���ӳɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int���� As Integer, strCode As String, strSpecify As String
    AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("û���ҵ��������" & strTable & "����Ҫ��������" & strTable & "����", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    err = 0: On Error GoTo ErrHand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(����)), 2) As Length FROM  " & strTable
    zlDataBase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int���� = rsTemp!Length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM  " & strTable
    zlDataBase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int���� = Len(strCode)
    strCode = strCode + 1
    
    If int���� >= Len(strCode) Then
    strCode = String(int���� - Len(strCode), "0") & strCode
    End If
    strSpecify = zlStr.GetCodeByVB(str����)
    
    
    gstrSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str���� & "','" & strSpecify & "')"
    zlDataBase.ExecuteProcedure gstrSQL, strTittle
    str���� = strCode
    AutoAddBaseItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub Logogram(ByVal staVal As StatusBar, ByVal bytType As Byte)
'���뷽ʽ
'staVal: StartusBar�ؼ�
'bytType: 0=ƴ��; 1=���;  ��ǰ����״̬
    Dim i As Integer
    For i = 1 To staVal.Panels.Count
        If staVal.Panels(i).Key = "PY" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDataBase.SetPara "���뷽ʽ", 0
                gSystem_Para.int���뷽ʽ = 0
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrRaised
            End If
        ElseIf staVal.Panels(i).Key = "WB" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrRaised
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDataBase.SetPara "���뷽ʽ", 1
                gSystem_Para.int���뷽ʽ = 1
            End If
        End If
    Next
End Sub

Public Function CheckQualifications(ByVal lngModule As Long, ByVal intType As Integer, ByVal strInput As String) As Boolean
    'У�����ģ������̣���Ӧ����Ϣ������Ч��
    'intType��0�����ģ�1�������̣�2����Ӧ��
    'strInput���ַ���ʱΪ���ƣ�����ʱΪID
    Dim rsTmp As ADODB.Recordset
    Dim strMsgInfo As String
    Dim strMsgDate As String
    Dim dateCurrent As Date
    Dim strMsg As String
    
    Dim intCheckType As Integer
    Dim arrColumn
    Dim strCheck As String
    Dim strCheck_���� As String
    Dim strCheck_������ As String
    Dim strCheck_��Ӧ�� As String
    Dim n As Integer
    Dim strTmp As String
    
    On Error GoTo errHandle
    If strInput = "" Then
        CheckQualifications = True
        Exit Function
    End If
        
    '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    strCheck = zlDataBase.GetPara("����У��", glngSys, lngModule, "")
    
    '����Ĳ�����ʽ����ȷʱ�˳�
    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�鷽ʽ��0-����飻1�����ѣ�2����ֹ
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    
    '�����ʱ�˳�
    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�����ݣ�
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '�ֱ�ȡ���ģ������̣���Ӧ����ҪУ�������
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
            If Split(arrColumn(n), ",")(0) = "����" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_���� = IIf(strCheck_���� = "", "", strCheck_���� & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "����������" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_������ = IIf(strCheck_������ = "", "", strCheck_������ & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "���Ĺ�Ӧ��" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_��Ӧ�� = IIf(strCheck_��Ӧ�� = "", "", strCheck_��Ӧ�� & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next
    
    '��У������ʱ�˳�
    If (intType = 0 And strCheck_���� = "") Or (intType = 1 And strCheck_������ = "") Or (intType = 2 And strCheck_��Ӧ�� = "") Then
        CheckQualifications = True
        Exit Function
    End If
    
    dateCurrent = CDate(Format(sys.Currentdate, "yyyy-mm-dd"))
    
    '����
    If intType = 0 Then
        gstrSQL = "Select ('[' || B.���� || ']' || B.����) AS ������Ϣ, A.���֤��, A.���֤��Ч��,ע��֤��,ע��֤��Ч�� " & _
            " From �շ���ĿĿ¼ B,�������� A " & _
            " Where B.ID = A.����ID And A.����ID = [1] "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "У����������", Val(strInput))
        
        If Not rsTmp.EOF Then
            If zlStr.NVL(rsTmp!���֤��) = "" And InStr(strCheck_����, "���֤��") > 0 Then
                strTmp = rsTmp!������Ϣ & "��" & "�����֤��"
            End If
            
            If zlStr.NVL(rsTmp!���֤��Ч��) <> "" Then
                If DateDiff("d", rsTmp!���֤��Ч��, dateCurrent) > 0 And InStr(strCheck_����, "���֤��Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������Ϣ & "��", strTmp & ",") & "���֤�ѹ���"
                End If
            End If
        End If
        If zlStr.NVL(rsTmp!ע��֤��) = "" And InStr(strCheck_����, "ע��֤��") > 0 Then
            strTmp = rsTmp!������Ϣ & "��" & "��ע��֤��"
        End If
        
        If zlStr.NVL(rsTmp!ע��֤��Ч��) <> "" Then
            If DateDiff("d", rsTmp!ע��֤��Ч��, dateCurrent) > 0 And InStr(strCheck_����, "ע��֤��Ч��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!������Ϣ & "��", strTmp & ",") & "ע��֤�ѹ���"
            End If
        End If
    End If
    
    '������
    If intType = 1 Then
        gstrSQL = "Select ('[' || A.���� || ']' || A.����) AS ������, A.������ҵ���֤, A.������ҵ���֤Ч��,a.��Ӫ���֤,a.��Ӫ���֤Ч��,a.��ҵ����ִ��,a.��ҵ����ִ��Ч�� " & _
                        " From ���������� A " & _
                        " Where A.���� = [1] "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "У����������", strInput)
        
        If Not rsTmp.EOF Then
            If zlStr.NVL(rsTmp!������ҵ���֤) = "" And InStr(strCheck_������ & ";", "������ҵ���֤" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "��������ҵ���֤"
            End If
            
            If zlStr.NVL(rsTmp!������ҵ���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "������ҵ���֤Ч��" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "������ҵ���֤�ѹ���"
                End If
            End If
        End If
        If Not rsTmp.EOF Then
            If zlStr.NVL(rsTmp!��Ӫ���֤) = "" And InStr(strCheck_������ & ";", "��Ӫ���֤" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "�޾�Ӫ���֤"
            End If
            
            If zlStr.NVL(rsTmp!��Ӫ���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "��Ӫ���֤Ч��" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "��Ӫ���֤�ѹ���"
                End If
            End If
        End If
        If Not rsTmp.EOF Then
            If zlStr.NVL(rsTmp!��ҵ����ִ��) = "" And InStr(strCheck_������ & ";", "��ҵ����ִ��" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "����ҵ����ִ��"
            End If
            
            If zlStr.NVL(rsTmp!��ҵ����ִ��Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "��ҵ����ִ��Ч��" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "��ҵ����ִ���ѹ���"
                End If
            End If
        End If
    End If
    
    '��Ӧ��
    If intType = 2 Then
        gstrSQL = "Select ('[' || ���� || ']' || ����) AS ��Ӧ��, ˰��ǼǺ�, ���֤��, ִ�պ�, ��Ȩ��, ������֤��, ������֤����, ҩ��ֱ�����, ҩ��ֱ�������, ���֤Ч��, ִ��Ч��, ��Ȩ�� " & _
            " From ��Ӧ�� " & _
            " Where (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And ID = [1] "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "��Ӧ����Ϣ", Val(strInput))
        
        strTmp = ""
        
        If Not rsTmp.EOF Then
            If zlStr.NVL(rsTmp!˰��ǼǺ�) = "" And InStr(strCheck_��Ӧ��, "˰��ǼǺ�") > 0 Then
                strTmp = rsTmp!��Ӧ�� & "��" & "��˰��ǼǺ�"
            End If
            
            If zlStr.NVL(rsTmp!���֤��) = "" And InStr(strCheck_��Ӧ��, "���֤��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "�����֤��"
            End If
            
            If zlStr.NVL(rsTmp!ִ�պ�) = "" And InStr(strCheck_��Ӧ��, "ִ�պ�") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��ִ�պ�"
            End If
            
            If zlStr.NVL(rsTmp!��Ȩ��) = "" And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "����Ȩ��"
            End If
            
            If zlStr.NVL(rsTmp!������֤��) = "" And InStr(strCheck_��Ӧ��, "������֤��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��������֤��"
            End If
            
            If zlStr.NVL(rsTmp!������֤����) <> "" Then
                If DateDiff("d", rsTmp!������֤����, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "������֤����") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "������֤���ѹ���"
                End If
            End If
            
            If zlStr.NVL(rsTmp!ҩ��ֱ�����) = "" And InStr(strCheck_��Ӧ��, "ҩ��ֱ�����") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��ҩ��ֱ�����"
            End If
            
            If zlStr.NVL(rsTmp!ҩ��ֱ�������) <> "" Then
                If DateDiff("d", rsTmp!ҩ��ֱ�������, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ҩ��ֱ�������") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "ҩ��ֱ������ѹ���"
                End If
            End If
            
            If zlStr.NVL(rsTmp!���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!���֤Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "���֤Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "���֤�ѹ���"
                End If
            End If
            
            If zlStr.NVL(rsTmp!ִ��Ч��) <> "" Then
                If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ִ��Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "ִ���ѹ���"
                End If
            End If
            
            If zlStr.NVL(rsTmp!��Ȩ��) <> "" Then
                If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��Ȩ�ѹ���"
                End If
            End If
        End If
    End If
    
    '��ʾ���ֹ
    If strTmp <> "" Then
        If intCheckType = 1 Then
            If MsgBox("δͨ������У�飬�Ƿ������" & vbCrLf & strTmp, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                CheckQualifications = True
                Exit Function
            Else
                Exit Function
            End If
        ElseIf intCheckType = 2 Then
            MsgBox "δͨ������У�飬������⣡" & vbCrLf & strTmp, vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    CheckQualifications = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж��룩
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
'        .ColData(intCol) = lngColWidth
        
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub



Public Sub GetPriceClass()
    '���ݵ�¼վ���ȡҩƷ�ļ۸�ȼ�
    Dim rsData As ADODB.Recordset
    
    If gstrNodeNo <> "" And gstrNodeNo <> "-" Then
        gstrSQL = " Select a.�۸�ȼ� " & _
            " From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ� B " & _
            " Where a.�۸�ȼ� = b.���� And a.���� = 0 And b.�Ƿ�����ҩƷ = 1 And a.վ�� = [1] And Nvl(b.����ʱ��, Sysdate + 1) > Sysdate "
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "GetPriceClass", gstrNodeNo)
        
        If rsData.RecordCount > 0 Then gstrPriceClass = rsData!�۸�ȼ�
    End If
End Sub


Public Function GetPriceClassString(strTableName As String) As String
    '���ݴ����ı������ؼ۸�ȼ���������
    GetPriceClassString = " And " & IIf(strTableName = "", "�۸�ȼ� Is Null ", strTableName & ".�۸�ȼ� Is Null ")
    
End Function

'ȡϵͳ����ֵ
Public Sub GetSysParms()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    
    'ȡ�������������
    gstrSQL = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum < 1"
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, "ȡҩƷ����")
    gtype_UserDrugDigits.Digit_��� = rs.Fields(0).NumericScale
    gtype_UserDrugDigits.Digit_�ɱ��� = rs.Fields(1).NumericScale
    gtype_UserDrugDigits.Digit_���ۼ� = rs.Fields(2).NumericScale
    gtype_UserDrugDigits.Digit_���� = rs.Fields(3).NumericScale
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function StuffWork_GetCheckStockRule(ByVal lng�ⷿID As Long) As Integer
    'ȡ���������
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ���ϳ����� Where �ⷿID=[1]"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ���������", lng�ⷿID)

    If Not rsData.EOF Then
        StuffWork_GetCheckStockRule = rsData!�����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
