Attribute VB_Name = "mdlOps"

Option Explicit

'��������
'######################################################################################################################

'ö��
'----------------------------------------------------------------------------------------------------------------------
Public Enum COLOR_NativeXpPlain
    BackgroundDark = 14054755
    BackgroundLight = 15180411
    HighlightBorderBottomRight = 8388608
    HighlightBorderTopLeft = 8388608
    HighlightHot = 12775167
    HighlightPressed = 4096254
    HighlightSelected = 7323903
    NormalGroupCaptionDark = 14215660
    NormalGroupCaptionLight = 14215660
    NormalGroupCaptionTextHot = 0
    NormalGroupCaptionTextNormal = 0
    NormalGroupClient = 16244694
    NormalGroupClientBorder = 16777215
    NormalGroupClientLink = 12999969
    NormalGroupClientLinkHot = 16748098
    NormalGroupClientText = 0
    SpecialGroupCaptionDark = 14215660
    SpecialGroupCaptionLight = 14215660
    SpecialGroupCaptionTextHot = 0
    SpecialGroupCaptionTextSpecial = 0
    SpecialGroupClient = 16244694
    SpecialGroupClientBorder = 16777215
    SpecialGroupClientLink = 12999969
    SpecialGroupClientLinkHot = 16748098
    SpecialGroupClientText = 0
End Enum
'----------------------------------------------------------------------------------------------------------------------
Public Enum COLOR
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E0E0
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
    ��ɫ = &HF5F5F5
    ����ɫ = 0
    ͣ��ɫ = 255
    �϶�ɫ = &HFFE0D9

End Enum
'----------------------------------------------------------------------------------------------------------------------
Public Enum REGISTER
    ע����Ϣ
    ˽��ģ��
    ˽��ȫ��
    ����ģ��
    ����ȫ��
End Enum

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    p�����¼���� = 1256
    pҽ�����ѹ��� = 1257
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
End Enum

'�Զ������Ͷ���
'----------------------------------------------------------------------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ���ű��� As String
    �������� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���ݿ��û� As String
    ģ��Ȩ�� As String
    ��λ���� As String
End Type
'----------------------------------------------------------------------------------------------------------------------
Public Type TYPE_ICONS_INFO

    ������� As String
    ѪҺƷ�� As String
    ѪҺ��� As String
    ��Ѫ���� As String
    ѪҺ�۸� As String
    
End Type

'----------------------------------------------------------------------------------------------------------------------
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
    support��Ժ��ʵ�ʽ��� = 29 '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
End Enum

'ϵͳ������Ϣ
'----------------------------------------------------------------------------------------------------------------------
Public Type SYSPARAM_INFO
    ���ý��С��λ�� As String
    �շ�������Ŀƥ�� As String
    ����Ʊ�ݺų��� As Integer
    �շ�Ʊ�ݺų��� As Integer
    ���￨���볤�� As Integer
    ���￨��ĸǰ׺ As String
    ���￨������ʾ As Boolean
    ��Ŀ����ƥ�䷽ʽ As Integer '0-˫��;1-����
    ϵͳ�� As Long
    ϵͳ���� As String
    ��Ʒ���� As String
    ģ��� As Long
    ������ As String
    �շ�Ʊ�� As Integer
    ����Ʊ�� As Integer
    ����Ʊ���ϸ���� As Boolean
    �շ�Ʊ���ϸ���� As Boolean
    ����HIS���� As Byte
End Type

'������������
'----------------------------------------------------------------------------------------------------------------------
Public ParamInfo As SYSPARAM_INFO
Public gobjKernel As New clsCISKernel       '�ٴ����Ĳ���
Public gobjRichEPR As New cRichEPR          '�������Ĳ���
Public IconInfo As TYPE_ICONS_INFO
Public UserInfo As TYPE_USER_INFO
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��

'ҽ������
'----------------------------------------------------------------------------------------------------------------------
Public gclsInsure As New clsInsure
Public gblnInsure As Boolean '�Ƿ�����ҽ��
Public gintInsure As Integer
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public gstrUnitName As String '�û���λ����
Public gfrmMain As Object
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public gstrSQL As String
Public gblnOK As Boolean
Public gblnShowInTaskBar As Boolean
Public glngOld As Long
Public glngFormW As Long
Public glngFormH As Long
Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Private mclsUnzip As New cUnzip

'�Զ�����̺ͺ���
'######################################################################################################################

Public Sub CloseRecord(rs As ADODB.Recordset)
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
    '******************************************************************************************************************
    '����: װ��������ָ�������������������е���������
    '******************************************************************************************************************
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Function GetUserInfo() As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡ��½�û���Ϣ
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        " Select A.ID,C.����ID,A.���,A.����,A.����,B.�û���,D.����,D.���� " & _
        " From ��Ա�� A,�ϻ���Ա�� B,������Ա C,���ű� D " & _
        " Where A.ID = B.��ԱID And A.ID = C.��ԱID And C.ȱʡ = 1 AND C.����id=D.ID And Upper(B.�û���) = USER And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOps")
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���ű��� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�������� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPara() As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strSQL = "Select DECODE(����ֵ,NULL,ȱʡֵ,����ֵ) As ����ֵ From ϵͳ������ Where ������=[1]"
    
    '���ý��С��λ��
    '------------------------------------------------------------------------------------------------------------------
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 9)
    If rs.BOF = False Then
        strTmp = Val(zlCommFun.NVL(rs.Fields(0).Value, 2))
        If Val(strTmp) > 0 Then
            strTmp = "0." & String(Val(strTmp), "0")
        Else
            strTmp = "0"
        End If
        
        ParamInfo.���ý��С��λ�� = strTmp
    End If
    
    'Ʊ�ݺų���
    '------------------------------------------------------------------------------------------------------------------
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 20)
    If rs.BOF = False Then
        
        strTmp = zlCommFun.NVL(rs.Fields(0).Value, "")
        If UBound(Split(strTmp, "|")) >= 2 Then ParamInfo.����Ʊ�ݺų��� = Val(Split(strTmp, "|")(2))
        If UBound(Split(strTmp, "|")) >= 0 Then ParamInfo.�շ�Ʊ�ݺų��� = Val(Split(strTmp, "|")(0))
        If UBound(Split(strTmp, "|")) >= 4 Then ParamInfo.���￨���볤�� = Val(Split(strTmp, "|")(4))
    End If
    
    'Ʊ���ϸ����
    '------------------------------------------------------------------------------------------------------------------
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 24)
    If rs.BOF = False Then
        strTmp = zlCommFun.NVL(rs.Fields(0).Value, "")
        If UBound(Split(strTmp, "|")) >= 2 Then ParamInfo.����Ʊ���ϸ���� = (Val(Split(strTmp, "|")(2)) = 1)
        If UBound(Split(strTmp, "|")) >= 0 Then ParamInfo.�շ�Ʊ���ϸ���� = (Val(Split(strTmp, "|")(0)) = 1)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '���￨��ĸǰ׺
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 27)
    If rs.BOF = False Then
        ParamInfo.���￨��ĸǰ׺ = zlCommFun.NVL(rs.Fields(0).Value, "")
    End If
    
    InitSysPara = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '******************************************************************************************************************
    '����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    '******************************************************************************************************************
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function GetNextCode(ByVal strTable As String, Optional ByVal strField As String = "����", Optional ByVal strFilter As String = "") As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strFormat As String

    GetNextCode = "1"
    strFormat = "00000000000000000000"
    gstrSQL = "select nvl(max(" & strField & "),0) as ���� from " & strTable & IIf(strFilter = "", "", " where " & strFilter)

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps")

    If rs.BOF = False Then
        strFormat = IIf(rs!���� = 0, "0000", Mid(strFormat, 1, Len(rs!����)))
        GetNextCode = Format(rs!���� + 1, strFormat)
    End If
    CloseRecord rs
End Function

Public Function CalcStorage(ByVal lngҩƷid As Long, ByVal lng�ⷿID As Long, ByVal vChangePrice As Boolean, ByVal vBatch As Boolean) As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset

    If lngҩƷid = 0 Then Exit Function

    If vChangePrice And vBatch = False Then
        'ֻ��ʵ��ҩƷ

        gstrSQL = "SELECT NVL(A.��������,0) AS �������� FROM ҩƷ��� A WHERE A.ҩƷid=[1] AND A.�ⷿID=[2]"

    ElseIf vChangePrice = False And vBatch Then
        'ֻ��ҩ����������ҩƷ

        gstrSQL = "Select Sum(Nvl(��������,0)) as �������� From ҩƷ���" & _
                    " Where ����=1 " & _
                    " And (Ч�� Is NULL Or Ч��>Trunc(Sysdate)) " & _
                    " And �ⷿID=[2]" & _
                    " And ҩƷID=[1]"

    ElseIf vChangePrice And vBatch Then
        '����ʵ��ҩƷ����ҩ����������ҩƷ

        gstrSQL = "Select Sum(Nvl(��������,0)) as �������� From ҩƷ���" & _
                    " Where ����=1 " & _
                    " And (Ч�� Is NULL Or Ч��>Trunc(Sysdate)) " & _
                    " And �ⷿID=[2]" & _
                    " And ҩƷID=[1]"

    Else
        '�Ȳ���ʵ��ҩƷ�ֲ���ҩ����������ҩƷ,��ֻ��ʵ��ҩƷһ����

        gstrSQL = "SELECT NVL(A.��������,0) AS �������� FROM ҩƷ��� A WHERE A.ҩƷid=[1] AND A.�ⷿID=[2]"

    End If

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngҩƷid, lng�ⷿID)

    If rs.BOF = False Then CalcStorage = zlCommFun.NVL(rs("��������").Value, 0)

    CloseRecord rs
End Function

Public Function CheckAllNumber(ByVal strKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long

    For lngLoop = 1 To Len(strKey)
        If Mid(strKey, lngLoop, 1) < "0" Or Mid(strKey, lngLoop, 1) > "9" Then
            Exit Function
        End If
    Next

    CheckAllNumber = True
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '******************************************************************************************************************
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '******************************************************************************************************************
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

Public Function CheckHaveOrder(ByVal lngKey As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset

    gstrSQL = "SELECT ҽ��״̬ FROM ����ҽ����¼ WHERE ID=[1]"

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)

    CheckHaveOrder = (rs.BOF = False)
    If rs.BOF = False Then
        CheckHaveOrder = (rs("ҽ��״̬").Value <> 4)
    End If

    CloseRecord rs
End Function

Public Function CheckAllowAudit(ByVal lngKey As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "SELECT 1 FROM ����ҽ������ WHERE ִ��״̬>0 AND ҽ��ID=[1]"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    
    CheckAllowAudit = (rs.BOF = True)
    If CheckAllowAudit = False Then
        MsgBox "����ҽ���Ѿ����Ͳ�������ִ�л��Ѿ�ִ����ɣ�", vbInformation, gstrSysName
    End If
    
    CloseRecord rs
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'���ܣ��ж��Ƿ����ָ�������������͵�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH

    strSQL = "Select ���ID From ҩƷ�������� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", bytBill)

    If Not rsTmp.EOF Then ExistIOClass = zlCommFun.NVL(rsTmp!���ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ActualMoney(�ѱ� As String, ������ĿID As Long, ��� As Currency) As Currency
'���ܣ����ݷѱ�,������ĿID,���,����ۺ�Ľ��
'˵��������ۿ۷�Χȡ����ֵ��Χ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    ActualMoney = ���
    If �ѱ� = "" Or ��� = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = _
        "Select " & ��� & "*ʵ�ձ���/100 as ��� From �ѱ���ϸ" & _
        " Where ������ĿID=[1] And �ѱ�=[2]" & _
        " And [3] Between Ӧ�ն���ֵ and Ӧ�ն�βֵ"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", ������ĿID, �ѱ�, Abs(���))
    If Not rsTmp.EOF Then ActualMoney = rsTmp!���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckChargeState(ByVal lngҽ��id As Long, ByVal lng���ͺ� As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '�շ�״̬
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    CheckChargeState = False
    
    strSQL = _
        "select NVL(COUNT(1), 0) AS ���� " & _
              "from ���˷��ü�¼ A, " & _
              "( " & _
                   "select no from ����ҽ������ where ҽ��id+0=" & lngҽ��id & " and ���ͺ�=[1] " & _
                   "Union " & _
                   "select no from ����ҽ������ where ҽ��id=" & lngҽ��id & " and ���ͺ�=[1] " & _
              ") B " & _
            "Where A.NO = B.NO AND NVL(A.��¼״̬,0)=0"
    
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lng���ͺ�)
    
    If rs.BOF Then Exit Function
    If rs("����").Value > 0 Then Exit Function
    
    CheckChargeState = True
    
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
    '******************************************************************************************************************
    '���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
    '������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    '******************************************************************************************************************
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
        End If
    End If
    FormatEx = strNumber
End Function


Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '******************************************************************************************************************
    '���ܣ�
    '******************************************************************************************************************
    MsgBox strInfo, vbInformation, ParamInfo.ϵͳ����
    
End Sub

Public Function ExecutePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    
    Select Case Control.ID
    Case conMenu_File_PrintSet '��ӡ����
    
        Call zlPrintSet
        
    Case conMenu_View_ToolBar_Button '������
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text '��ť����
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size '��ͼ��
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar '״̬��
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        
        Call zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail '���ͷ���
        
        Call zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About '����
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    
    Case conMenu_File_Exit '�˳�
        Unload frmMain
            
    End Select
    
    ExecutePublic = True
End Function

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function

Public Sub LocationObj(ByRef objTxt As Object)
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error Resume Next
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
End Sub

Public Sub LocationGrid(ByRef vsf As Object, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1)
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error Resume Next
    
    If lngRow <> -1 Then vsf.Row = lngRow
    If lngCol <> -1 Then vsf.Col = lngCol
    
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "0") As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.Cols = objVsf.Cols - lngNotPrintCols
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
            lngPrintCol = lngPrintCol + 1
            objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
            objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
            If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                objPrintVsf.ColAlignment(lngPrintCol) = 4
            Else
                objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
            End If
        End If
    Next
    
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                lngPrintCol = lngPrintCol + 1
                
                If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                    objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "��", "")
                Else
                    strFormat = objVsf.ColFormat(lngCol)
                    If strFormat = "" Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                    Else
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                    End If
                End If
            End If
        Next
        Call SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
End Function

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '
    '******************************************************************************************************************
    Dim intCol As Integer
    
    With msf
        
        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '����:��ȡ����ʱ��
    '����:
    '******************************************************************************************************************
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

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
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
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    gstrSQL = "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical")
    GetMaxLength = rs.Fields(0).DefinedSize

End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ָ������Ϣ������ע�����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strKeyValue-��ֵ
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        Call SaveSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue)
        
    Case ˽��ģ��

        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ˽��ȫ��

        Call SaveSetting("ZLSOFT", "˽��ȫ��\" & UserInfo.�û��� & "\" & strSection, strKey, strKeyValue)
        
    Case ����ģ��

        Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ����ȫ��
        
        Call SaveSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '���ܣ� ��ָ����ע����Ϣ��ȡ����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strDefKeyValue-ȱʡ��ֵ
    '���أ� strKeyValue-��ֵ
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        strValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ģ��

        strValue = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ȫ��

        strValue = GetSetting("ZLSOFT", "˽��ȫ��\" & UserInfo.�û��� & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ģ��

        strValue = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ȫ��
        
        strValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
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

Public Function ShowPubSelect(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnLeftSelect As Boolean = False) As Byte
    '******************************************************************************************************************
    '���ܣ�������+�б�ṹ,Ӧ���ڱ��ؼ�
    '������
    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
    '���أ�0:ȡ��ѡ��;1:ѡ��;2:�����ݷ���
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngObjHeight As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        ShowPubSelect = 2
        Exit Function
    End If
    
    If obj Is Nothing Then
        lngX = (Screen.Width - lngCX) / 2
        lngY = (Screen.Width - lngCY) / 2
        lngObjHeight = 0
    Else
        Call ClientToScreen(obj.hWnd, objPoint)
        
        Select Case TypeName(obj)
        Case "TextBox", "CommandButton"
        
            lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
            lngY = obj.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
            lngObjHeight = obj.Height
            
        Case Else
            lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
            lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
            lngObjHeight = obj.CellHeight
        End Select
    End If
    
    ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, blnLeftSelect, False, blnMuliSel, strFilterControl)
                                
    If ShowPubSelect = 1 Then
        Set rsResult = rsData
        
        If rsResult.BOF Then
            ShowPubSelect = 0
        End If
        
    End If

    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetApplyMode(ByVal strText As String) As Byte
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    If CheckStrType(strText, 1) And Left(ParamInfo.�շ�������Ŀƥ��, 1) = 1 Then
        '��ȫ���֣����������
            
        GetApplyMode = 1
        
    ElseIf CheckStrType(strText, 2) And Left(ParamInfo.�շ�������Ŀƥ��, 2) = 1 Then
        '��ȫ��ĸ�����������
        
        GetApplyMode = 2
    Else
        GetApplyMode = 3
    End If
End Function


Public Function AppendCode(ByVal strName As String, ByVal strCode As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    If strName <> "" And strCode <> "" Then
        AppendCode = "��" & strCode & "��" & strName
    Else
        AppendCode = strName
    End If
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
                If MsgBox("ҩƷ��" & strDrugName & "���ڿⷿ��" & strExecuteDept & "��ֻ��" & dbStorage & strUnit & "���Ƿ������", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                    bytWarn = 0
                Else
                    bytWarn = 1
                End If
            Case 2
                MsgBox "ҩƷ��" & strDrugName & "���ڿⷿ��" & strExecuteDept & "��ֻ��" & dbStorage & strUnit & "�������ֹ��", vbOKOnly + vbCritical, ParamInfo.ϵͳ����
                bytWarn = 1
            End Select
        End If
        
    End If
    
    PromptStorageWarn = bytWarn
    
End Function

Public Function BillExistBalance(ByVal strNO As String) As Boolean
    '******************************************************************************************************************
    '���ܣ��ж�ָ�����շѻ��۵��Ƿ�����Ѿ��շѵ�����
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH

    strSQL = "Select ID From ���˷��ü�¼ Where ��¼����=1 And ��¼״̬ IN(1,3) And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Between(X, a, b) As Boolean
    '******************************************************************************************************************
    '���ܣ��ж�x�Ƿ���a��b֮��
    '******************************************************************************************************************
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Public Function IntEx(vNumber As Variant) As Variant
    '******************************************************************************************************************
    '���ܣ�ȡ����ָ����ֵ����С����
    '******************************************************************************************************************
    
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function


Public Function GetDrugWarnOption(ByVal lngKey As Long, ByVal str��� As String) As Integer
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If str��� = "4" Then
        gstrSQL = "SELECT ��鷽ʽ FROM ���ϳ����� WHERE �ⷿID=[1]"
    Else
        gstrSQL = "SELECT ��鷽ʽ FROM ҩƷ������ WHERE �ⷿID=[1]"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)

    If rs.BOF = False Then
        GetDrugWarnOption = Val(IIf(IsNull(rs("��鷽ʽ").Value), 0, rs("��鷽ʽ").Value))
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function CalcTimePrice(ByVal lngҩƷid As Long, lngҩ��ID As Long, ByVal sng���� As Single) As Currency
    '******************************************************************************************************************
    '���ܣ�����ʵ��ҩƷ��ʵ�ʳ����
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim sngסԺ��װ As Single, sng�������� As Single
    Dim curָ�����ۼ� As Currency, cur������ As Currency
    
    sng�������� = sng����

    gstrSQL = "Select Nvl(����,0) as ����,Nvl(��������,0) as ���," & _
        " Nvl(Decode(Nvl(ʵ������,0),0,0,ʵ�ʽ��/ʵ������),0) as ʱ��" & _
        " From ҩƷ���" & _
        " Where ����=1 And �ⷿID=[2] And ҩƷID=[1]" & _
        " And (Nvl(����,0)=0 Or Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
        " Order by Nvl(����,0)"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngҩƷid, lngҩ��ID)
    If Not rsTmp.EOF Then
        While sng���� > 0 And Not rsTmp.EOF
            If rsTmp!��� > sng���� Then
                cur������ = cur������ + Format(sng���� * Format(rsTmp!ʱ��, "0.0000"), "0.00")
                sng���� = 0
            Else
                cur������ = cur������ + Format(rsTmp!��� * Format(rsTmp!ʱ��, "0.0000"), "0.00")
                sng���� = sng���� - rsTmp!���
            End If
            rsTmp.MoveNext
        Wend
        If sng���� <= 0 Then
            If sng�������� <> 0 Then
                CalcTimePrice = Format(cur������ / sng��������, "0.0000")
            Else
                CalcTimePrice = 0 '���Ϊ0
            End If
        Else
            CalcTimePrice = 0 '��治��
        End If
    End If

    CloseRecord rsTmp
End Function

Public Function GetWarnGrade(ByVal WarnGraded As Long, ByVal FeeClass As String, ByVal str�������� As String, ByVal lng����id As Long) As Long
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    GetWarnGrade = 0
    gstrSQL = "select MAX(����) as ���� FROM ("
    gstrSQL = gstrSQL & "select 1 AS ���� from ���ʱ����� where (������־1 like [3] or ������־1='-') And ���ò���=[2] AND ����id=[1]"
    gstrSQL = gstrSQL & " union select 2 AS ���� from ���ʱ����� where (������־2 like [3] or ������־2='-') And ���ò���=[2] AND ����id=[1]"
    gstrSQL = gstrSQL & " union select 3 AS ���� from ���ʱ����� where (������־3 like [3] or ������־3='-') And ���ò���=[2] AND ����id=[1]" & _
        ") A"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id, str��������, "%" & FeeClass & "%")
    
    If rs.BOF = False Then GetWarnGrade = IIf(WarnGraded > zlCommFun.NVL(rs!����, 0), WarnGraded, zlCommFun.NVL(rs!����, 0))
    
End Function

Public Function Ƿ�����(str���� As String, lng����id As Long, lng��ҳid As Long, Optional ByVal curMoney As Single = 0, Optional ByVal str�������� As String, Optional ByVal int������ʽ As Long, Optional ByVal blnǿ�Ƽ��� As Boolean, Optional strǿ�Ʊ������� As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset, strError As String
    Dim int������־ As Integer, int�������� As Integer, sng����ֵ As Single
    Dim sngʣ���ܶ� As Single, sng�������� As Single, sng������ As Single
    
    Ƿ����� = "δ֪"
        
    gstrSQL = "Select ��������,����ֵ From ���ʱ����� A,������ҳ B Where A.���ò���=[3] And A.����ID = B.��ǰ����ID And B.����id =[1] And B.��ҳid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id, lng��ҳid, str��������)
    If rsTmp.BOF Then Exit Function
    sng����ֵ = IIf(IsNull(rsTmp!����ֵ), 0, rsTmp!����ֵ)
    int�������� = IIf(IsNull(rsTmp!��������), 0, rsTmp!��������)
    int������־ = int������ʽ
    
    gstrSQL = "Select ������ From ������Ϣ A Where A.����ID =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id)
    If Not rsTmp.BOF Then sng������ = zlCommFun.NVL(rsTmp!������, 0)
    
    Select Case int��������
    Case 1 '�ۼƷ���
        Set rsTmp = Get������Ϣ(lng����id, lng��ҳid)
        If Not rsTmp.EOF Then sngʣ���ܶ� = zlCommFun.NVL(rsTmp!ʣ���ܶ�, 0)
        sngʣ���ܶ� = sng������ + sngʣ���ܶ� - curMoney
                
        Select Case int������־
        Case 1
            If sngʣ���ܶ� < sng����ֵ Then
                If blnǿ�Ƽ��� Then
                    Ƿ����� = "ǿ��"
                    strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")�ѵ�����ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
                Else
                    Ƿ����� = "����"
                    strError = str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")�ѵ�����ֵ(" & FormatEx(sng����ֵ, 2) & ")������Ҫ������"
                End If
                GoTo EndPoint
            End If
        Case 2
            If sngʣ���ܶ� <= 0 Then
                If blnǿ�Ƽ��� Then
                    Ƿ����� = "ǿ��"
                    strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "Ԥ������Ѿ����꣡"
                Else
                    Ƿ����� = "��"
                    strError = str���� & "Ԥ������Ѿ����꣬��ֹ���ʣ�"
                End If
                GoTo EndPoint
            End If
            
            If sngʣ���ܶ� < sng����ֵ Then
                If blnǿ�Ƽ��� Then
                    Ƿ����� = "ǿ��"
                    strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")С���˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
                Else
                    Ƿ����� = "����"
                    strError = str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")С���˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")������Ҫ������"
                End If
                GoTo EndPoint
            End If
        Case 3
            If sngʣ���ܶ� < sng����ֵ Then
                If blnǿ�Ƽ��� Then
                    Ƿ����� = "ǿ��"
                    strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")С���˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
                Else
                    Ƿ����� = "��"
                    strError = str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")С���˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")����ֹ���ʣ�"
                End If
                GoTo EndPoint
            End If
        End Select
    Case 2              'ÿ�շ���
        gstrSQL = "select sum(ʵ�ս��) as �������� from ���˷��ü�¼ where ����id=[1] and ��ҳid=[2] and trunc(����ʱ��)=[3] "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id, lng��ҳid, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If rsTmp.BOF = False Then
            sng�������� = zlCommFun.NVL(rsTmp!��������, 0) + curMoney
            Select Case int������־
            Case 1
                If sng�������� > sng����ֵ Then
                    If blnǿ�Ƽ��� Then
                        Ƿ����� = "ǿ��"
                        strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "����ķ�������(" & FormatEx(sng��������, 2) & ")�Ѿ������˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
                    Else
                        Ƿ����� = "����"
                        strError = str���� & "����ķ�������(" & FormatEx(sng��������, 2) & ")�Ѿ������˱���ֵ(" & FormatEx(sng����ֵ, 2) & "),����Ҫ������"
                    End If
                    GoTo EndPoint
                End If
            Case 3
                If sng�������� > sng����ֵ Then
                    If blnǿ�Ƽ��� Then
                        Ƿ����� = "ǿ��"
                        strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "����ķ�������(" & FormatEx(sng��������, 2) & ")�Ѿ������˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
                    Else
                        Ƿ����� = "��"
                        strError = str���� & "����ķ�������(" & FormatEx(sng��������, 2) & ")�Ѿ������˱���ֵ(" & FormatEx(sng����ֵ, 2) & "),��ֹ���ʣ�"
                    End If
                    GoTo EndPoint
                End If
            End Select
        End If
    End Select
    Exit Function
EndPoint:
    If Ƿ����� = "��" Then
        MsgBox strError, vbInformation, gstrSysName
    ElseIf Ƿ����� = "ǿ��" Then
        Ƿ����� = "����"
        If InStr(strǿ�Ʊ������� & ";", ";" & str���� & ";") = 0 Then
            strǿ�Ʊ������� = strǿ�Ʊ������� & ";" & str����
            MsgBox strError, vbInformation, gstrSysName
        End If
    Else
        If MsgBox(strError, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Ƿ����� = "��"
    End If
End Function

'Public Function GetWarnGrade(ByVal WarnGraded As Long, ByVal FeeClass As String, ByVal blnҽ�� As Boolean, ByVal lng����id As Long) As Long
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim rs As New ADODB.Recordset
'
'    GetWarnGrade = 0
'    gstrSQL = "select MAX(����) as ���� FROM ("
'    gstrSQL = gstrSQL & "select 1 AS ���� from ���ʱ����� where (������־1 like [3] or ������־1='-') And ���ò���=[2] AND ����id=[1]"
'    gstrSQL = gstrSQL & " union select 2 AS ���� from ���ʱ����� where (������־2 like [3] or ������־2='-') And ���ò���=[2] AND ����id=[1]"
'    gstrSQL = gstrSQL & " union select 3 AS ���� from ���ʱ����� where (������־3 like [3] or ������־3='-') And ���ò���=[2] AND ����id=[1]" & _
'        ") A"
'
'    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id, IIf(blnҽ��, 2, 1), "%" & FeeClass & "%")
'
'    If rs.BOF = False Then GetWarnGrade = IIf(WarnGraded > zlCommFun.NVL(rs!����, 0), WarnGraded, zlCommFun.NVL(rs!����, 0))
'
'End Function
'
'Public Function Ƿ�����(str���� As String, lng����id As Long, lng��ҳid As Long, _
'    Optional ByVal curMoney As Single = 0, Optional blnҽ�� As Boolean, Optional ByVal int������ʽ As Long, _
'    Optional ByVal blnǿ�Ƽ��� As Boolean, Optional strǿ�Ʊ������� As String) As String
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim rsTmp As New ADODB.Recordset, strError As String
'    Dim int������־ As Integer, int�������� As Integer, sng����ֵ As Single
'    Dim sngʣ���ܶ� As Single, sng�������� As Single, sng������ As Single
'
'    Ƿ����� = "δ֪"
'
'    gstrSQL = "Select ��������,����ֵ From ���ʱ����� A,������ҳ B Where A.���ò���=[3] And A.����ID = B.��ǰ����ID And B.����id =[1] And B.��ҳid = [2]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id, lng��ҳid, IIf(blnҽ��, 2, 1))
'    If rsTmp.BOF Then Exit Function
'    sng����ֵ = IIf(IsNull(rsTmp!����ֵ), 0, rsTmp!����ֵ)
'    int�������� = IIf(IsNull(rsTmp!��������), 0, rsTmp!��������)
'    int������־ = int������ʽ
'
'    gstrSQL = "Select ������ From ������Ϣ A Where A.����ID =[1]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id)
'    If Not rsTmp.BOF Then sng������ = zlCommFun.NVL(rsTmp!������, 0)
'
'    Select Case int��������
'    Case 1 '�ۼƷ���
'        Set rsTmp = Get������Ϣ(lng����id, lng��ҳid)
'        If Not rsTmp.EOF Then sngʣ���ܶ� = zlCommFun.NVL(rsTmp!ʣ���ܶ�, 0)
'        sngʣ���ܶ� = sng������ + sngʣ���ܶ� - curMoney
'
'        Select Case int������־
'        Case 1
'            If sngʣ���ܶ� < sng����ֵ Then
'                If blnǿ�Ƽ��� Then
'                    Ƿ����� = "ǿ��"
'                    strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")�ѵ�����ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
'                Else
'                    Ƿ����� = "����"
'                    strError = str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")�ѵ�����ֵ(" & FormatEx(sng����ֵ, 2) & ")������Ҫ������"
'                End If
'                GoTo EndPoint
'            End If
'        Case 2
'            If sngʣ���ܶ� <= 0 Then
'                If blnǿ�Ƽ��� Then
'                    Ƿ����� = "ǿ��"
'                    strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "Ԥ������Ѿ����꣡"
'                Else
'                    Ƿ����� = "��"
'                    strError = str���� & "Ԥ������Ѿ����꣬��ֹ���ʣ�"
'                End If
'                GoTo EndPoint
'            End If
'
'            If sngʣ���ܶ� < sng����ֵ Then
'                If blnǿ�Ƽ��� Then
'                    Ƿ����� = "ǿ��"
'                    strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")С���˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
'                Else
'                    Ƿ����� = "����"
'                    strError = str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")С���˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")������Ҫ������"
'                End If
'                GoTo EndPoint
'            End If
'        Case 3
'            If sngʣ���ܶ� < sng����ֵ Then
'                If blnǿ�Ƽ��� Then
'                    Ƿ����� = "ǿ��"
'                    strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")С���˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
'                Else
'                    Ƿ����� = "��"
'                    strError = str���� & "ʣ����ܶ�(" & FormatEx(sngʣ���ܶ�, 2) & ")С���˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")����ֹ���ʣ�"
'                End If
'                GoTo EndPoint
'            End If
'        End Select
'    Case 2              'ÿ�շ���
'        gstrSQL = "select sum(ʵ�ս��) as �������� from ���˷��ü�¼ where ����id=[1] and ��ҳid=[2] and trunc(����ʱ��)=[3] "
'
'        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id, lng��ҳid, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
'
'        If rsTmp.BOF = False Then
'            sng�������� = zlCommFun.NVL(rsTmp!��������, 0) + curMoney
'            Select Case int������־
'            Case 1
'                If sng�������� > sng����ֵ Then
'                    If blnǿ�Ƽ��� Then
'                        Ƿ����� = "ǿ��"
'                        strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "����ķ�������(" & FormatEx(sng��������, 2) & ")�Ѿ������˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
'                    Else
'                        Ƿ����� = "����"
'                        strError = str���� & "����ķ�������(" & FormatEx(sng��������, 2) & ")�Ѿ������˱���ֵ(" & FormatEx(sng����ֵ, 2) & "),����Ҫ������"
'                    End If
'                    GoTo EndPoint
'                End If
'            Case 3
'                If sng�������� > sng����ֵ Then
'                    If blnǿ�Ƽ��� Then
'                        Ƿ����� = "ǿ��"
'                        strError = "ǿ�Ƽ������ѣ�" & vbCrLf & vbTab & str���� & "����ķ�������(" & FormatEx(sng��������, 2) & ")�Ѿ������˱���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
'                    Else
'                        Ƿ����� = "��"
'                        strError = str���� & "����ķ�������(" & FormatEx(sng��������, 2) & ")�Ѿ������˱���ֵ(" & FormatEx(sng����ֵ, 2) & "),��ֹ���ʣ�"
'                    End If
'                    GoTo EndPoint
'                End If
'            End Select
'        End If
'    End Select
'    Exit Function
'EndPoint:
'    If Ƿ����� = "��" Then
'        MsgBox strError, vbInformation, gstrSysName
'    ElseIf Ƿ����� = "ǿ��" Then
'        Ƿ����� = "����"
'        If InStr(strǿ�Ʊ������� & ";", ";" & str���� & ";") = 0 Then
'            strǿ�Ʊ������� = strǿ�Ʊ������� & ";" & str����
'            MsgBox strError, vbInformation, gstrSysName
'        End If
'    Else
'        If MsgBox(strError, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Ƿ����� = "��"
'    End If
'End Function

Private Function Get������Ϣ(lngID As Long, Optional ByVal lngPageID As Long = 0) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ���ȡָ�����˵�ʣ���
    '******************************************************************************************************************
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If lngPageID = 0 Then
        strSQL = "Select Nvl(A.�������,0) as �������,Nvl(A.Ԥ�����,0) as Ԥ�����,Nvl(A.Ԥ�����,0)-Nvl(A.�������,0) AS ʣ���ܶ� " & _
                "From ������� A Where A.����=1 And A.����ID=[1]"
    Else
        strSQL = "Select Nvl(A.�������,0) as �������,Nvl(A.Ԥ�����,0) as Ԥ�����,Nvl(A.Ԥ�����,0)-Nvl(A.�������,0) + Nvl(B.���,0) AS ʣ���ܶ� " & _
                "From ������� A,(SELECT nvl(SUM(���),0) as ��� from ����ģ����� where ����id=[1] AND ��ҳid=[2]) B Where A.����=1 And A.����ID=[1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lngID, lngPageID)
    
    Set Get������Ϣ = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MakeChargeBill(ByVal lngKey As Long, ByVal int��¼���� As Integer, ByVal strMenuItem As String, Optional ByVal blnZeroBill As Boolean = False, Optional ByVal strPrivs As String) As String
    '******************************************************************************************************************
    '���ܣ�����ҩ�Ͳ��������ɸ��ӷ���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strNO As String
    Dim int��Դ As Integer
            
    Dim lngҽ��id As Long
    Dim int���� As Integer
    Dim lng��ĿID As Long
    Dim lngִ�в���ID As Long
    Dim lng���˲���ID As Long
    Dim lng���˿���ID As Long
    Dim lng���ID As Long
    Dim strDate As String
    Dim lngLoop As Long
    Dim int������Ŀ�� As Integer
    Dim lng���մ���ID As Long
    Dim str���ձ��� As String
    Dim curͳ���� As Currency
    Dim curӦ�� As Currency
    Dim curʵ�� As Currency
    Dim strMsg As String
    Dim dbl���� As Double
    Dim blnTran As Boolean
    Dim cur���� As Currency
    Dim lng�������� As Long
    Dim str�������� As String
    Dim lng�ѱ������� As Long
    Dim lng���� As Long
    Dim str��ǿ�Ʊ������� As String
    Dim blnҽ�� As Boolean
    Dim curMoneyTotal As Currency
    Dim str����С��λ As String
    Dim strSQL As String
    Dim rsSQL As ADODB.Recordset
    Dim blnǿ�Ƽ��� As Boolean
    Dim lng����id As Long
    Dim lng��ҳid As Long
    Dim lng���ͺ� As Long
    
    On Error GoTo errHand
    
    Screen.MousePointer = 11
    
    Call SQLRecord(rsSQL)
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select a.����id,a.��ҳid,a.������Դ,b.���ͺ� From ����ҽ����¼ a,����ҽ������ b Where a.ID=[1] And a.ID=b.ҽ��id"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    If rs.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    lng����id = rs("����id").Value
    lng��ҳid = zlCommFun.NVL(rs("��ҳid").Value, 0)
    int��Դ = rs("������Դ").Value
    lng���ͺ� = zlCommFun.NVL(rs("���ͺ�").Value, 0)
    
    'ȡ���ý���С��
    '------------------------------------------------------------------------------------------------------------------
    str����С��λ = ParamInfo.���ý��С��λ��
    blnǿ�Ƽ��� = (InStr(strPrivs, "Ƿ��ǿ�Ƽ���") > 0)
    
    '��ȡ���˵���Ϣ
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select A.����,A.�Ա�,A.����,Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�," & _
        " A.�����,A.סԺ��,Nvl(A.��ǰ����,B.��Ժ����) as ����," & _
        " Nvl(A.��ǰ����ID,B.��ǰ����ID) as ���˲���ID," & _
        " Nvl(A.��ǰ����ID,B.��Ժ����ID) as ���˿���ID," & _
        " Nvl(B.����,A.����) as ����,C.���� as ������" & _
        " From ������Ϣ A,������ҳ B,ҽ�Ƹ��ʽ C" & _
        " Where A.����ID=[1] And A.����ID=B.����ID(+)" & _
        " And B.��ҳID(+)=[2] And A.ҽ�Ƹ��ʽ=C.����(+)"
    
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����id, lng��ҳid)

    If rsPati.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    blnҽ�� = (Val(zlCommFun.NVL(rsPati!������, "0")) = 1)
    
    '���ܶ��շ���ΪҩƷ����
    '------------------------------------------------------------------------------------------------------------------
    lng���ID = ExistIOClass(IIf(int��¼���� = 1, 8, 9)) '8:���ﻮ�۵�;9:����/סԺ���ʵ�
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    gstrSQL = "SELECT B.ID AS �շ�ϸĿID," & _
                  "A.����,A.�ɷ����,A.����ϵ��,A.��װ," & _
                  "B.���㵥λ," & _
                  "B.���," & _
                  "C.�ּ� AS ����," & _
                  "D.�վݷ�Ŀ," & _
                  "C.������ĿID," & _
                  "A.ִ�п���id," & _
                  "DECODE(A.��ҳid,NULL,F.�����,0,F.�����,F.סԺ��) AS ��ʶ��," & _
                  "F.�ѱ�," & _
                  "A.���˿���id AS ��ǰ����ID," & _
                  "DECODE(F.��ǰ����ID,NULL,A.���˿���id,F.��ǰ����ID) AS ��ǰ����ID," & _
                  "F.��ǰ����," & _
                  "A.����ID," & _
                  "A.��ҳid," & _
                  "F.����," & _
                  "F.�Ա�," & _
                  "F.����," & _
                  "B.���� " & _
            "FROM   �շ���ĿĿ¼ B," & _
               "�շѼ�Ŀ C," & _
               "������Ŀ D," & _
               "������Ϣ F," & _
               "("
    
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "����"
    
        gstrSQL = gstrSQL & _
            "SELECT HH.�ɷ����,Decode(HH.����ϵ��,0,1,Null,1,HH.����ϵ��) As ����ϵ��,Decode(GG.������Դ,2,HH.סԺ��װ,HH.�����װ) As ��װ,GG.���˿���id,3 AS ���,AA.�շ�ϸĿid,AA.����,AA.ִ�п���id,GG.����id,GG.��ҳid ,0 AS ���� " & _
            "FROM ���������Ƽ� AA,����������¼ BB,ҩƷ��� HH,����ҽ����¼ GG " & _
            "Where AA.�շ�ϸĿID = HH.ҩƷid(+) And AA.��¼id = BB.ID And BB.ҽ��id = GG.ID And BB.ҽ��id=[1]"
    '------------------------------------------------------------------------------------------------------------------
    Case "��ҩ"
    
        gstrSQL = gstrSQL & _
            "SELECT HH.�ɷ����,Decode(HH.����ϵ��,0,1,Null,1,HH.����ϵ��) As ����ϵ��,Decode(GG.������Դ,2,HH.סԺ��װ,HH.�����װ) As ��װ,GG.���˿���id,1 AS ���,AA.ҩƷid AS �շ�ϸĿid,AA.ʹ������ AS ����,AA.ִ�п���id,BB.����id,BB.��ҳid ,0 AS ���� " & _
            "FROM ����������ҩ AA,����������¼ BB,ҩƷ��� HH,����ҽ����¼ GG " & _
            "Where AA.ҩƷid = HH.ҩƷid And AA.��¼id = BB.ID And BB.ҽ��id = GG.ID And BB.ҽ��id=[1] "
    '------------------------------------------------------------------------------------------------------------------
    Case "����"
    
        gstrSQL = gstrSQL & _
             "SELECT 0 As �ɷ����,1 As ����ϵ��,1 As ��װ,II.���˿���id,2 AS ���,CC.����id AS �շ�ϸĿid,CC.ʵ������ AS ����,CC.ִ�п���id,DD.����id,DD.��ҳid ,0 AS ���� " & _
             "FROM ������������ CC,����������¼ DD,����ҽ����¼ II " & _
             "Where CC.��¼id = DD.ID And II.ID = DD.ҽ��id And DD.ҽ��id =[1] "
             
    End Select
    
    gstrSQL = gstrSQL & _
               ") A " & _
            "Where C.�շ�ϸĿid = B.ID " & _
               "AND C.������ĿID = D.ID " & _
               "AND C.ִ������ <= SYSDATE " & _
               "AND A.���� > 0 " & _
               "AND (C.��ֹ���� >= SYSDATE OR C.��ֹ���� IS NULL) " & _
               "AND A.�շ�ϸĿid = B.ID " & _
               "AND F.����id=A.����id " & _
            "ORDER BY B.ID"
    
    Set rsCharge = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    If rsCharge.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    '��ɾ��ԭ����
    '------------------------------------------------------------------------------------------------------------------
    Select Case strMenuItem
    
    Case "����"
    
        If int��Դ = 1 And int��¼���� = 1 Then
            
            gstrSQL = "Select Distinct c.No As No From ����������¼ a,���˷��ü�¼ b,���������Ƽ� c " & _
                        "Where Nvl(b.��¼״̬,0) In (0,1) And b.No=c.No And b.��¼����=1 And a.ҽ��id=[1] And c.��¼id=a.ID And c.No Is Not Null "
            
        Else
            
            gstrSQL = "Select Distinct c.No As No From ����������¼ a,���˷��ü�¼ b,���������Ƽ� c " & _
                        "Where Nvl(b.��¼״̬,0)=1 And b.No=c.No And b.��¼����=2 And a.ҽ��id=[1] And c.��¼id=a.ID And c.No Is Not Null "
            
        End If
        
    Case "��ҩ"
        If int��Դ = 1 And int��¼���� = 1 Then
            
            gstrSQL = "Select a.��ҩNo As No From ����������¼ a,���˷��ü�¼ b " & _
                        "Where Nvl(b.��¼״̬,0) In (0,1) And b.No=a.��ҩNo And b.��¼����=1 And a.��ҩNo Is Not Null And a.ҽ��id=[1]"
            
        Else
            
            gstrSQL = "Select a.��ҩNo As No From ����������¼ a,���˷��ü�¼ b " & _
                        "Where Nvl(b.��¼״̬,0)=1 And b.No=a.��ҩNo And b.��¼����=2 And a.��ҩNo Is Not Null And a.ҽ��id=[1]"
            
        End If
            
    Case "����"
        If int��Դ = 1 And int��¼���� = 1 Then
            
            gstrSQL = "Select a.����No As No From ����������¼ a,���˷��ü�¼ b " & _
                        "Where Nvl(b.��¼״̬,0) In (0,1) And b.No=a.����No And b.��¼����=1 And a.����No Is Not Null And a.ҽ��id=[1]"
            
        Else
            
            gstrSQL = "Select a.����No As No From ����������¼ a,���˷��ü�¼ b " & _
                        "Where Nvl(b.��¼״̬,0)=1 And b.No=a.����No And b.��¼����=2 And a.����No Is Not Null And a.ҽ��id=[1]"
            
        End If
    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    If rs.BOF = False Then

        If int��Դ = 1 Then
            If int��¼���� = 1 Then
                '����
                strSQL = "zl_���ﻮ�ۼ�¼_Delete('" & rs("No").Value & "','')"
                Call SQLRecordAdd(rsSQL, strSQL)
            Else
                strSQL = "zl_������ʼ�¼_Delete('" & rs("No").Value & "','','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Else
            strSQL = "zl_סԺ���ʼ�¼_Delete('" & rs("No").Value & "','','" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    End If
    
    
    '
    '------------------------------------------------------------------------------------------------------------------
    With rsCharge
        
        '��ȡ��Ӧ��ҽ����Ϣ
        gstrSQL = "Select ҽ����Ч,���˿���ID,Ӥ��,ִ��Ƶ��,�Ƽ����� From ����ҽ����¼ Where ID=[1]"
        Set rsAdvice = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
        If rsAdvice.BOF Then
            Screen.MousePointer = 0
            Exit Function
        End If
        
        strNO = zlDatabase.GetNextNo(int��¼���� + 12)
         
        '����ҽ�����ӷ���
        '--------------------------------------------------------------------------------------------------------------
        strSQL = "ZL_����ҽ������_Insert(" & lngKey & "," & lng���ͺ� & "," & int��¼���� & ",'" & strNO & "')"
        Call SQLRecordAdd(rsSQL, strSQL)
        
        For lngLoop = 1 To .RecordCount
            
            dbl���� = zlCommFun.NVL(rsCharge("����").Value, 0)
            
            
            '���˲������ҡ�ִ�п���
            '----------------------------------------------------------------------------------------------------------
            lng���˲���ID = zlCommFun.NVL(rsPati!���˲���ID, 0)
            lng���˿���ID = zlCommFun.NVL(rsPati!���˿���ID, 0)
            If lng���˿���ID = 0 Then
                lng���˲���ID = zlCommFun.NVL(rsAdvice!���˿���ID, 0)
                lng���˿���ID = zlCommFun.NVL(rsAdvice!���˿���ID, 0)
            End If
            If lng���˿���ID = 0 Then
                lng���˲���ID = UserInfo.����ID
                lng���˿���ID = UserInfo.����ID
            End If
            
            lngִ�в���ID = !ִ�п���id
            
            cur���� = rsCharge("����").Value
            
            '�����ͨ�շ���Ŀ�Ŀ�棬����ʵ��ҩƷ/���ϵĵ���
            '----------------------------------------------------------------------------------------------------------
            Select Case rsCharge("���").Value
            Case "4", "5", "6", "7"
                Select Case rsCharge("���").Value
                Case "4"
                    gstrSQL = "SELECT NVL(B.�Ƿ���,0) AS ʵ��,NVL(���÷���,0) AS ���� FROM �������� A,�շ���ĿĿ¼ B WHERE A.����id=B.ID AND A.����id=[1] "
                Case "5", "6", "7"
                    '���з������
                    dbl���� = dbl����
                    
                    If zlCommFun.NVL(rsCharge("�ɷ����").Value, 0) = 0 Then
                        dbl���� = dbl���� / zlCommFun.NVL(rsCharge("����ϵ��").Value, 1)
                    Else
                        dbl���� = IntEx(dbl���� / zlCommFun.NVL(rsCharge("����ϵ��").Value, 1) / zlCommFun.NVL(rsCharge("��װ").Value, 1)) * zlCommFun.NVL(rsCharge("��װ").Value, 1)
                    End If
                                            
                    gstrSQL = "SELECT NVL(I.�Ƿ���,0) AS ʵ��,NVL(S.ҩ������,0) AS ���� FROM �շ���ĿĿ¼ I,ҩƷ��� S WHERE I.ID=S.ҩƷid AND S.ҩƷid=[1]"
                End Select
                
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", Val(!�շ�ϸĿid))
                If rs.BOF = False Then
                    If rs("����").Value <> 1 And rs("ʵ��").Value <> 1 Then
                        '����ͨ��Ŀ,Ҫ�����
                        If dbl���� > CalcStorage(!�շ�ϸĿid, lngִ�в���ID, False, False) Then
                            '�����������
                            Select Case GetDrugWarnOption(lngִ�в���ID, IIf(strMenuItem = "��ҩ", "567", "4"))
                            Case 1          '��治������
                                If MsgBox(!���� & "��治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Screen.MousePointer = 0
                                    Exit Function
                                End If
                            Case 2          '��治���ֹ
                                MsgBox !���� & "��治�㣡", vbInformation, gstrSysName
                                Screen.MousePointer = 0
                                Exit Function
                            End Select
                        End If
                    ElseIf rs("ʵ��") = 1 Then
                        cur���� = CalcTimePrice(!�շ�ϸĿid, lngִ�в���ID, dbl����)
                    End If
                End If
            End Select
                           
            '����Ӧ�պ�ʵ�ս��
            '----------------------------------------------------------------------------------------------------------
            curӦ�� = Format(dbl���� * cur����, str����С��λ)
            curʵ�� = IIf(blnZeroBill, 0, curӦ��)
            If rsPati("�ѱ�").Value <> "" And blnZeroBill = False Then curʵ�� = Format(ActualMoney(rsPati("�ѱ�").Value, !������ĿID, curӦ��), str����С��λ)
            
            'ÿ���շ���Ŀ�Ĵ���
            '----------------------------------------------------------------------------------------------------------
            If lng��ĿID <> !�շ�ϸĿid Then
            
                int���� = lngLoop '��ȡ�۸񸸺�
                
                '��ȡ������Ŀ��Ϣ
                '------------------------------------------------------------------------------------------------------
                If int��Դ = 2 And Not IsNull(rsPati!����) And gblnInsure Then
                    strMsg = gclsInsure.GetItemInsure(lng����id, !�շ�ϸĿid, curʵ��, False, rsPati!����)
                    If strMsg <> "" Then
                        int������Ŀ�� = Val(Split(strMsg, ";")(0))
                        lng���մ���ID = Val(Split(strMsg, ";")(1))
                        curͳ���� = Format(Val(Split(strMsg, ";")(2)), "0.00")
                        str���ձ��� = CStr(Split(strMsg, ";")(3))
                    End If
                End If
            End If
            lng��ĿID = !�շ�ϸĿid
            
            
            '����Ǽ��ʵ��ݣ����з��þ���
            '----------------------------------------------------------------------------------------------------------
            
            If int��¼���� = 2 And blnZeroBill = False Then
                
                '������ǰҽ������߱�������,�����ѱ�������Ƚ�
                
'                lng���� = GetWarnGrade(lng�ѱ�������, !���, blnҽ��, lng���˲���ID)
                
                str�������� = ""
                strSQL = "Select zl_PatiWarnScheme([1],[2]) As �������� From Dual"
                Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lng����id, lng��ҳid)
                If rs.BOF = False Then
                    str�������� = zlCommFun.NVL(rs("��������").Value)
                End If
                lng���� = GetWarnGrade(lng�ѱ�������, !���, str��������, lng���˲���ID)
                
                lng�������� = IIf(lng�������� > lng����, lng��������, lng����)
                lng�������� = IIf(lng�������� > lng�ѱ�������, lng��������, lng�ѱ�������)
                            
                '�ж��Ƿ�����Ƿ���
                curMoneyTotal = curMoneyTotal + curʵ��
                
                If lng�������� > lng�ѱ������� Then
                    If curMoneyTotal <> 0 Then
                        'If Ƿ�����(zlCommFun.NVL(rsPati!����), lng����id, lng��ҳid, curMoneyTotal, blnҽ��, lng��������, blnǿ�Ƽ���, str��ǿ�Ʊ�������) = "��" Then
                        If Ƿ�����(zlCommFun.NVL(rsPati!����), lng����id, lng��ҳid, curMoneyTotal, str��������, lng��������, blnǿ�Ƽ���, str��ǿ�Ʊ�������) = "��" Then
                            Screen.MousePointer = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            '��д��¼
            '----------------------------------------------------------------------------------------------------------
            If int��Դ = 1 Then
                If int��¼���� = 1 Then
                    '�������ﻮ�۵���
                    '--------------------------------------------------------------------------------------------------
                    strSQL = _
                        "zl_���ﻮ�ۼ�¼_Insert('" & strNO & "'," & lngLoop & "," & lng����id & ",NULL," & _
                        ZVal(zlCommFun.NVL(rsPati!�����, 0)) & ",'" & zlCommFun.NVL(rsPati!������) & "','" & zlCommFun.NVL(rsPati!����) & "'," & _
                        "'" & zlCommFun.NVL(rsPati!�Ա�) & "','" & zlCommFun.NVL(rsPati!����) & "','" & zlCommFun.NVL(rsPati!�ѱ�) & "',NULL," & _
                        lng���˲���ID & "," & lng���˿���ID & "," & UserInfo.����ID & ",'" & UserInfo.���� & "'," & _
                        "NULL," & lng��ĿID & ",'" & !��� & "','" & !���㵥λ & "',NULL,1," & dbl���� & "," & _
                        "0," & ZVal(lngִ�в���ID) & "," & IIf(int���� = lngLoop, "NULL", int����) & "," & _
                        !������ĿID & ",'" & zlCommFun.NVL(!�վݷ�Ŀ) & "'," & cur���� & "," & curӦ�� & "," & curʵ�� & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.���� & "'," & ZVal(lng���ID) & ",NULL," & _
                        lngKey & ",'" & zlCommFun.NVL(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!ҽ����Ч, 0) & "," & _
                        zlCommFun.NVL(rsAdvice!�Ƽ�����, 0) & ",1)"
                    Call SQLRecordAdd(rsSQL, strSQL)
                Else
                    '����������ʵ���
                    '--------------------------------------------------------------------------------------------------
                    strSQL = _
                        "zl_������ʼ�¼_Insert('" & strNO & "'," & lngLoop & "," & lng����id & "," & _
                        ZVal(zlCommFun.NVL(rsPati!�����, 0)) & ",'" & zlCommFun.NVL(rsPati!����) & "','" & zlCommFun.NVL(rsPati!�Ա�) & "'," & _
                        "'" & zlCommFun.NVL(rsPati!����) & "','" & zlCommFun.NVL(rsPati!�ѱ�) & "',NULL," & ZVal(rsAdvice!Ӥ��) & "," & _
                        lng���˲���ID & "," & lng���˿���ID & "," & UserInfo.����ID & "," & _
                        "'" & UserInfo.���� & "',NULL," & lng��ĿID & ",'" & !��� & "'," & _
                        "'" & !���㵥λ & "',1," & dbl���� & ",0," & ZVal(lngִ�в���ID) & "," & _
                        IIf(int���� = lngLoop, "NULL", int����) & "," & !������ĿID & ",'" & zlCommFun.NVL(!�վݷ�Ŀ) & "'," & cur���� & "," & _
                        curӦ�� & "," & curʵ�� & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.��� & "'," & _
                        "'" & UserInfo.���� & "'," & ZVal(lng���ID) & ",NULL,NULL," & lngKey & "," & _
                        "'" & zlCommFun.NVL(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!ҽ����Ч, 0) & "," & _
                        zlCommFun.NVL(rsAdvice!�Ƽ�����, 0) & ")"
                    Call SQLRecordAdd(rsSQL, strSQL)
                End If
            Else
                '����סԺ���ʵ���
                '------------------------------------------------------------------------------------------------------
                strSQL = _
                    "zl_סԺ���ʼ�¼_Insert('" & strNO & "'," & lngLoop & "," & lng����id & "," & ZVal(lng��ҳid) & "," & _
                    ZVal(zlCommFun.NVL(rsPati!סԺ��, 0)) & ",'" & zlCommFun.NVL(rsPati!����) & "','" & zlCommFun.NVL(rsPati!�Ա�) & "'," & _
                    "'" & zlCommFun.NVL(rsPati!����) & "','" & Trim(zlCommFun.NVL(rsPati!����)) & "','" & zlCommFun.NVL(rsPati!�ѱ�) & "'," & _
                    lng���˲���ID & "," & lng���˿���ID & ",NULL," & ZVal(rsAdvice!Ӥ��) & "," & _
                    UserInfo.����ID & ",'" & UserInfo.���� & "',NULL," & lng��ĿID & ",'" & !��� & "'," & _
                    "'" & !���㵥λ & "'," & int������Ŀ�� & "," & ZVal(lng���մ���ID) & ",'" & str���ձ��� & "'," & _
                    "1," & dbl���� & ",0," & ZVal(lngִ�в���ID) & "," & _
                    IIf(int���� = lngLoop, "NULL", int����) & "," & !������ĿID & ",'" & zlCommFun.NVL(!�վݷ�Ŀ) & "'," & cur���� & "," & _
                    curӦ�� & "," & curʵ�� & "," & curͳ���� & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.��� & "','" & UserInfo.���� & "',NULL," & ZVal(lng���ID) & ",NULL,NULL,NULL," & _
                    lngKey & ",'" & zlCommFun.NVL(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!ҽ����Ч, 0) & "," & _
                    zlCommFun.NVL(rsAdvice!�Ƽ�����, 0) & ",NULL)"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
            
            .MoveNext
            
        Next
        
        '
        '--------------------------------------------------------------------------------------------------------------
        Select Case strMenuItem
        Case "����"
            If .RecordCount > 0 Then
                strSQL = "zl_���������Ƽ�_No(" & lngKey & ",'" & strNO & "'," & int��¼���� & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
            
        Case "��ҩ", "����"
            If .RecordCount > 0 Then
                strSQL = "zl_����������¼_No(" & lngKey & ",'" & strNO & "'," & int��¼���� & ",'" & strMenuItem & "')"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        End Select
        
    End With
    
    '
    '------------------------------------------------------------------------------------------------------------------
        
    blnTran = True
    gcnOracle.BeginTrans
    
    If SQLRecordExecute(rsSQL, "mdlOps", False) = False Then GoTo errHand
        
    '���ύǰ����ҽ������
    '------------------------------------------------------------------------------------------------------------------
    If int��Դ = 2 And Not IsNull(rsPati!����) And gblnInsure Then
        If gclsInsure.GetCapability(support�����ϴ�, lng����id, rsPati!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, lng����id, rsPati!����) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, rsPati!����) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    blnTran = False
    
    '���ύ�����ҽ������
    '------------------------------------------------------------------------------------------------------------------
    If int��Դ = 2 And Not IsNull(rsPati!����) And gblnInsure Then
        If gclsInsure.GetCapability(support�����ϴ�, lng����id, rsPati!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, lng����id, rsPati!����) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, rsPati!����) Then
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                Else
                    MsgBox "����""" & strNO & """��������ҽ������ʧ��,�õ����ѱ��棡", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
        
    Screen.MousePointer = 0
    
    MakeChargeBill = strNO
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If blnTran Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1��ʾ��ʼ;2��ʾ����
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = ParamInfo.ϵͳ����
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
        
            strSQL = CStr(rs("SQL").Value)
            Call zlDatabase.ExecuteProcedure(strSQL, strTitle)
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
End Function

Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.STYLE = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars, Optional ByVal blnEnableCustomization As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
    cbsMain.VisualTheme = xtpThemeOffice2003
        
    With cbsMain.Options
        .ShowExpandButtonAlways = blnEnableCustomization
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization blnEnableCustomization

    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    cbsMain.Options.LargeIcons = False
    
    CommandBarInit = True
    
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '��ӡ����
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '��ӡ����,Ԥ������,�����Excel
        
        If objPrnVsf Is Nothing Then Exit Function
        
        Call SearchPrintData(objPrnVsf, frmPubResource.msfPrint)
        
        '���ô�ӡ��������
        Set objPrint.Body = frmPubResource.msfPrint
        objPrint.Title.Text = strPrintTitle
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add("")
        Call objAppRow.Add("��ӡʱ��:" & Now())
        Call objPrint.BelowAppRows.Add(objAppRow)

        Select Case Control.ID
        Case conMenu_File_Print
            bytMode = zlPrintAsk(objPrint)
            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
        Case conMenu_File_Preview
            zlPrintOrView1Grd objPrint, 2
        Case conMenu_File_Excel
            zlPrintOrView1Grd objPrint, 3
        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        frmMain.cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Size      '��ͼ��
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_StatusBar         '״̬��
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Help              '��������
    
        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((ParamInfo.ϵͳ��) / 100))
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Web_Home          'Web�ϵ�����
        
        Call zlHomePage(frmMain.hWnd)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Web_Forum         'Web�ϵ���̳
    
        Call zlWebForum(frmMain.hWnd)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Web_Mail          '���ͷ���
        
        Call zlMailTo(frmMain.hWnd)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_About             '����
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Exit              '�˳�
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              'ͼ������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '�����˵�����
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Call cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(ParamInfo.ϵͳ��, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

'################################################################################################################
'## ���ܣ�  ��ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  ������ݵ��ļ�����ʧ���򷵻��㳤��""
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]) as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If gobjFSO.FileExists(strZipPath & "TMP.RTF") Then gobjFSO.DeleteFile strZipPath & "TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
End Function

Public Sub ShowDocument(ByRef edt As Object, ByVal lngRecordId As Long, Optional ByVal blnPrivacyProtect As Boolean)
    '******************************************************************************************************************
    '���ܣ�ˢ�²�����ʾ���ݣ�
    '������lngRecordId�����Ӳ�����¼ID��blnPrivacyProtect���Ƿ�������˽����
    '******************************************************************************************************************
    
    Dim mstrPrivs As String
    Dim blnPrivacy As Boolean
    Dim Elements As New cEPRElements
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    
    If blnPrivacyProtect = True Then
        mstrPrivs = ";" & GetPrivFunc(ParamInfo.ϵͳ��, 1070) & ";"
        blnPrivacy = InStr(mstrPrivs, ";������˽����;") = 0     '������˽��Ŀ
    End If
    
    Dim strTemp As String
    Dim strZipFile As String

'    mlngRecordId = lngRecordId
    edt.Freeze
    edt.ReadOnly = False
    edt.NewDoc
    strZipFile = zlBlobRead(5, lngRecordId)
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strTemp) Then
            '���ļ�
            edt.OpenDoc strTemp
            '�����滻��Ŀ
            If blnPrivacy Then
                '��ȡ���е�Ҫ��
                gstrSQL = "Select A.ID,A.������ From ���Ӳ������� A, ��˽������Ŀ B,����������Ŀ C " & _
                    "Where A.�������� = 4 And A.�滻�� = 1 And A.�ļ�id = [1] And A.������� > 0 and B.��Ŀid = C.ID And A.Ҫ������ =C.������ And C.�滻�� = 1 "
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngRecordId)
                If Not rs.EOF Then
                    Do While Not rs.EOF
                        lngKey = Elements.Add(zlCommFun.NVL(rs("������"), 0))
                        Elements("K" & lngKey).GetElementFromDB cprET_�������༭, rs("ID"), True, "���Ӳ�������"
                        '�滻Ҫ������
                        Elements("K" & lngKey).�����ı� = String(Len(Elements("K" & lngKey).�����ı�), "*")
                        Elements("K" & lngKey).Refresh edt
                        rs.MoveNext
                    Loop
                End If
                rs.Close
            End If
            gobjFSO.DeleteFile strTemp, True
        End If
        gobjFSO.DeleteFile strZipFile, True
        edt.SelStart = 0
    End If
    
    If lngRecordId > 0 Then
        '����ҳ���ʽ
        Dim mEPRFileInfo As New cEPRFileDefineInfo
        gstrSQL = "Select c.ID, a.��ʽ From   ����ҳ���ʽ a, �����ļ��б� b, ���Ӳ�����¼ c " & _
                " Where  c.�ļ�id = b.id And a.���� = b.���� And a.��� = b.ҳ�� And c.ID = [1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngRecordId)
        If Not rs.EOF Then
            mEPRFileInfo.��ʽ = zlCommFun.NVL(rs("��ʽ").Value)
            mEPRFileInfo.SetFormat edt, mEPRFileInfo.��ʽ
            edt.ResetWYSIWYG
        End If
        Set mEPRFileInfo = Nothing
    End If
    edt.UnFreeze
    edt.RefreshTargetDC
    edt.ReadOnly = True
End Sub

Public Function GetDefaultDept(ByVal str��� As String, ByVal int������Դ As Integer) As Long
    Dim strTmp As String
    
    strTmp = ""
    Select Case str���
    Case "4"
'                strTmp = IIf(mint������Դ = 1, "����ȱʡ��ҩ��", "סԺȱʡ��ҩ��")
    Case "5"
        strTmp = IIf(int������Դ = 1, "����ȱʡ��ҩ��", "סԺȱʡ��ҩ��")
    Case "6"
        strTmp = IIf(int������Դ = 1, "����ȱʡ��ҩ��", "סԺȱʡ��ҩ��")
    Case "7"
        strTmp = IIf(int������Դ = 1, "����ȱʡ��ҩ��", "סԺȱʡ��ҩ��")
    End Select
    
    If strTmp <> "" Then
        GetDefaultDept = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, strTmp, "0"))
    End If
    
End Function
