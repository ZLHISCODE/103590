Attribute VB_Name = "mdlPACSWork"
Option Explicit
Public gobjRegist As Object
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long

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
Public glngTXTProc As Long
Public gbln�Ӱ�Ӽ� As Boolean
Public grsDuty As ADODB.Recordset '���ҽ��ְ��
Public grsSysPars As ADODB.Recordset

'ҽ������
Public gclsInsure As New clsInsure

'CISϵͳ����
Public gblnҩƷ�������ҽ�� As Boolean
Public gint�����Ǽ���Ч���� As Integer
Public gbln����ҽ��������Ч As Boolean
Public gblnҩ�ƻ��۵� As Boolean
Public gbln�������۵� As Boolean
Public gblnִ�к���� As Boolean

'HISϵͳ����
Public gbln��ҽ As Boolean '�Ƿ�ʹ����ҽ����
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"
Public gbytCardLen As String '���￨�ų���
Public gblnCardHide As Boolean '���￨��������ʾ
Public gstrCardMask As String  '���￨�������ĸǰ׺:AA|BB|CC...
Public gint�Һ����� As Integer '�Һŵ���Ч����
Public gbln�շ���� As Boolean '�Ƿ������������
Public gbln��Ʒ�� As Boolean '����ҩ�Ƿ���Ʒ����ʾ
Public gblnסԺ�Զ����� As Boolean 'סԺ������ɺ��Ƿ��Զ�����
Public gbln�����Զ����� As Boolean '���������ɺ��Ƿ��Զ�����
Public gbln��������ۿ� As Boolean '������Ŀ���ܼ����ۿ�
Public gbln�����������۷��� As Boolean '���ʱ����������۷���

'ҽ������վϵͳ���ò���
Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gblnStock As Boolean 'ָ��ҩ��ʱ�Ƿ��޶�����ҩƷ�Ŀ��
Public gstrҽ���������� As String 'ҽ����������ķ�������
Public gstr���ѷ������� As String '���Ѳ�������ķ�������

Public gintReportFormat As Integer     '��¼�����ӡ��ʽ
'-----------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

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
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

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
    Call SQLTest(App.ProductName, "mdlCureBase", strSQL)
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

Public Function GetMaxBedLen(Optional lng����ID As Long, Optional bln���� As Boolean) As Integer
'���ܣ���ȡָ�����ŵĴ�λ�ŵ���󳤶�
'������lng����ID=����ID�����ID,Ϊ0��ʾ���в��������
'      blnռ��=�Ƿ�ֻ�ܱ�ռ�õĴ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not bln���� Or lng����ID = 0 Then
        strSQL = "Select Max(Length(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIf(lng����ID = 0, " is Not NULL", "=" & lng����ID)
    Else
        strSQL = "Select Max(Length(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIf(lng����ID = 0, " is Not NULL", "=" & lng����ID)
    End If
    
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlCISWork", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!����), 0, rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ID(lng����ID As Long) As Long
'���ܣ��ӿ���ID��ȡ��Ӧ�Ĳ���ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ID From ��λ״����¼ Where ����ID=[1] Group by ����ID"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����ID)
    If Not rsTmp.EOF Then Get����ID = rsTmp!����ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser����IDs() As String
'���ܣ���ȡ����Ա�����Ĳ���(ֱ�����ڲ��������ڿ��������Ĳ���),�����ж��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = _
        "Select Distinct ����ID From (" & _
        " Select A.����ID as ����ID" & _
        " From ��������˵�� A,������Ա B" & _
        " Where A.����ID=B.����ID And B.��ԱID=[1]" & _
        " And A.������� in(1,2,3) And A.��������='����'" & _
        " Union" & _
        " Select A.����ID From ��λ״����¼ A,������Ա B" & _
        " Where A.����ID=B.����ID And B.��ԱID=[1])"
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
'���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
'�������Ƿ�ȡ���������µĿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select ����ID From ������Ա Where ��ԱID=[1]"
    If bln���� Then
        strSQL = strSQL & " Union" & _
            " Select Distinct B.����ID From ������Ա A,��λ״����¼ B" & _
            " Where A.����ID=B.����ID And A.��ԱID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitStockCheck(ByVal int��Χ As Integer) As Collection
'���ܣ���ȡ��ͬ�ⷿ�����鷽ʽ�ڼ�����
'������int��Χ=1-����,2-סԺ
    Dim colStock As Collection
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    '��ͬҩ��ҩƷ�����鷽ʽ
    Set colStock = New Collection
    colStock.Add 0, "_0" '��һ��,����һ����
    
    strSQL = _
        " Select Distinct A.ID,C.��鷽ʽ" & _
        " From ���ű� A,��������˵�� B,ҩƷ������ C" & _
        " Where B.����ID=A.ID And B.������� IN([1],3)" & _
        " And B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " And C.�ⷿID(+)=A.ID"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int��Χ)
    For i = 1 To rsTmp.RecordCount
        colStock.Add Nvl(rsTmp!��鷽ʽ, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    Set InitStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
        
    'HISϵͳ����
    '---------------------------------------------------------
    strSQL = "Select ������,������,����ֵ from ϵͳ������"
    Call OpenRecord(rsTmp, strSQL, "mdlCISWork")
    
    '���ý��С����λ��
    gbytDec = 2: gstrDec = "0.00"
    rsTmp.Filter = "������=9"
    If Not rsTmp.EOF Then
        gbytDec = Val(Nvl(rsTmp!����ֵ, 2))
        gstrDec = "0." & String(gbytDec, "0")
    End If
    
    '���￨��������ʾ
    rsTmp.Filter = "������=12"
    If Not rsTmp.EOF Then gblnCardHide = Nvl(rsTmp!����ֵ, 0) <> 0
    
    'ָ��ҩ��ʱ���ƿ��
    rsTmp.Filter = "������=18"
    If Not rsTmp.EOF Then gblnStock = Nvl(rsTmp!����ֵ, 0) <> 0
    
    '���￨����ĳ���
    gbytCardLen = 7
    rsTmp.Filter = "������=20"
    If Not rsTmp.EOF Then
        gbytCardLen = Val(Split(Nvl(rsTmp!����ֵ, "7|7|7|7|7"), "|")(4))
    End If
    
    '�Һ���Ч����
    rsTmp.Filter = "������=21"
    If Not rsTmp.EOF Then gint�Һ����� = Nvl(rsTmp!����ֵ, 0)
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    rsTmp.Filter = "������=23"
    If Not rsTmp.EOF Then gbytBillOpt = Nvl(rsTmp!����ֵ, 0)
    
    '���￨ʶ��ǰ׺
    rsTmp.Filter = "������=27"
    If Not rsTmp.EOF Then gstrCardMask = UCase(Nvl(rsTmp!����ֵ))
    
    '�Ƿ�ʹ����ҽ
    rsTmp.Filter = "������=31"
    If Not rsTmp.EOF Then gbln��ҽ = Nvl(rsTmp!����ֵ, 0) <> 0
    
    'ҽ����������
    rsTmp.Filter = "������=41"
    If Not rsTmp.EOF Then
        gstrҽ���������� = "'" & Replace(Nvl(rsTmp!����ֵ), "|", "','") & "'"
    End If

    '���ѷ�������
    rsTmp.Filter = "������=42"
    If Not rsTmp.EOF Then
        gstr���ѷ������� = "'" & Replace(Nvl(rsTmp!����ֵ), "|", "','") & "'"
    End If
    
    'סԺ�Զ�����
    rsTmp.Filter = "������=63"
    If Not rsTmp.EOF Then
        gblnסԺ�Զ����� = Nvl(rsTmp!����ֵ, 0) <> 0
    End If
    
    'ҩƷ�������ҽ��
    rsTmp.Filter = "������=69"
    If Not rsTmp.EOF Then gblnҩƷ�������ҽ�� = Val(Nvl(rsTmp!����ֵ, 0)) = 1
    
    'Ƥ�Խ����Чʱ��
    rsTmp.Filter = "������=70"
    If Not rsTmp.EOF Then gint�����Ǽ���Ч���� = Val(Nvl(rsTmp!����ֵ, 0))
    
    '����ҽ��������Ч
    rsTmp.Filter = "������=71"
    If Not rsTmp.EOF Then gbln����ҽ��������Ч = Val(Nvl(rsTmp!����ֵ, 0)) = 1
    
    '�Ƿ�Ҫ�������������
    rsTmp.Filter = "������=72"
    If Not rsTmp.EOF Then gbln�շ���� = Nvl(rsTmp!����ֵ, 1) <> 0
    
    '����ҩ�Ƿ���Ʒ����ʾ
    rsTmp.Filter = "������=74"
    If Not rsTmp.EOF Then gbln��Ʒ�� = Nvl(rsTmp!����ֵ, 0) <> 0
    
    'ҩ�����ɻ��۵�
    rsTmp.Filter = "������=79"
    If Not rsTmp.EOF Then gblnҩ�ƻ��۵� = Nvl(rsTmp!����ֵ, 0) <> 0
    
    '�������ɻ��۵�
    rsTmp.Filter = "������=80"
    If Not rsTmp.EOF Then gbln�������۵� = Nvl(rsTmp!����ֵ, 0) <> 0
    
    'ִ�к��Զ����
    rsTmp.Filter = "������=81"
    If Not rsTmp.EOF Then gblnִ�к���� = Nvl(rsTmp!����ֵ, 0) <> 0
            
    '�����Զ�����
    rsTmp.Filter = "������=92"
    If Not rsTmp.EOF Then
        gbln�����Զ����� = Nvl(rsTmp!����ֵ, 0) <> 0
    End If
    
    '������Ŀ���ܼ����ۿ�
    rsTmp.Filter = "������=93"
    If Not rsTmp.EOF Then gbln��������ۿ� = Nvl(rsTmp!����ֵ, 0) <> 0
    
    '���ʱ����������۷���
    rsTmp.Filter = "������=98"
    If Not rsTmp.EOF Then gbln�����������۷��� = Nvl(rsTmp!����ֵ, 0) <> 0
    
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    InitSysPar = False
End Function

Public Function GetPatiYear(lng����id As Long) As Integer
'���ܣ���ȡ���˵�׼ȷ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intYear As Integer
    
    On Error GoTo errH
    
    strSQL = "Select Sysdate as ��ǰ,��������,���� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����id)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!��������) Then
            intYear = Year(rsTmp!��ǰ) - Year(rsTmp!��������)
            If Format(rsTmp!��ǰ, "MMdd") < Format(rsTmp!��������, "MMdd") Then
                intYear = intYear - 1
            End If
            If intYear < 0 Then intYear = 0
        Else
            intYear = Val(Nvl(rsTmp!����))
        End If
    End If
    GetPatiYear = intYear
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��������(lng����ID As Long) As String
'���ܣ����ز�������
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ���� From ���ű� Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����ID)
    If Not rsTmp.EOF Then Get�������� = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��Ŀ����(lng��ĿID As Long) As String
'���ܣ�����������Ŀ����
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ���� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
    If Not rsTmp.EOF Then Get��Ŀ���� = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getȱʡ�÷�ID(ByVal int���� As Integer, ByVal int��Դ As Integer) As Long
'���ܣ�����ȱʡ�ĸ�ҩ;������ҩ�巨
'������int����=2-��ҩ;��,3-��ҩ�巨,4-��ҩ�÷�,6-�ɼ�����(����)
'      int��Դ=1-����,2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From ������ĿĿ¼" & _
        " Where ���='E' And ��������=[1]" & _
        " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
        " And ������� IN([2],3) And Rownum<100" & _
        " Order by ����"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", CStr(int����), int��Դ)
    If Not rsTmp.EOF Then
        Getȱʡ�÷�ID = rsTmp!ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check�ϰల��(ByVal blnҩ�� As Boolean) As Boolean
'���ܣ����ҽԺ�Ŀ����Ƿ�ʹ�����ϰల��
'������blnҩ��=�Ǽ��ҩ���ϰ໹����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If blnҩ�� Then
        strSQL = "Select Count(B.����ID) as NUM From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� IN('��ҩ��','��ҩ��','��ҩ��')"
    Else
        strSQL = "Select Count(B.����ID) as NUM From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� Not IN('��ҩ��','��ҩ��','��ҩ��')"
    End If
    Call OpenRecord(rsTmp, strSQL, "mdlCISWork")
    If Not rsTmp.EOF Then
        Check�ϰల�� = Nvl(rsTmp!Num, 0) > 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�շ�ִ�п���ID(ByVal str��� As String, ByVal lng��ĿID As Long, _
    ByVal intִ�п��� As Integer, ByVal lng����ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal lng���ϲ��� As Long) As Long
'���ܣ���ȡ��ҩ�շ���Ŀ��ִ�п���
'������int��Χ=1.����,2-סԺ
'      lng����ID=���˿���ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strҩ�� As String, lngҩ�� As Long
    Dim bytDay As Byte, strIDs As String
    
    On Error GoTo errH
    
    If str��� = "4" Then
        '�Լ�SQL�����Ĳ�֧�ִ洢�ⷿ����֮ǰ��
'        strSQL = "Select B.�������,A.����,A.ID From ���ű� A,��������˵�� B" & _
'            " Where A.ID=B.����ID And B.��������='���ϲ���' And B.������� IN([1],3)" & _
'            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
'            " Order by B.�������,A.����"
'        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int��Χ)
'        If Not rsTmp.EOF Then
'            If lng���ϲ��� <> 0 Then rsTmp.Filter = "ID=" & lng���ϲ���
'            If rsTmp.EOF Then rsTmp.Filter = 0
'            Get�շ�ִ�п���ID = rsTmp!ID
'        End If
        
        strSQL = _
            " Select Distinct" & _
            "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            " And B.������� IN([1],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And (A.������Դ is NULL Or A.������Դ=[1])" & _
            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
            " And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int��Χ, lng����ID, lng��ĿID)
        If Not rsTmp.EOF Then
            rsTmp.Filter = "��������ID=" & lng����ID
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
            For i = 1 To rsTmp.RecordCount
                If i = 1 Or rsTmp!ִ�п���ID = lng���ϲ��� Then
                    Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                End If
                rsTmp.MoveNext
            Next
        End If
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIf(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIf(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIf(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not Check�ϰల��(True) Then
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strҩ��, int��Χ, lng����ID, lng��ĿID, bytDay)
        If Not rsTmp.EOF Then
            strIDs = ""
            rsTmp.Filter = "��������ID=" & lng����ID
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
            For i = 1 To rsTmp.RecordCount
                strIDs = strIDs & "," & rsTmp!ִ�п���ID '�ռ����ڶ�̬����
                If i = 1 Or rsTmp!ִ�п���ID = lngҩ�� Then
                    Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                    If rsTmp!ִ�п���ID = lngҩ�� Then
                        strIDs = "": Exit For
                    End If
                End If
                rsTmp.MoveNext
            Next
            strIDs = Mid(strIDs, 2)
            If UBound(Split(strIDs, ",")) <= 0 Then strIDs = ""
        End If
    Else
        Select Case intִ�п���
            Case 0 '0-����ȷ����
                Get�շ�ִ�п���ID = UserInfo.����ID
            Case 1 '1-�������ڿ���
                Get�շ�ִ�п���ID = lng����ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get�շ�ִ�п���ID = lng����ID
                Else
                    Get�շ�ִ�п���ID = Get����ID(lng����ID)
                End If
            Case 3 '3-���������ڿ���
                Get�շ�ִ�п���ID = UserInfo.����ID
            Case 4 '4-ָ������
                strSQL = "Select Nvl(��������ID,0) as ��������ID,ִ�п���ID" & _
                    " From �շ�ִ�п��� Where �շ�ϸĿID=[1]" & _
                    " And (������Դ is NULL Or ������Դ=[2])" & _
                    " And (��������ID is NULL Or ��������ID=[3])" & _
                    " Order by Decode(������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID, int��Χ, lng����ID)
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "��������ID=" & lng����ID
                    If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
                    If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                End If
        End Select
        If Get�շ�ִ�п���ID = 0 Then Get�շ�ִ�п���ID = UserInfo.����ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ִ�п���ID(ByVal str��� As String, ByVal lng��ĿID As Long, _
    ByVal lngҩƷID As Long, ByVal intִ�п��� As Integer, ByVal lng����ID As Long, _
    ByVal int��Ч As Integer, Optional ByVal int��Χ As Integer = 2) As Long
'���ܣ�����������Ŀִ�п�����Ϣ����ȱʡ��ִ�п���ID
'������lngҩƷID=ҩƷID,ȷ�������ʱҪ��
'      intִ�п���=��Ŀִ�п��ұ�־
'      lng����ID=���˿���ID
'      lng��ҩ��,lng��ҩ��,lng��ҩ��=ҩƷȱʡҩ��,ҩƷ��ʱ��Ҫ
'      int��Ч=0-����,1-����,��������Ҫ�ж��ϰ�ʱ��
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strҩ�� As String, lngҩ�� As Long
    Dim bytDay As Byte, bln�ϰల�� As Boolean
    Dim bln��� As Boolean
    
    On Error GoTo errH
    
    If InStr(",5,6,7,", str���) > 0 Then
        bln��� = ((int��Ч = 1 Or gblnҩƷ�������ҽ��) And lngҩƷID <> 0) Or lngҩƷID <> 0
        
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIf(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIf(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIf(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If int��Χ = 1 Then bln�ϰల�� = Check�ϰల��(True) 'סԺҽ������ҩ���ϰల��
        If Not bln�ϰల�� Then
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From " & IIf(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                 IIf(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From " & IIf(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[6]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                IIf(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
                " Order by B.�������,C.����"
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strҩ��, int��Χ, lng����ID, lngҩƷID, lng��ĿID, bytDay)
        If Not rsTmp.EOF Then
            rsTmp.Filter = "��������ID=" & lng����ID
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
            For i = 1 To rsTmp.RecordCount
                If i = 1 Or rsTmp!ִ�п���ID = lngҩ�� Then
                    Get����ִ�п���ID = rsTmp!ִ�п���ID
                    If rsTmp!ִ�п���ID = lngҩ�� Then Exit For
                End If
                rsTmp.MoveNext
            Next
        End If
    Else
        Select Case intִ�п���
            Case 0 '0-��ִ�еĶ���
                Exit Function
            Case 1 '1-�������ڿ���
                Get����ִ�п���ID = lng����ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get����ִ�п���ID = lng����ID
                Else
                    Get����ִ�п���ID = Get����ID(lng����ID)
                End If
            Case 3 '3-���������ڿ���
                Get����ִ�п���ID = UserInfo.����ID
            Case 4 '4-ָ������
                If int��Ч = 1 Then bln�ϰల�� = Check�ϰల��(False)
                If Not bln�ϰల�� Then
                    strSQL = "Select Nvl(��������ID,0) as ��������ID,ִ�п���ID" & _
                        " From ����ִ�п���" & _
                        " Where ������ĿID=[1]" & _
                        " And (������Դ is NULL Or ������Դ=[2])" & _
                        " And (��������ID is NULL Or ��������ID=[3])" & _
                        " Order by Decode(������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                Else
                    bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
                    strSQL = _
                        " Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                        " From ����ִ�п��� A,���Ű��� B" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.����=[4]" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.��ʼʱ��,'HH24:MI:SS') and To_Char(B.��ֹʱ��,'HH24:MI:SS') " & _
                        " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                        " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                        " And A.������ĿID=[1]" & _
                        " Order by Decode(A.������Դ,Null,2,1)"
                End If
                Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID, int��Χ, lng����ID, bytDay)
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "��������ID=" & lng����ID
                    If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
                    If Not rsTmp.EOF Then Get����ִ�п���ID = rsTmp!ִ�п���ID
                End If
            Case 5 '5-Ժ��ִ��
                Exit Function
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ҩ��IDs(ByVal str��� As String, ByVal lng��ĿID As Long, _
    ByVal lngҩƷID As Long, ByVal lng����ID As Long, Optional ByVal int��Χ As Integer = 2) As String
'���ܣ���ȡҩƷ����Ч����ִ�п���ID��,�����ж�ȱʡִ�п���
'������lng����ID=���˿���ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strҩ�� As String
    Dim bytDay As Byte, bln�ϰల�� As Boolean
    Dim strҩ��IDs As String
    
    'ϵͳ����ָ��ҩƷִ�п���,������ȡ���п�ѡ�Ĺ���ѡ��
    If str��� = "5" Then
        strҩ�� = "��ҩ��"
    ElseIf str��� = "6" Then
        strҩ�� = "��ҩ��"
    ElseIf str��� = "7" Then
        strҩ�� = "��ҩ��"
    End If
        
    'ҩƷ��ϵͳָ���Ĵ���ҩ������
    If int��Χ = 1 Then bln�ϰల�� = Check�ϰల��(True) 'סԺҽ������ҩ���ϰల��
    If Not bln�ϰల�� Then
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIf(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
            " And B.������� IN([2],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And (A.������Դ is NULL Or A.������Դ=[2])" & _
            " And (A.��������ID is NULL Or A.��������ID=[3])" & _
            IIf(lngҩƷID <> 0, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]")
    Else
        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIf(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
            " And B.������� IN([2],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And D.����ID=C.ID And D.����=[6]" & _
            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
            " And (A.������Դ is NULL Or A.������Դ=[2])" & _
            " And (A.��������ID is NULL Or A.��������ID=[3])" & _
            IIf(lngҩƷID <> 0, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]")
    End If
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strҩ��, int��Χ, lng����ID, lngҩƷID, lng��ĿID, bytDay)
    Do While Not rsTmp.EOF
        strҩ��IDs = strҩ��IDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    Get����ҩ��IDs = Mid(strҩ��IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ִ�п���(objCbo As Object, ByVal str��� As String, ByVal lng��ĿID As Long, ByVal lngҩƷID As Long, _
    ByVal intִ�п��� As Integer, ByVal lng����ID As Long, ByVal lng��ǰִ��ID As Long, ByVal int��Ч As Integer, Optional ByVal int��Χ As Integer = 2) As Boolean
'���ܣ�����������Ŀִ�п�����Ϣ���ؿ��õ�ִ�п�����ָ����������
'������intִ�п���=��Ŀִ�п��ұ�־
'      lng����ID=���˿���ID
'      lng��ǰִ��ID=ҽ����ǰ��ִ�п���ID
'      int��Ч=0-����,1-����,��������Ҫ�ж��ϰ�ʱ��
'      int��Χ=1-����,2-סԺ(ȱʡ)
'˵�����Է�ҩҽ��,��ǰ��ִ�п��ҿ�����ǿ��ѡ�������,��Ҫ��ʾ��ѡ�����;��ѡ���������һ��������ѡ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strҩ�� As String
    Dim bytDay As Byte, bln�ϰల�� As Boolean
    Dim bln��� As Boolean, i As Long
    
    If InStr(",5,6,7,", str���) > 0 Then
        bln��� = ((int��Ч = 1 Or gblnҩƷ�������ҽ��) And lngҩƷID <> 0) Or lngҩƷID <> 0
        
        'ϵͳ����ָ��ҩƷִ�п���,������ȡ���п�ѡ�Ĺ���ѡ��
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
        End If
            
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If int��Χ = 1 Then bln�ϰల�� = Check�ϰల��(True) 'סԺҽ������ҩ���ϰల��
        If Not bln�ϰల�� Then
            strSQL = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.�������" & _
                " From " & IIf(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                IIf(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.�������" & _
                " From " & IIf(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[8]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                IIf(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
                " Order by B.�������,C.����"
        End If
    Else
        Select Case intִ�п���
            Case 0, 5 '0-��ִ�еĶ���,5-Ժ��ִ��
                Get����ִ�п��� = True: Exit Function
            Case 1 '1-�������ڿ���
                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([3],[6]) Order by ����"
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([3],[6]) Order by ����"
                Else
                    strSQL = _
                        " Select A.ID,A.����,A.����,A.����" & _
                        " From ���ű� A,��λ״����¼ B" & _
                        " Where Rownum<2 And A.ID=B.����ID And B.����ID=[3]" & _
                        " Union " & _
                        " Select ID,����,����,���� From ���ű� Where ID=[6]" & _
                        " Order by ����"
                End If
            Case 3 '3-���������ڿ���
                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([7],[6]) Order by ����"
            Case 4 '4-ָ������
                If int��Ч = 1 Then bln�ϰల�� = Check�ϰల��(False)
                If Not bln�ϰల�� Then
                    strSQL = _
                        " Select Distinct A.ID,A.����,A.����,A.����" & _
                        " From ���ű� A,����ִ�п��� B" & _
                        " Where A.ID=B.ִ�п���ID And B.������ĿID=[5]" & _
                        " And (B.������Դ is NULL Or B.������Դ=[2])" & _
                        " And (B.��������ID is NULL Or B.��������ID=[3])" & _
                        " Union Select ID,����,����,���� From ���ű� Where ID=[6]" & _
                        " Order by ����"
                Else
                    bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����" & _
                        " From ����ִ�п��� A,���Ű��� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.����ID=C.ID And B.����=[8]" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.��ʼʱ��,'HH24:MI:SS') and To_Char(B.��ֹʱ��,'HH24:MI:SS') " & _
                        " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                        " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                        " And A.������ĿID=[5]" & _
                        " Union Select ID,����,����,���� From ���ű� Where ID=[6]" & _
                        " Order by ����"
                End If
        End Select
    End If
        
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strҩ��, int��Χ, lng����ID, lngҩƷID, lng��ĿID, lng��ǰִ��ID, UserInfo.����ID, bytDay)
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        'ʹ��API���ټ���,��Ȼ�����е���
        AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, rsTmp!���� & "-" & rsTmp!����
        SetComboData objCbo.Hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
        If lng��ǰִ��ID = rsTmp!ID Then
            Call zlControl.CboSetIndex(objCbo.Hwnd, i - 1)
        End If
        rsTmp.MoveNext
    Next
    
    '����ҩҽ������ѡ��
    If InStr(",5,6,7,", str���) = 0 Then
        AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, "[����...]"
        SetComboData objCbo.Hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetƵ����Ϣ_����(ByVal str���� As String, strƵ�� As String, _
    intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String) As Boolean
'���ܣ�����Ƶ�ʵ������Ϣ
'������str����=Ƶ�ʱ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strƵ�� = ""
    intƵ�ʴ��� = 0
    intƵ�ʼ�� = 0
    str�����λ = ""
    
    strSQL = "Select * From ����Ƶ����Ŀ Where ����=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str����)
    If Not rsTmp.EOF Then
        strƵ�� = Nvl(rsTmp!����)
        intƵ�ʴ��� = Nvl(rsTmp!Ƶ�ʴ���, 0)
        intƵ�ʼ�� = Nvl(rsTmp!Ƶ�ʼ��, 0)
        str�����λ = Nvl(rsTmp!�����λ)
    End If
    GetƵ����Ϣ_���� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetƵ����Ϣ_����(ByVal strƵ�� As String, intƵ�ʴ��� As Integer, _
    intƵ�ʼ�� As Integer, str�����λ As String, str��Χ As String) As Boolean
'���ܣ�����Ƶ�ʵ������Ϣ
'������strƵ��=Ƶ������
'      str��Χ=1-��ҽ,2-��ҽ,-1-һ����,-2-������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    intƵ�ʴ��� = 0
    intƵ�ʼ�� = 0
    str�����λ = ""
    
    strSQL = "Select * From ����Ƶ����Ŀ Where ����=[1] And Instr([2],','||���÷�Χ||',')>0"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strƵ��, "," & str��Χ & ",")
    If Not rsTmp.EOF Then
        intƵ�ʴ��� = Nvl(rsTmp!Ƶ�ʴ���, 0)
        intƵ�ʼ�� = Nvl(rsTmp!Ƶ�ʼ��, 0)
        str�����λ = Nvl(rsTmp!�����λ)
    End If
    GetƵ����Ϣ_���� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetȱʡƵ��(ByVal int��Χ As Integer, strƵ�� As String, _
    intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String) As Boolean
'���ܣ�����������Ƶ����Ŀ��ȡһ����ΪȱʡƵ��
'������str��Χ=1-��ҽ,2-��ҽ,-1-һ����,-2-������
'���أ�ȱʡƵ����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strƵ�� = ""
    intƵ�ʴ��� = 0
    intƵ�ʼ�� = 0
    str�����λ = ""
    
    strSQL = "Select * From ����Ƶ����Ŀ Where ���÷�Χ=[1] Order by ����"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int��Χ)
    If Not rsTmp.EOF Then
        strƵ�� = Nvl(rsTmp!����)
        intƵ�ʴ��� = Nvl(rsTmp!Ƶ�ʴ���, 0)
        intƵ�ʼ�� = Nvl(rsTmp!Ƶ�ʼ��, 0)
        str�����λ = Nvl(rsTmp!�����λ)
    End If
    GetȱʡƵ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����Ƶ��(int��Χ As Integer, objCbo As Object, Optional strƵ�� As String) As Boolean
'���ܣ���ȡ����Ƶ����Ŀ��ָ����������,������ȱʡ��
'������int��Χ=1-��ҽ,2-��ҽ
'      strƵ��=ȱʡƵ������
'˵������ҩƷ��ѡƵ����Ŀʹ����ҩ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    strSQL = "Select Ӣ������,���� From ����Ƶ����Ŀ Where ���÷�Χ=[1] Order by ����"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int��Χ)
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        objCbo.AddItem Nvl(rsTmp!Ӣ������) & "-" & rsTmp!����
        If strƵ�� = rsTmp!���� Then
            Call zlControl.CboSetIndex(objCbo.Hwnd, objCbo.NewIndex)
        End If
        rsTmp.MoveNext
    Next
    Get����Ƶ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getȱʡʱ��(int��Χ As Integer, strƵ�� As String, Optional lng��ҩ;��ID As Long) As String
'���ܣ���ȡָ��ִ��Ƶ��ȱʡ��ִ��ʱ�䷽��
'������int��Χ=1-��ҽ;2-��ҽ;-1-һ����;-2-������
'      lng��ҩ;��ID=�Ƿ���԰�ָ����ҩ;������ȡ,����ȡ��ȷ����ҩ;���ķ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(A.��ҩ;��ID,0) as �÷�,A.ʱ�䷽��" & _
        " From ����Ƶ��ʱ�� A,����Ƶ����Ŀ B" & _
        " Where A.ִ��Ƶ��=B.���� And B.���÷�Χ=[1]" & _
        " And (A.��ҩ;��ID is NULL Or A.��ҩ;��ID=[2]) And B.����=[3]" & _
        " Order by �÷� Desc,A.�������"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int��Χ, lng��ҩ;��ID, strƵ��)
    If Not rsTmp.EOF Then Getȱʡʱ�� = Nvl(rsTmp!ʱ�䷽��)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getʱ�䷽��(objCbo As Object, int��Χ As Integer, strƵ�� As String, Optional lng��ҩ;��ID As Long) As Boolean
'���ܣ���ȡָ��Ƶ�ʿ��õ�����Ƶ��ʱ�䷽����ָ����������,������ȱʡ��(�򱣳�ԭ��ֵ)
'������int��Χ=1-��ҽ;2-��ҽ;-1-һ����;-2-������
'      strƵ��=����Ƶ����Ŀ����
'      lng��ҩ;��ID=�Ƿ�ֻ��ȡָ����ҩ;����ʱ�䷽��,����Ϊ��ҩƷ��ִ��ʱ�䷽��
'      strִ��ʱ��=ȱʡִ��ʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    '����ͬ����(�����Ƿ�ָ����ҩ;��)��ִ��ʱ��Ӧ�ò���ͬ,������ظ�����
    strSQL = "Select A.�������,A.ʱ�䷽��" & _
        " From ����Ƶ��ʱ�� A,����Ƶ����Ŀ B" & _
        " Where A.ִ��Ƶ��=B.���� And B.����=[1]" & _
        " And (A.��ҩ;��ID is NULL Or A.��ҩ;��ID=[2])" & _
        " And B.���÷�Χ=[3]" & _
        " Order by A.�������"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strƵ��, lng��ҩ;��ID, int��Χ)
    strSQL = objCbo.Text: objCbo.Clear 'Clear�ᵼ��Text���
    For i = 1 To rsTmp.RecordCount
        objCbo.AddItem rsTmp!ʱ�䷽��
        rsTmp.MoveNext
    Next
    objCbo.Text = strSQL: objCbo.Tag = ""
    Getʱ�䷽�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ҽ��(ByVal lng���˿���ID As Long, ByVal bln��ʿվ As Boolean, strȱʡҽ�� As String, lngҽ��ID As Long, _
    Optional objCbo As Object, Optional ByVal int��Χ As Integer = 2) As Boolean
'���ܣ���ȡ���õĿ���ҽ����ָ������������
'������lng���˿���ID=�������ڿ���ID
'      bln��ʿվ=�Ƿ��ɻ�ʿ��ҽ����ҽ��
'      objCbo=Ҫ����ҽ���嵥��������
'      strȱʡҽ��=ȱʡ��λ��ҽ��,�������objCbo,�������ȶ�λ,�ٷ���ȱʡҽ����ҽ��ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    If bln��ʿվ Then
        '�������ڿ��ҵ�ҽ��
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID=[1]" & _
            " Order by A.����"
        '�������ڲ������Ƶ�ҽ��
        strSQL = "Select Distinct ����ID From ��λ״����¼ Where ����ID=[1]"
        strSQL = "Select Distinct ����ID From ��λ״����¼ Where ����ID=(" & strSQL & ")"
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID IN(" & strSQL & ")" & _
            " Order by A.����"
        'ȫԺסԺ���ҵ�ҽ��
        strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN([2],3)"
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID IN(" & strSQL & ")" & _
            " Order by A.����"
    Else 'ҽ����ҽ��ʱ,����Ϊֻ��Ϊҽ������
        strSQL = "Select ID,���,����,���� From ��Ա�� Where ID=" & UserInfo.ID
    End If

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng���˿���ID, int��Χ)
    If objCbo Is Nothing Then
        If Not rsTmp.EOF Then
            If Not bln��ʿվ Then
                lngҽ��ID = rsTmp!ID
                strȱʡҽ�� = rsTmp!����
            ElseIf bln��ʿվ Then
                If strȱʡҽ�� <> "" Then
                    'ȱʡҽ��(סԺҽʦ)����
                    rsTmp.Filter = "����='" & strȱʡҽ�� & "'"
                Else
                    '���˿��ҵ�ҽ������
                    rsTmp.Filter = "����ID=" & lng���˿���ID
                End If
                If rsTmp.EOF Then rsTmp.Filter = 0
                lngҽ��ID = rsTmp!ID
                strȱʡҽ�� = rsTmp!����
            End If
        End If
    Else
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, Nvl(rsTmp!����) & "-" & rsTmp!����
            SetComboData objCbo.Hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
            If rsTmp!���� = strȱʡҽ�� Then
                Call zlControl.CboSetIndex(objCbo.Hwnd, i - 1)
            End If
            'objCbo.AddItem Nvl(rsTmp!����) & "-" & rsTmp!����
            'objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
'            If rsTmp!���� = strȱʡҽ�� Then
'                Call zlControl.CboSetIndex(objCbo.hwnd, objCbo.NewIndex)
'            End If
            rsTmp.MoveNext
        Next
    End If
    Get����ҽ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��������ID(ByVal lngҽ��ID As Long, ByVal lng���˿���ID As Long, Optional ByVal int��Χ As Integer = 2) As Long
'���ܣ���ҽ��ȷ����������
'������int��Χ=1-����,2-סԺ(ȱʡ)
'˵������ҽ���������ҷ�Χ��,����˳�����£�
'      1�����˿���
'      2������������/סԺ���˵Ŀ�����ΪĬ�Ͽ���
'      3������������/סԺ���˵Ŀ���
'      4��Ĭ�Ͽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr����ID(1 To 4) As Long
    
    '���ܲ���û������
    strSQL = "Select Distinct C.����,A.����ID,Nvl(A.ȱʡ,0) as ȱʡ,Nvl(B.�������,0) as �������" & _
        " From ������Ա A,��������˵�� B,���ű� C" & _
        " Where A.����ID=C.ID And A.����ID=B.����ID(+) And A.��ԱID=[1]" & _
        " Order by C.����"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!����ID = lng���˿���ID Then
            arr����ID(1) = rsTmp!����ID
        ElseIf InStr("," & int��Χ & ",3,", rsTmp!�������) > 0 And rsTmp!ȱʡ = 1 Then
            arr����ID(2) = rsTmp!����ID
        ElseIf InStr("," & int��Χ & ",3,", rsTmp!�������) > 0 Then
            If arr����ID(3) = 0 Then arr����ID(3) = rsTmp!����ID
        ElseIf rsTmp!ȱʡ = 1 Then
            arr����ID(4) = rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    For i = LBound(arr����ID) To UBound(arr����ID)
        If arr����ID(i) <> 0 Then
            Get��������ID = arr����ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check����ִ��(ByVal lngִ�п���ID As Long) As Boolean
'���ܣ�ȷ��ָ����ִ�п����Ƿ񱾿�(ҽ������)
'������lngִ�п���ID=ҽ����ִ�п���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr����ID(1 To 4) As Long
    
    strSQL = "Select ����ID From ������Ա Where ��ԱID=[1] And ����ID=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID, lngִ�п���ID)
    Check����ִ�� = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStock(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, Optional ByVal int��Χ As Integer = 2) As Double
'���ܣ���ȡָ���ָⷿ��ҩƷ���������(�������סԺ��λ)
'������int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strTmp = IIf(int��Χ = 1, "����", "סԺ")
    
    '��ȡҩƷ���(�����������ҩƷ),ҩ��������ҩƷ����Ч��
    strSQL = _
        " Select Nvl(Sum(A.��������),0)/Nvl(B." & strTmp & "��װ,1) as ���" & _
        " From ҩƷ��� A,ҩƷ��� B" & _
        " Where A.ҩƷID=B.ҩƷID(+) And A.����=1" & _
        " And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
        " And A.ҩƷID=[1] And A.�ⷿID=[2]" & _
        " Group by Nvl(B." & strTmp & "��װ,1)"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҩƷID, lng�ⷿID)
    If Not rsTmp.EOF Then
        GetStock = Format(rsTmp!���, "0.00000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemField(ByVal lng��ĿID As Long, ByVal strField As String) As Variant
'���ܣ���ȡָ��������Ŀ��ָ���ֶ���Ϣ
'˵����δ����NULLֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & strField & " From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
    If Not rsTmp.EOF Then GetItemField = rsTmp.Fields(strField).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetGroupCount(ByVal lng���ID As Long, ByVal int��Դ As Integer, Optional bln��Ч As Boolean = True) As Long
'���ܣ���ȡ�����Ŀ�е���Ŀ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Count(*) as NUM" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,�շ���ĿĿ¼ C" & _
        " Where A.������ĿID=B.ID And A.�շ�ϸĿID=C.ID(+) And A.�������ID=[1]" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is NULL) And B.������� IN([2],3)" & _
        " And (A.�շ�ϸĿID is NULL Or (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL) And C.������� IN([2],3))" & _
        IIf(bln��Ч And int��Դ = 1, " And A.��Ч=1", "")
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng���ID, int��Դ)
    If Not rsTmp.EOF Then GetGroupCount = Nvl(rsTmp!Num, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetGroupNone(ByVal lng�䷽ID As Long, ByVal int��Դ As Integer) As String
'���ܣ���ȡָ���䷽����Ч�������ҩ��ʾ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    strSQL = "Select B.����" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,ҩƷ��� C,�շ���ĿĿ¼ D" & _
        " Where A.������ĿID=B.ID And B.ID=C.ҩ��ID And C.ҩƷID=D.ID And A.�������ID=[1]" & _
        " And (Not (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is NULL)" & _
        " Or Nvl(B.�������,0) Not IN([2],3))" & _
        " And (Not (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� is NULL)" & _
        " Or Nvl(D.�������,0) Not IN([2],3))"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng�䷽ID, int��Դ)
    Do While Not rsTmp.EOF
        strMsg = strMsg & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    GetGroupNone = Mid(strMsg, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatientFileList(lngDeptID As Long, iFileType As Integer) As ADODB.Recordset
'���ܣ����ݿ��ҺͲ����ļ����ͻ�ȡ��ʹ�õĲ����ļ��嵥
'����˵����
'   lngDeptID������ID
'   iFileType���ļ����͡�0-���ﲡ��;1-סԺ����;2-�����¼;3-�����ļ�;4-���Ƶ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If iFileType = 3 Then
        strSQL = "Select * From �����ļ�Ŀ¼ Where ����=[1] And Ӧ��<>0 Order by ���"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1)
    Else
        strSQL = _
            " Select * From �����ļ�Ŀ¼ Where ����=[1] And " & _
            IIf(lngDeptID = -1, "Ӧ��=1", "Ӧ��=2 And ','||����ID||',' Like [2]") & _
            " Order by ���"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1, "%," & lngDeptID & ",%")
        If rsTmp.EOF Then  'ָ�������޸��ಡ������鹫�ò���
            strSQL = "Select * From �����ļ�Ŀ¼ Where ����=[1] And Ӧ��=1 Order by ���"
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1)
        End If
    End If
    Set GetPatientFileList = rsTmp
End Function

'�����˱��ξ�����ļ��鵵
Public Function PigePatiFile(ByVal lngPatientID As Long, ByVal vPageID As Variant) As Boolean
'lngPatientID������ID
'vPageID���Һŵ���String������ҳID��Long��
    PigePatiFile = False
    On Error GoTo DBError
    PigePatiFile_Proc lngPatientID, vPageID
    On Error GoTo 0
    PigePatiFile = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub PigePatiFile_Proc(ByVal lngPatientID As Long, ByVal vPageID As Variant)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo DBError
    
    gcnOracle.BeginTrans
    
    If TypeName(vPageID) = "String" Then
        strSQL = "Select ID From ���˲�����¼ Where ����ID=[1] And �Һŵ�=[2] And �鵵���� Is Null"
    Else
        strSQL = "Select ID From ���˲�����¼ Where ����ID=[1] And ��ҳID=[2] And �鵵���� Is Null"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngPatientID, vPageID)
    
    Do While Not rsTmp.EOF
        strSQL = "ZL_���˲���_�鵵(" & rsTmp(0) & ",'" + UserInfo.���� + "')"
        ExecuteProc strSQL, "ZL_���˲���_�鵵"
'        zlDatabase.ExecuteProcedure "ZL_���˲���_�鵵(" & rsTmp(0) & ",'" + UserInfo.���� + "')", ""
        rsTmp.MoveNext
    Loop

    gcnOracle.CommitTrans
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "�����ļ��鵵"
End Sub

Public Function CalcDrugPrice(ByVal lngҩƷID As Long, lngҩ��ID As Long, ByVal dbl���� As Double, _
    Optional ByVal str�ѱ� As String, Optional ByVal blnNone�Ӱ�Ӽ� As Boolean) As Double
'���ܣ�����ҩƷʵ��(��ȻҪ����ʵ��,ҩƷ��϶�Ϊ���)
'������dbl����=�ۼ�����,���ѱ����ʱ�������ʵ�ս��
'      str�ѱ�=�Ƿ񰴷ѱ������۵ļ۸�,��Ҫ��ֱ�Ӽ���ҩƷ�Ľ�������ʾ����ʱ��
'      gbln�Ӱ�Ӽ�=����ʱ���������,�����ط���ΪFalse
'      blnNone�Ӱ�Ӽ�=Ϊ��ʱ,"gbln�Ӱ�Ӽ�"��Ч
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl������ As Double, dbl��ǰ���� As Double
    Dim dbl�ܽ�� As Double, dblʱ�� As Double
        
    If dbl���� = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = _
        " Select Nvl(����,0) as ����,Nvl(��������,0) as ���," & _
        " Nvl(Decode(Nvl(ʵ������,0),0,0,ʵ�ʽ��/ʵ������),0) as ʱ��" & _
        " From ҩƷ���" & _
        " Where �ⷿID=[1] And ҩƷID=[2]" & _
        " And ����=1 And (Nvl(����,0)=0 Or Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
        " Order by Nvl(����,0)"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҩ��ID, lngҩƷID)
    
    dbl�ܽ�� = 0: dbl������ = dbl����
    For i = 1 To rsTmp.RecordCount
        If dbl������ = 0 Then Exit For
        If dbl������ <= rsTmp!��� Then
            dbl��ǰ���� = dbl������
        Else
            dbl��ǰ���� = rsTmp!���
        End If
        dbl�ܽ�� = dbl�ܽ�� + Format(dbl��ǰ���� * Format(rsTmp!ʱ��, "0.00000"), gstrDec)
        dbl������ = Val(dbl������) - Val(dbl��ǰ����)
        rsTmp.MoveNext
    Next
    If dbl������ <> 0 Then
        dblʱ�� = 0 '��治��
    Else
        dblʱ�� = Format(dbl�ܽ�� / dbl����, "0.00000")
        
        '�Ӱ�Ӽ۴���
        If gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ� Then
            strSQL = _
                " Select To_Number([2])" & _
                    IIf(gbln�Ӱ�Ӽ�, "*Decode(A.�Ӱ�Ӽ�,1,1+Nvl(B.�Ӱ�Ӽ���,0)/100,1)", "") & " as ���" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
                " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҩƷID, dblʱ��)
            If Not rsTmp.EOF Then dblʱ�� = Nvl(rsTmp!���, 0)
        End If
        
        If str�ѱ� <> "" Then
            '�����Ŀֻ��һ��������Ŀ
            strSQL = _
                "Select To_Number([1])*B.ʵ�ձ���/100 as ���" & _
                " From �շѼ�Ŀ A,�ѱ���ϸ B" & _
                " Where A.������ĿID=B.������ĿID And B.�ѱ�=[2]" & _
                " And [3] Between B.Ӧ�ն���ֵ And B.Ӧ�ն�βֵ" & _
                " And A.�շ�ϸĿID=[4]"
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", Val(Format(dblʱ�� * dbl����, gstrDec)), str�ѱ�, Abs(dblʱ��), lngҩƷID)
            If Not rsTmp.EOF Then dblʱ�� = Nvl(rsTmp!���, 0)
        End If
    End If
    CalcDrugPrice = dblʱ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcPrice(ByVal lng��ĿID As Long, Optional ByVal str�ѱ� As String, _
    Optional ByVal dbl���� As Double, Optional ByVal blnNone�Ӱ�Ӽ� As Boolean) As Double
'���ܣ���ȡ�շ�ϸĿ�ĵ�ǰ�ۼۼ۸���,��۷���0
'������str�ѱ�=�Ƿ񰴷ѱ������۵�ʵ�ս��
'      dbl����=���ѱ����ʱ,����Ҫ��������(���ۼ۵�λ),��ʱ�������ʵ�ս��
'      gbln�Ӱ�Ӽ�=����ʱ���������,�����ط���ΪFalse
'      blnNone�Ӱ�Ӽ�=Ϊ��ʱ,"gbln�Ӱ�Ӽ�"��Ч
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl��� As Double
    
    On Error GoTo errH
    
    If str�ѱ� = "" Then
        strSQL = _
            " Select Sum(Decode(Nvl(A.�Ƿ���,0),1,NULL," & _
                "B.�ּ�" & IIf(gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ�, "*Decode(A.�Ӱ�Ӽ�,1,1+Nvl(B.�Ӱ�Ӽ���,0)/100,1)", "") & ")) as ���" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
            " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
        If Not rsTmp.EOF Then dbl��� = Nvl(rsTmp!���, 0)
    Else
        '�������Խ�ActualMoney������SQLһ��д��������ѱ���ܱ�ɾ�����󲻳�����
        strSQL = _
            " Select B.������ĿID,Decode(Nvl(A.�Ƿ���,0),1,NULL," & _
                "B.�ּ�" & IIf(gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ�, "*Decode(A.�Ӱ�Ӽ�,1,1+Nvl(B.�Ӱ�Ӽ���,0)/100,1)", "") & ") as ���" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
            " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
        For i = 1 To rsTmp.RecordCount
            dbl��� = dbl��� + ActualMoney(str�ѱ�, rsTmp!������ĿID, Format(dbl���� * Format(Nvl(rsTmp!���, 0), "0.00000"), gstrDec))
            rsTmp.MoveNext
        Next
    End If
    CalcPrice = dbl���
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
        "Select To_Number([1])*ʵ�ձ���/100 as ��� From �ѱ���ϸ" & _
        " Where ������ĿID=[2] And �ѱ�=[3] And Abs([1]) Between Ӧ�ն���ֵ and Ӧ�ն�βֵ"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", ���, ������ĿID, �ѱ�)
    If Not rsTmp.EOF Then ActualMoney = rsTmp!���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�ٴ�����(ByVal int��Χ As Integer, Optional ByVal lng���˿���ID As Long, _
    Optional lngȱʡ����ID As Long, Optional objCbo As Object, Optional ByVal blnBed As Boolean) As Boolean
'���ܣ������ٴ������嵥��ȱʡ�ٴ�����
'������int��Χ=1-����,2-סԺ,3-�����סԺ
'      lng���˿���ID=���˵�ǰ�Ŀ���,����Ҫ�ſ��ÿ���
'      objCbo=Ҫ��������嵥��������,����ʱ,����ȱʡ����
'      lngȱʡ����ID=��objCboʱ,Ϊȱʡ��λ�Ŀ��ң�����ΪҪ���ص�ȱʡ����
'      blnBed=�Ƿ�ֻȡ�д�λ�Ŀ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
        
    On Error GoTo errH
    
    If int��Χ = 1 Then
        strTmp = "1,3"
    ElseIf int��Χ = 2 Then
        strTmp = "2,3"
    ElseIf int��Χ = 3 Then
        strTmp = "1,2,3"
    End If
    
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,A.���� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And A.ID=B.����ID And Instr([1],','||B.�������||',')>0 And B.��������='�ٴ�'" & _
        IIf(lng���˿���ID <> 0, " And A.ID<>[2]", "") & _
        IIf(blnBed, " And Exists(Select ����ID From ��λ״����¼ Where ����ID=A.ID)", "") & _
        " Order by A.����"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", "," & strTmp & ",", lng���˿���ID)
    
    If Not objCbo Is Nothing Then
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            objCbo.AddItem rsTmp!���� & "-" & rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
            If rsTmp!ID = lngȱʡ����ID Then
                Call zlControl.CboSetIndex(objCbo.Hwnd, objCbo.NewIndex)
            End If
            rsTmp.MoveNext
        Next
    ElseIf Not rsTmp.EOF Then
        lngȱʡ����ID = rsTmp!ID
    End If
    Get�ٴ����� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Have��Ա����(str���� As String) As Boolean
'���ܣ��жϵ�ǰ��¼��Ա�Ƿ����ָ������Ա����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    strSQL = "Select A.ID" & _
        " From ��Ա�� A,��Ա����˵�� B,�ϻ���Ա�� C" & _
        " Where A.ID=B.��ԱID And B.��Ա����=[1]" & _
        " And A.ID=C.��ԱID And Upper(C.�û���)=Upper(User)" & _
        " And Rownum=1"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str����)
    Have��Ա���� = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get������ĿID(ByVal lngҽ��ID As Long, bln��ҩ���� As Boolean) As Long
'���ܣ���ȡָ��ҽ����������ĿID
'������lngҽ��ID=���IDΪNULL��ҽ��ID(��ҩ,�����Ŀ,��Ҫ����,��ҩ�÷�,������ҽ��)
'      bln��ҩ����=ҽ��ID�Ƿ�Ϊ��ҩ�÷���ɼ�������ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bln��ҩ���� Then
        strSQL = "Select ���,������ĿID From ����ҽ����¼ Where ������� IN('7','C') And ���ID=[1]"
    Else
        strSQL = "Select ���,������ĿID From ����ҽ����¼ Where ID=[1]"
    End If
    strSQL = strSQL & " Union ALL " & Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    strSQL = strSQL & " Order by ���"
    
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    If Not rsTmp.EOF Then Get������ĿID = Nvl(rsTmp!������ĿID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcAdvicePrice(ByVal lngҽ��ID As Long, Optional ByVal str�ѱ� As String, _
    Optional ByVal bln�������� As Boolean, Optional ByVal dbl���� As Double) As Double
'���ܣ�����ָ����ҩƷҽ��Ҫ���͵��ܽ��,�����¼۸����
'������str�ѱ�=�Ƿ񰴷ѱ������۵Ľ��
'      dbl����=���͵�����,���ѱ����ʱ��Ҫ(���ۼ۵�λ),��ʱ�������ʵ�ս��
'      bln��������=�Ƿ񸽼�����ҽ���ļƼ�
'      gbln�Ӱ�Ӽ�=����ʱ���������,�����ط���ΪFalse
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl���� As Double, dblӦ�� As Double, dblʵ�� As Double
    Dim dbl��� As Double, lngִ�п���ID As Long
    Dim lng������ID As Long, blnHaveSub As Boolean
    
    On Error GoTo errH
    
    If str�ѱ� = "" And dbl���� = 0 Then dbl���� = 1 '���Է�������,����ҽ������
    
    strSQL = _
        " Select M.��ҳID,M.���˿���ID,M.ִ�п���ID,C.ID,C.���,C.�Ƿ���,D.��������,B.������ĿID,A.����," & _
        " Nvl(A.����,0) as ����,C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.�����շ���,Decode(Nvl(C.�Ƿ���,0),1,A.����,B.�ּ�) as ����" & _
        " From ����ҽ����¼ M,����ҽ���Ƽ� A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D" & _
        " Where A.�շ�ϸĿID=B.�շ�ϸĿID And B.�շ�ϸĿID=C.ID" & _
        " And C.ID=D.����ID(+) And M.ID=A.ҽ��ID And M.ID=[1]" & _
        " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
        " Order by ����"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    If gbln��������ۿ� And Not rsTmp.EOF And str�ѱ� <> "" Then
        rsTmp.Filter = "����=1"
        If Not rsTmp.EOF Then blnHaveSub = True
        rsTmp.Filter = 0
    End If
    For i = 1 To rsTmp.RecordCount
        If InStr(",5,6,7,", rsTmp!���) > 0 And Nvl(rsTmp!�Ƿ���, 0) = 1 Then
            '�趨�ļƼ���ʱ��ҩƷ���ۼ���
            lngִ�п���ID = Get�շ�ִ�п���ID(rsTmp!���, rsTmp!ID, 4, Nvl(rsTmp!���˿���ID, 0), IIf(Not IsNull(rsTmp!��ҳID), 2, 1))
            dbl���� = Format(CalcDrugPrice(rsTmp!ID, lngִ�п���ID, dbl���� * Nvl(rsTmp!����, 0), , True), "0.00000")
        ElseIf rsTmp!��� = "4" And Nvl(rsTmp!�Ƿ���, 0) = 1 And Nvl(rsTmp!��������, 0) = 1 Then
            '�趨�ļƼ���ʱ�����ĵ��ۼ���
            lngִ�п���ID = Get�շ�ִ�п���ID(rsTmp!���, rsTmp!ID, 4, Nvl(rsTmp!���˿���ID, 0), IIf(Not IsNull(rsTmp!��ҳID), 2, 1), Nvl(rsTmp!ִ�п���ID, 0))
            dbl���� = Format(CalcDrugPrice(rsTmp!ID, lngִ�п���ID, dbl���� * Nvl(rsTmp!����, 0), , True), "0.00000")
        Else
            dbl���� = Format(Nvl(rsTmp!����, 0), "0.00000")
        End If
        
        '����Ӧ�ս��
        dblӦ�� = dbl���� * Nvl(rsTmp!����, 0) * dbl����
        If bln�������� Then
            dblӦ�� = dblӦ�� * Nvl(rsTmp!�����շ���, 100) / 100
        End If
        If gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
            dblӦ�� = dblӦ�� * (1 + Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100)
        End If
        
        '����ʵ�ս��
        If str�ѱ� = "" Then
            dblӦ�� = Format(dblӦ��, "0.00000")
            dblʵ�� = dblӦ��
        Else
            dblӦ�� = Format(dblӦ��, gstrDec)
        
            If gbln��������ۿ� And blnHaveSub Then
                If rsTmp!���� = 0 And lng������ID = 0 Then lng������ID = rsTmp!������ĿID
                dblʵ�� = dblӦ��
            Else
                dblʵ�� = Format(ActualMoney(str�ѱ�, rsTmp!������ĿID, CCur(dblӦ��)), gstrDec)
            End If
        End If
        
        dbl��� = dbl��� + dblʵ��
        
        rsTmp.MoveNext
    Next
    
    '�ײ����������
    If gbln��������ۿ� And blnHaveSub And lng������ID <> 0 And str�ѱ� <> "" Then
        dbl��� = Format(ActualMoney(str�ѱ�, lng������ID, CCur(dbl���)), gstrDec)
    End If
    
    CalcAdvicePrice = dbl���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDrugInfo(lngҩ��ID As Long, lngҩƷID As Long, lngҩ��ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal blnͣ�� As Boolean = True) As ADODB.Recordset
'���ܣ���ȡָ��ҩƷ�����Ϣ
'������int��Χ=1-����,2-סԺ(ȱʡ)
'      blnͣ��=�Ƿ��ȡ��ͣ��ҩƷ,���ڳ���ҩƷ���ʹ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strTmp = IIf(int��Χ = 1, "����", "סԺ")
    
    strSQL = _
        " Select ҩƷID,Sum(Nvl(��������,0)) as ��� From ҩƷ���" & _
        " Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))" & _
        " And ���� = 1 And �ⷿID=[1]" & IIf(lngҩƷID <> 0, " And ҩƷID=[2]", "") & _
        " Group by ҩƷID Having Sum(Nvl(��������,0))<>0"
    strSQL = "Select A.ҩƷID,A.����ϵ��,A." & strTmp & "��װ,A." & strTmp & "��λ,A.�ɷ����," & _
        " A.ҩ������,B.�Ƿ���,C.���/A." & strTmp & "��װ as ���,B.����,Nvl(D.����,B.����) as ����,B.���,B.����,B.����ʱ��,B.�������" & _
        " From ҩƷ��� A,�շ���ĿĿ¼ B,(" & strSQL & ") C,�շ���Ŀ���� D" & _
        " Where A.ҩƷID=B.ID And A.ҩƷID=C.ҩƷID(+)" & _
        " And B.ID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=[5]" & _
        IIf(blnͣ��, " And B.������� IN([3],3) And (B.����ʱ�� is NULL Or B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))", "") & _
        " And A.ҩ��ID=[4]" & IIf(lngҩƷID <> 0, " And A.ҩƷID=[2]", "") & _
        " Order by B.����"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҩ��ID, lngҩƷID, int��Χ, lngҩ��ID, IIf(gbln��Ʒ��, 3, 1))
    Set GetDrugInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function NextNo(intBillID As Integer) As Variant
'���ܣ������ض���������µĺ���,�������£�
'   һ����Ŀ��ţ�
'   1   ����ID         ����
'   2   סԺ��         ����
'   3   �����         ����
'   10  ҽ�����ͺ�     ����,˳��������
'   x   �������ݺ�     �ַ�,���ݱ�Ź���˳��������,���Զ���ȱ
'   �������λȷ��ԭ��:
'       ��1990Ϊ���������������������0��9/A��Z��˳����Ϊ��ȱ���

    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim vntNo As Variant, strSQL As String
    Dim intYear, strYear As String
    Dim curDate As Date, blnByDate As Boolean
ReStart:
    Err = 0
    On Error GoTo errHand

    If intBillID = 1 Then '����ID
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From ������Ʊ� Where ��Ŀ���=" & intBillID
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!������), 0, !������)
            
            strSQL = "Select Nvl(Max(����ID),0)+1 as ����ID From ������Ϣ Where ����ID>=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
            With rsTmp
                'If .State = adStateOpen Then .Close
                'Call SQLTest(App.ProductName, "NextNo", strSQL)
                '.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                'Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            .Update "������", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 2 Then 'סԺ��
        '˳���Ż������ڱ��
        strSQL = "Select A.*,Sysdate as ���� From ϵͳ������ A Where A.������=27"
        With rsTmp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            Call SQLTest
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!����ֵ), 1, !����ֵ) = 2)
                curDate = !����
            End If
        End With
        
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From ������Ʊ� Where ��Ŀ���=" & intBillID
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!������), 0, !������)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(סԺ��),0)+1 as סԺ�� From ������Ϣ Where סԺ��>=[1]"
            Else
                strSQL = "Select Nvl(Max(סԺ��),To_Number(To_Char(Sysdate,'YYMM')||'0000'))+1 as סԺ��" & _
                    " From ������Ϣ Where סԺ�� Like To_Number(To_Char(Sysdate,'YYMM'))||'%' And סԺ��>=[1]"
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
            With rsTmp
                'If .State = adStateOpen Then .Close
                'Call SQLTest(App.ProductName, "NextNo", strSQL)
                '.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                'Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "������", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "������", IIf(vntNo - 10 > Val(Format(curDate, "YYMM0000")), vntNo - 10, Val(Format(curDate, "YYMM0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 3 Then '�����
        '˳���Ż������ڱ��
        strSQL = "Select A.*,Sysdate as ���� From ϵͳ������ A Where A.������=46"
        With rsTmp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            Call SQLTest
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!����ֵ), 1, !����ֵ) = 2)
                curDate = !����
            End If
        End With
    
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From ������Ʊ� Where ��Ŀ���=" & intBillID
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!������), 0, !������)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(�����),0)+1 as ����� From ������Ϣ Where �����>=[1]"
            Else
                strSQL = "Select Nvl(Max(�����),To_Number(To_Char(Sysdate,'YYMMDD')||'0000'))+1 as �����" & _
                    " From ������Ϣ Where ����� Like To_Number(To_Char(Sysdate,'YYMMDD'))||'%' And �����>=[1]"
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
            With rsTmp
                'If .State = adStateOpen Then .Close
                'Call SQLTest(App.ProductName, "NextNo", strSQL)
                '.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                'Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "������", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "������", IIf(vntNo - 10 > Val(Format(curDate, "YYMMDD0000")), vntNo - 10, Val(Format(curDate, "YYMMDD0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 10 Then 'ҽ�����ͺ�
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From ������Ʊ� C Where C.��Ŀ���=" & intBillID
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            vntNo = Val(IIf(IsNull(!������), 0, !������)) + 1
            
            On Error Resume Next
            .Update "������", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    Else
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From ������Ʊ� C Where C.��Ŀ���=" & intBillID
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            intYear = Format(!Today, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIf(IsNull(!������), "", !������)
            
            If IIf(IsNull(!��Ź���), 0, !��Ź���) = 1 Then
                '����˳����
                If vntNo < strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000" Then
                    vntNo = strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000"
                End If
                vntNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + 1), 4)
            Else
                '����˳����
                If Left(vntNo, 1) < strYear Then
                    vntNo = strYear & "0000000"
                End If
                vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
            End If
            
            If Not (UCase(strYear) >= "A" And UCase(strYear) <= "Z") Or zlCommFun.ActualLen(vntNo) > 8 Then GoTo ReStart
            
            On Error Resume Next
            .Update "������", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    NextNo = Null
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'���ܣ��ж��Ƿ����ָ�������������͵�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���ID From ҩƷ�������� Where ����=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!���ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng����id As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����id)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetClinicBillID(ByVal lng��ĿID As Long, ByVal int���� As Integer) As Long
'���ܣ���ȡ������Ŀ��Ӧ�����Ƶ���(���ܸ���,�������ɷ���NO)
'������int����=1-����,2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select �����ļ�ID From ���Ƶ���Ӧ�� Where ������ĿID=[1] And Ӧ�ó���=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID, int����)
    If Not rsTmp.EOF Then GetClinicBillID = Nvl(rsTmp!�����ļ�ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptIsWoman(ByVal lng����ID As Long) As Boolean
'���ܣ��ж�ָ�������Ƿ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select * From ��������˵�� Where ��������='����' And ����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����ID)
    DeptIsWoman = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng����id As Long, ByVal lng��ҳID As Long) As String
'���ܣ���鲡����ҽ�������Ƿ���δִ�����(δִ�л�����ִ��)����Ŀ
'���أ�ҽ����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    strSQL = _
        " Select Distinct A.ID From ���ű� A,��������˵�� B" & _
        " Where B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� IN('���','����','����','����','Ӫ��')" & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
    strSQL = "Select C.���� as ��Ŀ,D.���� as ����,B.ִ��״̬" & _
        " From ����ҽ����¼ A,����ҽ������ B,������ĿĿ¼ C,���ű� D" & _
        " Where A.����ID=[1] And Nvl(A.��ҳID,0)=[2]" & _
        " And B.ҽ��ID=A.ID And B.ִ�в���ID+0 IN(" & strSQL & ")" & _
        " And B.ִ��״̬ IN(0,3) And A.������ĿID=C.ID And B.ִ�в���ID=D.ID" & _
        " And Not (A.������� IN('F','G','D') And A.���ID is Not NULL)" & _
        " And Not (A.�������='Z' And Nvl(C.��������,'0')<>'0')"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����id, lng��ҳID)
    strSQL = ""
    For i = 1 To rsTmp.RecordCount
        If i > 10 Then
            strSQL = strSQL & vbCrLf & "... ..."
            Exit For
        Else
            strSQL = strSQL & vbCrLf & rsTmp!��Ŀ & "����" & rsTmp!���� & Decode(Nvl(rsTmp!ִ��״̬, 0), 0, "δִ��", 3, "����ִ��")
        End If
        rsTmp.MoveNext
    Next
    ExistWaitExe = Mid(strSQL, 3)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitDrug(ByVal lng����id As Long, ByVal lng��ҳID As Long) As String
'���ܣ���鲡����ҩ���Ƿ���δ��ҩ��ҩƷ
'���أ�ҩ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    '��ҩƷ�շ���¼�д���δ��ҩƷΪ׼
    strSQL = "Select Distinct C.���� as ҩ��" & _
        " From ���˷��ü�¼ A,ҩƷ�շ���¼ B,���ű� C" & _
        " Where A.NO=B.NO And B.�ⷿID+0=C.ID(+) And A.�շ���� IN('5','6','7')" & _
        " And B.���� IN(9,10) And Mod(B.��¼״̬,3)=1 And B.����� IS NULL" & _
        " And A.����ID=[1] And Nvl(A.��ҳID,0)=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����id, lng��ҳID)
    strSQL = ""
    For i = 1 To rsTmp.RecordCount
        strSQL = strSQL & "," & Nvl(rsTmp!ҩ��, "[δ��ҩ��]")
        rsTmp.MoveNext
    Next
    ExistWaitDrug = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicePause(ByVal lngҽ��ID As Long) As String
'���ܣ���ȡָ��ҽ������ͣʱ��μ�¼
'���أ�"��ͣʱ��,��ʼʱ��;...."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    strSQL = "Select ��������,����ʱ�� From ����ҽ��״̬" & _
        " Where �������� IN(6,7) And ҽ��ID=[1]" & _
        " Order by ����ʱ��"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!�������� = 6 Then
            strTmp = strTmp & ";" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & ","
        ElseIf rsTmp!�������� = 7 Then
            '���õ���һ�벻����ͣ�ķ�Χ֮��
            strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
        End If
        rsTmp.MoveNext
    Next
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub PrintDiagReport(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, objParent As Object, Optional ByVal PrtMode As Integer = 2, Optional objPic As Object = Nothing, Optional blnMoved As Boolean = False)
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
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, lng���ͺ�)
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
                        Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("�û���1")), Nvl(rsTmp("����1"))
                    End If
                    
                    If strDeviceNO2 <> rsTmp("�豸��2") Then
                        strDeviceNO2 = rsTmp("�豸��2")
                        Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("�û���2")), Nvl(rsTmp("����2"))
                    End If
                    
                    If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
                        Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
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
                        Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("�û���1")), Nvl(rsTmp("����1"))
                    End If
                    
                    If strDeviceNO2 <> rsTmp("�豸��2") Then
                        strDeviceNO2 = rsTmp("�豸��2")
                        Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("�û���2")), Nvl(rsTmp("����2"))
                    End If
                    
                    'Inet.strIPAddress = Nvl(rsTmp(2)): Inet.strUser = Nvl(rsTmp(6)): Inet.strPsw = Nvl(rsTmp(7))
                    If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
    '                    strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
    '                        Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
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
        
        If Not objPic Is Nothing Then
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
            'If rsTmp.State <> adStateClosed Then rsTmp.Close
            If blnMoved Then
                strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, lng���ͺ�)
            iFlagCount = rsTmp.RecordCount
            objPic.Cls
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = strTempPath & objFileSystem.GetTempName
                
                '���������ߴ�
                On Error Resume Next
                Set objPic.Picture = ReadCaseMap(rsTmp(2))
                objPic.Width = objPic.ScaleX(objPic.Picture.Width, vbHimetric, vbTwips): objPic.Height = objPic.ScaleY(objPic.Picture.Height, vbHimetric, vbTwips)
                If objPic.Width / objPic.Height > rsTmp(4) / rsTmp(5) Then
                    objPic.Width = objPic.Height * rsTmp(4) / rsTmp(5)
                Else
                    objPic.Height = objPic.Width / (rsTmp(4) / rsTmp(5))
                End If
                objPic.Cls: Set objPic.Picture = Nothing
                On Error GoTo DBError
                Call ShowMapInOjbect(objPic, rsTmp(2), rsTmp(3), blnMoved:=blnMoved)
                SavePicture objPic.Image, strTmpFile
                objPic.Cls
            
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
                "ReportFormat=" & gintReportFormat, _
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
                "ReportFormat=" & gintReportFormat, _
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
'        For i = 0 To iFlagCount - 1
'            objFileSystem.DeleteFile aFlagImages(1, i), True
'        Next
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub PrintDiagRpt_New(ByVal lng����ID As Long, objParent As Object, Optional ByVal PrtMode As Integer = 2, Optional objPic As Object = Nothing, Optional ByVal blnMoved As Boolean)
'���ܣ���ӡ���ﱨ��
'������blnMoved=�ò������������Ƿ���ת��
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strRptName As String
    Dim aImages(1, 8) As Variant, aFlagImages(1, 8) As Variant, i As Integer
    Dim strTempPath As String, lngBuffSize As Long
    Dim intReportFormatItem As Integer
    
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strNO As String, lng��¼���� As Long, lngҽ��ID As Long, lng���ͺ� As Long
    Dim iTmpFileCount As Integer, iFlagCount As Integer
    
    On Error GoTo DBError
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    strTempPath = Mid(strTempPath, 1, InStrRev(strTempPath, "\"))
    
    strSQL = "Select A.ҽ��ID,A.���ͺ�,A.NO,A.��¼����,'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' As ������,D.���ID,E.���,E.��������" & _
        " From ����ҽ������ A,���˲�����¼ B,�����ļ�Ŀ¼ C,����ҽ����¼ D,������ĿĿ¼ E" & _
        " Where A.����ID=B.ID And B.�ļ�ID=C.ID And A.ҽ��ID=D.ID And D.������ĿID=E.ID And A.����ID=[1] Order By D.���ID Desc"
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����ID)
    If rsTmp.EOF Then
        MsgBox "������δ��д���棬���ܴ�ӡ��", vbInformation, gstrSysName
    Else
        lngҽ��ID = rsTmp("ҽ��ID"): lng���ͺ� = rsTmp("���ͺ�")
        strRptName = rsTmp("������"): strNO = rsTmp("NO"): lng��¼���� = rsTmp("��¼����")
        If PrtMode = 2 Then
            If Not ReportPrintSet(gcnOracle, glngSys, strRptName, objParent) Then Exit Sub
            intReportFormatItem = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\zl9Report\LocalSet\" & strRptName, "Format", 1)
        Else
            intReportFormatItem = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\zl9Report\frmReport" & strRptName, "��ʽ", 1)
        End If
        
        '����
        If Nvl(rsTmp("���") = "E") And Nvl(rsTmp("��������")) = "6" Then
            strSQL = "Select A.ҽ��ID,A.���ͺ�,A.NO,A.��¼����" & _
                " From ����ҽ������ A,����걾��¼ C,������Ŀ�ֲ� D" & _
                " Where D.�걾ID+0=C.ID And C.ҽ��ID+0=A.ҽ��ID And D.ҽ��ID=[1]"
            If blnMoved Then
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
                strSQL = Replace(strSQL, "����걾��¼", "H����걾��¼")
                strSQL = Replace(strSQL, "������Ŀ�ֲ�", "H������Ŀ�ֲ�")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
            If Not rsTmp.EOF Then
                lngҽ��ID = rsTmp("ҽ��ID"): lng���ͺ� = rsTmp("���ͺ�")
                strNO = rsTmp("NO"): lng��¼���� = rsTmp("��¼����")
            End If
        End If
        
        strSQL = "Select B.���,B.����,A.�û���1,A.����1,A.Host1,A.Root1,A.URL1,A.�û���2,A.����2,A.Host2,A.Root2,A.URL2," & _
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
            " Order By A.���) A," & _
            " (select z.���,z.����,rownum as seq " & _
            " from " & _
            " (Select A.���,B.����,Rownum As Seq" & _
            " From zlReports A,zlRPTItems B,�����ļ�Ŀ¼ C,����ҽ����¼ D,���Ƶ���Ӧ�� E" & _
            " Where A.ID=B.����ID And A.���='ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2'" & _
            " And C.ID=E.�����ļ�ID And D.������ĿID=E.������ĿID And Nvl(B.����,0)=1 And B.����=11" & _
            " And E.Ӧ�ó���=D.������Դ And D.ID=[1]" & _
            " And B.���� Not Like '���%' and b.��ʽ��=[3]" & _
            " Order BY b.���� ) z ) B Where A.Seq=B.Seq"
        If blnMoved Then
            strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
            strSQL = Replace(strSQL, "���˲����ⲿͼ", "H���˲����ⲿͼ")
        End If
        'If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, lng���ͺ�, intReportFormatItem)
        iTmpFileCount = rsTmp.RecordCount
        If PrtMode = 2 Then
            If Not ReportPrintSet(gcnOracle, glngSys, strRptName, objParent) Then Exit Sub
        End If
        For i = 0 To rsTmp.RecordCount - 1
            If i > 8 Then Exit For
            
            strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & objFileSystem.GetParentFolderName(rsTmp("URL1"))
            strTmpFile = Replace(strTmpFile, "/", "\")
            MkLocalDir strTmpFile
            strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(rsTmp("URL1"))
            
            If Dir(strTmpFile, vbDirectory) = "" Then
                If strDeviceNO1 <> rsTmp("�豸��1") Then
                    strDeviceNO1 = rsTmp("�豸��1")
                    Inet1.FuncFtpConnect rsTmp("Host1"), rsTmp("�û���1"), rsTmp("����1")
                End If
                
                If strDeviceNO2 <> rsTmp("�豸��2") Then
                    strDeviceNO2 = rsTmp("�豸��2")
                    Inet2.FuncFtpConnect rsTmp("Host2"), rsTmp("�û���2"), rsTmp("����2")
                End If
                
                'Inet.strIPAddress = Nvl(rsTmp(2)): Inet.strUser = Nvl(rsTmp(6)): Inet.strPsw = Nvl(rsTmp(7))
                If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
'                    strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
'                        Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                    Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
                End If
            End If
            
            aImages(0, i) = rsTmp(1)
            aImages(1, i) = strTmpFile
            rsTmp.MoveNext
        Next
        For i = rsTmp.RecordCount To 8
            aImages(0, i) = "1"
            aImages(1, i) = "1"
        Next
        If Not objPic Is Nothing Then
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
            If blnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
                strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
            End If
            'If rsTmp.State <> adStateClosed Then rsTmp.Close
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, lng���ͺ�)
            iFlagCount = rsTmp.RecordCount
            objPic.Cls
             For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = strTempPath & objFileSystem.GetTempName
                
                '���������ߴ�
                On Error Resume Next
                Set objPic.Picture = ReadCaseMap(rsTmp(2))
                objPic.Width = objPic.ScaleX(objPic.Picture.Width, vbHimetric, vbTwips): objPic.Height = objPic.ScaleY(objPic.Picture.Height, vbHimetric, vbTwips)
                If objPic.Width / objPic.Height > rsTmp(4) / rsTmp(5) Then
                    objPic.Width = objPic.Height * rsTmp(4) / rsTmp(5)
                Else
                    objPic.Height = objPic.Width / (rsTmp(4) / rsTmp(5))
                End If
                objPic.Cls: Set objPic.Picture = Nothing
                On Error GoTo DBError
                Call ShowMapInOjbect(objPic, rsTmp(2), rsTmp(3), blnMoved:=blnMoved)
                SavePicture objPic.Image, strTmpFile
                objPic.Cls
            
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
        For i = 0 To iTmpFileCount - 1
            objFileSystem.DeleteFile aImages(1, i), True
        Next
        For i = 0 To iFlagCount - 1
            objFileSystem.DeleteFile aFlagImages(1, i), True
        Next
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ItemIsVarPrice(ByVal lng��ĿID As Long) As Boolean
'���ܣ��ж�ָ����Ŀ�Ƿ���(��ҩƷ�͸������õ�����)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select A.���,A.�Ƿ���,B.�������� From �շ���ĿĿ¼ A,�������� B Where A.ID=B.����ID(+) And A.ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
    If Not rsTmp.EOF Then
        If Not (InStr(",5,6,7,", rsTmp!���) > 0 Or rsTmp!��� = "4" And Nvl(rsTmp!��������, 0) = 1) Then
            ItemIsVarPrice = Nvl(rsTmp!�Ƿ���, 0) <> 0
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlClinicCodeRepeat(str���� As String, Optional lng��ĿID As Long) As Boolean
'���ܣ����������Ŀ������Ƿ������б����ظ����ظ��������ʾ
'��Σ�str����-����ı��룻lng��ĿID-�Լ���ID�ţ����޸�ʱ����Ҫ��������������ж�
'���Σ��ظ�����True��������Flase
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select K.����||' ['||I.����||']'||I.���� as ����" & _
        " From ������ĿĿ¼ I,������Ŀ��� K" & _
        " Where I.���=K.���� And I.����=[1] And I.ID<>[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str����, lng��ĿID)
    If Not rsTmp.EOF Then
        MsgBox "����Ŀ�����롰" & rsTmp!���� & "���ı����ظ���", vbInformation, gstrSysName
        zlClinicCodeRepeat = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetSymbol(StrInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & StrInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & StrInput & "') from dual"
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

Public Function GetSendMoneyState(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, Optional ByVal blnMoved As Boolean) As String
'���ܣ���ȡָ��ҽ��ĳ�η���֮��ļƷ�״̬����Ҫ����һЩ���ҽ���ж��ּƷѵ�״̬
'������lngҽ��ID=�����������Ŀ,��������Ŀ,��һ��������Ŀ��ҽ��ID(����ҽ��վ����ʾ����Ŀ��)
'���أ�",-1,0,1,"������-1=����Ʒ�,1=�ѼƷ�,0=δ�Ʒ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���ID From ����ҽ����¼ Where �������='C' And ID=[1]"
    strSQL = _
        " Select ID From ����ҽ����¼ Where ID=[1] Or (���ID=[1] And ������� IN('F','D'))" & _
        " Union ALL " & _
        " Select ID From ����ҽ����¼ Where �������='C' And ���ID=(" & strSQL & ")"
    strSQL = "Select Distinct �Ʒ�״̬ From ����ҽ������ Where ҽ��ID IN(" & strSQL & ") And ���ͺ�=[2]"
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, lng���ͺ�)
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & Nvl(rsTmp!�Ʒ�״̬, 0)
        rsTmp.MoveNext
    Loop
    If strSQL <> "" Then GetSendMoneyState = strSQL & ","
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillMax���(ByVal strNO As String, ByVal int��¼���� As Integer, str�Ǽ�ʱ�� As String) As Integer
'���ܣ���ȡָ�����ݵ�ǰ��������+1
'������str�Ǽ�ʱ��=���ҽ��ֻ�����˲���������ʱ����Ҫ�����ɵ��շѻ��۵�(NO��ͬ)��ʱ���������ɵ�һ�¡�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    str�Ǽ�ʱ�� = ""
    strSQL = "Select Max(���) as ���,Max(�Ǽ�ʱ��) as ʱ�� From ���˷��ü�¼ Where NO=[1] And ��¼����=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strNO, int��¼����)
    If Not rsTmp.EOF Then
        GetBillMax��� = Nvl(rsTmp!���, 0) + 1
        If Not IsNull(rsTmp!ʱ��) Then
            str�Ǽ�ʱ�� = Format(rsTmp!ʱ��, "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        GetBillMax��� = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistBalance(ByVal strNO As String) As Boolean
'���ܣ��ж�ָ�����շѻ��۵��Ƿ�����Ѿ��շѵ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From ���˷��ü�¼ Where ��¼����=1 And ��¼״̬ IN(1,3) And NO=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRegistRoom(ByVal str�ű� As String, ByVal lng�ű�ID As Long, ByVal int���� As Integer) As String
'���ܣ����ݺű�ķ��﷽ʽ��ȡ�ű������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If int���� = 0 Then Exit Function '������
    
    On Error GoTo errH
    
    '�������
    If int���� = 1 Then
        'ָ������
        strSQL = "Select �������� From �ҺŰ������� Where �ű�ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng�ű�ID)
        If Not rsTmp.EOF Then GetRegistRoom = rsTmp!��������
    ElseIf int���� = 2 Then
        '��̬����ø��ű���Һ�δ�������ٵ�����   //todoδ����ԤԼ�Һ�
        strSQL = _
            " Select ��������,Sum(NUM) as NUM From (" & _
                " Select ��������,0 as NUM From �ҺŰ������� Where �ű�ID=[1]" & _
                " Union ALL" & _
                " Select ����,Count(����) as NUM From ���˹Һż�¼" & _
                " Where Nvl(ִ��״̬,0)=0" & _
                " And �Ǽ�ʱ�� Between Trunc(Sysdate) And Sysdate And �ű�=[2]" & _
                " And ���� IN(Select �������� From �ҺŰ������� Where �ű�ID=[1])" & _
                " Group by ����)" & _
            " Group by �������� Order by Num"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng�ű�ID, str�ű�)
        If Not rsTmp.EOF Then GetRegistRoom = rsTmp!��������
    ElseIf int���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        strSQL = "Select * From �ҺŰ������� Where �ű�ID=" & lng�ű�ID
        Call OpenRecord(rsTmp, strSQL, "mdlPublic", adOpenStatic, adLockOptimistic) '��д��¼��
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If Nvl(rsTmp!��ǰ����, 0) = 1 Then
                    GetRegistRoom = rsTmp!��������
                    rsTmp!��ǰ���� = 0
                    
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!��ǰ���� = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '�����һ��ƽ������
            If GetRegistRoom = "" Then
                rsTmp.MoveFirst
                GetRegistRoom = rsTmp!��������
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!��ǰ���� = 1
                rsTmp.Update
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadCaseMap(lngID As Long) As StdPicture
'���ܣ����ݱ��ͼID����ͼ�ζ���
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ͼ�� From �������ͼ Where Ԫ��ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngID)
        
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
    Kill strFile
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptExist(ByVal str�������� As String, ByVal int������� As Integer) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ID From ��������˵�� Where ��������=[1] And ������� IN([2],3) And Rownum<2"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str��������, int�������)
    DeptExist = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExpend(ByVal strNO As String) As Boolean
'���ܣ��жϹҺŵ��Ƿ��Ѿ�������Ч�Һ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If gint�Һ����� = 0 Then Exit Function
        
    On Error GoTo errH
    
    '��ʱ����
    strSQL = "Select Sysdate-�Ǽ�ʱ�� as ��� From ���˹Һż�¼ Where NO=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strNO)
    If Not rsTmp.EOF Then
        BillExpend = rsTmp!��� > gint�Һ�����
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��
'������vDate=ʱ����ʱ��εĿ�ʼʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select �ϴ����� From zlDataMove Where ϵͳ=[1] And ���=1 And �ϴ����� is Not Null"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '�ϴ�����û��ʱ��,"<"�ж���ת��������һ��
        If vDate < rsTmp!�ϴ����� Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByNO(ByVal strNO As String, ByVal strTable As String, Optional ByVal strWhere As String) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��
'������vDate=ʱ����ʱ��εĿ�ʼʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select NO From H" & strTable & " Where NO=[1] And Rownum<2" & IIf(strWhere <> "", " And " & strWhere, "")
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strNO)
    If Not rsTmp.EOF Then
        MovedByNO = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedBySend(ByVal lngҽ��ID As Long, Optional ByVal lng���ͺ� As Long) As Boolean
'���ܣ����ĳ�η��͵�ҽ���еķ����Ƿ��Ѿ�ִ��������ת��
'������lng���ͺ�=��Ϊ����ҽ��ֻ��һ�η���,���Բ�����
'˵����1.��ҽ��δת��������£�ִ�л��˻����ϲ���ʱ�����������ת���ķ��ã����ֹ
'      2.����סԺ�����ж�η��͵������ֻ�жϵ�ǰҪ���˵����ҽ�����ͷ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1]"
    strSQL = "Select B.NO From ����ҽ������ A,H���˷��ü�¼ B" & _
        " Where A.��¼����=B.��¼���� And A.NO=B.NO" & _
        IIf(lng���ͺ� <> 0, " And A.���ͺ�+0=[2]", "") & _
        " And A.ҽ��ID IN(" & strSQL & ")" & _
        " Group by B.NO"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, lng���ͺ�)
    If Not rsTmp.EOF Then MovedBySend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getҽ�Ƹ�����(ByVal str���� As String) As String
'���ܣ�����ҽ�Ƹ��ʽ���ƻ�ȡҽ�Ƹ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If str���� = "" Then Exit Function
    
    strSQL = "Select ���� From ҽ�Ƹ��ʽ Where ����=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str����)
    If Not rsTmp.EOF Then Getҽ�Ƹ����� = Nvl(rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

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
        strSQL = "Select ����,Nvl(Ƹ�μ���ְ��,0) as ְ�� From ��Ա�� Where ����=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strҽ��)
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

Public Function PatiHaveOut(ByVal lng����id As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
'���ܣ��жϲ����Ƿ��ѳ�Ժ(����Ԥ��Ժ),���ڲ��������ж�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ����ID From ������ҳ Where (��Ժ���� is Not Null Or Nvl(״̬,0)=3) And ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng����id, lng��ҳID)
    If Not rsTmp.EOF Then PatiHaveOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlExpense", intNum)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '������
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'��һ������Ӱ���ļ�ת�Ƶ���������ȥ
Public Function MergeImageFiles(ByVal strCurrUID As String, ByVal strNewUID As String, _
    Optional ByVal strReceiveDate As String = "", Optional ByVal strMoveFiles As String = "") As Boolean
    
    Dim objSrcFtp As New clsFtp, objDestFtp As New clsFtp
    Dim strSrcPath As String, strDestPath As String
    Dim rsTmp As New ADODB.Recordset, strSQL As String, strTmpFile As String
    Dim aFiles() As String, i As Integer, objFile As New Scripting.FileSystemObject
    '�洢ԭ���UID��FTP����
    Dim strFTPUser As String, strFTPPassw As String, strFTPHost As String, strFTPRoot As String
    On Error GoTo errH
    MergeImageFiles = True
    If strCurrUID = strNewUID Then Exit Function
    
    '��ʼ��ԴFtp
    strSQL = "Select D.�û��� As FtpUser,D.���� As FtpPwd," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ�ö�,Null,C.λ��һ,C.λ�ö�)=D.�豸��(+)" & _
        "And C.���UID= [1] Union All " & _
        "Select D.�û��� As FtpUser,D.���� As FtpPwd," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ�ö�,Null,C.λ��һ,C.λ�ö�)=D.�豸��(+)" & _
        "And C.���UID= [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ZLPACSWork", strCurrUID)
    If rsTmp.EOF Then
        MergeImageFiles = False
        Exit Function
    End If
    
    '�洢ԭ��FTP��������
    strFTPHost = rsTmp("Host")
    strFTPPassw = rsTmp("FtpPwd")
    strFTPRoot = rsTmp("Root")
    strFTPUser = rsTmp("FtpUser")
    
    With objSrcFtp
        '.strIPAddress = Nvl(rsTmp("Host")): .strUser = Nvl(rsTmp("FtpUser")): .strPsw = Nvl(rsTmp("FtpPwd"))
        .FuncFtpConnect rsTmp("Host"), rsTmp("FtpUser"), rsTmp("FtpPwd")
        strSrcPath = rsTmp("Root") & Nvl(rsTmp("URL"))
    End With
    
    '��ʼ��Ŀ��Ftp,���Ŀ��UID�����ڣ�����һ����·��
    Set rsTmp = OpenSQLRecord(strSQL, "ZLPACSWork", strNewUID)
    If rsTmp.EOF Then
        If strReceiveDate <> "" Then
            With objDestFtp
                .FuncFtpConnect strFTPHost, strFTPUser, strFTPPassw
                strDestPath = strFTPRoot & Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
                '����FTPĿ¼
                .FuncFtpMkDir strFTPRoot, Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
            End With
        Else
            MergeImageFiles = False
            Exit Function
        End If
    Else
        With objDestFtp
    '        .strIPAddress = Nvl(rsTmp("Host")): .strUser = Nvl(rsTmp("FtpUser")): .strPsw = Nvl(rsTmp("FtpPwd"))
            .FuncFtpConnect rsTmp("Host"), rsTmp("FtpUser"), rsTmp("FtpPwd")
            strDestPath = rsTmp("Root") & Nvl(rsTmp("URL"))
        End With
    End If
    
    '��ȡ��Ҫ�ƶ����ļ���
    If strMoveFiles <> "" Then
        aFiles = Split(strMoveFiles, "|")
    Else
        aFiles = Split(objSrcFtp.FuncDirFiles(strSrcPath), "|")
    End If
    
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\" & aFiles(i)
        Call objSrcFtp.FuncDownloadFile(strSrcPath, strTmpFile, aFiles(i))
        Call objDestFtp.FuncUploadFile(strDestPath, strTmpFile, aFiles(i))
        
        Kill strTmpFile
        Call objSrcFtp.FuncDelFile(strSrcPath, aFiles(i))
    Next
    
    '��Ҫ����һ�£������ͼ����Ŀ¼�У�Ŀ¼�Ƿ��ɾ����
    Call objSrcFtp.FuncFtpDelDir(Replace(strSrcPath, strCurrUID, ""), strCurrUID)
    
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    MergeImageFiles = False
    Call SaveErrLog
End Function

Public Sub MkLocalDir(ByVal strDir As String)
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

Public Function GetAllImageFiles(ByVal CheckUID As String, _
    Optional ByVal strSerials As String = "", Optional ByVal blnMoved As Boolean = False, _
    Optional strFTPHost As String, Optional strDicomPath As String, Optional strLocalPath As String, _
    Optional strFTPUser As String, Optional strFtpPwd As String) As String()
    Dim strSQL As String, lngSeqUID As String
    Dim strURL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double
    Dim FrameCount As Integer
    Dim iCols As Integer, iRows As Integer
    Dim objFile As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim blnFirst As Boolean
    Dim strCachePath As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strGetFilesName As String
    
    Dim aFiles() As String
    
    Dim bln1stDev As Boolean
    bln1stDev = True
    ReDim Preserve aFiles(0) As String
    
    On Error GoTo DBError
    Screen.MousePointer = vbHourglass
    
    strFTPHost = "": strDicomPath = ""
    
    strCachePath = App.Path & "\TmpImage\"
    If Not objFileSystem.FolderExists(strCachePath) Then objFileSystem.CreateFolder strCachePath
            
    strSQL = "Select A.ͼ���,D.�û��� As User1,D.���� As Pwd1," & _
        "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As FtpPath1,d.�豸�� as �豸��1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/' As Path1,A.ͼ��UID As URL1, " & _
        "E.�û��� As User2,E.���� As Pwd2," & _
        "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As FtpPath2," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/' As Path2,A.ͼ��UID As URL2,e.�豸�� as �豸��2 " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And C.���UID=[1] "
    If Len(strSerials) = 0 Then
        strSQL = strSQL & "Order By A.ͼ���"
    Else
        strSQL = strSQL & "And A.����UID In(" & strSerials & ") Order By b.���к�,A.ͼ���"
    End If
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡӰ���ļ�", CheckUID)

    If rsTmp.RecordCount > 0 Then
        ClearCacheFolder strCachePath
        MkLocalDir strCachePath & rsTmp("Path1")
        blnFirst = True
        Do While Not rsTmp.EOF
            strFTPHost = "ftp://" & rsTmp("Host1"): strDicomPath = rsTmp("FtpPath1") & rsTmp("Path1")
            strLocalPath = rsTmp("Path1"): strFTPUser = Nvl(rsTmp("User1")): strFtpPwd = Nvl(rsTmp("Pwd1"))
            
'            If Dir(strCachePath & rsTmp("Path1") & rsTmp("URL1")) = vbNullString Then
                If strDeviceNO1 <> rsTmp("�豸��1") Then
                    strDeviceNO1 = rsTmp("�豸��1")
                    Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
                End If

                If strDeviceNO2 <> rsTmp("�豸��2") Then
                    strDeviceNO2 = rsTmp("�豸��2")
                    Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
                End If
                
                '---�ƽ��޸���2007-1-29----------
                blnFirst = False
                If rsTmp("�豸��1") <> "" Then
                    strTmpFile = strCachePath & rsTmp("Path1") & rsTmp("URL1")
                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
                    aFiles(UBound(aFiles)) = rsTmp("URL1")
                ElseIf rsTmp("�豸��2") <> "" Then
                    strTmpFile = strCachePath & rsTmp("Path2") & rsTmp("URL2")
                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
                    aFiles(UBound(aFiles)) = rsTmp("URL2")
                End If
                
                '---�ƽ��޸���2007-1-29-----����-----
                
                
                '--------������д---�ƽ��޸���2007-1-29----------------
'                blnFirst = False
'                strTmpFile = strCachePath & rsTmp("Path1") & rsTmp("URL1")
'                strGetFilesName = Inet1.FuncDirFiles(strLocalPath)
'                If InStr(1, strGetFilesName, rsTmp("URL1")) > 0 Then
'                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                    aFiles(UBound(aFiles)) = rsTmp("URL1")
'                Else
'                    strTmpFile = strCachePath & rsTmp("Path2") & rsTmp("URL2")
'                    strGetFilesName = Inet2.FuncDirFiles(strLocalPath)
'                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                    aFiles(UBound(aFiles)) = rsTmp("URL2")
'                End If
                '--------������д------�ƽ��޸���2007-1-29--------����-----
                
'                Inet.strIPAddress = Nvl(rsTmp("Host1")): Inet.strUser = Nvl(rsTmp("User1")): Inet.strPsw = Nvl(rsTmp("Pwd1"))
'                If Inet1.FuncDownloadFile(strDicomPath, strTmpFile, rsTmp("URL1")) <> 0 Then
'                    strFtpHost = "ftp://" & Nvl(rsTmp("Host2")): strDicomPath = rsTmp("FtpPath2") & rsTmp("Path2")
'                    strLocalPath = rsTmp("Path2"): strFtpUser = Nvl(rsTmp("User2")): strFtpPwd = Nvl(rsTmp("Pwd2"))
'
'                    strTmpFile = strCachePath & rsTmp("URL2")
''                    Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
'                    If Inet2.FuncDownloadFile(strDicomPath, strTmpFile, rsTmp("URL2")) = 0 Then
'                        ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                        aFiles(UBound(aFiles)) = rsTmp("URL1")
'                    End If
'                Else
'                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                    aFiles(UBound(aFiles)) = rsTmp("URL1")
'                End If
'            Else
'                ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                aFiles(UBound(aFiles)) = rsTmp("URL1")
'            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    Screen.MousePointer = vbDefault
    GetAllImageFiles = aFiles
    Exit Function

ReadURLError:
    If bln1stDev Then
        bln1stDev = False
        strFTPHost = rsTmp("Host2"): strDicomPath = rsTmp("FtpPath2") & rsTmp("Path2")
        Resume
    Else
        If ErrCenter() = 1 Then Resume
        Screen.MousePointer = vbDefault
        Call SaveErrLog
    End If
    Exit Function

DBError:
    If ErrCenter() = 1 Then Resume
    Screen.MousePointer = vbDefault
    Call SaveErrLog
End Function

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

'---����Ϊ����Ǽ���Ҫ
'-------------------------------------------------------------------------------------------------
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

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function InputIsCard(txtInput As Object, KeyAscii As Integer) As Boolean
'���ܣ��ж�ָ���ı����е�ǰ�����Ƿ���ˢ��,���ݴ���������ʾ
    Dim strText As String, blnCard As Boolean
    Dim arrMask As Variant, i As Long

    '��ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = UCase(strText & Chr(KeyAscii))
    End If
        
    '�ж��Ƿ���ˢ��
    blnCard = False
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then
        blnCard = True
    ElseIf gstrCardMask <> "" Then
        arrMask = Split(gstrCardMask, "|")
        For i = 0 To UBound(arrMask)
            If strText Like arrMask(i) & "*" Then
                If IsNumeric(Mid(strText, Len(arrMask(i)) + 1)) And IsNumeric(Mid(strText, Len(arrMask(i)) + 1, 1)) Then
                    blnCard = True
                End If
            End If
        Next
    End If
    
    'ˢ��ʱ�����Ƿ�������ʾ
    If blnCard Then
        txtInput.PasswordChar = IIf(gblnCardHide, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    
    InputIsCard = blnCard
End Function

Public Function GetSysParVal(Optional ByVal int������ As Integer = -9999, Optional ByVal strDefault As String) As String
'���ܣ���ȡָ��ϵͳ������ֵ
'������int������=Ϊ-9999ʱ����ʼ��������
'      strDefault=���û��ֵ��Ϊ�յ�ȱʡֵ
    Dim blnDo As Boolean, strSQL As String
    
    On Error GoTo errH
    
    blnDo = True
    If Not grsSysPars Is Nothing Then
        If grsSysPars.State = 1 Then blnDo = False
    End If
    If blnDo Then
        strSQL = "Select ������,������,����ֵ From ϵͳ������"
        Set grsSysPars = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(grsSysPars, strSQL, "GetSysParVal")
    End If
    
    If int������ <> -9999 Then
        grsSysPars.Filter = "������=" & int������
        If Not grsSysPars.EOF Then
            GetSysParVal = Nvl(grsSysPars!����ֵ, strDefault)
        Else
            GetSysParVal = strDefault
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetTrayHeight() As Long
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ĸ߶�
    '------------------------------------------------------------------------------------------------------------------
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
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

