Attribute VB_Name = "mdlTechCore"
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
Public gbln�Ӱ�Ӽ� As Boolean
Public grsSysPars As ADODB.Recordset
Public grsDuty As ADODB.Recordset '���ҽ��ְ��
Public gstr��̬�ѱ� As String '������ﵱǰ���ҿ��ö�̬�ѱ�,�ڹ���������ʹ��,ʹ��ʱ�Ÿ�ֵ:CalcDrugPrice,CalcPrice

'ҽ������
Public gclsInsure As New clsInsure

'CISϵͳ����
Public gblnҩƷ�������ҽ�� As Boolean
Public gint�����Ǽ���Ч���� As Integer
Public gbln����ҽ��������Ч As Boolean
Public gstr���ͻ��۵� As String
Public gblnҩ�ƻ��۵� As Boolean
Public gbln�������۵� As Boolean
Public gblnִ�к���� As Boolean
Public gintRXCount As Integer

'HISϵͳ����
Public gbln��ҽ As Boolean '�Ƿ�ʹ����ҽ����
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"

Public gint�Һ����� As Integer '�Һŵ���Ч����
Public gbln�շ���� As Boolean '�Ƿ������������
Public gbln��Ʒ�� As Boolean '����ҩ�Ƿ���Ʒ����ʾ
Public gblnסԺ�Զ����� As Boolean 'סԺ������ɺ��Ƿ��Զ�����
Public gbln�����Զ����� As Boolean '���������ɺ��Ƿ��Զ�����
Public gbln��������ۿ� As Boolean '������Ŀ���ܼ����ۿ�
Public gbln�����������۷��� As Boolean '���ʱ����������۷���
Public gbln�������Ҷ��� As Boolean '�����Ϳ����Ƿ��������
Public gint���Ʊ��� As Integer '0-˳����,1-����+�����+˳����
Public gint�����Դ As Integer '1-��ҽ��ѡ��������Դ,2-������ϱ�׼����,3-���ռ�����������
Public gint������� As Integer '1-������������,2-�����ݿ���ȡ����,3-��ҽ�����˴����ݿ�����
Public gstrMatchMode As String '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
Public gbyt���δִ�� As Byte '��Ժת��ʱ�Ƿ�����δִ����Ŀ��δ��ҩƷ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gintҽ������ As Integer '�Ƿ��סԺҽ�����˵���Ŀ����������м��:0-�����,1-��鲢����,2-��鲢��ֹ

'ҽ������վϵͳ���ò���
Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gblnStock As Boolean 'ָ��ҩ��ʱ�Ƿ��޶�����ҩƷ�Ŀ��
Public gstrҽ���������� As String 'ҽ����������ķ�������
Public gstr���ѷ������� As String '���Ѳ�������ķ�������

'����ǩ��
Public gintCA As Integer '����ǩ����֤����
Public gstrESign As String '����ǩ�����Ƴ���
Public gobjESign As Object '����ǩ���ӿڲ���

'-----------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ����ID As Long
    ������ As String
    ������ As String
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
    support�������� = 35            '�Ƿ����������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
End Enum
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

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
    Call SQLTest(App.ProductName, "mdlTechCore", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIF(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIF(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIF(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�û��� = IIF(IsNull(rsTmp!�û���), "", rsTmp!�û���)
        gstrDBUser = UserInfo.�û���
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
        strSQL = "Select Max(Length(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIF(lng����ID = 0, " is Not NULL", "=" & lng����ID)
    Else
        strSQL = "Select Max(Length(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIF(lng����ID = 0, " is Not NULL", "=" & lng����ID)
    End If
    
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlCISWork", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    If Not rsTmp.EOF Then GetMaxBedLen = IIF(IsNull(rsTmp!����), 0, rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ID(Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, _
    Optional ByVal lng����ID As Long) As Long
'���ܣ����ݿ���ID���˻�ȡ��Ӧ�Ĳ���ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng����ID <> 0 And lng��ҳID <> 0 Then
        strSQL = "Select ��ǰ����ID as ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID, lng��ҳID)
    Else
        strSQL = "Select ����ID From ��λ״����¼ Where ����ID=[1] Group by ����ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID)
    End If
    If Not rsTmp.EOF Then Get����ID = Nvl(rsTmp!����ID, 0)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
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

Public Function InitStockCheck(ByVal int��Χ As Integer, Optional ByVal bln���� As Boolean) As Collection
'���ܣ���ȡ��ͬ�ⷿ�����鷽ʽ�ڼ�����
'������int��Χ=1-����,2-סԺ
'      bln����=�Ƿ�������ĵķ��ϲ���
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
        " And B.�������� in('��ҩ��','��ҩ��','��ҩ��'" & IIF(bln����, ",'���ϲ���'", "") & ")" & _
        " And C.�ⷿID(+)=A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int��Χ)
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

Public Function GetSysParVal(Optional ByVal int������ As Integer = -9999, Optional ByVal strDefault As String) As String
'���ܣ���ȡָ��ϵͳ������ֵ
'������int������=Ϊ-9999ʱ����ʼ��������
'      strDefault=���û��ֵ��Ϊ�յ�ȱʡֵ
    Dim blnDo As Boolean, strSQL As String
    
    On Error GoTo errH
    
    blnDo = True
    If int������ <> -9999 Then
        If Not grsSysPars Is Nothing Then
            If grsSysPars.State = 1 Then blnDo = False
        End If
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

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    'HISϵͳ����
    '---------------------------------------------------------
    Call GetSysParVal
    
    '���ý��С����λ��
    gbytDec = 2: gstrDec = "0.00"
    grsSysPars.Filter = "������=9"
    If Not grsSysPars.EOF Then
        gbytDec = Val(Nvl(grsSysPars!����ֵ, 2))
        gstrDec = "0." & String(gbytDec, "0")
    End If
    
    'ָ��ҩ��ʱ���ƿ��
    grsSysPars.Filter = "������=18"
    If Not grsSysPars.EOF Then gblnStock = Nvl(grsSysPars!����ֵ, 0) <> 0
    
    '�Һ���Ч����
    grsSysPars.Filter = "������=21"
    If Not grsSysPars.EOF Then gint�Һ����� = Nvl(grsSysPars!����ֵ, 0)
    
    '���δִ����Ŀ
    grsSysPars.Filter = "������=22"
    If Not grsSysPars.EOF Then gbyt���δִ�� = Nvl(grsSysPars!����ֵ, 0)
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    grsSysPars.Filter = "������=23"
    If Not grsSysPars.EOF Then gbytBillOpt = Nvl(grsSysPars!����ֵ, 0)
    
    '����ǩ����֤����
    grsSysPars.Filter = "������=25"
    If Not grsSysPars.EOF Then gintCA = Val(Nvl(grsSysPars!����ֵ, "0"))
    
    '����ǩ�����Ƴ���
    grsSysPars.Filter = "������=26"
    If Not grsSysPars.EOF Then gstrESign = Nvl(grsSysPars!����ֵ)
    
    '�Ƿ�ʹ����ҽ
    grsSysPars.Filter = "������=31"
    If Not grsSysPars.EOF Then gbln��ҽ = Nvl(grsSysPars!����ֵ, 0) <> 0
    
    'ҽ����������
    grsSysPars.Filter = "������=41"
    If Not grsSysPars.EOF Then
        gstrҽ���������� = "'" & Replace(Nvl(grsSysPars!����ֵ), "|", "','") & "'"
    End If

    '���ѷ�������
    grsSysPars.Filter = "������=42"
    If Not grsSysPars.EOF Then
        gstr���ѷ������� = "'" & Replace(Nvl(grsSysPars!����ֵ), "|", "','") & "'"
    End If
    
    '���ﴦ����������
    grsSysPars.Filter = "������=56"
    If Not grsSysPars.EOF Then gintRXCount = Val(Nvl(grsSysPars!����ֵ, 0))

    'ҽ��������
    grsSysPars.Filter = "������=59"
    gintҽ������ = 1
    If Not grsSysPars.EOF Then gintҽ������ = Val(Nvl(grsSysPars!����ֵ, 1))

    '���Ʊ������ģʽ
    grsSysPars.Filter = "������=61"
    If Not grsSysPars.EOF Then gint���Ʊ��� = Val(Nvl(grsSysPars!����ֵ, 0))
    
    'סԺ�Զ�����
    grsSysPars.Filter = "������=63"
    If Not grsSysPars.EOF Then
        gblnסԺ�Զ����� = Nvl(grsSysPars!����ֵ, 0) <> 0
    End If
    
    'ҩƷ�������ҽ��
    grsSysPars.Filter = "������=69"
    If Not grsSysPars.EOF Then gblnҩƷ�������ҽ�� = Val(Nvl(grsSysPars!����ֵ, 0)) = 1
    
    'Ƥ�Խ����Чʱ��
    grsSysPars.Filter = "������=70"
    If Not grsSysPars.EOF Then gint�����Ǽ���Ч���� = Val(Nvl(grsSysPars!����ֵ, 0))
    
    '����ҽ��������Ч
    grsSysPars.Filter = "������=71"
    If Not grsSysPars.EOF Then gbln����ҽ��������Ч = Val(Nvl(grsSysPars!����ֵ, 0)) = 1
    
    '�Ƿ�Ҫ�������������
    grsSysPars.Filter = "������=72"
    If Not grsSysPars.EOF Then gbln�շ���� = Nvl(grsSysPars!����ֵ, 1) <> 0
    
    '����ҩ�Ƿ���Ʒ����ʾ
    grsSysPars.Filter = "������=74"
    If Not grsSysPars.EOF Then gbln��Ʒ�� = Nvl(grsSysPars!����ֵ, 0) <> 0
    
    'ҩ�����ɻ��۵�
    grsSysPars.Filter = "������=79"
    If Not grsSysPars.EOF Then gblnҩ�ƻ��۵� = Nvl(grsSysPars!����ֵ, 0) <> 0
    
    '�������ɻ��۵�
    grsSysPars.Filter = "������=80"
    If Not grsSysPars.EOF Then gbln�������۵� = Nvl(grsSysPars!����ֵ, 0) <> 0
    
    'ִ�к��Զ����
    grsSysPars.Filter = "������=81"
    If Not grsSysPars.EOF Then gblnִ�к���� = Nvl(grsSysPars!����ֵ, 0) <> 0
            
    '�����Զ�����
    grsSysPars.Filter = "������=92"
    If Not grsSysPars.EOF Then
        gbln�����Զ����� = Nvl(grsSysPars!����ֵ, 0) <> 0
    End If
    
    '������Ŀ���ܼ����ۿ�
    grsSysPars.Filter = "������=93"
    If Not grsSysPars.EOF Then gbln��������ۿ� = Nvl(grsSysPars!����ֵ, 0) <> 0
    
    '���ʱ����������۷���
    grsSysPars.Filter = "������=98"
    If Not grsSysPars.EOF Then gbln�����������۷��� = Nvl(grsSysPars!����ֵ, 0) <> 0
    
    '�����Ϳ����Ƿ��������
    grsSysPars.Filter = "������=99"
    If Not grsSysPars.EOF Then gbln�������Ҷ��� = Nvl(grsSysPars!����ֵ, 0) <> 0
    
    '����ǩ����ʼ��:ֻҪʹ�ü�����
    If gintCA <> 0 Then
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            Err.Clear: On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle)
            End If
        End If
    Else
        Set gobjESign = Nothing
    End If
    
    InitSysPar = True
End Function

Public Function GetPatiYear(lng����ID As Long) As Integer
'���ܣ���ȡ���˵�׼ȷ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intYear As Integer
    
    On Error GoTo errH
    
    strSQL = "Select Sysdate as ��ǰ,��������,���� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID)
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

Public Function GET��������(lngID As Long, Optional ByRef rs���� As ADODB.Recordset) As String
'���ܣ���ȡ��������
'������lngID=����ID
'���أ���������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If rs���� Is Nothing Then
        strSQL = "Select ���� from ���ű� Where ID=" & lngID
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlCISWork")
    Else
        Set rsTmp = rs����
        rsTmp.Filter = "ID=" & lngID
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select ���� from ���ű� Where ID=" & lngID
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlCISWork")
        End If
    End If
    
    If Not rsTmp.EOF Then GET�������� = rsTmp!����
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", CStr(int����), int��Դ)
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
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlCISWork")
    If Not rsTmp.EOF Then
        Check�ϰల�� = Nvl(rsTmp!Num, 0) > 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�շ�ִ�п���ID(ByVal lng����ID As Long, lng��ҳID As Long, _
    ByVal str��� As String, ByVal lng��ĿID As Long, ByVal intִ�п��� As Integer, _
    ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal lng���ϲ��� As Long) As Long
'���ܣ���ȡ��ҩ�շ���Ŀ��ִ�п���
'������int��Χ=1.����,2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strҩ�� As String, lngҩ�� As Long
    Dim bytDay As Byte, strIDs As String
    
    On Error GoTo errH
    
    If str��� = "4" Then
        '�г����з��ϲ���ʱ
'        strSQL = "Select B.�������,A.����,A.ID From ���ű� A,��������˵�� B" & _
'            " Where A.ID=B.����ID And B.��������='���ϲ���' And B.������� IN([1],3)" & _
'            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
'            " Order by B.�������,A.����"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int��Χ)
'        If Not rsTmp.EOF Then
'            If lng���ϲ��� <> 0 Then rsTmp.Filter = "ID=" & lng���ϲ���
'            If rsTmp.EOF Then rsTmp.Filter = 0
'            Get�շ�ִ�п���ID = rsTmp!ID
'        End If
    
        '��ִ�п�������ʱ
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int��Χ, lng���˿���ID, lng��ĿID)
        If Not rsTmp.EOF Then
            rsTmp.Filter = "��������ID=" & lng���˿���ID
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
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strҩ��, int��Χ, lng���˿���ID, lng��ĿID, bytDay)
        If Not rsTmp.EOF Then
            strIDs = ""
            rsTmp.Filter = "��������ID=" & lng���˿���ID
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
                Get�շ�ִ�п���ID = lng���˿���ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get�շ�ִ�п���ID = lng���˿���ID
                Else
                    Get�շ�ִ�п���ID = Get����ID(lng����ID, lng��ҳID)
                End If
            Case 3 '3-����Ա���ڿ���
                Get�շ�ִ�п���ID = UserInfo.����ID
            Case 4 '4-ָ������
                strSQL = "Select Nvl(��������ID,0) as ��������ID,ִ�п���ID" & _
                    " From �շ�ִ�п��� Where �շ�ϸĿID=[1]" & _
                    " And (������Դ is NULL Or ������Դ=[2])" & _
                    " And (��������ID is NULL Or ��������ID=[3])" & _
                    " Order by Decode(������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID, int��Χ, lng���˿���ID)
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "��������ID=" & lng���˿���ID
                    If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
                    If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                End If
            Case 6 '6-���������ڿ���
                Get�շ�ִ�п���ID = lng��������ID
        End Select
        If Get�շ�ִ�п���ID = 0 Then Get�շ�ִ�п���ID = UserInfo.����ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function




Public Function Get����ִ�п���ID(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal str��� As String, ByVal lng��ĿID As Long, ByVal lngҩƷID As Long, ByVal intִ�п��� As Integer, _
    ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, ByVal int��Ч As Integer, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal blnByȱʡ As Boolean) As Long
'���ܣ�����������Ŀִ�п�����Ϣ����ȱʡ��ִ�п���ID
'������lngҩƷID=ҩƷID,ȷ�������ʱҪ��
'      intִ�п���=��Ŀִ�п��ұ�־
'      lng���˿���ID=���˿���ID
'      lng��ҩ��,lng��ҩ��,lng��ҩ��=ҩƷȱʡҩ��,ҩƷ��ʱ��Ҫ
'      int��Ч=0-����,1-����,��������Ҫ�ж��ϰ�ʱ��
'      int��Χ=1-����,2-סԺ(ȱʡ)
'      blnByȱʡ=��ȡȱʡҩ��ʱ�����������ָ�����Ƿ񰴱���ȱʡָ����ҩ������û���򲻷���
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
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Χ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If int��Χ = 1 Then bln�ϰల�� = Check�ϰల��(True) 'סԺҽ������ҩ���ϰల��
        If Not bln�ϰల�� Then
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                 IIF(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[6]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                IIF(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
                " Order by B.�������,C.����"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strҩ��, int��Χ, lng���˿���ID, lngҩƷID, lng��ĿID, bytDay)
        If Not rsTmp.EOF Then
            If blnByȱʡ And lngҩ�� <> 0 Then
                rsTmp.Filter = "ִ�п���ID=" & lngҩ��
            Else
                Get����ִ�п���ID = rsTmp!ִ�п���ID
                rsTmp.Filter = "ִ�п���ID=" & lngҩ��
                If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            End If
            If Not rsTmp.EOF Then Get����ִ�п���ID = rsTmp!ִ�п���ID
        End If
    Else
        Select Case intִ�п���
            Case 0 '0-��ִ�еĶ���
                Exit Function
            Case 1 '1-�������ڿ���
                Get����ִ�п���ID = lng���˿���ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get����ִ�п���ID = lng���˿���ID
                Else
                    Get����ִ�п���ID = Get����ID(lng����ID, lng��ҳID)
                End If
            Case 3 '3-����Ա���ڿ���
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
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID, int��Χ, lng���˿���ID, bytDay)
                If Not rsTmp.EOF Then
                    Get����ִ�п���ID = rsTmp!ִ�п���ID
                    rsTmp.Filter = "��������ID=" & lng���˿���ID
                    If Not rsTmp.EOF Then Get����ִ�п���ID = rsTmp!ִ�п���ID
                End If
            Case 5 '5-Ժ��ִ��
                Exit Function
            Case 6 '6-���������ڿ���
                Get����ִ�п���ID = lng��������ID
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Have��������(ByVal lng����ID As Long, ByVal str���� As String) As Boolean
'���ܣ����ָ�������Ƿ����ָ����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select ����ID From ��������˵�� Where ����ID=[1] And ��������=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, str����)
    Have�������� = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Load��̬�ѱ�(lng����ID As Long) As String
'���ܣ�Ȩ��ָ�����Ҷ�ȡ��ǰ��Ч�Ķ�̬�ѱ�(Ŀǰֻ��������)
'���أ��ѱ�="���˽�,��һ��"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select ����,����,���� From �ѱ�" & _
        " Where Nvl(����,1)=2 And Nvl(���ÿ���,1)=1 And Nvl(�������,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Union ALL" & _
        " Select Distinct A.����,A.����,A.����" & _
        " From �ѱ� A,�ѱ����ÿ��� B" & _
        " Where A.����=B.�ѱ� And B.����ID=[1]" & _
        " And Nvl(A.����,1)=2 And Nvl(A.���ÿ���,1)=2 And Nvl(A.�������,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(A.��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(A.��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Load��̬�ѱ�", lng����ID)
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    Load��̬�ѱ� = Mid(strTmp, 2)
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
            " From " & IIF(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
            " And B.������� IN([2],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And (A.������Դ is NULL Or A.������Դ=[2])" & _
            " And (A.��������ID is NULL Or A.��������ID=[3])" & _
            IIF(lngҩƷID <> 0, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]")
    Else
        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIF(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
            " And B.������� IN([2],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And D.����ID=C.ID And D.����=[6]" & _
            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
            " And (A.������Դ is NULL Or A.������Դ=[2])" & _
            " And (A.��������ID is NULL Or A.��������ID=[3])" & _
            IIF(lngҩƷID <> 0, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]")
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strҩ��, int��Χ, lng����ID, lngҩƷID, lng��ĿID, bytDay)
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

Public Function Get����ִ�п���(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    objCbo As Object, ByVal str��� As String, ByVal lng��ĿID As Long, ByVal lngҩƷID As Long, _
    ByVal intִ�п��� As Integer, ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, _
    ByVal lng��ǰִ��ID As Long, ByVal int��Ч As Integer, Optional ByVal int��Χ As Integer = 2) As Boolean
'���ܣ�����������Ŀִ�п�����Ϣ���ؿ��õ�ִ�п�����ָ����������
'������intִ�п���=��Ŀִ�п��ұ�־
'      lng���˿���ID=���˿���ID
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
                " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                IIF(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct C.ID,C.����,C.����,C.����,B.�������" & _
                " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[8]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                IIF(bln���, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]") & _
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
                        " From ���ű� A,������ҳ B" & _
                        " Where A.ID=B.��ǰ����ID And B.����ID=[9] And B.��ҳID=[10]" & _
                        " Union " & _
                        " Select ID,����,����,���� From ���ű� Where ID=[6]" & _
                        " Order by ����"
                End If
            Case 3 '3-����Ա���ڿ���
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
            Case 6 '6-���������ڿ���
                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([11],[6]) Order by ����"
        End Select
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strҩ��, int��Χ, lng���˿���ID, lngҩƷID, lng��ĿID, lng��ǰִ��ID, UserInfo.����ID, bytDay, lng����ID, lng��ҳID, lng��������ID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str����)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strƵ��, "," & str��Χ & ",")
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int��Χ)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int��Χ)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int��Χ, lng��ҩ;��ID, strƵ��)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strƵ��, lng��ҩ;��ID, int��Χ)
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
    Optional objCbo As Object, Optional ByVal int��Χ As Integer = 2, Optional blnAppend As Boolean) As Boolean
'���ܣ���ȡ���õĿ���ҽ����ָ������������
'������lng���˿���ID=�������ڿ���ID
'      bln��ʿվ=�Ƿ��ɻ�ʿ��ҽ����ҽ��
'      objCbo=Ҫ����ҽ���嵥��������
'      strȱʡҽ��=ȱʡ��λ��ҽ��,�������objCbo,�������ȶ�λ,�ٷ���ȱʡҽ����ҽ��ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
'      blnAppend=�Ƿ񽫵�ǰȱʡҽ�����ӵ��б��е���ʽ,��ʱ"bln��ʿվ,strȱʡҽ��,objCbo"����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    If bln��ʿվ Then
        If blnAppend And strȱʡҽ�� <> "" Then
            strSQL = "Select ID,���,����,���� From ��Ա�� Where ����=[4]"
        Else
            '�������ڿ��ҵ�ҽ��
            strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIF(objCbo Is Nothing, ",B.����ID", "") & _
                " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
                " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
                " And B.����ID=[1]" & _
                " Order by A.����"
            'ȫԺסԺ���ҵ�ҽ��
            strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN([2],3)"
            strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIF(objCbo Is Nothing, ",B.����ID", "") & _
                " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
                " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
                " And B.����ID IN(" & strSQL & ")" & _
                " Order by A.����"
        End If
    Else 'ҽ����ҽ��ʱ,����Ϊֻ��Ϊҽ������
        strSQL = "Select ID,���,����,���� From ��Ա�� Where ID=[3]"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng���˿���ID, int��Χ, UserInfo.ID, strȱʡҽ��)
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
        If blnAppend Then
            '��ɾ��"����"
            i = SeekCboIndex(objCbo, -1)
            If i <> -1 Then objCbo.RemoveItem objCbo.ListCount - 1
            
            '��λ���������ѡ��
            If Not rsTmp.EOF Then
                i = SeekCboIndex(objCbo, rsTmp!ID)
                If i = -1 Then
                    objCbo.AddItem Nvl(rsTmp!����) & "-" & rsTmp!����
                    objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
                    Call zlControl.CboSetIndex(objCbo.Hwnd, objCbo.NewIndex)
                Else
                    Call zlControl.CboSetIndex(objCbo.Hwnd, i)
                End If
            End If
            
            '����"����"��ѡ��
            AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, "[����...]"
            SetComboData objCbo.Hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
        Else
            'ȫ���¼���
            objCbo.Clear
            For i = 1 To rsTmp.RecordCount
                AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, Nvl(rsTmp!����) & "-" & rsTmp!����
                SetComboData objCbo.Hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
                If rsTmp!���� = strȱʡҽ�� Then
                    Call zlControl.CboSetIndex(objCbo.Hwnd, i - 1)
                End If
                rsTmp.MoveNext
            Next
        End If
    End If
    Get����ҽ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID, lngִ�п���ID)
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
    
    strTmp = IIF(int��Χ = 1, "����", "סԺ")
    
    '��ȡҩƷ���(�����������ҩƷ),ҩ��������ҩƷ����Ч��
    strSQL = _
        " Select Nvl(Sum(A.��������),0)/Nvl(B." & strTmp & "��װ,1) as ���" & _
        " From ҩƷ��� A,ҩƷ��� B" & _
        " Where A.ҩƷID=B.ҩƷID(+) And A.����=1" & _
        " And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
        " And A.ҩƷID=[1] And A.�ⷿID=[2]" & _
        " Group by Nvl(B." & strTmp & "��װ,1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҩƷID, lng�ⷿID)
    If Not rsTmp.EOF Then
        GetStock = Format(rsTmp!���, "0.00000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemField(ByVal strTable As String, ByVal lngID As Long, Optional ByVal strField As String) As Variant
'���ܣ���ȡָ����ָ���ֶ���Ϣ
'˵����δ����NULLֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If strField = "" Then
        strSQL = "Select * From " & strTable & " Where ID=[1]"
    Else
        strSQL = "Select " & strField & " From " & strTable & " Where ID=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngID)
    If Not rsTmp.EOF Then
        If strField = "" Then
            Set GetItemField = rsTmp
        Else
            GetItemField = rsTmp.Fields(strField).Value
        End If
    End If
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
        IIF(bln��Ч And int��Դ = 1, " And A.��Ч=1", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng���ID, int��Դ)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng�䷽ID, int��Դ)
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1)
    Else
        strSQL = _
            " Select * From �����ļ�Ŀ¼ Where ����=[1] And " & _
            IIF(lngDeptID = -1, "Ӧ��=1", "Ӧ��=2 And ','||����ID||',' Like [2]") & _
            " Order by ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1, "%," & lngDeptID & ",%")
        If rsTmp.EOF Then  'ָ�������޸��ಡ������鹫�ò���
            strSQL = "Select * From �����ļ�Ŀ¼ Where ����=[1] And Ӧ��=1 Order by ���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngPatientID, vPageID)
    
    Do While Not rsTmp.EOF
        zlDatabase.ExecuteProcedure "ZL_���˲���_�鵵(" & rsTmp(0) & ",'" + UserInfo.���� + "')", ""
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҩ��ID, lngҩƷID)
    
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
                    IIF(gbln�Ӱ�Ӽ�, "*Decode(A.�Ӱ�Ӽ�,1,1+Nvl(B.�Ӱ�Ӽ���,0)/100,1)", "") & " as ���" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
                " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҩƷID, dblʱ��)
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
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", Val(Format(dblʱ�� * dbl����, gstrDec)), str�ѱ�, Abs(dblʱ��), lngҩƷID)
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
    Optional ByVal dbl���� As Double, Optional ByVal blnNone�Ӱ�Ӽ� As Boolean, Optional ByVal lngִ�п���ID As Long) As Double
'���ܣ���ȡ�շ�ϸĿ�ĵ�ǰ�ۼۼ۸���,��۷���0
'������str�ѱ�=�Ƿ񰴷ѱ������۵�ʵ�ս��
'      dbl����=���ѱ����ʱ,����Ҫ��������(���ۼ۵�λ),��ʱ�������ʵ�ս��
'      lngִ�п���ID=�������˷ѱ�ʱ��Ҫ,���ܰ��ɱ��Ӵ��ۼ���
'      gbln�Ӱ�Ӽ�=����ʱ���������,�����ط���ΪFalse
'      blnNone�Ӱ�Ӽ�=Ϊ��ʱ,"gbln�Ӱ�Ӽ�"��Ч
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl��� As Double
    
    On Error GoTo errH
    
    If str�ѱ� = "" Then
        strSQL = _
            " Select Sum(Decode(Nvl(A.�Ƿ���,0),1,NULL," & _
                "B.�ּ�" & IIF(gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ�, "*Decode(A.�Ӱ�Ӽ�,1,1+Nvl(B.�Ӱ�Ӽ���,0)/100,1)", "") & ")) as ���" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
            " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
        If Not rsTmp.EOF Then dbl��� = Nvl(rsTmp!���, 0)
    Else
        '�������Խ�ActualMoney������SQLһ��д��������ѱ���ܱ�ɾ�����󲻳�����
        strSQL = _
            " Select A.���ηѱ�,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.������ĿID,Decode(Nvl(A.�Ƿ���,0),1,NULL," & _
                "B.�ּ�" & IIF(gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ�, "*Decode(A.�Ӱ�Ӽ�,1,1+Nvl(B.�Ӱ�Ӽ���,0)/100,1)", "") & ") as ���" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
            " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
        For i = 1 To rsTmp.RecordCount
            If Nvl(rsTmp!���ηѱ�, 0) = 1 Then
                dbl��� = dbl��� + Format(dbl���� * Format(Nvl(rsTmp!���, 0), "0.00000"), gstrDec)
            Else
                dbl��� = dbl��� + ActualMoney(str�ѱ� & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsTmp!������ĿID, Format(dbl���� * Format(Nvl(rsTmp!���, 0), "0.00000"), gstrDec), _
                    lng��ĿID, lngִ�п���ID, dbl����, IIF(gbln�Ӱ�Ӽ� And Not blnNone�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1, Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0))
            End If
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


Public Function ActualMoney(str�ѱ� As String, ByVal lng������ĿID As Long, ByVal curӦ�ս�� As Currency, _
    Optional ByVal lngҩƷID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal dbl���� As Double, Optional ByVal dbl�Ӱ�Ӽ��� As Double) As Currency
'���ܣ����ݷѱ�,������ĿID,Ӧ�ս��,���ֶα������۹������ʵ�ս������ҩƷ�����Ϣ���ɱ����ձ����������ʵ�ս��
'������str�ѱ�=���˷ѱ�����ǰ���̬�ѱ����,�����ʽΪ"���˷ѱ�,��̬�ѱ�1,��̬�ѱ�2,..."
'      str���,lngҩƷID,lng�ⷿID,dbl����,dbl�Ӱ�Ӽ���=ҩƷ����Ŀ��Ҫ����
'      dbl����=�����������ڵ��ۼ�����
'      dbl�Ӱ�Ӽ���=С������,�����Ӧ�ս���Ѱ��Ӱ�Ӽۼ���ʱ��Ҫ�����ڻ�ԭ������
'���أ������۹���ͱ��������ʵ�ս��,����Ƕ�̬�ѱ�,��"str�ѱ�"�������Żݷѱ�(ע�����δ���ۼ���,����ԭ������)
'˵����
'���ɱ��ۼ��ձ������۵����ּ��㷽��(ʵ����һ��)��
'1.���۽�� = �ɱ���� * (1 + ���ձ���)
'2.���۽�� = �ɱ��� * (1 + ���ձ���) * ��������
'��صļ��㹫ʽ��
'      �ɱ��� = ҩƷ�ۼ� * (1 - �����)
'      �ɱ���� = �ۼ۽�� * (1 - �����) = �ɱ��� * ��������
'      �п����ʱ:����� = ����� / �����,����:����� = ָ�������
'      ���ڷ���ҩƷ��Ӧÿ���������ηֱ����ɱ��ۺͳɱ����
'        ����ʱ�۷�����"ҩƷ�ۼ�=ʵ�ʽ��/ʵ������"��������ʱ��ҩƷ��治��ʱ��������ۼ��㡣
    Dim rsTmp As New ADODB.Recordset
    Dim rsBase As New ADODB.Recordset
    Dim rsDrug As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dblCost As Double, dblRate As Double
    Dim dblCurTime As Double, dblRebate As Double
    
    Dim blnDynamic As Boolean, dblCurMoney As Double
    Dim strMin�ѱ� As String, dblMinMoney As Double
    
    On Error GoTo errH
    
    '���ò����ۼ���ʱ��ȱʡֵ
    ActualMoney = curӦ�ս��
    If str�ѱ� = "" Or curӦ�ս�� = 0 Then Exit Function
       
    blnDynamic = InStr(str�ѱ�, ",") > 0
    If Not blnDynamic Then
        strSQL = _
            " Select �ѱ�,Nvl(ʵ�ձ���,0) as ʵ�ձ���,[3]*Nvl(ʵ�ձ���,0)/100 as ʵ�ս��,Nvl(���㷽��,0) as ���㷽��" & _
            " From �ѱ���ϸ Where ������ĿID=[1] and �ѱ�=[2] and Abs([3]) Between Ӧ�ն���ֵ and Ӧ�ն�βֵ"
        Set rsBase = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", lng������ĿID, str�ѱ�, curӦ�ս��)
    Else
        strSQL = _
            " Select A.�ѱ�,Nvl(A.ʵ�ձ���,0) as ʵ�ձ���,[3]*Nvl(A.ʵ�ձ���,0)/100 as ʵ�ս��,Nvl(A.���㷽��,0) as ���㷽��" & _
            " From �ѱ���ϸ A,�ѱ� B" & _
            " Where A.�ѱ�=B.���� And A.������ĿID=[1] And Instr([2],','||B.����||',')>0" & _
            " And Abs([3]) Between A.Ӧ�ն���ֵ and A.Ӧ�ն�βֵ" & _
            " Order by Nvl(A.���㷽��,0),Nvl(A.ʵ�ձ���,0),Nvl(B.����,1),B.����"
        Set rsBase = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", lng������ĿID, "," & str�ѱ� & ",", curӦ�ս��)
    End If
    If rsBase.EOF Then Exit Function
    
    dblMinMoney = 9999999999999#
        
    '���ֶα������۹������:���������۹����д��۱�����С�ķѱ�
    If blnDynamic Then rsBase.Filter = "���㷽��=0"
    If Not rsBase.EOF Then
        If Nvl(rsBase!���㷽��, 0) = 0 Then
            strMin�ѱ� = rsBase!�ѱ�
            dblMinMoney = rsBase!ʵ�ս��
        End If
    End If
    
    '���ɱ����ձ����������:���ɱ����չ����м��ձ�����С�ķѱ�
    If blnDynamic Then rsBase.Filter = "���㷽��=1"
    If Not rsBase.EOF Then
        If Nvl(rsBase!���㷽��, 0) = 1 And lngҩƷID <> 0 And dbl���� <> 0 Then
            dblRate = Nvl(rsBase!ʵ�ձ���, 0) / 100
            curӦ�ս�� = curӦ�ս�� / (1 + dbl�Ӱ�Ӽ���) '�����Ӧ���Ǹ��ݼӰ�Ӽۼ������,������Ҫ��ԭ
            
            'ȡҩƷ��Ϣ
            strSQL = _
                " Select B.���,A.ָ�������,Nvl(C.�ּ�,0) as �ۼ�," & _
                " Nvl(A.ҩ������,0) as ����,Nvl(B.�Ƿ���,0) as ���" & _
                " From ҩƷ��� A,�շ���ĿĿ¼ B,�շѼ�Ŀ C" & _
                " Where A.ҩƷID=B.ID And B.ID=C.�շ�ϸĿID And A.ҩƷID=[1]" & _
                " And Sysdate Between C.ִ������ And Nvl(C.��ֹ����,To_Date('3000-01-01', 'YYYY-MM-DD'))"
            Set rsDrug = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", lngҩƷID)
            If rsDrug.EOF Then GoTo EndCalc
            If InStr(",5,6,7,", rsDrug!���) = 0 Then GoTo EndCalc '��ҩƷ������
            
            If lng�ⷿID = 0 Then
                'û��ȷ��ҩ��ʱ(�����ģʽ),�����򲻷���ҩƷ����ָ���������
                dblCurMoney = curӦ�ս�� * (1 - Nvl(rsDrug!ָ�������, 0) / 100) * (1 + dblRate) * (1 + dbl�Ӱ�Ӽ���)
                If dblCurMoney < dblMinMoney Then strMin�ѱ� = rsBase!�ѱ�: dblMinMoney = dblCurMoney
            ElseIf rsDrug!���� = 0 Then
                '������ҩƷ:
                strSQL = "[1]*(1-Nvl(Decode(Sign(Nvl(A.ʵ�ʽ��,0)),1,A.ʵ�ʲ��/A.ʵ�ʽ��,B.ָ�������/100),0))"
                strSQL = "Select " & strSQL & " as �ɱ���� From ҩƷ��� A,ҩƷ��� B" & _
                         " Where A.ҩƷID(+)=B.ҩƷID And A.�ⷿID(+)=[2] And A.ҩƷID(+)=[3] And A.����(+)=1 And B.ҩƷID=[3]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", curӦ�ս��, lng�ⷿID, lngҩƷID)
                dblCost = Nvl(rsTmp!�ɱ����, 0)
                If dblCost <> 0 Then
                    dblCurMoney = dblCost * (1 + dblRate) * (1 + dbl�Ӱ�Ӽ���)
                    If dblCurMoney < dblMinMoney Then strMin�ѱ� = rsBase!�ѱ�: dblMinMoney = dblCurMoney
                End If
            ElseIf rsDrug!���� = 1 Then
                '����ҩƷ:ÿ��������ɱ�����
                strSQL = _
                    " Select Nvl(����,0) as ����,Nvl(��������,0) as ���," & _
                    " Nvl(Decode(Nvl(ʵ������,0),0,0,ʵ�ʽ��/ʵ������),0) as ʱ��," & _
                    " Nvl(ʵ�ʲ��,0) as ʵ�ʲ��,Nvl(ʵ�ʽ��,0) as ʵ�ʽ��" & _
                    " From ҩƷ���" & _
                    " Where (Nvl(����,0)=0 Or Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
                    " And Nvl(��������,0)<>0 And ����=1 And �ⷿID=[1] And ҩƷID=[2]" & _
                    " Order by Nvl(����,0)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", lng�ⷿID, lngҩƷID)
                
                dblRebate = 0: dblCurTime = 0
                For i = 1 To rsTmp.RecordCount
                    If dbl���� = 0 Then Exit For
                    
                    'ȡС��
                    If dbl���� <= rsTmp!��� Then
                        dblCurTime = dbl����
                    Else
                        dblCurTime = rsTmp!���
                    End If
                    
                    If rsTmp!ʵ�ʽ�� <> 0 Then
                        '�����γɱ���:�����������
                        dblCost = IIF(rsDrug!��� = 1, rsTmp!ʱ��, rsDrug!�ۼ�) * (1 - rsTmp!ʵ�ʲ�� / rsTmp!ʵ�ʽ��)
                    Else
                        '�޿���ָ���������:�޿���������SQL�����ſ�
                        dblCost = IIF(rsDrug!��� = 1, rsTmp!ʱ��, rsDrug!�ۼ�) * (1 - Nvl(rsDrug!ָ�������, 0) / 100)
                    End If
                    If dblCost <> 0 Then
                        dblRebate = dblRebate + dblCost * (1 + dblRate) * dblCurTime
                        dbl���� = dbl���� - dblCurTime
                    End If
                    rsTmp.MoveNext
                Next
                If dbl���� <> 0 Then GoTo EndCalc '����δ�ֽ����,��治��,������
                dblCurMoney = dblRebate * (1 + dbl�Ӱ�Ӽ���)
                If dblCurMoney < dblMinMoney Then strMin�ѱ� = rsBase!�ѱ�: dblMinMoney = dblCurMoney
            End If
        End If
    End If
    
EndCalc:
    If dblMinMoney <> 9999999999999# Then
        str�ѱ� = strMin�ѱ�
        ActualMoney = Format(dblMinMoney, gstrDec)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
        IIF(lng���˿���ID <> 0, " And A.ID<>[2]", "") & _
        IIF(blnBed, " And Exists(Select ����ID From ��λ״����¼ Where ����ID=A.ID)", "") & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", "," & strTmp & ",", lng���˿���ID)
    
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

Public Function GetAuditName(ByVal strName As String) As String
'���ܣ���"ʵϰҽ��/���ҽ��"��ȡ���ҽ����
    GetAuditName = Mid(strName, InStr(strName, "/") + 1)
End Function

Public Function HaveAuditPriv(Optional ByVal str���� As String) As Boolean
'���ܣ��жϵ�ǰ��Ա�Ƿ����"ִҵҽʦ"���ʸ�
'������str����=�ж�ָ����Ա��������ʱ�жϵ�ǰ��Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strHave As String, strNone As String
    Dim strSQL As String
    
    If str���� = "" Then str���� = UserInfo.����
    
    If InStr(strHave & "|", "|" & str���� & "|") > 0 Then
        HaveAuditPriv = True: Exit Function
    ElseIf InStr(strNone & "|", "|" & str���� & "|") > 0 Then
        HaveAuditPriv = False: Exit Function
    End If
    
    On Error GoTo errH
    strSQL = "Select B.���� From ��Ա�� A,ִҵ��� B Where A.����=[1] And A.ִҵ���=B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "HaveAuditPriv", str����)
    If Not rsTmp.EOF Then
        If rsTmp!���� = "ִҵҽʦ" Or rsTmp!���� = "ִҵ����ҽʦ" Then HaveAuditPriv = True
    End If
    If HaveAuditPriv Then
        strHave = strHave & "|" & str����
    Else
        strNone = strNone & "|" & str����
    End If
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str����)
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
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
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
        " Select M.����ID,M.��ҳID,M.���˿���ID,M.ִ�п���ID,C.ID,C.���,C.�Ƿ���,D.��������,B.������ĿID,A.����," & _
        " Nvl(A.����,0) as ����,C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.�����շ���,Decode(Nvl(C.�Ƿ���,0),1,A.����,B.�ּ�) as ����" & _
        " From ����ҽ����¼ M,����ҽ���Ƽ� A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D" & _
        " Where A.�շ�ϸĿID=B.�շ�ϸĿID And B.�շ�ϸĿID=C.ID" & _
        " And C.ID=D.����ID(+) And M.ID=A.ҽ��ID And M.ID=[1]" & _
        " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
        " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    If gbln��������ۿ� And Not rsTmp.EOF And str�ѱ� <> "" Then
        rsTmp.Filter = "����=1"
        If Not rsTmp.EOF Then blnHaveSub = True
        rsTmp.Filter = 0
    End If
    For i = 1 To rsTmp.RecordCount
        If InStr(",5,6,7,", rsTmp!���) > 0 And Nvl(rsTmp!�Ƿ���, 0) = 1 Then
            '�趨�ļƼ���ʱ��ҩƷ���ۼ���
            lngִ�п���ID = Get�շ�ִ�п���ID(rsTmp!����ID, Nvl(rsTmp!��ҳID, 0), rsTmp!���, rsTmp!ID, 4, Nvl(rsTmp!���˿���ID, 0), 0, IIF(Not IsNull(rsTmp!��ҳID), 2, 1))
            dbl���� = Format(CalcDrugPrice(rsTmp!ID, lngִ�п���ID, dbl���� * Nvl(rsTmp!����, 0), , True), "0.00000")
        ElseIf rsTmp!��� = "4" And Nvl(rsTmp!�Ƿ���, 0) = 1 And Nvl(rsTmp!��������, 0) = 1 Then
            '�趨�ļƼ���ʱ�����ĵ��ۼ���
            lngִ�п���ID = Get�շ�ִ�п���ID(rsTmp!����ID, Nvl(rsTmp!��ҳID, 0), rsTmp!���, rsTmp!ID, 4, Nvl(rsTmp!���˿���ID, 0), 0, IIF(Not IsNull(rsTmp!��ҳID), 2, 1), Nvl(rsTmp!ִ�п���ID, 0))
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
    
    strTmp = IIF(int��Χ = 1, "����", "סԺ")
    
    strSQL = _
        " Select ҩƷID,Sum(Nvl(��������,0)) as ��� From ҩƷ���" & _
        " Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))" & _
        " And ���� = 1 And �ⷿID=[1]" & IIF(lngҩƷID <> 0, " And ҩƷID=[2]", "") & _
        " Group by ҩƷID Having Sum(Nvl(��������,0))<>0"
    strSQL = "Select A.ҩƷID,A.����ϵ��,A." & strTmp & "��װ,A." & strTmp & "��λ,A.�ɷ����," & _
        " A.ҩ������,B.�Ƿ���,C.���/A." & strTmp & "��װ as ���,B.����,Nvl(D.����,B.����) as ����,B.���,B.����,B.����ʱ��,B.�������" & _
        " From ҩƷ��� A,�շ���ĿĿ¼ B,(" & strSQL & ") C,�շ���Ŀ���� D" & _
        " Where A.ҩƷID=B.ID And A.ҩƷID=C.ҩƷID(+)" & _
        " And B.ID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=[5]" & _
        IIF(blnͣ��, " And B.������� IN([3],3) And (B.����ʱ�� is NULL Or B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))", "") & _
        " And A.ҩ��ID=[4]" & IIF(lngҩƷID <> 0, " And A.ҩƷID=[2]", "") & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҩ��ID, lngҩƷID, int��Χ, lngҩ��ID, IIF(gbln��Ʒ��, 3, 1))
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
    On Error GoTo ErrHand

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
            vntNo = IIF(IsNull(!������), 0, !������)
            
            strSQL = "Select Nvl(Max(����ID),0)+1 as ����ID From ������Ϣ Where ����ID>=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
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
            .Update "������", IIF(vntNo - 10 > 0, vntNo - 10, 1)
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
                blnByDate = (IIF(IsNull(!����ֵ), 1, !����ֵ) = 2)
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
            vntNo = IIF(IsNull(!������), 0, !������)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(סԺ��),0)+1 as סԺ�� From ������Ϣ Where סԺ��>=[1]"
            Else
                strSQL = "Select Nvl(Max(סԺ��),To_Number(To_Char(Sysdate,'YYMM')||'0000'))+1 as סԺ��" & _
                    " From ������Ϣ Where סԺ�� Like To_Number(To_Char(Sysdate,'YYMM'))||'%' And סԺ��>=[1]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
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
                .Update "������", IIF(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "������", IIF(vntNo - 10 > Val(Format(curDate, "YYMM0000")), vntNo - 10, Val(Format(curDate, "YYMM0001")))
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
                blnByDate = (IIF(IsNull(!����ֵ), 1, !����ֵ) = 2)
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
            vntNo = IIF(IsNull(!������), 0, !������)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(�����),0)+1 as ����� From ������Ϣ Where �����>=[1]"
            Else
                strSQL = "Select Nvl(Max(�����),To_Number(To_Char(Sysdate,'YYMMDD')||'0000'))+1 as �����" & _
                    " From ������Ϣ Where ����� Like To_Number(To_Char(Sysdate,'YYMMDD'))||'%' And �����>=[1]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
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
                .Update "������", IIF(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "������", IIF(vntNo - 10 > Val(Format(curDate, "YYMMDD0000")), vntNo - 10, Val(Format(curDate, "YYMMDD0001")))
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
            
            vntNo = Val(IIF(IsNull(!������), 0, !������)) + 1
            
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
            strYear = IIF(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIF(IsNull(!������), "", !������)
            
            If IIF(IsNull(!��Ź���), 0, !��Ź���) = 1 Then
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
ErrHand:
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!���ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID, int����)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID)
    DeptIsWoman = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer) As String
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
        " Where A.����ID=[1] And Nvl(A.��ҳID,0)=[2] And Nvl(A.Ӥ��,0)=[3]" & _
        " And B.ҽ��ID=A.ID And B.ִ�в���ID+0 IN(" & strSQL & ")" & _
        " And B.ִ��״̬ IN(0,3) And A.������ĿID=C.ID And B.ִ�в���ID=D.ID" & _
        " And Not (A.������� IN('F','G','D') And A.���ID is Not NULL)" & _
        " And Not (A.�������='Z' And Nvl(C.��������,'0')<>'0')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID, lng��ҳID, intӤ��)
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

Public Function ExistWaitDrug(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer) As String
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
        " And A.����ID=[1] And Nvl(A.��ҳID,0)=[2] And Nvl(A.Ӥ����,0)=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID, lng��ҳID, intӤ��)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
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

Public Function ItemIsVarPrice(ByVal lng��ĿID As Long) As Boolean
'���ܣ��ж�ָ����Ŀ�Ƿ���(��ҩƷ�͸������õ�����)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select A.���,A.�Ƿ���,B.�������� From �շ���ĿĿ¼ A,�������� B Where A.ID=B.����ID(+) And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str����, lng��ĿID)
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
    On Error GoTo ErrHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIF(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

ErrHand:
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, lng���ͺ�)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO, int��¼����)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)

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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng�ű�ID)
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng�ű�ID, str�ű�)
        If Not rsTmp.EOF Then GetRegistRoom = rsTmp!��������
    ElseIf int���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        strSQL = "Select * From �ҺŰ������� Where �ű�ID=" & lng�ű�ID
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlPublic", adOpenStatic, adLockOptimistic) '��д��¼��
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngID)
        
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptExist(ByVal str�������� As String, ByVal int������� As Integer) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ID From ��������˵�� Where ��������=[1] And ������� IN([2],3) And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str��������, int�������)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
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
    
    strSQL = "Select NO From H" & strTable & " Where NO=[1] And Rownum<2" & IIF(strWhere <> "", " And " & strWhere, "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)
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
        IIF(lng���ͺ� <> 0, " And A.���ͺ�+0=[2]", "") & _
        " And A.ҽ��ID IN(" & strSQL & ")" & _
        " Group by B.NO"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, lng���ͺ�)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str����)
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strҽ��)
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
        strMsg = """" & strҽ�� & """Ҫ��Ĵ���ְ�����㣺" & vbCrLf & vbCrLf & IIF(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """δ����ְ��"
    ElseIf intְ��B < intְ��A Then
        '��ֵԽСְ��Խ��
        strMsg = """" & strҽ�� & """Ҫ��Ĵ���ְ�����㣺" & vbCrLf & vbCrLf & IIF(blnҽ��, "��ҽ���򹫷Ѳ���,", "") & _
            "��ҩƷҪ��ְ������Ϊ""" & Split(STR_ְ��, ",")(intְ��B - 1) & """�����´�,��ҽ��""" & strҽ�� & """��ְ��Ϊ""" & Split(STR_ְ��, ",")(intְ��A - 1) & """��"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiHaveOut(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
'���ܣ��жϲ����Ƿ��ѳ�Ժ(����Ԥ��Ժ),���ڲ��������ж�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ����ID From ������ҳ Where (��Ժ���� is Not Null Or Nvl(״̬,0)=3) And ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then PatiHaveOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadAdviceSignSource(ByVal int�������� As Integer, _
    ByVal lng����ID As Long, ByVal varTime As Variant, strIDs As String, _
    ByVal lngǩ��ID As Long, ByVal blnMoved As Boolean, strSource As String, _
    Optional ByVal lngǰ��ID As Long, Optional ByVal colStopTime As Collection) As Integer
'���ܣ���ȡ�������ڵ���ǩ��/��֤��ҽ��Դ������
'������
'  int��������=Ҫǩ��/��֤ǩ����ҽ��״̬
'  ǩ��ʱ���룺
'    lng����ID
'    varTime=���˹Һŵ��Ż���ҳID
'    strIDs=ָ��Ҫǩ����ҽ��ID����(��ID)
'    lngǰ��ID=�¿�ҽ��Ҫǩ����ҽ����Դ(�Ƿ�ҽ��)
'    colStopTime=ֹͣҽ��ǩ��ʱ���������ҽ��ִ����ֹʱ�������
'  ��֤ǩ��ʱ��
'    lngǩ��ID=ǩ����¼��ID
'    blnMoved=�Ƿ�ҽ��������ת��
'���أ�ǩ��/��֤ǩ����Դ�����ɹ���
'      strIDs=ǩ��/��֤ǩ����ҽ��ID����(ÿ����ϸID)
'      strSource=ǩ��/��֤ǩ����ҽ��Դ��
    Dim rsTmp As New ADODB.Recordset
    Dim str��IDs As String, strSQL As String, i As Long
    Dim arrField As Variant, strField As String
    Dim strLine As String, intRule As Integer
    
    On Error GoTo errH
    
    str��IDs = strIDs
    strSource = "": strIDs = ""
    intRule = 1 '�������µ�ҽ��ǩ��Դ�����ɹ�����
    
    If lngǩ��ID = 0 Then
        'ǩ��ʱ
        If int�������� = 1 Then
            '���¿���ҽ������ǩ�������ξ���/סԺ��ǰҽ�����´��δǩ��ҽ��
            strSQL = _
                " Select A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID is Null And B.��������=1" & _
                " And A.ҽ��״̬=1 And A.����ҽ��=[3] And Nvl(A.ǰ��ID,0)=[5] And A.����ID=[1]" & _
                IIF(TypeName(varTime) = "String", " And A.�Һŵ�=[2]", " And A.��ҳID=[2]") & _
                IIF(str��IDs <> "", " And Instr([4],','||Nvl(A.���ID,A.ID)||',')>0", "") & _
                " Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID, varTime, UserInfo.����, "," & str��IDs & ",", lngǰ��ID)
        Else
            '��Ҫ���ϻ�ֹͣ��ҽ������ǩ�����¿�ʱǩ������ָ��ҽ������һ���ǵ�ǰҽ���´�
            strSQL = _
                " Select A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID is Not Null And B.��������=1 And A.����ID=[1]" & _
                IIF(TypeName(varTime) = "String", " And A.�Һŵ�=[2]", " And A.��ҳID=[2]") & _
                IIF(str��IDs <> "", " And Instr([3],','||Nvl(A.���ID,A.ID)||',')>0", "") & _
                " Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID, varTime, "," & str��IDs & ",")
        End If
    Else
        '��֤ǩ��ʱ:�ȶ�ȡǩ��ʱ��Դ�����ɹ���
        strSQL = "Select ǩ������ From ҽ��ǩ����¼ Where ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "ҽ��ǩ����¼", "Hҽ��ǩ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngǩ��ID)
        If Not rsTmp.EOF Then intRule = Nvl(rsTmp!ǩ������, 1)
        '--
        strSQL = _
            " Select A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID=[1] Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
        If blnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngǩ��ID)
    End If
    
    'ҽ��Դ�ĵĲ�ͬ���ɹ���
    If intRule = 1 Then
        If int�������� = 8 Then
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,ִ����ֹʱ��,ִ������,������־,����ҽ��,����ʱ��"
        Else
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,ִ������,������־,����ҽ��,����ʱ��"
        End If
    End If
    arrField = Split(strField, ",")
        
    '����ҽ��ǩ��Դ��
    Do While Not rsTmp.EOF
        strLine = ""
        For i = 0 To UBound(arrField)
            If lngǩ��ID = 0 And int�������� = 8 And arrField(i) = "ִ����ֹʱ��" Then
                'ֹͣҽ��ǩ��ʱ,����ֹʱ�����⴦����������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                strLine = strLine & vbTab & colStopTime("_" & Nvl(rsTmp!���ID, rsTmp!ID))
            Else
                If IsNull(rsTmp.Fields(arrField(i)).Value) Then
                    strLine = strLine & vbTab & ""
                Else
                    If rsTmp.Fields(arrField(i)).Type = adDBTimeStamp Then
                        strLine = strLine & vbTab & Format(rsTmp.Fields(arrField(i)).Value, "yyyy-MM-dd HH:mm:ss")
                    Else
                        strLine = strLine & vbTab & rsTmp.Fields(arrField(i)).Value
                    End If
                End If
            End If
        Next
        strSource = strSource & vbCrLf & Mid(strLine, 2)
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    strSource = Mid(strSource, 3)
    strIDs = Mid(strIDs, 2)
    
    ReadAdviceSignSource = intRule
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistsDiagNoses(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal int���� As Integer) As Boolean
'���ܣ���鲡��ָ��������Ƿ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��¼��Դ,����ID,���ID,�������,�Ƿ����� From ������ϼ�¼" & _
        " Where ����ID=[1] And Nvl(��ҳID,0)=[2] And �������=[3] And ȡ��ʱ�� Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID, int����)
    ExistsDiagNoses = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistsSpecAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ��жϲ��˱���סԺ�Ƿ��Ѿ��´���ȷ�ϵ�����ҽ��(��Ժ��תԺ������)
'���أ�������ڣ�����ҽ����ʾ��Ϣ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    strSQL = "Select A.����,A.Ӥ��,A.ҽ������ From ����ҽ����¼ A,������ĿĿ¼ B" & _
        " Where A.������ĿID=B.ID And A.�������='Z' And B.�������� IN('5','6','11')" & _
        " And A.ҽ��״̬ Not IN(1,2,4) And A.����ID=[1] And A.��ҳID=[2]" & _
        " Order by A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng����ID, lng��ҳID)
    Do While Not rsTmp.EOF
        strMsg = strMsg & vbCrLf & "������" & rsTmp!ҽ������ & IIF(Nvl(rsTmp!Ӥ��, 0) <> 0, "(Ӥ��" & Nvl(rsTmp!Ӥ��, 0) & ")", "")
        rsTmp.MoveNext
    Loop
    If strMsg <> "" Then
        rsTmp.MoveFirst
        strMsg = "������������""" & rsTmp!���� & """�Ѿ�ȷ���´���������ҽ����" & vbCrLf & strMsg & vbCrLf & ""
    End If
    ExistsSpecAdvice = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check�����÷�(ByVal lng�÷�ID As Long, ByVal lng��ĿID As Long, ByVal int��Դ As Integer) As Boolean
'���ܣ����ָ�����÷��Ƿ�������ָ������Ŀ
'������int��Դ=1-����,2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(A.�÷�ID) as ����,Max(Decode(A.�÷�ID,[2],1,0)) as ָ��" & _
        " From �����÷����� A,������ĿĿ¼ B" & _
        " Where A.�÷�ID=B.ID And B.������� IN([3],3) And A.��ĿID=[1] And A.����>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID, lng�÷�ID, int��Դ)
    If Nvl(rsTmp!����, 0) <= 1 Then
        Check�����÷� = True
    ElseIf Nvl(rsTmp!ָ��, 0) = 1 Then
        Check�����÷� = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceTime(ByVal lngҽ��ID As Long, ByVal int���� As Integer) As Date
'���ܣ���ȡҽ��ָ��������ʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select Max(����ʱ��) as ʱ�� From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, int����)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!ʱ��) Then
            GetAdviceTime = rsTmp!ʱ��
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetAdviceSign(ByVal lngҽ��ID As Long, ByVal int���� As Integer, ByVal str��Ա As String, ByVal datʱ�� As Date) As Long
'���ܣ���ȡָ��ҽ��������ǩ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ǩ��ID From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=[2] And ������Ա=[3] And ����ʱ��=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID, int����, str��Ա, datʱ��)
    If Not rsTmp.EOF Then
        GetAdviceSign = Nvl(rsTmp!ǩ��ID, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
