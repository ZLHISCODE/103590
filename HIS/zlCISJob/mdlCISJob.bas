Attribute VB_Name = "mdlCISJob"
Option Explicit
Public gblnShowInTaskBar As Boolean         '�Ƿ���ʾ��������������
Public gobjRichEPR As New cRichEPR          '�������Ĳ���
Public gobjKernel As New zlPublicAdvice.clsPublicAdvice       '�ٴ����Ĳ���
Public gobjPath As New zlPublicPath.clsPublicPath             '�ٴ�·������
Public gobjRegist As Object
Public gobjCommunity As Object              '���������ӿڶ���
Public gclsInsure As New clsInsure          'ҽ������
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gstrSysName As String                'ϵͳ����
Public glngSys As Long
Public glngModul As Long
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String            'OEM��Ʒ����
Public gfrmMain As Object                   '����̨����
Public gblnOK As Boolean
Public gobjCISBase As Object                '��ʿվ��ҽ��վ���������շѶ���
Public gobjPlugIn As Object                 '��ҹ��ܶ���
Public gobjRis As Object                    'RIS�ӿڲ���
Public gblnKSSStrict As Boolean             '����ҩ���ϸ����
Public gblnKSSAuditType As Boolean             '����ҩ����˷�ʽ����������ҽ��С����п���ҩ����� 0��Ĭ�ϣ�1����ҽ��С��
Public gbln�����ּ����� As Boolean  '�Ƿ����������ּ�����
Public gbln��Ѫ�ּ����� As Boolean  '�Ƿ�������Ѫ�ּ�����
Public gblnѪ��ϵͳ As Boolean  '�Ƿ�װѪ��ϵͳ
Public gobjEmr  As Object                   '�°没������
Public gbln�������Һ���Ч���� As Boolean   '���������Һ���Ч�����Ĳ���
Public gobjLIS As Object                    'LIS���벿��
Public gobjPublicPacs As Object                  'PACS��������
Public gobjExchange As Object               'HL7���ݽ�������
Public gobjPublicExpense As Object           '���ù�������
Public gobjNurseIntegrate As Object        '���廤��ӿڲ���
Public gobjPublicBlood As Object             'Ѫ�⹫������
Public gblnGetPath As Boolean

'����ǩ��
Public gintCA As Integer '����ǩ����֤����
Public gstrESign As String '����ǩ�����Ƴ���
Public grsSign As Recordset  '����ǩ�����ò���

Public gbln��Ѫ����������� As Boolean  '��Ѫ�����������
'������ҩ�ӿ�����,0-δʹ��,1-����,2-��ͨ,3-̫Ԫͨ
Public gbytPass As Byte
'0-ҽ��ѡ��1-��ҩƷĿ¼���룬2-������Դ����
Public gint����������Դ As Integer
'̫Ԫͨ�ӿڶ���
Public gobjPass As Object

Public glngPreHWnd As Long '����֧�������ֹ���

Public Enum ���ó���
    E������� = 1
    EסԺ���� = 2
End Enum

'ϵͳ����
Public gstrLike As String   '�����˫��ƥ�䣬��Ϊ%
Public gint���� As Integer  '����ƥ�䷽ʽ��0-ƴ��,1-���
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"

Public gbytCardLen As Byte '���￨�ų���
Public gblnCardHide As Boolean '���￨��������ʾ

Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gint��ͨ�Һ����� As Integer '��ͨ�Һŵ���Ч����
Public gint����Һ����� As Integer '����Һŵ���Ч����

Public gdblԤ��������鿨 As Double 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
                                                      'Ϊ����(-N)ʱ��ʾ,NԪ������֧��,��ʾ����������NԪ�ڱ���ˢ��,�����������뼴��֧��;���������������
Public gblnִ��ǰ�Ƚ��� As Boolean '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������

Public gstrҽ���˶� As String    '��ѪƤ��ҽ����Ҫ�˶� ��λ��ȡ11����һλΪ ��Ѫҽ�����ڶ�λΪ Ƥ��ҽ��
Public gstr��Һ�������� As String          '��-�����ã���������
Public gblnDo As Boolean  '�Ƿ�ʹ�ø��Ի�����
Public gintҽ��ִ����Ч���� As Integer '�����޸�n���ڵǼǵ�ҽ��ִ�м�¼

Public gint�����Դ As Integer '1-��ҽ��ѡ��������Դ,2-������ϱ�׼����,3-���ռ�����������
Public gstr������� As String '1����/2סԺ��1-������������,2-�����ݿ���ȡ����,3-��ҽ�����˴����ݿ�����
Public gbln����Ӱ����Ϣϵͳ�ӿ� As Boolean
Public gbln����Ӱ����ϢϵͳԤԼ As Boolean
Public gbln�������廤��ӿ� As Boolean
Public gbln�ҺŰ��� As Boolean '����ϵͳ�������Һ��Ű�ģʽ   true�°棬false�ϰ�
Public gblnPatiByID As Boolean 'ϵͳ������ͬһ���ֻ֤�ܶ�Ӧһ����������


'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    p���Ӳ������� = 2250
    p�°�סԺ���� = 2252
    p�°����ﲡ�� = 2251
    p����������д = 1249
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    p�ٴ�·��Ӧ�� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p���Ӳ������� = 1259
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    P�°滤ʿվ = 1265
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    p��Ƭ���߹��� = 1289
    p������� = 1132
    pסԺ���� = 1133
    p���ò�ѯ = 1139
    p���������� = 1113
    p�Ŷӽк�����ģ�� = 1160
    p������ҩ��� = 1266
    p������˹��� = 1267
    p���Ӳ������ = 1560
    p��Ѫ��˹��� = 1268
    p����ӿ� = 2425
    p������Ȩ���� = 1080
    p��Һ�������� = 1345
    P����·��Ӧ�� = 1248
    P�������Ĵ�ӡ = 1566
End Enum

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ������ As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.��� = rsTmp!���
            UserInfo.����ID = Nvl(rsTmp!����ID, 0)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.�û��� = rsTmp!User & ""
            GetUserInfo = True
        End If
    End If
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean, Optional ByVal lngSys As Long) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
'      lngSys=ָ��ϵͳ���ڲ�ģ��Ȩ�ޣ���0�򲻴���Ĭ���ǵ�ǰϵͳ
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    If lngSys = 0 Then lngSys = glngSys
    On Error Resume Next
    strPrivs = gcolPrivs(lngSys & "_" & lngProg)
    If err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove lngSys & "_" & lngProg
        End If
    Else
        err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(lngSys, lngProg)
        gcolPrivs.Add strPrivs, lngSys & "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    '55928:������,2012-11-20
    gblnDo = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0
    
    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    gint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
    
    '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '���￨��������ʾ
    strSQL = "Select ���ų���, Nvl(��������, 0) �������� From ҽ�ƿ���� Where �ض���Ŀ = '���￨'"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���￨")
    If rsTmp.RecordCount > 0 Then
        gblnCardHide = rsTmp!�������� <> "0"
        gbytCardLen = Val("" & rsTmp!���ų���)
    Else
        gblnCardHide = False
        gbytCardLen = 8
    End If
    
    
    '�Һ���Ч����
    strTmp = zlDatabase.GetPara(21, glngSys)
    If Len(strTmp) = 1 Then strTmp = strTmp & strTmp
    gint��ͨ�Һ����� = Val(Mid(strTmp, 1, 1))
    gint����Һ����� = Val(Mid(strTmp, 2, 1))
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    'һ��ͨ������֤
    strTmp = zlDatabase.GetPara(28, glngSys) & "|"
    gdblԤ��������鿨 = Val(Split(strTmp, "|")(0))
    
    
    '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
    gblnִ��ǰ�Ƚ��� = Val(zlDatabase.GetPara(163, glngSys)) <> 0
    
    '����ǩ����֤����
    gintCA = Val(zlDatabase.GetPara(25, glngSys))
    
    '����ǩ�����Ƴ���
    gstrESign = zlDatabase.GetPara(26, glngSys)
    
    '��ȡ������������
    If glngModul = p����ҽ��վ Or glngModul = pסԺҽ��վ Or glngModul = pסԺ��ʿվ Or glngModul = pҽ������վ Or _
        glngModul = P�°滤ʿվ Or glngModul = p������ҩ��� Then
        '��ȡ������������
        Set grsSign = New ADODB.Recordset
        grsSign.Fields.Append "����ID", adBigInt
        grsSign.Fields.Append "����", adBigInt
        grsSign.Fields.Append "�Ƿ�����", adBigInt
        grsSign.CursorLocation = adUseClient
        grsSign.LockType = adLockOptimistic
        grsSign.CursorType = adOpenStatic
        grsSign.Open
    End If
    
    '��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�
    gstrҽ���˶� = zlDatabase.GetPara(186, glngSys)
    
    '����ҩ��ּ�����
    gblnKSSStrict = Val(zlDatabase.GetPara(187, glngSys)) <> 0
    
    '��ҽ��С����п���ҩ�����
    gblnKSSAuditType = Val(zlDatabase.GetPara(248, glngSys)) <> 0
    
    '�Ƿ����������ּ�����
    gbln�����ּ����� = Val(zlDatabase.GetPara(209, glngSys)) <> 0
    
    '�Ƿ�������Ѫ�ּ�����
    gbln��Ѫ�ּ����� = Val(zlDatabase.GetPara(216, glngSys)) <> 0
    
    '�Ƿ�װѪ��ϵͳ
    gblnѪ��ϵͳ = (IsSysSetUp(2200) And Val(zlDatabase.GetPara(236, glngSys)) <> 0)
    
    '���������Һ���Ч�����Ĳ���
    gbln�������Һ���Ч���� = Val(zlDatabase.GetPara(210, glngSys)) <> 0
    
    '61762:������,2012-05-20
    gstr��Һ�������� = Get��Һ��������

    '��Ѫ�����������
    gbln��Ѫ����������� = Val(zlDatabase.GetPara(218, glngSys)) <> 0
    
    '�����޸�n���ڵǼǵ�ҽ��ִ�м�¼
    gintҽ��ִ����Ч���� = Val(zlDatabase.GetPara(220, glngSys))
    '������ҩ�ӿ����ͣ�0-δ���ã�1-�Ĵ���ͨ��2-��ͨ��3-̫Ԫͨ
    gbytPass = Val(zlDatabase.GetPara(30, glngSys))
    
    '����������Դ����
    gint����������Դ = Val(zlDatabase.GetPara(224, glngSys))
    
    '���������Դ
    gint�����Դ = Val(zlDatabase.GetPara(55, glngSys, , 1))
    
    '������뷽ʽ
    gstr������� = zlDatabase.GetPara(65, glngSys, , "11")
    
    gbln����Ӱ����Ϣϵͳ�ӿ� = Val(zlDatabase.GetPara(255, glngSys)) = 1
    
    gblnGetPath = Val(zlDatabase.GetPara(54, glngSys, glngModul)) = 1
    
    strTmp = ""
    gbln�ҺŰ��� = False
    strTmp = zlDatabase.GetPara(256, glngSys) & "|" 'ϵͳ�������Һ��Ű�ģʽ
    If 0 <> Val(Split(strTmp, "|")(0)) Then
        If Split(strTmp, "|")(1) <> "" Then
            strTmp = Format(Split(strTmp, "|")(1), "YYYY-MM-DD")
            If Format(zlDatabase.Currentdate, "YYYY-MM-DD") >= strTmp Then
                gbln�ҺŰ��� = True
            End If
        End If
    End If
    
    'ͬһ���ֻ֤�ܶ�Ӧһ����������
    gblnPatiByID = Val(zlDatabase.GetPara(279, glngSys)) = 1

    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�Һ�ID(ByVal strNO As String) As Long
'���ܣ����ݹҺŵ���ȡ�Һ�ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From ���˹Һż�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get�Һ�ID", strNO)
    If Not rsTmp.EOF Then Get�Һ�ID = rsTmp!ID
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

Public Function GetPatiDiagnose(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int��Դ As Integer) As String
'���ܣ���ȡ����ָ���ξ�����������
'������lng����ID=�Һ�ID����ҳID
'      int��Դ=1-����,2-סԺ
'���أ���"��"�ŷָ��Ķ����ϴ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��¼��Դ,�������,��ϴ���,�������,�Ƿ�����,Mod(�������,10) as ���� From ������ϼ�¼" & _
        " Where ����ID=[1] And ��ҳID=[2] And NVL(�������,1) = 1 And ������� IN(" & IIf(int��Դ = 1, "1,11", "1,2,3,11,12,13") & ")" & _
        " Order by ��¼��Դ,�������,��ϴ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiDiagnose", lng����ID, lng����ID)
    
    '�Ȱ���Դ����˳�����
    rsTmp.Filter = "��¼��Դ=3" '��ҳ����
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=2" '��Ժ�Ǽ�
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=1" '����
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=4" '������¼��
    
    'סԺ�ٰ���������˳�����
    If Not rsTmp.EOF And int��Դ = 2 Then
        strSQL = rsTmp.Filter
        rsTmp.Filter = strSQL & " And ����=3"
        If rsTmp.EOF Then rsTmp.Filter = strSQL & " And ����=2"
        If rsTmp.EOF Then rsTmp.Filter = strSQL & " And ����=1"
    End If
    
    strSQL = ""
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!�������) Then
            strSQL = strSQL & "��" & rsTmp!������� & IIf(Nvl(rsTmp!�Ƿ�����, 0) = 1, "������", "")
        End If
        rsTmp.MoveNext
    Loop
    
    GetPatiDiagnose = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadAdviceSignSource(ByVal int�������� As Integer, _
    ByVal lng����ID As Long, ByVal varTime As Variant, strIDs As String, _
    ByVal lngǩ��ID As Long, ByVal blnMoved As Boolean, strSource As String, _
    Optional ByVal lngǰ��ID As Long, Optional ByVal colSomeTime As Collection) As Integer
'���ܣ���ȡ�������ڵ���ǩ��/��֤��ҽ��Դ������
'������
'  int��������=Ҫǩ��/��֤ǩ����ҽ��״̬
'  ǩ��ʱ���룺
'    lng����ID
'    varTime=���˹Һŵ��Ż���ҳID
'    strIDs=ָ��Ҫǩ����ҽ��ID����(��ID)
'    lngǰ��ID=�¿�ҽ��Ҫǩ����ҽ����Դ(�Ƿ�ҽ��)
'    colSomeTime=ĳҽ����ʱ�����ݣ���ֹͣҽ��ǩ��ʱ���������ҽ��ִ����ֹʱ������ݣ�У��ʱ����У��ʱ������
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
                " Select /*+ Rule*/ A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID is Null And B.��������=1" & _
                " And A.ҽ��״̬=1 And Nvl(A.ǰ��ID,0)=[5]" & _
                " And Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))=[3]" & _
                " And Exists(Select M.���� From ��Ա�� M,ִҵ��� N" & _
                "       Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
                "         And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')" & _
                "   )" & _
                IIf(TypeName(varTime) = "String", " And A.����ID+0=[1] And A.�Һŵ�=[2]", " And A.����ID=[1] And A.��ҳID=[2]") & _
                IIf(str��IDs <> "", " And Nvl(A.���ID,A.ID) IN(Select Column_Value From Table(f_Num2list([4])))", "") & _
                " Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, varTime, UserInfo.����, str��IDs, lngǰ��ID)
        Else
            '��Ҫ���ϡ�ֹͣ��У�Ե�ҽ������ǩ�����¿�ʱǩ������ָ��ҽ������һ���ǵ�ǰҽ���´�
            strSQL = _
                " Select /*+ Rule*/ A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID is Not Null And B.��������=1" & _
                IIf(TypeName(varTime) = "String", " And A.����ID+0=[1] And A.�Һŵ�=[2]", " And A.����ID=[1] And A.��ҳID=[2]") & _
                IIf(str��IDs <> "", " And Nvl(A.���ID,A.ID) IN(Select Column_Value From Table(f_Num2list([3])))", "") & _
                " Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, varTime, str��IDs)
        End If
    Else
        '��֤ǩ��ʱ:�ȶ�ȡǩ��ʱ��Դ�����ɹ���
        strSQL = "Select ǩ������ From ҽ��ǩ����¼ Where ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "ҽ��ǩ����¼", "Hҽ��ǩ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngǩ��ID)
        If Not rsTmp.EOF Then intRule = Nvl(rsTmp!ǩ������, 1)
        '--
        strSQL = _
            " Select A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID=[1] Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
        If blnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngǩ��ID)
    End If
    
    'ҽ��Դ�ĵĲ�ͬ���ɹ���
    If intRule = 1 Then
        If int�������� = 3 Then
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,У��ʱ��,ִ������,������־,����ҽ��,����ʱ��"
        ElseIf int�������� = 8 Then
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
            If lngǩ��ID = 0 And int�������� = 3 And arrField(i) = "У��ʱ��" Then
                'У��ҽ��ǩ��ʱ,��У��ʱ�����⴦����������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                strLine = strLine & vbTab & colSomeTime("_" & Nvl(rsTmp!���ID, rsTmp!ID))
            ElseIf lngǩ��ID = 0 And int�������� = 8 And arrField(i) = "ִ����ֹʱ��" Then
                'ֹͣҽ��ǩ��ʱ,����ֹʱ�����⴦����������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                strLine = strLine & vbTab & colSomeTime("_" & Nvl(rsTmp!���ID, rsTmp!ID))
            Else
                If IsNull(rsTmp.Fields(arrField(i)).Value) Then
                    strLine = strLine & vbTab & ""
                Else
                    If Rec.IsType(rsTmp.Fields(arrField(i)).Type, adDBTimeStamp) Then
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

Public Function GetPatiDept(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal bytMode As Byte) As Long
'���ܣ���ȡ���˵�ǰ�����Ϳ���
'������bytMode=0-�����,1=�鲡��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & IIf(bytMode = 0, "��ǰ����id", "��ǰ����id") & " as ����ID" & vbNewLine & _
            "From ������Ϣ" & vbNewLine & _
            "Where ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then GetPatiDept = Val("" & rsTmp!����ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiLog(lng����ID As Long, lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˱䶯��¼
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ��ֹԭ��,��ֹʱ��,��ʼԭ��,Decode(��ʼԭ��, 1, '��Ժ��ס', 2, '��ס', 3," & _
            " Decode(��ʼʱ��, Null, 'ת��', 'ת����ס'), 4, '����', 5, '��λ�ȼ��䶯', 6, '����ȼ��䶯', 7," & vbNewLine & _
            "               '����ҽʦ�ı�', 8, '���λ�ʿ�ı�', 9, 'תΪסԺ����', 10, 'Ԥ��Ժ', 11, '����ҽʦ�䶯'," & _
            " 12, '����ҽʦ�䶯', 13, '�����䶯',14,'תҽ��С��',15,Decode(��ʼʱ��, Null, 'ת����', 'ת������ס')) ����" & vbNewLine & _
            "From ���˱䶯��¼" & vbNewLine & _
            "Where Nvl(���Ӵ�λ, 0) = 0 And ����id = [1] And ��ҳid = [2]" & vbNewLine & _
            "Order By ��ֹʱ�� Desc, ��ʼʱ�� Desc"
    Set GetPatiLog = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPati������Ϣ(ByVal lng����ID As Long, lng��ҳID As Long) As String
'���ܣ���ȡ��ǰ���˵ķ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lng�������� As Long
    
    strSQL = "Select �������� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPati������Ϣ", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        lng�������� = Val(rsTmp!�������� & "")
    End If
    strSQL = _
        " Select �������,Ԥ�����,0 as Ԥ�����,0 as ������ From ������� Where ����=1 And ����ID=[1] And ���� = [3]" & _
        " Union ALL" & _
        " Select 0,0,0, Sum(������) as ������ From ���˵�����¼ Where ����id = [1] And ��ҳid = [2] And ɾ����־ = 1 And (Sysdate <= ����ʱ�� Or ����ʱ�� Is Null)" & _
        " Union ALL" & _
        " Select 0,0,Sum(���),0 From ����ģ����� A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2]"
    strSQL = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ�����,Sum(Ԥ�����) as Ԥ�����,sum(������) as ������ From (" & strSQL & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPati������Ϣ", lng����ID, lng��ҳID, IIf(lng�������� = 1, 1, 2))
    If Not rsTmp.EOF Then
        GetPati������Ϣ = _
            "Ԥ�����:" & FormatEx(Nvl(rsTmp!Ԥ�����, 0), 2) & ",δ�����:" & FormatEx(Nvl(rsTmp!�������, 0), 2) & _
            IIf(Nvl(rsTmp!Ԥ�����, 0) <> 0, ",Ԥ�����:" & FormatEx(Nvl(rsTmp!Ԥ�����, 0), 2), "") & _
            ",ʣ���:" & FormatEx(Nvl(rsTmp!Ԥ�����, 0) - Nvl(rsTmp!�������, 0) + Nvl(rsTmp!Ԥ�����, 0), 2) & ",������:" & Nvl(rsTmp!������, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetסԺ����ҩռ��(ByVal lng����ID As Long, lng��ҳID As Long) As String
'���ܣ���ȡ��ǰ���˵�סԺ����ҩռ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select /*+ RULE */" & vbNewLine & _
            " c.���� As ����, Sum(Decode(a.�շ����, '5', a.ʵ�ս��, 0)) As ��ҩ��, Sum(Decode(a.�շ����, '6', a.ʵ�ս��, 0)) As ��ҩ��," & vbNewLine & _
            " Sum(Decode(a.�շ����, '7', a.ʵ�ս��, 0)) As ��ҩ��, Sum(Decode(a.�շ����, '5', 0, '6', 0, '7', 0, a.ʵ�ս��)) As ��ҩ��," & vbNewLine & _
            " Sum(a.ʵ�ս��) As ���з�" & vbNewLine & _
            "From סԺ���ü�¼ A, Table(f_Num2list2([1])) B, ������Ϣ C" & vbNewLine & _
            "Where a.����id = b.C1 And a.��ҳid = b.C2 And b.C1 = c.����id And a.��¼״̬ <> 0 Having Sum(a.ʵ�ս��) > 0" & vbNewLine & _
            "Group By c.����" & vbNewLine & _
            "Order By ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetסԺ����ҩռ��", lng����ID & ":" & lng��ҳID)
    If Not rsTmp.EOF Then
        GetסԺ����ҩռ�� = ",ҩռ��:" & Format((Val(rsTmp!��ҩ��) + Val(rsTmp!��ҩ��) + Val(rsTmp!��ҩ��)) / Val(rsTmp!���з�) * 100, "0.0") & "%"
    Else
        GetסԺ����ҩռ�� = ",ҩռ��:0.0%"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    GetסԺ����ҩռ�� = ",ҩռ��:0.0%"
End Function

Public Function LoadPatiAllergy(ByVal lng����ID As Long, Optional ByRef objCbo As Object, Optional ByRef rsAller As ADODB.Recordset) As Boolean
'���ܣ���ȡ���˵Ĺ�����¼����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
        
    strSQL = "Select Distinct B.����ʱ�� as �Һ�ʱ��,D.���� as �Һſ���,C.��ҳID,E.���� as סԺ����," & _
        " A.ҩ����,Nvl(A.����ʱ��,A.��¼ʱ��) as ����ʱ��,B.NO as �Һŵ�,A.ҩ��ID,A.����Դ����,A.������Ӧ,(max(lengthB(a.ҩ����)) over()-lengthB(a.ҩ����)+4) AS �ո񳤶�" & _
        " From ���˹�����¼ A,���˹Һż�¼ B,������ҳ C,���ű� D,���ű� E" & _
        " Where A.����ID=B.����ID(+) And A.��ҳID=B.ID(+) And B.��¼����(+)=1 And B.��¼״̬(+)=1" & _
        " And A.����ID=C.����ID(+) And A.��ҳID=C.��ҳID(+)" & _
        " And B.ִ�в���ID=D.ID(+) And C.��Ժ����ID=E.ID(+)" & _
        " And A.���=1 And ҩ���� is Not NULL And A.����ID=[1] And Not Exists" & vbNewLine & _
        " (Select ҩ��id" & vbNewLine & _
        "       From ���˹�����¼" & vbNewLine & _
        "       Where (Nvl(ҩ��id, 0) = Nvl(a.ҩ��id, 0) Or Nvl(ҩ����, 'Null') = Nvl(a.ҩ����, 'Null')) And Nvl(���, 0) = 0 And" & vbNewLine & _
        "             ��¼ʱ��>A.��¼ʱ�� And ����id = [1])" & _
        " Order by Nvl(A.����ʱ��,A.��¼ʱ��) Desc"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LoadPatiAllergy", lng����ID)
        
    If Not objCbo Is Nothing Then
        objCbo.Clear
        Do While Not rsTmp.EOF
'            If Not IsNull(rsTmp!�Һ�ʱ��) Then
'                strTmp = Format(rsTmp!����ʱ��, "yyyy-MM-dd") & "," & Nvl(rsTmp!ҩ����) & ",�������:" & Nvl(rsTmp!�Һſ���)
'            Else
'                strTmp = Format(rsTmp!����ʱ��, "yyyy-MM-dd") & "," & Nvl(rsTmp!ҩ����) & ",��" & rsTmp!��ҳID & "��סԺ:" & Nvl(rsTmp!סԺ����)
'            End If
            strTmp = Nvl(rsTmp!ҩ����) & String(Val(rsTmp!�ո񳤶�), " ") & Format(rsTmp!����ʱ��, "yyyy-MM-dd") & String(4, " ")

            If Not IsNull(rsTmp!������Ӧ) Then strTmp = strTmp & IIf(Nvl(rsTmp!������Ӧ) = "", "", "������Ӧ:" & rsTmp!������Ӧ)

            objCbo.AddItem strTmp
            
            rsTmp.MoveNext
        Loop
        If objCbo.ListCount = 0 Then
            objCbo.AddItem "�޼�¼"
        End If
        objCbo.ListIndex = 0
        objCbo.ForeColor = vbRed
    End If
    
    If Not rsAller Is Nothing Then Set rsAller = rsTmp
        
    LoadPatiAllergy = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRefuseReason(ByVal lng����ID As Long, lng��ҳID As Long) As String
'���ܣ���ȡ���˵Ĳ����ύ��������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '�Ըô�סԺ���һ�α��ܵ�Ϊ׼
    strSQL = "Select �������� From (Select �������� From �����ύ��¼ Where ����ID=[1] And ��ҳID=[2] And ��¼״̬=2 Order by ID Desc) Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRefuseReason", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then GetRefuseReason = Nvl(rsTmp!��������)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiMedRecHaveSubmit(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����ָ�����˵Ĳ����Ƿ��Ѿ��ύ(ͨ���ύ��¼)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '�Ըô�סԺ���һ�α��ܵ�Ϊ׼
    strSQL = "Select 1 From �����ύ��¼ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiMedRecHaveSubmit", lng����ID, lng��ҳID)
    PatiMedRecHaveSubmit = Not rsTmp.EOF
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

Private Function XWWebViewerOpen(lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��RIS��WEB Viewer
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    Dim strIp As String
    Dim strUrl As String
    
    On Error GoTo err
    
    strIp = zlDatabase.GetPara("XWWEB������IP", glngSys, 1288, "")
    
    If strIp <> "" Then
        strUrl = "C:\Program Files\Internet Explorer\iexplore.exe http://" & strIp & ":8080/imageweb/imageAction.action?ColID0=22&ColValue0=" & lngOrderID
        
        Shell strUrl, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        MsgBox "WEB������IP��ַΪ�գ��������ú�WEB��������", vbOKOnly, "��ʾ��Ϣ"
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
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

Public Function zlGetLocaleComputerNamePara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDeFault As String, _
        Optional strComputerName As String = "") As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ����������
    '��Σ�varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      lngSys=ʹ�øò�����ϵͳ��ţ���100
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
    '     strComputerName-��ȡָ������������
    '���Σ�
    '���أ�����ֵ���ַ�����ʽ
    '���ƣ����˺�
    '���ڣ�2010-06-07 13:56:22
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Integer, rsPara As ADODB.Recordset, rsUserPara As ADODB.Recordset
    Dim blnNew As Boolean, blnEnabled As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select ID,Nvl(����ֵ,ȱʡֵ) as ����ֵ,SYS_CONTEXT('USERENV','TERMINAL') as MName From zlParameters where ģ��=[1] and ϵͳ=[2]"
    If TypeName(varPara) = "String" Then
        strSQL = strSQL & " and ������=[3]"
    Else
        strSQL = strSQL & " and ������=[4]"
    End If
    Set rsPara = zlDatabase.OpenSQLRecord(strSQL, "GetPara", lngModual, lngSys, CStr(varPara), Val(varPara))
    If rsPara.EOF = False Then
        strSQL = _
            "   Select ����ֵ " & _
            "   From zlUserParas Where ����ID=[1]  and  ������=[2]"
        Set rsUserPara = zlDatabase.OpenSQLRecord(strSQL, "GetPara", Val(Nvl(rsPara!ID)), IIf(strComputerName = "", CStr(Nvl(rsPara!MName)), strComputerName))
         If Not rsUserPara.EOF Then
                zlGetLocaleComputerNamePara = Nvl(rsUserPara!����ֵ, strDeFault)
         Else
                zlGetLocaleComputerNamePara = Nvl(rsPara!����ֵ, strDeFault)
         End If
    Else
        zlGetLocaleComputerNamePara = strDeFault
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function CheckDoctorPatisIsValid() As Byte
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ҽ�����������Ƿ���Ч
    '���أ�0-����̨�������;1-ҽ����������;2-�ȷ���̨���ƽ�,����ҽ������
    '���ƣ����˺�
    '���ڣ�2010-06-07 14:32:47
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean, strComputerName As String

    '���˺�:Ӧ�����Ŷӽкŵĺ����˴�:��Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪ���������ŶӺ���վ��=1ʱ��Ч
     
     '��Ҫ����Ƿ�Ϊҽ���������з�ʽ
     '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
     blnValid = Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, p����������)) = 1
    If blnValid Then
         '����Ҫ���:�ŶӺ���վ��=1
         '�ŶӺ���վ��: 0-�������̨�������;1-����ҽ����������
         strComputerName = zlDatabase.GetPara("Զ�˺���վ��", glngSys, p�Ŷӽк�����ģ��)
        blnValid = Val(zlGetLocaleComputerNamePara("�ŶӺ���վ��", glngSys, p����������, "0", strComputerName)) = 1
    End If
    CheckDoctorPatisIsValid = blnValid
End Function

Public Sub PrintInMedRec(ByRef objClsMedRec As zlMedRecPage.clsInOutMedRec, ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
        ByRef objReport As Object, ByVal lng����ID As Long, ByRef objForm As Object, Optional intPage As Integer, Optional strPDFFile As String, _
          Optional ByRef objReportForm As Object)
'���ܣ���ҳ��ӡ��Ԥ��
'������intType=2����ӡ����=1��Ԥ����0=����,4-PDF;5-����Ƕ�봰�����
'     mobjReport-��ӡ������lng����ID-���˿��ң�mobjForm-������
'     intPage=1-4��ӡ��ҳ������ʽ��=5��ӡ����+��ҳ1��=6��ӡ����+��ҳ2
'     strPDFFile-intType=4 ʱ PDF���·��; intType=2 ʱ Ϊ��ӡ��
'     objReportForm-����Ƕ�봰��ı������
'    If lng����ID <> 0 Then
'        If objClsMedRec Is Nothing Then
'            Set objClsMedRec = New clsInOutMedRec
'            Call objClsMedRec.InitMedRec(gcnOracle, glngSys, gobjCommunity, gclsInsure)
'        End If
'        Call objClsMedRec.PrintOrPriviewInMedRec(intType, lng����ID, lng��ҳID, objReport, lng����ID, objForm, intPage)
'    End If
'    Exit Sub
    Dim strName As String
    Dim lngPage As Long
    
    If lng����ID <> 0 Then
        If objReport Is Nothing Then Set objReport = New clsReport
        Select Case Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
    
            Case 0 '��������׼
                If Sys.DeptHaveProperty(lng����ID, "��ҽ��") Then
                    strName = "ZL1_INSIDE_1261_4"
                Else
                    strName = "ZL1_INSIDE_1261_1"
                End If
            Case 1    '�Ĵ�ʡ��׼
                If Sys.DeptHaveProperty(lng����ID, "��ҽ��") Then
                    strName = "ZL1_INSIDE_1261_6"
                Else
                    strName = "ZL1_INSIDE_1261_5"
                End If
            Case 2    '����ʡ��׼
                If Sys.DeptHaveProperty(lng����ID, "��ҽ��") Then
                    strName = "ZL1_INSIDE_1261_8"
                Else
                    strName = "ZL1_INSIDE_1261_7"
                End If
            Case 3    '����ʡ��׼
                If Sys.DeptHaveProperty(lng����ID, "��ҽ��") Then
                    strName = "ZL1_INSIDE_1261_10"
                Else
                    strName = "ZL1_INSIDE_1261_9"
                End If
        End Select
        If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 0) = 0 And intPage = 0 Then
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 1)
        End If
        If intType = 0 Then
            Call ReportPrintSet(gcnOracle, glngSys, strName, objForm)
        Else
            If intPage = 5 Then
                lngPage = 1
            ElseIf intPage = 6 Then
                lngPage = 2
            Else
                lngPage = intPage
            End If
            If intType = 4 Then
                Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), "PDF=" & strPDFFile, intType)
                If intPage > 4 Then
                    Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), "PDF=" & strPDFFile, intType)
                End If
            ElseIf intType = 5 Then
                If strPDFFile <> "" Then
                    Call objReport.SetReportPrintSet(gcnOracle, glngSys, strName, "printer", strPDFFile) '����ָ����ӡ��
                End If
                Call objReport.LoadReport(gcnOracle, glngSys, strName, objForm, objReportForm, Nothing, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), 1)
                If intPage > 4 Then
                    Call objReport.LoadReport(gcnOracle, glngSys, strName, objForm, objReportForm, Nothing, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), 1)
                End If
            Else
                If intType = 2 And strPDFFile <> "" Then
                    Call objReport.SetReportPrintSet(gcnOracle, glngSys, strName, "printer", strPDFFile) '����ָ����ӡ��
                End If
                Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), intType)
                If intPage > 4 Then
                    Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), intType)
                End If
            End If
        End If
    End If
End Sub

Public Function CheckDiseaseFile(ByRef frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intCurDeptID As Long, _
ByVal str����IDs As String, ByVal str���IDs As String, Optional ByRef lngFileID As Long, Optional ByVal blnOnlyCheck As Boolean, Optional ByRef blnNo As Boolean, Optional ByVal lngFrom As Long = 1) As Boolean

'���ܣ���鲡����Щ����֤������û����д����ʾ������д
'����:frmParent    ������
'     lng����ID    ����ID
'     lng��ҳID    ���ﴫ�Һ�ID��סԺ����ҳID
'     intCurDeptID ��д��������ID
'     lngҽ��ID    ҽ��ID�����ڼ�鱨�棩
'     blnOnlyCheck true-ֻ���δ��д���������������б�,false-�����δ��д�����򵯳��б�
'     blnNo        �Ƿ�Ҫ��д��Ⱦ�����濨
'     lngFrom      ��Դ��1-���2-סԺ
   Dim rsTmp As ADODB.Recordset
   
   On Error GoTo errH
   
    If str����IDs = "" And str���IDs = "" Then
        CheckDiseaseFile = True
        Exit Function
    End If
    Dim strSQL As String
    If str����IDs <> "" Then
        strSQL = strSQL & " Union Select �ļ�ID From ��������ǰ�� Where ����ID IN (Select Column_Value From Table(f_Num2list([3])))"
    End If
    If str���IDs <> "" Then
        strSQL = strSQL & " Union Select �ļ�ID From ��������ǰ�� Where ���ID IN (Select Column_Value From Table(f_Num2list([4])))"
    End If
    On Error GoTo errH
    strSQL = "(" & Mid(strSQL, 8) & ") Minus Select �ļ�ID From ���Ӳ�����¼ Where ����ID=[1] And ��ҳID=[2] And ��������=5"
    strSQL = "Select /*+ Rule*/" & vbNewLine & _
            " a.Id, a.����, a.���, a.����, a.����, a.˵��" & vbNewLine & _
            "From �����ļ��б� A ,(" & strSQL & ") B Where A.ID=B.�ļ�ID  And nvl(A.����,0)=0 And " & vbNewLine & _
            "(a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����Ӧ�ÿ��� C Where c.�ļ�id = a.Id And c.����id = [5]))" & vbNewLine & _
            "Order By a.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDiseaseFile", lng����ID, lng��ҳID, str����IDs, str���IDs, intCurDeptID)
    blnNo = False
    
    If lngFrom = 1 Then
        If InStr(";" & GetPrivFunc(glngSys, p���ﲡ������) & ";", ";������д;") <= 0 Then
            rsTmp.Filter = "����=4"
        End If
    ElseIf lngFrom = 2 Then
        If InStr(";" & GetPrivFunc(glngSys, pסԺ��������) & ";", ";������д;") <= 0 Then
            rsTmp.Filter = "����=4"
        End If
    End If
        
    If rsTmp.RecordCount = 0 Then
        CheckDiseaseFile = True
        Exit Function
    Else
        strSQL = ""
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            strSQL = strSQL & vbCrLf & "��" & rsTmp!���� & "��"
            rsTmp.MoveNext
        Loop
    End If

    If rsTmp.RecordCount = 1 Then
        If blnOnlyCheck Then
            If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strSQL, 3) & vbCrLf & vbCrLf & "Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnNo = True: Exit Function
        Else
            If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strSQL, 3) & vbCrLf & vbCrLf & "Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                rsTmp.MoveFirst
                lngFileID = Val(rsTmp!ID & "")
            Else
                blnNo = True
            End If
        End If
    ElseIf rsTmp.RecordCount > 1 Then
        If blnOnlyCheck Then
            If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strSQL, 3) & vbCrLf & vbCrLf & "Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnNo = True: Exit Function
        Else
            If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strSQL, 3) & vbCrLf & vbCrLf & "Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                CheckDiseaseFile = True
                                Exit Function
            Else
                blnNo = True
            End If
        End If
    End If
    CheckDiseaseFile = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub Set�ٴ��Թ�ҩ(objFrom As Object)
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������(ZLCISBase)û����ȷ��װ���ù����޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.SetMedList(objFrom, gcnOracle, glngSys, gstrDBUser)
End Sub

Public Function CheckMecRed(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strfrmCation As String, Optional ByVal strOperateName As String) As Boolean
'���ܣ���鲡���Ƿ��Ѿ���Ŀ,�����Ƿ��ڴ������������(��ʱ��ҳ��������״̬���������޸�)
'       lng����ID:��ǰ����ID
'       lng��ҳID:��ǰ������ҳID
'       strfrmCation:���øú����Ĵ�������
'       strOperateName:���øú����Ĳ������ơ�strOperateNameΪ��ʱ����������ʾ
    Dim strSQL As String, rsTmp As Recordset
    Dim int����״̬ As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    '��ȡ����״̬
    strSQL = "Select Nvl(����״̬, 0) ����״̬ From ������ҳ Where ����id = [1] And ��ҳid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, strfrmCation, lng����ID, lng��ҳID)
    rsTmp.MoveFirst
    int����״̬ = rsTmp!����״̬
    '��ҳ���������ж�
    Select Case int����״̬
        Case 1 '�ȴ����
            strMsg = "�ò����ȴ������,����"
        Case 3 '�������
            strMsg = "�ò������������,����"
        Case 5 '���鵵
            strMsg = "�ò����Ѿ����鵵,����"
        Case 10 '���մ���
            strMsg = "�ò����ڽ��մ�����,����"
        Case Else '2-�ܾ����4-��鷴��;6-�������;13-���ڳ��;14-��鷴��;16-�������
            strMsg = ""
    End Select
    
    If strMsg = "" Then
        strSQL = "Select ��Ŀ���� from ������ҳ where ����ID=[1] And ��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, strfrmCation, lng����ID, lng��ҳID)
        If Not IsNull(rsTmp!��Ŀ����) Then
            strMsg = "�ò��˵Ĳ����Ѿ���Ŀ������"
        End If
    End If
    
    If strMsg <> "" Then  '������ҳ
        If strOperateName <> "" Then
            MsgBox strMsg & strOperateName & "��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    CheckMecRed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CanUnExec(ByVal datExec As Date, Optional ByVal datNow As Date) As Boolean
'���ܣ�����ִ�м�¼��ִ��ʱ���ж��ܷ�ȡ��ִ�л�ȡ�����
'������datExec=ִ�м�¼��ִ��ʱ��
'      datNow =��ǰʱ��
'���أ�CanUnExec=true-����ȡ��ִ�У�false-������ȡ��ִ��

    Dim lngDatDiff As Long
    If datExec <> CDate(Format("0", "yyyy-MM-dd HH:mm")) Then
        If datNow = CDate(0) Then
            datNow = zlDatabase.Currentdate
        End If
        lngDatDiff = DateDiff("D", datExec, datNow)
        CanUnExec = lngDatDiff <= gintҽ��ִ����Ч����
    Else
        CanUnExec = True
    End If
    
End Function

Public Function GetPatiDiagnoseByDept(ByVal lng����ID As Long, Optional ByVal intType As Integer = 1) As ADODB.Recordset
'���ܣ���ȡָ��������Ժ���˵������������
'������
'      lng����id=����id/����id
'      intType=0-��������ʾ��1-��������ʾ,Ĭ�ϰ�������ʾ
'���أ���¼��
    Dim strSQL As String
    
    strSQL = " Select A.����ID,A.�������, A.�������" & _
        " From ������ϼ�¼ A,������ҳ B,������Ϣ C,��Ժ���� R" & _
        " Where a.������� In (1, 2, 3, 11, 12, 13) And NVL(A.�������,1) = 1 And A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.����ID=C.����ID And C.��ҳID=B.��ҳID And C.����ID=R.����ID And C.��ǰ����ID=R.����ID " & _
        " And ��ϴ���=1 And" & IIf(intType = 1, " (R.����ID=[1] Or b.Ӥ������ID=[1])", " (r.����id = [1] Or b.Ӥ������id = [1])") & _
        " Order by A.����ID asc,A.��¼��Դ desc,A.������� desc"
    On Error GoTo errH
    Set GetPatiDiagnoseByDept = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub

Public Function ISPassShowCard() As Boolean
'���ܣ��Ƿ�������ʾ���￨��
'����:True ������ʾ,False ������
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnPassShowCard As Boolean
    
    On Error GoTo errHandle
    strSQL = "Select �������� From ҽ�ƿ���� where ����='���￨' and �Ƿ�̶�=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ����")
    If Not rsTemp.EOF Then
        blnPassShowCard = Nvl(rsTemp!��������) <> ""
    End If
    
    ISPassShowCard = blnPassShowCard
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ReadPatPricture(ByVal lng����ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '������lng����ID=��ȡָ�����˵���Ƭ
    '           imgPatient=��Ƭ����λ��
    '           strFile=��Ƭ�ı���·��
    '74421,������,2014-07-04,��ȡ������Ƭ��Ϣ
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = Sys.Readlob(glngSys, 27, lng����ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'���ܣ�֧�ֹ��ֵĹ���
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '���¹�
            zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '���Ϲ�
            zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
End Function

Public Function Get����ҽ����ӡ(ByVal lng����ID As Long, ByVal lng��ҳID) As Integer
'���ܣ��ж�ĳ�����˵�ҽ���Ƿ��ӡ��
'���أ�0-δ��ӡ��1-���ִ�ӡ��2-ȫ����ӡ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngҽ��ID As Long
    Dim dat���� As Date
    Dim bytPrint As Byte
    Dim blnDo As Boolean
    Dim arrӤ�� As Variant
    Dim strӤ�� As String
    Dim lngPrintType As Long
    Dim blnKey As Boolean
    Dim lng��� As Long
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = "select count(1) as ��ӡ from ����ҽ����ӡ a where a.����id=[1] and a.��ҳid=[2] and a.��ӡʱ�� is not null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If (rsTmp!��ӡ & "") = 0 Then
            Get����ҽ����ӡ = 0
            Exit Function
        End If
    End If
    
    strSQL = "select 1 from ����ҽ����ӡ a where a.����id=[1] and a.��ҳid=[2] and a.��ӡʱ�� is not null and Exists" & _
            " (select 1 from ����ҽ����ӡ where ����id=[1] and ��ҳid=[2] and ��ӡʱ�� is null and rownum<2) and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        Get����ҽ����ӡ = 1
        Exit Function
    End If
   
    Get����ҽ����ӡ = 1
    
    '�ж��ǲ���ȫ���Ѿ���ӡ
    '��ȡ�������ҽ����������
    lngPrintType = Val(zlDatabase.GetPara("ҽ������ӡģʽ", glngSys, pסԺҽ���´�))
    dat���� = CDate("1900-01-01")
    strSQL = "Select ҽ������ʱ�� as ʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID, lng��ҳID)
    If Not IsNull(rsTmp!ʱ��) Then dat���� = CDate(rsTmp!ʱ�� & "")
    
    strSQL = "Select ���,Ӥ������ From ������������¼ Where ����ID=[1] And ��ҳID=[2] Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID, lng��ҳID)
    strӤ�� = "0"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strӤ�� = strӤ�� & "," & rsTmp!���
            rsTmp.MoveNext
        Next
    End If
    arrӤ�� = Split(strӤ��, ",")
    
    For i = 0 To 1 '��������
        For j = 0 To UBound(arrӤ��) 'Ӥ��
            'ͣ����ӡ��ֻ��Ҫ�ж�һ��
            If i = 0 Then
                strSQL = "Select 1 From ����ҽ����ӡ A, ����ҽ����¼ B" & _
                    " Where A.ҽ��id=B.ID And A.��Ч = 0 And A.����id=[1] And A.��ҳid=[2] And Nvl(A.Ӥ��,0)=[3] And a.��ӡʱ�� is not null And (B.ȷ��ͣ��ʱ�� Is Not Null And" & _
                    " Not Exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� = 2) " & _
                    IIf(lngPrintType = 1, " Or B.ִ����ֹʱ�� Is Not Null And Not exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� in (1,2))  or b.У��ʱ�� is not null and not exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� in(1,2,3))", "") & ") And Rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID, lng��ҳID, Val(arrӤ��(j)))
                If Not rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            End If
        
            'δ��ӡ��Ҳû���� ����ҽ����ӡ �в���
            lngҽ��ID = 0
            lng��� = 0
            strSQL = "Select ҽ��id From (Select ҽ��id From ����ҽ����ӡ Where ����id =[1] And ��ҳid =[2] And Nvl(Ӥ��, 0)=[3] And ��Ч =[4]" & _
            " And ��ӡʱ�� + 0 >= [5] Order By ҳ�� Desc, �к� Desc) Where Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID, lng��ҳID, Val(arrӤ��(j)), i, dat����)
            If Not rsTmp.EOF Then lngҽ��ID = Val(rsTmp!ҽ��ID & "")
            ' lngҽ��id=0 ֻ������������һ��Ҳû��
            If lngҽ��ID <> 0 Then
                strSQL = "Select Nvl(Max(���), 0) as ��� From (Select ��� From ����ҽ����¼ Where (���id =[1] Or ID =[1])" & _
                    " Union All Select b.��� From ����ҽ����¼ A, ����ҽ����¼ B Where a.������� In ('5', '6') And a.Id =[1] And a.���id = b.Id)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lngҽ��ID)
                If Not rsTmp.EOF Then lng��� = Val(rsTmp!��� & "")
            End If
            
            If dat���� = CDate("1900-01-01") Then
                strSQL = "Select 1 From ����ҽ����¼ A, ������ĿĿ¼ B Where a.����id =[1] And a.��ҳid =[2] And Nvl(a.Ӥ��, 0) =[3] And" & vbNewLine & _
                        " a.������Ŀid = b.Id(+) And a.ҽ��״̬ Not In (-1, 2) and a.ҽ����Ч =[4] And" & vbNewLine & _
                        " ([6] = 1 And a.ҽ��״̬ = 1 Or a.ҽ��״̬ <> 1) And Nvl(a.���δ�ӡ, 0) = 0 And" & vbNewLine & _
                        " Not Exists (Select 1 From ����ҽ����¼ Where ������� = 'F' And ID = a.ǰ��id) And a.��� >[5] And a.������Դ = 2 and rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID, lng��ҳID, Val(arrӤ��(j)), i, lng���, lngPrintType)
                If Not rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            Else
                strSQL = "Select 1 From ����ҽ����¼ A, ������ĿĿ¼ B Where a.����id =[1] And a.��ҳid =[2] And Nvl(a.Ӥ��, 0) =[3] And" & vbNewLine & _
                        " a.������Ŀid = b.Id(+) And a.ҽ��״̬ Not In (-1, 2) and a.ҽ����Ч =[4] And" & vbNewLine & _
                        " ([6] = 1 And a.ҽ��״̬ = 1 Or a.ҽ��״̬ <> 1) And Nvl(a.���δ�ӡ, 0) = 0 And" & vbNewLine & _
                        " Not Exists (Select 1 From ����ҽ����¼ Where ������� = 'F' And ID = a.ǰ��id) And a.��� >[5] And a.������Դ = 2 and" & _
                        " Exists (Select 1 From ����ҽ��״̬ C Where a.Id = c.ҽ��id And c.����ʱ�� >=[7]) and rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng����ID, lng��ҳID, Val(arrӤ��(j)), i, lng���, lngPrintType, dat����)
                If Not rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            End If
        Next
        If blnKey Then Exit For
    Next
    
    If Not blnKey Then Get����ҽ����ӡ = 2
 
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
        Call SaveErrLog
End Function

Public Function Get������Һ��(ByVal lng����ID As Long, ByVal lng��ҳID) As String
'���ܣ���ȡָ�����˽�����������Һ������Ӧ���ز���  �ö��ŷָ� "200,300"
'˵����ҩƷ������λΪml����ý����ҩ;��Ϊ��Һ�� �������ѷ���+δ���ͣ���δ�����к��¿��ģ�
'      �����¿��Ļ��δ���͵�����ҽ����������㣺Ƶ��Ϊ��Ҫʱ�ĳ�����Ƶ��Ϊ��Ҫʱ��һ���Ե�����
'      �ݴ棬У�����ʣ������� ҽ��������ͳ�ƣ�����ʹ�ù���ͣ/���ù��ܵĲ��������ǣ�һ�ɵ���û����ͣ��
    Dim rsTmp As ADODB.Recordset
    Dim rsִ��ʱ�� As ADODB.Recordset
    Dim strSQL As String, str�ֽ�ʱ�� As String, strҽ��IDs As String
    Dim dblToday As Double, dblTomorrow As Double, dblTmp As Double
    Dim datCur As Date, datBegin As Date, datEnd As Date
    Dim lng���� As Long
    Dim i As Long, j As Long
    Dim varArr As Variant
    
    strSQL = "Select a.��������,a.�״�����,a.��ʼִ��ʱ��,a.�ϴ�ִ��ʱ��,Nvl(a.ִ����ֹʱ��,[4]) as ִ����ֹʱ��,a.Ƶ�ʼ��,a.ִ��ʱ�䷽��,a.Ƶ�ʴ���,a.�����λ,a.ִ��Ƶ��," & vbNewLine & _
        "     a.ҽ����Ч,a.ͣ��ʱ��,a.����,nvl(a.�ɷ����,d.סԺ�ɷ����) as ����,a.�ܸ�����,d.����ϵ��,a.���id" & vbNewLine & _
        "From ����ҽ����¼ A,������ĿĿ¼ B,ҩƷ���� C,ҩƷ��� D" & vbNewLine & _
        "Where a.������Ŀid = b.Id And b.Id = c.ҩ��id And a.�շ�ϸĿid=d.ҩƷid(+) And a.������� In ('5','6') and" & vbNewLine & _
        "     Upper(Nvl(b.���㵥λ,'NULL')) = 'ML' And c.��ý=1 And a.����id =[1] And a.��ҳid=[2] And" & vbNewLine & _
        "     a.��ʼִ��ʱ�� <= [4] And a.ҽ��״̬ Not In (-1,2,4) And" & vbNewLine & _
        "     (a.ҽ����Ч = 1 And" & vbNewLine & _
        "     (a.ִ��Ƶ�� = 'һ����' And a.��ʼִ��ʱ�� >= [3] Or" & vbNewLine & _
        "     a.ͣ��ʱ�� Is Null And a.ִ��ʱ�䷽�� Is Not Null Or" & vbNewLine & _
        "     a.ͣ��ʱ�� Is Not Null And a.ִ����ֹʱ�� >= [3] And (a.ִ��ʱ�䷽�� Is Not Null Or a.ִ��Ƶ�� = '��Ҫʱ')) Or" & vbNewLine & _
        "     a.ҽ����Ч = 0 And" & vbNewLine & _
        "     (a.�ϴ�ִ��ʱ�� Is Null And a.ִ��ʱ�䷽�� Is Not Null And Nvl(a.ִ����ֹʱ��,[3])>=[3] Or" & vbNewLine & _
        "     a.�ϴ�ִ��ʱ�� >= [3] ))"
    '��ʱ��Ҫ����˳�����7��ҩƷҽ����
    '1.Ƶ��Ϊһ���Ե��������ѷ��ͺ�δ���ͣ�
    '2.Ƶ��Ϊָ��������������δ���ͣ�
    '3.Ƶ��Ϊָ���������������ѷ��ͣ�
    '4.Ƶ��Ϊ��Ҫʱ���������ѷ��ͣ�
    '5.Ƶ��Ϊָ�������ĳ�������δ���ͣ�
    '6.Ƶ��Ϊ��Ҫʱ���������ٷ���һ�Σ�
    '7.Ƶ��Ϊָ�������ĳ��������ٷ���һ�Σ�
    '��������û�б����͹���ҽ����������Ҫʱ��������Ҫʱ��������ҽ�����û�з����򲻲�������㣬SQL��ѯ��Ҳ���ᱻ���˳���
    
    On Error GoTo errH
    datCur = zlDatabase.Currentdate
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get������Һ��", lng����ID, lng��ҳID, CDate(Format(datCur, "YYYY-MM-DD 00:00:00")), CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")))
    If rsTmp.EOF Then Get������Һ�� = "0,0": Exit Function
    
    '�� ҽ��ִ��ʱ�� ����ȡִ��ʱ���
    For i = 1 To rsTmp.RecordCount
        '3.Ƶ��Ϊָ���������������ѷ��ͣ�'6.Ƶ��Ϊ��Ҫʱ���������ٷ���һ�Σ�
        If Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!ͣ��ʱ�� & "" <> "" Or _
           Val(rsTmp!ҽ����Ч & "") = 0 And rsTmp!ִ��ʱ�䷽�� & "" = "" And rsTmp!�ϴ�ִ��ʱ�� & "" <> "" Then
           
            If InStr("," & strҽ��IDs & ",", "," & Val(rsTmp!���ID & "") & ",") = 0 Then strҽ��IDs = strҽ��IDs & "," & Val(rsTmp!���ID & "")
        End If
        rsTmp.MoveNext
    Next
    strҽ��IDs = Mid(strҽ��IDs, 2)
    If strҽ��IDs <> "" Then
        strSQL = "select a.ҽ��id,a.Ҫ��ʱ�� from ҽ��ִ��ʱ�� a where a.ҽ��id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) and a.Ҫ��ʱ��<=[2]"
        Set rsִ��ʱ�� = zlDatabase.OpenSQLRecord(strSQL, "Get������Һ��", strҽ��IDs, CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")))
    End If
    rsTmp.MoveFirst
    
    '��ʼ����
    For i = 1 To rsTmp.RecordCount
        '1.Ƶ��Ϊһ���Ե��������ѷ��ͺ�δ���ͣ�����ʼʱ�����һ��Ϊ׼ֻ��һ�Σ�����
        If Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��Ƶ�� & "" = "һ����" Then
            If Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD") = Format(datCur, "YYYY-MM-DD") Then
                dblToday = dblToday + Val(rsTmp!�������� & "")
            Else
                dblTomorrow = dblTomorrow + Val(rsTmp!�������� & "")
            End If
        '2.Ƶ��Ϊָ��������������δ���ͣ����ȼ�������ٷֽ�ʱ���
        ElseIf Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!ͣ��ʱ�� & "" = "" Then
            If Nvl(rsTmp!����, 0) <> 0 And Not IsNull(rsTmp!ִ��Ƶ��) Then
                'һ��Ƶ�����ڵĴ���
                If rsTmp!�����λ = "��" Then
                    lng���� = IntEx(rsTmp!���� * (rsTmp!Ƶ�ʴ��� / 7))
                ElseIf rsTmp!�����λ = "��" Then
                    lng���� = IntEx(rsTmp!���� * (rsTmp!Ƶ�ʴ��� / rsTmp!Ƶ�ʼ��))
                ElseIf rsTmp!�����λ = "Сʱ" Then
                    lng���� = IntEx(rsTmp!���� * (rsTmp!Ƶ�ʴ��� / rsTmp!Ƶ�ʼ��) * 24)
                ElseIf rsTmp!�����λ = "����" Then
                    lng���� = IntEx(rsTmp!���� * (rsTmp!Ƶ�ʴ��� / rsTmp!Ƶ�ʼ��) * (24 * 60))
                End If
            Else
                '�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,���ɷ�����һ����ʹ��ҩƷʱ���������ԣ����������ϵ����ֵȡ�����ı��������ҩ;���Ĵ�����
                '����һ��Ƶ�����ڵĴ�������
                If Nvl(rsTmp!����, 0) = 0 And Nvl(rsTmp!��������, 0) <> 0 Then
                    lng���� = IntEx(rsTmp!�ܸ����� * rsTmp!����ϵ�� / rsTmp!��������)
                ElseIf (Nvl(rsTmp!����, 0) = 1 Or Nvl(rsTmp!����, 0) = 2) And Nvl(rsTmp!��������, 0) <> 0 Then
                    lng���� = IntEx(rsTmp!�ܸ����� / IntEx(rsTmp!�������� / rsTmp!����ϵ��))
                Else
                    lng���� = Nvl(rsTmp!Ƶ�ʴ���, 0)
                End If
            End If
            If Not IsNull(rsTmp!ִ��ʱ�䷽��) Or Nvl(rsTmp!�����λ) = "����" Then
                str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsTmp!��ʼִ��ʱ��, CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")), "", Nvl(rsTmp!ִ��ʱ�䷽��), rsTmp!Ƶ�ʴ���, rsTmp!Ƶ�ʼ��, rsTmp!�����λ)
            End If
            If str�ֽ�ʱ�� <> "" Then
                varArr = Split(str�ֽ�ʱ��, ",")
                For j = 0 To UBound(varArr)
                    If Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + Val(rsTmp!�������� & "")
                    ElseIf Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + Val(rsTmp!�������� & "")
                    End If
                Next
            End If
        '3.Ƶ��Ϊָ���������������ѷ��ͣ����� ҽ��ִ��ʱ�� ����ִ��ʱ��㼴��
        ElseIf Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!ͣ��ʱ�� & "" <> "" Then
            If Not rsִ��ʱ�� Is Nothing Then
                rsִ��ʱ��.Filter = "ҽ��id=" & Val(rsTmp!���ID & "")
                For j = 1 To rsִ��ʱ��.RecordCount
                    If Between(Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + Val(rsTmp!�������� & "")
                    ElseIf Between(Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + Val(rsTmp!�������� & "")
                    End If
                    rsִ��ʱ��.MoveNext
                Next
            End If
        '4.Ƶ��Ϊ��Ҫʱ���������ѷ��ͣ�������ҽ��ֻ��һ�Σ��ҵ�����Ч��ֱ���ÿ�ʼʱ���жϼ���
        ElseIf Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��Ƶ�� & "" = "��Ҫʱ" And rsTmp!ͣ��ʱ�� & "" <> "" Then
            If Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD") = Format(datCur, "YYYY-MM-DD") Then
                dblToday = dblToday + Val(rsTmp!�������� & "")
            Else
                dblTomorrow = dblTomorrow + Val(rsTmp!�������� & "")
            End If
        '6.Ƶ��Ϊ��Ҫʱ���������ٷ���һ�Σ����� ҽ��ִ��ʱ�� ����ִ��ʱ��㣬����������Ҫ����ȡ����Сʱ��㣬���״�ִ��ʱ���
        ElseIf Val(rsTmp!ҽ����Ч & "") = 0 And rsTmp!ִ��ʱ�䷽�� & "" = "" And rsTmp!�ϴ�ִ��ʱ�� & "" <> "" Then
            If Not rsִ��ʱ�� Is Nothing Then
                rsִ��ʱ��.Filter = "ҽ��id=" & Val(rsTmp!���ID & "")
                rsִ��ʱ��.Sort = "Ҫ��ʱ��"
                For j = 1 To rsִ��ʱ��.RecordCount
                    dblTmp = Val(rsTmp!�������� & "")
                    If j = 1 And Val(rsTmp!�״����� & "") <> 0 Then dblTmp = Val(rsTmp!�״����� & "")
                    If Between(Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + dblTmp
                    ElseIf Between(Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + dblTmp
                    ElseIf Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS") > Format(datCur + 1, "YYYY-MM-DD 23:59:59") Or _
                        Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS") > Format(rsTmp!ִ����ֹʱ�� & "", "YYYY-MM-DD HH:MM:SS") Then
                        Exit For
                    End If
                    rsִ��ʱ��.MoveNext
                Next
            End If
        '7.Ƶ��Ϊָ�������ĳ��������ٷ���һ�Σ�
        '5.Ƶ��Ϊָ�������ĳ�������δ���ͣ�
        '7��5��һ���Ĵ���ʽ�����¼���ֽ�ʱ���
        ElseIf Val(rsTmp!ҽ����Ч & "") = 0 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!�ϴ�ִ��ʱ�� & "" <> "" Or _
            Val(rsTmp!ҽ����Ч & "") = 0 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!�ϴ�ִ��ʱ�� & "" = "" Then
            '����״�������Ϊ0����ʼʱ����ҽ����ʼִ��ʱ��Ϊ׼Ϊ�˼�����״�ִ��ʱ��������ж�
            If Val(rsTmp!�״����� & "") = 0 And Format(datCur, "YYYY-MM-DD 00:00:00") > Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS") Then
                datBegin = Format(datCur, "YYYY-MM-DD 00:00:00")
            Else
                datBegin = rsTmp!��ʼִ��ʱ��
            End If
        
            If Format(rsTmp!ִ����ֹʱ�� & "", "YYYY-MM-DD HH:MM:SS") > Format(datCur + 1, "YYYY-MM-DD 23:59:59") Then
                datEnd = CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59"))
            Else
                datEnd = CDate(Format(rsTmp!ִ����ֹʱ�� & "", "YYYY-MM-DD HH:MM:SS"))
            End If
            
            str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, datEnd, "", Nvl(rsTmp!ִ��ʱ�䷽��), Nvl(rsTmp!Ƶ�ʴ���, 0), Nvl(rsTmp!Ƶ�ʼ��, 0), Nvl(rsTmp!�����λ), rsTmp!��ʼִ��ʱ��)
            If str�ֽ�ʱ�� <> "" Then
                varArr = Split(str�ֽ�ʱ��, ",")
                For j = 0 To UBound(varArr)
                    dblTmp = Val(rsTmp!�������� & "")
                    If j = 0 And Val(rsTmp!�״����� & "") <> 0 Then dblTmp = Val(rsTmp!�״����� & "")
                    If Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + dblTmp
                    ElseIf Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + dblTmp
                    End If
                Next
            End If
        End If
        rsTmp.MoveNext
    Next
    Get������Һ�� = dblToday & "," & dblTomorrow
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function HaveOperateAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intType As Integer) As Boolean
'���ܣ��ж�ָ�������Ƿ񻹴��ڿ��Բ�����ҽ��
'������intType 0-У�ԣ�1��ȷ��ֹͣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    
    On Error GoTo errH
    If intType = 0 Then
        If gblnKSSStrict Or gbln�����ּ����� Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ Then
            strWhere = strWhere & " And (Nvl(A.���״̬,0) Not in(1,3,7" & IIf(gblnѪ��ϵͳ = True, "", ",4,5") & ") or a.ҽ����Ч=0 and a.���״̬=1 and a.������־=1 and (instr(',5,6,',A.�������)>0 or A.�������='E' and B.��������='2'))"
        End If
        strSQL = "select 1 from ����ҽ����¼ a,������ĿĿ¼ b where a.������Ŀid=b.id(+) and A.ҽ��״̬=1 and a.����id=[1] and a.��ҳid=[2]" & strWhere & _
                " And Exists ( Select M.���� From ��Ա�� M,ִҵ��� N" & _
                " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
                " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')) And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3 and Rownum<2"
    ElseIf intType = 1 Then
        strSQL = "select 1 from ����ҽ����¼ a where A.ҽ��״̬=8 and Nvl(a.ҽ����Ч,0)=0 And a.����id=[1] and a.��ҳid=[2] And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3 and Rownum<2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob.HaveOperateAdvice", lng����ID, lng��ҳID)
    HaveOperateAdvice = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub PlugInInSideBar(ByRef cbsMain As Object, ByVal strFuncName As String, Optional ByVal intInSide As Integer)
'���ܣ����ù�������ť����ҿ�Ƭ�����еĹ��ܰ�ť
'intInSide ��ҪҪ��ӹ�������ť��Ĭ����Ҫ���
    Dim objControl As CommandBarControl
    Dim objMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim varArr As Variant
    Dim strTmp As String
    Dim lngTmp As Long
    Dim objCbs As Object
    Dim lngidx As Long
    Dim i As Long
    Dim strName As String, lngIcon As Long
    
    If strFuncName = "" Then Exit Sub
    varArr = Split(strFuncName, "|")
    
    Set objCbs = cbsMain
    
    '��չ:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenuBar = objCbs.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenuBar Is Nothing Then Set objMenuBar = objCbs.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
 
    Set objMenuBar = objCbs.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ(&A)", objMenuBar.Index + 1, False)
 
    With objMenuBar.CommandBar.Controls
        For i = 0 To UBound(varArr)
            strTmp = varArr(i)
            
            strName = strTmp
            lngIcon = conMenu_Tool_PlugIn
            
            If InStr(strTmp, ",") > 0 Then
                strName = Split(strTmp, ",")(0)
                lngIcon = Val(Split(strTmp, ",")(1))
            End If
 
            Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + 1 + i, strName)
                objControl.IconId = lngIcon
                objControl.ToolTipText = strName
                objControl.Style = xtpButtonIconAndCaption
        Next
    End With
    
    If intInSide = 0 Then
        '���������
        '�ҵ�Ҫ��ӵ�λ��
        For Each objControl In objCbs(2).Controls '�����ǰ������һ��Control
            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
                Set objControl = objCbs(2).Controls(objControl.Index - 1)
                lngidx = objControl.Index
                Exit For
            End If
        Next
        
        With objCbs(2).Controls
            For i = UBound(varArr) To 0 Step -1
                strTmp = varArr(i)
                
                strName = strTmp
                lngIcon = conMenu_Tool_PlugIn
                
                If InStr(strTmp, ",") > 0 Then
                    strName = Split(strTmp, ",")(0)
                    lngIcon = Val(Split(strTmp, ",")(1))
                End If
                
                Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + 1 + i, strName, lngidx + 1)
                    objControl.IconId = lngIcon
                    objControl.ToolTipText = strName
                    objControl.Style = xtpButtonIconAndCaption
            Next
        End With
    End If
    cbsMain.RecalcLayout
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function Get��Ⱦ��״̬(ByVal lng��¼ As Long, ByVal lng��д As Long, ByVal lng״̬ As Long) As String
'���ܣ���ȡ��ǰ���˵Ĵ�Ⱦ��״̬
    Dim strTmp As String
    If lng״̬ <> 0 Then
'       -1-�Ѿܾ� 1-�ѽ���;2-�ѳʱ�;3-���ͨ��;4-��ҽ�����ޣ�5-ҽ���ѷ�����ɴ����
        Select Case lng״̬
        Case -1
            strTmp = "�Ѿܾ�"
        Case 1
            strTmp = "�ѽ���"
        Case 2
            strTmp = "�ѳʱ�"
        Case 3
            strTmp = "���ͨ��"
        Case 4
            strTmp = "��ҽ���޸�"
        Case 5
            strTmp = "ҽ������Ĵ����"
        End Select
    ElseIf lng��д > 0 Then
        strTmp = "����д"
    ElseIf lng��¼ > 0 Then
        strTmp = "�����Խ��"
    End If
    Get��Ⱦ��״̬ = strTmp
End Function

Public Sub FuncEPRReport(frmMain As Form, ByVal lngҽ��ID As Long, ByVal str������� As String, _
        Optional ByVal lng����ID As Long, Optional ByVal str��鱨��ID As String, _
        Optional ByVal int���� As Integer = 1)
'���ܣ����ı���
'����: int����:1-����һ��,2-סԺһ��
    Dim strPrivs As String, strSQL As String
    Dim lngRet As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    '�������ݽ���ƽ̨����LIS,PACS���ı���
    If gobjExchange Is Nothing Then
        On Error Resume Next
        Set gobjExchange = CreateObject("zlExchange.clsExchange")
        If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
        err.Clear: On Error GoTo 0
    End If
    If Not gobjExchange Is Nothing Then
        '�����д���ǲɼ��������������ΪE��������ֻ�жϼ����
        Call gobjExchange.SendMsg(IIf(str������� = "D", 4, 3), "ҽ��ID::" & lngҽ��ID & "||����Ա����::" & UserInfo.���� & "||����Աȱʡ����::" & UserInfo.������)
        Exit Sub
    End If
    strPrivs = GetInsidePrivs(IIf(int���� = 1, p����ҽ���´�, pסԺҽ���´�))
    '���ж��Ƿ���Լ�������
    Select Case CheckEPRReport(lngҽ��ID, lng����ID)
    Case 0
        MsgBox "��ҽ���ı���û����д��", vbInformation, gstrSysName
        Exit Sub
    Case 2
        If InStr(strPrivs, ";����δ��ɱ���;") > 0 Then
            MsgBox "ע�⣺��ҽ���ı��滹û����ʽǩ����", vbInformation, gstrSysName
        Else
            MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)����û��Ȩ�޲�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End Select
    
    If str������� = "D" Then
        If HaveRIS Then
            'RIS�������
            strSQL = "select 1 from ����ҽ������ a,���Ӳ�����¼ b,���Ӳ�����ʽ c where a.����id=b.id and b.id=c.�ļ�id and a.ҽ��id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmMain.Caption, lngҽ��ID)
            If rsTmp.EOF Then
                lngRet = gobjRis.ShowViewReport(frmMain.hwnd, lngҽ��ID, InStr(strPrivs, ";�����ӡ;") > 0)
                If lngRet = 0 Then Exit Sub
            End If
        End If
    End If
    'ִ�в���
    '�°�PACS���棬ֱ��ǿ��ʹ���°�PACS����༭��
    If str��鱨��ID <> "" Then
        If CreateObjectPacs(gobjPublicPacs) Then
            Call gobjPublicPacs.zlDocShowReport(lngҽ��ID, , False, frmMain)
        End If
    Else
        '���ı���
        Call gobjRichEPR.ViewDocument(frmMain, lng����ID, False)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function InitNurseIntegrate(Optional blnMsg As Boolean = False) As Boolean
'�ж�������廤����Ϊ�վͳ�ʼ��
    If gobjNurseIntegrate Is Nothing Then
        On Error Resume Next
        Set gobjNurseIntegrate = CreateObject("zlNurseIntegrate.clsNurseIntegrate")
        If Not gobjNurseIntegrate Is Nothing Then
            If gobjNurseIntegrate.zlInitCommon(gcnOracle, gstrDBUser) = False Then
                Set gobjNurseIntegrate = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    If blnMsg = True And gobjNurseIntegrate Is Nothing Then
        MsgBox "���廤��ӿڲ�����zlNurseIntegrate  ����ʧ�ܣ�", vbInformation, gstrSysName
    End If
    InitNurseIntegrate = Not gobjNurseIntegrate Is Nothing
End Function



Public Function CheckStructAddr(ByVal objCtl As PatiAddress, ByVal lngLen As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ṹ����ַ�ؼ��е���Ϣ¼���Ƿ���ȷ
    '���:objCtl-�ṹ����ַ�ؼ���lngLen-���Ƴ���
    '����:True-������Ϣ�Ϸ�
    '����:���ϴ�
    '����:2015-12-7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If zlCommFun.ActualLen(objCtl.Value) > lngLen Then
        MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "���ֻ������" & lngLen \ 2 & "������,���顣", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    If objCtl.CheckNullValue(, True, False) <> "" Then
        MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "��" & objCtl.CheckNullValue & "��δ����,���顣", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    CheckStructAddr = True
End Function

Public Function zlReadAddrInfo(ByVal objCtrl As PatiAddress, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                               ByVal intType As Integer, Optional ByVal strAddress As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���Ĳ��˵�ַ��Ϣ���ؼ���
    '���:objCtrl-�ṹ����ַ�ؼ�,intType -��ַ����1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ
    '����:
    '����:���ϴ�
    '����:2015/12/3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    strSQL = "Select ʡ,��,��,����,���� From ���˵�ַ��Ϣ Where ����ID=[1] and Nvl(��ҳID,0)=[2] and ��ַ���=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ṹ����ַ", lng����ID, lng��ҳID, intType)
    If rsTmp.RecordCount > 0 Then
        Call objCtrl.LoadStructAdress(Nvl(rsTmp!ʡ), Nvl(rsTmp!��), Nvl(rsTmp!��), Nvl(rsTmp!����), Nvl(rsTmp!����))
    Else
        objCtrl.Value = strAddress
    End If
    zlReadAddrInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function FuncTraReaction(ByVal lngҽ��ID As Long, ByVal lngMoudle As Long, ByVal blnMoved As Boolean, Optional ByVal lng�շ�id As Long) As Boolean
'���ܣ���Ѫ��Ӧ
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, lng��ҳID As Long, lng������Դ As Long
    If InitObjBlood(True) = False Then Exit Function
    On Error GoTo errH
    strSQL = "Select B.����ID,B.��ҳID,B.������Դ,A.ID ����ID From ���˹Һż�¼ A,����ҽ����¼ B where B.�Һŵ�=A.NO(+) And  B.id=[1]"
    If blnMoved = True Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����Ӧ", lngҽ��ID)
    lng����ID = Val("" & rsTemp!����ID)
    If IsNull(rsTemp!��ҳID) Then
        lng��ҳID = Val("" & rsTemp!����ID)
    Else
        lng��ҳID = Val("" & rsTemp!��ҳID)
    End If
    lng������Դ = Val("" & rsTemp!������Դ)
    Call gobjPublicBlood.zlShowBloodReaction(Nothing, glngSys, lngMoudle, 1, lng����ID, lng��ҳID, lng������Դ, 1, lng�շ�id)
    FuncTraReaction = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FuncTraReactionRecord(ByVal frmParent As Object, ByVal lng���� As Long, ByVal lngMoudle As Long) As Boolean
'���ܣ���Ѫ��Ӧ���ýӿ�
    On Error GoTo errH
    If InitObjBlood(True) = False Then Exit Function
    Call gobjPublicBlood.zlShowBloodReactionRecord(frmParent, glngSys, lngMoudle, lng����)
    FuncTraReactionRecord = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckZLPass(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��жϵ�ǰ�����Ƿ���δ���ͨ����ҽ��
'���أ�true ͨ����false����δͨ��ҩƷҽ��

    Dim rsTmp As ADODB.Recordset
    Dim strErr As String
    Dim blnTmp As Boolean
    
    CheckZLPass = True
    
    
    On Error Resume Next
    
    If gobjPass Is Nothing Then
        Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "������ҩ���", True)
    End If
    err.Clear
    
    On Error GoTo errH
    
    If Not gobjPass Is Nothing Then
        blnTmp = gobjPass.ZLPharmReviewResultView(lng����ID, lng��ҳID, rsTmp, strErr)
    End If
    
    If strErr = "" Then
        If blnTmp Then
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF Then
                    CheckZLPass = False
					Call gobjPass.ZLPharmReviewResultShow(frmParent, lng����ID, lng��ҳID)
                End If
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function