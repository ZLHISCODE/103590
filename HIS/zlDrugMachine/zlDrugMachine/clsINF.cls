VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'˵�����������ӿڻ���
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Public mobjComLib As Object
Public mobjLog As clsLog
Public mstrNOs As String        '��ʽ������1,������1[[;����2,������2]...]

Private mstrURL As String
Private mblnAction As Boolean
Private mstrKey As String
Private mstrToken As String
Private mstrUser As String
Private mstrInfCode As String

'ZLHIS���ݿ��ʻ���
Public Property Get UserName() As String
    UserName = mstrUser
End Property
Public Property Let UserName(ByVal strUser As String)
    mstrUser = strUser
End Property

'�ӿ�����״̬��True���ӣ�Falseδ����
Public Property Get Action() As Boolean
    Action = mblnAction
End Property
Public Property Let Action(ByVal blnValue As Boolean)
    mblnAction = blnValue
End Property

'OLEDB���Ӵ�������WebService�����ַ��Ŀǰֻ֧��������
Public Property Get URL() As String
    URL = mstrURL
End Property
Public Property Let URL(ByVal strValue As String)
    mstrURL = strValue
End Property

'��Ϣ����ƽ̨��Ҫ��KEY
Public Property Get KEY() As String
    KEY = mstrKey
End Property
Public Property Let KEY(ByVal strKey As String)
    mstrKey = strKey
End Property

'��Ϣ����ƽ̨��Ҫ��Token
Public Property Get Token() As String
    Token = mstrToken
End Property
Public Property Let Token(ByVal strToken As String)
    mstrToken = strToken
End Property

Public Function GetData(ByVal intType As Integer, ByVal strData As String, ByRef blnFinish As Boolean) As ADODB.Recordset
'���ܣ���ȡZLHIS�ı�׼����
'������
'  intType��ҵ�����
'           1-������Ϣ��
'           2-��Ա��Ϣ��
'           3-ҩƷĿ¼��
'           4-ҩƷ������λ��
'           5-��ҩ���ڣ�
'           21-��ҩ[�����סԺ������ϸ�ϴ�]��
'           22-��ʼ��ҩ��
'           23-��ɷ�ҩ��
'           24-����(����/����)��ҩ��
'           25-����������ҩ��
'  strData��ҵ������
'           ҵ�����=1���ӿڱ��|��������1;��������2;��
'           ҵ�����=2���ӿڱ��|��Ա����1;��Ա����2;��
'           ҵ�����=3���ӿڱ��|���ͱ���1;���ͱ���2;��
'           ҵ�����=4���ӿڱ��|�ⷿid1;�ⷿid2;��
'           ҵ�����=5���ӿڱ��|�ⷿid1;�ⷿid2;��
'           ҵ�����=21 to 23��
'               ���1|����1,������1;����2,������2;��
'               סԺ��2|�շ�ID1;�շ�ID2;��
'           ҵ�����=24���շ�ID1,��ҩ����1;�շ�ID2,��ҩ����2;��   ��ҩ�����ĵ�λΪZLHISҩƷ��С��λ
'           ҵ�����=25������1,������1,�ⷿid1;����2,������2,�ⷿid2;��
'  blnFinish��True������ɣ�False�쳣�ж�
'���أ���¼������

    Dim objDB As Object
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strINF As String, strInfo As String, strDetail As String
    Dim strTmp As String
    Dim intIO As Integer
    Dim arrTmp As Variant, arrInfo As Variant, arrDetail As Variant, arrNew As Variant
    Dim l As Long
    
    mobjLog.Add "׼����ȡ��׼����", 1

    If TypeName(mobjComLib) = "clsComLib" Then
        'ZLHIS����������ҩƷ�Զ����豸��������
        Set objDB = mobjComLib.zlDatabase
    ElseIf TypeName(mobjComLib) = "clsPublic" Then
        '֧����
        Set objDB = mobjComLib
    Else
        mobjLog.Add "�������Ͳ���ȷ����ֹ��ȡ��׼����", 1
        mobjLog.Add "" & intType, 1, 1
        mobjLog.Save
        Exit Function
    End If
    
    '���strData����
    mobjLog.Add strData, 1, 1
    
    arrTmp = Split(strData, "|")
    Select Case intType
    Case 1 To 5
        If UBound(arrTmp) = 1 Then
            strINF = arrTmp(0)          '�ӿڱ���
            strInfo = Trim(arrTmp(1))   'ҵ������
        ElseIf UBound(arrTmp) = 0 Then
            strINF = arrTmp(0)
            strInfo = ""
        Else
            mobjLog.Add "���������ݡ����������ݣ���ֹ��ȡ��׼����", 1
            mobjLog.Save
            Exit Function
        End If
    
    Case 21 To 23
        If UBound(arrTmp) = 1 Then
            If Trim(arrTmp(1)) = "" Then
errExit:
                mobjLog.Add "��ҵ�����ݡ������޵�����Ϣ����ֹ��ȡ��׼����", 1
                mobjLog.Save
                Exit Function
            Else
                intIO = Val(arrTmp(0))  '���סԺ
                strInfo = arrTmp(1)     'ҵ������
                If intIO < 1 Or intIO > 2 Then
                    mobjLog.Add "��ҵ�����ݡ������޷�ʶ��������סԺ����ֹ��ȡ��׼����", 1
                    mobjLog.Save
                    Exit Function
                End If
            End If
        Else
            GoTo errExit
        End If
    
    Case 24
        If UBound(arrTmp) = 0 Then
            If Trim(arrTmp(0)) = "" Then
                GoTo errExit
            Else
                If Not arrTmp(0) Like "*,*" Then
                    mobjLog.Add "��ҵ�����ݡ������޷�ʶ����ҩ���ݣ���ֹ��ȡ��׼����", 1
                    mobjLog.Save
                    Exit Function
                End If
            End If
            strInfo = arrTmp(0)
            intIO = 1
        Else
            GoTo errExit
        End If
        
    Case 25
        If UBound(arrTmp) = 0 Then
            If Trim(arrTmp(0)) = "" Then
                GoTo errExit
            Else
                If Not arrTmp(0) Like "*,*,*" Then
                    mobjLog.Add "��ҵ�����ݡ������޷�ʶ��������ҩ���ݣ���ֹ��ȡ��׼����", 1
                    mobjLog.Save
                    Exit Function
                End If
            End If
            
            '��������,������,�ⷿid;...��ת�ɡ�������,�ⷿid;...����ʽ
            arrNew = Split(arrTmp(0), ";")
            strTmp = ""
            For l = LBound(arrNew) To UBound(arrNew)
                If arrNew(l) Like "*,*" Then
                    strTmp = strTmp & ";" & Mid(arrNew(l), InStr(arrNew(l), ",") + 1)
                End If
            Next
            If Left(strTmp, 1) = ";" Then strTmp = Mid(strTmp, 2)
            Erase arrNew
            
            strInfo = strTmp
            intIO = 1
        Else
            GoTo errExit
        End If
        
    Case Else
        mobjLog.Add "��ҵ����𡱲�������ȷ����ֹ��ȡ��׼����", 1
        mobjLog.Save
        Exit Function
        
    End Select
    
    On Error GoTo hErr
    
    mobjLog.Add "strInfo��" & strInfo, 1
    
    '��ҵ������ȡ��׼����
    Select Case intType
    Case Val("1-������Ϣ")
        strSQL = _
            "Select b.Id, b.�ϼ�id, b.����, b.����, b.����," & vbNewLine & _
            "    f_List2str(Cast(Collect(Cast(a.������� As Varchar2(20))) As t_Strlist)) �������," & vbNewLine & _
            "    f_List2str(Cast(Collect(a.��������) As t_Strlist)) ��������, " & vbNewLine & _
            "    Decode(To_Char(Nvl(b.����ʱ��, To_Date('3000-1-1', 'yyyy-mm-dd')), 'yyyy'), '3000', 0, 1) ͣ�� " & vbNewLine & _
            "From ��������˵�� A, ���ű� B " & _
            IIf(strInfo = "", "", ", Table(f_Str2list([1], ';')) C ") & vbNewLine & _
            "Where a.����id(+) = b.Id " & _
            IIf(strInfo = "", "", " And a.�������� = c.Column_Value ") & vbNewLine & _
            "Group By b.Id, b.�ϼ�id, b.����, b.����, b.����, b.����ʱ�� "
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡ������Ϣ", strInfo)
        mobjLog.Add "��ȡ������Ϣ���", 1
        
    Case Val("2-��Ա��Ϣ")
        strSQL = _
            "Select b.Id, b.���, b.����, b.����, b.��������, b.�Ա�, b.����," & vbNewLine & _
            "    f_List2str(Cast(Collect(a.��Ա����) As t_Strlist)) ��Ա����," & vbNewLine & _
            "    f_List2str(Cast(Collect(Cast(c.����id As Varchar2(20))) As t_Strlist)) ����id," & vbNewLine & _
            "    f_List2str(Cast(Collect(d.����) As t_Strlist)) ���ű���, " & vbNewLine & _
            "    Decode(To_Char(Nvl(b.����ʱ��, To_Date('3000-1-1', 'yyyy-mm-dd')), 'yyyy'), '3000', 0, 1) ͣ�� " & vbNewLine & _
            "From ��Ա����˵�� A, ��Ա�� B, ������Ա C, ���ű� D " & _
            IIf(strInfo = "", "", ", Table(f_Str2list([1], ';')) E") & vbNewLine & _
            "Where a.��Աid(+) = b.Id And b.Id = c.��Աid(+) And c.����id = d.Id " & _
            IIf(strInfo = "", "", " And a.��Ա���� = e.Column_Value ") & vbNewLine & _
            "Group By b.Id, b.���, b.����, b.����, b.��������, b.�Ա�, b.����, b.����ʱ�� "
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡ��Ա��Ϣ", strInfo)
        mobjLog.Add "��ȡ��Ա��Ϣ���", 1
        
    Case Val("3-ҩƷĿ¼")
        strSQL = _
            "Select Decode(Nvl(To_Char(a.����ʱ��, 'yyyy'), '3000'), '3000', 0, 1) ͣ��, a.Id, a.����, a.���� ͨ����," & vbNewLine & _
            "    a.���, a.���㵥λ, c.ҩ�ⵥλ, c.סԺ��λ, c.���ﵥλ, c.ҩ���װ, c.סԺ��װ, c.�����װ, c.����ϵ��," & vbNewLine & _
            "    c.��׼�ĺ�, d.ҩƷ����, d.��������, e.���㵥λ ������λ, f.���� �����̱���, f.���� ������," & vbNewLine & _
            "     Decode(a.���, '5', '����ҩ', '6', '�г�ҩ', '�в�ҩ') ҩƷ���," & vbNewLine & _
            "     Max(Decode(b.����, 1, b.����, Null)) ͨ��������, Max(Decode(b.����, 3, b.����, Null)) ��Ʒ��," & vbNewLine & _
            "     g.���� Ӣ���� " & vbNewLine & _
            "From �շ���ĿĿ¼ A, �շ���Ŀ���� B, ҩƷ��� C, ҩƷ���� D, ������ĿĿ¼ E, ҩƷ������ F, ������Ŀ���� G " & vbNewLine & _
            IIf(strInfo = "", "", ", Table(f_Str2list([1], ';')) H ") & _
            "Where a.Id = b.�շ�ϸĿid(+) And a.Id = c.ҩƷid And c.ҩ��id = d.ҩ��id And c.ҩ��id = e.Id " & vbNewLine & _
            "    And a.���� = f.����(+) And e.Id = g.������Ŀid(+) " & _
            IIf(strInfo = "", "", " And d.ҩƷ���� = h.Column_Value ") & vbNewLine & _
            "    And g.����(+) = 2 And g.����(+) = 0 And a.��� In ('5', '6', '7') And b.����(+) = 1" & vbNewLine & _
            "Group By a.����ʱ��, a.Id, a.����, a.����, a.���, a.���㵥λ, a.����, c.ҩ�ⵥλ, c.סԺ��λ, c.���ﵥλ," & vbNewLine & _
            "    c.ҩ���װ, c.סԺ��װ, c.�����װ, c.����ϵ��, c.��׼�ĺ�, d.ҩƷ����, d.��������, e.���㵥λ, a.���," & vbNewLine & _
            "    f.����, f.����, g.���� "

        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡҩƷĿ¼", strInfo)
        mobjLog.Add "��ȡҩƷĿ¼���", 1
    
    Case Val("4-ҩƷ������λ")
        strSQL = _
            "Select a.�ⷿid, a.ҩƷid, a.����, a.�ϴ�����, a.�ϴ���������, a.Ч��, a.ʵ������, b.���� ҩƷ����, b.���㵥λ," & vbNewLine & _
            "    c.ҩ�ⵥλ, c.סԺ��λ, c.���ﵥλ, c.ҩ���װ, c.סԺ��װ, c.�����װ, d.�ⷿ��λ, e.���� �����̱���," & vbNewLine & _
            "    e.���� ������, f.���� ��Ӧ�̱���, f.���� ��Ӧ�� " & vbNewLine & _
            "From ҩƷ��� A, �շ���ĿĿ¼ B, ҩƷ��� C, ҩƷ�����޶� D, ҩƷ������ E, ��Ӧ�� F " & _
            IIf(strInfo = "", "", " , Table(f_Num2list([1], ';')) G") & vbNewLine & _
            "Where a.ҩƷid = b.Id And a.ҩƷid = c.ҩƷid And a.ҩƷid = d.ҩƷid(+) And a.�ⷿid = d.�ⷿid(+) And " & vbNewLine & _
            "    a.�ϴβ��� = e.����(+) And a.�ϴι�Ӧ��id = f.Id(+) " & _
            IIf(strInfo = "", "", " And a.�ⷿid = g.Column_Value ")
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡҩƷ�����Ϣ", strInfo)
        mobjLog.Add "��ȡҩƷ�����Ϣ���", 1
    
    Case Val("5-��ҩ����")
        strSQL = _
            "Select a.����, a.����, a.�ϰ�� �ϰ�, a.ҩ��id " & vbNewLine & _
            "From ��ҩ���� A " & _
            IIf(strInfo = "", "", ", Table(f_Num2list([1], ';')) B") & vbNewLine & _
            "Where 1 = 1 " & _
            IIf(strInfo = "", "", " And a.ҩ��id = b.Column_Value ")
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡҩ����ҩ����", strInfo)
        mobjLog.Add "��ȡҩ����ҩ�������", 1
    
    Case Val("21-��ҩ")
        If intIO = 1 Then
            strTmp = "�����䣨�ڣ�ҩ"
            strSQL = _
                "Select /*+cardinality(z,10)*/ " & vbNewLine & _
                "    a.No ������, a.����id, a.�Ǽ�ʱ�� ����ʱ��, a.�Ƿ���, b.ID �շ�ID, b.����, b.���, b.�ⷿid, b.ҩƷid, Y1.���� ��ҩҩ��," & vbNewLine & _
                "    c.��������, a.����, a.�Ա�, a.����, c.���￨��, c.ҽ�Ƹ��ʽ ҽ�����, c.�ѱ�, d.���ȼ�, a.����ʱ�� �ɷ�ʱ��, E1.��������id," & vbNewLine & _
                "    Y2.���� ��������, E1.����ҽ��, E1.Id ҽ��id, f.����, f.���, f.���� ͨ����, g.���� ��Ʒ��, f.���� ������," & vbNewLine & _
                "    h.���ﵥλ, Cast(b.ʵ������ / h.�����װ As Number(16, 5)) ���ﵥλ����, b.�÷�, b.���� ��������, i.���㵥λ ������λ," & vbNewLine & _
                "    E1.ִ��Ƶ��, j.ҩƷ����, j.��������, E1.ҽ������, l.�������, c.����, " & vbNewLine & _
                "    Decode(Nvl(m.��־, 0), 11, 1, 12, 1, 0, 1, 0) �Ƿ���, " & vbNewLine & _
                "    Zl_Fun_Drug_Machine(b.�ⷿid, j.ҩƷ����, b.ID) �ӿڱ�� " & vbNewLine & _
                "From ������ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C, ���� D, ����ҽ����¼ E1, �շ���ĿĿ¼ F," & vbNewLine & _
                "    �շ���Ŀ���� G, ҩƷ��� H, ������ĿĿ¼ I, ҩƷ���� J, �������ҽ�� K, ������ϼ�¼ L, " & vbNewLine & _
                "    ҩƷ�շ������־ M, ���ű� Y1, ���ű� Y2, " & vbNewLine & _
                "    Table(f_Str2list2([1], ';', ',')) Z " & vbNewLine & _
                "Where a.Id = b.����id And a.����id = c.����id(+) And c.���� = d.����(+) " & vbNewLine & _
                "    And a.ҽ����� = E1.Id(+) And b.�ⷿid = Y1.Id And E1.��������id = Y2.Id(+) And a.�շ�ϸĿid = f.Id(+) " & vbNewLine & _
                "    And f.Id = g.�շ�ϸĿid(+) And a.�շ�ϸĿid = h.ҩƷid(+) And E1.���id = k.ҽ��id(+) And h.ҩ��id = i.Id(+) " & vbNewLine & _
                "    And h.ҩ��id = j.ҩ��id(+) And k.���id = l.Id(+) " & vbNewLine & _
                "    And b.no = m.������(+) And b.���� = m.����(+) And b.�ⷿid = m.�ⷿid(+) And m.ҵ�����(+) = 1 " & vbNewLine & _
                "    And b.���� = z.C1 And b.No = z.C2 " & vbNewLine & _
                "    And (Mod(b.��¼״̬, 3) = 0 Or b.��¼״̬ = 1) And a.�շ���� In ('5', '6', '7') " & vbNewLine & _
                "    And E1.���id(+) Is Not Null And g.����(+) = 3 And g.����(+) = 1 " & vbNewLine & _
                "Order By �ӿڱ��, a.����id, a.NO, b.����, b.�ⷿid "
            
        Else
            strTmp = "סԺ�䣨����ҩ"
            strSQL = _
                "Select /*+cardinality(z,10)*/ Distinct Zl_Fun_Drug_Machine(c.�ⷿid, g.ҩƷ����, c.Id) �ӿڱ��, c.�ⷿid, I1.���� As ��ҩҩ��," & vbNewLine & _
                "    Decode(Nvl(o.��־, 0), 11, 1, 12, 1, 0, 1, 0) �Ƿ���, b.��ʶ�� סԺ��, " & vbNewLine & _
                "    c.���ܷ�ҩ�� As ��ҩ��, Sysdate As ����ʱ��, c.�Է�����id As ��ҩ����id, I2.���� As ��ҩ����, j.����id, j.���� As ��������," & vbNewLine & _
                "    j.�Ա� As �����Ա�, j.���� As ��������, j.���￨��, b.���˲���id As ����id, I3.���� As ����, j.��ǰ���� As ����," & vbNewLine & _
                "    k.���ȼ�, a.����ҽ��, a.����ʱ��, a.Id As ҽ��id, g.ҩƷ����, c.ҩƷid, e.���� As ҩƷ����, f.���� As ��Ʒ����, c.�÷�," & vbNewLine & _
                "    c.����, h.���㵥λ As ������λ, d.סԺ��λ, a.ҽ����Ч, m.Ҫ��ʱ�� As ִ��ʱ��, a.ҽ������, g.��������, a.������־," & vbNewLine & _
                "    Decode(m.Ҫ��ʱ��, Null, Round(c.ʵ������ / d.סԺ��װ, 2), Round((a.�������� / d.����ϵ��) / d.סԺ��װ, 2)) As סԺ��λ����," & vbNewLine & _
                "    d.����ϵ��, a.У��ʱ�� ��ʿУ��ʱ��, c.������ �շ�������, c.����� �շ������, c.ID �շ�ID " & vbNewLine & _
                "From ����ҽ����¼ A, סԺ���ü�¼ B, ҩƷ�շ���¼ C, ҩƷ��� D, �շ���ĿĿ¼ E, �շ���Ŀ���� F, ҩƷ���� G, ������ĿĿ¼ H," & vbNewLine & _
                "    ���ű� I1, ���ű� I2, ���ű� I3, ������Ϣ J, ���� K, ����ҽ������ L, ҽ��ִ��ʱ�� M, ����ҽ����¼ N," & vbNewLine & _
                "    ҩƷ�շ�סԺ��־ O, Table(f_Num2list([1], ';')) Z" & vbNewLine & _
                "Where a.Id = b.ҽ����� And b.Id = c.����id And c.ҩƷid = d.ҩƷid And c.ҩƷid = e.Id And e.Id = f.�շ�ϸĿid(+) " & vbNewLine & _
                "    And d.ҩ��id = g.ҩ��id And g.ҩ��id = h.Id And c.�ⷿid = I1.Id And c.�Է�����id = I2.Id And b.���˲���id = I3.Id " & vbNewLine & _
                "    And b.����id = j.����id And j.���� = k.����(+) And a.���id = n.Id And n.Id = m.ҽ��id " & vbNewLine & _
                "    And b.ҽ����� = l.ҽ��id(+) And b.No = l.No And b.ִ�в���id = l.ִ�в���id " & vbNewLine & _
                "    And l.���ͺ� = m.���ͺ�(+) And c.Id = o.�շ�id(+) And f.����(+) = 3 And f.����(+) = 1 " & vbNewLine & _
                "    And c.Id = z.Column_Value " & vbNewLine & _
                "Order By c.�ⷿid, c.�Է�����id, j.����id"

        End If
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡ" & strTmp & "����", strInfo)
        mobjLog.Add "��ȡ" & strTmp & "�������", 1
        
    Case Val("22-��ʼ��ҩ"), Val("23-��ɷ�ҩ")
        If intIO = 1 Then
            If intType = 22 Then
                strTmp = "���￪ʼ��ҩ"
            Else
                strTmp = "������ɷ�ҩ"
            End If
            
            strSQL = _
                "Select Distinct a.����, a.No ������, b.�Ǽ�ʱ�� ����ʱ��, a.�ⷿid, a.ҩƷid, b.����id, d.ҩƷ����," & _
                "    Decode(Nvl(e.��־, 0), 11, 1, 12, 1, 0, 1, 0) �Ƿ���, f.���� ���ڱ���, " & _
                "    Zl_Fun_Drug_Machine(a.�ⷿid, d.ҩƷ����) �ӿڱ�� " & vbNewLine & _
                "From ҩƷ�շ���¼ A, ������ü�¼ B, ҩƷ��� C, ҩƷ���� D, ҩƷ�շ������־ E, ��ҩ���� F " & vbNewLine & _
                "   , Table(f_Str2list2([1], ';', ',')) Z " & vbNewLine & _
                "Where a.����id = b.Id And a.ҩƷid = c.ҩƷid And c.ҩ��id = d.ҩ��id(+) " & vbNewLine & _
                "    And a.��ҩ���� = f.����(+) And a.�ⷿid = f.ҩ��id(+) And a.���� = z.C1 " & vbNewLine & _
                "    And a.NO = e.������(+) And a.���� = e.����(+) And a.�ⷿid = e.�ⷿid(+) And e.ҵ�����(+) = [2] " & vbNewLine & _
                "    And a.No = z.C2 And (Mod(a.��¼״̬, 3) = 0 Or a.��¼״̬ = 1) " & vbNewLine & _
                "Order By �ӿڱ��, a.NO, a.����, a.�ⷿid "

            mobjLog.Add strSQL, 1, 1
            Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡ" & strTmp & "����", strInfo, intType - 20)
            mobjLog.Add "��ȡ" & strTmp & "�������", 1
        Else
            strTmp = "סԺ��ʼ��ҩ��֧��"
            
            Set rsTmp = Nothing
            mobjLog.Add strTmp, 1
        End If
    
    Case Val("24-������ҩ")
        If intIO = 1 Then
            strTmp = "���ﴦ����ҩ"
            strSQL = _
                "Select /*+cardinality(h,10)*/ " & vbNewLine & _
                "    Decode(Mod(B.��¼����, 10), 1, 8, 9) ����, b.No ������, b.ִ�в���id �ⷿid, b.�Ǽ�ʱ�� ����ʱ��, " & vbNewLine & _
                "    b.ҽ����� ҽ��id, e.���� ҩƷ����, " & vbNewLine & _
                "    Cast((b1.���� * b1.���� - Decode(b1.����, 0, 0, 1) * h.C2) / c.�����װ As Number(16, 5)) ���ﵥλ����,  " & vbNewLine & _
                "    Decode(Nvl(f.��־, 0), 11, 1, 12, 1, 0, 1, 0) �Ƿ���, " & vbNewLine & _
                "    Zl_Fun_Drug_Machine(b.ִ�в���id, d.ҩƷ����, Null) �ӿڱ�� " & vbNewLine & _
                "From ������ü�¼ B, ������ü�¼ B1, ҩƷ��� C, ҩƷ���� D, �շ���ĿĿ¼ E, ҩƷ�շ������־ F " & vbNewLine & _
                "    , Table(f_Str2list2([1], ';', ',')) H " & vbNewLine & _
                "Where b.NO = b1.NO(+) And b.��¼���� = b1.��¼����(+) And b.ҽ����� = b1.ҽ�����(+) " & vbNewLine & _
                "    And b.�շ�ϸĿid = c.ҩƷid And c.ҩ��id = d.ҩ��id(+) And b.�շ�ϸĿid = e.Id And b.No = f.������(+) " & vbNewLine & _
                "    And Decode(Mod(B.��¼����, 10), 1, 8, 9) = f.���� And b.ִ�в���id = f.�ⷿid(+) And b.Id = h.C1 " & vbNewLine & _
                "    And b1.��¼״̬(+) = 3 And f.ҵ�����(+) = 4 " & vbNewLine & _
                "Order By �ӿڱ��, ������, ����, �ⷿid "
                
            mobjLog.Add strSQL, 1, 1
            Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡ" & strTmp & "����", strInfo)
            mobjLog.Add "��ȡ" & strTmp & "�������", 1
        Else
            strTmp = "סԺ��ҩ��֧��"

            Set rsTmp = Nothing
            mobjLog.Add strTmp, 1
        End If
        
    Case Val("25-����������ҩ")
        If intIO = 1 Then
            strTmp = "����������ҩ"
            strSQL = _
                "Select /*+cardinality(h,10)*/ Distinct " & vbNewLine & _
                "    Decode(Mod(B.��¼����, 10), 1, 8, 9) ����, b.No ������, b.ִ�в���id �ⷿid, b.�Ǽ�ʱ�� ����ʱ��, " & vbNewLine & _
                "    Decode(Nvl(f.��־, 0), 11, 1, 12, 1, 0, 1, 0) �Ƿ���, " & vbNewLine & _
                "    Zl_Fun_Drug_Machine(b.ִ�в���id, d.ҩƷ����, Null) �ӿڱ�� " & vbNewLine & _
                "From ������ü�¼ B, ҩƷ��� C, ҩƷ���� D, ҩƷ�շ������־ F, " & vbNewLine & _
                "    Table(f_Str2list2([1], ';', ',')) H " & vbNewLine & _
                "Where b.�շ�ϸĿid = c.ҩƷid And c.ҩ��id = d.ҩ��id(+) And " & vbNewLine & _
                "    b.No = f.������(+) And Decode(Mod(B.��¼����, 10), 1, 8, 9) = f.����(+) And " & vbNewLine & _
                "    b.ִ�в���id = f.�ⷿid(+) And b.No = h.C1 And b.ִ�в���id = h.C2 " & vbNewLine & _
                "Order By �ӿڱ��, ������, ����, �ⷿid "
            
            mobjLog.Add strSQL, 1, 1
            Set rsTmp = objDB.OpenSQLRecord(strSQL, "��ȡ" & strTmp & "����", strInfo)
            mobjLog.Add "��ȡ" & strTmp & "�������", 1
        Else
            strTmp = "סԺ��ҩ��֧��"

            Set rsTmp = Nothing
            mobjLog.Add strTmp, 1
        End If
    
    End Select
    
    mobjLog.Save
    If Not rsTmp Is Nothing Then
        Set GetData = rsTmp.Clone
    End If
    
    blnFinish = True
    
    Exit Function
    
hErr:
    blnFinish = False
    mobjLog.Add "��ȡ��׼����ʧ��", 1
    mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjLog.Save
End Function

Private Sub Class_Initialize()
    Set mobjLog = New clsLog
End Sub

Private Sub Class_Terminate()
    Set mobjLog = Nothing
End Sub

Public Function FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, ByRef rsFinish As ADODB.Recordset) As Boolean
'���ܣ����˳�ҩƷ�Զ����ӿڵ�����
'������
'  intType��ҵ�����
'           1-������Ϣ��
'           2-��Ա��Ϣ��
'           3-ҩƷĿ¼��
'           4-ҩƷ������λ��
'           5-��ҩ���ڣ�
'           21-��ҩ[�����סԺ������ϸ�ϴ�]��
'           22-��ʼ��ҩ��
'           23-��ɷ�ҩ��
'           24-����(����/����)��ҩ��
'  rsData����׼��ҵ���¼������
'  rsFinish���ѹ�����ɵļ�¼������
'���أ�True��ɣ�Falseʧ��
'˵��������������ڶ�Ӧ�Ľӿ���ģ���д

    Dim strTemp As String

    On Error GoTo hErr
    
    mobjLog.Add "��ʼ��������", 1
    
    Set rsFinish = mdlDrugMachine.CopyStructure(rsData.Fields)
    mobjLog.Add "��ɼ�¼���ṹ����", 1
    
    If intType > 20 Then
        rsData.Filter = "�ӿڱ��='" & strCode & "'"
        mobjLog.Add rsData.Filter, 1, 1
    End If
    Do While rsData.EOF = False
        strTemp = CopyRecord(rsData.Fields, rsFinish)
    
        rsData.MoveNext
    Loop
    
    FilterData = True
    
    mobjLog.Add "��ɹ�������", 1
    mobjLog.Save

    Exit Function
    
hErr:
    mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjLog.Save
End Function

Public Function SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, ByRef strMessage As String, _
    Optional ByRef strWinInfo As String) As Boolean
'���ܣ���ҩƷ�Զ����ӿڴ�������
'������
'  intAppType��Ӧ�����1-ZLHIS��2-�ӿڹ������ߣ���ʱ��������ӿڴ������ݣ���3-�ƶ���֧��������4-��������ƽ̨
'  intType��ҵ�����ͣ���� GetData �����Ĳ���
'  intIO��1-���2-סԺ
'  strCode���ӿڱ��
'  rsData���ѹ�����ɵļ�¼������
'  strMessage��������Ϣ
'  strWinInfo���ƶ���֧�����������
'���أ�True��ɣ�Falseʧ��
'˵��������������ڶ�Ӧ�Ľӿ���ģ���д

    Dim strMsg As String

    On Error GoTo hErr
    
    mstrInfCode = strCode
    If intAppType = Val("3-�ƶ�") And intType = Val("21-�շ�") And intIO = Val("1-����") Then
        '�ƶ���֧�������ɷѣ��ӿڷ��ش�����Ϣ���ƶ�ҵ��
        mobjLog.Add "�ӿڷ��ش�����Ϣ���ƶ�ҵ��", 1
        strWinInfo = MobilePayment(mobjLog, mstrNOs, strMsg)
        strMessage = IIf(strMessage <> "", strMessage & vbNewLine, "") & strMsg
    End If
    
    mobjLog.Save
    SendData = True
    Exit Function
    
hErr:
    strWinInfo = ""
    mobjLog.Add Err.Description, 1
    mobjLog.Save
End Function

Private Function MobilePayment(ByVal objLog As clsLog, ByVal strNO As String, _
    Optional ByRef strMessage As String) As String
    
'���ܣ���ȡ���뵥�ݺŵĲ���ID����ҩҩ������ҩ���ڡ�ҩƷ������Ϣ����Ҫ���ƶ�ҵ����ʹ��
'������
'  objLog����־�����
'  strNO��������Ϣ����ʽ������ò�˵��
'  strMessage��ʵ�Σ����쳣��Ϣ
'���أ�����ID����ҩҩ������ҩ���ڡ�ҩƷ������Ϣ��XML�ṹ�ַ���
'XML��ʽ��
'<OUTPUT>
'  <BRID>����ID</BRID>
'  <ITEM>
'    <YFMC>ҩ������</YFMC>
'    <YFCK>��ҩ����</YFCK>
'    <YFMX>
'      <ITEM>
'        <MC>ҩƷ����1</MC>
'      </ITEM>
'      <ITEM>
'        <MC>ҩƷ����2</MC>
'      </ITEM>
'      <ITEM>
'        <MC>ҩƷ����...</MC>
'      </ITEM>
'    </YFMX>
'  </ITEM>
'  <ITEM>
'    ...
'  </ITEM>
'</OUTPUT>

    Const STR_OUT As String = "OUTPUT"
    Const STR_BRID As String = "BRID"
    Const STR_ITEM As String = "ITEM"
    Const STR_YFMC As String = "YFMC"
    Const STR_YFCK As String = "YFCK"
    Const STR_YFMX As String = "YFMX"
    Const STR_MC As String = "MC"

    Dim objXML As New clsXML
    Dim rsSQL As ADODB.Recordset
    Dim strSQL As String, strWin As String
    Dim lngStoreID As Long
    
    On Error GoTo hErr
    
    strSQL = "Select Distinct b.����id, a.�ⷿid, d.���� As ҩ������, a.��ҩ����, c.���� As ҩƷ���� " & vbCr & _
             "From ҩƷ�շ���¼ A, ������ü�¼ B, �շ���ĿĿ¼ C, ���ű� D, Table(f_Str2list2([1], ';', ',')) E " & vbCr & _
             "Where a.����id = b.Id And a.ҩƷid = c.Id And a.�ⷿid = d.Id And a.���� = e.C1 And a.No = e.C2 " & vbCr & _
             "Order By Nvl(b.����id, 0) Desc, a.�ⷿid, c.���� "
    Set rsSQL = mobjComLib.OpenSQLRecord(strSQL, "��ȡҩ����ҩ��Ϣ", strNO)
    With rsSQL
        If .RecordCount > 0 Then
            objXML.AppendNode STR_OUT, False
            objXML.AppendData STR_BRID, mdlDrugMachine.NVL(!����id)
            
            lngStoreID = -9999
            Do While .EOF = False
                strWin = mdlDrugMachine.NVL(!��ҩ����)
                If lngStoreID <> Val(mdlDrugMachine.NVL(!�ⷿid)) Then
                    objXML.AppendNode STR_ITEM, False
                    objXML.AppendData STR_YFMC, mdlDrugMachine.NVL(!ҩ������)
                    objXML.AppendData STR_YFCK, strWin
                    
                    objXML.AppendNode STR_YFMX, False
                    GoSub makAdd
                Else
makAdd:
                    objXML.AppendNode STR_ITEM, False
                    objXML.AppendData STR_MC, mdlDrugMachine.NVL(!ҩƷ����)
                    objXML.AppendNode STR_ITEM, True
                End If
                
                lngStoreID = Val(mdlDrugMachine.NVL(!�ⷿid))
                .MoveNext
                
                If .EOF Then
                    GoTo makEnd
                Else
                    If lngStoreID <> Val(mdlDrugMachine.NVL(!�ⷿid)) Then
makEnd:
                        objXML.AppendNode STR_YFMX, True
                        objXML.AppendNode STR_ITEM, True
                    End If
                End If
            Loop
            objXML.AppendNode STR_OUT, True
        End If
        .Close
    End With
    
    MobilePayment = objXML.XmlText
    
    objLog.Add MobilePayment, 1
    objLog.Add "���ء�����ID����ҩҩ������ҩ���ڡ�ҩƷ���ơ���Ϣ�ɹ�", 1
    objLog.Save
    
    Set objXML = Nothing
    Exit Function
    
hErr:
    strMessage = "���ء�����ID����ҩҩ������ҩ���ڡ�ҩƷ���ơ���Ϣʧ��"
    objLog.Add strNO, 2, 1
    objLog.Add strSQL, 2, 1
    objLog.Add strMessage, 1
    objLog.Save
    MobilePayment = ""
End Function

