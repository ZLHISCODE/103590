VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Group"
Attribute VB_Ext_KEY = "Member0" ,"Group"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarѡ������ As Integer
Public p����ִ��ʱ�� As Date
Public pִ��ժҪ As String

Public p��ˮ�� As String
Public p��ҩ�� As String
Public p��ϵ�� As Integer
Public p�ӵ��� As String
Public p��ʱ As Integer
Public p���� As Integer

'�ֲ����������漯��
Private mCol As Collection

Public Function GetGroups(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long, ByVal intִ����� As Integer, _
        ByVal datAdviceB As Date, ByVal datAdviceE As Date, ByVal str�Һŵ� As String, _
        ByVal strKey As String, ByVal bytPatiFrom As Byte) As Boolean
'����: ȡ��ִ�е���Ŀ�������
    
    Dim strSQL As String
    Dim rsGroups As New ADODB.Recordset
    Dim rsBill As New ADODB.Recordset
    Dim objBIlls As Bills, cur��� As Currency
   
    Dim lng��ִ������ As Long, dat�ϴ�ִ��ʱ�� As Date
    Dim intĬ�ϵ��� As Integer, lngʱ�� As Long
    Dim strҽ������  As String, strTmp As String
    Dim int�Ʒ�״̬  As Integer
    
    On Error GoTo errHandle
    intĬ�ϵ��� = Val(zlDatabase.GetPara("Ĭ�ϵ���", glngSys, 1264))
    If intĬ�ϵ��� < 10 Or intĬ�ϵ��� > 100 Then intĬ�ϵ��� = 40
    'ȡ���е�ִ�е�,�ӵ�groups������
    
    If bytPatiFrom = 1 Then
        '��������
        strSQL = "Select A.ҽ��id, A.���ͺ�, A.��������,c.��������,Decode(C.ִ�з���, 1, '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����') As ִ�з���," & vbNewLine & _
                " B.ҽ������ As �÷�, B.Ƥ�Խ��, B.ִ��Ƶ��, B.Ƶ�ʴ���, A.ִ��״̬, A.�Ʒ�״̬,B.�շ�ϸĿid, b.ҽ������, D.��¼״̬, B.ִ������, " & vbNewLine & _
                " Sum(Decode(Sign(D.��¼״̬), 0, 0, Decode(d.����״̬,1,0,nvl(D.ʵ�ս��,0)))) ʵ�ս��" & vbNewLine & _
                " From ������ü�¼ D, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ������ A" & vbNewLine & _
                " Where  A.NO=D.No(+) And A.ҽ��id = D.ҽ�����(+) And B.������Ŀid = C.ID And A.ҽ��id = B.ID And C.���='E' And A.ִ�в���id = [2] " & vbNewLine & _
                " And nvl(C.ִ�з���,0) = [3] And D.��¼״̬(+)<>2 And A.����ʱ�� Between [4] And [5] And B.������Դ = 2 " & vbNewLine & _
                " And B.����id = [1] and b.��ҳid=[6] " & vbNewLine & _
                " Group By A.ҽ��id, A.���ͺ�, A.��������,c.��������,Decode(C.ִ�з���, 1, '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����'), B.ҽ������," & vbNewLine & _
                " B.Ƥ�Խ��, B.ִ��Ƶ��, B.Ƶ�ʴ���, A.ִ��״̬, A.�Ʒ�״̬,B.�շ�ϸĿid, b.ҽ������, D.��¼״̬, B.ִ������" & _
                " Order by ҽ��ID"
        Set rsGroups = zlDatabase.OpenSQLRecord(strSQL, "Groups.getGroups", lng����ID, lngִ�в���ID, intִ�����, datAdviceB, datAdviceE, Val(Split(strKey, "_")(1)))
    Else
        strSQL = "Select A.ҽ��id, A.���ͺ�, A.��������,c.��������,Decode(C.ִ�з���, 1, '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����') As ִ�з���," & vbNewLine & _
                " B.ҽ������ As �÷�, B.Ƥ�Խ��, B.ִ��Ƶ��, B.Ƶ�ʴ���, A.ִ��״̬, A.�Ʒ�״̬,B.�շ�ϸĿid, b.ҽ������, D.��¼״̬, B.ִ������, " & vbNewLine & _
                " Sum(Decode(Sign(D.��¼״̬), 0, 0, Decode(d.����״̬,1,0,nvl(D.ʵ�ս��,0)))) ʵ�ս��" & vbNewLine & _
                " From ������ü�¼ D, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ������ A" & vbNewLine & _
                " Where  A.NO=D.No(+) And A.ҽ��id = D.ҽ�����(+) And B.������Ŀid = C.ID And A.ҽ��id = B.ID And C.���='E' And " & _
                " A.ִ�в���id = [2] And B.�Һŵ�=[6] And nvl(C.ִ�з���,0) = [3] And " & vbNewLine & _
                " D.��¼״̬(+)<>2 And A.����ʱ�� Between [4] And [5] And B.������Դ = 1 And B.����id = [1]" & vbNewLine & _
                " Group By A.ҽ��id, A.���ͺ�, A.��������,c.��������,Decode(C.ִ�з���, 1, '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����'), B.ҽ������," & vbNewLine & _
                " B.Ƥ�Խ��, B.ִ��Ƶ��, B.Ƶ�ʴ���, A.ִ��״̬, A.�Ʒ�״̬,B.�շ�ϸĿid, b.ҽ������, D.��¼״̬, B.ִ������" & _
                " Order by ҽ��ID"
        Set rsGroups = zlDatabase.OpenSQLRecord(strSQL, "Groups.getGroups", lng����ID, lngִ�в���ID, intִ�����, datAdviceB, datAdviceE, str�Һŵ�)
    End If
            
    cur��� = 0
    Do Until rsGroups.EOF
        lng��ִ������ = 0
        dat�ϴ�ִ��ʱ�� = CDate(0)
        '�ڼӵ�group����֮ǰ,��ҪȡҩƷ���ݺ�ִ�����,�浽bills��execrecords��
        
        '����bills����
        Select Case intִ�����
        Case 0
            '����  ���ָ�ҩ;���ͷǸ�ҩ;��
            If Val(rsGroups.Fields("��������") & "") = 2 Then
                '���Ǹ�ҩ;��ʱ�� ��Һ,ע�� ��ѯ��SQL��ͬ
                strSQL = "Select C.���, B.ID, Rownum ˳��, B.ִ��Ƶ��, B.ҽ������, B.��������, C.���㵥λ," & vbNewLine & _
                        "       Decode(Sign(Instr(',5,6,', ',' || C.��� || ',')), 1, ( B.��������/G.����ϵ��) * G.����, 0) As ��������," & vbNewLine & _
                        "       B.ִ��Ƶ��, B.ҽ������,D.��¼״̬,Decode(D.����״̬,1,0,D.ʵ�ս��) as ʵ�ս��,E.�Ʒ�״̬,B.ִ������" & vbNewLine & _
                        "From ����ҽ������ E,������ü�¼ D,ҩƷ��� G, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ������ A" & vbNewLine & _
                        "Where B.id=E.ҽ��ID and E.NO=D.NO(+) And  E.ҽ��id = D.ҽ�����(+) And D.��¼״̬(+)<>2 And B.�շ�ϸĿid = G.ҩƷid(+) And B.������Ŀid = C.ID" & _
                        " And A.ҽ��id = B.���id And A.ҽ��id=[1] And A.���ͺ�=[2] And A.ִ�в���id=[3] and E.���ͺ�=[2]" & vbNewLine & _
                        "Order By B.���"
            Else
                strSQL = "Select m.���,m.ID,m.ִ��Ƶ��,m.ҽ������,m.��������,m.���㵥λ,m.ҽ������,m.��¼״̬,m.ʵ�ս��,m.�Ʒ�״̬,m.ִ������,Rownum As ˳��,0 As �������� " & vbNewLine & _
                        "From (Select C.���, B.Id, B.ִ��Ƶ��, B.ҽ������ || '��' || B1.ҽ������ ҽ������, B1.��������, C.���㵥λ, B.ҽ������, D.��¼״̬, " & _
                        "             Sum(Decode(D.����״̬,1,0,D.ʵ�ս��)) As ʵ�ս��, A.�Ʒ�״̬, B.ִ������ " & vbNewLine & _
                        "       From ������ü�¼ D, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ����¼ B1, ����ҽ������ A " & vbNewLine & _
                        "       Where A.No = D.No(+) And A.ҽ��id = D.ҽ�����(+) And D.��¼״̬(+) <> 2 And B1.������Ŀid = C.Id(+) And " & vbNewLine & _
                        "             A.ҽ��id = B.Id And A.ҽ��id = B1.���ID(+) And A.ҽ��id=[1] And A.���ͺ�=[2] And A.ִ�в���id=[3]" & vbNewLine & _
                        "       Group By C.���, B.Id, B.ִ��Ƶ��, B.ҽ������, B1.��������, C.���㵥λ, B.ִ��Ƶ��, B.ҽ������, D.��¼״̬, A.�Ʒ�״̬, B.ִ������, b1.ҽ������" & vbNewLine & _
                        "      ) M " & vbNewLine & _
                        "Order By ˳�� "
            End If            
        Case 3
            'Ƥ��
            strSQL = "Select m.���,m.ID,m.ִ��Ƶ��,m.ҽ������,m.��������,m.���㵥λ,m.ҽ������,m.��¼״̬,m.ʵ�ս��,m.�Ʒ�״̬,m.ִ������, Rownum As ˳��, 0 As ��������" & vbNewLine & _
                    "From (Select C.���, B.Id, B.ִ��Ƶ��, B.ҽ������, B.��������, C.���㵥λ, B.ҽ������, D.��¼״̬, " & _
                    "             Sum(Decode(D.����״̬,1,0,D.ʵ�ս��)) As ʵ�ս��, A.�Ʒ�״̬, B.ִ������" & vbNewLine & _
                    "       From ������ü�¼ D, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ������ A" & vbNewLine & _
                    "       Where A.No = D.No(+) And A.ҽ��id = D.ҽ�����(+) And D.��¼״̬(+) <> 2 And B.������Ŀid = C.Id And A.ҽ��id = B.Id" & vbNewLine & _
                    "            And A.ҽ��id=[1] And A.���ͺ�=[2] And A.ִ�в���id=[3] " & vbNewLine & _
                    "       Group By C.���, B.Id, B.ִ��Ƶ��, B.ҽ������, B.��������, C.���㵥λ, B.ִ��Ƶ��, B.ҽ������, D.��¼״̬, A.�Ʒ�״̬, B.ִ������) M" & vbNewLine & _
                    "Order By ˳��"
        
        Case Else
            '��Һ,ע��
            strSQL = "Select C.���, B.ID, Rownum ˳��, B.ִ��Ƶ��, B.ҽ������, B.��������, C.���㵥λ," & vbNewLine & _
                    "       Decode(Sign(Instr(',5,6,', ',' || C.��� || ',')), 1, ( B.��������/G.����ϵ��) * G.����, 0) As ��������," & vbNewLine & _
                    "       B.ִ��Ƶ��, B.ҽ������,D.��¼״̬,Decode(D.����״̬,1,0,D.ʵ�ս��) as ʵ�ս��,E.�Ʒ�״̬,B.ִ������" & vbNewLine & _
                    "From ����ҽ������ E,������ü�¼ D,ҩƷ��� G, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ������ A" & vbNewLine & _
                    "Where B.id=E.ҽ��ID and E.NO=D.NO(+) And  E.ҽ��id = D.ҽ�����(+) And D.��¼״̬(+)<>2 And B.�շ�ϸĿid = G.ҩƷid(+) And B.������Ŀid = C.ID" & _
                    " And A.ҽ��id = B.���id And A.ҽ��id=[1] And A.���ͺ�=[2] And A.ִ�в���id=[3] and E.���ͺ�=[2]" & vbNewLine & _
                    "Order By B.���"

        End Select
        If intִ����� = 1 Then
            'ҽ������ ���� X��/���� �ģ��ݴ˸���Ĭ�ϵ���
            strҽ������ = "" & rsGroups.Fields("ҽ������")
            If InStr(strҽ������, "��/����") > 0 Then
                If Val(strҽ������) > 10 And Val(strҽ������) < 100 Then
                    intĬ�ϵ��� = Val(strҽ������)
                End If
            End If
        End If
        
        Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "Groups.getGroups", Val(rsGroups.Fields("ҽ��ID")), Val(rsGroups.Fields("���ͺ�")), lngִ�в���ID)
        With rsBill
            Set objBIlls = New Bills
            Do Until .EOF
                
                lngʱ�� = CacleTransTime(Val(zlCommFun.NVL(.Fields("��������"), 0)), p��ϵ��, intĬ�ϵ���)
                int�Ʒ�״̬ = Val("" & .Fields("�Ʒ�״̬"))
                If Val("" & .Fields("ִ������")) = 5 And int�Ʒ�״̬ <> -1 Then
                    int�Ʒ�״̬ = -1 'Ժ���ҩ�ģ���ͬΪ����Ʒ�
                ElseIf int�Ʒ�״̬ > 0 Then
                    '�ѼƷѵ�,�ж��Ƿ������
                    If .Fields("��¼״̬") = 1 And Val("" & .Fields("ʵ�ս��")) = 0 Then
                        int�Ʒ�״̬ = -2
                    ElseIf .Fields("��¼״̬") = 3 Then
                        int�Ʒ�״̬ = -3    '���˷�
                    End If
                End If
                Call objBIlls.Add(lng����ID, Val(.Fields("ID")), zlCommFun.NVL(.Fields("ҽ������")), zlCommFun.NVL(.Fields("��������"), 0), _
                            zlCommFun.NVL(.Fields("���㵥λ")), Val(zlCommFun.NVL(.Fields("��������"), 0)), _
                            lngʱ��, zlCommFun.NVL(.Fields("ҽ������")), Val(.Fields("˳��")), _
                            IIf(Val("" & .Fields("��¼״̬")) = 0, 0, Val("" & .Fields("ʵ�ս��"))), _
                            int�Ʒ�״̬, Val(rsGroups.Fields("ҽ��ID")) & "_" & .Fields("ID"))
                
                .MoveNext
            Loop
        End With
        

        strSQL = "Select  max(ִ��ʱ��) as �ϴ�ִ��ʱ��,sum(nvl(B.��������,0)) as ִ������" & vbNewLine & _
                "From ����ҽ��ִ�� B" & vbNewLine & _
                "Where B.ҽ��id = [1] And B.���ͺ� = [2] "
        Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "Groups.getGroups", Val(rsGroups.Fields("ҽ��ID")), Val(rsGroups.Fields("���ͺ�")))
        With rsBill

        Do Until .EOF
            lng��ִ������ = zlCommFun.NVL(.Fields("ִ������"), 0)
            dat�ϴ�ִ��ʱ�� = zlCommFun.NVL(.Fields("�ϴ�ִ��ʱ��"), CDate(0))
            .MoveNext
        Loop
        End With
        
        '����ִ�е���Ŀ ����group����
        With rsGroups
            If Val(.Fields("��������")) - lng��ִ������ > 0 Then
                '���� ִ����Ŀ(�羲����Һ) �����Ľ��
                cur��� = Val("" & .Fields("ʵ�ս��"))
                int�Ʒ�״̬ = Val("" & .Fields("�Ʒ�״̬"))
                
                If Val("" & .Fields("ִ������")) = 5 And int�Ʒ�״̬ <> -1 Then
                    int�Ʒ�״̬ = -1 'Ժ���ҩ�ģ���ͬΪ����Ʒ�
                ElseIf int�Ʒ�״̬ > 0 Then
                    '�ѼƷѵ�,�ж��Ƿ������
                    If .Fields("��¼״̬") = 1 And cur��� = 0 Then
                        int�Ʒ�״̬ = -2       '
                    ElseIf .Fields("��¼״̬") = 3 Then
                        int�Ʒ�״̬ = -3
                    End If
                End If
                If Me.Item(CStr(.Fields("ҽ��ID") & "_" & .Fields("���ͺ�"))) Is Nothing Then
                    Call Add(.Fields("ҽ��ID"), .Fields("���ͺ�"), Val(.Fields("��������")), lng��ִ������, dat�ϴ�ִ��ʱ��, _
                              .Fields("ִ�з���"), zlCommFun.NVL(.Fields("�÷�")), zlCommFun.NVL(.Fields("Ƥ�Խ��")), _
                              zlCommFun.NVL(.Fields("ִ��Ƶ��"), ""), zlCommFun.NVL(.Fields("Ƶ�ʴ���"), 0), _
                              intĬ�ϵ���, zlCommFun.NVL(.Fields("ִ��״̬"), 0), cur���, objBIlls, int�Ʒ�״̬, _
                              .Fields("ҽ��ID") & "_" & .Fields("���ͺ�"))
                              
                    Call GetOtherValue(.Fields("ҽ��ID"), .Fields("���ͺ�"))
                End If
           End If
        End With
        
        rsGroups.MoveNext
    Loop
    GetGroups = True
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Add(ByVal ִ��ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal �������� As Integer, ByVal ��ִ������ As Integer, ByVal �ϴ�ִ��ʱ�� As Date, _
        ByVal ִ�з��� As String, ByVal �÷� As String, ByVal Ƥ�Խ�� As String, ByVal ִ��Ƶ�� As String, ByVal Ƶ�ʴ��� As Integer, _
        ByVal ���� As Integer, ByVal ִ��״̬ As Integer, ByVal �շѽ�� As Currency, ByVal clsBills As Bills, ByVal �Ʒ�״̬ As Integer, _
        Optional sKey As String) As Group
    '�����¶���
    Dim objNewMember As Group, objNewBills As Bills
    Set objNewMember = New Group
    
    '���ô��뷽��������
    objNewMember.ִ��ҽ��ID = ִ��ҽ��ID
    objNewMember.���ͺ� = ���ͺ�
    objNewMember.�������� = ��������
    objNewMember.��ִ������ = ��ִ������
    objNewMember.�ϴ�ִ��ʱ�� = �ϴ�ִ��ʱ��
    objNewMember.ִ�з��� = ִ�з���
    objNewMember.�÷� = �÷�
    objNewMember.Ƥ�Խ�� = Ƥ�Խ��
    objNewMember.ִ��Ƶ�� = ִ��Ƶ��
    objNewMember.Ƶ�ʴ��� = Ƶ�ʴ���
    objNewMember.ִ��״̬ = ִ��״̬
    objNewMember.���� = ����
    objNewMember.�շѽ�� = �շѽ��
    objNewMember.�Ʒ�״̬ = �Ʒ�״̬
    If clsBills Is Nothing Then
        Set objNewBills = New Bills
    Else
        Set objNewBills = clsBills
    End If
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
        Call objNewMember.AddBills(objNewBills, 0)
    Else
        mCol.Add objNewMember, sKey
        Call objNewMember.AddBills(objNewBills, sKey)
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Let ѡ������(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ��ҽ��ID = 5
    mvarѡ������ = vData
End Property

Public Property Get ѡ������() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��ҽ��ID
    ѡ������ = mvarѡ������
End Property

Public Property Get Item(vntIndexKey As Variant) As Group
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    ' x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub

Public Sub CheckGroup(ByVal strBillKey As String, blnCheck As Boolean)
    'ѡ��������һ������ʱִ��
    Dim lng��� As Long
    If Me.ѡ������ < 0 Then Me.ѡ������ = 0
    Me.ѡ������ = Me.ѡ������ + IIf(blnCheck, 1, -1)
    If blnCheck Then
        Me.Item(strBillKey).��� = Me.ѡ������
    Else
        lng��� = Me.Item(strBillKey).���
        Me.Item(strBillKey).��� = 0
    End If
    
    If Me.ѡ������ > 0 Then
        Call Sort(lng���)
    End If
    
End Sub

Private Sub Sort(ByVal lng��� As Long)
    Dim blnҪ�� As Boolean
    Dim objGroup As Group
    Dim lng�� As Long
    For Each objGroup In Me
        If objGroup.��� > Me.ѡ������ Then
            blnҪ�� = True
        End If
    Next
    If blnҪ�� Then
        For Each objGroup In Me
            If objGroup.��� <> 0 Then
                If Me.Item(objGroup.ִ��ҽ��ID & "_" & objGroup.���ͺ�).��� > lng��� Then
                    lng�� = Me.Item(objGroup.ִ��ҽ��ID & "_" & objGroup.���ͺ�).���
                    Me.Item(objGroup.ִ��ҽ��ID & "_" & objGroup.���ͺ�).��� = IIf(objGroup.��� - 1 < 0, 0, IIf(objGroup.��� = 1, 1, objGroup.��� - 1))
                End If
            End If
        Next
        If lng�� > 0 Then
            Call Sort(lng��)
        End If
    End If
End Sub

Public Function SelectGroupThingNew(ByVal lngExcDeptID As Long, Optional ByVal blnPrint As Boolean, _
            Optional ByVal intType As Integer, Optional ByVal frmParent As Object, _
            Optional ByVal blnView As Boolean, Optional ByVal blnLabel As Boolean = False, _
            Optional ByVal blnWristband As Boolean = False) As Boolean
'����:    ��¼ִ�������������ѡ���ִ�����ݵ�����ҽ��ִ�б�
'����
'  lngExcDeptID : ִ�в���ID 47323
'  blnPrint����ӡ
'  intType����������ţ�
'  frmParent�����������
'  blnView��TrueԤ����False��Ԥ��
'  blnLabel��True��ӡ��Һƿǩ��False����ӡ
'  blnWristband��True��ӡ��Һ�����False����ӡ
'���أ�True�ɹ���Falsʧ��

    Dim objGroup As Group
    Dim strSQL As String, lng��ˮ�� As Long, rsTmp As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim strRpt As String
    Dim objBIlls As Bills, i As Integer, Y As Integer
    Dim objBIll As Bill
    Dim cur���� As Currency
    Dim arrSQL As Variant
    
    On Error GoTo errH
    
    'ȡ��ˮ��
    strSQL = "Select ����ҽ��ִ��_��ˮ��.nextval as ��ˮ�� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Group.SelectGroupThingNew")
    If rsTmp.EOF Then
        MsgBox "ȡ��ˮ��ʧ�ܣ����飡", vbInformation, gstrSysName
    Else
        lng��ˮ�� = zlCommFun.NVL(rsTmp.Fields("��ˮ��"), 1)
    End If
    
    arrSQL = Array()
    
    For Each objGroup In Me
        With objGroup
        If .��� > 0 Then
            '0-δִ��;1-��ȫִ��;2-�ܾ�ִ��;3-����ִ��(�����ֽܷ�Ϊ����ʵ�ʲ���)
            If .ִ��״̬ = 1 Then
                MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
                Exit Function
            End If
            
            '����ִ��ʱ��Ӧ�����ϴ�ִ��ʱ��
            Call GetOtherValue(.ִ��ҽ��ID, .���ͺ�)
            If .�������� <= 0 Then .�������� = 1
            If p����ִ��ʱ�� <= .�ϴ�ִ��ʱ�� Then
                MsgBox "����ִ��ʱ��Ӧ�����ϴ�ִ��ʱ�� " & Format(.�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                Exit Function
            End If
            
            If .�������� > .��ִ������ Then
                'ִ���˲�������ʱ����˱�ʾ���
                'ZL_����ҽ��ִ��_Insert(ҽ��ID,���ͺ�,Ҫ��ʱ��,��������,ִ��ժҪ,ִ����,ִ��ʱ��)
                strSQL = "ZL_����ҽ��ִ��_Insert(" & .ִ��ҽ��ID & "," & .���ͺ� & "," & _
                    "To_Date('" & Format(.Ҫ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    Val(.��������) & ",'" & DelInvalidChar(pִ��ժҪ, "'") & "',''," & _
                    "To_Date('" & Format(p����ִ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),0,0,1,Null,Null,Null," & lngExcDeptID & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL '
                strSQL = "ZL_����ҽ��ִ��_Transfusion(" & .ִ��ҽ��ID & "," & .���ͺ� & "," & _
                         "To_Date('" & Format(p����ִ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & lng��ˮ�� & ",'" & p��ҩ�� & "'," & _
                         Me.ѡ������ & "," & .��� & "," & .���� & "," & p��ϵ�� & "," & .Һ���� & ",'" & DelInvalidChar(.˵��, "'") & "','" & p�ӵ��� & "'," & p��ʱ & "," & p���� & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                
                For i = 1 To objGroup.BillsCount
                    Set objBIlls = objGroup.BillsItem(i)
                    For Y = 1 To objBIlls.Count
                        Set objBIll = objBIlls.Item(Y)
                        '����ݴ�ҩƷ������У�����дʹ�ü�¼
                                
                        strSQL = "Select B.NO, B.���, B.ҩƷid,C.��������, D.����ϵ��, Sum(Nvl(B.����* B.���ϵ��, 0)) As �Ĵ�����, max(d.����ɷ����) ����ɷ���� " & vbNewLine & _
                                "From ҩƷ��� D, ����ҽ������ C, �ݴ�ҩƷ��¼ B, ����ҽ����¼ A" & vbNewLine & _
                                "Where C.ҽ��id = B.ҽ��id And C.���ͺ� = B.���ͺ� And D.ҩƷid = B.ҩƷid And" & vbNewLine & _
                                "      B.ҩƷid = A.�շ�ϸĿid And A.ID = B.ҽ��id And A.ID = [1] And B.���ͺ� = [2]" & vbNewLine & _
                                "Group By B.NO, B.���, B.ҩƷid, C.�������� ,D.����ϵ��" & vbNewLine & _
                                "Having Sum(Nvl(B.����, 0)) >0"

                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ݴ�ҩƷ", objBIll.ҽ��ID, .���ͺ�)
                        If Not rsTmp.EOF Then
                            '����ִ�е��������ܴ��ڼĴ�����
                            If zlCommFun.NVL(rsTmp!����ɷ����, 0) = 0 Then
                                cur���� = .�������� / .�������� * (Val("" & rsTmp.Fields("��������")) / rsTmp.Fields("����ϵ��"))
                            Else
                                cur���� = .�������� * Abs(Int(0 - objBIll.���� / rsTmp!����ϵ��))
                            End If
                            
                            If rsTmp.Fields("�Ĵ�����") - cur���� < 0 Then cur���� = rsTmp.Fields("�Ĵ�����")
                            
                            strSQL = "Zl_�ݴ�ҩƷ��¼_Adviceused('" & rsTmp.Fields("NO") & "'," & rsTmp.Fields("���") & "," & objBIll.ҽ��ID & "," & .���ͺ� & "," & _
                                     rsTmp.Fields("ҩƷID") & "," & cur���� & ",'" & p�ӵ��� & "',To_Date('" & Format(p����ִ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = strSQL
                        End If
                    Next
                Next
            Else
                MsgBox "ҽ�� " & .�÷� & " ����ִ�� " & .�������� & " �Σ���ǰ�Ѿ�ִ���� " & .��ִ������ & " �Σ�������ִ�С�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        End With
    Next
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "Groups.SelectGroupThingNew")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    SelectGroupThingNew = True

    If lng��ˮ�� <> 0 Then
        If blnPrint Then
            strRpt = "ZL1_BILL_1264_" & intType
            Call ReportOpen(gcnOracle, glngSys, strRpt, frmParent, "�ӵ���ˮ��=" & lng��ˮ��, IIf(blnView, 1, 2))
        End If
        If blnLabel Then
            '��ӡ��Һƿǩ
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_4", frmParent, "�ӵ���ˮ��=" & lng��ˮ��, IIf(blnView, 1, 2))
        End If
        If blnWristband Then
            '��ӡ��Һ���
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_5", frmParent, "�ӵ���ˮ��=" & lng��ˮ��, IIf(blnView, 1, 2))
        End If
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetOtherValue(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long)
'���ܣ�ȡȡ �ϴ�ִ��ʱ��,��ִ�д���,�ϴ���Σ�Ҫ��ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objGroup As Group
    Dim strData As String
 
    On Error GoTo errH
    
    Set objGroup = Me.Item(lngҽ��ID & "_" & lng���ͺ�)
    If objGroup Is Nothing Then
        MsgBox "δ�ҵ�ҽ���������쳣��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�·�ʽ����
    strSQL = "Select �ϴ�ִ��ʱ��,�ϴ�Ҫ��ʱ��,�ϴ����,��ִ������ " & vbNewLine & _
        " From (Select ִ��ʱ�� as �ϴ�ִ��ʱ��, Ҫ��ʱ�� as �ϴ�Ҫ��ʱ��, ��� as �ϴ����," & vbNewLine & _
        " Count(��������) Over(Partition By ҽ��id, ���ͺ�) as ��ִ������ " & vbNewLine & _
        " From ����ҽ��ִ�� " & vbNewLine & _
        " Where ҽ��id = [1] And ���ͺ� = [2] " & vbNewLine & _
        " Order By ��ˮ�� Desc) " & vbNewLine & _
        " Where Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Groups.GetOtherValue1", lngҽ��ID, lng���ͺ�)
    If rsTmp.EOF Then
        objGroup.�ϴ�ִ��ʱ�� = CDate(0)
        objGroup.��ִ������ = 0
        objGroup.�ϴ���� = 0
    Else
        objGroup.�ϴ�ִ��ʱ�� = zlCommFun.NVL(rsTmp!�ϴ�ִ��ʱ��, CDate(0))
        objGroup.��ִ������ = zlCommFun.NVL(rsTmp!��ִ������, 0)
        objGroup.�ϴ���� = zlCommFun.NVL(rsTmp!�ϴ����, 0)
        strData = Format(zlCommFun.NVL(rsTmp!�ϴ�Ҫ��ʱ��), "yyyy-MM-dd HH:mm:ss")
    End If
    objGroup.Ҫ��ʱ�� = GetAdviceҪ��ʱ��(lngҽ��ID, lng���ͺ�, strData)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RecalculateTime(ByVal datʱ�� As Date, ByVal str�����λ As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal int���ʹ��� As Integer) As String
'���ܣ����̶�һ����ʼʱ��㣬���¼���ִ��ʱ�䷽��
'������
'���أ�ִ��ʱ�䷽��

    Const DBL_START As Double = 0.333333    '��08��00
    
    Dim i As Integer
    Dim dblBase As Double
    Dim strTimes As String
    
    Select Case Trim(str�����λ)
        Case "��"
            dblBase = intƵ�ʼ�� * 7 / intƵ�ʴ���
        Case "��"
            dblBase = intƵ�ʼ�� / intƵ�ʴ���
        Case "Сʱ"
            dblBase = intƵ�ʼ�� / intƵ�ʴ��� / 24
    End Select
    
    If Format(datʱ��, "yyyy-mm-dd hh:mm:ss") = Format(datʱ��, "yyyy-mm-dd 00:00:00") Then
        datʱ�� = datʱ�� + DBL_START   '���û��ʱ�䣬�̶���08��00
    End If
    
    For i = 0 To int���ʹ��� - 1
        strTimes = strTimes & "," & Format(datʱ�� + i * dblBase, "yyyy-mm-dd hh:mm:ss")
    Next
    If strTimes <> "" Then strTimes = Mid(strTimes, 2)

    RecalculateTime = strTimes
End Function

Private Function GetAdviceҪ��ʱ��(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal strCurDate As String) As Date
'���ܣ�����ҽ��ID�����ͺź��ϴ�ִ��ʱ����㱾��Ҫ��ʱ��
'      ҽ��ID�����ͺţ�
'      strCurDate �ϴ�ִ��ʱ�䣬Ϊ�ձ�ʾ��δִ�� �̶���ʽ��yyyy-MM-dd HH:mm:ss
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim datҪ��ʱ�� As Date
    Dim strPause As String
    Dim strTime  As String
    Dim vDate As Date
    Dim arrTime As Variant
    Dim i As Long
    
    On Error GoTo errH
    
    '���㱾��ִ��Ӧ�õ�Ҫ��ʱ��
    strSQL = "Select A.�״�ʱ��,A.ĩ��ʱ��," & _
        " B.��ʼִ��ʱ��,B.ִ����ֹʱ��,B.�ϴ�ִ��ʱ��,B.ִ��ʱ�䷽��," & _
        " B.ִ��Ƶ��,B.Ƶ�ʴ���,B.Ƶ�ʼ��,B.�����λ,B.ҽ����Ч" & _
        " From ����ҽ������ A,����ҽ����¼ B" & _
        " Where A.ҽ��ID=B.ID And A.ҽ��ID=[1] And A.���ͺ�=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Groups.GetAdviceҪ��ʱ��", lngҽ��ID, lng���ͺ�)
        
    If rsTmp!ִ��Ƶ�� & "" = "һ����" Or rsTmp!ִ��Ƶ�� & "" = "��Ҫʱ" Then
        'Ϊһ����ִ�е�����
        datҪ��ʱ�� = rsTmp!��ʼִ��ʱ��
    ElseIf strCurDate = "" Then
        '��һ��ִ��ʱ,��Ϊ�״�ʱ��
        datҪ��ʱ�� = rsTmp!�״�ʱ��
    Else
        '����ִ��Ƶ�ʷֽ�ʱ��
        If Val(rsTmp!ҽ����Ч & "") = 0 Then
            '�����Ż�����ͣʱ���
            strPause = GetAdvicePause(lngҽ��ID)
        End If
        If IsNull(rsTmp!ִ��ʱ�䷽��) And (NVL(rsTmp!Ƶ�ʴ���, 0) = 0 Or NVL(rsTmp!Ƶ�ʼ��, 0) = 0 Or IsNull(rsTmp!�����λ)) Then
            '�����Գ���
            strTime = ""
            vDate = Format(rsTmp!�״�ʱ��, "yyyy-MM-dd")
            Do While vDate <= Format(rsTmp!ĩ��ʱ��, "yyyy-MM-dd")
                If Not DateIsPause(vDate, strPause) Then
                    If Format(vDate, "yyyy-MM-dd") > Format(strCurDate, "yyyy-MM-dd") And strTime = "" Then
                        strTime = Format(vDate, "yyyy-MM-dd")
                    End If
                End If
                vDate = vDate + 1
            Loop
            If strTime = "" Then strTime = "0"
            datҪ��ʱ�� = CDate(strTime)
        Else
            vDate = Calc�����ڿ�ʼʱ��(rsTmp!��ʼִ��ʱ��, rsTmp!�״�ʱ��, rsTmp!Ƶ�ʼ��, rsTmp!�����λ)
            strTime = Calc���ڷֽ�ʱ��(vDate, rsTmp!ĩ��ʱ��, strPause, rsTmp!ִ��ʱ�䷽��, rsTmp!Ƶ�ʴ���, rsTmp!Ƶ�ʼ��, rsTmp!�����λ, rsTmp!��ʼִ��ʱ��)
            arrTime = Split(strTime, ",")
            datҪ��ʱ�� = rsTmp!��ʼִ��ʱ��
            For i = 0 To UBound(arrTime)
                If arrTime(i) > strCurDate Then
                    datҪ��ʱ�� = CDate(arrTime(i))
                    Exit For '�Ե�һ��ʱ��ΪҪ��ʱ��
                End If
            Next
            If i > UBound(arrTime) Then
                datҪ��ʱ�� = CDate(arrTime(0))
            End If
        End If
    End If
    GetAdviceҪ��ʱ�� = datҪ��ʱ��
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function