VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExecRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ExecutiveGroup"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function GetExecGroups(ByVal objPati As cPatient, ByVal lngִ�в���ID As Long, ByVal intִ����� As Integer, _
    ByVal dateBeging As Date, ByVal DateEnd As Date) As Boolean
    
    Dim strSQL As String, i As Integer
    Dim rsExecGroups As ADODB.Recordset, rsGroup As ADODB.Recordset, rsBill As ADODB.Recordset
    Dim objexecGroup As ExecutiveGroup, objGroup As Group, objBIlls As Bills, lngʱ�� As Long, Date�ϴ�ִ��ʱ�� As Date
    Dim objBIll As Bill, lng�ܺ�ʱ As Long, lngҺ���� As Long, int��ִ�д��� As Integer
    Dim rsTmp As ADODB.Recordset, str˵�� As String, lng_Count As Long
    Dim int�Ʒ�״̬ As Integer, cur��� As Currency
    
    On Error GoTo errHandle
    
    If objPati.������Դ = 1 Then
        '��������
        strSQL = "Select distinct A.��ˮ��, A.��ҩ��, A.ִ��ʱ��, A.����, A.�ӵ���, A.��ϵ��," & _
                 "  Decode(C.ִ�з���, 1, '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����') As ִ�з���,d.�Ʒ�״̬" & vbNewLine & _
                 "From ����ҽ������ D, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ��ִ�� A" & vbNewLine & _
                 "Where A.ҽ��id = D.ҽ��id And A.���ͺ� = D.���ͺ� And A.ҽ��id = B.ID And B.������Ŀid = C.ID And B.��ҳid=[5] And " & vbNewLine & _
                 "   Nvl(A.���,0) <= 1 And (A.��ˮ�� is not Null or A.��ˮ��<>0) And D.����ʱ�� Between [3] And [4] And " & _
                 "   D.ִ�в���id = [2] And B.����id = [1] " & vbCr & _
                 "Order by A.��ˮ�� "
        Set rsExecGroups = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", objPati.����ID, lngִ�в���ID, dateBeging, DateEnd, objPati.����ID)
    Else
        strSQL = "Select distinct A.��ˮ��, A.��ҩ��, A.ִ��ʱ��, A.����, A.�ӵ���, A.��ϵ��,Decode(C.ִ�з���, 1, '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����') As ִ�з���,d.�Ʒ�״̬" & vbNewLine & _
                "From ����ҽ������ D, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ��ִ�� A" & vbNewLine & _
                "Where A.ҽ��id = D.ҽ��id And A.���ͺ� = D.���ͺ� And A.ҽ��id = B.ID And B.������Ŀid = C.ID And B.�Һŵ�=[5] And " & vbNewLine & _
                "   Nvl(A.���,0) <= 1 And (A.��ˮ�� is not Null or A.��ˮ��<>0) And D.����ʱ�� Between [3] And [4] And D.ִ�в���id = [2] And B.����id = [1] Order by A.��ˮ�� "
        Set rsExecGroups = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", objPati.����ID, lngִ�в���ID, dateBeging, DateEnd, objPati.�Һŵ�)
    End If
    For i = 1 To mCol.Count
        mCol.Remove 1
    Next
    Do Until rsExecGroups.EOF
    
        Set objexecGroup = New ExecutiveGroup
        objexecGroup.��ˮ�� = zlCommFun.NVL(rsExecGroups.Fields("��ˮ��"), 0)
        objexecGroup.��ҩ�� = zlCommFun.NVL(rsExecGroups.Fields("��ҩ��"))
        objexecGroup.ִ��ʱ�� = zlCommFun.NVL(rsExecGroups.Fields("ִ��ʱ��"), CDate(0))
        objexecGroup.���� = zlCommFun.NVL(rsExecGroups.Fields("����"), 0)
        objexecGroup.�ӵ��� = zlCommFun.NVL(rsExecGroups.Fields("�ӵ���"))
        objexecGroup.��ϵ�� = zlCommFun.NVL(rsExecGroups.Fields("��ϵ��"), 20)
        objexecGroup.ִ�з��� = zlCommFun.NVL(rsExecGroups.Fields("ִ�з���"), "0-����")
        
        
        lng�ܺ�ʱ = 0
        
        '2012-07-17 ��ʾ�����
        strSQL = "Select d.��ˮ��, d.���, d.����, d.Һ����, a.ҽ��id, d.ִ����,d.�˶���, a.���ͺ�, a.��������, d.��������, b.�ϴ�ִ��ʱ��, d.Ҫ��ʱ��," & vbNewLine & _
                "       c.��������,Decode(c.ִ�з���, 1, '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����') As ִ�з���, d.����, b.ҽ������ As �÷�, b.Ƥ�Խ��, b.ִ��Ƶ��, b.Ƶ�ʴ���," & vbNewLine & _
                "       a.ִ��״̬, a.�Ʒ�״̬, b.�շ�ϸĿid, a.��¼����, b.ִ������, Sum(Nvl(Decode(e.��¼״̬,9,0,3,0,e.ʵ�ս��), 0)) As ���, e.��¼״̬" & vbNewLine & _
                "From ����ҽ��ִ�� D, ������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ������ A, ������ü�¼ E" & vbNewLine & _
                "Where a.ҽ��id = d.ҽ��id And a.���ͺ� = d.���ͺ� And b.������Ŀid = c.Id And a.ҽ��id = b.Id And c.��� = 'E' And a.ִ�в���id = [2] And" & vbNewLine & _
                "      a.����ʱ�� Between [3] And [4] And d.��ˮ�� = [5] And b.������Դ = [6] And b.����id = [1] And a.ҽ��id = e.ҽ�����(+) And" & vbNewLine & _
                "      a.��¼��� = e.���(+) And a.No = e.No(+) And e.��¼״̬(+) <> 2" & vbNewLine & _
                "Group By d.��ˮ��, d.���, d.����, d.Һ����, a.ҽ��id, d.ִ����,d.�˶���, a.���ͺ�, a.��������, d.��������, b.�ϴ�ִ��ʱ��, d.Ҫ��ʱ��," & vbNewLine & _
                "         c.��������,Decode(c.ִ�з���, 1, '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����'), d.����, b.ҽ������, b.Ƥ�Խ��, b.ִ��Ƶ��, b.Ƶ�ʴ���, a.ִ��״̬, a.�Ʒ�״̬," & vbNewLine & _
                "         b.�շ�ϸĿid, a.��¼����, b.ִ������, e.��¼״̬" & vbNewLine & _
                "Order By d.���"

        Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", _
                                               objPati.����ID, lngִ�в���ID, dateBeging, DateEnd, objexecGroup.��ˮ��, _
                                               IIf(objPati.������Դ = 1, 2, 1))
        With rsGroup
        Do Until .EOF
            strSQL = "Select A.˵�� From ����ҽ��ִ�� A Where A.ҽ��id = [1] And A.���ͺ� = [2] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", Val(.Fields("ҽ��ID")), Val(.Fields("���ͺ�")))
            str˵�� = ""
            lng_Count = 0
            Do Until rsTmp.EOF
                lng_Count = lng_Count + 1
                str˵�� = str˵�� & rsTmp!˵��
                If lng_Count <> rsTmp.RecordCount Then
                    str˵�� = str˵�� & vbNewLine
                End If
                rsTmp.MoveNext
            Loop
            
            Select Case .Fields("ִ�з���")
            Case "0-����"
                '����  ���ָ�ҩ;���ͷǸ�ҩ;��
                If Val(.Fields("��������") & "") = 2 Then
                    '���Ǹ�ҩ;��ʱ�� ��Һ,ע�� ��ѯ��SQL��ͬ
                    strSQL = "Select ����, Һ����, ���, Id, Rownum ˳��, ִ��Ƶ��, ҽ������, ��������, ���㵥λ, ��������, ִ��Ƶ��,ʵ�ս��,�Ʒ�״̬,��¼״̬,ִ������" & vbNewLine & _
                            "From (Select Distinct d.����, d.Һ����, c.���, b.Id, b.ҽ������, b.��������, c.���㵥λ," & vbNewLine & _
                            "        Decode(Sign(Instr(',5,6,', ',' || c.��� || ',')), 1, ( b.��������/g.����ϵ��) * g.����, 0) As ��������," & vbNewLine & _
                            "        b.ִ��Ƶ�� , b.���,Decode(E.��¼״̬,9,0,3,0,E.ʵ�ս��) as ʵ�ս��,f.�Ʒ�״̬,e.��¼״̬,b.ִ������" & vbNewLine & _
                            "      From ����ҽ������ f,������ü�¼ e,����ҽ��ִ�� d, ҩƷ��� g, ������ĿĿ¼ c, ����ҽ����¼ b, ����ҽ������ a" & vbNewLine & _
                            "      Where b.id=f.ҽ��id and f.���ͺ�=[5] and f.ҽ��id=E.ҽ�����(+) And F.NO=E.NO(+) And E.��¼״̬(+)<>2 and a.ҽ��id = d.ҽ��id And a.���ͺ� = d.���ͺ� And b.�շ�ϸĿid = g.ҩƷid(+) And b.������Ŀid = c.Id And" & vbNewLine & _
                            "        a.ҽ��id = b.���id And a.ִ�в���id = [1] And" & vbNewLine & _
                            "        a.����ʱ�� Between [2] And [3] And a.ҽ��id = [4] And a.���ͺ� = [5] And" & vbNewLine & _
                            "        b.������Դ = [8] And b.����id = [6] And D.��ˮ��=[7]" & vbNewLine & _
                            "      Order By b.��� ) "
                Else
                    strSQL = "Select Һ����,���,Id,Rownum ˳��,ִ��Ƶ��,ҽ������,��������,���㵥λ,��������,ִ��Ƶ��,ʵ�ս��,�Ʒ�״̬,��¼״̬,ִ������ " & vbNewLine & _
                            "From (" & vbNewLine & _
                            "  Select Distinct 0 as Һ����,C.���, B.ID, B.ִ��Ƶ��, B.ҽ������ || '��' || b1.ҽ������ ҽ������, B1.��������, C.���㵥λ, " & _
                            "      0 As ��������, B.���, Decode(D.��¼״̬,9,0,3,0,D.ʵ�ս��) as ʵ�ս��,a.�Ʒ�״̬,d.��¼״̬,b.ִ������" & vbNewLine & _
                            "  From  ������ü�¼ D,������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ����¼ B1, ����ҽ������ A" & vbNewLine & _
                            "  Where A.ҽ��id=D.ҽ�����(+) And A.NO=D.NO(+) and D.��¼״̬(+)<>2 and B1.������Ŀid = C.ID(+) And A.ҽ��id = B.ID And " & vbNewLine & _
                            "      A.ҽ��id = B1.���ID(+) And A.ִ�в���id = [1] And A.����ʱ�� Between [2] And [3] And A.ҽ��id = [4] And " & vbNewLine & _
                            "      A.���ͺ� = [5] And B.������Դ = [8] And B.����id = [6]" & vbNewLine & _
                            "  Order By B.��� ) "
                End If
            Case "3-Ƥ��"
                'Ƥ��
                strSQL = "Select  Һ����,���,Id,Rownum ˳��,ִ��Ƶ��,ҽ������,��������,���㵥λ,��������,ִ��Ƶ��,ʵ�ս��,�Ʒ�״̬,��¼״̬,ִ������" & vbNewLine & _
                        "From (" & vbNewLine & _
                        "  Select Distinct 0 as Һ����,C.���, B.ID,  B.ִ��Ƶ��, B.ҽ������, B.��������, C.���㵥λ, 0 As ��������, B.���, " & _
                        "    Decode(D.��¼״̬,9,0,3,0,D.ʵ�ս��) as ʵ�ս��,a.�Ʒ�״̬,d.��¼״̬,b.ִ������" & vbNewLine & _
                        "  From  ������ü�¼ D,������ĿĿ¼ C, ����ҽ����¼ B, ����ҽ������ A" & vbNewLine & _
                        "  Where A.ҽ��id=D.ҽ�����(+) And A.NO=D.NO(+) and D.��¼״̬(+)<>2 and B.������Ŀid = C.ID And A.ҽ��id = B.ID And A.ִ�в���id = [1] And" & vbNewLine & _
                        "      A.����ʱ�� Between [2] And [3] And A.ҽ��id = [4] And A.���ͺ� = [5] And B.������Դ = [8] And" & vbNewLine & _
                        "      B.����id = [6]" & vbNewLine & _
                        "  Order By B.��� ) "
            Case Else
                '��Һ,ע��
                strSQL = "Select ����, Һ����, ���, Id, Rownum ˳��, ִ��Ƶ��, ҽ������, ��������, ���㵥λ, ��������, ִ��Ƶ��,ʵ�ս��,�Ʒ�״̬,��¼״̬,ִ������" & vbNewLine & _
                        "From (Select Distinct d.����, d.Һ����, c.���, b.Id, b.ҽ������, b.��������, c.���㵥λ," & vbNewLine & _
                        "        Decode(Sign(Instr(',5,6,', ',' || c.��� || ',')), 1, ( b.��������/g.����ϵ��) * g.����, 0) As ��������," & vbNewLine & _
                        "        b.ִ��Ƶ�� , b.���,Decode(E.��¼״̬,9,0,3,0,E.ʵ�ս��) as ʵ�ս��,f.�Ʒ�״̬,e.��¼״̬,b.ִ������" & vbNewLine & _
                        "      From ����ҽ������ f,������ü�¼ e,����ҽ��ִ�� d, ҩƷ��� g, ������ĿĿ¼ c, ����ҽ����¼ b, ����ҽ������ a" & vbNewLine & _
                        "      Where b.id=f.ҽ��id and f.���ͺ�=[5] and f.ҽ��id=E.ҽ�����(+) And F.NO=E.NO(+) And E.��¼״̬(+)<>2 and a.ҽ��id = d.ҽ��id And a.���ͺ� = d.���ͺ� And b.�շ�ϸĿid = g.ҩƷid(+) And b.������Ŀid = c.Id And" & vbNewLine & _
                        "        a.ҽ��id = b.���id And a.ִ�в���id = [1] And" & vbNewLine & _
                        "        a.����ʱ�� Between [2] And [3] And a.ҽ��id = [4] And a.���ͺ� = [5] And" & vbNewLine & _
                        "        b.������Դ = [8] And b.����id = [6] And D.��ˮ��=[7]" & vbNewLine & _
                        "      Order By b.��� ) "

            End Select
            Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", lngִ�в���ID, dateBeging, DateEnd, _
                                                  Val(.Fields("ҽ��ID")), Val(.Fields("���ͺ�")), _
                                                  objPati.����ID, objexecGroup.��ˮ��, _
                                                  IIf(objPati.������Դ = 1, 2, 1))
            
            Set objBIlls = New Bills
            lngʱ�� = 0
            lngҺ���� = 0
            Do Until rsBill.EOF
                If rsBill.Fields("˳��") = 1 Then
                    lngʱ�� = CacleTransTime(Val(rsBill.Fields("Һ����")), objexecGroup.��ϵ��, .Fields("����"))
                    lngҺ���� = Val(rsBill.Fields("Һ����"))
                Else
                    lngʱ�� = 0
                    lngҺ���� = 0
                End If
                int�Ʒ�״̬ = Val("" & rsBill.Fields("�Ʒ�״̬"))

                cur��� = Val("" & rsBill.Fields("ʵ�ս��"))
                
                If Val("" & rsBill.Fields("��¼״̬")) = 0 Then cur��� = 0
                
                If Val("" & rsBill.Fields("ִ������")) = 5 And int�Ʒ�״̬ <> -1 Then
                    int�Ʒ�״̬ = -1 'Ժ���ҩ�ģ���ͬΪ����Ʒ�
                ElseIf Val("" & rsBill.Fields("��¼״̬")) = 1 And cur��� = 0 Then
                    int�Ʒ�״̬ = -2 '�����
                ElseIf Val("" & rsBill.Fields("��¼״̬")) = 3 Then
                    int�Ʒ�״̬ = -3 '���˷�
                End If
                
                Call objBIlls.Add(objPati.����ID, Val(rsBill.Fields("ID")), zlCommFun.NVL(rsBill.Fields("ҽ������")), zlCommFun.NVL(rsBill.Fields("��������"), 0), zlCommFun.NVL(rsBill.Fields("���㵥λ")), lngҺ����, _
                             lngʱ��, str˵��, Val(rsBill.Fields("˳��")), cur���, int�Ʒ�״̬, Val(.Fields("ҽ��ID")) & "_" & rsBill.Fields("ID"))
                rsBill.MoveNext
            Loop
            Date�ϴ�ִ��ʱ�� = zlCommFun.NVL(.Fields("�ϴ�ִ��ʱ��"), CDate(0))
            
            For Each objBIll In objBIlls
                lng�ܺ�ʱ = lng�ܺ�ʱ + objBIll.ʱ��
            Next
            
            strSQL = "Select Sum(nvl(��������,0)) as ��ִ������ From ����ҽ��ִ�� Where ҽ��id=[1] And ���ͺ�=[2]"
            Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "", Val(.Fields("ҽ��ID")), Val(.Fields("���ͺ�")))
            If rsBill.RecordCount > 0 Then
                int��ִ�д��� = IIf(IsNull(rsBill.Fields("��ִ������")), 0, Val(rsBill.Fields("��ִ������"))) - Val(.Fields("��������"))
            End If
            

            int�Ʒ�״̬ = Val("" & .Fields("�Ʒ�״̬"))
            cur��� = Val("" & .Fields("���"))
            
            If Val("" & .Fields("ִ������")) = 5 And int�Ʒ�״̬ <> -1 Then
                int�Ʒ�״̬ = -1 'Ժ���ҩ�ģ���ͬΪ����Ʒ�
            ElseIf Val("" & .Fields("��¼״̬")) = 1 And cur��� = 0 Then
                int�Ʒ�״̬ = -2 '�����
            ElseIf Val("" & .Fields("��¼״̬")) = 3 Then
                int�Ʒ�״̬ = -3 '���˷�
            End If
            
            If objexecGroup.Item(Val(.Fields("ҽ��ID")) & "_" & .Fields("���ͺ�")) Is Nothing Then
                Call objexecGroup.Add(Val(.Fields("ҽ��ID")), Val(.Fields("���ͺ�")), Val(.Fields("ִ��״̬")), zlCommFun.NVL(.Fields("ִ����")), CStr(.Fields("�÷�")), zlCommFun.NVL(.Fields("Ƶ�ʴ���"), 0), _
                                      Val(.Fields("����")), .Fields("ִ��Ƶ��"), zlCommFun.NVL(.Fields("Ƥ�Խ��")), CStr(.Fields("ִ�з���")), int��ִ�д���, _
                                      Val(.Fields("��������")), Val(.Fields("��������")), Val(.Fields("���")), objBIlls, Date�ϴ�ִ��ʱ��, CDate(.Fields("Ҫ��ʱ��")), _
                                      int�Ʒ�״̬, Val("" & .Fields("��¼����")), "" & .Fields("�˶���"), Val(.Fields("ҽ��ID")) & "_" & .Fields("���ͺ�"))
            End If
            
            .MoveNext
        Loop
        End With
        objexecGroup.�ܺ�ʱ = lng�ܺ�ʱ
        
        Call Add(objexecGroup, zlCommFun.NVL(rsExecGroups.Fields("��ˮ��"), 0))
        rsExecGroups.MoveNext
    Loop
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Add(ByVal ObjExecutiveGroup As ExecutiveGroup, ByVal sKey As String) As ExecutiveGroup
    '�����¶���
    Dim objNewMember As ExecutiveGroup
    Set objNewMember = ObjExecutiveGroup
    On Error GoTo errH
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
errH:
    If Err.Number = 457 Then
        Exit Function
    ElseIf ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Property Get Item(vntIndexKey As Variant) As ExecutiveGroup
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


