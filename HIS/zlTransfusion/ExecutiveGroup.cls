VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExecutiveGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Group"
Attribute VB_Ext_KEY = "Member0" ,"Group"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection
'��������ֵ�ľֲ�����
Private mvar��ˮ�� As Long '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar��ҩ�� As String '�ֲ�����
Private mvarִ��ʱ�� As Date '�ֲ�����
Private mvar���� As Integer '�ֲ�����
Private mvar�ӵ��� As String '�ֲ�����
Private mvar��ϵ�� As Integer '�ֲ�����
Private mvar�ܺ�ʱ As Long
Private mvarִ�з��� As String
 
Public Property Let ��ϵ��(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�� = 5
    mvar��ϵ�� = vData
End Property

Public Property Get ��ϵ��() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.��ϵ��
    ��ϵ�� = mvar��ϵ��
End Property

Public Property Let �ܺ�ʱ(ByVal vData As Long)
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.��ϵ��
    mvar�ܺ�ʱ = vData
End Property

Public Property Get �ܺ�ʱ() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.��ϵ��
    Dim objItem As Group
    mvar�ܺ�ʱ = 0
    For Each objItem In mCol
     mvar�ܺ�ʱ = mvar�ܺ�ʱ + objItem.��ʱ
    Next
    �ܺ�ʱ = mvar�ܺ�ʱ
End Property

Public Property Let �ӵ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ӵ��� = 5
    mvar�ӵ��� = vData
End Property

Public Property Get �ӵ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�ӵ���
    �ӵ��� = mvar�ӵ���
End Property

Public Property Let ִ�з���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ӵ��� = 5
    mvarִ�з��� = vData
End Property

Public Property Get ִ�з���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.�ӵ���
    ִ�з��� = mvarִ�з���
End Property

Public Property Let ����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property

Public Property Let ִ��ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ��ʱ�� = 5
    mvarִ��ʱ�� = vData
End Property

Public Property Get ִ��ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��ʱ��
    ִ��ʱ�� = mvarִ��ʱ��
End Property

Public Property Let ��ҩ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ҩ�� = 5
    mvar��ҩ�� = vData
End Property

Public Property Get ��ҩ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.��ҩ��
    ��ҩ�� = mvar��ҩ��
End Property

Public Property Let ��ˮ��(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ˮ�� = 5
    mvar��ˮ�� = vData
End Property

Public Property Get ��ˮ��() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.��ˮ��
    ��ˮ�� = mvar��ˮ��
End Property

Public Function Add(ִ��ҽ��ID As Long, ���ͺ� As Long, ִ��״̬ As Integer, ִ���� As String, �÷� As String, Ƶ�ʴ��� As Integer, ���� As Integer, ִ��Ƶ�� As String, Ƥ�Խ�� As String, ִ�з��� As String, ��ִ������ As Integer, �������� As Integer, �������� As Integer, ��� As Integer, BillsItem As Bills, �ϴ�ִ��ʱ�� As Date, Ҫ��ʱ�� As Date, �Ʒ�״̬ As Integer, ��¼���� As Integer, �˶��� As String, Optional sKey As String) As Group
'�����¶���
    Dim objNewMember As Group, rsTmp As New ADODB.Recordset, strSQL As String
    Set objNewMember = New Group
    On Error GoTo hErr
    

    '���ô��뷽��������
    objNewMember.ִ��ҽ��ID = ִ��ҽ��ID
    objNewMember.���ͺ� = ���ͺ�
    objNewMember.ִ��״̬ = ִ��״̬
    objNewMember.ִ���� = ִ����
    objNewMember.�÷� = �÷�
    objNewMember.Ƶ�ʴ��� = Ƶ�ʴ���
    objNewMember.���� = ����
    objNewMember.ִ��Ƶ�� = ִ��Ƶ��
    objNewMember.Ƥ�Խ�� = Ƥ�Խ��
    objNewMember.ִ�з��� = ִ�з���
    objNewMember.��ִ������ = ��ִ������
    objNewMember.�������� = ��������
    objNewMember.�������� = ��������
    objNewMember.��� = ���
    objNewMember.�Ʒ�״̬ = �Ʒ�״̬
    objNewMember.��¼���� = ��¼����
    objNewMember.�˶��� = �˶���
    
    '-- �շѽ�� 2007-07-18
    objNewMember.�շѽ�� = 0
    '--   2011-09-21 ����ִ��״̬<>9
    strSQL = "Select Sum(Nvl(B.ʵ�ս��, 0)) As ���" & vbNewLine & _
            "From ������ü�¼ B, ����ҽ������ A" & vbNewLine & _
            "Where A.ҽ��id = B.ҽ����� And A.NO = B.NO And B.��¼״̬<>0 And B.����״̬<>1 And a.���ͺ� = [2] And A.ҽ��id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecutiveGroup.Add", ִ��ҽ��ID, ���ͺ�)
    Do Until rsTmp.EOF
        objNewMember.�շѽ�� = IIf(IsNull(rsTmp.Fields("���")), 0, rsTmp.Fields("���"))
        rsTmp.MoveNext
    Loop
    
    If IsObject(BillsItem) Then
        Call objNewMember.AddBills(BillsItem, ִ��ҽ��ID & "_" & ���ͺ�)
    End If
    objNewMember.�ϴ�ִ��ʱ�� = �ϴ�ִ��ʱ��
    objNewMember.Ҫ��ʱ�� = Ҫ��ʱ��
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

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
    On Error Resume Next
End Sub

Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub

Public Sub Undo(ByVal lng��ˮ�� As Long, ByVal lngExcDeptID As Long, Optional ByRef lngErrNo As Long)
    '�����ӵ�
    'lngExcDeptID :ִ�п���ID 47323
    Dim strSQL() As String, i As Integer, bln����� As Boolean
    Dim strBill() As String
    
    Dim objGroup As Group
    Dim intBills As Integer, intBill As Integer
    Dim objBIlls As Bills, objBIll As Bill, strTmp As String, rsTmp As New ADODB.Recordset, Y As Integer
    Dim blnTrans As Boolean
    On Error GoTo errHandle
    If mvar��ˮ�� <> lng��ˮ�� Then Exit Sub
    
    ReDim strSQL(mCol.Count)
    
    For Each objGroup In mCol
        If objGroup.ִ��״̬ = 1 Then
        
            bln����� = True
            strSQL(i) = ""
        Else
            '��ˮ�Ŷ�Ӧ��ҽ����ִ��ʱ����ܻ�ı�
            strTmp = "Select ִ��ʱ�� From ����ҽ��ִ�� Where ҽ��ID=[1] And ���ͺ�=[2] And ��ˮ��=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "ȡҽ��ִ��ʱ��", objGroup.ִ��ҽ��ID, objGroup.���ͺ�, lng��ˮ��)
            If Not rsTmp.EOF Then mvarִ��ʱ�� = rsTmp!ִ��ʱ��
            
            strSQL(i) = "ZL_����ҽ��ִ��_Delete(" & objGroup.ִ��ҽ��ID & "," & objGroup.���ͺ� & ",to_date('" & Format(mvarִ��ʱ��, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),0,0," & lngExcDeptID & ")"
            '�ݴ�ҩƷ����
            For intBills = 1 To objGroup.BillsCount
                Set objBIlls = objGroup.BillsItem(intBills)
                For intBill = 1 To objBIlls.Count
                    Set objBIll = objBIlls.Item(intBill)
                    strTmp = "Select NO,���,to_char(�Ǽ�ʱ��,'yyyy-MM-dd HH24:MI:SS') as �Ǽ�ʱ�� ,���ϵ�� From  �ݴ�ҩƷ��¼ Where ���ϵ��=-1 And �Ǽ�ʱ�� = [3] And ҽ��ID=[1] And ���ͺ�=[2] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "ExecutiveGroup.Undo", objBIll.ҽ��ID, objGroup.���ͺ�, mvarִ��ʱ��)
                    Do Until rsTmp.EOF
                        ReDim Preserve strBill(Y + 1)
                        
                        strBill(Y) = "zl_�ݴ�ҩƷ��¼_undouse('" & rsTmp.Fields("NO") & "'," & rsTmp.Fields("���") & ",-1,TO_date('" & rsTmp.Fields("�Ǽ�ʱ��") & "','yyyy-MM-dd HH24:MI:SS'))"
                        rsTmp.MoveNext
                        Y = Y + 1
                    Loop
                Next
            Next
        End If
        i = i + 1
    Next
    LoadPicture
    If bln����� Then
        If MsgBox("�������ĵ����к����������Ŀ����Щ��Ŀ���ܳ������Ƿ������" & vbNewLine & "��[��]��ֻ����δ�����Ŀ����[��]����ִ���κβ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    blnTrans = False
    gcnOracle.BeginTrans
    blnTrans = True
    For i = 0 To mCol.Count - 1
        If strSQL(i) <> "" Then
            If strSQL(i) <> "" Then
                Call zlDatabase.ExecuteProcedure(strSQL(i), "executivegroup.Undo")
            End If
        End If
    Next
    
    If Y > 0 Then
        For i = 0 To UBound(strBill) - 1
            If strBill(i) <> "" Then
                Call zlDatabase.ExecuteProcedure(strBill(i), "executivegroup.Undo")
            End If
        Next
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    Exit Sub
    
errHandle:
    Call ErrCenter
    If blnTrans = True Then gcnOracle.RollbackTrans
    Call SaveErrLog
    lngErrNo = Err.Number
End Sub

Public Function ExecuteFinish(ByVal strGroupKey As String, ByVal lngExcDeptID As Long, Optional strSQLTest As String = "", Optional frmMain As Object, Optional objCardSquare As Object) As Boolean
    '����ҽ����ִ����ɹ���
    ' strGroupKey: ��λһ��ҽ����KEY
    ' frmMain    : Ϊ����һ��ͨ�ӿڴ���������� ����
    'objCardSquare: Ϊ����һ��ͨ�ӿڴ����һ��ͨ����, objCardSquare ΪNothing ���ʾ����δ��ʼ�ɹ���
    'strSQLTest :��Ƥ�Ե�ִ�й�����������Ŀ��ͬ�������������Ҳ������Ƥ����ɹ��̵���,����Ƥ�Թ��̵���ʱ���˲�����Ϊ�ա�
    'lngExcDeptID : ִ�п���ID��47323�����޸�
    Dim strSQL As String, rsTmp As New ADODB.Recordset, strInfo As String

    Dim blnTrans As Boolean, intReturn As Integer, strPrivs As String
    Dim lngPatiId As Long
    '�Ƿ��������δ�շѲ��˵���Ŀ
    On Error GoTo errHandle
 
    '2012-07-17 ȡ����ID����������ȡ����λռ��
    
    strSQL = "Select ����ID,������� From ����ҽ����¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ҽ��ִ�����", Item(strGroupKey).ִ��ҽ��ID)
    If Not rsTmp.EOF Then lngPatiId = Val("" & rsTmp!����ID)
    
    
    '����һ��ͨ��������
    intReturn = OneCardCheck(Item(strGroupKey).ִ��ҽ��ID, Item(strGroupKey).���ͺ�, frmMain, objCardSquare)
    
    If intReturn = 2 Then
        '������ʧ��,�ڲ�����ʾ���˴�ֱ���˳�
        Exit Function
    ElseIf intReturn = 0 Then
        '--������
        If Val(zlDatabase.GetPara("δ�շ����", glngSys, 1264)) = 0 Then
                If Item(strGroupKey).�շѽ�� = 0 Then
                    '����ֻ����շѼ�¼�����ۼ�¼��Ϊ��ִ�к�������������ǰ���Ѽ��
                    
                    strSQL = "Select B.��¼״̬,A.�Ʒ�״̬,b.����״̬ ,D.�շ�ϸĿid, a.NO,B.No as �շ�NO, B.�վݷ�Ŀ, D.ҽ������, To_Char(A.����ʱ��, 'yyyy-MM-dd HH24:MI:SS') As ����ʱ��," & vbNewLine & _
                            "       Sum(Nvl(B.ʵ�ս��, 0)) As ���" & vbNewLine & _
                            "From ����ҽ����¼ D, ������ü�¼ B, ����ҽ������ A" & vbNewLine & _
                            "Where A.ҽ��id = D.Id And A.ҽ��id = B.ҽ�����(+) And A.No = B.No(+) And  nvl(B.��¼״̬(+),0)<=1 And A.���ͺ� = [2] And" & vbNewLine & _
                            "      mod(B.��¼����(+),10) = 1 And A.ҽ��id = [1]" & vbNewLine & _
                            "Group By B.��¼״̬,A.�Ʒ�״̬, b.����״̬, D.�շ�ϸĿid, A.No,B.NO, B.�վݷ�Ŀ, D.ҽ������, To_Char(A.����ʱ��, 'yyyy-MM-dd HH24:MI:SS')"
    
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecutiveGroup.ExecuteFinish", Item(strGroupKey).ִ��ҽ��ID, Item(strGroupKey).���ͺ�)
                    With rsTmp
                    Do Until .EOF
                        If !�Ʒ�״̬ = -1 Then
                            '����Ʒѵ���Ŀ
                            
                            '2011-09-21 ����  (Val("" & !��¼״̬ <> 0) And Val("" & !ִ��״̬ = 9))
                        ElseIf Val("" & !��¼״̬) = 0 Or (Val("" & !��¼״̬ <> 0) And Val(zlCommFun.NVL(!����״̬, 0) = 1)) Then
                            If strInfo = "" Then
                                strInfo = " NO: " & .Fields("NO") & " ����ʱ��:" & .Fields("����ʱ��") & vbNewLine
                            End If
                            strInfo = strInfo & "   ҽ��:" & .Fields("ҽ������")
                            If Trim("" & .Fields("�շ�ϸĿid")) = "" Then
                                strInfo = strInfo & "   δ��Ӧ������ϸ!"
                            ElseIf Trim("" & .Fields("�շ�NO")) = "" Then
                                strInfo = strInfo & "   δ���ɷ�����ϸ!"
                            Else
                                strInfo = strInfo & " ���:" & Format(.Fields("���"), "0.00") & vbNewLine
                            End If
                        End If
                        .MoveNext
                    Loop
                    End With
                    If strInfo <> "" Then
                        MsgBox "����Ŀ��δ�շѣ�������ɡ�" & vbNewLine & strInfo, vbInformation, gstrSysName
                        Exit Function
                    End If
            End If
        End If  'ģ����Ȩ�ޣ���δ�շ���ɡ�
    End If  '�������̼��
    
    blnTrans = False
    gcnOracle.BeginTrans
    blnTrans = True
    If strSQLTest <> "" Then
        Call zlDatabase.ExecuteProcedure(strSQLTest, "ExecutiveGroup.ExecuteFinish")
    End If

    '2012-07-17 51194���� �˹��� �����Զ������Ĺ��� ��ϵͳ����[�ٴ�Ӧ��]��[ִ��֮������������Զ�����]��Ҫ���ϲ���Ч��
    strSQL = "ZL_����ҽ��ִ��_Finish(" & Item(strGroupKey).ִ��ҽ��ID & "," & Item(strGroupKey).���ͺ� & ",Null,0,Null,Null," & lngExcDeptID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.ExecuteFinish")
    gcnOracle.CommitTrans
    blnTrans = False
        
    ExecuteFinish = True
    Exit Function
errHandle:
    
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    
    Call SaveErrLog
End Function

Public Sub Update(ByVal str��ˮ�� As String, ByVal strGroupKey As String, ByVal lngExcDeptID As Long, Optional ByRef lngErrNo As Long)
    '�����޸�����
    'lngExcDeptID :ִ�п���ID 47323����
    
    Dim objGroup As Group, strSQL As String, rsTmp As ADODB.Recordset
    Dim dateԭִ��ʱ�� As Date, strԭժҪ As String, blnBegin As Boolean
    On Error GoTo errHandle
    
    strSQL = "Select ִ��ʱ��,ִ��ժҪ From ����ҽ��ִ�� Where ��ˮ��=[1] and ҽ��ID=[2] and ���ͺ�=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecutiveGroup.Update", ��ˮ��, Item(strGroupKey).ִ��ҽ��ID, Item(strGroupKey).���ͺ�)
    Do Until rsTmp.EOF
        dateԭִ��ʱ�� = zlCommFun.NVL(rsTmp!ִ��ʱ��, CDate(0))
        strԭժҪ = zlCommFun.NVL(rsTmp!ִ��ժҪ)
        rsTmp.MoveNext
    Loop
    
    strSQL = "ZL_����ҽ��ִ��_Update(to_date('" & Format(dateԭִ��ʱ��, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                            Item(strGroupKey).ִ��ҽ��ID & "," & Item(strGroupKey).���ͺ� & "," & _
                                            "to_date('" & Format(Item(strGroupKey).Ҫ��ʱ��, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                            Item(strGroupKey).�������� & ",'" & strԭժҪ & "','" & Item(strGroupKey).ִ���� & "'," & _
                                            "to_date('" & Format(ִ��ʱ��, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,Null,0,Null,Null," & lngExcDeptID & ")"
    
    gcnOracle.BeginTrans
    blnBegin = True
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.Update")
    strSQL = "Zl_����ҽ��ִ��_Modify(" & ��ˮ�� & "," & Item(strGroupKey).ִ��ҽ��ID & "," & Item(strGroupKey).���ͺ� & "," & _
                                Item(strGroupKey).���� & "," & Item(strGroupKey).Һ���� & "," & ��ϵ�� & "," & Item(strGroupKey).��ʱ & ",'" & _
                                Item(strGroupKey).˵�� & "','" & _
                                UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.Update")
    gcnOracle.CommitTrans
    blnBegin = False
    Exit Sub
errHandle:
    lngErrNo = Err.Number
    Call ErrCenter
    If blnBegin Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Public Function ExecuteTestFinish(ByVal strGroupKey As String, ByVal frmMain As Object, ByVal objCardSquare As Object, ByRef strƤ�Խ�� As String) As Boolean
    'Ƥ�����
    Dim strSQL As String, strTest As String, strOwner As String
    Dim str��� As String, lngDeptID As Long
    Dim strContent As String, strLabel As String
    Dim strSelect As String, strSelectInput As String, strTextInput As String
    Dim datInput As Date
    Dim objBIll As Bill
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer, intResult  As Integer
    
    For Each objBIll In Item(strGroupKey).BillsItem(strGroupKey)
        If objBIll.ҽ��ID = Val(strGroupKey) Then
            strContent = objBIll.ҽ������
            Exit For
        End If
    Next
    
    On Error GoTo errHandle
    strSQL = "Select A.�������,A.Ƥ�Խ��,B.��������,Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ " & _
             "From ����ҽ����¼ A, ������ĿĿ¼ B " & _
             "Where A.������ĿID=B.ID And A.ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Ƥ������", Val(strGroupKey))
    '����
    For i = 0 To UBound(Split(Split(rsTemp!�걾��λ & "", ";")(0), ","))
        strSelect = strSelect & "," & Split(Split(rsTemp!�걾��λ & "", ";")(0), ",")(i) & "|0"
    Next
    '����
    For i = 0 To UBound(Split(Split(rsTemp!�걾��λ & "", ";")(1), ","))
        strSelect = strSelect & "," & Split(Split(rsTemp!�걾��λ & "", ";")(1), ",")(i) & "|0|2"
    Next
    strSelect = Mid(strSelect, 2)
    
    str��� = zlCommFun.ShowMsgBox("Ƥ�Խ��", _
                  strContent & "��^^����ݹ���������ѡ����Ӧ�İ�ť������", _
                  "ȷ��(&O),?ȡ��(&C)", frmMain, vbQuestion, _
                  "Ƥ��ʱ��", datInput, "yyyy-MM-dd HH:mm", _
                  "Ƥ�Խ��(&P):" & strSelect, strSelectInput, _
                  "������Ӧ(&F)", 100, strTextInput)
    
    If str��� = "" Then Exit Function
    If strSelectInput = "" Then Exit Function
    
    GetTestLabel rsTemp!�걾��λ, strSelectInput, strLabel, intResult
    lngDeptID = frmMain.cboDept.ItemData(frmMain.cboDept.ListIndex)
    
    strTest = "ZL_����ҽ����¼_Ƥ��(" & _
              Item(strGroupKey).ִ��ҽ��ID & _
              ",'" & strLabel & "'" & _
              "," & intResult & _
              ",''" & _
              ",to_date('" & datInput & "','YYYY-MM-DD HH24:MI:SS')" & _
              ",'" & strTextInput & "') "

    ExecuteTestFinish = ExecuteFinish(strGroupKey, lngDeptID, strTest, frmMain, objCardSquare)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function ExecCanle(ByVal strGroupKey As String, ByVal blnƤ����֤ As Boolean, ByVal lngExcDeptID As Long, ByVal objMain As Object) As Boolean
    'ȡ�����,Ƥ����ĿҲ������
    '
    '���ﲻ�ü���Ƿ����ȡ����
    'lngExcDeptID :ִ�п���ID 47323
    
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strUserName As String
    Dim cnNew As ADODB.Connection
    
    On Error GoTo errHandle
    If Item(strGroupKey).ִ��״̬ <> 1 Then
        MsgBox "��ִ����Ŀ��ǰ��������ִ��״̬������ȡ��ִ�С�", vbInformation, gstrSysName
        Exit Function
    End If
    '----- ������жϣ���Ҫ�ž�����Ҫzl_����ҽ��ִ��_Cancel�д���
    If Item(strGroupKey).�Ʒ�״̬ = -3 Then
        MsgBox "���˷ѣ�����ִ��ȡ��ִ��"
        Exit Function
    End If
    
    If Val(Item(strGroupKey).ִ�з���) = 3 Then
        'Ƥ����
        If blnƤ����֤ Then
            Set cnNew = New ADODB.Connection
            strUserName = zlDatabase.UserIdentify(objMain, "����Ƥ�Խ��ǰ�������������û������������������֤��", glngSys, 1263, "ȡ��ִ�����", cnNew)
            If strUserName = "" Then Exit Function
        End If
        strSQL = "ZL_����ҽ��ִ��_Cancel(" & Item(strGroupKey).ִ��ҽ��ID & "," & Item(strGroupKey).���ͺ� & "," & IIf(blnƤ����֤, 1, 0) & ",0," & lngExcDeptID & ")"
    Else
        strSQL = "ZL_����ҽ��ִ��_Cancel(" & Item(strGroupKey).ִ��ҽ��ID & "," & Item(strGroupKey).���ͺ� & ",Null,0," & lngExcDeptID & ")"
    End If
    
    'strSQL = "ZL_����ҽ��ִ��_Cancel(" & Item(strGroupKey).ִ��ҽ��ID & "," & Item(strGroupKey).���ͺ� & "," & IIf(blnƤ����֤, 1, 0) & ",0," & lngExcDeptID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.ExecCanle")
    strSQL = "Select ִ��״̬ From ����ҽ������ Where ���ͺ�+0=[1] and ҽ��ID=[2] And Rownum=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecutiveGroup.ExecCanle", Item(strGroupKey).���ͺ�, Item(strGroupKey).ִ��ҽ��ID)
    Do Until rsTmp.EOF
        Item(strGroupKey).ִ��״̬ = Val(rsTmp.Fields("ִ��״̬"))
        Item(strGroupKey).Ƥ�Խ�� = ""
        rsTmp.MoveNext
    Loop
    ExecCanle = True
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Public Sub FuncExecRefuse(strGroupKey As String)
''���ܣ��ܾ�ִ��
'    Dim lngҽ��ID As Long, lng���ͺ� As Long
'    Dim strSQL As String
'
'    '����ִ�л���ִ�в������ܾ�
'    If Item(strGroupKey).ִ��״̬ = 2 Then
'        MsgBox "��ִ����Ŀ��ǰ�Ѿ��ܾ�ִ�С�", vbInformation, gstrSysName
'        Exit Sub
'    End If
'    If Item(strGroupKey).ִ��״̬ = 3 Then
'        MsgBox "��ִ����Ŀ��ǰ����ִ�У����ܾܾ���", vbInformation, gstrSysName
'        Exit Sub
'    End If
'    If Item(strGroupKey).ִ��״̬ = 1 Then
'        MsgBox "��ִ����Ŀ��ǰ�Ѿ�ִ�У����ܾܾ���", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    If MsgBox("ȷʵҪ�ܾ�ִ�и���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    lngҽ��ID = Item(strGroupKey).ִ��ҽ��ID
'    lng���ͺ� = Item(strGroupKey).���ͺ�
'
'    strSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & lngҽ��ID & "," & lng���ͺ� & ")"
'
'    On Error GoTo errH
'
'    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.FuncExecRefuse")
'
'    On Error GoTo 0
'    Item(strGroupKey).ִ��״̬ = 2
'    Exit Sub
'errH:
'
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

'Public Sub FuncExecRestore(strGroupKey As String)
''���ܣ�ȡ���ܾ�ִ��
'    Dim lngҽ��ID As Long, lng���ͺ� As Long
'    Dim strSQL As String, i As Long
'
'    If Item(strGroupKey).ִ��״̬ <> 2 Then
'        MsgBox "��ִ����Ŀû�б��ܾ�ִ�С�", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    If MsgBox("ȷʵҪȡ���ܾ�ִ�и���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    lngҽ��ID = Item(strGroupKey).ִ��ҽ��ID
'    lng���ͺ� = Item(strGroupKey).���ͺ�
'
'    strSQL = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & lngҽ��ID & "," & lng���ͺ� & ")"
'
'    On Error GoTo errH
'
'    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.FuncExecRestore")
'
'    On Error GoTo 0
'    Item(strGroupKey).ִ��״̬ = 0
'    Exit Sub
'errH:
'
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Public Sub ExecStart(ByVal intUndo As Integer, strGroupKey As String, dateStart As Date, strOper As String)
    '��д��ʼʱ��
    Dim strSQL As String
    Dim lngID As Long, lngSendNo As Long, lngSeqNo As Long
    
    On Error GoTo hErr
    lngID = Item(strGroupKey).ִ��ҽ��ID
    lngSendNo = Item(strGroupKey).���ͺ�
    lngSeqNo = Me.��ˮ��
 
    If intUndo = 1 Then
        If Me.ִ��ʱ�� <> dateStart Then Me.ִ��ʱ�� = dateStart
        strSQL = "Zl_����ҽ��ִ��_Start(1," & lngSeqNo & "," & lngID & "," & lngSendNo & ",to_date('" & Format(dateStart, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),'" & strOper & "')"
    Else
        strSQL = "Zl_����ҽ��ִ��_Start(2," & lngSeqNo & "," & lngID & "," & lngSendNo & ")"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, "��д��ʼʱ��")
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub SaveDispenseUser(ByVal intUndo As Integer, dateStart As Date, strOper As String)
    '��д��ҩ�ˣ���ҩʱ��
    Dim strSQL As String
    Dim lngID As Long, lngSendNo As Long, lngSeqNo As Long
    Dim Y As Integer
    On Error GoTo hErr
    
    For Y = 1 To Me.Count
        If Item(Y).ִ�з��� = "1-��Һ" Then
            lngID = Item(Y).ִ��ҽ��ID
            lngSendNo = Item(Y).���ͺ�
            lngSeqNo = Me.��ˮ��
        
            If intUndo = 1 Then
                If Me.ִ��ʱ�� <> dateStart Then Me.ִ��ʱ�� = dateStart
                strSQL = "Zl_����ҽ��ִ��_Start(3," & lngSeqNo & "," & lngID & "," & lngSendNo & ",to_date('" & Format(dateStart, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),'" & strOper & "')"
            Else
                strSQL = "Zl_����ҽ��ִ��_Start(4," & lngSeqNo & "," & lngID & "," & lngSendNo & ")"
            End If
            Call zlDatabase.ExecuteProcedure(strSQL, "��д��ʼʱ��")
       End If
    Next
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub