VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Seatings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Seating"
Attribute VB_Ext_KEY = "Member0" ,"Seating"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection
Private mlng����ID As Long
Private mstr���� As String

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    ����ID = mlng����ID
End Property

Public Property Get ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    �������� = mstr����
End Property

Public Function Add(��� As Integer, ����ID As Long, ����� As String, ���� As String, �Ա� As String, ��� As String, _
                    ��� As Integer, ״̬ As Integer, �ּ� As Currency, �շ�ϸĿID As Long, �շ���Ŀ As String, _
                    ��ע As String, ���� As Long, ���� As String, ������ As String, �Һŵ� As String, Optional sKey As String) As Seating
    '�����¶���
    Dim objNewMember As Seating
    On Error GoTo errHandle
    If ��� = 0 Then
        Dim strSQL As String
        '���������ݣ�Ҫд������
        strSQL = "ZL_��λ״����¼_INSERT(" & mlng����ID & ",'" & ��� & "'," & Val(���) & "," & Val(״̬) & "," & IIf(Val(�շ�ϸĿID) = 0, "Null", Val(�շ�ϸĿID)) & ",'" & ��ע & "'," & ���� & ",'" & ���� & "','" & ������ & "')"
        zldatabase.ExecuteProcedure strSQL, "seatings.add"
        Call getSeatings(mlng����ID)
    Else
        Set objNewMember = New Seating
        '���ô��뷽��������
        objNewMember.��� = ���
        objNewMember.����ID = ����ID
        objNewMember.����� = �����
        objNewMember.���� = ����
        objNewMember.�Ա� = �Ա�
        objNewMember.��� = ���
        objNewMember.��� = ���
        objNewMember.״̬ = ״̬
        objNewMember.�ּ� = �ּ�
        objNewMember.�շ�ϸĿID = �շ�ϸĿID
        objNewMember.�շ���Ŀ = IIf(IsNull(�շ���Ŀ), "", �շ���Ŀ)
        objNewMember.��ע = IIf(IsNull(��ע), "", ��ע)
        objNewMember.���� = ����
        objNewMember.��������� = ������
        objNewMember.���� = ����
        objNewMember.�Һŵ� = �Һŵ�
        If �Һŵ� = "" Then
            objNewMember.PatientKey = "1_" & ����ID & "_" & GetPatiPageID(mlng����ID, ����ID)
        Else
            objNewMember.PatientKey = "0_" & ����ID & "_" & �Һŵ�
        End If
        objNewMember.Key = sKey
        
        If Len(sKey) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sKey
        End If
        '�����Ѵ����Ķ���
        Set Add = objNewMember
        Set objNewMember = Nothing
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetNextNo(ByVal ��� As Integer) As String
    
    Dim strSQL As String, strMaxNO As String, strǰ׺ As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select Max(���) As maxno From ��λ״����¼ Where  ����ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.GetNextNO", mlng����ID)
'    Select Case ���
'        Case 1
'            strǰ׺ = "��"
'        Case 2
'            strǰ׺ = "��"
'        Case 3
'            strǰ׺ = "��"
'        Case Else
'            strǰ׺ = "��"
'    End Select
    strǰ׺ = ""
    If rsTmp.RecordCount > 0 Then
        strMaxNO = zlCommFun.NVL(rsTmp.Fields("maxno"))
        If strMaxNO = "" Then
            GetNextNo = strǰ׺ & "001"
        Else
            GetNextNo = zlCommFun.IncStr(strMaxNO)
        End If
    Else
        GetNextNo = strǰ׺ & "001"
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ClearSeatings()
    '�������������
    Dim lngIndex As Long
    For lngIndex = 1 To Me.Count
        Remove 1
    Next
End Function

Private Function GetPatiPageID(ByVal lngDeptID As Long, ByVal lngPatiID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Max(��ҳid) ��ҳid From �ŶӼ�¼ Where ����id = [1] And ����id = [2]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "��ȡ���˵���ҳID", lngDeptID, lngPatiID)
    If rsTemp.EOF = False Then
        GetPatiPageID = Val(zlCommFun.NVL(rsTemp!��ҳid))
    End If
    rsTemp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Function SetSeating(ByVal lng����ID As Long, ByVal strRegNo As String, ByVal SeatingKey As String) As Boolean
    '������λ
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strBillNO As String
    Dim blnTrans As Boolean
    Dim lngPageID As Long
    
    On Error GoTo errHandle
    
    If lng����ID = 0 Or SeatingKey = "" Then
        MsgBox "��λ���ŵ���δ�������!", vbInformation, gstrSysName
        Exit Function
    End If
    '-- ������λ��Ϣ
    Call Me.getSeatings(Me.����ID)
    SeatingKey = Replace(SeatingKey, "-", "_")
    If Me.Item(SeatingKey).����ID = 0 And Me.Item(SeatingKey).״̬ = 0 Then
        Me.Item(SeatingKey).����ID = lng����ID
        Me.Item(SeatingKey).״̬ = 1
        Me.Item(SeatingKey).�Һŵ� = strRegNo
        
        If (strRegNo = "" Or strRegNo Like "*��*") Then
            '���ŶӼ�¼ȡ�������۲��˵���ҳID
            lngPageID = GetPatiPageID(Me.����ID, lng����ID)
            If lngPageID > 0 Then
                Me.Item(SeatingKey).PatientKey = "1_" & lng����ID & "_" & lngPageID
            Else
                Me.Item(SeatingKey).PatientKey = "0_" & lng����ID & "_"
            End If
        Else
            Me.Item(SeatingKey).PatientKey = "0_" & lng����ID & "_" & strRegNo
        End If
        
        strSQL = "Select ����ID From ��λ״����¼ Where ����ID=[1] and ���=[2] and ���=[3]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "SetSeating", Me.����ID, Me.Item(SeatingKey).���, Me.Item(SeatingKey).���)
        If Not rsTmp.EOF Then
            If zlCommFun.NVL(rsTmp.Fields("����ID"), 0) = 0 Then
                
                If Me.Item(SeatingKey).�շ�ϸĿID <> 0 Then
                    '�������۵�
                    strBillNO = zldatabase.NextNo(13)
                    With Me.Item(SeatingKey)
                        Dim str���ʽ As String, str���� As String, str�ѱ� As String, lng����ID As Long, lng����ID As Long
                        Dim rsBill As ADODB.Recordset
                        Dim str���� As String, str�Ա� As String
                        strSQL = "Select B.����, A.ҽ�Ƹ��ʽ, A.����, A.�ѱ�, A.��ǰ����id, A.��ǰ����id, A.����, A.�Ա�" & vbNewLine & _
                                "From ҽ�Ƹ��ʽ B, ������Ϣ A" & vbNewLine & _
                                "Where A.ҽ�Ƹ��ʽ = B.���� And A.����ID=[1]"

                        Set rsBill = zldatabase.OpenSQLRecord(strSQL, "SetSeating", lng����ID)
                        If Not rsBill.EOF Then
                            str���ʽ = rsBill.Fields("����")
                            str���� = rsBill.Fields("����")
                            str�Ա� = rsBill.Fields("�Ա�")
                            str���� = zlCommFun.NVL(rsBill.Fields("����"))
                            str�ѱ� = zlCommFun.NVL(rsBill.Fields("�ѱ�"))
                            lng����ID = zlCommFun.NVL(rsBill.Fields("��ǰ����ID"), 0)
                            lng����ID = zlCommFun.NVL(rsBill.Fields("��ǰ����ID"), 0)
                        End If
                        
                        Dim str�շ���� As String, str���㵥λ As String, lng������ĿID As Long, str�վݷ�Ŀ As String, lng�۸񸸺� As Long, lng�Ӱ��־ As Long
                        Dim cur���� As Currency, cur��� As Currency
                        Dim int��� As Integer
                        strSQL = "Select a.���, a.���㵥λ, a.�Ӱ�Ӽ�, b.�ּ�, b.������Ŀid, c.�վݷ�Ŀ" & vbNewLine & _
                                "From ������Ŀ c," & vbNewLine & _
                                "        (Select ������Ŀid, �ּ�, �շ�ϸĿid" & vbNewLine & _
                                "            From �շѼ�Ŀ a" & vbNewLine & _
                                "            Where ��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'yyyy-MM-dd')) b, �շ���ĿĿ¼ a" & vbNewLine & _
                                "Where b.������Ŀid = c.Id And a.Id = b.�շ�ϸĿid And a.Id = [1]"
                        Set rsBill = zldatabase.OpenSQLRecord(strSQL, "SetSeating", .�շ�ϸĿID)
                        
                        gcnOracle.BeginTrans
                        blnTrans = True
                        Do Until rsBill.EOF
                            int��� = int��� + 1
                            If int��� = 1 Then
                                lng�۸񸸺� = rsBill.Fields("������ĿID")
                            End If
                            str�շ���� = rsBill.Fields("���")
                            str���㵥λ = rsBill.Fields("���㵥λ")
                            cur���� = rsBill.Fields("�ּ�")
                            lng������ĿID = rsBill.Fields("������Ŀid")
                            str�վݷ�Ŀ = rsBill.Fields("�վݷ�Ŀ")
                            
                            strSQL = "Zl_���ﻮ�ۼ�¼_Insert('" & strBillNO & "',1," & lng����ID & ",Null," & .����� & ",'" & str���ʽ & "','" & _
                                     str���� & "','" & str�Ա� & "','" & str���� & "','" & _
                                     str�ѱ� & "'," & lng�Ӱ��־ & "," & IIf(lng����ID = 0, "Null", lng����ID) & "," & Me.����ID & ",'" & _
                                     UserInfo.���� & "',Null," & .�շ�ϸĿID & ",'" & str�շ���� & "','" & str���㵥λ & "',Null," & _
                                     "Null,1,Null," & Me.����ID & "," & IIf(int��� = 1, "Null", lng�۸񸸺�) & "," & lng������ĿID & ",'" & str�վݷ�Ŀ & "'," & cur���� & "," & _
                                     cur���� & "," & cur���� & ",sysdate,sysdate,Null,'" & UserInfo.���� & "')"
                            Call zldatabase.ExecuteProcedure(strSQL, "SetSeating")
                            cur��� = cur��� + cur����
                            rsBill.MoveNext
                        Loop
                        gcnOracle.CommitTrans
                        blnTrans = False
                    End With
                    MsgBox "������[" & str���� & "]���շ���ȥ�ɷѡ�" & vbNewLine & "���ݺţ�" & strBillNO & "����" & Format(cur���, "0.00"), vbQuestion, gstrSysName
                End If
                strSQL = "Zl_��λ״����¼_Setseating(" & Me.����ID & "," & Me.Item(SeatingKey).��� & ",'" & Me.Item(SeatingKey).��� & "'," & lng����ID & ",'" & strBillNO & "')"
                Call zldatabase.ExecuteProcedure(strSQL, "Seatings.SetSeating")
                
                SetSeating = True
                Call SaveOperLog(Me.����ID, Me.Item(SeatingKey).PatientKey, SEAT, "������λ" & Me.Item(SeatingKey).��� & IIf(strBillNO <> "", "�������۵�" & strBillNO & "��" & Format(cur���, "0.00"), ""))
            ElseIf zlCommFun.NVL(rsTmp.Fields("����ID"), 0) = lng����ID Then
                SetSeating = True
            Else
                MsgBox "��λ�Ѿ��������û���ռ�ã�����ѡ��λ��", vbInformation, gstrSysName
                Exit Function

            End If
        Else
            MsgBox "��λ�Ѿ������ڣ������°��ţ�", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        MsgBox "��λ�Ѿ���������ռ�ã�����ѡ��λ��", vbInformation, gstrSysName
        Exit Function
    End If
    Exit Function
errHandle:
    
    Call ErrCenter
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Sub getSeatings(ByVal deptID As Long)
    'ȡ��λ��Ϣ
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsPati As ADODB.Recordset
    Dim str�Һŵ� As String
    
    On Error GoTo errHandle
    Call ClearSeatings
    
    strSQL = "Select Rownum As ���, Nvl(A.����id, 0) As ����id, Nvl(B.�����, 0) �����, Nvl(B.����, '') As ����, " & vbNewLine & _
            "       Nvl(B.�Ա�, '') �Ա�, B.���� ,A.���, A.���, Nvl(A.�շ�ϸĿid, 0) �շ�ϸĿid, A.״̬, A.��ע, Nvl(A.����,0) as ����, Nvl(C.�ּ�, 0) As �ּ�, a.����, a.���������," & vbNewLine & _
            "       Replace('[' || D.���� || ']' || D.����, '[]', '') As �շ���Ŀ" & vbNewLine & _
            "From �շ���ĿĿ¼ D, �շѼ�Ŀ C, ������Ϣ B, ��λ״����¼ A " & vbNewLine & _
            "Where A.�շ�ϸĿid = D.ID(+) And (C.��ֹ���� Is Null Or C.��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "      A.�շ�ϸĿid = C.�շ�ϸĿid(+) And A.����id = B.����id(+) And A.����id = [1] " & _
            GetPriceGradeSQL(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�, "D", "C", "2", "3", "4") & _
            "Order by A.���,A.���"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Seatings.getSeatings", deptID, gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�)
    With rsTmp
    Do Until .EOF
        str�Һŵ� = ""
        If Val(zlCommFun.NVL(!����ID)) > 0 Then
            strSQL = "Select NO From ���˹Һż�¼ " & _
                     "Where (����id, ID) In (Select ����id, Max(ID) ID From ���˹Һż�¼ " & _
                     "                       Where ��¼���� = 1 And ��¼״̬ = 1 And ����id = [1] And ִ�в���id = [2] " & _
                     "                       Group By ����id)"
            Set rsPati = zldatabase.OpenSQLRecord(strSQL, "��ȡ���˹Һŵ�", !����ID, deptID)
            If rsPati.EOF = False Then
                str�Һŵ� = zlCommFun.NVL(rsPati!NO)
            End If
            rsPati.Close
        End If
    
        Add .Fields("���"), .Fields("����ID"), .Fields("�����"), _
            Trim$("" & .Fields("����") & " " & .Fields("����")), "" & .Fields("�Ա�"), .Fields("���"), _
            "" & .Fields("���"), .Fields("״̬"), .Fields("�ּ�"), .Fields("�շ�ϸĿID"), _
            IIf(IsNull(.Fields("�շ���Ŀ")), "", .Fields("�շ���Ŀ")), _
             IIf(IsNull(.Fields("��ע")), "", .Fields("��ע")), _
             .Fields("����"), "" & .Fields("����"), "" & .Fields("���������"), _
             str�Һŵ�, "" & .Fields("���") & "_" & .Fields("���")
        .MoveNext
    Loop
    End With
    'ȡ��������
    mlng����ID = deptID
    strSQL = "Select ����,���� From ���ű� Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.getSeatings", deptID)
    If rsTmp.RecordCount > 0 Then
        mstr���� = "[" & rsTmp.Fields("����") & "] " & rsTmp.Fields("����")
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As Seating
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
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

Public Function Delete(ByVal strKey As String) As Boolean
    'ɾ����λ
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim objSeating As Seating
    On Error GoTo errHandle
        

    
    Set objSeating = Item(strKey)
    
    strSQL = "Select A.����ID,B.����,A.״̬  From ��λ״����¼ A,������Ϣ B Where  A.����ID=B.����ID(+) And A.����ID=[1] And A.���=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.Delete", mlng����ID, CStr(objSeating.���))
    If rsTmp.EOF Then
        Call Remove(strKey)
        Delete = True
        Exit Function
    Else
        If rsTmp!����ID <> 0 Or rsTmp!״̬ = 1 Then
            MsgBox objSeating.��� & "��" & rsTmp!���� & "ռ�ã�����ɾ��!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    
    If strKey <> "" And objSeating.����ID = 0 And objSeating.״̬ <> 1 Then
        If MsgBox("�Ƿ�ɾ��[" & objSeating.��� & "]��λ��", vbExclamation + vbYesNo, "����") = vbYes Then
            strSQL = "Zl_��λ״����¼_Delete(" & mlng����ID & ",'" & objSeating.��� & "')"
            Call zldatabase.ExecuteProcedure(strSQL, "seatings.delete")
            Call Remove(strKey)
            Delete = True
        End If
    Else
        If strKey = "" Then
            MsgBox "���Ϊ�յ���λ������ɾ��!", vbInformation, gstrSysName
            Exit Function
        End If
        If objSeating.����ID <> 0 Or objSeating.״̬ = 1 Then
            MsgBox objSeating.��� & "��" & objSeating.���� & "ռ�ã�����ɾ��!", vbInformation, gstrSysName
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Clear(ByVal strKey As String, Optional blnInfo As Boolean = True) As Boolean
    '�����λ
    'blnInfo:�Ƿ�����
    Dim strSQL As String, rsTmp As New ADODB.Recordset, strNo As String
    Dim blnTrans As Boolean, strSwap As String
    On Error GoTo errHandle
    If strKey <> "" And Item(strKey).����ID <> 0 And Item(strKey).״̬ = 1 Then
        
        If blnInfo Then
            If MsgBox("�Ƿ������[" & Item(strKey).���� & "]ռ�õ�[" & Item(strKey).��� & "]��λ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            strSwap = "���"
        Else
            '����λʱ,����Ƿ��л��۵�,���δ�շ������,���շ�����ʾȥ�˷�
            strSQL = "Select ����ID,NO From ��λ״����¼ Where ����ID=[1] and ���=[2] and ���=[3]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.Clear", Me.����ID, Me.Item(strKey).���, Me.Item(strKey).���)
            If Not rsTmp.EOF Then
                strNo = zlCommFun.NVL(rsTmp.Fields("NO"))
                If strNo <> "" Then
                    strSQL = "Select nvl(ʵ�ս��,0) as ��� From ������ü�¼ Where mod(��¼����,10)=1 And ��¼״̬=1 and ����״̬<>1 And NO=[1] "
                    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.Clear", strNo)
                    If Not rsTmp.EOF Then
                        MsgBox "������[" & Item(strKey).���� & "]���շ����˷�" & vbNewLine & "���ݺţ�" & strNo & "����" & Format(Item(strKey).�ּ�, "0.00") & "��", vbInformation, gstrSysName
                    Else
                        strSQL = "Select NO,��� From ������ü�¼ Where mod(��¼����,10)=1 And ��¼״̬=0 And NO=[1] "
                        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.Clear", strNo)
                        gcnOracle.BeginTrans
                        blnTrans = True
                        Do Until rsTmp.EOF
                            strSQL = "zl_���ﻮ�ۼ�¼_Delete('" & strNo & "'," & rsTmp.Fields("���") & ")"
                            Call zldatabase.ExecuteProcedure(strSQL, "Seatings.Clear")
                            rsTmp.MoveNext
                        Loop
                        gcnOracle.CommitTrans
                        blnTrans = False
                    End If
                End If
            End If
            strSwap = "����λ�����"
        End If
        
        strSQL = "ZL_��λ״����¼_Clear(" & mlng����ID & ",'" & Item(strKey).��� & "')"
        Call zldatabase.ExecuteProcedure(strSQL, "Seatings.Clear")
        
        Clear = True
        Call SaveOperLog(Me.����ID, Item(strKey).PatientKey, SEAT, strSwap & IIf(Item(strKey).���� <> "", "��[" & Item(strKey).���� & "]ռ�õ�", "") & "[" & Item(strKey).��� & "]��λ")
    
        Item(strKey).����ID = 0
        Item(strKey).���� = ""
        Item(strKey).�Ա� = ""
        Item(strKey).״̬ = 0
        Item(strKey).����� = 0
        Item(strKey).��� = ""
        Item(strKey).��ʼʱ�� = ""
        Item(strKey).�Һŵ� = ""
        Item(strKey).PatientKey = ""
        
    End If
    Exit Function
errHandle:
    
    Call ErrCenter
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function SwapSeating(strSourceKey, strObjKey) As Boolean
    '
    Dim lng����ID As Long, strRegNo As String
    On Error GoTo errHandle
    
    If strSourceKey <> "" And strObjKey <> "" Then
        lng����ID = Item(strSourceKey).����ID
        strRegNo = Item(strSourceKey).�Һŵ�
        
        If Not Clear(strSourceKey, False) Then Exit Function
        If Not SetSeating(lng����ID, strRegNo, strObjKey) Then Exit Function
        
        SwapSeating = True
    End If
    
    Exit Function
errHandle:

    Call ErrCenter
End Function

