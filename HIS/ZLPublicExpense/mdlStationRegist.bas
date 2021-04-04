Attribute VB_Name = "mdlStationRegist"
Option Explicit
'ҽ��վ�ҺŴ������

Public Function InitRegist(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnOracle As ADODB.Connection, ByVal strDbUser As String, _
                Optional objRegist As clsRegist, _
                Optional objExseSvr As clsExpenceSvr, _
                Optional objService As clsService) As Boolean
    '��ʼ���Һ�
    Dim strDept As String
    On Error GoTo errH:
    Set objRegist = New clsRegist
    If objRegist.zlInitCommon(lngSys, cnOracle, strDbUser) = False Then Exit Function
    
    Set objExseSvr = New clsExpenceSvr
    If objExseSvr.zlInitCommon(lngSys, lngModul, cnOracle, strDbUser) = False Then Exit Function
    
    Set objService = New zlPublicExpense.clsService
    If objService.zlInitCommon(lngSys, lngModul, cnOracle, strDbUser) = False Then Exit Function
    
    InitRegist = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function ReadLastAppoint(ByVal lng����ID As Long, ByVal lng�ƻ�ID As Long, _
                                ByVal lng��¼ID As Long, ByVal datDay As Date, _
                                ByVal bln��ʱ�� As Boolean, strLastAppoint As String) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    If lng��¼ID <> 0 Then
        If bln��ʱ�� Then
            strSQL = "Select ����, ����ʱ��, ��ʼʱ��, ��ֹʱ��" & vbNewLine & _
                        "From (Select b.��� as ����, a.����ʱ��, a.�Ǽ�ʱ��, b.��ʼʱ��, b.��ֹʱ��" & vbNewLine & _
                        "       From ���˹Һż�¼ a, �ٴ�������ſ��� b" & vbNewLine & _
                        "       Where a.�����¼id = [1] And a.��¼״̬ = 1 And a.�����¼id = b.��¼id And (a.���� = b.��� or a.���� = zl_To_number(b.��ע)) And a.����ʱ�� Between [2] And [3]" & vbNewLine & _
                        "       Order By �Ǽ�ʱ�� Desc)" & vbNewLine & _
                        "Where Rownum < 2"
        Else
            strSQL = "Select ����, ����ʱ��" & vbNewLine & _
                        "From (Select ����, ����ʱ��, �Ǽ�ʱ��" & vbNewLine & _
                        "       From ���˹Һż�¼" & vbNewLine & _
                        "       Where �����¼id = [1] And ��¼״̬ = 1 And ����ʱ�� Between [2] And [3]" & vbNewLine & _
                        "       Order By �Ǽ�ʱ�� Desc)" & vbNewLine & _
                        "Where Rownum < 2"
        End If
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ReadLastAppoint", lng��¼ID, CDate(Format(datDay, "yyyy-MM-dd")), CDate(Format(datDay, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
    Else
        If bln��ʱ�� Then
            strSQL = "Select ����, ����ʱ��, ��ʼʱ��, ��ֹʱ��" & vbNewLine & _
                        "From (Select a.����, a.����ʱ��, a.�Ǽ�ʱ��, c.��ʼʱ��, c.����ʱ�� As ��ֹʱ��" & vbNewLine & _
                        "       From ���˹Һż�¼ a, �ҺŰ��� b, �ҺŰ���ʱ�� c" & vbNewLine & _
                        "       Where b.Id = [1] And a.�ű� = b.���� And a.��¼״̬ = 1 And b.Id = c.����id And a.���� = c.���(+) And" & vbNewLine & _
                        "             c.����(+) = Decode(To_Char([2], 'D')," & vbNewLine & _
                        "                              '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) And" & vbNewLine & _
                        "             ����ʱ�� Between [2] And [3]" & vbNewLine & _
                        "       Order By �Ǽ�ʱ�� Desc)" & vbNewLine & _
                        "Where Rownum < 2"
        Else
            strSQL = "Select ����, ����ʱ��" & vbNewLine & _
                        "From (Select a.����, a.����ʱ��, a.�Ǽ�ʱ��" & vbNewLine & _
                        "       From ���˹Һż�¼ a, �ҺŰ��� b" & vbNewLine & _
                        "       Where b.Id = [1] And a.�ű� = b.���� And a.��¼״̬ = 1 And ����ʱ�� Between [2] And" & vbNewLine & _
                        "             [3]" & vbNewLine & _
                        "       Order By �Ǽ�ʱ�� Desc)" & vbNewLine & _
                        "Where Rownum < 2"
        End If
        If lng�ƻ�ID <> 0 Then
            strSQL = Replace(strSQL, "�ҺŰ���ʱ��", "�Һżƻ�ʱ��")
            strSQL = Replace(strSQL, "�ҺŰ���", "�ҺŰ��żƻ�")
            strSQL = Replace(strSQL, "c.����id", "c.�ƻ�id")
        End If
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ReadLastAppoint", IIf(lng�ƻ�ID > 0, lng�ƻ�ID, lng����ID), CDate(Format(datDay, "yyyy-MM-dd")), CDate(Format(datDay, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
    End If
    
    If Not rsTemp.EOF Then
        If bln��ʱ�� Then
            strLastAppoint = Nvl(rsTemp!����) & "(" & Format(Nvl(rsTemp!��ʼʱ��), "HH:MM") & "-" & Format(Nvl(rsTemp!��ֹʱ��), "HH:MM") & ")"
        Else
            strLastAppoint = Nvl(rsTemp!����) & "(" & Format(Nvl(rsTemp!����ʱ��), "HH:MM") & ")"
        End If
    End If
    ReadLastAppoint = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetSNState(ByVal bytRegistMode As Byte, lng��¼ID As Long, Optional str�ű� As String, Optional datThis As Date, Optional lngSN As Long) As ADODB.Recordset
    If bytRegistMode = 0 Then
        Set GetSNState = GetSNState_Visit(lng��¼ID)
    Else
        Set GetSNState = GetSNState_Normal(str�ű�, datThis, lngSN)
    End If
End Function

Private Function GetSNState_Normal(str�ű� As String, datThis As Date, Optional lngSN As Long) As ADODB.Recordset
    Dim strSQL           As String
    Dim datStart         As Date
    Dim datEnd           As Date
    On Error GoTo errH
    datStart = CDate(Format(datThis, "yyyy-MM-dd"))
    datEnd = DateAdd("s", -1, DateAdd("d", 1, datStart))
    strSQL = "    " & vbNewLine & " Select ���,״̬,����Ա����,Nvl(ԤԼ,0) as ԤԼ,TO_Char(����,'hh24:mi:ss') as ����  "
    strSQL = strSQL & vbNewLine & " From �Һ����״̬ "
    strSQL = strSQL & vbNewLine & " Where ����=[1]"
    strSQL = strSQL & vbNewLine & IIf(datThis = CDate(0), " And ���� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 ", " And ���� Between  [2] And [3]")
    strSQL = strSQL & vbNewLine & IIf(lngSN > 0, " And ���=[4]", "")
    Set GetSNState_Normal = gobjDatabase.OpenSQLRecord(strSQL, "GetSNState_Normal", str�ű�, datStart, datEnd, lngSN)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetSNState_Visit(lng��¼ID As Long) As ADODB.Recordset
    Dim strSQL           As String
    On Error GoTo errH

    strSQL = "    " & vbNewLine & " Select A.���,A.�Һ�״̬,A.����Ա����,Decode(A.�Һ�״̬,2,1,0) as ԤԼ,To_Char(B.��������,'hh24:mi:ss') as ����  "
    strSQL = strSQL & vbNewLine & " From �ٴ�������ſ��� A, �ٴ������¼ B "
    strSQL = strSQL & vbNewLine & " Where B.ID=[1] And B.ID=A.��¼ID"
    Set GetSNState_Visit = gobjDatabase.OpenSQLRecord(strSQL, "GetSNState_Visit", lng��¼ID)

    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAllҽ��(rsDoctor As ADODB.Recordset) As Boolean
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.����, Upper(a.����) As ����,b.����id,a.���" & _
            " From ��Ա�� a, ������Ա b, ��Ա����˵�� c" & _
            " Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = 'ҽ��' And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order By a.���� Desc"
    Set rsDoctor = gobjDatabase.OpenSQLRecord(strSQL, "GetAllҽ��")
    GetAllҽ�� = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function zlGetActiveViewSql(ByVal bytRegistMode As Byte) As String
    Dim strSQL As String
    
    If bytRegistMode = 0 Then
        strSQL = _
        "       Select   Havedata, ����id" & vbNewLine & _
        "       From (" & vbNewLine & _
        "               Select 1 As Havedata, b.Id As ����id " & vbNewLine & _
        "               From �ҺŰ���ʱ�� A, �ҺŰ��� B" & vbNewLine & _
        "               Where B.����=[1] And A.����id = b.ID " & _
        "                And   Decode(To_Char([2], 'D'), '1', '����', '2'," & _
        "                   '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6','����', '7', '����', Null) =a.���� " & vbNewLine & _
        "                       And Not Exists" & vbNewLine & _
        "                     (Select 1 From �ҺŰ��żƻ� C " & vbNewLine & _
        "                         Where c.����id = b.Id And c.���ʱ�� Is Not Null And [2] Between " & _
        "                               Nvl(c.��Чʱ��, [2]) And" & _
        "                          c.ʧЧʱ��)" & vbNewLine & _
        "               Union All " & vbNewLine & _
        "               Select 1 As Havedata, c.Id As ����id" & vbNewLine & _
        "               From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C,(" & vbNewLine & _
        "                   SELECT MAX(a.��Чʱ�� ) ��Ч FROM �ҺŰ��żƻ� a,�ҺŰ��� B  WHERE a.����Id=b.ID AND b.����=[1] AND a.���ʱ�� IS NOT NULL" & vbNewLine & _
        "             And [2] Between nvl(a.��Чʱ��,to_date('1900-01-01','yyyy-mm-dd')) And a.ʧЧʱ��" & vbNewLine & _
        "           ) D  " & vbNewLine & _
        "               Where  C.����=[1] And c.Id = b.����id And b.Id = a.�ƻ�id And b.��Чʱ��=d.��Ч And b.���ʱ�� Is Not Null" & _
        "                    And   Decode(To_Char([2], 'D'), '1', '����', '2'," & _
        "                   '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6','����', '7', '����', Null) =a.���� " & vbNewLine & _
        "                       And [2] Between Nvl(b.��Чʱ��,[2]) And b.ʧЧʱ��) B"
    Else
        strSQL = "Select 1 From �ٴ������¼ Where ID=[1] And Nvl(�Ƿ��ʱ��,0)=1 "
    End If
    zlGetActiveViewSql = strSQL
End Function

Public Function zlGetTimeSnSql(ByVal bytRegistMode As Byte) As String
    Dim strSQL As String
    
    If bytRegistMode = 0 Then
        strSQL = "" & _
        " Select Distinct a.��� As ID, A.���,To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
        " From �ҺŰ���ʱ�� A, �ҺŰ��� B, �ҺŰ������� C" & vbNewLine & _
        " Where a.����id = b.Id And b.���� = [1] And" & vbNewLine & _
        " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
        "      Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',Null) = a.����(+)  " & _
        "      And b.Id = c.����Id(+) And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',Null) = c.������Ŀ(+)" & _
        "      And Not Exists (Select Count(1) From �Һ����״̬ Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like Rpad(a.���, Length(a.���)+ length(Nvl(c.�޺���,0)), '_')) Having Count(1) - a.�������� >= 0) " & _
        "      And Not Exists (Select 1 From �ҺŰ��żƻ� E Where e.����id = b.Id And e.���ʱ�� Is Not Null And [2] Between Nvl(e.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And e.ʧЧʱ��)" & _
        "      And Not Exists (Select 1 From ������λ���ſ��� Where ����id = b.Id And ��� = a.��� And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�','4', '����', '5', '����', '6', '����', '7', '����', Null) = ������Ŀ)"
        
        strSQL = strSQL & " Union " & _
        "Select Distinct a.��� As ID,A.���,To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
        "From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C, �Һżƻ����� E," & vbNewLine & _
        "     (Select Max(a.��Чʱ��) ��Ч" & vbNewLine & _
        "       From �ҺŰ��żƻ� A, �ҺŰ��� B" & vbNewLine & _
        "       Where a.����id = b.Id And b.���� = [1] And a.���ʱ�� Is Not Null And" & vbNewLine & _
        "             [2] Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
        "             a.ʧЧʱ��) D" & vbNewLine & _
        "Where a.�ƻ�id = b.Id And b.����id = c.Id And c.���� = [1] And b.��Чʱ�� = d.��Ч And b.���ʱ�� Is Not Null And" & vbNewLine & _
        " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
        "      [2] Between Nvl(b.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And b.ʧЧʱ��" & vbNewLine & _
        "      And b.Id = e.�ƻ�Id(+) And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',Null) = e.������Ŀ(+)" & vbNewLine & _
        "      And Not Exists" & vbNewLine & _
        " (Select Count(1)" & vbNewLine & _
        "       From �Һ����״̬" & vbNewLine & _
        "       Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like Rpad(a.���, Length(a.���)+ length(Nvl(e.�޺���,0)), '_')) Having" & vbNewLine & _
        "        Count(1) - a.�������� >= 0) And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5'," & vbNewLine & _
        "                                           '����', '6', '����', '7', '����', Null) = a.����(+) And Not Exists" & vbNewLine & _
        " (Select 1 From ������λ�ƻ����� Where �ƻ�id = b.Id And ��� = a.��� And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) = ������Ŀ)" & vbNewLine & _
        "Order By ��ʼʱ��"
    Else
        strSQL = "" & _
        " Select Rownum As Id, ���, To_Char(��ʼʱ��, 'hh24') || ':00' As ʱ���, To_Char(��ʼʱ��, 'hh24:mi') As ��ʼʱ��," & vbNewLine & _
        "       To_Char(��ֹʱ��, 'hh24:mi') As ����ʱ��, ��ʼʱ�� As ��ϸ��ʼʱ��, ��ֹʱ�� As ��ϸ����ʱ�� " & vbNewLine & _
        " From �ٴ�������ſ��� A" & vbNewLine & _
        " Where ��¼id = [1] And Nvl(�Һ�״̬,0) = 0 And Nvl(�Ƿ�ԤԼ,0)=1 And Trunc(��ʼʱ��) = [2] And Not Exists " & vbNewLine & _
        "(Select 1 From �ٴ�����Һſ��Ƽ�¼ B Where b.��¼id = a.��¼id And b.���Ʒ�ʽ = 3 And a.��� = b.���)" & vbNewLine & _
        "Order By ��ϸ��ʼʱ��"
    End If
    zlGetTimeSnSql = strSQL
End Function

Public Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset, _
                                ByVal rsExpenses As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL  As String
    
    Err = 0: On Error GoTo errHand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic

        rsItems.Filter = 0
        If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
        Do While Not rsItems.EOF
            rsIncomes.Filter = "��ĿID=" & rsItems!��ĿID
            rsMoney.Filter = "�շ����='" & Nvl(rsItems!���, "��") & "'"
            If rsMoney.EOF Then
                .AddNew
            End If
            !�շ���� = Nvl(rsItems!���, "��")
            Do While Not rsIncomes.EOF
                !��� = Val(Nvl(!���)) + Val(Nvl(rsIncomes!ʵ��))
                rsIncomes.MoveNext
            Loop
            .Update
            rsItems.MoveNext
        Loop
        
        If Not rsExpenses Is Nothing Then
            If rsExpenses.RecordCount > 0 Then rsExpenses.MoveFirst
            Do While Not rsExpenses.EOF
                rsMoney.Filter = "�շ����='" & Nvl(rsExpenses!���, "��") & "'"
                If rsMoney.EOF Then
                    .AddNew
                End If
                !�շ���� = Nvl(rsExpenses!���, "��")
                !��� = Val(Nvl(!���)) + Val(Nvl(rsExpenses!ʵ��))
                .Update
                rsExpenses.MoveNext
            Loop
        End If
    End With
    rsMoney.Filter = 0
    zlGetClassMoney = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CreatePublicPatient(ByVal frmMain As Form, objPubPatient As clsInterFacePatient) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����zlPublicPatient����
    '����:�����ɹ�,����True,���򷵻�False
    '����:Ƚ����
    '����:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set objPubPatient = New clsInterFacePatient
    If objPubPatient.Init(frmMain, glngSys, glngModul, gcnOracle, gstrDBUser) = False Then Exit Function
    CreatePublicPatient = True
End Function

Public Function zlInsure_Check(ByVal str���ս��� As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:str���ս���-���ս���
    '       strAdvance-ҽ�����صĽ���
    '����:
    '����:��Ҫ�϶�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varTemp As Variant
    Dim varData As Variant
    Dim varData1 As Variant
    Dim varTemp1 As Variant
    
    On Error GoTo errHandle
    If Not (strAdvance <> "" And str���ս��� <> strAdvance) Then Exit Function
    '��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    blnMedicareCheck = True
    varData = Split(str���ս���, "|")
    varData1 = Split(strAdvance, "|")
    If UBound(varData) = UBound(varData1) Then
    
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, ",")
            
            For j = 0 To UBound(varData1)
                strTmp = varData1(j)
                varTemp1 = Split(strTmp, ",")
                
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next

    End If
    zlInsure_Check = blnMedicareCheck
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlMakeBillRecord(ByVal lng����ID As Long, ByVal str�ѱ� As String, ByVal bln���� As Boolean, _
            ByVal blnHav������ As Boolean, _
            ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset, _
            ByVal rsExpense As ADODB.Recordset, ByVal datDate As Date, _
            ByRef rsDetail As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݹҺ��շ���Ŀ������ҽ����¼����ϸ��Ϣ(���ۼ۵�λ)
    '���: datDate:����ʱ��,
    '����:rsDetail-���ص�ҽ�������ϸ����
    '      �������(1--n),�ѱ�,NO,���,ʵ��Ʊ��,����ʱ��,����ID,�շ����,�վݷ�Ŀ,���㵥λ,������,
    '      �շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��,���ձ���,ժҪ,��������ID,
    '      ִ�в���ID
    '����:ҽ��������ݵ����ݼ�()
    '����:���˺�
    '����:2011-08-15 16:40:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNo As String
    Dim i As Long, j As Long, lngSort As Long
    
    Err = 0: On Error GoTo errHand:
    
    Set rsDetail = New ADODB.Recordset
    If rsItems Is Nothing Or rsIncomes Is Nothing Then
        MsgBox "����ѡ��Һ���Ŀ", vbInformation, gstrSysName
        Exit Function
    End If
    
    rsDetail.Fields.Append "�ѱ�", adVarChar, 50, adFldIsNullable
    rsDetail.Fields.Append "���", adBigInt, , adFldIsNullable '����:42961
    rsDetail.Fields.Append "ʵ��Ʊ��", adVarChar, 20, adFldIsNullable
    rsDetail.Fields.Append "����ʱ��", adDBTimeStamp, , adFldIsNullable
    rsDetail.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    rsDetail.Fields.Append "�վݷ�Ŀ", adVarChar, 50, adFldIsNullable
    rsDetail.Fields.Append "���㵥λ", adVarChar, 50, adFldIsNullable
    rsDetail.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsDetail.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "����", adDouble, , adFldIsNullable
    rsDetail.Fields.Append "����", adDouble, , adFldIsNullable
    rsDetail.Fields.Append "ʵ�ս��", adDouble, , adFldIsNullable
    rsDetail.Fields.Append "ͳ����", adDouble, , adFldIsNullable
    rsDetail.Fields.Append "����֧������ID", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "�Ƿ�ҽ��", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "���ձ���", adVarChar, 50, adFldIsNullable
    rsDetail.Fields.Append "ժҪ", adVarChar, 200, adFldIsNullable
    rsDetail.Fields.Append "�Ƿ���", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "��������ID", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "ִ�в���ID", adBigInt, , adFldIsNullable
    rsDetail.CursorLocation = adUseClient
    rsDetail.LockType = adLockOptimistic
    rsDetail.CursorType = adOpenStatic
    rsDetail.Open
    
    If Not blnHav������ Then rsItems.Filter = "���� <> 3"
    If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        rsIncomes.Filter = "��ĿID=" & rsItems!��ĿID
        For j = 1 To rsIncomes.RecordCount
            lngSort = lngSort + 1
            rsDetail.AddNew
            rsDetail!�ѱ� = str�ѱ�
            rsDetail!��� = lngSort
            rsDetail!����ʱ�� = datDate
            rsDetail!����ID = lng����ID
            rsDetail!�շ���� = rsItems!���
            rsDetail!�վݷ�Ŀ = rsIncomes!�վݷ�Ŀ
            rsDetail!�շ�ϸĿID = rsIncomes!��ĿID
            rsDetail!���㵥λ = rsItems!���㵥λ
            rsDetail!���� = rsItems!����
            rsDetail!���� = rsIncomes!����
            rsDetail!ʵ�ս�� = rsIncomes!ʵ��
            rsDetail!ͳ���� = rsIncomes!ͳ����
            rsDetail!����֧������ID = rsItems!���մ���ID
            rsDetail!�Ƿ�ҽ�� = rsItems!������Ŀ��
            rsDetail!���ձ��� = rsItems!���ձ���
            rsDetail!ժҪ = Null
            rsDetail!�Ƿ��� = bln����
            rsDetail!��������ID = UserInfo.����ID
            rsDetail!ִ�в���ID = rsItems!ִ�п���ID
            rsDetail!������ = UserInfo.����
            rsDetail.Update
            rsIncomes.MoveNext
        Next
        rsItems.MoveNext
    Next
    
    If Not rsExpense Is Nothing Then
        If lngSort <> 0 Then lngSort = lngSort - 1
        '141815:���ϴ���2019/6/10��Ԥ����ʱ�ȶ�λ�б�
        If rsExpense.RecordCount <> 0 Then rsExpense.MoveFirst
        For i = 1 To rsExpense.RecordCount
            lngSort = lngSort + 1
            rsDetail.AddNew
            rsDetail!�ѱ� = str�ѱ�
            rsDetail!��� = lngSort
            rsDetail!����ʱ�� = datDate
            rsDetail!����ID = lng����ID
            rsDetail!�շ���� = rsExpense!���
            rsDetail!�վݷ�Ŀ = rsExpense!�վݷ�Ŀ
            rsDetail!�շ�ϸĿID = rsExpense!��ĿID
            rsDetail!���㵥λ = rsExpense!���㵥λ
            rsDetail!���� = rsExpense!����
            rsDetail!���� = rsExpense!����
            rsDetail!ʵ�ս�� = rsExpense!ʵ��
            rsDetail!ͳ���� = rsExpense!ͳ����
            rsDetail!����֧������ID = rsExpense!���մ���ID
            rsDetail!�Ƿ�ҽ�� = rsExpense!������Ŀ��
            rsDetail!���ձ��� = rsExpense!���ձ���
            rsDetail!ժҪ = Null
            rsDetail!�Ƿ��� = bln����
            rsDetail!��������ID = UserInfo.����ID
            rsDetail!ִ�в���ID = rsExpense!ִ�п���ID
            rsDetail!������ = UserInfo.����
            rsDetail.Update
            rsExpense.MoveNext
        Next
    End If
    If rsDetail.RecordCount > 0 Then rsDetail.MoveFirst
    zlMakeBillRecord = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function AddPayToList(ByVal objPay As clsPayInfo, ByVal vsfPay As VSFlexGrid, _
                Optional ByVal bln�쳣���� As Boolean, Optional ByVal byt֧������ As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ϣ���µ�֧���б���
    '���:objPayInfo-������Ϣ
    '����:
    '����:���ϴ�
    '����:2019/1/29 9:42:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim objSubPay As clsSubPayInfo
    If objPay Is Nothing Then AddPayToList = True: Exit Function
    If objPay.֧����� = 0 Then AddPayToList = True: Exit Function
    With vsfPay
        If objPay.Count > 0 Then
            .RemoveItem objPay.PayRow
            For Each objSubPay In objPay
                .Rows = .Rows + 1
                .RowData(objPay.PayRow) = objPay.��������
                .TextMatrix(.Rows - 1, .ColIndex("֧����ʽ")) = objSubPay.���㷽ʽ
                .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(objSubPay.������, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = objSubPay.���㷽ʽ
                .TextMatrix(.Rows - 1, .ColIndex("�������")) = objSubPay.�������
                .TextMatrix(.Rows - 1, .ColIndex("�����ID")) = objPay.�ӿ����
                .TextMatrix(.Rows - 1, .ColIndex("���ѿ�")) = IIf(objPay.���ѿ�, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("���ѿ�ID")) = objPay.���ѿ�ID
                .TextMatrix(.Rows - 1, .ColIndex("����")) = objPay.����
                .TextMatrix(.Rows - 1, .ColIndex("��������ID")) = objPay.��������ID
                .TextMatrix(.Rows - 1, .ColIndex("�޸�")) = IIf(byt֧������ = 0, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("����޸�")) = IIf(byt֧������ = 0, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("У�Ա�־")) = IIf(byt֧������ = 1, 2, 1)
                .Cell(flexcpData, .Rows - 1, .ColIndex("У�Ա�־")) = IIf(byt֧������ = 1, 1, 0) '�̶�
                .TextMatrix(.Rows - 1, .ColIndex("��������")) = IIf(objPay.��������, 1, 0)
                If byt֧������ = 2 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                Else
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = 0
                End If
            Next
        Else
            '���ָ������Ч������뵽���һ����
            If objPay.PayRow = 0 Or objPay.PayRow > .Rows - 1 Then
                objPay.PayRow = 0
                For lngRow = 1 To .Rows - 1
                    If Trim(.TextMatrix(lngRow, .ColIndex("֧����ʽ"))) = "" Then
                        objPay.PayRow = lngRow: Exit For
                    End If
                Next
                If objPay.PayRow = 0 Then
                    objPay.PayRow = .Rows
                    .Rows = .Rows + 1
                End If
            ElseIf bln�쳣���� Then
                .RemoveItem objPay.PayRow
                objPay.PayRow = .Rows
                .Rows = .Rows + 1
            End If
            .RowData(objPay.PayRow) = objPay.��������
            .TextMatrix(objPay.PayRow, .ColIndex("֧����ʽ")) = objPay.����
            .TextMatrix(objPay.PayRow, .ColIndex("���")) = Format(Val(.TextMatrix(objPay.PayRow, .ColIndex("���"))) + objPay.֧�����, "0.00")
            .TextMatrix(objPay.PayRow, .ColIndex("���㷽ʽ")) = objPay.���㷽ʽ
            .TextMatrix(objPay.PayRow, .ColIndex("�������")) = objPay.�������
            .TextMatrix(objPay.PayRow, .ColIndex("�����ID")) = objPay.�ӿ����
            .TextMatrix(objPay.PayRow, .ColIndex("���ѿ�")) = IIf(objPay.���ѿ�, 1, 0)
            .TextMatrix(objPay.PayRow, .ColIndex("���ѿ�ID")) = objPay.���ѿ�ID
            .TextMatrix(objPay.PayRow, .ColIndex("����")) = objPay.����
            .TextMatrix(objPay.PayRow, .ColIndex("��������ID")) = objPay.��������ID
            .TextMatrix(objPay.PayRow, .ColIndex("�޸�")) = IIf(byt֧������ = 1, 0, 1)
            .TextMatrix(objPay.PayRow, .ColIndex("����޸�")) = IIf(byt֧������ = 0, 1, 0)
            .TextMatrix(objPay.PayRow, .ColIndex("У�Ա�־")) = IIf(byt֧������ = 1, 2, 1)
            .Cell(flexcpData, objPay.PayRow, .ColIndex("У�Ա�־")) = IIf(byt֧������ = 1, 1, 0) '�̶�
            .TextMatrix(objPay.PayRow, .ColIndex("��������")) = IIf(objPay.��������, 1, 0)
            If byt֧������ = 2 Then
                .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = 0
            End If
        End If
    End With
End Function

Public Function GetPayInfo(ByVal colCardPayMode As Collection, ByVal str���㷽ʽ As String, _
                            objPayInfo As clsPayInfo) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���ݽ������ƻ�ȡ������Ϣ
    ' ��� : colCardPayMode:֧����Ϣ���ϣ��������֧����ʽʱ��ʼ��
    '      : str���㷽ʽ :��Ҫ��ȡ�Ľ��㷽ʽ
    ' ���� : objPayInfo�������������ʡ��������ʡ����㷽ʽ���ӿ���š��Ƿ����ѿ���֧������
    '      : bln��������:���㷽ʽ�Ƿ�����������������ͬ����
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2018/11/20 09:30
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    If objPayInfo Is Nothing Then Set objPayInfo = New clsPayInfo
    If str���㷽ʽ = "" Then GetPayInfo = True: Exit Function
    
    If str���㷽ʽ = "Ԥ����" Or str���㷽ʽ = "��Ԥ��" Then
        objPayInfo.���� = str���㷽ʽ
        objPayInfo.֧������ = Pay_AccountPay
        objPayInfo.�������� = 11
        GetPayInfo = True: Exit Function
    End If
    
    '���ȳ伯���в���,������һ�����⣬ҽ�ƿ����ƺ����ѿ�������ͬ
    'colCardPayMode:���ƣ����ʣ����㷽ʽ�������ID���Ƿ����ѿ����Ƿ��������
    For i = 1 To colCardPayMode.Count
        If colCardPayMode(i)(0) = str���㷽ʽ Then
            objPayInfo.���� = str���㷽ʽ
            objPayInfo.�������� = Val(colCardPayMode(i)(1))
            objPayInfo.���㷽ʽ = colCardPayMode(i)(2)
            objPayInfo.�ӿ���� = Val(colCardPayMode(i)(3))
            objPayInfo.���ѿ� = Val(colCardPayMode(i)(4)) = 1
            objPayInfo.�������� = Val(colCardPayMode(i)(5)) = 1
            If objPayInfo.�ӿ���� > 0 Then
                If objPayInfo.���ѿ� Then
                    objPayInfo.֧������ = Pay_SquarePay
                Else
                    objPayInfo.֧������ = Pay_ThreePay
                End If
            Else
                objPayInfo.֧������ = Pay_CashPay
            End If
            GetPayInfo = True
            Exit Function
        End If
    Next
    ' ʲôʱ��ѡ���˽��㷽ʽ�����ǻ�����û�е�
    MsgBox "��Ч�Ľ��㷽ʽ", vbInformation, gstrSysName
    Exit Function
    
    strSQL = "Select ���� From ���㷽ʽ Where ����=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "GetPayInfo", str���㷽ʽ)
    If Not rsTemp.EOF Then
        objPayInfo.���� = str���㷽ʽ
        objPayInfo.�������� = Val(rsTemp!����)
        objPayInfo.���㷽ʽ = objPayInfo.����
    Else
'        strSQL = "Select 1 As ����, a.Id, b.����, b.����, A.�Ƿ��������" & vbNewLine & _
'                "From ҽ�ƿ���� a, ���㷽ʽ b" & vbNewLine & _
'                "Where a.���㷽ʽ = b.���� And a.���� = [1]" & vbNewLine & _
'                "Union" & vbNewLine & _
'                "Select 2 As ����, c.��� As Id, d.����, d.����, 0 as �Ƿ��������" & vbNewLine & _
'                "From ���ѿ����Ŀ¼ c, ���㷽ʽ d" & vbNewLine & _
'                "Where c.���㷽ʽ = d.���� And c.���� = [1]"
'
'        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "GetPayInfo", str���㷽ʽ)
'        If rsTemp.EOF Then
'            MsgBox str���㷽ʽ & "����Ч�Ľ��㷽ʽ����ѡ������֧����ʽ��", vbInformation, gstrSysName
'            Exit Function
'        End If
'        objPayInfo.���� = str���㷽ʽ
'        objPayInfo.�������� = Val(Nvl(rsTemp!����))
'        objPayInfo.���㷽ʽ = Nvl(rsTemp!����)
'        objPayInfo.�ӿ���� = Val(Nvl(rsTemp!ID))
'        objPayInfo.���ѿ� = Val(Nvl(rsTemp!����)) = 2
'        objPayInfo.�������� = Val(Nvl(rsTemp!�Ƿ��������)) = 1
'        If objPayInfo.�ӿ���� > 0 Then
'            If objPayInfo.���ѿ� Then
'                objPayInfo.֧������ = Pay_SquarePay
'            Else
'                objPayInfo.֧������ = Pay_ThreePay
'            End If
'        Else
'            objPayInfo.֧������ = Pay_CashPay
'        End If
    End If
    GetPayInfo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceSQLByVsf(ByVal strNo As String, lng����ID As Long, ByVal vsfPay As VSFlexGrid, _
                                    cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� :
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/2/26 16:29
    '---------------------------------------------------------------------------------------
    Dim dbl��� As Double
    Dim str���㷽ʽ As String, str������Ϣ As String, strSQL As String
    Dim PayType As gPagePay
    Dim i As Long
    On Error GoTo errH
    If cllPro Is Nothing Then Set cllPro = New Collection
    With vsfPay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("У�Ա�־"))) = 1 And (Val(.TextMatrix(i, .ColIndex("���ѿ�"))) = 1 Or Val(.TextMatrix(i, .ColIndex("�����ID"))) = 0) Then
                dbl��� = Val(.TextMatrix(i, .ColIndex("���")))
                str���㷽ʽ = .TextMatrix(i, .ColIndex("���㷽ʽ"))
                If FormatEx(dbl���, 6) > 0 Then
                    If Val(.TextMatrix(i, .ColIndex("���ѿ�"))) = 1 Then
                        str������Ϣ = str���㷽ʽ & "," & dbl���
                        PayType = Pay_SquarePay
                    Else
                        str������Ϣ = str���㷽ʽ & "," & dbl��� & "," & .TextMatrix(i, .ColIndex("�������")) & ", "
                        PayType = Pay_CashPay
                    End If
                    strSQL = zlGetRegFeeModifySQL(strNo, lng����ID, str������Ϣ, PayType, , , , , _
                                Val(.TextMatrix(i, .ColIndex("�����ID"))), .TextMatrix(i, .ColIndex("����")))
                    Call zlAddArray(cllPro, strSQL)
                End If
            End If
        Next
    End With
    zlGetBalanceSQLByVsf = True
    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegistSql(ByVal int�Һ�ģʽ As Integer, _
                ByVal lng����ID As Long, ByVal str����� As String, ByVal str���� As String, _
                ByVal str�Ա� As String, ByVal str���� As String, ByVal str���ʽ As String, _
                ByVal str�ѱ� As String, ByVal str���ݺ� As String, ByVal lngִ�в���ID As Long, _
                ByVal str����ʱ�� As String, _
                ByVal str�Ǽ�ʱ�� As String, ByVal strҽ������ As String, ByVal byt���� As Byte, _
                ByVal str�ű� As String, ByVal str���� As String, _
                ByVal strժҪ As String, ByVal blnԤԼ�Һ� As Boolean, ByVal byt���� As Byte, _
                ByVal lng���� As Long, ByVal int���� As Integer, ByVal blnԤԼ���� As Boolean, _
                ByVal strԤԼ��ʽ As String, ByVal int�������� As Integer, ByVal int���� As Integer, _
                ByVal lng�Һ���ĿID As Long, Optional ByVal lng�����¼ID As Long, Optional ByVal int����ģʽ As Integer, _
                Optional ByVal str�շѵ� As String, Optional ByVal str������ˮ�� As String, _
                Optional ByVal str����˵�� As String, Optional ByVal str������λ As String) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���˹Һż�¼SQL
    ' ��� : �Һ���Ϣ
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/31 20:47
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If int�Һ�ģʽ = 0 Then
        strSQL = "Zl_���˹Һż�¼_Insert_S("
    Else
        strSQL = "Zl_���˹Һż�¼_����_Insert_S("
    End If
    '    ����id_In        ���˹Һż�¼.����id%Type,
    strSQL = strSQL & "" & ZVal(lng����ID) & ","
    '    �����_In        ���˹Һż�¼.�����%Type,
    strSQL = strSQL & "" & IIf(str����� = "", "NULL", str�����) & ","
    '    ����_In          ���˹Һż�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '    �Ա�_In          ���˹Һż�¼.�Ա�%Type,
    strSQL = strSQL & "'" & str�Ա� & "',"
    '    ����_In          ���˹Һż�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '    ���ʽ_In      ���˹Һż�¼.ҽ�Ƹ��ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ����
    strSQL = strSQL & "'" & str���ʽ & "',"
    '    �ѱ�_In          ���˹Һż�¼.�ѱ�%Type,
    strSQL = strSQL & "'" & str�ѱ� & "',"
    '    ���ݺ�_In        ���˹Һż�¼.No%Type,
    strSQL = strSQL & "'" & str���ݺ� & "',"
    '    ִ�в���id_In    ���˹Һż�¼.ִ�в���ID%Type,
    strSQL = strSQL & "" & lngִ�в���ID & ","
    '    ����Ա���_In    ���˹Һż�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '    ����Ա����_In    ���˹Һż�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '    ����ʱ��_In      ���˹Һż�¼.����ʱ��%Type,
    strSQL = strSQL & "" & "To_Date('" & str����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '    �Ǽ�ʱ��_In      ���˹Һż�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "" & "To_Date('" & str�Ǽ�ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '    ҽ������_In      �ҺŰ���.ҽ������%Type,
    strSQL = strSQL & "'" & strҽ������ & "',"
    '    ����_In          Number,
    strSQL = strSQL & "" & byt���� & ","
    '    �ű�_In          �ҺŰ���.����%Type,
    strSQL = strSQL & "'" & str�ű� & "',"
    '    ����_In          ���˹Һż�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '    ժҪ_In          ���˹Һż�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
    strSQL = strSQL & "'" & strժҪ & "',"
    '    ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
    strSQL = strSQL & "" & IIf(blnԤԼ�Һ�, 1, 0) & ","
    '    ����_In          ���˹Һż�¼.����%Type := 0,
    strSQL = strSQL & "" & byt���� & ","
    '    ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
    strSQL = strSQL & "" & ZVal(lng����) & ","
    '    ����_In          ���˹Һż�¼.����%Type := Null,
    strSQL = strSQL & "" & ZVal(int����) & ","
    '    ԤԼ����_In      Number := 0,
    strSQL = strSQL & "" & IIf(blnԤԼ����, 1, 0) & ","
    '    ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
    strSQL = strSQL & "'" & strԤԼ��ʽ & "',"
    '    ��������_In      Number := 0,
    strSQL = strSQL & "" & int�������� & ","
    '    ����_In          ���˹Һż�¼.����%Type := Null,
    strSQL = strSQL & "" & ZVal(int����) & ","
    '    �Һ���ĿID_In    ���˹Һż�¼.�Һ���ĿID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng�Һ���ĿID) & ","
    '    �����¼id_In    ���˹Һż�¼.�����¼ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng�����¼ID) & ","
    '    ����ģʽ_In      ���˹Һż�¼.����ģʽ%Type := 0,
    strSQL = strSQL & "" & int����ģʽ & ","
    '    �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null,
    strSQL = strSQL & "'" & str�շѵ� & "',"
    '    ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '    ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & str����˵�� & "',"
    '    ������λ_In      ����Ԥ����¼.������λ%Type := Null
    strSQL = strSQL & "'" & str������λ & "')"
    
    zlGetRegistSql = strSQL
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegistCollectSql(ByVal strҽ������ As String, ByVal lngҽ��ID As Long, _
                ByVal lng��Ŀid As Long, ByVal lngִ�в���ID As Long, ByVal str����ʱ�� As String, ByVal bytԤԼ��־ As Byte, _
                ByVal str�ű� As String, Optional ByVal lng��¼ID As Long) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���˹ҺŻ���SQL
    ' ��� : �Һ���Ϣ
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/31 20:47
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHand
    
     strSQL = "zl_���˹ҺŻ���_Update("
    '  ҽ������_In   �ҺŰ���.ҽ������%Type,
    strSQL = strSQL & "'" & strҽ������ & "',"
    '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
    strSQL = strSQL & "" & ZVal(lngҽ��ID) & ","
    '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
    strSQL = strSQL & "" & lng��Ŀid & ","
    '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
    strSQL = strSQL & "" & lngִ�в���ID & ","
    '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
    strSQL = strSQL & "" & "To_Date('" & str����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����,3-�շ�ԤԼ
    strSQL = strSQL & bytԤԼ��־ & ","
    '  ����_In       �ҺŰ���.����%Type := Null
    strSQL = strSQL & "'" & str�ű� & "',"
    '  ��������_In   Number := 0
    strSQL = strSQL & "" & "Null" & ","
    '  �����¼id_In �ٴ������¼.Id%Type := Null
    strSQL = strSQL & "" & ZVal(lng��¼ID) & ")"
    
    zlGetRegistCollectSql = strSQL
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegistFeeSql(ByVal cllPro As Collection, ByVal lng����ID As Long, ByVal str����� As String, ByVal str���� As String, _
                ByVal str�Ա� As String, ByVal str���� As String, ByVal str���ʽ As String, _
                ByVal str�ѱ� As String, ByVal str���ݺ� As String, ByVal strƱ�ݺ� As String, _
                ByVal lng��� As Long, ByVal int�۸񸸺� As Long, ByVal int�������� As Long, _
                ByVal str�շ���� As String, ByVal lng�շ�ϸĿid As Long, ByVal int���� As Integer, _
                ByVal dbl��׼���� As Double, ByVal lng������Ŀid As Long, ByVal str�վݷ�Ŀ As String, _
                ByVal dblӦ�ս�� As Double, ByVal dblʵ�ս�� As Double, ByVal lng���˿���ID As Long, _
                ByVal lng��������ID As Long, ByVal lngִ�в���ID As Long, ByVal str�Ǽ�ʱ�� As String, _
                ByVal str����ʱ�� As String, ByVal strҽ������ As String, ByVal lng����ID As Long, _
                ByVal dbl���ʽ�� As Double, ByVal lng���մ���ID As Long, ByVal int������Ŀ�� As Integer, _
                ByVal dblͳ���� As Double, ByVal str���ձ��� As String, ByVal bln������ As Boolean, _
                ByVal byt���� As Byte, ByVal str�ű� As String, ByVal str���� As String, _
                ByVal lng���� As Long, ByVal blnԤԼ�Һ� As Boolean, ByVal strԤԼ��ʽ As String, _
                ByVal strժҪ As String, ByVal int�Ʒѷ�ʽ As Integer, Optional ByVal str�շѵ� As String, _
                Optional ByVal str���㵥λ As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���˹Һŷ���SQL
    ' ��� : �Һŷ�����Ϣ
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/31 20:47
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHand
    If cllPro Is Nothing Then Set cllPro = New Collection
    
    strSQL = "Zl_���˹Һŷ���_Insert_S("
    '    ����id_In        ���˹Һż�¼.����id%Type,
    strSQL = strSQL & "" & ZVal(lng����ID) & ","
    '    �����_In        ���˹Һż�¼.�����%Type,
    strSQL = strSQL & "" & IIf(str����� = "", "NULL", str�����) & ","
    '    ����_In          ���˹Һż�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '    �Ա�_In          ���˹Һż�¼.�Ա�%Type,
    strSQL = strSQL & "'" & str�Ա� & "',"
    '    ����_In          ���˹Һż�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '    ���ʽ_In      ���˹Һż�¼.ҽ�Ƹ��ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
    strSQL = strSQL & "'" & str���ʽ & "',"
    '    �ѱ�_In          ���˹Һż�¼.�ѱ�%Type,
    strSQL = strSQL & "'" & str�ѱ� & "',"
    '    ���ݺ�_In        ���˹Һż�¼.No%Type,
    strSQL = strSQL & "'" & str���ݺ� & "',"
    '    Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
    strSQL = strSQL & "'" & strƱ�ݺ� & "',"
    '    ���_In          ������ü�¼.���%Type,
    strSQL = strSQL & "" & lng��� & ","
    '    �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
    strSQL = strSQL & "" & ZVal(int�۸񸸺�) & ","
    '    ��������_In      ������ü�¼.��������%Type,
    strSQL = strSQL & "" & ZVal(int��������) & ","
    '    �շ����_In      ������ü�¼.�շ����%Type,
    strSQL = strSQL & "'" & str�շ���� & "',"
    '    �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
    strSQL = strSQL & "" & lng�շ�ϸĿid & ","
    '    ����_In          ������ü�¼.����%Type,
    strSQL = strSQL & "" & int���� & ","
    '    ��׼����_In      ������ü�¼.��׼����%Type,
    strSQL = strSQL & "" & dbl��׼���� & ","
    '    ������Ŀid_In    ������ü�¼.������Ŀid%Type,
    strSQL = strSQL & "" & lng������Ŀid & ","
    '    �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
    strSQL = strSQL & "'" & str�վݷ�Ŀ & "',"
    '    Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
    strSQL = strSQL & "" & IIf(str�շѵ� <> "", 0, dblӦ�ս��) & ","
    '    ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & IIf(str�շѵ� <> "", 0, dblʵ�ս��) & ","
    '    ���˿���id_In    ������ü�¼.���˿���id%Type,
    strSQL = strSQL & "" & lng���˿���ID & ","
    '    ��������id_In    ������ü�¼.��������id%Type,
    strSQL = strSQL & "" & lng��������ID & ","
    '    ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
    strSQL = strSQL & "" & lngִ�в���ID & ","
    '    ����Ա���_In    ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '    ����Ա����_In    ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '    �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "" & "To_Date('" & str�Ǽ�ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '    ����ʱ��_In      ������ü�¼.����ʱ��%Type,
    strSQL = strSQL & "" & "To_Date('" & str����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '    ҽ������_In      ������ü�¼.ִ����%Type,
    strSQL = strSQL & "'" & strҽ������ & "',"
    '    ����id_In        ������ü�¼.����id%Type,
    strSQL = strSQL & "" & ZVal(lng����ID) & ","
    '    ���ʽ��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
    strSQL = strSQL & "" & IIf(str�շѵ� <> "", 0, dbl���ʽ��) & ","
    '    ���մ���id_In    ������ü�¼.���մ���id%Type,
    strSQL = strSQL & "" & ZVal(lng���մ���ID) & ","
    '    ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
    strSQL = strSQL & "" & ZVal(int������Ŀ��) & ","
    '    ͳ����_In      ������ü�¼.ͳ����%Type,
    strSQL = strSQL & "" & ZVal(dblͳ����) & ","
    '    ���ձ���_In      ������ü�¼.���ձ���%Type,
    strSQL = strSQL & "'" & str���ձ��� & "',"
    '    ������_In Number, --������¼�Ƿ���������
    strSQL = strSQL & "" & IIf(bln������, 1, 0) & ","
    '    ����_In          Number,
    strSQL = strSQL & "" & byt���� & ","
    '    �ű�_In          ������ü�¼.���㵥λ%Type,
    strSQL = strSQL & "'" & str�ű� & "',"
    '    ����_In          ������ü�¼.��ҩ����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '    ����_In          ������ü�¼.��ҩ����%Type,
    strSQL = strSQL & "" & ZVal(lng����) & ","
    '    ԤԼ�Һ�_In      Number := 0,
    strSQL = strSQL & "" & IIf(blnԤԼ�Һ�, 1, 0) & ","
    '    ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
    strSQL = strSQL & "'" & strԤԼ��ʽ & "',"
    '    ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
    strSQL = strSQL & "'" & strժҪ & "',"
    '    �Ʒѷ�ʽ_In      Number := 0,
    strSQL = strSQL & "" & int�Ʒѷ�ʽ & ","
    '    �շѵ�_In        ������ü�¼.No%Type := Null
    strSQL = strSQL & "'" & str�շѵ� & "')"
    zlAddArray cllPro, strSQL
    
    If str�շѵ� <> "" Then
        strSQL = "Zl_���ﻮ�ۼ�¼_Insert_S("
        '    No_In           ������ü�¼.No%Type,
        strSQL = strSQL & "'" & str�շѵ� & "',"
        '    ���_In         ������ü�¼.���%Type,
        strSQL = strSQL & "" & lng��� & ","
        '    ����id_In       ������ü�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '    ��ҳid_In       סԺ���ü�¼.��ҳid%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '    ��ʶ��_In       ������ü�¼.��ʶ��%Type,
        strSQL = strSQL & "" & IIf(str����� = "", "NULL", str�����) & ","
        '    ���ʽ_In     ������ü�¼.���ʽ%Type,
        strSQL = strSQL & "'" & str���ʽ & "',"
        '    ����_In         ������ü�¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '    �Ա�_In         ������ü�¼.�Ա�%Type,
        strSQL = strSQL & "'" & str�Ա� & "',"
        '    ����_In         ������ü�¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '    �ѱ�_In         ������ü�¼.�ѱ�%Type,
        strSQL = strSQL & "'" & str�ѱ� & "',"
        '    �Ӱ��־_In     ������ü�¼.�Ӱ��־%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '    ���˿���id_In   ������ü�¼.���˿���id%Type,
        strSQL = strSQL & "" & lng���˿���ID & ","
        '    ��������id_In   ������ü�¼.��������id%Type,
        strSQL = strSQL & "" & lng��������ID & ","
        '    ������_In       ������ü�¼.������%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '    ��������_In     ������ü�¼.��������%Type,
        strSQL = strSQL & "" & ZVal(int��������) & ","
        '    �շ�ϸĿid_In   ������ü�¼.�շ�ϸĿid%Type,
        strSQL = strSQL & "" & lng�շ�ϸĿid & ","
        '    �շ����_In     ������ü�¼.�շ����%Type,
        strSQL = strSQL & "'" & str�շ���� & "',"
        '    ���㵥λ_In     ������ü�¼.���㵥λ%Type,
        strSQL = strSQL & "'" & str���㵥λ & "',"
        '    ��ҩ����_In     ������ü�¼.��ҩ����%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '    ����_In         ������ü�¼.����%Type,
        strSQL = strSQL & "" & 1 & ","
        '    ����_In         ������ü�¼.����%Type,
        strSQL = strSQL & "" & int���� & ","
        '    ���ӱ�־_In     ������ü�¼.���ӱ�־%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '    ִ�в���id_In   ������ü�¼.ִ�в���id%Type,
        strSQL = strSQL & "" & lngִ�в���ID & ","
        '    �۸񸸺�_In     ������ü�¼.�۸񸸺�%Type,
        strSQL = strSQL & "" & ZVal(int�۸񸸺�) & ","
        '    ������Ŀid_In   ������ü�¼.������Ŀid%Type,
        strSQL = strSQL & "" & lng������Ŀid & ","
        '    �վݷ�Ŀ_In     ������ü�¼.�վݷ�Ŀ%Type,
        strSQL = strSQL & "'" & str�վݷ�Ŀ & "',"
        '    ��׼����_In     ������ü�¼.��׼����%Type,
        strSQL = strSQL & "" & dbl��׼���� & ","
        '    Ӧ�ս��_In     ������ü�¼.Ӧ�ս��%Type,
        strSQL = strSQL & "" & dblӦ�ս�� & ","
        '    ʵ�ս��_In     ������ü�¼.ʵ�ս��%Type,
        strSQL = strSQL & "" & dblʵ�ս�� & ","
        '    ����ʱ��_In     ������ü�¼.����ʱ��%Type,
        strSQL = strSQL & "" & "To_Date('" & str����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
        '    �Ǽ�ʱ��_In     ������ü�¼.�Ǽ�ʱ��%Type,
        strSQL = strSQL & "" & "To_Date('" & str�Ǽ�ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
        '    ����Ա����_In   ������ü�¼.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '    ����id_In       ������ü�¼.Id%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '    ����ժҪ_In     ������ü�¼.ժҪ%Type := Null,
        strSQL = strSQL & "'" & "�Һ�:" & str���ݺ� & "')"
        zlAddArray cllPro, strSQL
    End If
    zlGetRegistFeeSql = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegFeeModifySQL(ByVal strNo As String, ByVal lng����ID As Long, _
                ByVal strBalance As String, Optional ByVal Pay�������� As gPagePay = Pay_CashPay, _
                Optional ByVal intУ�Ա�־ As Integer = 2, Optional ByVal bln��ɽ��� As Boolean, _
                Optional ByVal bln�������� As Boolean, _
                Optional ByVal lng����ID As Long, Optional ByVal lng�����ID As Long, _
                Optional ByVal str���� As String, Optional str������ˮ�� As String, _
                Optional ByVal str����˵�� As String, Optional ByVal bln��ͨ���� As Boolean) As String
    Dim strSQL As String
    '����У�Ա�־����ɹҺ��շ�
    strSQL = "Zl_���˹Һ��շ�_Modify_S("
    '���ݺ�_In         ������ü�¼.No%Type
    strSQL = strSQL & "'" & strNo & "',"
    '����id_In         ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '������Ϣ_In       Varchar2,
    strSQL = strSQL & "'" & strBalance & "',"
    '��������_In       Number := 0,
    strSQL = strSQL & "" & Pay�������� & ","
    '��ɱ�־_In       Number := 0,
    strSQL = strSQL & "" & IIf(bln��ɽ���, 1, 0) & ","
    '��������_In       Number := 0,
    strSQL = strSQL & "" & IIf(bln��������, 1, 0) & ","
    '��������ID_In     ����Ԥ����¼.��������ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng����ID) & ","
    '�����ID_In       ����Ԥ����¼.�����ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng�����ID) & ","
    '����_In           ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & str���� & "',"
    '������ˮ��_In     ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '����˵��_In       ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & str����˵�� & "',"
    '��ͨ����_In       Number := 0
    strSQL = strSQL & "" & IIf(bln��ͨ����, 1, 0) & ","
    'У�Ա�־_In       Number := 2
    strSQL = strSQL & "" & intУ�Ա�־ & ")"
    
    zlGetRegFeeModifySQL = strSQL
End Function
    
Public Function zlGetRegDoneSQL(ByVal strNo As String, ByVal bln��Ժ As Boolean, ByVal blnԤԼ As Boolean, _
                Optional ByVal bln���ɶ��� As Boolean) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ��ɹҺŹ���SQL
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/1 20:15
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHand
    
    strSQL = "Zl_���˹Һż�¼_��ɹҺ�_S("
    '    ���ݺ�_In     ������ü�¼.No%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '    ��Ժ����_In   Number := 0,
    strSQL = strSQL & "" & IIf(bln��Ժ, 1, 0) & ","
    '    ԤԼ��־_In   Number := 0,
    strSQL = strSQL & "" & IIf(blnԤԼ, 1, 0) & ","
    '    ���ɶ���_In Number:=0
    strSQL = strSQL & "" & IIf(bln���ɶ���, 1, 0) & ")"

    zlGetRegDoneSQL = strSQL
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function zlGetCancelSql(ByVal lngReg����ID As Long, cllBack As Collection, _
                Optional ByVal blnԤԼ As Boolean, Optional ByVal bln��ɾ�����㷽ʽ As Boolean) As Boolean
    Dim strSQL As String
    Set cllBack = New Collection
    
    If lngReg����ID <> 0 Then
        strSQL = "Zl_���˹Һż�¼_Cancel(" & lngReg����ID & ", " & IIf(blnԤԼ, 2, 0) & "," & IIf(bln��ɾ�����㷽ʽ, 1, 0) & ")"
        zlAddArray cllBack, strSQL
    End If
    
End Function

Public Function zlReadAddrInfo(ByVal objService As clsService, _
                            ByVal objCtrl As PatiAddress, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                            ByVal intTYPE As Integer, Optional ByVal strAddress As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���Ĳ��˵�ַ��Ϣ���ؼ���
    '���:objCtrl-�ṹ����ַ�ؼ�,intType -��ַ����1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ
    '����:
    '����:���ϴ�
    '����:2015/12/3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str�������� As String, str��ַ_ʡ As String, str��ַ_�� As String
    Dim str��ַ_�� As String, str��ַ_�� As String, str��ַ_���� As String
    On Error GoTo errHandle
    If objService Is Nothing Then Exit Function
    If lng����ID = 0 Then zlReadAddrInfo = True: Exit Function
    If objService.zlPatiSvr_GetPatiAddrssInfo(lng����ID, lng��ҳID, intTYPE, str��ַ_ʡ, str��ַ_��, _
                    str��ַ_��, str��ַ_��, str��ַ_����, str��������) Then
         Call objCtrl.LoadStructAdress(str��ַ_ʡ, str��ַ_��, str��ַ_��, str��ַ_��, str��ַ_����)
    Else
        objCtrl.value = strAddress
    End If
    zlReadAddrInfo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Calc_Age(ByVal lng����ID As Long, ByVal str�������� As String, Optional ByVal str�������� As String) As String
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    On Error GoTo errH
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    Zl_Calc_Age = objOneCardComLib.Zl_Calc_Age(lng����ID, str��������, str��������)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function SimilarIDs(str���֤�� As String) As String
    '���ܣ���鲡���Ƿ����������Ϣ
    '���أ����Ƽ�¼�Ĳ���ID��,��"234,235,236"
    Dim i As Integer
    Dim cllPati As Collection
    Dim cllFilter As New Collection
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    On Error GoTo errH
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    cllFilter.Add Array("���֤��", str���֤��)
    If objOneCardComLib.zlGetPatiInfsByFilter(2, False, cllFilter, cllPati) = False Then Exit Function
    If cllPati Is Nothing Then Exit Function
    
    For i = 1 To cllPati.Count
        SimilarIDs = SimilarIDs & "|ID:" & cllPati("����ID")
        SimilarIDs = SimilarIDs & ",����:" & cllPati("����")
        SimilarIDs = SimilarIDs & ",�����:" & IIf(cllPati("�����") = "", "��", cllPati("�����"))
        SimilarIDs = SimilarIDs & ",���֤��:" & IIf(cllPati("���֤��") = "", cllPati("���֤��"), "δ�Ǽ�")
        SimilarIDs = SimilarIDs & ",��ַ:" & IIf(cllPati("��ͥ��ַ") = "", cllPati("��ͥ��ַ"), "δ�Ǽ�")
        SimilarIDs = SimilarIDs & ",�Ǽ�����:" & Format(cllPati("�Ǽ�ʱ��"), "YYYY-MM-DD")
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetCardByName(ByVal strCardName As String, ByVal bln���ѿ� As Boolean, ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���ݿ�������ƻ�ȡ������
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/25 16:28
    '---------------------------------------------------------------------------------------
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    If objOneCardComLib.zlGetCardFromTypeName(strCardName, bln���ѿ�, objCard) = False Then Exit Function
    If objCard Is Nothing Then GetCardByName = True
End Function

Public Function GetPatiIDByName(ByVal frmMain As Object, ByVal objControl As Object, ByVal strName As String, _
    ByVal str�Ա� As String, ByRef lngPatiID As Long, _
    Optional ByVal blnContסԺ As Boolean, Optional ByVal blnAddNewPati As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�����������ȡ������Ϣ
    '���:objControl-���õĿؼ�
    '     strName-������Ϣ
    '     frmMain-���õ�������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnUserCancel As Boolean
    Dim rsPati As ADODB.Recordset, rsSel As ADODB.Recordset
    Dim intNameDays As Integer
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errHand
    If LenB(strName) < 4 Then Exit Function
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    intNameDays = Val(gobjDatabase.GetPara("������������", glngSys, 9000, 0))
    
    If objOneCardComLib.zlGetPatiIdFromPatiName(objControl, strName, lngPatiID, frmMain, intNameDays, , IIf(blnContסԺ, 0, 1), 1, blnUserCancel, blnAddNewPati) = False Then Exit Function
    If blnUserCancel Then Exit Function
    If blnAddNewPati Then GetPatiIDByName = True: Exit Function

    GetPatiIDByName = lngPatiID <> 0
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function CheckUsed�����(ByVal lng����ID As Long, ByVal str����� As String, _
                ByRef blnUsedByOther As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���������Ƿ�ʹ��
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/12 11:33
    '---------------------------------------------------------------------------------------
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    CheckUsed����� = objOneCardComLib.zlCheckOutNoIsExist(lng����ID, str�����, blnUsedByOther)
End Function

Public Function CheckMobile(str�ֻ��� As String, Optional ByVal lng����ID As Long, _
                    Optional ByVal blnShowMsg As Boolean = True, Optional ByRef strErrMsg As String) As Boolean
    '���ܣ��ж�ָ���ֻ����Ƿ�����ȷ���ֻ��Ÿ�ʽ�Լ��Ƿ��Ѿ����������ݿ���
    '��Σ�str�ֻ���-���м����ֻ���
    '      lng����ID - ����ֻ����ظ��ԣ�����Ҫ���ʹ�0
    '      blnShowMsg-����ʱ�Ƿ���ʾ��ʾ
    '���Σ�strErrMsg -������Ϣ
    '���أ��ֻ�������ʹ��-true�����򷵻�False
    Dim blnUsedByOther As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMobileRange As String
    Dim blnQuery As Boolean
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errH
    '127941:���ϴ�,2018/8/10,�����ֻ��Ŷμ���ֻ����Ƿ�Ϸ�
    If str�ֻ��� = "" Then CheckMobile = True: Exit Function
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    strSQL = "Select Max(�����) As �����" & vbNewLine & _
                "From (Select Decode(���볤��, Length([1]), 1, 2) As �����" & vbNewLine & _
                "       From �ֻ��ų��úŶα�" & vbNewLine & _
                "       Where �Ŷ� = Substr([1], 1, Length(�Ŷ�)))"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "�ֻ��ż��", str�ֻ���)
    
    If rsTmp.RecordCount = 0 Then
        strErrMsg = "δ���ڡ��ֻ��ų��úŶα�������������ֻ��Ÿ�ʽ��������¼�룡"
    ElseIf Val(Nvl(rsTmp!�����)) = 0 Then
        strErrMsg = "δ���ڡ��ֻ��ų��úŶα�������������ֻ��Ÿ�ʽ��������¼�룡"
    ElseIf Val(Nvl(rsTmp!�����)) = 2 Then
        strErrMsg = "������ֻ���λ������ȷ��������¼�룡"
    End If
    
    If gSysPara.bln����ֻ����ظ� Then
        If objOneCardComLib.zlCheckPhoneIsExist(lng����ID, str�ֻ���, blnUsedByOther, Not blnShowMsg, strErrMsg) = False Then Exit Function
        If blnUsedByOther Then
            strErrMsg = "������ֻ��������������ظ����Ƿ�ȷ��¼�룿"
            blnQuery = True
        End If
    End If
    
    If strErrMsg <> "" Then
        If Not blnShowMsg Then Exit Function
        If blnQuery Then
            If MsgBox(strErrMsg, vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
                strErrMsg = "": Exit Function
            End If
            strErrMsg = ""
        Else
            MsgBox strErrMsg, vbInformation, gstrSysName
            strErrMsg = "": Exit Function
        End If
    End If
    CheckMobile = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckUsedMCNO(ByVal strMCNO As String) As Boolean
    '����:���ҽ�����Ƿ��Ѵ���
    Dim blnUsed As Boolean
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    If objOneCardComLib.zlCheckMCNOIsExist(strMCNO, blnUsed) = False Then Exit Function

    If blnUsed Then
        MsgBox "ҽ����" & strMCNO & "�Ѵ��ڣ����顣", vbInformation, gstrSysName
    End If
    CheckUsedMCNO = Not blnUsed
End Function

Public Function GetPatientInfo(objPati As clsPatientInfo, ByVal frmMain As Object, ByVal objControl As Object, _
                ByVal str��ѯ��ʽ As String, ByVal lng�����ID As Long, ByVal strInput As String, _
                Optional ByVal blnCard As Boolean, Optional ByVal blnCont��Ժ As Boolean = True, _
                Optional ByVal blnSeekName As Boolean, Optional ByVal strName As String, Optional ByVal strSex As String, _
                Optional ByRef strPassWord As String, Optional ByRef blnUserCancel As Boolean, Optional ByRef intCardStatus As Integer, _
                Optional ByRef strValidTime As String, Optional ByVal bln������֤ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������Ϣ
    ' ��� : int��ѯ��ʽ��0��������id����;-1:����ģʽ����;>0 ����������
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/6 16:19
    '---------------------------------------------------------------------------------------
    Dim blnValidTime As Boolean '��֤����Ч��ֹʱ��
    Dim cllPati As Collection, cllOtherFindCons As Collection
    Dim strErrMsg As String
    Dim lngDefaultCardTypeID As Long, lng����ID As Long
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errHand
    strPassWord = ""
    Set objPati = New clsPatientInfo
    Set cllOtherFindCons = New Collection
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    If str��ѯ��ʽ = "����ID" Then
        lng����ID = Val(strInput)
    ElseIf str��ѯ��ʽ = "����" Or str��ѯ��ʽ = "��������￨" Then
        If blnCard Then
            If objOneCardComLib.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, , objControl, frmMain, , , , blnUserCancel, , blnValidTime, intCardStatus, strValidTime) = False Then lng����ID = 0
            
        ElseIf blnSeekName Then
            If GetPatiIDByName(frmMain, objControl, strInput, strSex, lng����ID, blnCont��Ժ, True) = False Then Exit Function
            If lng����ID = 0 Then GetPatientInfo = True: Exit Function  '�²���
        End If
        If lng����ID = 0 Then
            If objOneCardComLib.zlIsMobileNo(strInput) Then
                lng�����ID = 0
                str��ѯ��ʽ = "�ֻ���"
            End If
        End If
    ElseIf str��ѯ��ʽ = "�ֻ���" Then
        If objOneCardComLib.zlIsMobileNo(strInput) = False Then Exit Function
    ElseIf lng�����ID > 0 Then
        str��ѯ��ʽ = lng�����ID
        blnValidTime = True
    End If
    
    If (str��ѯ��ʽ = "����" Or str��ѯ��ʽ = "��������￨") And Not blnCard And lng����ID = 0 Then
        GetPatientInfo = True: Exit Function   '�²���
    End If
    If lng����ID = 0 Then
       If objOneCardComLib.zlGetPatiID(str��ѯ��ʽ, strInput, False, lng����ID, strPassWord, , , objControl, frmMain, , , , blnUserCancel, , blnValidTime, intCardStatus, strValidTime) = False Then Exit Function
    End If
    If lng����ID = 0 Then Exit Function
    
    If strPassWord <> "" Then
        If Not VerifyPassWord(frmMain, strPassWord) Then
            MsgBox "���������֤ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If objOneCardComLib.zlGetPatiInforFromPatiID(lng����ID, objPati, strErrMsg) = False Then lng����ID = 0: Exit Function
    If Not blnCont��Ժ And objPati.��Ժ Then
        If objOneCardComLib.zlGetInpatiState(objPati.����ID, objPati.��ҳID, , cllPati) Then
            If Val(cllPati("����״̬")) = 0 Then
                Set objPati = New clsPatientInfo: lng����ID = 0: Exit Function
            End If
        End If
    End If
    GetPatientInfo = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiID(ByVal frmMain As Object, ByVal objControl As Object, _
                ByVal strCardTypes As String, ByVal strInput As String, ByRef lng����ID As Long, _
                Optional ByVal strCardPassWord As String, Optional ByVal blnUserCancel As Boolean, _
                Optional ByVal blnNotShowErr As Boolean, Optional ByRef intCardStatus As Integer, _
                Optional ByVal blnShowMergePati As Boolean, Optional ByRef strValidTime As String) As Boolean
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    If objOneCardComLib.zlGetPatiID(strCardTypes, strInput, blnNotShowErr, lng����ID, strCardPassWord, , , objControl, frmMain, blnShowMergePati, True, , blnUserCancel, , True, intCardStatus, strValidTime) = False Then Exit Function
End Function

Public Function GetPatientOtherInfo(ByVal lng����ID As Long, cllDrug As Collection, _
                                cllImmune As Collection, cllOther As Collection, cllContact As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���˽���ҳ��Ϣ
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/8 16:48
    '---------------------------------------------------------------------------------------
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    On Error GoTo errHand
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    If objOneCardComLib.zlGetPatiOtherInforFromPatiID(lng����ID, , , , True, True, , True, _
                    , cllDrug, cllImmune, , cllOther, cllContact) = False Then Exit Function
    GetPatientOtherInfo = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function zlUpdateOutMedRec(ByVal lng����ID As Long, Optional ByVal str������ As String, Optional ByVal str�������� As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �������ﲡ����û�м�¼ʱ������¼,����str������ʱɾ��������¼
    ' ��� : str������-�����
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/21 13:53
    '---------------------------------------------------------------------------------------
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib

    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    zlUpdateOutMedRec = objOneCardComLib.zlUpdateOutMedRec(lng����ID, str������, str��������)
End Function

Public Function zlCheckMzLgPatiUseDeposit(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������۲����Ƿ���ʹ������Ԥ��
    '���:lng����ID-����ID
    '����:
    '����:�����������ܹ�ʹ������Ԥ������true,���򷵻�False
    '����:���ϴ�
    '����:2019/9/6 15:57:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnLimitDeposit As Boolean, blnMzLgPati As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errHandle
    blnLimitDeposit = Val(gobjDatabase.GetPara(Val("323-�������۲���Ԥ����ʹ�ÿ���"), glngSys)) <> 0
    If Not blnLimitDeposit Then zlCheckMzLgPatiUseDeposit = True: Exit Function
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    If objOneCardComLib.zlCheckMzLgPati(lng����ID, lng��ҳID, blnMzLgPati, True) = False Then Exit Function
    zlCheckMzLgPatiUseDeposit = Not blnMzLgPati
    
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetMoneyInfoRegist(lng����ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵�����Ԥ��ʣ���
    '���:
    '       curModiMoney=�޸�ʱ,ԭ���ݵĵ�ǰ���˵ķ��úϼ�
    '       int����:����(0-�����סԺ����;1-����;2-סԺ),-1��ʾ����
    '       bytModiMoneyType-�޸ķ��õ����(�ڰ����ͳ��ʱ��Ч)
    '       blnFamilyMoney-�Ƿ��ȡ�������
    '����:
    '����:����ʣ���
    '����:���˺�
    '����:2011-07-21 15:33:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, blnFamilyMoney As Boolean
    Dim strSQL As String, strFamilyIds As String
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errH
    blnFamilyMoney = True
    
    If blnFamilyMoney Then
        If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
        Call objOneCardComLib.ZlGetPatiFamilyMember(1, lng����ID, strFamilyIds)
    End If
    
    strSQL = "Select " & IIf(blnFamilyMoney, "0 As ����,", "") & _
            "       Nvl(�������,0) As �������,Nvl(Ԥ�����,0) As Ԥ�����" & _
            " From �������" & _
            " Where ����=1 And ����ID=[1] And ���� = 1"
    '79868,��ȡ���˼������
    If blnFamilyMoney And strFamilyIds <> "" Then
        strSQL = strSQL & " Union All " & _
                " Select /*+cardinality(B,10) */ " & IIf(blnFamilyMoney, "1 As ����,", "") & _
                "       Nvl(a.�������, 0) As �������, Nvl(a.Ԥ�����, 0) As Ԥ�����" & _
                " From ������� A, (Select Column_Value as ����ID from Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) B" & _
                " Where a.����id = b.����id And a.���� = 1 And a.���� = 1 "
    End If

    strSQL = "Select " & IIf(blnFamilyMoney, "����,", "") & _
            "       nvl(Sum(�������),0) as �������,nvl(Sum(Ԥ�����),0) as Ԥ����� " & _
            " From (" & strSQL & ")" & vbCrLf & _
                IIf(blnFamilyMoney, " Group by ����", "")
    
    Set GetMoneyInfoRegist = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, strFamilyIds)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Check����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As Boolean
'����:�жϲ����Ƿ��ٴε�����ͬ�ٴ����ʵ��ٴ����ҡ��Һ�
'     �����ҹ��ŵ�,��ס��Ժ��,���ﲻ��ȷ��ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHand
    strSQL = "Select Zl1_Fun_Getreturnvisit([1],[2]) As �����־ From Dual"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������", lng����ID, lngִ�в���ID)
    Check���� = Val(Nvl(rsTmp!�����־)) = 1
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ZValStr(ByVal lngTmp As Long) As String
    ZValStr = IIf(lngTmp = 0, "", lngTmp)
End Function
