Attribute VB_Name = "mdlExpense"
Option Explicit

Public Function MoneyOverFlow(objBill As ExpenseBill) As Boolean
'���ܣ���鵥�ݺϼƽ���Ƿ����
'˵������Currency����922337203685477Ϊ׼
    Dim dblӦ�� As Double, dblʵ�� As Double
    Dim i As Integer, j As Integer
    
    'Ҫ��VALתΪDouble��������
    For i = 1 To objBill.Details.Count
        For j = 1 To objBill.Details(i).InComes.Count
            If Abs(dblӦ�� + Val(objBill.Details(i).InComes(j).Ӧ�ս��)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            If Abs(dblʵ�� + Val(objBill.Details(i).InComes(j).ʵ�ս��)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            dblӦ�� = dblӦ�� + Val(objBill.Details(i).InComes(j).Ӧ�ս��)
            dblʵ�� = dblʵ�� + Val(objBill.Details(i).InComes(j).ʵ�ս��)
        Next
    Next
End Function

Public Function GetBillTotal(objBill As ExpenseBill) As Currency
'���ܣ���ȡ���ݷ�Ŀ�ϼƽ��
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    
    For Each objBillDetail In objBill.Details
        For Each objBillIncome In objBillDetail.InComes
            GetBillTotal = GetBillTotal + objBillIncome.ʵ�ս��
        Next
    Next
End Function

Public Function GetServiceDept(str�շ�ϸĿIDs As String) As ADODB.Recordset
'����:��ȡ���ҩ���Ĵ洢�ⷿ��������
    Dim strSQL As String, rsTmp As New ADODB.Recordset
        
    strSQL = " Select Distinct �շ�ϸĿID,Nvl(��������ID,0) as ��������ID,ִ�п���id From �շ�ִ�п��� Where �շ�ϸĿID In (" & str�շ�ϸĿIDs & ") "
    
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    If Not rsTmp.EOF Then Set GetServiceDept = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDrugTotal(ByVal objBill As ExpenseBill, ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long) As Double
'���ܣ���ȡ������ָ��ҩƷ��ͬһҩ�����е�������
    Dim i As Integer, dblCount As Double
    
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).�շ�ϸĿID = lngҩƷID And objBill.Details(i).ִ�в���ID = lngҩ��ID Then
            dblCount = dblCount + objBill.Details(i).���� * objBill.Details(i).����
        End If
    Next
    GetDrugTotal = dblCount
End Function

Public Function GetFirstRow(curBill As ExpenseBill, Optional strClass As String) As Integer
'���ܣ���ȡ��ǰ�����е�һ��ΪҩƷ���շ��к�
'������strClass=ȡ��һ��ҩ����ҩ��,��ΪҩƷ
'���أ�0=û��ҩƷ�շ���
    Dim i As Long
    If curBill.Details.Count = 0 Then GetFirstRow = 0
    For i = 1 To curBill.Details.Count
        If strClass = "" Then
            If InStr(",5,6,7,", curBill.Details(i).�շ����) > 0 Then
                GetFirstRow = i: Exit Function
            End If
        Else
            If curBill.Details(i).�շ���� = strClass Then
                GetFirstRow = i: Exit Function
            End If
        End If
    Next
End Function

Public Function Get������׼��Ŀ(lng����ID As Long, strField As String) As String
    Dim rsTmp As New ADODB.Recordset
    Dim lng����ID As Long, int���� As Integer, strSQL As String
    Dim strA1 As String, strA2 As String, strB1 As String, strB2 As String
    
    On Error GoTo errH
            
    '��ȡ���˲���,�Ǹò����Ƿ��и�����׼��Ŀ����
    strSQL = _
        " Select A.����,A.����ID,Nvl(B.����,0) as ����,B.����,Count(*)" & _
        " From �����ʻ� A,������׼��Ŀ B" & _
        " Where Nvl(A.����ID,0)=B.����ID And Nvl(A.����ID,0)<>0" & _
        " And B.���� IN(1,2) And A.����ID=[1]" & _
        " Group by A.����,A.����ID,Nvl(B.����,0),B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID)
    If rsTmp.EOF Then Exit Function
    
    lng����ID = rsTmp!����ID
    int���� = rsTmp!����
    
    '������շ�ϸĿ
    rsTmp.Filter = "����=0 And ����=1"
    If Not rsTmp.EOF Then
        strA1 = strField & _
            " IN (" & _
            "   Select �շ�ϸĿID From ����֧����Ŀ" & _
            "   Where ���� = " & int���� & _
            "   And �շ�ϸĿID IN (" & _
            "       Select �շ�ϸĿID From ������׼��Ŀ Where Nvl(����,0)=0 And ����=1 And ����ID=" & lng����ID & ")" & _
            ")"
    End If
    
    '����ı��մ���
    rsTmp.Filter = "����=1 And ����=1"
    If Not rsTmp.EOF Then
        strA2 = strField & _
            " IN (" & _
            "   Select �շ�ϸĿID From ����֧����Ŀ" & _
            "   Where ���� = " & int���� & _
            "   And ����ID IN (" & _
            "       Select �շ�ϸĿID From ������׼��Ŀ Where Nvl(����,0)=1 And ����=1 And ����ID=" & lng����ID & ")" & _
            ")"
    End If
    
    '��ֹ���շ�ϸĿ
    rsTmp.Filter = "����=0 And ����=2"
    If Not rsTmp.EOF Then
        strB1 = strField & _
            " Not IN (" & _
            "   Select �շ�ϸĿID From ����֧����Ŀ" & _
            "   Where ���� = " & int���� & _
            "   And �շ�ϸĿID IN (" & _
            "       Select �շ�ϸĿID From ������׼��Ŀ Where Nvl(����,0)=0 And ����=2 And ����ID=" & lng����ID & ")" & _
            ")"
    End If
    
    '��ֹ�ı��մ���
    rsTmp.Filter = "����=1 And ����=2"
    If Not rsTmp.EOF Then
        strB2 = strField & _
            " Not IN (" & _
            "   Select �շ�ϸĿID From ����֧����Ŀ" & _
            "   Where ���� = " & int���� & _
            "   And ����ID IN (" & _
            "       Select �շ�ϸĿID From ������׼��Ŀ Where Nvl(����,0)=1 And ����=2 And ����ID=" & lng����ID & ")" & _
            ")"
    End If
    
    '���SQL(������Ҫ��Or)
    strSQL = ""
    If strA1 <> "" And strA2 <> "" Then
        strSQL = " And (" & strA1 & " Or " & strA2 & ")"
    Else
        If strA1 <> "" Then strSQL = " And " & strA1
        If strA2 <> "" Then strSQL = " And " & strA2
    End If
    If strB1 <> "" Then strSQL = strSQL & " And " & strB1
    If strB2 <> "" Then strSQL = strSQL & " And " & strB2
        
    Get������׼��Ŀ = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ְ��(lngҩƷID As Long) As String
'���ܣ�����ҩƷID��ȡ�䴦��ְ��
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    Get����ְ�� = "00"
    strSQL = "Select Nvl(B.����ְ��,'00') as ����ְ�� From ҩƷ��� A,ҩƷ���� B Where A.ҩ��ID=B.ҩ��ID And A.ҩƷID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngҩƷID)
    If Not rsTmp.EOF Then Get����ְ�� = rsTmp!����ְ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��������(lngID As Long) As Double
'���ܣ���ȡָ��ҩƷ�Ĵ�������,�����۵�λ���ء�
'������lngID=ҩƷID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(A.��������,0) as ��������" & _
        " From ҩƷ���� A,ҩƷ��� B Where A.ҩ��ID=B.ҩ��ID And B.ҩƷID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngID)
    If Not rsTmp.EOF Then Get�������� = rsTmp!��������
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemExistInsure(ByVal lng�շ�ϸĿID As Long, ByVal int���� As Integer) As Boolean
'���ܣ��ж��շ���Ŀ�Ƿ������˱���֧����Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    If gclsInsure.GetCapability(support��������ҽ����Ŀ, , int����) Then
        ItemExistInsure = True: Exit Function
    End If
    
    strSQL = "Select * From ����֧����Ŀ Where �շ�ϸĿID=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng�շ�ϸĿID, int����)
    ItemExistInsure = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckLimit(ByVal objBill As ExpenseBill, Optional ByVal intRow As Integer, Optional ByVal blnҩ����λ As Boolean) As Boolean
'���ܣ����õ���ҩƷ�����������
'˵����
'   1.ȫ��û���������������棻���г���ҩƷ�����ں�������ʾ�������ؼ١�
'   2.���ʱ���Ϊÿ�����˵������
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim tmpDetail As BillDetail, curDetail As BillDetail
    Dim strItemIDs As String, i As Integer, j As Integer
    Dim dblTime As Double, dbl���� As Double
    
    CheckLimit = True
    If objBill.Details.Count = 0 Then Exit Function
    
    On Error GoTo errH
    
    '�ռ�����
    For i = 1 To objBill.Details.Count
        If intRow = 0 Or (intRow > 0 And i = intRow) Then
            With objBill.Details(i)
                '�ռ�ҩƷID
                If InStr(strItemIDs & ",", "," & .�շ�ϸĿID & ",") = 0 And InStr(",5,6,7,", .�շ����) > 0 Then
                    strItemIDs = strItemIDs & "," & .�շ�ϸĿID
                End If
            End With
        End If
    Next
    If strItemIDs = "" Then Exit Function
    strItemIDs = Mid(strItemIDs, 2)
        
    strSQL = "Select A.ҩƷID,A.����ϵ��,B.���㵥λ as ������λ" & _
        " From ҩƷ��� A,������ĿĿ¼ B" & _
        " Where A.ҩ��ID=B.ID And A.ҩƷID IN (" & strItemIDs & ")"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlExpense") 'In
    
    strItemIDs = ""
    For j = 1 To objBill.Details.Count
        If intRow = 0 Or (intRow > 0 And j = intRow) Then
            Set tmpDetail = objBill.Details(j)
            If InStr(",5,6,7,", tmpDetail.�շ����) > 0 And tmpDetail.Detail.�������� > 0 Then
                If InStr(strItemIDs, "," & tmpDetail.�շ�ϸĿID) = 0 Then
                    dblTime = 0
                    For Each curDetail In objBill.Details
                        If InStr(",5,6,7,", curDetail.�շ����) > 0 And tmpDetail.�շ�ϸĿID = curDetail.�շ�ϸĿID Then
                            dblTime = dblTime + curDetail.���� * curDetail.����
                        End If
                    Next
                    rsTmp.Filter = "ҩƷID=" & tmpDetail.�շ�ϸĿID
                    If Not rsTmp.EOF Then
                        If blnҩ����λ Then
                            dbl���� = dblTime * tmpDetail.Detail.ҩ����װ * rsTmp!����ϵ��
                        Else
                            dbl���� = dblTime * rsTmp!����ϵ��
                        End If
                        If dbl���� > tmpDetail.Detail.�������� Then
                            MsgBox "ҩƷ """ & tmpDetail.Detail.���� & """ ���ܼ��� " & _
                                FormatEx(dbl����, 5) & rsTmp!������λ & "(" & FormatEx(dblTime, 5) & IIF(blnҩ����λ, tmpDetail.Detail.ҩ����λ, tmpDetail.Detail.���㵥λ) & ") ������������ " & _
                                FormatEx(tmpDetail.Detail.��������, 5) & rsTmp!������λ & " ��", vbInformation, gstrSysName
                            CheckLimit = False: Exit Function
                        End If
                    End If
                    strItemIDs = strItemIDs & "," & tmpDetail.�շ�ϸĿID
                End If
            End If
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStockInfo(lngҩƷID As Long, blnҩ�� As Boolean, blnҩ�� As Boolean, _
    Optional ByVal blnҩ����λ As Boolean, Optional strҩ����װ As String) As String
'���ܣ���ȡҩƷ�ڸ���ҩ����ҩ��Ŀ����Ϣ
'������"blnҩ��/blnҩ��"����Ҫ��һ������Ϊ��
'���أ�������Ϣ
    Dim strSQL As String, strSQL2 As String
    Dim str���� As String, i As Long
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    If blnҩ�� And blnҩ�� Then
        str���� = "'��ҩ��','��ҩ��','��ҩ��','��ҩ��','��ҩ��','��ҩ��'"
    ElseIf blnҩ�� Then
        str���� = "'��ҩ��','��ҩ��','��ҩ��'"
    ElseIf blnҩ�� Then
        str���� = "'��ҩ��','��ҩ��','��ҩ��'"
    End If
    
    '�ų�������ʵ����,���������סԺ
    strSQL = _
        " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And Instr([1],B.��������)>0"
    'ҩ��������ҩƷ����Ч��
    strSQL2 = "Select ����ID From ��������˵�� Where �������� IN('��ҩ��','��ҩ��','��ҩ��')"
    '�����������ҩƷ
    strSQL = _
        " Select B.����,B.����,A.�ⷿID," & _
        " Nvl(Sum(A.��������),0)" & IIF(blnҩ����λ, "/Nvl(C." & strҩ����װ & ",1)", "") & " as ���" & _
        " From ҩƷ��� A,(" & strSQL & ") B,ҩƷ��� C" & _
        " Where A.�ⷿID=B.ID And A.ҩƷID=C.ҩƷID" & _
        " And ((A.Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
        " Or (Nvl(C.ҩ������,0)=0 And A.�ⷿID IN(" & strSQL2 & ")))" & _
        " And A.����=1 And A.ҩƷID=[2]" & _
        " Group by B.����,B.����,A.�ⷿID,Nvl(C." & strҩ����װ & ",1)" & _
        " Having Sum(Nvl(A.��������,0))<>0" & _
        " Order By B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", str����, lngҩƷID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!���� & ":" & rsTmp!���
        rsTmp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    GetStockInfo = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function OverTime() As Boolean
'���ܣ��жϵ�ǰ�Ƿ��ڼӰ�ʱ�䷶Χ��
'���أ���-��ǰ���ڼӰ�ʱ����,��-������
    Dim str���� As String, str���� As String
    Dim DateBegin As Date, DateEnd As Date
    Dim curTime As Date
    
    str���� = GetSysParVal(1): str���� = GetSysParVal(2)
    curTime = CDate(Format(zlDatabase.Currentdate, "HH:MM:SS"))
    
    If str���� <> "" Then
        DateBegin = CDate(Trim(Split(UCase(str����), "AND")(0)))
        DateEnd = CDate(Trim(Split(UCase(str����), "AND")(1)))
    End If
    
    If Not (curTime >= DateBegin And curTime <= DateEnd) Then
        If str���� <> "" Then
            DateBegin = CDate(Trim(Split(UCase(str����), "AND")(0)))
            DateEnd = CDate(Trim(Split(UCase(str����), "AND")(1)))
        End If
        If Not (curTime >= DateBegin And curTime <= DateEnd) Then OverTime = True
    End If
End Function

Public Function GetBillRows(str���ݺ� As String, int��¼���� As Integer) As Integer
'���ܣ���ȡһ�ŷ��õ�����δ���ϵķ�������
'������int��¼����=1-�շ�(����),2-����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    strSQL = _
        " Select ���,Sum(����) as ʣ������" & _
        " From (" & _
        " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���," & _
        " Avg(Nvl(����, 1) * ����) As ����" & _
        " From ���˷��ü�¼" & _
        " Where NO=[1] And ��¼����=[2]" & _
        " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", str���ݺ�, int��¼����)
    If Not rsTmp.EOF Then GetBillRows = rsTmp.RecordCount
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistInsure(strNO As String) As Integer
'���ܣ��ж�ָ����סԺ���ʵ����Ƿ��ҽ�����˼ǵ���
'������strNO=���ʵ��ݺ�
'���أ�������򷵻ز�������
'˵����1.ֻ��סԺҽ������,�������ﲡ�˵�ҽ������
'      2.���ʱ�ֻ���ص�һ�����˵�����,������ҲӦ��ֻ��һ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.���� From ���˷��ü�¼ A,������ҳ B" & _
        " Where A.��¼����=2 And A.��¼״̬ IN(0,1,3) And B.���� is Not NULL" & _
        " And A.NO=[1] And A.����ID=B.����ID And A.��ҳID=B.��ҳID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then BillExistInsure = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsureName(intInsure As Integer) As String
'���ܣ����ݱ��������Ż�ȡ�����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select * From ������� Where ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", intInsure)
    If Not rsTmp.EOF Then GetInsureName = Nvl(rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'���ܣ���ȡҩƷ�����ĳ�����ļ���
'������bytType:0-ҩƷ��1-����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '�������
    
    strSQL = _
        " Select Distinct A.ID,C.��鷽ʽ" & _
        " From ���ű� A,��������˵�� B," & IIF(bytType = 0, "ҩƷ������", "���ϳ�����") & " C" & _
        " Where B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� " & IIF(bytType = 0, "IN('��ҩ��','��ҩ��','��ҩ��')", "='���ϲ���'") & _
        " And C.�ⷿID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetStockCheck")
    For i = 1 To rsTmp.RecordCount
        colStock.Add Nvl(rsTmp!��鷽ʽ, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function CheckDisable(objBill As ExpenseBill) As String
'���ܣ���鵥���е�ҩƷ�Ľ������
'���أ�ҩƷ���������ʾ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim i As Long, j As Long, k As Long
    Dim strGroup As String, strIDs As String
    Dim blnStop As Boolean
    
    For i = 1 To objBill.Details.Count
        If InStr(",5,6,7,", objBill.Details(i).�շ����) > 0 Then
            strIDs = strIDs & "," & objBill.Details(i).�շ�ϸĿID
        End If
    Next
    strIDs = Mid(strIDs, 2)
    If strIDs = "" Or UBound(Split(strIDs, ",")) < 1 Then Exit Function
    
    strSQL = _
        " Select A.����,Count(Distinct A.��ĿID) as ������" & _
        " From ���ƻ�����Ŀ A,ҩƷ��� B" & _
        " Where A.��ĿID=B.ҩ��ID And B.ҩƷID IN(" & strIDs & ")" & _
        " Having Count(Distinct A.��ĿID)>1 Group by A.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlExpense") 'In
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strGroup = strGroup & "," & rsTmp!����
            rsTmp.MoveNext
        Next
        strGroup = Mid(strGroup, 2)
        
        For i = 0 To UBound(Split(strGroup, ","))
            strSQL = _
                "Select Distinct C.����,C.����,D.����,D.����,D.���" & _
                " From ҩƷ��� A,������ĿĿ¼ B,���ƻ�����Ŀ C,�շ���ĿĿ¼ D" & _
                " Where A.ҩ��ID=B.ID And B.ID=C.��ĿID And A.ҩƷID=D.ID" & _
                " And C.����=" & Split(strGroup, ",")(i) & _
                " And A.ҩƷID IN(" & strIDs & ")" & _
                " Order by C.����,C.����,D.����"
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlExpense") 'In
            If Not rsTmp.EOF Then
                rsTmp.Filter = "����=1"
                If rsTmp.RecordCount > 1 Then
                    k = k + 1
                    strInfo = strInfo & vbCrLf & "�� " & k & " ��(��������)��" & vbCrLf
                    For j = 1 To rsTmp.RecordCount
                        strInfo = strInfo & "[" & rsTmp!���� & "]" & rsTmp!���� & IIF(IsNull(rsTmp!���), "", "(" & rsTmp!��� & ")") & "                 " & vbCrLf
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Filter = "����=2"
                If rsTmp.RecordCount > 1 Then
                    blnStop = True
                    k = k + 1
                    strInfo = strInfo & vbCrLf & "�� " & k & " ��(�������)��" & vbCrLf
                    For j = 1 To rsTmp.RecordCount
                        strInfo = strInfo & "[" & rsTmp!���� & "]" & rsTmp!���� & IIF(IsNull(rsTmp!���), "", "(" & rsTmp!��� & ")") & "                 " & vbCrLf
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Filter = 0
            End If
        Next
        If strInfo <> "" Then
            If blnStop Then
                CheckDisable = "���ֵ���������ҩƷ������û����ã�" & vbCrLf & strInfo & vbCrLf & "���޸Ľ���ҩƷ���ټ�����"
            Else
                CheckDisable = "���ֵ���������ҩƷ������û����ã�" & vbCrLf & strInfo & vbCrLf & "Ҫ������"
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ImportBill(ByVal int��Դ As Integer, ByVal str���ݺ� As String, _
    ByVal int��¼���� As Integer, Optional ByVal bln��� As Boolean) As ExpenseBill
'���ܣ���ȡ���õ��ݵ����ݶ�����(Ŀǰ���Դ�����Ŀ,����������Ŀ),�����޸Ļ���ʱ��
'������int��¼����=1-�շ�(����),2-����
'      bln���=�Ƿ���ķ��õǼ�,ʵ�ս��Ϊ0
'���أ���ŵ�����Ϣ�ĵ��ݶ���
'˵������Ϊ������ʱ��Ŀ�۸���Ϣ��������,���Է�������������¼���
'      �����ǵ��뻹���޸ĵ���,����Ӧ������ͣ���շ�ϸĿ
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    Dim intCurNo As Integer, strInfo As String
    Dim int��� As Integer, blnDo As Boolean, i As Integer
    
    Dim dblAllTime As Double, dblCurTime As Double
    Dim dblPrice As Double, strҩ�� As String
    
    Dim colSerial As New Collection '���ڴ����������
    Dim strSQL As String
    
    Dim lng��ҩ�� As Long, lng��ҩ�� As Long, lng��ҩ�� As Long
    Dim blnҩ����λ As Boolean, strҩ����λ As String, strҩ����װ As String
    
    'ȱʡҩ��
    lng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
    lng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
    lng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(int��Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", 0))
    
    'ҩƷ��λ
    blnҩ����λ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ҩƷ��λ", 0)) <> 0
    If int��Դ = 1 Then
        strҩ����λ = "���ﵥλ": strҩ����װ = "�����װ"
    Else
        strҩ����λ = "סԺ��λ": strҩ����װ = "סԺ��װ"
    End If
        
    '------------------------------------------------------------------------------------------
    '�շѼ�Ŀ����:�¼���۸�,����ж���۸�,��һ���շ�ϸĿID�оͻ��ж��������ͬ�ļ�¼
    '�۸񸸺� is NULL:ֻȡÿ���շ�ϸĿID�ĵ�һ��(ҩƷֻ��һ��),��ΪҪ����۸�
        
    'ʹ��ָ����ҩ����ȡ��ȷ�Ŀ��
    strҩ�� = "Decode(A.�շ����,'5'," & IIF(lng��ҩ�� <> 0, lng��ҩ��, "A.ִ�в���ID") & "," & _
        "'6'," & IIF(lng��ҩ�� <> 0, lng��ҩ��, "A.ִ�в���ID") & "," & _
        "'7'," & IIF(lng��ҩ�� <> 0, lng��ҩ��, "A.ִ�в���ID") & ",A.ִ�в���ID)"
    
    'ҩ�������������ҩƷ����Ч��
    strSQL = _
        " Select X.ҩƷID,W.����ID,W.��������," & _
        " A.��� As ���,A.��������,A.NO,A.��¼����,A.��¼״̬,A.�ಡ�˵�,A.Ӥ����,A.�ѱ�,A.����,A.�Ա�,A.����," & _
        " A.����,A.��ʶ��,A.����ID,A.��ҳID,A.���˲���ID,A.���˿���ID,A.��������ID,A.�����־,A.�Ӱ��־," & _
        " A.���ӱ�־,A.�շ����,A.�շ�ϸĿID,A.��ҩ����,Nvl(����,1) as ����,Nvl(A.����,0) as ����," & _
        " A.��׼���� As ��׼����," & strҩ�� & " as ִ�в���ID,A.������,A.������,A.����Ա���,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��,A.ժҪ," & _
        " B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(F.����,B.����) as ����,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
        " B.���ηѱ�,B.˵��,B.ִ�п���,Nvl(A.��������,B.��������) ��������,D.�ּ�,D.ԭ��,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
        " E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
        " Decode(A.�շ����,'4',1,X." & strҩ����װ & ") as ҩ����װ," & _
        " Decode(A.�շ����,'4',B.���㵥λ,X." & strҩ����λ & ") as ҩ����λ," & _
        " Decode(A.�շ����,'4',Nvl(W.���÷���,0),Nvl(X.ҩ������,0)) as ����,Nvl(Y.���,0) As ���,B.¼������" & _
        " From ���˷��ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E,�շ���Ŀ���� F,�������� W,ҩƷ��� X," & _
        "   (Select A.ҩƷID,A.�ⷿID,Sum(Nvl(A.��������,0)) as ��� From ҩƷ��� A" & _
        "       Where A.����=1 And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
        "       And A.ҩƷID IN(Select �շ�ϸĿID From ���˷��ü�¼ Where ��¼����=[2] And ��¼״̬ IN(0,1,3) And NO=[1])" & _
        "    Group by A.ҩƷID,A.�ⷿID) Y" & _
        " Where A.��¼����=[2] And A.��¼״̬ IN(0,1,3) And A.NO=[1]" & _
        " And A.�۸񸸺� Is Null And A.�շ�ϸĿID=B.ID And A.�շ�ϸĿID=D.�շ�ϸĿID" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is NULL)" & _
        " And A.�շ����=C.���� And A.�շ�ϸĿID=X.ҩƷID(+) And A.�շ�ϸĿID=W.����ID(+) And D.������ĿID=E.ID" & _
        " And A.�շ�ϸĿID=Y.ҩƷID(+) And " & strҩ�� & "=Y.�ⷿID(+)" & _
        " And A.�շ�ϸĿID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=[3]" & _
        " And ((Sysdate Between D.ִ������ And D.��ֹ����) Or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"

    strSQL = "Select * From (" & strSQL & ") Order by ���"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", str���ݺ�, int��¼����, IIF(gbln��Ʒ��, 3, 1))
    
    'û�м�¼���ǿյ���
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    If rsTmp.RecordCount <> 0 Then
        With rsTmp
            i = 1
            Do While Not .EOF
                '����������=====================================================
                If i = 1 Then
                    objBill.NO = !NO
                    objBill.����ID = Nvl(!����ID, 0)
                    objBill.��ҳID = Nvl(!��ҳID, 0)
                    objBill.����ID = Nvl(!���˲���ID, 0)
                    objBill.����ID = Nvl(!���˿���ID, 0)
                    objBill.���� = Nvl(!����)
                    objBill.�Ա� = Nvl(!�Ա�)
                    objBill.���� = Nvl(!����)
                    objBill.�ѱ� = Nvl(!�ѱ�)
                    objBill.��ʶ�� = Nvl(!��ʶ��, 0)
                    objBill.���� = Nvl(!����)
                    objBill.�ѱ� = Nvl(!�ѱ�)
                    objBill.�����־ = Nvl(!�����־, 0)
                    objBill.�Ӱ��־ = Nvl(!�Ӱ��־, 0)
                    objBill.Ӥ���� = Nvl(!Ӥ����, 0)
                    objBill.��������ID = Nvl(!��������ID, 0)
                    objBill.������ = Nvl(!������)
                    objBill.������ = Nvl(!������)
                    objBill.����Ա��� = Nvl(!����Ա���)
                    objBill.����Ա���� = Nvl(!����Ա����)
                    objBill.����ʱ�� = !����ʱ��
                    objBill.�Ǽ�ʱ�� = !�Ǽ�ʱ��
                    objBill.�ಡ�˵� = Nvl(!�ಡ�˵�, 0) <> 0
                End If
                
                '�����շ�ϸĿ=====================================================
                Set objBillDetail = New BillDetail
                Set objBillDetail.Detail = New Detail
                            
                '�������,��������
                intCurNo = intCurNo + 1
                objBillDetail.��� = intCurNo 'ʵ�����к�
                colSerial.Add intCurNo, "_" & !��� '��¼ԭ������ڵ��к�
                If Not IsNull(!��������) Then
                    objBillDetail.�������� = colSerial("_" & !��������)
                End If
                                                                    
                'ʹ��ԭ���Ķ�̬�ѱ�
                objBillDetail.�շ���� = !�շ����
                objBillDetail.�շ�ϸĿID = !�շ�ϸĿID
                objBillDetail.���㵥λ = IIF(IsNull(!���㵥λ), "", !���㵥λ)
                
                objBillDetail.���� = Nvl(!����, 1)
                If InStr(",5,6,7,", !�շ����) > 0 And blnҩ����λ Then
                    objBillDetail.���� = Nvl(!����, 0) / Nvl(!ҩ����װ, 1)
                Else
                    objBillDetail.���� = Nvl(!����, 0)
                End If
                
                objBillDetail.���ӱ�־ = Nvl(!���ӱ�־, 0)
                objBillDetail.ժҪ = Nvl(!ժҪ)
                objBillDetail.ִ�в���ID = Nvl(!ִ�в���ID, 0)
                objBillDetail.��ҩ���� = Nvl(!��ҩ����)
                objBillDetail.Detail.ID = !�շ�ϸĿID
                objBillDetail.Detail.���� = !����
                objBillDetail.Detail.��� = Nvl(!�Ƿ���, 0) = 1
                objBillDetail.Detail.�������� = 0 '!!!Ŀǰ���Դ�����Ŀ,����������Ŀ
                objBillDetail.Detail.���д��� = 0 '!!!Ŀǰ���Դ�����Ŀ,����������Ŀ
                objBillDetail.Detail.��� = Nvl(!���)
                objBillDetail.Detail.���㵥λ = Nvl(!���㵥λ)
                
                objBillDetail.Detail.ҩ����λ = Nvl(!ҩ����λ)
                objBillDetail.Detail.ҩ����װ = Nvl(!ҩ����װ, 1)
                If InStr(",5,6,7,", !�շ����) > 0 And blnҩ����λ Then
                    objBillDetail.Detail.��� = Nvl(!���, 0) / Nvl(!ҩ����װ, 1)
                Else
                    objBillDetail.Detail.��� = Nvl(!���, 0)
                End If
                objBillDetail.Detail.¼������ = Val("" & !¼������)
                
                objBillDetail.Detail.�Ӱ�Ӽ� = Nvl(!�Ӱ�Ӽ�, 0) <> 0
                objBillDetail.Detail.��� = Nvl(!���)
                objBillDetail.Detail.������� = Nvl(!�������)
                objBillDetail.Detail.���� = Nvl(!����)
                objBillDetail.Detail.���ηѱ� = Nvl(!���ηѱ�, 0) <> 0
                objBillDetail.Detail.˵�� = Nvl(!˵��)
                objBillDetail.Detail.ִ�п��� = Nvl(!ִ�п���, 0)
                objBillDetail.Detail.���� = Nvl(!��������)
                objBillDetail.Detail.����ְ�� = Get����ְ��(objBillDetail.Detail.ID)
                
                objBillDetail.Detail.ҩ��ID = Nvl(!ҩ��ID, 0)
                objBillDetail.Detail.��� = Nvl(!�Ƿ���, 0) <> 0
                objBillDetail.Detail.���� = Nvl(!����, 0) <> 0
                objBillDetail.Detail.�������� = Nvl(!��������, 0) = 1
                objBillDetail.Detail.Ҫ������ = 0
                
                '����۸񲿷�=====================================================
                Set objBillDetail.InComes = New BillInComes
                Do
                    '�������еļ۸��������¼���
                    If !�Ƿ��� = 1 Then
                        If InStr(",5,6,7,", !�շ����) > 0 Or (!�շ���� = "4" And Nvl(!��������, 0) = 1) Then
                            '----------------------------------------------------------------------------------------------
                            'ʱ��ҩƷ����۸�(�����ɲ�����)
                            dblAllTime = !���� * !���� '�������ۼ�����
                            If dblAllTime <> 0 Then
                                strSQL = _
                                    " Select Nvl(A.����,0) as ����,Nvl(A.��������,0) as ���," & _
                                    "   Nvl(Decode(Nvl(A.ʵ������,0),0,0,A.ʵ�ʽ��/A.ʵ������),0) as ʱ��" & _
                                    " From ҩƷ��� A" & _
                                    " Where A.�ⷿID=[1] And A.ҩƷID=[2] And Nvl(A.��������,0)>0" & _
                                    "   And A.����=1 And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
                                    " Order by Nvl(A.����,0)"
                                Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", objBillDetail.ִ�в���ID, Val(!�շ�ϸĿID))
                                'ʱ��=�ܽ��/������
                                dblPrice = 0
                                For i = 1 To rsPrice.RecordCount
                                    If dblAllTime = 0 Then Exit For
                                    'ȡС��
                                    If dblAllTime <= rsPrice!��� Then
                                        dblCurTime = dblAllTime
                                    Else
                                        dblCurTime = rsPrice!���
                                    End If
                                    dblPrice = dblPrice + Format(dblCurTime * Format(rsPrice!ʱ��, "0.00000"), gstrDec)
                                    dblAllTime = Val(dblAllTime) - Val(dblCurTime)
                                    rsPrice.MoveNext
                                Next
                                If dblAllTime <> 0 Then
                                    '����δ�ֽ����
                                    If !�շ���� = "4" Then
                                        MsgBox "ʱ����������""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    Else
                                        MsgBox "ʱ��ҩƷ""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.��׼���� = 0
                                Else
                                    objBillIncome.��׼���� = Format(dblPrice / (!���� * !����), "0.00000") '�������ۼۼ۸�
                                End If
                            Else
                                objBillIncome.��׼���� = 0
                            End If
                            '----------------------------------------------------------------------------------------------
                        Else
                            If Abs(!��׼����) > Abs(Nvl(!�ּ�, 0)) Then
                                objBillIncome.��׼���� = Nvl(!ԭ��, 0)
                            Else
                                objBillIncome.��׼���� = !��׼����
                            End If
                        End If
                    Else
                        objBillIncome.��׼���� = !�ּ�
                    End If
                                        
                    If InStr(",5,6,7,", !�շ����) > 0 And blnҩ����λ Then
                        objBillIncome.��׼���� = Format(objBillIncome.��׼���� * Nvl(!ҩ����װ, 1), "0.00000")
                    Else
                        objBillIncome.��׼���� = Format(objBillIncome.��׼����, "0.00000")
                    End If
                    objBillIncome.�ּ� = Nvl(!�ּ�, 0) '�ּ�ԭ�۶�ҩƷ�������
                    objBillIncome.ԭ�� = Nvl(!ԭ��, 0)
                    objBillIncome.������ĿID = Nvl(!������ID, 0)
                    objBillIncome.������Ŀ = Nvl(!������Ŀ)
                    objBillIncome.�վݷ�Ŀ = Nvl(!�ַ�Ŀ)
                    
                    'Ӧ�ս��=����*����*����
                    If !�Ƿ��� = 1 And (InStr(",5,6,7,", !�շ����) > 0 Or !�շ���� = "4" And Nvl(!��������, 0) = 1) Then
                        objBillIncome.Ӧ�ս�� = dblPrice '��֤Ӧ�ս�������۽��û�����
                    Else
                        objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * objBillDetail.���� * objBillDetail.����
                    End If
                    
                    '�������������ü���(����������Ŀ)
                    If Nvl(!���ӱ�־, 0) = 1 And Nvl(!�շ����) = "F" Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * Nvl(!�����շ���, 100) / 100
                    End If
                    
                    '�Ӱ�����ʼ���
                    If Nvl(!�Ӱ��־, 0) = 1 And Nvl(!�Ӱ�Ӽ�, 0) = 1 Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * (1 + Nvl(!�Ӱ�Ӽ���, 0) / 100)
                    End If
                    objBillIncome.Ӧ�ս�� = Format(objBillIncome.Ӧ�ս��, gstrDec)
                    
                    '����ʵ�ս��
                    If bln��� Then
                        objBillIncome.ʵ�ս�� = 0
                    Else
                        If Nvl(!���ηѱ�, 0) = 1 Then
                            objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                        Else
                            'ʹ��ԭ���Ķ�̬�ѱ�
                            objBillIncome.ʵ�ս�� = ActualMoney(objBill.�ѱ�, !������ID, objBillIncome.Ӧ�ս��, objBillDetail.�շ�ϸĿID, _
                                objBillDetail.ִ�в���ID, !���� * !����, IIF(Nvl(!�Ӱ��־, 0) = 1 And Nvl(!�Ӱ�Ӽ�, 0) = 1, Nvl(!�Ӱ�Ӽ���, 0) / 100, 0))
                        End If
                    End If
                    
                    With objBillIncome
                        objBillDetail.InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
                    End With
                    
                    '�ж���һ����¼�Ƿ����ڵ�ǰ��
                    blnDo = False
                    int��� = !���
                    .MoveNext
                    If Not .EOF Then blnDo = (int��� = !���)
                    i = i + 1
                Loop While blnDo And Not .EOF
               
                With objBillDetail
                    objBill.Details.Add .InComes, .Detail, .�շ�ϸĿID, .���, .��������, .�շ����, .���㵥λ, .����, .����, .���ӱ�־, .ִ�в���ID, .��ҩ����, , , , .ժҪ
                End With
            Loop
        End With
    End If
    
    Set ImportBill = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillMoney(strNO As String) As Currency
'���ܣ���ȡһ���������ʵ��ĵ��ݽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Sum(ʵ�ս��) as ��� From ���˷��ü�¼ Where NO=[1] And ��¼����=2 And ��¼״̬ IN(0,1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then GetBillMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(lng����ID As Long) As Currency
'����:��ȡָ�����˵ļ��ʻ��۵����ϼ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ� From ���˷��ü�¼ Where ��¼״̬=0 And ���ʷ���=1 And ����ID=" & lng����ID
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlExpense")
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!���۷��úϼ�
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAuditRecord(lng����ID As Long, lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡָ�����˵ķ���������Ŀ
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ��ĿId From ����������Ŀ Where ����ID=[1] And ��ҳID=[2]"
    Set GetAuditRecord = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMoneyInfo(lng����ID As Long, Optional curModiMoney As Currency, Optional blnInsure As Boolean) As ADODB.Recordset
'���ܣ���ȡָ�����˵�ʣ���
'������blnInsure=�Ƿ��ſ�ҽ�����˵�Ԥ�����
    Dim rsTmp As New ADODB.Recordset
    Dim blnҽ�� As Boolean, lng��ҳID As Long
    Dim strSQL As String
        
    On Error GoTo errH
    
    If blnInsure Then
        strSQL = "Select A.����,A.��ҳID From ������ҳ A,������Ϣ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� And B.����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID)
        If Not rsTmp.EOF Then
            blnҽ�� = Not IsNull(rsTmp!����)
            lng��ҳID = rsTmp!��ҳID
        End If
    End If
    
    strSQL = "Select Nvl(�������,0) as �������,Nvl(Ԥ�����,0) as Ԥ�����" & _
            " From ������� Where ����=1 And ����ID=" & lng����ID
    
    If curModiMoney <> 0 Then   '����Ҫ��Union��ʽ,���ֱ��ȥ��,�ڲ�������޼�¼ʱ,���᷵�ؼ�¼
        strSQL = strSQL & " Union All " & " Select -1* " & curModiMoney & " as �������,0 as Ԥ����� From Dual"
        strSQL = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ����� From (" & strSQL & ")"
    End If
            
    '���Ϊҽ��סԺ���ˣ����ڷ���������ſ�Ԥ���еķ���(���ڱ���)
    If blnInsure And blnҽ�� Then
        strSQL = strSQL & " Union All " & _
            " Select -1*Nvl(Sum(���),0) as �������,0 as Ԥ�����" & _
            " From ����ģ����� Where ����ID=[1] And ��ҳID=[2]"
        strSQL = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ����� From (" & strSQL & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiUnit(lngPatiID As Long) As Long
'���ܣ����ز�����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.��ǰ����ID From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.סԺ����=B.��ҳID And A.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngPatiID)
    If Not rsTmp.EOF Then GetPatiUnit = Nvl(rsTmp!��ǰ����ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AdjustCpt(lngID As Long)
'���ܣ�ҩƷ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    
    strSQL = _
        "Select ID From �շѼ�Ŀ" & _
        " Where ((Sysdate Between ִ������ and ��ֹ����) Or (Sysdate>=ִ������ And ��ֹ���� is NULL))" & _
        " And Nvl(�䶯ԭ��,0)=0 And �շ�ϸĿID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngID)
    Do While Not rsTmp.EOF
        strSQL = "zl_ҩƷ�շ���¼_Adjust(" & rsTmp!ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "mdlExpense")
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function BillisZeroLog(ByVal strNO As String) As Boolean
'���ܣ��ж�ָ�������Ƿ�������ķ��õǼ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long

    On Error GoTo errH

    strSQL = "Select ʵ�ս�� From ���˷��ü�¼ Where ��¼״̬ In(0,1,3) And ��¼����=2 And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    BillisZeroLog = True
    For i = 1 To rsTmp.RecordCount
        If Nvl(rsTmp!ʵ�ս��, 0) <> 0 Then
            BillisZeroLog = False: Exit For
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiCanBilling(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
'���ܣ����ָ�������Ƿ�������Ȩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") > 0 _
        And InStr(strPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
        Exit Function
    End If
    
    strSQL = "Select A.����,B.��Ժ����,B.״̬,X.�������" & _
        " From ������Ϣ A,������ҳ B,������� X" & _
        " Where A.����ID=B.����ID And A.����ID=X.����ID(+)" & _
        " And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!��Ժ����) And Nvl(rsTmp!״̬, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") = 0 Then
            If Nvl(rsTmp!�������, 0) <> 0 Then
                strMsg = """" & rsTmp!���� & """�ķ���δ���壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If InStr(strPrivs, "��Ժ����ǿ�Ƽ���") = 0 Then
            If Nvl(rsTmp!�������, 0) = 0 Then
                strMsg = """" & rsTmp!���� & """�ķ����ѽ��壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillIdentical(ByVal strNO As String) As Boolean
'���ܣ��ж�ָ���ļ��ʵ����е�״̬�Ƿ�һ��,���Ƿ�ͬʱ������˺�δ��˵�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    BillIdentical = True
    
    On Error GoTo errH

    strSQL = _
        " Select Count(Distinct �Ǽ�ʱ��) as ʱ����," & _
        " Sum(Decode(��¼״̬,0,1,0)) as δ���," & _
        " Sum(Decode(��¼״̬,0,0,1)) as �����" & _
        " From ���˷��ü�¼" & _
        " Where ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!δ���, 0) <> 0 And Nvl(rsTmp!�����, 0) <> 0 Then
            BillIdentical = False
        ElseIf Nvl(rsTmp!ʱ����, 0) > 1 Then
            BillIdentical = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckValidity(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, Optional ByVal blnAsk As Boolean = True) As Boolean
'���ܣ�����������ϵ����Ч���Ƿ����
'˵����blnAsk=��ʾ�Ƿ�ѯ���Ƿ����,����Ϊ����
    Dim rsTmp As New ADODB.Recordset
    Dim curDate As Date, minDate As Date
    Dim strSQL As String, strName As String
    
    CheckValidity = True
    
    '��һ���Բ��ϲ��ж�
    '��Ϊ���ܸ��������Ч�ڲ�ͬ,���Ҫ�õ�����������С��Ч��
    strSQL = _
        " Select C.����,Nvl(B.����,0) as ����," & _
        " B.�������� as ���,B.���Ч��,Sysdate as ʱ��" & _
        " From �������� A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
        " Where A.����ID=B.ҩƷID And A.����ID=C.ID And A.һ���Բ���=1" & _
        " And B.����=1 And Nvl(B.��������,0)>0 And A.���Ч�� is Not NULL" & _
        " And A.����ID=[1] And B.�ⷿID=[2]" & _
        " Order by Nvl(B.����,0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng�ⷿID)
    If Not rsTmp.EOF Then
        strName = rsTmp!����
        curDate = rsTmp!ʱ��
        minDate = CDate("3000-01-01")
            
        Do While Not rsTmp.EOF
            If rsTmp!���Ч�� < minDate Then
                minDate = rsTmp!���Ч��
            End If
            If Nvl(rsTmp!���, 0) < dbl���� Then
                dbl���� = dbl���� - Nvl(rsTmp!���, 0)
            Else
                dbl���� = 0
            End If
            If dbl���� = 0 Then Exit Do
            rsTmp.MoveNext
        Loop

        If curDate > minDate Then
            If blnAsk Then
                If MsgBox("��������""" & strName & """�����Ч��""" & Format(minDate, "yyyy-MM-dd") & """�ѹ���,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckValidity = False
                End If
            Else
                MsgBox "���ѣ�" & vbCrLf & vbCrLf & "��������""" & strName & """�����Ч��""" & Format(minDate, "yyyy-MM-dd") & """�ѹ��ڡ�", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveBilling(ByVal strNO As String, Optional ByVal blnALL As Boolean = True, Optional ByVal strTime As String) As Integer
'���ܣ��ж�һ�ż��ʵ�/���Ƿ��Ѿ�����
'������strNO=���ʵ��ݺ�,�������ＰסԺ
'      blnALL=�Ƿ�����ŵ������ݽ����ж�,����ֻ��δ���ʲ��ֽ����ж�(����ʱ)
'���أ�0-δ����,1=��ȫ������,2-�Ѳ��ֽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    
    On Error GoTo errH
        
    '��δ���ϵķ�����
    strSQL = _
        " Select ��� From (" & _
        " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���," & _
        " Avg(Nvl(����, 1) * ����) As ����" & _
        " From ���˷��ü�¼" & _
        " Where NO=[1] And ��¼����=2" & _
        " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    
    '��ÿ�еĽ������
    strSQL = _
        "Select Nvl(�۸񸸺�,���) as ���,Sum(Nvl(���ʽ��,0)) as ���ʽ��" & _
        " From ���˷��ü�¼" & _
        " Where NO=[1] And ��¼���� IN(2,12)" & _
        IIF(Not blnALL, " And Nvl(�۸񸸺�,���) IN(" & strSQL & ")", "") & _
        IIF(strTime <> "", " And �Ǽ�ʱ��=[2]", "") & _
        " Group by Nvl(�۸񸸺�,���)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO, CDate(IIF(strTime = "", "1990-01-01", strTime)))
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '��������
        rsTmp.Filter = "���ʽ��<>0"
        If rsTmp.EOF Then
            HaveBilling = 0 '�޽�����
        ElseIf rsTmp.RecordCount = lngTmp Then
            HaveBilling = 1 'ȫ�����ѽ���
        ElseIf rsTmp.RecordCount > 0 Then
            HaveBilling = 2 '�������ѽ���
        End If
    End If
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", intNum)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate(Format(rsTmp!����, "YYYY-MM-dd")) - CDate(Format(rsTmp!����, "YYYY") & "-01-01") + 1, "000")
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

Public Function CheckAdviceBalanceRoll(ByVal lng���ͺ� As Long, ByVal lngҽ��ID As Long, Optional ByVal blnBat As Boolean) As Boolean
'���ܣ�(סԺ)��Ҫ���˵�ҽ����Ӧ�ķ��õĽ���������м��(һ������һ��סԺ��)
'������blnBat=�Ƿ�Ҫ������������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intInsure As Integer
    
    On Error GoTo errH
        
    'ȡҪ���˵ļ���NO
    If blnBat Then
        strSQL = "Select Distinct NO From ����ҽ������ Where ��¼����=2 And ���ͺ�=[1]"
    Else
        strSQL = "Select Distinct A.NO From ����ҽ������ A,����ҽ����¼ B" & _
            " Where A.ҽ��ID=B.ID And A.��¼����=2 And A.���ͺ�=[1] And (B.ID=[2] Or B.���ID=[2])"
    End If
    'ȡ��ЩNO�Ľ������(�ǻ���δ����)
    strSQL = "Select A.NO,Nvl(A.�۸񸸺�,A.���) as ���,Sum(Nvl(A.���ʽ��,0)) as ���ʽ��" & _
        " From ���˷��ü�¼ A,(" & strSQL & ") B Where A.NO=B.NO And A.��¼���� IN(2,12) And A.��¼״̬=1" & _
        " Group by A.NO,Nvl(A.�۸񸸺�,A.���) Having Sum(Nvl(A.���ʽ��,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng���ͺ�, lngҽ��ID)
    If Not rsTmp.EOF Then
        strSQL = "Select A.���� From ������ҳ A,����ҽ����¼ B" & _
            " Where Rownum=1 And A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngҽ��ID)
        If Not rsTmp.EOF Then intInsure = Nvl(rsTmp!����, 0)
        If intInsure <> 0 Then '�ȶ�ҽ�������ƽ��м��
            If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , intInsure) Then
                MsgBox "�ò���Ϊҽ�����ˣ�Ҫ����ҽ���ķ��ͷ����д����ѽ��ʵķ��ã����ܻ��ˡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If gbytBillOpt <> 0 Then
            If gbytBillOpt = 1 Then
                If MsgBox("Ҫ����ҽ���ķ��ͷ����д����ѽ��ʵķ��ã�ȷʵҪ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf gbytBillOpt = 2 Then
                MsgBox "Ҫ����ҽ���ķ��ͷ����д����ѽ��ʵķ��ã����ܻ��ˡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    CheckAdviceBalanceRoll = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBalanceRevoke(ByVal lngҽ��ID As Long) As Boolean
'���ܣ�(����)��Ҫ���ϵ�ҽ����Ӧ�ķ��õĽ���������м��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng���ͺ� As Long
    
    If gbytBillOpt = 0 Then
        CheckAdviceBalanceRevoke = True
        Exit Function
    End If
    
    On Error GoTo errH
    
    'ҽ��IDΪ����ֵ������ҽ����һ�������˵�,�����޷��͡�
    strSQL = "Select Distinct ���ͺ� From ����ҽ������" & _
        " Where ҽ��ID IN(Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngҽ��ID)
    If rsTmp.EOF Then Exit Function
    lng���ͺ� = rsTmp!���ͺ�
    
    '����������"ZL_����ҽ����¼_����"
    strSQL = "Select A.NO,Nvl(A.�۸񸸺�,A.���) as ���,Sum(Nvl(A.���ʽ��,0)) as ���ʽ��" & _
        " From ���˷��ü�¼ A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ I" & _
        " Where A.NO=B.NO And A.��¼���� IN(2,12) And A.��¼״̬=1 And B.ҽ��ID=C.ID" & _
        " And B.��¼����=2 And C.������ĿID=I.ID And B.���ͺ�=[1] And (C.ID=[2] Or C.���ID=[2])" & _
        " And (" & _
            " A.�շ���� Not In ('5','6','7','E')" & _
            " Or A.�շ����='E' And I.�������� Not In ('2','3','4')" & _
            " Or A.�շ���� In ('5','6','7') And Nvl(A.ִ��״̬,0)=0" & _
            " Or Exists(Select ����ֵ From ϵͳ������ Where ������=68 And Nvl(����ֵ,0)=0)" & _
            " )" & _
        " Group by A.NO,Nvl(A.�۸񸸺�,A.���) Having Sum(Nvl(A.���ʽ��,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng���ͺ�, lngҽ��ID)
    If Not rsTmp.EOF Then
        If gbytBillOpt = 1 Then
            If MsgBox("Ҫ����ҽ���Ķ�Ӧ�����д����ѽ��ʵķ��ã�ȷʵҪ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        ElseIf gbytBillOpt = 2 Then
            MsgBox "Ҫ����ҽ���Ķ�Ӧ�����д����ѽ��ʵķ��ã��������ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckAdviceBalanceRevoke = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetExamineItem(ByVal strItems As String, ByVal lngMediCareID As Long) As ADODB.Recordset
'����:����ָ��������շ���ĿҪ�������ļ�¼��
'����:strItems-�շ�ϸĿID��,����:"2369,2367,2368"
'     lngMediCareID-����,����:901
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select A.�շ�ϸĿid" & vbNewLine & _
            "From ����֧����Ŀ A ,Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B" & vbNewLine & _
            "Where A.���� = [1] And A.Ҫ������ = 1 And A.�շ�ϸĿid = B.Column_Value"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngMediCareID, strItems)
    
    Set GetExamineItem = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRowByFeeItemID(ByRef ObjBillDetails As BillDetails, ByRef lngItemID As Long) As Long
'����:�����շ���ĿID�������ڵ����е��к�,������ظ���,ֻ���ص�һ��
    Dim i As Long
    
    For i = 1 To ObjBillDetails.Count
        If lngItemID = ObjBillDetails(i).�շ�ϸĿID Then
            GetRowByFeeItemID = i: Exit Function
        End If
    Next
End Function

Public Function CheckExamine(ByRef ObjBillDetails As BillDetails, ByRef rsMedAudit As ADODB.Recordset, ByRef lngMediCareID As Long) As Boolean
'����:���ݸ������շ���Ŀ���󼯺Ͳ���������Ŀ��¼�������Ӧ���շ���Ŀ�Ƿ���Ҫ����
    Dim i As Long, strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    For i = 1 To ObjBillDetails.Count
        strTmp = strTmp & "," & ObjBillDetails(i).�շ�ϸĿID
    Next
    Set rsTmp = GetExamineItem(Mid(strTmp, 2), lngMediCareID)
    
    strTmp = ""
    For i = 1 To rsTmp.RecordCount
        rsMedAudit.Filter = "��ĿID=" & rsTmp!�շ�ϸĿID
        If rsMedAudit.RecordCount = 0 Then strTmp = strTmp & "," & GetRowByFeeItemID(ObjBillDetails, rsTmp!�շ�ϸĿID)
        rsTmp.MoveNext
    Next
    
    If strTmp <> "" Then
        MsgBox "��" & Mid(strTmp, 2) & "���շ���ĿҪ������,��ǰ����δ����׼ʹ��!", vbInformation, gstrSysName
        CheckExamine = False: Exit Function
    End If
    CheckExamine = True
End Function

