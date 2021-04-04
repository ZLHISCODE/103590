Attribute VB_Name = "mdlDockExpense"
Option Explicit
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
    g��������ģ�� = 5
    g����˽��ģ�� = 6
End Enum
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public gobjInExse As Object
Private mlng���ű���ƽ������ As Long
Public grs������Ŀ As ADODB.Recordset
Public glngMainHwnd As Long

Public Function GetFeeKind() As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select ����, ����, ���� From �շ���Ŀ���"
    Set GetFeeKind = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ�շ����")
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

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
    
    If InStr(1, str�շ�ϸĿIDs, ",") = 0 Then
        strSQL = "" & _
        "   Select Distinct /*+ Rule*/ �շ�ϸĿID,Nvl(��������ID,0) as ��������ID,ִ�п���id " & _
        "   From �շ�ִ�п��� A " & _
        "   Where   A.�շ�ϸĿID  =[2] "
    Else
        strSQL = "" & _
        "   Select Distinct /*+ Rule*/ �շ�ϸĿID,Nvl(��������ID,0) as ��������ID,ִ�п���id " & _
        "   From �շ�ִ�п��� A," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        "   Where   A.�շ�ϸĿID  = j.Column_Value"
    End If
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡִ�п�����Ϣ", Replace(str�շ�ϸĿIDs, "'", ""), Val(str�շ�ϸĿIDs))
    If Not rsTmp.EOF Then Set GetServiceDept = rsTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub LoadPatientBaby(ByRef cboBaby As ComboBox, ByVal lngPatient As Long, lngPatientPage As Long)
    Dim rsTmp As ADODB.Recordset, i As Long
    
    cboBaby.Clear
    cboBaby.AddItem "0-���˱���"
    cboBaby.ItemData(cboBaby.NewIndex) = 0
    Call gobjControl.CboSetIndex(cboBaby.hWnd, 0)
    
    If lngPatient <> 0 Then
        Set rsTmp = GetPatientBaby(lngPatient, lngPatientPage)
        With rsTmp
            For i = 1 To .RecordCount
                If Not IsNull(!Ӥ������) Then
                    cboBaby.AddItem !��� & "-" & !Ӥ������
                Else
                    cboBaby.AddItem !��� & "-��" & !��� & "��Ӥ��"
                End If
                cboBaby.ItemData(cboBaby.NewIndex) = !���
                .MoveNext
            Next
        End With
    End If
End Sub

Public Function GetPatientBaby(ByVal lngPatient As Long, lngPatientPage As Long) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select ���, Ӥ������ From ������������¼ Where ����id = [1] And ��ҳID = [2]"
    On Error GoTo errH
    Set GetPatientBaby = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��������¼", lngPatient, lngPatientPage)

    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetDrugTotal(ByVal objBill As ExpenseBill, ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long, _
    Optional lng���� As Long = 0) As Double
'���ܣ���ȡ������ָ��ҩƷ��ͬһҩ�����е�������
    Dim i As Integer, dblCount As Double
    
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).�շ�ϸĿID = lngҩƷID _
            And objBill.Details(i).ִ�в���ID = lngҩ��ID And objBill.Details(i).Detail.���� = lng���� Then
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

Public Function Getҽ������(ByVal lng�շ�ϸĿID As Long, ByVal int���� As Integer) As String
'���ܣ���ȡָ���շ���Ŀ�ı��մ�������
'������
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select N.����" & _
        " From ����֧����Ŀ M,����֧������ N " & _
        " Where M.�շ�ϸĿID=[1] And M.����=[2] And M.����ID=N.ID"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lng�շ�ϸĿID, int����)
    If rsTmp.RecordCount > 0 Then Getҽ������ = rsTmp!����
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zl_Check��׼��Ŀ(ByVal objclsInsure As Object, ByVal intInsure As Integer, ByVal lng����ID As Long, Optional ByVal bln���� As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ҽ�������Ƿ���Ҫ�����׼��Ŀ
    '���:objInsure-������ҽ������
    '     intInsure-����
    '     lng����ID-����ID
    '     bln����-�Ƿ�����
    '����:
    '����:�����Ҫ�����׼��Ŀ,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:24862
    '����:2009-08-12 10:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zl_Check��׼��Ŀ = False
     If bln���� Then
        If objclsInsure.GetCapability(support���ﲡ�˲�����׼��Ŀ����, lng����ID, intInsure) = False Then zl_Check��׼��Ŀ = True
        Exit Function
     End If
    If objclsInsure.GetCapability(supportסԺ���˲�����׼��Ŀ����, lng����ID, intInsure) = False Then zl_Check��׼��Ŀ = True
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
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID)
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get����ְ��(lngҩƷID As Long) As String
'���ܣ�����ҩƷID��ȡ�䴦��ְ��
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    Get����ְ�� = "00"
    strSQL = "Select Nvl(B.����ְ��,'00') as ����ְ�� From ҩƷ��� A,ҩƷ���� B Where A.ҩ��ID=B.ҩ��ID And A.ҩƷID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngҩƷID)
    If Not rsTmp.EOF Then Get����ְ�� = rsTmp!����ְ��
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get��������(lngID As Long) As Double
'���ܣ���ȡָ��ҩƷ�Ĵ�������,�����۵�λ���ء�
'������lngID=ҩƷID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(A.��������,0) as ��������" & _
        " From ҩƷ���� A,ҩƷ��� B Where A.ҩ��ID=B.ҩ��ID And B.ҩƷID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngID)
    If Not rsTmp.EOF Then Get�������� = rsTmp!��������
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function ItemExistInsure(ByVal lng����ID As Long, ByVal lng�շ�ϸĿID As Long, ByVal int���� As Integer) As Boolean
'���ܣ��ж��շ���Ŀ�Ƿ������˱���֧����Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    If gclsInsure.GetCapability(support��������ҽ����Ŀ, lng����ID, int����) Then
        ItemExistInsure = True: Exit Function
    End If
    
    strSQL = "Select 1 From ����֧����Ŀ Where �շ�ϸĿID=[1] And ����=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng�շ�ϸĿID, int����)
    ItemExistInsure = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    
    strSQL = "Select  /*+ RULE */  A.ҩƷID,A.����ϵ��,B.���㵥λ as ������λ" & _
        " From ҩƷ��� A,������ĿĿ¼ B," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        " Where A.ҩ��ID=B.ID And A.ҩƷID  = j.Column_Value"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strItemIDs)
    
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
                                FormatEx(dbl����, 5) & rsTmp!������λ & "(" & FormatEx(dblTime, 5) & IIf(blnҩ����λ, tmpDetail.Detail.ҩ����λ, tmpDetail.Detail.���㵥λ) & ") ������������ " & _
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
        " Nvl(Sum(A.��������),0)" & IIf(blnҩ����λ, "/Nvl(C." & strҩ����װ & ",1)", "") & " as ���" & _
        " From ҩƷ��� A,(" & strSQL & ") B,ҩƷ��� C" & _
        " Where A.�ⷿID=B.ID And A.ҩƷID=C.ҩƷID" & _
        " And ((A.Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
        " Or (Nvl(C.ҩ������,0)=0 And A.�ⷿID IN(" & strSQL2 & ")))" & _
        " And A.����=1 And A.ҩƷID=[2]" & _
        " Group by B.����,B.����,A.�ⷿID,Nvl(C." & strҩ����װ & ",1)" & _
        " Having Sum(Nvl(A.��������,0))<>0" & _
        " Order By B.����"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", str����, lngҩƷID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!���� & ":" & rsTmp!���
        rsTmp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    GetStockInfo = strSQL
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function OverTime() As Boolean
'���ܣ��жϵ�ǰ�Ƿ��ڼӰ�ʱ�䷶Χ��
'���أ���-��ǰ���ڼӰ�ʱ����,��-������
    Dim str���� As String, str���� As String
    Dim DateBegin As Date, DateEnd As Date
    Dim curTime As Date
    
    str���� = gobjDatabase.GetPara(1, glngSys): str���� = gobjDatabase.GetPara(2, glngSys)
    curTime = CDate(Format(gobjDatabase.Currentdate, "HH:MM:SS"))
    
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

Public Function GetBillRows(str���ݺ� As String, int��¼���� As Integer, int������Դ As Integer) As Integer
'���ܣ���ȡһ�ŷ��õ�����δ���ϵķ�������
'������int��¼����=1-�շ�(����),2-����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIf(int��¼���� = 1 Or (int��¼���� = 2 And int������Դ = 1), "������ü�¼", "סԺ���ü�¼")

    On Error GoTo errH
    
    
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    strSQL = _
        " Select ���,Sum(����) as ʣ������" & _
        " From (" & _
        " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���," & _
        " Avg(Nvl(����, 1) * ����) As ����" & _
        " From " & strTab & _
        " Where NO=[1] And ��¼����=[2]" & _
        " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", str���ݺ�, int��¼����)
    If Not rsTmp.EOF Then GetBillRows = rsTmp.RecordCount
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    
    strSQL = "Select B.���� From סԺ���ü�¼ A,������ҳ B" & _
        " Where A.��¼����=2 And A.��¼״̬ IN(0,1,3) And B.���� is Not NULL" & _
        " And A.NO=[1] And A.����ID=B.����ID And A.��ҳID=B.��ҳID"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then BillExistInsure = rsTmp!����
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function BillExistDelete(strNO As String, int��¼���� As Integer, int������Դ As Integer) As Boolean
'���ܣ��ж�ָ�������Ƿ����(����)�˷ѻ����ʵ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIf(int��¼���� = 1 Or (int��¼���� = 2 And int������Դ = 1), "������ü�¼", "סԺ���ü�¼")
    
    On Error GoTo errH
    
    strSQL = "Select NO From " & strTab & " Where NO=[1] And ��¼����=[2] And ��¼״̬=2"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "BillExistDelete", strNO, int��¼����)
    BillExistDelete = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetInsureName(intInsure As Integer) As String
'���ܣ����ݱ��������Ż�ȡ�����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ������� Where ���=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", intInsure)
    If Not rsTmp.EOF Then GetInsureName = Nvl(rsTmp!����)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
        " From ���ű� A,��������˵�� B," & IIf(bytType = 0, "ҩƷ������", "���ϳ�����") & " C" & _
        " Where B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� " & IIf(bytType = 0, "IN('��ҩ��','��ҩ��','��ҩ��')", "='���ϲ���'") & _
        " And C.�ⷿID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "GetStockCheck")
    For i = 1 To rsTmp.RecordCount
        colStock.Add Nvl(rsTmp!��鷽ʽ, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
        " Select /*+ RULE */  A.����,Count(Distinct A.��ĿID) as ������" & _
        " From ���ƻ�����Ŀ A,ҩƷ��� B," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        " Where A.��ĿID=B.ҩ��ID And B.ҩƷID  = j.Column_Value" & _
        " Having Count(Distinct A.��ĿID)>1  " & _
        "  Group by A.����"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strIDs)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strGroup = strGroup & "," & rsTmp!����
            rsTmp.MoveNext
        Next
        strGroup = Mid(strGroup, 2)
        
        For i = 0 To UBound(Split(strGroup, ","))
            strSQL = _
            "Select /*+ RULE */   Distinct C.����,C.����,D.����,D.����,D.���" & _
            " From ҩƷ��� A,������ĿĿ¼ B,���ƻ�����Ŀ C,�շ���ĿĿ¼ D," & _
            "          (Select Column_Value From Table(Cast(f_num2list([2]) As Zltools.t_Numlist ))) J " & _
            " Where A.ҩ��ID=B.ID And B.ID=C.��ĿID And A.ҩƷID=D.ID" & _
            "           And C.����=[1]" & _
            "           And A.ҩƷID  = j.Column_Value" & _
            " Order by C.����,C.����,D.����"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", Val(Split(strGroup, ",")(i)), strIDs)
            
            If Not rsTmp.EOF Then
                rsTmp.Filter = "����=1"
                If rsTmp.RecordCount > 1 Then
                    k = k + 1
                    strInfo = strInfo & vbCrLf & "�� " & k & " ��(��������)��" & vbCrLf
                    For j = 1 To rsTmp.RecordCount
                        strInfo = strInfo & "[" & rsTmp!���� & "]" & rsTmp!���� & IIf(IsNull(rsTmp!���), "", "(" & rsTmp!��� & ")") & "                 " & vbCrLf
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Filter = "����=2"
                If rsTmp.RecordCount > 1 Then
                    blnStop = True
                    k = k + 1
                    strInfo = strInfo & vbCrLf & "�� " & k & " ��(�������)��" & vbCrLf
                    For j = 1 To rsTmp.RecordCount
                        strInfo = strInfo & "[" & rsTmp!���� & "]" & rsTmp!���� & IIf(IsNull(rsTmp!���), "", "(" & rsTmp!��� & ")") & "                 " & vbCrLf
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function ImportBill(ByVal int��Դ As Integer, ByVal str���ݺ� As String, _
    ByVal int��¼���� As Integer, Optional ByVal bln��� As Boolean, _
    Optional ByVal strҩƷ�۸�ȼ� As String, _
    Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ�۸�ȼ� As String) As ExpenseBill
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
    Dim intCurNo As Integer
    Dim int��� As Integer, blnDo As Boolean, i As Integer
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim dblAllTime As Double
    
    Dim colSerial As New Collection '���ڴ����������
    Dim strSQL As String
    Dim strTab As String
    
    Dim lng��ҩ�� As Long, lng��ҩ�� As Long, lng��ҩ�� As Long, strҩ�� As String
    Dim blnҩ����λ As Boolean, strҩ����λ As String, strҩ����װ As String
    Dim strWherePriceGrade As String
        
    strTab = IIf(int��¼���� = 1 Or (int��¼���� = 2 And int��Դ = 1), "������ü�¼", "סԺ���ü�¼")
    
    '�۸�ȼ�
    If strҩƷ�۸�ȼ� <> "" Or str���ļ۸�ȼ� <> "" Or str��ͨ�۸�ȼ� <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [4])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [5])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And d.�۸�ȼ� = [6])" & vbNewLine & _
            "            Or (d.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From �շѼ�Ŀ" & vbNewLine & _
            "                                Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [4])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [5])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And �۸�ȼ� = [6])))))"
    Else
        strWherePriceGrade = " And d.�۸�ȼ� Is Null "
    End If
    
    'ȱʡҩ��
    lng��ҩ�� = Val(gobjDatabase.GetPara(IIf(int��Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, pҽ�����ѹ���))
    lng��ҩ�� = Val(gobjDatabase.GetPara(IIf(int��Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, pҽ�����ѹ���))
    lng��ҩ�� = Val(gobjDatabase.GetPara(IIf(int��Դ = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, pҽ�����ѹ���))
    
    'ҩƷ��λ
    blnҩ����λ = Val(gobjDatabase.GetPara("ҩƷ��λ", glngSys, pҽ�����ѹ���)) <> 0
    If int��Դ = 1 Then
        strҩ����λ = "���ﵥλ": strҩ����װ = "�����װ"
    Else
        strҩ����λ = "סԺ��λ": strҩ����װ = "סԺ��װ"
    End If
        
    '------------------------------------------------------------------------------------------
    '�շѼ�Ŀ����:�¼���۸�,����ж���۸�,��һ���շ�ϸĿID�оͻ��ж��������ͬ�ļ�¼
    '�۸񸸺� is NULL:ֻȡÿ���շ�ϸĿID�ĵ�һ��(ҩƷֻ��һ��),��ΪҪ����۸�
        
    'ʹ��ָ����ҩ����ȡ��ȷ�Ŀ��
    strҩ�� = "Decode(A.�շ����,'5'," & IIf(lng��ҩ�� <> 0, lng��ҩ��, "A.ִ�в���ID") & "," & _
        "'6'," & IIf(lng��ҩ�� <> 0, lng��ҩ��, "A.ִ�в���ID") & "," & _
        "'7'," & IIf(lng��ҩ�� <> 0, lng��ҩ��, "A.ִ�в���ID") & ",A.ִ�в���ID)"
    
    'ҩ�������������ҩƷ����Ч��
    strSQL = _
    " Select X.ҩƷID,W.����ID,W.��������,A.��� As ���,A.��������,A.NO,A.��¼����,A.��¼״̬," & IIf(strTab = "סԺ���ü�¼", "A.�ಡ�˵�", " 0 as �ಡ�˵�") & ",A.Ӥ����,A.�ѱ�,A.����,A.�Ա�,A.����," & _
            IIf(strTab = "סԺ���ü�¼", "A.����,A.���˲���ID,A.��ҳID", "A.���ʽ as ����,0 as ���˲���ID,0 as ��ҳID") & _
    "       ,A.��ʶ��,A.����ID,A.���˿���ID,A.��������ID,A.�����־,A.�Ӱ��־," & _
    "       A.���ӱ�־,A.�շ����,A.�շ�ϸĿID,A.��ҩ����,Nvl(����,1) as ����,Nvl(A.����,0) as ����," & _
    "       A.��׼���� As ��׼����," & strҩ�� & " as ִ�в���ID,A.������,A.������,A.����Ա���,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��,A.ժҪ," & _
    "       B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(F.����,B.����) as ����,E1.���� as ��Ʒ��,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
    "       B.���ηѱ�,B.˵��,B.ִ�п���,B.�������,Nvl(A.��������,B.��������) ��������,D.�ּ�,D.ԭ��,D.ȱʡ�۸�,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
    "       E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
    "       Decode(A.�շ����,'4',1,X." & strҩ����װ & ") as ҩ����װ," & _
    "       Decode(A.�շ����,'4',B.���㵥λ,X." & strҩ����λ & ") as ҩ����λ," & _
    "       Decode(A.�շ����,'4',Nvl(W.���÷���,0),Nvl(X.ҩ������,0)) as ����,Nvl(Y.���,0) As ���,B.¼������" & _
    " From " & strTab & " A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E,�շ���Ŀ���� F,�շ���Ŀ���� E1,�������� W,ҩƷ��� X," & _
    "       (Select A.ҩƷID,A.�ⷿID,Sum(Nvl(A.��������,0)) as ��� From ҩƷ��� A" & _
    "        Where A.����=1 And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
    "               And A.ҩƷID IN(Select �շ�ϸĿID From " & strTab & " Where ��¼����=[2] And ��¼״̬ IN(0,1,3) And NO=[1])" & _
                ""
    strSQL = strSQL & _
    "        Group by A.ҩƷID,A.�ⷿID) Y" & _
    " Where A.��¼����=[2] And A.��¼״̬ IN(0,1,3) And A.NO=[1]" & _
    "       And A.�۸񸸺� Is Null And A.�շ�ϸĿID=B.ID And A.�շ�ϸĿID=D.�շ�ϸĿID" & _
    "       And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is NULL)" & _
    "       And A.�շ����=C.���� And A.�շ�ϸĿID=X.ҩƷID(+) And A.�շ�ϸĿID=W.����ID(+) And D.������ĿID=E.ID" & _
    "       And A.�շ�ϸĿID=Y.ҩƷID(+) And " & strҩ�� & "=Y.�ⷿID(+)" & _
    "       And A.�շ�ϸĿID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=[3]" & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
    "       And ((Sysdate Between D.ִ������ And D.��ֹ����) Or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & strWherePriceGrade

    strSQL = "Select * From (" & strSQL & ") Order by ���"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", str���ݺ�, int��¼����, IIf(gSysPara.bytҩƷ������ʾ = 1, 3, 1), _
        strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ�۸�ȼ�)
    
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
                    objBill.����ID = Nvl(!���˿���id, 0)
                    objBill.���� = Nvl(!����)
                    objBill.�Ա� = Nvl(!�Ա�)
                    objBill.���� = Nvl(!����)
                    objBill.�ѱ� = Nvl(!�ѱ�)
                    objBill.��ʶ�� = Nvl(!��ʶ��)
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
                objBillDetail.���㵥λ = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                
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
                objBillDetail.Detail.��Ʒ�� = Nvl(!��Ʒ��)
                objBillDetail.Detail.���ηѱ� = Nvl(!���ηѱ�, 0) <> 0
                objBillDetail.Detail.˵�� = Nvl(!˵��)
                objBillDetail.Detail.ִ�п��� = Nvl(!ִ�п���, 0)
                objBillDetail.Detail.������� = Val(Nvl(!�������))
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
                    If InStr(",5,6,7,", !�շ����) > 0 Or (!�շ���� = "4" And Nvl(!��������, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        'ʱ��ҩƷ����۸�(�����ɲ�����)
                        dblAllTime = !���� * !���� '�������ۼ�����
                        If dblAllTime <> 0 Or Nvl(!�Ƿ���, 0) = 0 Then
                            Set rsPrice = gobjDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                        "��ȡҩƷ��ǰ�ۼ�", CLng(!�շ�ϸĿID), objBillDetail.ִ�в���ID, dblAllTime)
                            If rsPrice.EOF Then
                                '��ȡ�۸�ʧ��
                                If !�շ���� = "4" Then
                                    MsgBox "��������""" & Nvl(!����) & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                                Else
                                    MsgBox "ҩƷ""" & Nvl(!����) & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                                End If
                                objBillIncome.��׼���� = 0
                            Else
                                strPrice = Nvl(rsPrice!Price) & "|||"
                                varPrice = Split(strPrice, "|")
                                objBillIncome.��׼���� = Val(varPrice(0))
                                dblʣ������ = Val(varPrice(2))
                                
                                If dblʣ������ <> 0 And Nvl(!�Ƿ���, 0) = 1 Then
                                    '����δ�ֽ����
                                    If !�շ���� = "4" Then
                                        MsgBox "ʱ����������""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    Else
                                        MsgBox "ʱ��ҩƷ""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.��׼���� = 0
                                End If
                            End If
                        Else
                            objBillIncome.��׼���� = 0
                        End If
                    ElseIf Nvl(!�Ƿ���, 0) = 1 Then
                        If Abs(!��׼����) > Abs(Val(Nvl(!�ּ�))) Then
                            objBillIncome.��׼���� = Val(Nvl(!ȱʡ�۸�))
                        Else
                            objBillIncome.��׼���� = !��׼����
                        End If
                    Else
                        objBillIncome.��׼���� = !�ּ�
                    End If
                                        
                    If InStr(",5,6,7,", !�շ����) > 0 And blnҩ����λ Then
                        objBillIncome.��׼���� = Format(objBillIncome.��׼���� * Nvl(!ҩ����װ, 1), gSysPara.Price_Decimal.strFormt_VB)
                    Else
                        objBillIncome.��׼���� = Format(objBillIncome.��׼����, gSysPara.Price_Decimal.strFormt_VB)
                    End If
                    objBillIncome.�ּ� = Nvl(!�ּ�, 0) '�ּ�ԭ�۶�ҩƷ�������
                    objBillIncome.ԭ�� = Nvl(!ԭ��, 0)
                    objBillIncome.������ĿID = Nvl(!������ID, 0)
                    objBillIncome.������Ŀ = Nvl(!������Ŀ)
                    objBillIncome.�վݷ�Ŀ = Nvl(!�ַ�Ŀ)
                    
                    'Ӧ�ս��=����*����*����
                    objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * objBillDetail.���� * objBillDetail.����
                    
                    '�������������ü���(����������Ŀ)
                    If Nvl(!���ӱ�־, 0) = 1 And Nvl(!�շ����) = "F" Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * Nvl(!�����շ���, 100) / 100
                    End If
                    
                    '�Ӱ�����ʼ���
                    If Nvl(!�Ӱ��־, 0) = 1 And Nvl(!�Ӱ�Ӽ�, 0) = 1 Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * (1 + Nvl(!�Ӱ�Ӽ���, 0) / 100)
                    End If
                    objBillIncome.Ӧ�ս�� = Format(objBillIncome.Ӧ�ս��, gSysPara.Money_Decimal.strFormt_VB)
                    
                    '����ʵ�ս��
                    If bln��� Then
                        objBillIncome.ʵ�ս�� = 0
                    Else
                        If Nvl(!���ηѱ�, 0) = 1 Then
                            objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                        Else
                            'ʹ��ԭ���Ķ�̬�ѱ�
                            objBillIncome.ʵ�ս�� = ActualMoney(objBill.�ѱ�, !������ID, objBillIncome.Ӧ�ս��, objBillDetail.�շ�ϸĿID, _
                                objBillDetail.ִ�в���ID, !���� * !����, IIf(Nvl(!�Ӱ��־, 0) = 1 And Nvl(!�Ӱ�Ӽ�, 0) = 1, Nvl(!�Ӱ�Ӽ���, 0) / 100, 0))
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function ImportStuffBill(ByVal int��Դ As Integer, ByVal str���ݺ� As String, _
    ByVal int��¼���� As Integer, ByVal lng����ⷿID As Long) As ExpenseBill
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
    Dim intCurNo As Integer
    Dim int��� As Integer, blnDo As Boolean, i As Integer
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim dblAllTime As Double

    Dim colSerial As New Collection '���ڴ����������
    Dim strSQL As String, strStock As String
    Dim strTab As String
    Dim byt���� As Byte
    
    strTab = IIf(int��¼���� = 1 Or (int��¼���� = 2 And int��Դ = 1), "������ü�¼", "סԺ���ü�¼")
    '------------------------------------------------------------------------------------------
    '�շѼ�Ŀ����:�¼���۸�,����ж���۸�,��һ���շ�ϸĿID�оͻ��ж��������ͬ�ļ�¼
    '�۸񸸺� is NULL:ֻȡÿ���շ�ϸĿID�ĵ�һ��(ҩƷֻ��һ��),��ΪҪ����۸�
        
    byt���� = IIf(int��¼���� = 1, 24, 25)
    strStock = _
        " Select A.����ID,Max( A.ҩƷID) as ҩƷID,Max(A.����) as ����,Max(A.��Ʒ����) as ��Ʒ���� ,Max(A.�ڲ�����) as �ڲ����� " & _
        " From ҩƷ�շ���¼ A" & _
        " Where A.NO=[1]  And  ���� =[5] And MOD(A.��¼״̬,3) in (0,1)" & _
        " Group by A.����ID "
        
    strStock = "" & _
    "   Select A.����ID,A.����,A.��Ʒ����,A.�ڲ�����,sum(b.��������) as �������� " & _
    "   From (" & strStock & ") A,ҩƷ��� B " & _
    "   Where A.ҩƷid=b.ҩƷID(+) And B.�ⷿID(+)=[4] " & _
    "   Group by A.����ID,A.����,A.��Ʒ����,A.�ڲ�����"
    
 
    'ҩ�������������ҩƷ����Ч��
    strSQL = _
    " Select X.ҩƷID,W.����ID,W.��������,A.��� As ���,A.��������,A.NO,A.��¼����,A.��¼״̬," & IIf(strTab = "סԺ���ü�¼", "A.�ಡ�˵�", " 0 as �ಡ�˵�") & ",A.Ӥ����,A.�ѱ�,A.����,A.�Ա�,A.����," & _
            IIf(strTab = "סԺ���ü�¼", "A.����,A.���˲���ID,A.��ҳID", "A.���ʽ as ����,0 as ���˲���ID,0 as ��ҳID") & _
    "       ,A.��ʶ��,A.����ID,A.���˿���ID,A.��������ID,A.�����־,A.�Ӱ��־," & _
    "       A.���ӱ�־,A.�շ����,A.�շ�ϸĿID,A.��ҩ����,Nvl(����,1) as ����,Nvl(A.����,0) as ����," & _
    "       A.��׼���� As ��׼����,A.ִ�в���ID,A.������,A.������,A.����Ա���,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��,A.ժҪ," & _
    "       B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(F.����,B.����) as ����,E1.���� as ��Ʒ��,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
    "       B.���ηѱ�,B.˵��,B.ִ�п���,B.�������,Nvl(A.��������,B.��������) ��������,D.�ּ�,D.ԭ��,D.ȱʡ�۸�,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
    "       E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
    "       1 as ҩ����װ, B.���㵥λ as ҩ����λ, Nvl(W.���÷���,0) as ����, nvl(y.����,0) as ����,y.��Ʒ����,y.�ڲ����� ,Nvl(Y.��������,0) As ���,B.¼������" & _
    " From " & strTab & " A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E,�շ���Ŀ���� F,�շ���Ŀ���� E1,�������� W,ҩƷ��� X," & _
    "       (" & strStock & ") Y" & _
    " Where A.��¼����=[2] And A.��¼״̬ IN(0,1,3) And A.NO=[1]" & _
    "       And A.�۸񸸺� Is Null And A.�շ�ϸĿID=B.ID And A.�շ�ϸĿID=D.�շ�ϸĿID" & _
    "       And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is NULL)" & _
    "       And A.�շ����=C.���� And A.�շ�ϸĿID=X.ҩƷID(+) And A.�շ�ϸĿID=W.����ID(+) And D.������ĿID=E.ID" & _
    "       And A.ID=y.����ID(+) " & _
    "       And A.�շ�ϸĿID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=[3]" & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
    "       And ((Sysdate Between D.ִ������ And D.��ֹ����) Or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"

    strSQL = "Select * From (" & strSQL & ") Order by ���"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", str���ݺ�, int��¼����, IIf(gSysPara.bytҩƷ������ʾ = 1, 3, 1), lng����ⷿID, byt����)
    
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
                    objBill.����ID = Nvl(!���˿���id, 0)
                    objBill.���� = Nvl(!����)
                    objBill.�Ա� = Nvl(!�Ա�)
                    objBill.���� = Nvl(!����)
                    objBill.�ѱ� = Nvl(!�ѱ�)
                    objBill.��ʶ�� = Nvl(!��ʶ��)
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
                objBillDetail.���㵥λ = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                
                objBillDetail.���� = Nvl(!����, 1)
                objBillDetail.���� = Nvl(!����, 0)
                
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
                objBillDetail.Detail.��� = Nvl(!���, 0)
                objBillDetail.Detail.¼������ = Val("" & !¼������)
                
                objBillDetail.Detail.�Ӱ�Ӽ� = Nvl(!�Ӱ�Ӽ�, 0) <> 0
                objBillDetail.Detail.��� = Nvl(!���)
                objBillDetail.Detail.������� = Nvl(!�������)
                objBillDetail.Detail.���� = Nvl(!����)
                objBillDetail.Detail.��Ʒ�� = Nvl(!��Ʒ��)
                objBillDetail.Detail.���ηѱ� = Nvl(!���ηѱ�, 0) <> 0
                objBillDetail.Detail.˵�� = Nvl(!˵��)
                objBillDetail.Detail.ִ�п��� = Nvl(!ִ�п���, 0)
                objBillDetail.Detail.������� = Val(Nvl(!�������))
                objBillDetail.Detail.���� = Nvl(!��������)
                objBillDetail.Detail.����ְ�� = Get����ְ��(objBillDetail.Detail.ID)
                
                objBillDetail.Detail.ҩ��ID = Nvl(!ҩ��ID, 0)
                objBillDetail.Detail.��� = Nvl(!�Ƿ���, 0) <> 0
                objBillDetail.Detail.���� = Nvl(!����, 0) <> 0
                objBillDetail.Detail.�������� = Nvl(!��������, 0) = 1
                objBillDetail.Detail.Ҫ������ = 0
                objBillDetail.Detail.���� = Nvl(!����, 0)
                objBillDetail.Detail.��Ʒ���� = Nvl(!��Ʒ����)
                objBillDetail.Detail.�ڲ����� = Nvl(!�ڲ�����)
                '����۸񲿷�=====================================================
                Set objBillDetail.InComes = New BillInComes
                Do
                    '�������еļ۸��������¼���
                    If InStr(",5,6,7,", !�շ����) > 0 Or (!�շ���� = "4" And Nvl(!��������, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        'ʱ��ҩƷ����۸�(�����ɲ�����)
                        dblAllTime = !���� * !���� '�������ۼ�����
                        If dblAllTime <> 0 Or Nvl(!�Ƿ���, 0) = 0 Then
                            Set rsPrice = gobjDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                        "��ȡҩƷ��ǰ�ۼ�", CLng(!�շ�ϸĿID), objBillDetail.ִ�в���ID, dblAllTime)
                            If rsPrice.EOF Then
                                '��ȡ�۸�ʧ��
                                MsgBox "��������""" & Nvl(!����) & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                                objBillIncome.��׼���� = 0
                            Else
                                strPrice = Nvl(rsPrice!Price) & "|||"
                                varPrice = Split(strPrice, "|")
                                objBillIncome.��׼���� = Val(varPrice(0))
                                dblʣ������ = Val(varPrice(2))
                                
                                If dblʣ������ <> 0 And Nvl(!�Ƿ���, 0) = 1 Then
                                    '����δ�ֽ����
                                    MsgBox "ʱ����������""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    objBillIncome.��׼���� = 0
                                End If
                            End If
                        Else
                            objBillIncome.��׼���� = 0
                        End If
                    ElseIf Nvl(!�Ƿ���, 0) = 1 Then
                        If Abs(!��׼����) > Abs(Val(Nvl(!�ּ�))) Then
                            objBillIncome.��׼���� = Val(Nvl(!ȱʡ�۸�))
                        Else
                            objBillIncome.��׼���� = !��׼����
                        End If
                    Else
                        objBillIncome.��׼���� = !�ּ�
                    End If
                                        
                    objBillIncome.��׼���� = Format(objBillIncome.��׼����, gSysPara.Price_Decimal.strFormt_VB)
                    
                    objBillIncome.�ּ� = Nvl(!�ּ�, 0) '�ּ�ԭ�۶�ҩƷ�������
                    objBillIncome.ԭ�� = Nvl(!ԭ��, 0)
                    objBillIncome.������ĿID = Nvl(!������ID, 0)
                    objBillIncome.������Ŀ = Nvl(!������Ŀ)
                    objBillIncome.�վݷ�Ŀ = Nvl(!�ַ�Ŀ)
                    
                    'Ӧ�ս��=����*����*����
                    objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * objBillDetail.���� * objBillDetail.����
                    
                    '�������������ü���(����������Ŀ)
                    If Nvl(!���ӱ�־, 0) = 1 And Nvl(!�շ����) = "F" Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * Nvl(!�����շ���, 100) / 100
                    End If
                    
                    '�Ӱ�����ʼ���
                    If Nvl(!�Ӱ��־, 0) = 1 And Nvl(!�Ӱ�Ӽ�, 0) = 1 Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * (1 + Nvl(!�Ӱ�Ӽ���, 0) / 100)
                    End If
                    objBillIncome.Ӧ�ս�� = Format(objBillIncome.Ӧ�ս��, gSysPara.Money_Decimal.strFormt_VB)
                    
                    '����ʵ�ս��
                    If Nvl(!���ηѱ�, 0) = 1 Then
                        objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                    Else
                        'ʹ��ԭ���Ķ�̬�ѱ�
                        objBillIncome.ʵ�ս�� = ActualMoney(objBill.�ѱ�, !������ID, objBillIncome.Ӧ�ս��, objBillDetail.�շ�ϸĿID, _
                            objBillDetail.ִ�в���ID, !���� * !����, IIf(Nvl(!�Ӱ��־, 0) = 1 And Nvl(!�Ӱ�Ӽ�, 0) = 1, Nvl(!�Ӱ�Ӽ���, 0) / 100, 0))
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
    
    Set ImportStuffBill = objBill
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function GetBillMoney(strNO As String, Optional ByVal int���� As Integer = 2, Optional ByVal bln���� As Boolean = False, Optional lng����ID As Long) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ���������ݽ��
    '���:strNO-���ݺ�
    '     int����=1-�շѵ�,2-���ʵ�,3-���ʵ�(�Զ����ʵ�),4-�Һŵ�
    '     lng����ID=����ID
    '     bln����=true:���ﲡ��:false-סԺ����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-10-17 17:03:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    If lng����ID = 0 Then
        strSQL = "Select Sum(ʵ�ս��) as ��� From  " & IIf(bln����, "������ü�¼", " סԺ���ü�¼") & " Where NO=[1] And ��¼����=[2] And ��¼״̬ IN(0,1)"
    Else
        strSQL = "Select Sum(ʵ�ս��) as ��� From " & IIf(bln����, "������ü�¼", " סԺ���ü�¼") & " Where NO=[1] And ��¼����=[2] And ��¼״̬ IN(0,1) And ����ID=[3]"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, int����, lng����ID)
    If Not rsTmp.EOF Then GetBillMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
 
 
Public Function GetPriceMoneyTotal(ByVal intType As Byte, lng����ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵Ļ��۵����ϼ�
    '���:intType:0-����;1-סԺ
    '����:���ػ����ܶ�
    '����:���˺�
    '����:2014-03-20 18:08:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String, blnAllFee As Boolean
    '���ʱ�����������סԺ���۷���
    If intType = 1 Then
        blnAllFee = Val(gobjDatabase.GetPara("���ʱ�����������סԺ���۷���", glngSys, 1150)) = 1
        If blnAllFee Then
            strWhere = ""
        Else
            strWhere = " And Nvl(��ҳID,0) = (Select Nvl(��ҳID,0) From ������Ϣ Where ����ID = [1])"
        End If
    Else
        strWhere = ""
    End If
        
    On Error GoTo errH
    If intType = 1 Then
        strSQL = "" & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ�  " & _
        "   From סԺ���ü�¼ " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1] and �����־=2" & strWhere
    Else
        '78226,Ƚ����,2014-9-24,�޸�SQL���
        '"   From סԺ���ü�¼ and �����־<>2 " & _
        '"   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1]"
        strSQL = "" & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ� " & _
        "   From ������ü�¼  " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1]  and �����־<>2" & _
        "   Union ALL   " & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ�  " & _
        "   From סԺ���ü�¼ " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1] and �����־<>2 "
        strSQL = "" & _
        "   Select Sum(nvl(���۷��úϼ�,0)) as ���۷��úϼ�  " & _
        "   From ( " & strSQL & ")"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡָ�����˵Ļ����ܶ�", lng����ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!���۷��úϼ�
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAuditRecord(lng����ID As Long, lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡָ�����˵ķ���������Ŀ
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ��ĿId,ʹ������,��������,ʹ������-�������� �������� From ����������Ŀ Where ����ID=[1] And ��ҳID=[2]"
    Set GetAuditRecord = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetMoneyInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional curModiMoney As Currency) As ADODB.Recordset
'���ܣ���ȡָ�����˵�ʣ���
'������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(�������,0) as �������,Nvl(Ԥ�����,0) as Ԥ�����" & _
            " From ������� Where ����=1 And ���� = " & IIf(lng��ҳID = 0, 1, 2) & " And ����ID= [1] "
    
    If curModiMoney <> 0 Then   '����Ҫ��Union��ʽ,���ֱ��ȥ��,�ڲ�������޼�¼ʱ,���᷵�ؼ�¼
        strSQL = strSQL & " Union All  Select -1* " & curModiMoney & " as �������,0 as Ԥ����� From Dual"
        strSQL = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ����� From (" & strSQL & ")"
    End If
            
    '���Ϊҽ��סԺ���ˣ����ڷ���������ſ�Ԥ���еķ���(���ڱ���)
    If lng��ҳID <> 0 Then
        strSQL = strSQL & " Union All " & _
            " Select -1*Nvl(Sum(���),0) as �������,0 as Ԥ�����" & _
            " From ����ģ����� Where ����ID=[1] And ��ҳID=[2]"
        strSQL = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ����� From (" & strSQL & ")"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetPatiUnit(lngPatiID As Long) As Long
'���ܣ����ز�����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.��ǰ����ID From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngPatiID)
    If Not rsTmp.EOF Then GetPatiUnit = Nvl(rsTmp!��ǰ����ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub AdjustCpt(lngID As Long)
'���ܣ�ҩƷ����
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "zl_ҩƷ�շ���¼_Adjust(" & lngID & ")"
    Call gobjDatabase.ExecuteProcedure(strSQL, "mdlExpense")
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Public Function BillisZeroLog(ByVal strNO As String, ByVal byt��Դ As Byte) As Boolean
'���ܣ��ж�ָ�������Ƿ�������ķ��õǼ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTab As String
    strTab = IIf(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")

    On Error GoTo errH

    strSQL = "Select ʵ�ս�� From " & strTab & " Where ��¼״̬ In(0,1,3) And ��¼����=2 And NO=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    BillisZeroLog = True
    For i = 1 To rsTmp.RecordCount
        If Nvl(rsTmp!ʵ�ս��, 0) <> 0 Then
            BillisZeroLog = False: Exit For
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function BillIdentical(ByVal strNO As String, byt��Դ As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ļ��ʵ����е�״̬�Ƿ�һ��,���Ƿ�ͬʱ������˺�δ��˵�����
    '���:strNO-���ݺ�
    '     byt��Դ-������Դ:1-����;2-סԺ
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-09 14:25:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIf(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
    BillIdentical = True
    
    
    On Error GoTo errHandle
    strSQL = _
        " Select Count(Distinct �Ǽ�ʱ��) as ʱ����," & _
        " Sum(Decode(��¼״̬,0,1,0)) as δ���," & _
        " Sum(Decode(��¼״̬,0,0,1)) as �����" & _
        " From " & strTab & _
        " Where ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=2"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!δ���, 0) <> 0 And Nvl(rsTmp!�����, 0) <> 0 Then
            BillIdentical = False
        ElseIf Nvl(rsTmp!ʱ����, 0) > 1 Then
            BillIdentical = False
        End If
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function CheckValidity(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, _
    Optional ByVal blnAsk As Boolean = True, Optional lng���� As Long = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ϵ����Ч���Ƿ����
    '���:blnAsk=��ʾ�Ƿ�ѯ���Ƿ����,����Ϊ����
    '       lng����:-1��ʾ���������������;>=0�������������Ч��
    '����:
    '����:
    '����:���˺�
    '����:2010-12-14 10:21:54
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As New ADODB.Recordset
    Dim curDate As Date, minDate As Date
    Dim strSQL As String, strName As String
    
    CheckValidity = True
    
    '��һ���Բ��ϲ��ж�
    '��Ϊ���ܸ��������Ч�ڲ�ͬ,���Ҫ�õ�����������С��Ч��
    strSQL = _
        " Select C.����,Nvl(B.����,0) as ����," & _
        "           B.�������� as ���,B.���Ч��,Sysdate as ʱ��" & _
        " From �������� A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
        " Where A.����ID=B.ҩƷID And A.����ID=C.ID And A.һ���Բ���=1" & _
        "       And B.����=1 And Nvl(B.��������,0)>0 And A.���Ч�� is Not NULL" & _
        "       And A.����ID=[1] And B.�ⷿID=[2] " & IIf(lng���� >= 0, " And nvl(b.����,0)=[3] ", "") & _
        " Order by Nvl(B.����,0)"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng�ⷿID, lng����)
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function HaveExecute(ByVal strNO As String, ByVal intFlag As Integer, ByVal blnAll As Boolean, ByVal byt��Դ As Byte) As Boolean
'���ܣ��жϷ��õ����Ƿ������ȫִ�л򲿷�ִ�е�����
'������strNO=���õ��ݺ�,intFlag=��¼����
'      blnALL=�б𵥾����Ƿ�ȫ��Ϊ��ȫִ�л򲿷�ִ�е�����
'      byt��Դ:1-���2-סԺ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    strTab = IIf(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
    
    On Error GoTo errH
    strSQL = "Select Nvl(Count(ID),0) as ��Ŀ" & _
        " From " & strTab & _
        " Where NO=[1] And ��¼����=[2] And ��¼״̬ IN(0,1,3) And " & IIf(blnAll, " Not", "") & " ִ��״̬ IN(1,2)"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "HaveExecute", strNO, intFlag)
    
    If blnAll Then
        HaveExecute = (rsTmp!��Ŀ = 0)
    Else
        HaveExecute = (rsTmp!��Ŀ > 0)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
    '���:intNum=��Ŀ���,Ϊ0ʱ�̶��������
    '����:
    '����:���ز�ȫ�ĵ��ݺ�
    '����:���˺�
    '����:2014-04-09 14:34:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim dtCurDate As Date, strMaxNo As String
    Dim strYearStr As String
    
    Err = 0: On Error GoTo errH:
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ����,������ From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, intNum)
    dtCurDate = Date
    If Not rsTmp.EOF Then
        intType = Val("" & rsTmp!��Ź���)
        dtCurDate = rsTmp!����
        strMaxNo = Nvl(rsTmp!������)
    End If
    strYearStr = PreFixNO
    If strMaxNo = "" Then strMaxNo = strYearStr & "000001"
    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate(Format(dtCurDate, "YYYY-MM-dd")) - CDate(Format(dtCurDate, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
        Exit Function
    End If
    '������
    If Len(strNO) = 6 Then
        GetFullNO = Left(strMaxNo, 2) & strNO: Exit Function
    End If
    GetFullNO = Left(strMaxNo, 2) & zlLeftPad(Right(strNO, 6), 6, "0")
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlLeftPad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���������ƿո�
    '����:�����ִ�
    '����:���˺�
    '����:2012-02-22 17:58:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = zlSubstr(strCode, 1, lngLen)
    End If
    zlLeftPad = Replace(strTmp, Chr(0), strChar)
End Function

Private Function zlSubstr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '���:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '����:�Ӵ�
    '����:���˺�
    '����:2012-02-22 18:00:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo Errhand:
    zlSubstr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    zlSubstr = Replace(zlSubstr, Chr(0), " ")
    Exit Function
Errhand:
    zlSubstr = ""
End Function

Public Function CheckAdviceDrugSurplus(ByVal lng���ͺ� As Long, Optional ByVal lngҽ��ID As Long) As String
'���ܣ���������ҩƷҽ���������Ƿ���ڵ�ǰ���������
'������lng���ͺ�=Ҫ���˵ķ��ͺ�
'      lngҽ��ID=Ҫ���˵�һ��ҩƷҽ����ID�������ָ���ɱ�ʾ�������˶���ҽ��
'���أ���ʾ��Ϣ
'˵������ʿ���ܻ���ҽ���Ĳ���������ֻ�漰סԺ���ü�¼(ҽ���ſ��ܷ�������Ϊ�������)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select C.ҽ������ as ҩƷ,A.�շ�ϸĿID as ҩƷID,A.���˲���ID as ����ID,A.ִ�в���ID as �ⷿID,Sum(A.����) as ��������" & _
        " From סԺ���ü�¼ A,����ҽ������ B,����ҽ����¼ C" & _
        " Where A.ҽ�����=B.ҽ��ID And A.NO=B.NO And A.��¼����=B.��¼����" & _
        " And B.ҽ��ID=C.ID And A.�շ���� In('5','6') And A.�۸񸸺� Is Null" & _
        " And B.���ͺ�=[1] And C.������� IN('5','6') And (C.���ID=[2] Or [2]=0)" & _
        " Group by C.ҽ������,A.�շ�ϸĿID,A.���˲���ID,A.ִ�в���ID"
    strSQL = _
        " Select A.ҩƷ,D.���� as �ⷿ,C.סԺ��װ,C.סԺ��λ,A.��������,B.��������" & _
        " From (" & strSQL & ") A,ҩƷ����ƻ� B,ҩƷ��� C,���ű� D" & _
        " Where A.�ⷿID=D.ID And A.ҩƷID=C.ҩƷID" & _
        " And A.����ID=B.����ID(+) And A.�ⷿID=B.�ⷿID(+) And A.ҩƷID=B.ҩƷID(+) And B.״̬(+)=0"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "CheckAdviceDrugSurplus", lng���ͺ�, lngҽ��ID)
    Do While Not rsTmp.EOF
        If Nvl(rsTmp!��������, 0) > Nvl(rsTmp!��������, 0) And Nvl(rsTmp!��������, 0) <> 0 Then
            strMsg = strMsg & vbCrLf & "��[" & rsTmp!ҩƷ & "]��""" & rsTmp!�ⷿ & """�Ļ������� " & _
                FormatEx(Nvl(rsTmp!��������, 0) / Nvl(rsTmp!סԺ��װ, 1), 5) & rsTmp!סԺ��λ & "����ǰ�������� " & _
                FormatEx(Nvl(rsTmp!��������, 0) / Nvl(rsTmp!סԺ��װ, 1), 5) & rsTmp!סԺ��λ
        End If
        rsTmp.MoveNext
    Loop
    
    If strMsg <> "" Then strMsg = "����ҩƷ�Ļ���������������������" & vbCrLf & strMsg & vbCrLf & vbCrLf & "Ҫ������"
    CheckAdviceDrugSurplus = strMsg
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckAdviceBillingRevoke(ByVal lngҽ��ID As Long) As Boolean
'���ܣ�(����)��Ҫ���ϵ�ҽ����Ӧ�ļ��ʷ��õ����������м��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng���ͺ� As Long
    Dim intPara As Integer
    
    On Error GoTo errH
    
    'ҽ��IDΪ����ֵ������ҽ����һ�������˵�,�����޷��͡�
    strSQL = "Select Distinct ���ͺ� From ����ҽ������" & _
        " Where ҽ��ID IN(Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1])"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngҽ��ID)
    If rsTmp.EOF Then Exit Function
    lng���ͺ� = rsTmp!���ͺ�
    
    intPara = Val(gobjDatabase.GetPara(68, glngSys))

   
    '����������"ZL_����ҽ����¼_����"
    strSQL = "Select A.NO,A.���" & _
        " From ������ü�¼ A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ I" & _
        " Where A.NO=B.NO And A.��¼���� IN(2,12) And A.��¼״̬=1" & _
        " And A.������ Is Not NULL And A.������<>A.����Ա����" & _
        " And A.ҽ�����=B.ҽ��ID And B.ҽ��ID=C.ID And B.��¼����=2" & _
        " And C.������ĿID=I.ID And B.���ͺ�=[1] And (C.ID=[2] Or C.���ID=[2])" & _
        " And (" & _
            " A.�շ���� Not In ('5','6','7','E')" & _
            " Or A.�շ����='E' And I.�������� Not In ('2','3','4')" & _
            " Or A.�շ���� In ('5','6','7') And Nvl(A.ִ��״̬,0)=0" & _
            " Or 0=[3])"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng���ͺ�, lngҽ��ID, intPara)
    If Not rsTmp.EOF Then Exit Function
    
    CheckAdviceBillingRevoke = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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

Public Function CheckFeeItemLimitDept(ByVal lngFeeItem As Long) As Boolean
'����:����շ���Ŀ,���������,�Ƿ������ڵ�ǰ���˿��һ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select ����id From �շ����ÿ��� Where ��Ŀid = [1] And (Select Count(����id) From �շѴ�����Ŀ Where ����id = [1]) > 0"

    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItem)
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!����ID = UserInfo.����ID Then
                CheckFeeItemLimitDept = True
                Exit For
            End If
            rsTmp.MoveNext
        Next
    Else
        CheckFeeItemLimitDept = True
    End If
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPatiIS�����ѱ�Ŀ(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional blnMsgbox As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�ŵ��ݵ�ʵ�ս��ϼ�,��һ�ż��ʱ���ָ�����˵�ʵ�ս��ϼ�
    '���أ��ѱ�Ŀ,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-08-12 11:26:28
    '˵����28725
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    strSQL = "Select NVL(A.����,b.����) ���� From ������ҳ A,������Ϣ B where a.����id=b.����id and  A.����id=[1] and a.��ҳid=[2] and ��Ŀ���� IS NOT NULL"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ��Ѿ�����", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        zlPatiIS�����ѱ�Ŀ = False
    Else
        zlPatiIS�����ѱ�Ŀ = True
        If blnMsgbox Then
                MsgBox "���ˡ�" & Nvl(rsTemp!����) & " ���Ѿ���Ŀ,��������м��ʻ����ʲ���!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Function
        End If
    End If
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then Resume
End Function
Public Function zlIs��������(ByVal strNO As String, ByVal lng��¼���� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ񱸻����ϼ���
    '���:strNO-���ݺ�
    '       lng��������:1-�շ�;2-����;
    '����:
    '����:����Ǳ������ϼ���,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-15 11:01:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If lng��¼���� = 1 Then
        strSQL = "Select  /*+ rule*/ 1 From ҩƷ�շ���¼ A,������ü�¼ B Where A.����ID=b.ID and A.����=21 And b.NO=[1] and b.��¼����=[2] and rownum <=1"
    Else
        strSQL = "Select  /*+ rule*/ 1 From ҩƷ�շ���¼ A,סԺ���ü�¼ B Where A.����ID=b.ID and A.����=21 And b.NO=[1] and b.��¼����=[2] and rownum <=1"
    End If
    
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����Ƿ���ڱ������ϼ���", strNO, lng��¼����)
    zlIs�������� = Not rsTemp.EOF
    rsTemp.Close
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_vsGrid_Para_Save(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional blnǿ�Ʊ��� As Boolean = False, Optional blnHaveParaPrivs As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '����:����vsFlex�Ŀ�ȵ�ע���
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If blnSaveToDataBase = False Then
        zl_vsGrid_Para_Save = True
        If blnǿ�Ʊ��� = False Then
            If Val(gobjDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
        End If
    End If
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '�����ʽ:������,�п�,������|������,�п�,������|...
    If blnSaveToDataBase Then
        gobjDatabase.SetPara strKey, strCol, glngSys, lngModule, blnHaveParaPrivs
    Else
        Call SaveRegInFor(g˽��ģ��, strCaption, strKey, strCol)
    End If
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�����ݿ��лָ�����Ŀ�ȵ���Ϣ
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '     blnSaveToDataBase-�Ƿ��������ݿ��б������(����������ݿ��б���,��ǿ�Ʊ���Ϊtrue,��������Ƿ�ʹ�ø��Ի������ȷ��)
    '     blnǿ�ƻָ�����-�����Ƿ񽫱���ע���Ĳ���ֵ,����ǿ�ƻָ�
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        'ֻ���ڱ���ע����вŻᴦ����Ի�����
        zl_vsGrid_Para_Restore = True
        If blnǿ�ƻָ����� = False Then
            If Val(gobjDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g˽��ģ��, strCaption, strKey, strParaValue)
    Else
        strParaValue = gobjDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
Errhand:
End Function
 




Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
End Sub
Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�������߶�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    GetTaskbarHeight = gobjComlib.OS.TaskbarHeight
End Function
Public Function GetVsGridBoolColVal(ByVal vsGrid As VSFlexGrid, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '����:��ȡbool�е�ֵ
    '����:�Ǹõ�Ԫ��Ϊtrue,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/28
    '------------------------------------------------------------------------------
    GetVsGridBoolColVal = gobjComlib.Grid.BoolVal(vsGrid, lngRow, lngCol)
End Function
Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub
Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln������� As Boolean = True, Optional bln���� As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     bln�������     �Ƿ���и������
    '     bln����         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    zlDblIsValid = gobjCommFun.DblIsValid(strInput, intMax, bln�������, bln����, hWnd, str��Ŀ)
End Function

Public Function zlIsAllowFeeChange(lng����ID As Long, lng��ҳID As Long, _
   Optional int״̬ As Integer = -1, Optional blnNotMsgBox As Boolean, Optional strOutErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�������ñ䶯
    '���:int״̬-(-1��ʾ�����ݿ��ж�ȡ��˱�־�����ж�;>0��ʾ,ֱ�Ӹ��ݸ�״̬�����ж�)
    '    blnNotMsgbox-�Ƿ���ʾ������ʾ��
    '����:strOutErrMsg-���ش�����Ϣ
    '����:����䶯����true,���򷵻�False
    '����:���˺�
    '����:2012-05-21 15:44:47
    '����:49501,51612
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If gSysPara.byt������˷�ʽ = 0 And gSysPara.blnδ��ƽ�ֹ���� = False Then
        ''����Ǹ��
        zlIsAllowFeeChange = True: Exit Function
    End If
    
    strSQL = "" & _
    " Select Nvl(��˱�־,0) as ��˱�־,nvl(״̬,0) as ״̬" & _
    " From ������ҳ " & _
    " Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        strOutErrMsg = "δ�ҵ���Ӧ�Ĳ�����Ϣ,��������з��ñ䶯����!"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '���δ��Ʋ��˲��������
    If gSysPara.blnδ��ƽ�ֹ���� And Val(Nvl(rsTemp!״̬)) = 1 Then
        '51612
        strOutErrMsg = "����δ���(��" & lng��ҳID & "��סԺ) ,���ܶԸò��˽��м��˻����˲�����"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '�����ؼ��
    If gSysPara.byt������˷�ʽ = 0 Then zlIsAllowFeeChange = True: Exit Function
    
    If int״̬ < 0 Then
        int״̬ = Val(Nvl(rsTemp!��˱�־))
    End If
    '������״̬
    If int״̬ = 1 Then
        strOutErrMsg = "�����ڵ�" & lng��ҳID & "��סԺ���Ѿ���ʼ��˷���,���ܶԸò��˽��з��ñ䶯��"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If int״̬ = 2 Then
        strOutErrMsg = "�Ѿ�����˶Բ��˵�" & lng��ҳID & "��סԺ���õ����,���ܶԸò��˽��з��ñ䶯��"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zlIsAllowFeeChange = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ش�д�ĵ��ݺ���ǰ׺
    '����:��ǰ׺
    '����:���˺�
    '����:2014-04-09 14:34:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PreFixNO = gobjComlib.zlStr.PreFixNO(curDate)
End Function



Public Function GetPatiUnitID(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Long
'���ܣ����ݲ��˻�ȡ��Ӧ�Ĳ���ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��ǰ����ID as ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    GetPatiUnitID = Nvl(rsTmp!����ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Check�ϰల��(ByVal blnҩ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽԺ�Ŀ����Ƿ�ʹ�����ϰల��
    '���:��blnҩ��=�Ǽ��ҩ���ϰ໹����������
    '����:
    '����:�������ϰ�ʱ��ķ���true,���򷵻�False
    '����:���˺�
    '����:2014-04-09 14:52:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Static blnҩ��Load As Boolean
    Static blnҩ��Last As Boolean
    Static bln��ҩLoad As Boolean
    Static bln��ҩLast As Boolean
    
    If blnҩ�� Then '�Ƿ��а���ֻ���ȡһ��
        If blnҩ��Load Then Check�ϰల�� = blnҩ��Last: Exit Function
    Else
        If bln��ҩLoad Then Check�ϰల�� = bln��ҩLast: Exit Function
    End If
    
    On Error GoTo errH
    
    If blnҩ�� Then
        strSQL = "Select 1 From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� IN('��ҩ��','��ҩ��','��ҩ��') And Rownum<2"
    Else
        strSQL = "Select 1 From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� Not IN('��ҩ��','��ҩ��','��ҩ��') And Rownum<2"
    End If
    Call gobjDatabase.OpenRecordset(rsTmp, strSQL, "Check�ϰల��")
    Check�ϰల�� = rsTmp.RecordCount > 0
    
    If blnҩ�� Then
        blnҩ��Load = True: blnҩ��Last = Check�ϰల��
    Else
        bln��ҩLoad = True: bln��ҩLast = Check�ϰల��
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function Get����Ա����ID(ByVal int������� As Integer, Optional ByVal lngĬ�ϲ��� As Long = 0) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ����Ա���������ָ������Ĳ��ţ�ȱʡ��������
    '����:���ز���Ա��ȱʡ����ID
    '����:���˺�
    '����:2014-04-09 14:53:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    
    If blnNew Then
        strSQL = "Select Distinct B.����ID,Nvl(B.ȱʡ,0) as ȱʡ,C.������� From ������Ա B,��������˵�� C" & _
            " Where B.��ԱID = [1] And B.����ID=C.����ID" & _
            " Order by ȱʡ Desc"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    
    '74794,Ƚ����,2014-7-18,��ʿ�ڼ���ʱ����ó��׷���ʱδʹ�ó��׷����ڵ�ִ�п���
    If lngĬ�ϲ��� <> 0 Then
        rsTmp.Filter = "(������� = 3 and ����ID = " & lngĬ�ϲ��� & ") " & _
                    "or (������� = " & int������� & " and ����ID = " & lngĬ�ϲ��� & ")"
        If Not rsTmp.EOF Then Get����Ա����ID = rsTmp!����ID: Exit Function
    End If
    
    rsTmp.Filter = "������� = 3 or ������� = " & int�������
    
    If Not rsTmp.EOF Then
        Get����Ա����ID = rsTmp!����ID
    Else
        Get����Ա����ID = UserInfo.����ID
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetPatiDayMoneyDetail(rsMoneyDay As ADODB.Recordset, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal byt��Դ As Byte, _
         Optional ByVal lng������ĿID As Long, Optional ByVal lng�շ�ϸĿID As Long, Optional ByVal date���ղ���ȡ As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵��켰֮��ҽ�������ķ�����Ŀ��ϸ
    '���:lng��ҳID=סԺ���˲�ʹ��
    '      byt��Դ:1-����(��סԺ�������͵�����)��2-סԺ
    '      str�״�ʱ��=����ҽ�����ͣ��״�ִ�е�ʱ��
    '      date���ղ���ȡ=����������ղ���ȡ����Ŀ��������Ƶ���ֲ���ÿ��һ�εģ�ʵ����ÿ��һ�εģ��������һ�Σ�ÿ24Сʱһ�ε�

    '����:rsMoneyDay������"������ĿID,�շ���ĿID,ִ�в���ID,ִ�з�,�շ�ʱ��"�ֶ�
    '����: ����Ƿ��͵���֮ǰ��ҽ�����򱾹�����ʱû�п��������������鵱���Ƿ���ִ��ʱ���鲻��
    '����:���˺�
    '����:2014-04-09 14:55:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strToDay As String, strDay As String
        
    On Error GoTo errH
    
    If lng������ĿID = 0 Then
        Set rsMoneyDay = New ADODB.Recordset '�������Filter����
        strToDay = Format(gobjDatabase.Currentdate, "yyyy-MM-dd")
        'ִ���жϣ�
        '1.������ǽ�������ü�¼�е�ִ�в��ţ����Ҳ�Է��ü�¼�е�ִ�в���Ϊ׼�жϡ�
        '2.���͸��������⣬ҽ�����õ�ִ�п�����ҽ��ִ�п�����ͬ���Ժ������ͬ�ˣ��ú���Ҳ������Ӧ
        '3.ҽ��ִ��ʱ����Ӧ���õ�ִ��״̬Ҳ��ͬ����ǡ�
        '4.�״β��յ���Ŀ�����Ƶ����һ��ֻ��һ�Σ���û�в������ü�¼������ҽ�����ͼ�¼��,��Ҫ���������������ɵģ��Ա������״β��յ���Ŀ�ж�
        If byt��Դ = 1 Then
            strSQL = "Select A.������ĿID,C.�շ�ϸĿID as �շ���ĿID,C.ִ�в���ID,Decode(Nvl(C.ִ��״̬,0),0,0,1) as ִ�з�,To_Char(C.����ʱ��,'yyyy-mm-dd') as �շ�ʱ��,0 as �շѷ�ʽ" & _
                " From ����ҽ����¼ A,����ҽ������ B,������ü�¼ C" & _
                " Where A.����ID=[1] And Nvl(A.��ҳID,0) = [2] And a.ҽ����Ч = 1 And A.ID=B.ҽ��ID And B.��¼����=C.��¼���� And B.NO=C.NO" & _
                " And B.ҽ��ID=C.ҽ����� And C.��¼״̬ IN(0,1) And C.����ʱ��>=[3]" & _
                " Union " & _
                " Select A.������ĿID,D.�շ�ϸĿid,D.ִ�п���ID as ִ�в���ID,0 as ִ�з�,To_Char(B.�״�ʱ��,'yyyy-mm-dd') as �շ�ʱ��,-1 as �շѷ�ʽ" & _
                " From ����ҽ����¼ A,����ҽ������ B,����ҽ���Ƽ� D" & _
                " Where A.����ID=[1] And Nvl(A.��ҳID,0) = [2] And a.ҽ����Ч = 1 " & _
                " And A.ID=B.ҽ��ID And NVL(B.�״�ʱ��,a.��ʼִ��ʱ��)>=[3] And A.ID=D.ҽ��ID And D.�շѷ�ʽ=7" & vbNewLine & _
                " And Not Exists (Select 1 From ������ü�¼ C Where c.�շ�ϸĿid=d.�շ�ϸĿid  And b.��¼���� = c.��¼���� And b.No = c.No And a.Id = c.ҽ�����)" & vbNewLine & _
                " Order by ������ĿID,�շ���ĿID"
            Set rsMoneyDay = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���켰������ҽ��", lng����ID, lng��ҳID, CDate(strToDay))
            Set rsMoneyDay = gobjDatabase.CopyNewRec(rsMoneyDay)
        Else
            '����������ҽ����¼.�ϴ�ִ��ʱ��Ϊ��
            '����������ҽ������ͬ���ã����ܲ�ͬʱ���η���,Unionȥ�����ظ���¼
            '�״β��յ���Ŀ�����Ƶ����һ��ֻ��һ�Σ���û�в������ü�¼������ҽ�����ͼ�¼��,��Ҫ���������������ɵģ��Ա������״β��յ���Ŀ�ж�
            strSQL = "Select a.������Ŀid, c.�շ�ϸĿid As �շ���Ŀid, c.ִ�в���id, Decode(Nvl(c.ִ��״̬, 0), 0, 0, 1) As ִ�з�," & vbNewLine & _
                "     Decode(a.ҽ����Ч, 0, b.�״�ʱ��, c.����ʱ��) As �״�ʱ��, Decode(b.�״�ʱ��,null, 1,Trunc(b.ĩ��ʱ��) - Trunc(b.�״�ʱ��) + 1) As ����,0 as �շѷ�ʽ" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ������ B, סԺ���ü�¼ C" & vbNewLine & _
                "Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.ҽ��id And b.��¼���� = c.��¼���� And b.No = c.No And b.ҽ��id = c.ҽ����� And" & vbNewLine & _
                "      c.��¼״̬ In (0, 1) And ((b.�״�ʱ�� > [3] Or b.ĩ��ʱ�� > [3]) Or a.ҽ����Ч = 1 And C.����ʱ�� >= [3])" & vbNewLine & _
                " Union " & vbNewLine & _
                "Select a.������Ŀid, D.�շ�ϸĿid, D.ִ�п���ID as ִ�в���id, 0 As ִ�з�," & vbNewLine & _
                "     b.�״�ʱ��, Decode(a.ҽ����Ч, 0, Trunc(b.ĩ��ʱ��) - Trunc(b.�״�ʱ��) + 1, 1) As ����,-1 as �շѷ�ʽ" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ������ B, ����ҽ���Ƽ� D" & vbNewLine & _
                "Where a.����id = [1] And a.��ҳid = [2]" & vbNewLine & _
                "   And a.Id = b.ҽ��id And ((b.�״�ʱ�� > [3] Or b.ĩ��ʱ�� > [3]) Or (a.ҽ����Ч = 1 And b.�״�ʱ�� is null and a.��ʼִ��ʱ�� >= [3]))" & vbNewLine & _
                "   And A.ID=D.ҽ��ID And D.�շѷ�ʽ=7" & vbNewLine & _
                " And Not Exists (Select 1 From סԺ���ü�¼ C Where c.�շ�ϸĿid=d.�շ�ϸĿid  And b.��¼���� = c.��¼���� And b.No = c.No And a.Id = c.ҽ�����)" & vbNewLine & _
                "Order By ������Ŀid, �շ���Ŀid"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���켰������ҽ��", lng����ID, lng��ҳID, CDate(strToDay))
            '���ݿ�ʼʱ�������������¼����ִ��ʱ��ֳɶ�����¼
            Set rsMoneyDay = InitPatiExecDays
                    
            For i = 1 To rsTmp.RecordCount
                For j = 1 To rsTmp!����
                    If j = 1 Then
                        strDay = Format(rsTmp!�״�ʱ��, "yyyy-MM-dd")
                    Else
                        strDay = Format(DateAdd("d", j - 1, CDate(rsTmp!�״�ʱ��)), "yyyy-MM-dd")
                    End If
                    If strDay >= strToDay Then
                        rsMoneyDay.Filter = "������ĿID=" & Val("" & rsTmp!������ĿID) & " And �շ���ĿID=" & Val("" & rsTmp!�շ���ĿID) & _
                                            " And �շ�ʱ��='" & strDay & "' And ִ�з�=" & Val("" & rsTmp!ִ�з�) & " And �շѷ�ʽ=" & Val("" & rsTmp!�շѷ�ʽ)
                        If rsMoneyDay.RecordCount = 0 Then
                            rsMoneyDay.AddNew
                            rsMoneyDay!������ĿID = Val("" & rsTmp!������ĿID)
                            rsMoneyDay!�շ���ĿID = Val("" & rsTmp!�շ���ĿID)
                            rsMoneyDay!ִ�в���ID = Val("" & rsTmp!ִ�в���ID)
                            rsMoneyDay!ִ�з� = Val("" & rsTmp!ִ�з�)
                            rsMoneyDay!�շѷ�ʽ = Val("" & rsTmp!�շѷ�ʽ)
                            rsMoneyDay!�շ�ʱ�� = strDay
                            rsMoneyDay.Update
                        End If
                    End If
                Next
                rsTmp.MoveNext
            Next
            rsMoneyDay.Filter = ""
        End If
    Else
        '���﷢��ʱ�����ж�ÿ���״β���ȡ����Ŀ�����Ƿ�ִ�д���=1,���=1��û���շѣ�˵�������״��Ѿ�û����ȡ��
        strSQL = "Select d.ִ�п���id As ִ�в���id" & vbNewLine & _
                "From ����ҽ����¼ A,����ҽ������ B, ����ҽ���Ƽ� D" & vbNewLine & _
                "Where A.����ID=[1] And Nvl(A.��ҳID,0) = [2] And a.Id = b.ҽ��id And A.id = d.ҽ��id And A.������ĿID = [6] And d.�շѷ�ʽ = 7 And d.�շ�ϸĿid = [3] And Not Exists" & vbNewLine & _
                " (Select 1" & vbNewLine & _
                "       From " & IIf(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼") & " C" & vbNewLine & _
                "       Where c.�շ�ϸĿid = d.�շ�ϸĿid And b.��¼���� = c.��¼���� And b.No = c.No And d.ҽ��id = c.ҽ�����) And" & vbNewLine & _
                "      Zl_Adviceexecount(d.ҽ��id, [4], [5],1) = 1"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���켰������ҽ��", lng����ID, lng��ҳID, lng�շ�ϸĿID, CDate(Format(date���ղ���ȡ, "yyyy-MM-dd")), CDate(Format(date���ղ���ȡ, "yyyy-MM-dd 23:59:59")), lng������ĿID)
        If rsTmp.RecordCount > 0 Then
            rsMoneyDay.Filter = "������ĿID=" & lng������ĿID & " And �շ���ĿID=" & lng�շ�ϸĿID & _
                                " And �շ�ʱ��='" & Format(date���ղ���ȡ, "yyyy-MM-dd") & "' And ִ�з�=0" & " And �շѷ�ʽ=-1"
            If rsMoneyDay.RecordCount = 0 Then
                rsMoneyDay.AddNew
                rsMoneyDay!������ĿID = lng������ĿID
                rsMoneyDay!�շ���ĿID = lng�շ�ϸĿID
                rsMoneyDay!ִ�в���ID = Val("" & rsTmp!ִ�в���ID)
                rsMoneyDay!ִ�з� = 0
                rsMoneyDay!�շѷ�ʽ = -1
                rsMoneyDay!�շ�ʱ�� = Format(date���ղ���ȡ, "yyyy-MM-dd")
                rsMoneyDay.Update
            End If
        End If
    End If
    
    GetPatiDayMoneyDetail = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Private Function InitPatiExecDays() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ����ط���ִ�еļ�¼��
    '����:ҽ����ط���ִ�еļ�¼��
    '����:���˺�
    '����:2014-04-09 14:56:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "������ĿID", adBigInt
    rsTmp.Fields.Append "�շ���ĿID", adBigInt
    rsTmp.Fields.Append "ִ�в���ID", adBigInt
    rsTmp.Fields.Append "�շѷ�ʽ", adInteger
    rsTmp.Fields.Append "ִ�з�", adInteger
    rsTmp.Fields.Append "�շ�ʱ��", adVarChar, 10
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set InitPatiExecDays = rsTmp
End Function


Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��������Ƿ���ԭ�ۺ��ִ��޶��ķ�Χ��
    '���:varL=ԭ��,varR=�ּ�,varI=������
    '����:������ڷ�Χ��,��Ϊ��ʾ��Ϣ,����Ϊ�մ�
    '����:���˺�
    '����:2014-04-09 15:44:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '�����ֵ������ͬ,���þ���ֵ�ж�
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "����ļ۸����ֵ���ڷ�Χ(" & FormatEx(Abs(varL), 5) & "-" & FormatEx(Abs(varR), 5) & ")��."
        End If
    Else
        '������Ų���ͬ,����ԭʼ��Χ�ж�
        If varI < varL Or varI > varR Then
            CheckScope = "����ļ۸�ֵ���ڷ�Χ(" & FormatEx(varL, 5) & "-" & FormatEx(varR, 5) & ")��."
        End If
    End If
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲿����Ϣ�Ƿ���ر���
    '����:��ʾ����,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 13:11:01
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    If mlng���ű���ƽ������ = 0 Then
        strSQL = "Select Avg(length(����)) As ���� From ���ű�"
        On Error GoTo errH
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ȡ���ű����ƽ������")
        mlng���ű���ƽ������ = Val(Nvl(rsTemp!����))
    End If
    '���ڱ��볤�ȿ��ܹ���,�޷���ʾ���ŵ�����,����Զ���ʾ�Ͳ���ʾ����,������5ʱ,����ʾ.С��5ʱ,��ʾ
   zlIsShowDeptCode = mlng���ű���ƽ������ <= 5
      
   Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get�շ�ִ�п���ID(ByVal lng����ID As Long, lng��ҳID As Long, _
    ByVal str��� As String, ByVal lng��Ŀid As Long, ByVal intִ�п��� As Integer, _
    ByVal lng���˿���ID As Long, ByVal lng��������id As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal lngִ�п���ID As Long, _
    Optional ByVal bytMode As Byte, Optional ByVal bytCallBy As Byte, _
    Optional ByVal int���ó��� As Integer = 1, _
    Optional lng����ȱʡִ�п��� As Long = 0) As Long
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ���Ŀ��ִ�п���
    '���:int��Χ=1.����,2-סԺ
    '      lngִ�п���ID=ָ����ȱʡִ�п���ID(����ҩƷ������)
    '      bytMode=1-Ҫ����ȱʡֵ,0-����
    '      bytCallBy=0-ҽ���������,1-���ѳ������
    '      int���ó���=1-����,2-סԺ
    '      lng����ȱʡִ�п���-ȱʡִ�п���ID
    '����:
    '����:����ָ����ִ�п���ID
    '����:���˺�
    '����:2014-04-09 13:58:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strҩ�� As String, lngҩ�� As Long
    Dim lng���˲���ID As Long, bytDay As Byte
    
    On Error GoTo errH
    
    If str��� = "4" Then
        lngҩ�� = Val(gobjDatabase.GetPara(IIf(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ���ϲ���", glngSys, _
            IIf(bytCallBy = 1, pҽ�����ѹ���, IIf(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�))))
        
        '��ִ�п�������ʱ
        strSQL = _
            " Select Distinct" & _
            "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            " And B.������� IN([1],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And (A.������Դ is NULL Or A.������Դ=[1])" & _
            " And (A.��������ID is NULL Or A.��������ID=[2]   " & _
            "       Or Exists(select 1 From �������Ҷ�Ӧ M where A.��������ID=M.����ID And M.����ID=[2] ))" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int��Χ, lng���˿���ID, lng��Ŀid)
        If Not rsTmp.EOF Then
            If bytMode = 1 Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID  '�����û�У��򷵻ص�һ�����õ�ִ�п���
            
            '1:ȱʡΪָ����(ҽ����)ִ�п���,�����Ƿ�����ڲ��˿���
            rsTmp.Filter = "ִ�п���ID=" & lngִ�п���ID
            
            '2.ȱʡΪ����ָ����ȱʡ����
            If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lngҩ��
            
            '3:�����ɷ����ڲ��˿��ҵ�ִ�п���
            If rsTmp.EOF Then
                '2.0 ��������д���ȱʡ��ִ�п���,��ȱʡΪ����ָ����ȱʡ����
                If lng����ȱʡִ�п��� <> 0 Then
                    rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                    If Not rsTmp.EOF Then
                            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                    End If
                End If
                '2.1:����ȱʡΪ���˿���
                If lngִ�п���ID <> lng���˿���ID And lngҩ�� <> lng���˿���ID Then
                    rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˿���ID
                End If
                '3.2:����ȱʡΪ���˲���
                If rsTmp.EOF And lng��ҳID <> 0 Then
                    lng���˲���ID = GetPatiUnitID(lng����ID, lng��ҳID)
                    If lng���˲���ID <> 0 And lng���˲���ID <> lng���˿���ID And lng���˲���ID <> lngִ�п���ID And lng���˲���ID <> lngҩ�� Then
                        rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˲���ID
                    End If
                End If
            End If
            '3.3:�ɷ����ڲ��˿��ҵ�һ��ִ�п���
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            
            '3.4�ɷ��������п��ҵĵ�ǰ���˿���ִ��
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0 And ִ�п���ID=" & lng���˿���ID
            
            '4:�����û�У��򷵻�0���ڼ��
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(gobjDatabase.GetPara(IIf(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIf(bytCallBy = 1, pҽ�����ѹ���, IIf(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�))))
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(gobjDatabase.GetPara(IIf(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIf(bytCallBy = 1, pҽ�����ѹ���, IIf(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�))))
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(gobjDatabase.GetPara(IIf(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIf(bytCallBy = 1, pҽ�����ѹ���, IIf(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�))))
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
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                " And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(gobjDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
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
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                " And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        End If
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strҩ��, int��Χ, lng���˿���ID, lng��Ŀid, bytDay)
        If Not rsTmp.EOF Then
            If lng����ȱʡִ�п��� <> 0 Then
                rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                If Not rsTmp.EOF Then
                        Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                End If
            End If
            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
            rsTmp.Filter = "ִ�п���ID=" & lngִ�п���ID
            If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lngҩ��
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    Else
        Select Case intִ�п���
            Case 0 '0-����ȷ����
                '1 ������Ŀѡ���Ҵ���ȱʡ��ִ�п��ҵ� ������Ŀ��ִ�в���ID
                If lng����ȱʡִ�п��� <> 0 Then
                    Get�շ�ִ�п���ID = lng����ȱʡִ�п���: Exit Function
                End If
                '101736,�ֹ�����ȱʡִ�п���
                '2 �շ���Ŀ.ȱʡ����(�ֹ�����ȱʡִ�п���)
                If int��Χ = 2 Then
                    strSQL = "Select a.ִ�п���id" & vbNewLine & _
                            " From �շ�ִ�п��� A, ���ű� C" & vbNewLine & _
                            " Where a.ִ�п���id + 0 = c.Id And a.�շ�ϸĿid = [1]" & vbNewLine & _
                            "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
                            "       And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null)" & vbNewLine & _
                            "       And a.������Դ = [2] And a.��������id Is Null"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng��Ŀid, 2)
                    If Not rsTmp.EOF Then
                        If Val(Nvl(rsTmp!ִ�п���ID)) <> 0 Then
                            Get�շ�ִ�п���ID = Val(Nvl(rsTmp!ִ�п���ID)): Exit Function
                        End If
                    End If
                    '3 ���˿���
                    If lng���˿���ID <> 0 Then Get�շ�ִ�п���ID = lng���˿���ID: Exit Function
                    '4 ��������
                    If lng��������id <> 0 Then Get�շ�ִ�п���ID = lng��������id: Exit Function
                End If
                '5 ����Ա��������ID
                Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ)
            Case 1 '1-�������ڿ���
                Get�շ�ִ�п���ID = lng���˿���ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get�շ�ִ�п���ID = lng���˿���ID
                Else
                    Get�շ�ִ�п���ID = GetPatiUnitID(lng����ID, lng��ҳID)
                End If
            Case 3 '3-����Ա���ڿ���
                Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ, lng����ȱʡִ�п���)
            Case 4 '4-ָ������
                strSQL = "Select Distinct Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID,Decode(A.������Դ,Null,2,1) as ����" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.�շ�ϸĿID=[1] And A.ִ�п���ID=B.����ID" & _
                    " And B.������� IN([2],3) And (A.������Դ is NULL Or A.������Դ=[2])" & _
                    " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                    " And A.ִ�п���ID=C.ID And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " Order by ����" 'Ĭ�Ͽ�������
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid, int��Χ, lng���˿���ID)
                If Not rsTmp.EOF Then
                    If lng����ȱʡִ�п��� <> 0 Then
                         rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                         If Not rsTmp.EOF Then
                                 Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                         End If
                     End If
                    Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                    rsTmp.Filter = "��������ID=" & lng���˿���ID
                    If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                End If
            Case 6 '6-���������ڿ���
                Get�շ�ִ�п���ID = lng��������id
        End Select
        If Get�շ�ִ�п���ID = 0 Then Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function PatiCanBilling(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String, Optional ByVal lngModual As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�������Ƿ�������Ȩ��
    '���:lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     strPrivs-Ȩ�޴�
    '     lngModual-ģ���
    '����:�������Ȩ��,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-09 14:13:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(strPrivs, ";��Ժ����ǿ�Ƽ���;") > 0 Then Exit Function
    
    On Error GoTo errH
    strSQL = "Select NVL(B.����,A.����) ����,B.��Ժ����,B.״̬,X.�������" & _
        " From ������Ϣ A,������ҳ B,������� X" & _
        " Where A.����ID=B.����ID And A.����ID=X.����ID(+) And X.����(+) = 2" & _
        " And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!��Ժ����) And Nvl(rsTmp!״̬, 0) <> 3 Then Exit Function
        If InStr(strPrivs, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
            If Nvl(rsTmp!�������, 0) <> 0 Then
                strMsg = """" & rsTmp!���� & """�ķ���δ���壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If InStr(strPrivs, ";��Ժ����ǿ�Ƽ���;") = 0 Then
            If Nvl(rsTmp!�������, 0) = 0 Then
                strMsg = """" & rsTmp!���� & """�ķ����ѽ��壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If lngModual = pҽ�����ѹ��� Or lngModual = pסԺҽ������ Or lngModual = pסԺҽ���´� Then
            '68081�������Ժ���˴���ҽ������
            strMsg = """" & rsTmp!���� & """�Ѿ���Ժ(��Ԥ��Ժ)�����ܶԸò��˵�ҽ�����з��͡������ջء�ִ�С����ˡ�"
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function



Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng����ID As Long, _
    ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal cur��� As Currency, ByVal str��� As String, ByVal str����� As String) As Boolean
'���ܣ���ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
'������str���="CDE..."����������漰�����շ����
'      str�����="���,����,..."����Ӧ�������������ʾ
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSQL As String, intR As Integer, i As Long
    Dim cur���� As Currency
    
    On Error GoTo errH
    
    If lng��ҳID <> 0 Then
        'סԺ���˱���
        strSQL = _
            " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1] And ���� = 2" & _
            " Union ALL" & _
            " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
        strSQL = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSQL & ") Group by ����ID"
        
        strSQL = "Select NVL(B.����,A.����) ����, Nvl(B.סԺ��,A.סԺ��) As סԺ��, Nvl(B.��Ժ����,A.��ǰ����) As ����,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���,C.ʣ���," & _
            " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
            " From ������Ϣ A,������ҳ B,(" & strSQL & ") C" & _
            " Where A.����ID=B.����ID And A.����ID=C.����ID(+)" & _
            " And A.����ID=[1] And B.��ҳID=[2]"
        Set rsPati = gobjDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID, lng��ҳID)
    Else
        '���������ﱨ��
        strSQL = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ����ID=[1] And ���� = 1"
        strSQL = "Select A.����,A.סԺ��,A.��ǰ���� As ����,zl_PatiWarnScheme(A.����ID) as ���ò���,A.������," & _
            " Nvl(B.Ԥ�����,0)-Nvl(B.�������,0)+Nvl(E.�ʻ����,0) as ʣ���" & _
            " From ������Ϣ A,(" & strSQL & ") B,ҽ�����˹����� D,ҽ�����˵��� E" & _
            " Where A.����ID=B.����ID(+) And A.����id = D.����id(+) And A.����=D.����(+)" & _
            " And D.����=E.����(+) And D.����=E.����(+) And D.ҽ����=E.ҽ����(+) And D.��־(+)=1" & _
            " And A.����ID=[1]"
        Set rsPati = gobjDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID)
    End If
    
    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ
    'ִ�б���:���ﲡ�˲���ID=0
    strSQL = "Select Nvl(��������,1) as ��������,����ֵ,������־1,������־2,������־3 From ���ʱ����� Where Nvl(����ID,0)=[1] And ���ò���=[2]"
    Set rsWarn = gobjDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID, CStr(Nvl(rsPati!���ò���)))
    If Not rsWarn.EOF Then
        If Val(Nvl(rsWarn!��������)) = 2 Then cur���� = GetPatiDayMoney(lng����ID)
        str����� = Mid(str�����, 2)
        For i = 1 To Len(str���)
            intR = BillingWarn(frmParent, strPrivs, rsWarn, Nvl(rsPati!����) & IIf(Nvl(rsPati!סԺ��) = "", "", "(סԺ��:" & Nvl(rsPati!סԺ��) & " ����:" & Nvl(rsPati!����) & ")"), Nvl(rsPati!ʣ���, 0), cur����, cur���, Nvl(rsPati!������, 0), Mid(str���, i, 1), Split(str�����, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function





Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���в��� As String = "", _
    Optional blnSendKeys As Boolean = True, Optional blnAddItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str���в���-���в�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim intIndex As Integer
    Dim strIDs As String, str���� As String, strLike As String
    strLike = IIf(Val(gobjDatabase.GetPara("����ƥ��")) = 0, "*", "")
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = gobjDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = strLike & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf gobjCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���в��� <> "" Then
        str���� = gobjCommFun.SpellCode(str���в���)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str���в���) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0:  Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp): Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 And lngDeptID <= 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    End Select
    
    '����ѡ����
    If gobjDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    If gobjControl.CboLocate(cboDept, lngDeptID, True) = False Then
        If blnAddItem = True Then
            If rsTemp.RecordCount = 1 Then
                cboDept.RemoveItem cboDept.ListCount - 1
                cboDept.AddItem IIf(zlIsShowDeptCode, rsTemp!���� & "-", "") & rsTemp!����
                cboDept.ItemData(cboDept.ListCount - 1) = Val(Nvl(rsTemp!ID))
                intIndex = cboDept.NewIndex
                cboDept.AddItem "�������ҡ�"
                cboDept.ItemData(cboDept.ListCount - 1) = 0
                cboDept.ListIndex = intIndex
            Else
                cboDept.RemoveItem cboDept.ListCount - 1
                cboDept.AddItem IIf(zlIsShowDeptCode, rsReturn!���� & "-", "") & rsReturn!����
                cboDept.ItemData(cboDept.ListCount - 1) = Val(Nvl(rsReturn!ID))
                intIndex = cboDept.NewIndex
                cboDept.AddItem "�������ҡ�"
                cboDept.ItemData(cboDept.ListCount - 1) = 0
                cboDept.ListIndex = intIndex
            End If
            rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
            zlSelectDept = True
            Exit Function
        Else
            GoTo GoNotSel
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing

    If blnSendKeys Then gobjCommFun.PressKey vbKeyTab
    zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.TxtSelAll cboDept
End Function


Public Function Getҽ����������(ByVal lngҽ��ID As Long, ByVal str������ As String) As String
'����:����ҽ��ID��Ԫ�����ơ�����ҽ���Ķ�ӦԪ�ص����븽������
'����:str������  ����������Ŀ.������

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
     
    strSQL = "Select a.���� From ����ҽ������ A, ����������Ŀ B" & _
        " Where a.Ҫ��id = b.Id And a.ҽ��id = [1] And b.������ = [2]"

    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID, str������)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strTmp = IIf(strTmp = "", "", strTmp & ",") & rsTmp!����
            rsTmp.MoveNext
        Next
    End If
    
    Getҽ���������� = strTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function GetStock(ByVal lngҩƷID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal int��Χ As Integer = 2, _
        Optional ByVal strDepartments As String, Optional ByVal lng���� As Double, Optional ByVal lng���� As Long = -1) As Double
'���ܣ���ȡָ���ָⷿ��ҩƷ���������(�������סԺ��λ)
'������int��Χ=1-����,2-סԺ(ȱʡ),0-��ʾ���ۼ�
'      strDepartments����ִ�п����ַ���������������ѯ���
'      lng���� ���lng������Ϊ�գ����ѯ�Ƿ��п������������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    '��ȡҩƷ���(�����������ҩƷ),ҩ��������ҩƷ����Ч��
    If int��Χ = 0 Or int��Χ = 3 Then
        If lng���� = 0 Or lng���� = -1 Then
            strSQL = _
                " Select Nvl(Sum(A.��������),0) as ���" & _
                " From ҩƷ��� A" & _
                " Where A.����=1" & _
                " And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
                " And A.ҩƷID=[1] And Instr([2],',' || a.�ⷿid || ',')>0 " & _
                " Group By A.�ⷿID"
        Else
            strSQL = _
                " Select Nvl(Sum(A.��������),0) as ���" & _
                " From ҩƷ��� A" & _
                " Where A.����=1" & _
                " And Nvl(A.����,0) = [3] And (A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
                " And A.ҩƷID=[1] And (Instr([2],',' || a.�ⷿid || ',')>0 Or A.�ⷿID In (Select ����ⷿid From ����ⷿ���� Where Instr([2],',' || ����id || ',')>0 And Rownum < 2)) " & _
                " Group By A.�ⷿID Order By Sign(Nvl(Sum(A.��������),0)) Desc "
        End If
    Else
        strTmp = IIf(int��Χ = 1, "����", "סԺ")
        If lng���� = 0 Or lng���� = -1 Then
            strSQL = _
                " Select Nvl(Sum(A.��������),0)/Nvl(B." & strTmp & "��װ,1) as ���" & _
                " From ҩƷ��� A,ҩƷ��� B" & _
                " Where A.ҩƷID=B.ҩƷID(+) And A.����=1" & _
                " And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
                " And A.ҩƷID=[1] And Instr([2],',' || a.�ⷿid || ',')>0" & _
                " Group by Nvl(B." & strTmp & "��װ,1),A.�ⷿID"
        Else
            strSQL = _
                " Select Nvl(Sum(A.��������),0)/Nvl(B." & strTmp & "��װ,1) as ���" & _
                " From ҩƷ��� A,ҩƷ��� B" & _
                " Where A.ҩƷID=B.ҩƷID(+) And A.����=1" & _
                " And Nvl(A.����,0) = [3] And (A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
                " And A.ҩƷID=[1] And (Instr([2],',' || a.�ⷿid || ',')>0 Or A.�ⷿID In (Select ����ⷿid From ����ⷿ���� Where Instr([2],',' || ����id || ',')>0 And Rownum < 2)) " & _
                " Group by Nvl(B." & strTmp & "��װ,1),A.�ⷿID Order By Sign(Nvl(Sum(A.��������),0)) Desc"
        End If
    End If
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҩƷID, IIf(strDepartments = "", "," & lng�ⷿID & ",", "," & strDepartments & ","), lng����)
    
    Do While Not rsTmp.EOF
    
        If strDepartments = "" Then
            GetStock = Format(rsTmp!���, "0.00000")
            Exit Function
        Else
            If Val(rsTmp!���) & "" > lng���� Then
                GetStock = Format(rsTmp!���, "0.00000")
                Exit Function
            End If
        End If
        rsTmp.MoveNext
    
    Loop
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function



Public Function Get��������(lngID As Long, Optional ByRef rs���� As ADODB.Recordset) As String
'���ܣ���ȡ��������
'������lngID=����ID
'���أ���������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If rs���� Is Nothing Then
        strSQL = "Select ���� from ���ű� Where ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngID)
    Else
        Set rsTmp = rs����
        rsTmp.Filter = "ID=" & lngID
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select ���� from ���ű� Where ID=[1]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngID)
        End If
    End If
    If Not rsTmp.EOF Then Get�������� = rsTmp!����
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get��Ŀ����(lng��Ŀid As Long) As String
'���ܣ�����������Ŀ����
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ���� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng��Ŀid)
    If Not rsTmp.EOF Then Get��Ŀ���� = rsTmp!����
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckFeeItemAvailable(ByVal lngFeeItemID As Long, ByVal bytFlag As Byte) As Boolean
'����:����շ���Ŀ�Ƿ�δͣ��,���ҷ����ڲ���
'����:bytFlag:�������:1-����,2-סԺ
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From �շ���ĿĿ¼ Where ID = [1] And (����ʱ�� is Null Or ����ʱ�� > Sysdate) And ������� In (" & bytFlag & ",3)"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItemID)
    CheckFeeItemAvailable = rsTmp.RecordCount > 0
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function Get�վݷ�Ŀ(ByVal lng������ĿID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�վݷ�Ŀ
    '����:�����վݷ�Ŀ
    '����:���˺�
    '����:2014-04-11 16:33:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    'On Error GoTo errHandle
    '�������쳣���ú�����Ҫ���ڱ�������ʱʹ�ã��������崦���쳣
    If grs������Ŀ Is Nothing Then
        strSQL = "Select ID,����,����,�վݷ�Ŀ From ������Ŀ��Where  (����ʱ�� is Null Or ����ʱ�� > Sysdate)"
        Set grs������Ŀ = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ŀ")
    ElseIf grs������Ŀ.State <> 1 Then
        strSQL = "Select ID,����,����,�վݷ�Ŀ From ������Ŀ ��Where  (����ʱ�� is Null Or ����ʱ�� > Sysdate)"
        Set grs������Ŀ = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ŀ")
    End If
    grs������Ŀ.Filter = "ID=" & lng������ĿID
    If grs������Ŀ.EOF = False Then Get�վݷ�Ŀ = grs������Ŀ!�վݷ�Ŀ
'Get�վݷ�Ŀ = True
'    Exit Function
'errHandle:
'    If gobjComlib.ErrCenter() = 1 Then
'        Resume
'    End If
End Function
Public Function GetPatiInforFromAdvice(ByVal lngҽ��ID As Long, _
    ByVal lng����ID As Long, _
    ByVal lng��ҳID As Long, ByRef lng���˿���ID As Long, _
    ByRef lng���˲���ID As Long, _
    ByRef lngҽ��С��ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ����ȡ��ʱ��ҽ���Ĳ������״̬(���˿���ID,��ǰ����ID,ҽ��С��ID)
    '���:lngҽ��ID-ҽ��ID
    '     lng����ID-����ID
    '     lng��ҳID-��ҳID
    '����:lng���˿���ID-���ص�ʱҽ���Ĳ��˿���id
    '     lng���˲���ID-���ص�ʱҽ���Ĳ��˲���ID
    '     lngҽ��С��ID-���ص�ʱҽ����ҽ��С��ID
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-21 15:05:19
    '����:70896
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    'ֻ��סԺ�Ż����
    lng���˿���ID = 0: lng���˲���ID = 0: lngҽ��С��ID = 0
    If Not (lng����ID <> 0 And lng��ҳID <> 0) Then GetPatiInforFromAdvice = True: Exit Function
    
    strSQL = " " & _
    " Select * " & _
    " From (Select a.����id, Nvl(b.���˿���id, a.����id) As ���˿���id, a.ҽ��С��id, a.��ʼʱ�� " & _
    "        From ���˱䶯��¼ A, (Select ����ʱ��, ���˿���id From ����ҽ����¼ Where ID = [3]) B " & _
    "        Where a.����id = [1] And ��ҳid = [2] And b.����ʱ�� Between ��ʼʱ�� And Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
    "              Nvl(a.����id, 0) <> 0 " & _
    "        Order By ��ʼʱ�� Desc) " & _
    " Where Rownum < 2"
 
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "����ҽ����ȡ���������˿���Id", lng����ID, lng��ҳID, lngҽ��ID)
    If rsTemp.EOF Then GetPatiInforFromAdvice = True: Exit Function
    lng���˿���ID = Nvl(rsTemp!���˿���id, 0)
    lng���˲���ID = Nvl(rsTemp!����ID, 0)
    lngҽ��С��ID = Nvl(rsTemp!ҽ��С��id, 0)
    GetPatiInforFromAdvice = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function BillOperCheck(bytNO As Byte, strOperator As String, Datadd As Date, Optional strMessage As String = "����", _
    Optional ByVal strNO As String, Optional ByVal lngPatientID As Long, _
    Optional ByVal bytFlag As Byte = 2, Optional ByVal blnOnlyCheckLimit As Boolean, Optional ByVal blnCheckOperator As Boolean = True, _
    Optional ByVal blnCheckCur As Boolean = True, Optional blnNotMsgBox As Boolean, Optional strOutErrMsg As String) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ��Ա�Ե����Ƿ��в���Ȩ��
    '���: bytNO��1-�Һŵ���,2-�շѵ�,3-���۵�,4-�������,5-סԺ����,6-Ԥ����,7-���ʵ���,8-���￨
    '   strOperator������ʵ�ʵĲ���Ա
    '   DatAdd�����ݵĵǼ�ʱ��
    '   strNO   �����������ʱ����ȷ������
    '   lngPatientID�����������ʱ�����ڼ��ʱ�����ȷ�������еĲ���
    '   bytFlag��1-�շѵ�,2-���ʵ�,3-���ʵ�
    '   blnOnlyCheckLimit��ֻ���������
    '   blnCheckOperator��Ҫ����Ƿ�����������˵���
    '   blnCheckCur���Ƿ���������
    '   blnNotMsgBox��True-����ʾ��Ϣ��ʾ��;False-��ʾ��Ϣ��ʾ��
    '����:strOutErrMsg-���ش�����Ϣ
    '����:�Ƿ��в���Ȩ��,�з���true,���򷵻�False
    '����:���˺�
    '����:2016-10-17 17:13:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
 

    Dim strSQL As String, strBill As String
    Dim rsTmp As ADODB.Recordset
    Dim curTmp As Currency
    Dim int��Դ As Integer
    
    If bytNO = 1 Or bytNO = 2 Or (bytNO = 3 And bytFlag = 1) Or bytNO = 4 Then
        int��Դ = 1
    Else
        int��Դ = 2
    End If
    
    If glngSys Like "8??" Then
        strBill = Switch(bytNO = 1, "�Һŵ���", bytNO = 2, "�շѵ���", bytNO = 3, _
            "���۵���", bytNO = 4, "���ʵ���", bytNO = 5, "���ʵ���", _
            bytNO = 6, "Ԥ�����", bytNO = 7, "���ʵ���", bytNO = 8, "��Ա��")
    Else
        strBill = Switch(bytNO = 1, "�Һŵ���", bytNO = 2, "�շѵ���", bytNO = 3, _
            "���۵���", bytNO = 4, "���ʵ���", bytNO = 5, "���ʵ���", _
            bytNO = 6, "Ԥ�����", bytNO = 7, "���ʵ���", bytNO = 8, "���￨")
    End If
        
    On Error GoTo errH
    
    strSQL = "" & _
    "   Select Nvl(ʱ������,0) as ʱ������,Nvl(���˵���,0) as ���˵���,Nvl(�������,0) as ������� " & _
    "   From ���ݲ������� Where ��ԱID=[1] And ����=[2]"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, UserInfo.ID, bytNO)
    If rsTmp.EOF Then
        BillOperCheck = True
        Exit Function
    End If

    If Not blnOnlyCheckLimit Then
        If rsTmp!���˵��� = 0 And blnCheckOperator Then
            If strOperator <> UserInfo.���� Then
                strOutErrMsg = "��û��Ȩ�޶�" & strOperator & "�����" & strBill & "����" & strMessage & "��"
                If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If rsTmp!ʱ������ > 0 Then
            If Int(gobjDatabase.Currentdate) - Int(CDate(Datadd)) + 1 > rsTmp!ʱ������ Then
                strOutErrMsg = "��ֻ�ܶ� " & rsTmp!ʱ������ & " ���ڴ����" & strBill & "����" & strMessage & "��"
                If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    If rsTmp!������� > 0 And blnCheckCur Then
        If strNO <> "" Then
            curTmp = GetBillMoney(strNO, 2, IIf(int��Դ = 1, True, False), lngPatientID)
            If curTmp >= rsTmp!������� Then
                strOutErrMsg = "��ֻ�ܶ� " & rsTmp!������� & " Ԫ���µ�" & strBill & "����" & strMessage & "��" & _
                vbCrLf & "����[" & strNO & "]��ʵ�ս��ϼ�Ϊ:" & curTmp
                If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    BillOperCheck = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Check���۲���(ByVal strNO As String, ByVal strPrivs As String, Optional ByVal strTime As String, Optional ByVal bytFlag As Byte = 2, _
    Optional blnNotMsgBox As Boolean, Optional strOutErrMsg As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ƿ���������۲����˽��м���,�Լ��ʵ�/����м��
    '     ��Ҫ���ڼ��ʵ�/���޸�,���ʡ����ڼ��ʱ�,ֻҪ����һ�����۲�����Ȩ��,��������ֹ
    '���:
    '   blnNotMsgBox��True-����ʾ��Ϣ��ʾ��;False-��ʾ��Ϣ��ʾ��
    '����:strOutErrMsg-���ش�����Ϣ
    '����:û��Ȩ�޵����۲���,��"���۲���","�������۲���","סԺ���۲���"
    '����:���˺�
    '����:2016-10-17 17:44:16
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset
    Dim bln�������� As Boolean
    Dim blnסԺ���� As Boolean
    Dim strSQL As String
    
    bln�������� = gSysPara.bln�������ۼ��� And InStr(strPrivs, ";�������ۼ���;") > 0
    blnסԺ���� = gSysPara.blnסԺ���ۼ��� And InStr(strPrivs, ";סԺ���ۼ���;") > 0
        
    If bln�������� And blnסԺ���� Then Exit Function
    
    If Not bln�������� And Not blnסԺ���� Then
        strSQL = "1,2"
    ElseIf Not bln�������� Then
        strSQL = "1"
    ElseIf Not blnסԺ���� Then
        strSQL = "2"
    End If
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(B.��������,0) as ��������" & _
        " From סԺ���ü�¼ A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
        " And A.NO=[1] And A.��¼����=[2]" & _
        " And Nvl(B.��������,0) IN(" & strSQL & ") And A.��¼״̬ IN(0,1,3)" & _
        IIf(strTime <> "", " And A.�Ǽ�ʱ��=[3]", "")
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    If Not rsTmp.EOF Then
        If rsTmp.RecordCount = 2 Then
            Check���۲��� = "���۲���"
        ElseIf rsTmp!�������� = 1 Then
            Check���۲��� = "�������۲���"
        ElseIf rsTmp!�������� = 2 Then
            Check���۲��� = "סԺ���۲���"
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function CheckDelPriv(ByVal strNO As String, ByVal strPrivs As String, Optional ByVal strTime As String, _
        Optional ByVal bytFlag As Byte = 2, Optional ByVal bytMode As Byte = 1, Optional ByVal blnNotMsgBox As Boolean, Optional ByRef strOutErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�Ȩ�޳���סԺ���ʵ�
    '���:bytMode,����Ȩ�޲���ʱ�Ƿ����ʾ,1-�������,������,0-���ʼ���,���ؼ�
    '   blnNotMsgBox��True-����ʾ��Ϣ��ʾ��;False-��ʾ��Ϣ��ʾ��
    '����:strOutErrMsg=���ش�����Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-10-17 17:22:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ֻ�ж�δ���ʷ�����
    strSQL = "Select Nvl(Sum(Decode(�շ����,'5',1,'6',1,'7',1,0)),0) as ҩƷ��," & _
        " Nvl(Sum(Decode(�շ����,'4',1,0)),0) as ������," & _
        " Nvl(Sum(Decode(�շ����,'4',0,'5',0,'6',0,'7',0,1)),0) as ������" & _
        " From סԺ���ü�¼" & _
        " Where ��¼����=[2] And ��¼״̬ IN(0,1) And NO=[1]" & _
        IIf(strTime <> "", " And �Ǽ�ʱ��=[3]", "")
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytFlag)
    End If
    
    If rsTmp.EOF Then CheckDelPriv = True: Exit Function
    'û��סԺ����Ȩ��ʱ,�˵��Ͱ�ť������Ϊ���ɼ�
    Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
    Dim strNotPrivs As String, strNote As String
    
    blnYP = InStr(strPrivs, ";ҩƷ����;") > 0
    blnZL = InStr(strPrivs, ";��������;") > 0
    blnWC = InStr(strPrivs, ";��������;") > 0
    
    If blnYP = False And blnZL = False And blnWC = False Then
        strOutErrMsg = "��û��ҩƷ���ʻ��������ʻ��������ʵ�Ȩ��,���ܶԵ���[" & strNO & "]�������ʣ�"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    strNotPrivs = ""
    If Not blnYP Then strNotPrivs = strNotPrivs & "��ҩƷ����"
    If Not blnWC Then strNotPrivs = strNotPrivs & "����������"
    If Not blnZL Then strNotPrivs = strNotPrivs & "����������"
    strNotPrivs = Mid(strNotPrivs, 2)
    strNote = ""
    
    If blnYP Then strNote = strNote & "��ҩƷ����"
    If blnWC Then strNote = strNote & "����������"
    If blnZL Then strNote = strNote & "����������"
    strNote = Mid(strNote, 2)
 
    If rsTmp!ҩƷ�� > 0 And Not blnYP Then
        strOutErrMsg = "��û��" & strNotPrivs & "Ȩ��,ֻ�ܶԵ���[" & strNO & "]�е�" & strNote & "�������ʣ�"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    If rsTmp!������ > 0 And Not blnWC Then
        strOutErrMsg = "��û��" & strNotPrivs & "Ȩ��,ֻ�ܶԵ���[" & strNO & "]�е�" & strNote & "�������ʣ�"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    If rsTmp!������ > 0 And Not blnZL Then
        strOutErrMsg = "��û��" & strNotPrivs & "Ȩ��,ֻ�ܶԵ���[" & strNO & "]�е�" & strNote & "�������ʣ�"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    CheckDelPriv = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function




Public Function BillCanBeOperate(ByVal strNO As String, ByVal strPriv As String, _
    ByVal strNote As String, Optional ByVal strTime As String, _
    Optional str����IDs As String, Optional ByVal bytType As Byte = 2, Optional ByVal blnNotMsgBox As Boolean, _
    Optional ByRef strOutErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݵĲ�����Ϣ�ж��Ƿ���Ȩ�޲����õ���
    '���:strNote=������������,������ʾ������ʱ�����⴦��
    '     str����IDs=����ʱ��������������Ĳ���ID��,��Ϊ���в���
    '   blnNotMsgBox��True-����ʾ��Ϣ��ʾ��;False-��ʾ��Ϣ��ʾ��
    '����:strOutErrMsg=���ش�����Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-10-17 18:07:00
    '˵��:��Ҫ�ǲ��˳�Ժ(��Ԥ��Ժ)��,���û��Ȩ��,���������
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnOut As Boolean
    Dim strInfo As String
    
    str����IDs = ""
    If InStr(strPriv, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(strPriv, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        BillCanBeOperate = True: Exit Function
    End If
    
    On Error GoTo errH
    
    '����޶�Ӧ��ҳ,�����ѳ�Ժ����(�����ﲡ��ҽ������)
    If strNote Like "*����" Then
        '���ʲ���ʱ,ֻ�Կ������ʲ������ݽ����ж�
        strSQL = _
            " Select ��� From סԺ���ü�¼" & _
            " Where ��¼����=[2] And NO=[1] And Nvl(ִ��״̬,0)<>1 And �۸񸸺� is NULL" & _
            " Group by ��� Having Nvl(Sum(Nvl(����,1)*����),0)<>0"
    ElseIf strNote Like "*���" Then
        '��˲���ʱ,ֻ��δ��˲������ݽ����ж�
        strSQL = _
            " Select ��� From סԺ���ü�¼" & _
            " Where ��¼����=2 And �۸񸸺� is NULL And ��¼״̬=0 And NO=[1]"
    End If
    strSQL = "Select Distinct ����,����ID,��ҳID From סԺ���ü�¼" & _
        " Where ��¼����=[2] And NO=[1] And ��¼״̬ IN(0,1,3)" & _
        IIf(strTime <> "", " And �Ǽ�ʱ��=[3]", "") & _
        IIf(strSQL <> "", " And Nvl(�۸񸸺�,���) IN(" & strSQL & ")", "")

    strSQL = "Select B.����ID,B.����," & _
    " Decode(A.����ID,NULL,Sysdate,A.��Ժ����) as ��Ժ����," & _
    " Nvl(A.״̬,0) as ״̬,Nvl(C.�������,0) as ���" & _
    " From ������ҳ A,(" & strSQL & ") B,������� C" & _
    " Where B.����ID=A.����ID(+) And C.����(+)=1 And C.����(+)=2  And B.��ҳID=A.��ҳID(+) And B.����ID=C.����ID(+) And C.����(+)=1 And C.����(+)=2 "
        
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytType, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytType)
    End If
    
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!��Ժ����) Or rsTmp!״̬ = 3 Then
            If rsTmp!��� = 0 And InStr(strPriv, ";��Ժ����ǿ�Ƽ���;") = 0 Then
                strInfo = strInfo & vbCrLf & "����""" & rsTmp!���� & """�ѳ�Ժ(��Ԥ��Ժ)�ҷ����Ѿ����塣"
            ElseIf rsTmp!��� <> 0 And InStr(strPriv, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
                strInfo = strInfo & vbCrLf & "����""" & rsTmp!���� & """�ѳ�Ժ(��Ԥ��Ժ)�ҷ�����δ���塣"
            Else
                str����IDs = str����IDs & "," & rsTmp!����ID
            End If
        Else
            str����IDs = str����IDs & "," & rsTmp!����ID
        End If
        rsTmp.MoveNext
    Loop
    str����IDs = Mid(str����IDs, 2)
        
    'ֻ�м��ʱ����ʿ��Բ��ݼ���
    If strInfo <> "" Then
        strOutErrMsg = Mid(strInfo, 3) & vbCrLf & "��û��Ȩ�޶Ե���""" & strNO & """����" & strNote & "��"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    BillCanBeOperate = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function



Public Function GetBillingBalanceStatu(ByVal int��Դ As Integer, ByVal strNO As String, Optional ByVal blnAll As Boolean = True, _
    Optional ByVal strTime As String, Optional ByVal bytFlag As Byte = 2, Optional ByRef intOutStatu As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ�ż��ʵ�/��Ľ���״̬
    '��Σ�int��Դ-1-����;2-סԺ
    '      strNO=���ʵ��ݺ�,�������ＰסԺ
    '      blnALL=�Ƿ�����ŵ������ݽ����ж�,����ֻ��δ���ʲ��ֽ����ж�
    '����:intOutStatu-���ؽ���״̬��0-δ����,1=��ȫ������,2-�Ѳ��ֽ���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-10-18 10:50:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    
    On Error GoTo errH
    intOutStatu = 0
    
    '��δ���ϵķ�����
    strSQL = _
        " Select ��� From (" & _
            " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���, Avg(Nvl(����, 1) * ����) As ����" & _
            " From " & IIf(int��Դ = 1, "������ü�¼", "סԺ���ü�¼") & _
            " Where NO=[1] And ��¼����=[2]" & _
            " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    
    '��ÿ�еĽ������
    strSQL = _
        "Select Nvl(�۸񸸺�,���) as ���,Sum(Nvl(���ʽ��,0)) as ���ʽ��" & _
        " From " & IIf(int��Դ = 1, "������ü�¼", "סԺ���ü�¼") & _
        " Where NO=[1] And mod(��¼����,10)= [2]" & _
        IIf(Not blnAll, " And Nvl(�۸񸸺�,���) IN(" & strSQL & ")", "") & _
        IIf(strTime <> "", " And �Ǽ�ʱ��=[3]", "") & _
        " Group by Nvl(�۸񸸺�,���)"
    
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytFlag)
    End If
    
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '��������
        rsTmp.Filter = "���ʽ��<>0"
        If rsTmp.EOF Then
            intOutStatu = 0 '�޽�����
        ElseIf rsTmp.RecordCount = lngTmp Then
            intOutStatu = 1 'ȫ�����ѽ���
        ElseIf rsTmp.RecordCount > 0 Then
            intOutStatu = 2 '�������ѽ���
        End If
    End If
    GetBillingBalanceStatu = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
 
Public Function zlCheckIsExistsApplied(ByVal strNO As String, ByVal str��� As String, _
    ByRef str����IDs As String, Optional ByRef str������s As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʵ��������Ƿ������������
    '���:strNo-���ݺ�
    '       str���-�������ʵ����(Ϊ��Ϊ����)
    '����:str����IDs-����ķ���ID
    '����:������������,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-20 15:51:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct A.ID,B.������ " & _
    "   From סԺ���ü�¼ A,���˷������� B  " & _
    "   Where A.ID=B.����ID and A.NO=[1] and A.��¼����=2 And nvl(B.״̬,0)=0 " & IIf(str��� <> "", " and Instr([2],','||���||',')>0 ", "")
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ����״̬", strNO, "," & str��� & ",")
    If rsTemp.EOF Then
        rsTemp.Close: Set rsTemp = Nothing: Exit Function
    End If
    str������s = "": str����IDs = ""
    With rsTemp
        Do While Not .EOF
            str����IDs = str����IDs & "," & Val(Nvl(rsTemp!ID))
            If InStr(1, str������s & vbCrLf, vbCrLf & Nvl(rsTemp!������) & vbCrLf) = 0 Then
                str������s = str������s & vbCrLf & Nvl(rsTemp!������)
            End If
            .MoveNext
        Loop
    End With
    If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    zlCheckIsExistsApplied = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function GetBillInsures(strInsure As String, ByVal strNO As String, _
    Optional ByVal strTime As String, Optional ByVal blnAuditing As Boolean, _
    Optional ByVal blnGetNoneInsure As Boolean, Optional ByVal bytFlag As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ʱ��е����മ"10,20,30,...",Ҳ�����ڼ��ʵ�
    '���:��strNO=���ʵ��ݺ�
    '      blnAuditing=�Ƿ����ڼ������,ֻ���δ��˵Ĳ�������
    '      blnGetNoneInsure=�Ƿ񽫷Ǳ��շ��÷���Ϊ0����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-10-18 13:44:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strInsure = ""
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(B.����,0) as ����" & _
        " From סԺ���ü�¼ A,������ҳ B" & _
        " Where A.��¼����=[2] And A.��¼״̬" & IIf(blnAuditing, "=0", " IN(0,1,3)") & _
            IIf(blnGetNoneInsure, "", " And B.���� is Not NULL") & _
        " And A.NO=[1] And A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
        IIf(strTime <> "", " And A.�Ǽ�ʱ��=[3]", "")
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���ʵ�����ر�����Ϣ", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���ʵ�����ر�����Ϣ", strNO, bytFlag)
    End If
    
    Do While Not rsTmp.EOF
        strInsure = strInsure & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    strInsure = Mid(strInsure, 2)
    GetBillInsures = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetPriceGradeStartType() As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�۸�ȼ�����������
    '����:
    '   0-δ����
    '   1-ֻ������վ��
    '   2-ֻ������ҽ�Ƹ��ʽ
    '   3-վ���ҽ�ƿʽ��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandler
    GetPriceGradeStartType = 0
    strSQL = _
        " Select Nvl(Max(Decode(b.վ��, Null, 0, 1)), 0) As ����վ��," & vbNewLine & _
        "        Nvl(Max(Decode(b.ҽ�Ƹ��ʽ, Null, 0, 1)), 0) As ����ҽ�Ƹ��ʽ" & vbNewLine & _
        " From �շѼ۸�ȼ� A, �շѼ۸�ȼ�Ӧ�� B" & vbNewLine & _
        " Where a.���� = b.�۸�ȼ� And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Ч�ļ۸�ȼ�����")
    If rsTmp.EOF Then Exit Function
    
    If Val(Nvl(rsTmp!����վ��)) = 1 Then
        If Val(Nvl(rsTmp!����ҽ�Ƹ��ʽ)) = 1 Then
            GetPriceGradeStartType = 3 'վ���ҽ�ƿʽ��������
        Else
            GetPriceGradeStartType = 1 'ֻ������վ��
        End If
    Else
        If Val(Nvl(rsTmp!����ҽ�Ƹ��ʽ)) = 1 Then
            GetPriceGradeStartType = 2 'ֻ������ҽ�Ƹ��ʽ
        End If
    End If
    Exit Function
errHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetPriceGrade(ByVal strվ�� As String, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    Optional ByVal strҽ�Ƹ��ʽ As String, _
    Optional ByRef strҩƷ�۸�ȼ�_Out As String, _
    Optional ByRef str���ļ۸�ȼ�_Out As String, _
    Optional ByRef str��ͨ��Ŀ�۸�ȼ�_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ����ҽ�Ƹ��ʽ��վ�㣬��ȡ��Ӧ�ļ۸�ȼ�
    '���:strվ��-��½��վ�㣬���봫�룬����NULLʱ���۸�ȼ�Ϊ���ؿ�
    '     lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     strҽ�Ƹ��ʽ:�������ǿգ����Դ���ҽ�Ƹ��ʽ_In��ʽ����ȡ�۸�ȼ�;�����Բ���ID_In����ҳID����ȡ��Ӧ�Ĳ��˵�ҽ�Ƹ��ʽ��
    
    '����:strҩƷ�۸�ȼ�_out-����ҩƷ�۸�ȼ�
    '     str���ļ۸�ȼ�_out-�������ļ۸�ȼ�
    '     str��ͨ��Ŀ�۸�ȼ�_out-������ͨ�շ���Ŀ�۸�ȼ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim varPara As Variant
    
    strҩƷ�۸�ȼ�_Out = "": str���ļ۸�ȼ�_Out = "": str��ͨ��Ŀ�۸�ȼ�_out = ""
    On Error GoTo errHandle
    '    Zl_Get_Pricegrade
    '  վ��_In         In �շѼ۸�ȼ�����.վ��%Type,
    '  ����id_In       In ������Ϣ.����id%Type := Null,
    '  ��ҳid_In       In ������ҳ.��ҳid%Type := Null,
    '  ҽ�Ƹ��ʽ_In In �շѼ۸�ȼ�����.��������%Type := Null
    
    strSQL = "Select Zl_Get_Pricegrade([1],[2],[3],[4]) as �۸�ȼ� From dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ�۸�ȼ�", strվ��, lng����ID, lng��ҳID, strҽ�Ƹ��ʽ)
    If Nvl(rsTemp!�۸�ȼ�) = "" Then GetPriceGrade = True: Exit Function
    '��ʽ:��ͨ�۸�ȼ�|ҩƷ�۸�ȼ�|�������ϼ۸�ȼ�
    varPara = Split(rsTemp!�۸�ȼ� & "||||", "|")
    
    str��ͨ��Ŀ�۸�ȼ�_out = varPara(0)
    strҩƷ�۸�ȼ�_Out = varPara(1)
    str���ļ۸�ȼ�_Out = varPara(2)
    GetPriceGrade = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetRetailPrice(ByVal lng�շ�ϸĿID As Long, _
    ByVal str�۸�ȼ� As String, ByRef dbl���ۼ�_out As Double, ByRef dblδ�ֽ���_out As Double, _
    Optional ByVal lng�ⷿID As Long = 0, _
    Optional ByVal dbl���� As Double = 0) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ���ݼ۸�ȼ���ȡָ���շ���Ŀ�����ۼ۵������Ϣ
    '���:lng�շ�ϸĿid-�շ�ϸĿID
    '     str�۸�ȼ�-�շѼ۸�ȼ�
    '     lng�ⷿid-�ⷿID��ҩƷ���������ϴ���)
    '     dbl����:��ǰ��������(ҩƷ���������ϴ���)��
    '����:dbl���ۼ�_out-�������ۼ۸�
    '     dblδ�ֽ���_out-���ҩƷ������������Ч����ʾ���ݵ�ǰ����ĳ���������������dbl����)���зֽ�ʱ��δ�ֽ��������.
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim varPara As Variant
    
    On Error GoTo errHandle
    'Zl_Get_Retailprice
    '  �շ�ϸĿid_In In �շ���ĿĿ¼.Id%Type,
    '  �۸�ȼ�_In   In �շѼ۸�ȼ�.����%Type,
    '  �ⷿid_In     In ���ű�.Id%Type := 0,
    '  ����_In       In Number := 0
    ') Return Varchar2
    '  --      a.ҩƷ����������:���ۼ�|δ�ֽ���
    '  --      b.�����ͨ����:���ۼ�(ʵ��Ϊȱʡ�۸�)|0
  
    strSQL = "Select Zl_Get_Retailprice([1],[2],[3],[4]) as �۸� From dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "���ݼ۸�ȼ���ȡ�۸�", lng�շ�ϸĿID, str�۸�ȼ�, lng�ⷿID, dbl����)
    If Nvl(rsTemp!�۸�) = "" Then Exit Function
    
    varPara = Split(rsTemp!�۸� & "||||", "|")
    dbl���ۼ�_out = FormatEx(varPara(0), 5)
    dblδ�ֽ���_out = FormatEx(varPara(1), 5)
    GetRetailPrice = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
