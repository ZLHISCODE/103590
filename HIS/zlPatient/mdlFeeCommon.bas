Attribute VB_Name = "mdlFeeCommon"
Option Explicit
Private mrs���� As ADODB.Recordset
Public Type Ty_FactProperty
    lngShareUseID As Long   '������������ID
    strUseType As String ' ʹ�����
    intInvoiceFormat As Integer '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
    intInvoicePrint As Integer     '��ӡ��ʽ:0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
End Type
Public grsҽ�Ƹ��ʽ As ADODB.Recordset
Public Type TY_PatiMaxLenInfor
    intPatiName As Integer  '������󳤶�
    intPatiAge  As Integer   '������󳤶�
    intPatiSex As Integer   '�Ա���󳤶�
    intPatiMzNo As Integer   '�������󳤶�
End Type
Public grsOneCard As ADODB.Recordset

Private gPatiMaxLen As TY_PatiMaxLenInfor

 Public Function zlGetPatiInforMaxLen() As TY_PatiMaxLenInfor
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ����󳤶�
    '����TY_PatiMaxLenInfor
    '����:���˺�
    '����:2013-11-11 11:44:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If gPatiMaxLen.intPatiName <> 0 Then
        zlGetPatiInforMaxLen = gPatiMaxLen: Exit Function
    End If
    With gPatiMaxLen
        .intPatiName = 100
        .intPatiMzNo = 18
        .intPatiAge = 20
        .intPatiSex = 4
    End With
    '�����ݿ��ж�ȡ
    
    strSQL = "" & _
    "   Select /*+ rule */  A.Column_Name ,Nvl(A.Data_Precision, A.Data_Length) as PatiMaxLen " & _
    "   From All_Tab_Columns A,Table(f_str2list([2])) J " & _
    "   Where A.Table_Name = [1] And A.Column_Name=J.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, "������Ϣ", "����,�����,�Ա�,����")
    With rsTemp
        Do While Not .EOF
            Select Case nvl(!Column_Name)
            Case "����"
                gPatiMaxLen.intPatiName = Val(nvl(rsTemp!PatiMaxLen))
            Case "�����"
                gPatiMaxLen.intPatiMzNo = Val(nvl(rsTemp!PatiMaxLen))
            Case "�Ա�"
                gPatiMaxLen.intPatiSex = Val(nvl(rsTemp!PatiMaxLen))
            Case "����"
                gPatiMaxLen.intPatiAge = Val(nvl(rsTemp!PatiMaxLen))
            End Select
            .MoveNext
        Loop
    End With
    rsTemp.Close: Set rsTemp = Nothing
    zlGetPatiInforMaxLen = gPatiMaxLen: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetColumnLength(strTable As String, strColumn As String) As Long
    GetColumnLength = Sys.FieldsLength(strTable, strColumn)
End Function
Public Function zlExcuteUploadSwap(ByVal lng����ID As Long, ByRef strOutPut As String, Optional objExcuteObject As Object = Nothing) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����UploadSwap�ӿ�
    '���:strCardNo
    '     objExcuteObject-���õĶ���
    '����:
    '����:���óɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2009-07-24 10:32:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnNothing As Boolean, rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0: On Error GoTo Errhand:
    strSQL = "Select ��� From һ��ͨĿ¼ where nvl(����,0)=2 and rownum<=1"
    If mrs���� Is Nothing Then
        Set mrs���� = zlDatabase.OpenSQLRecord(strSQL, "���һ��ͨ")
    ElseIf mrs����.State <> 1 Then
        Set mrs���� = zlDatabase.OpenSQLRecord(strSQL, "���һ��ͨ")
    End If
    If mrs����.EOF Then zlExcuteUploadSwap = True: Exit Function
    
    If objExcuteObject Is Nothing Then
        Set objExcuteObject = CreateObject("zlICCard.clsICCard")
        Set objExcuteObject.gcnOracle = gcnOracle
        blnNothing = True
    End If
    If objExcuteObject Is Nothing Then Exit Function
    'UploadSwap(ByVal strCardNO As String, ByVal lng����ID As Long, ByRef strOut As String) As Boolean'Ŀǰֻ��,û��ʲô����ֵ
    Call objExcuteObject.UploadSwap(lng����ID, strOutPut)
    If blnNothing Then Set objExcuteObject = Nothing
    
    zlExcuteUploadSwap = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
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

Public Sub SetNOInputLimit(ByRef objThis As Object, ByRef KeyAscii As Integer, Optional BytType As Byte)
'����:�����ݺŻ�Ʊ�ݺ�����ؼ��Ŀ�����ֵ,Ŀǰ���ݺ������һλ����ĸ,�����������,Ʊ�ݺ�����ǰ��λ����ĸ������,�����������
'����:objThis:������txtbox�����ֵ��combox
'     bytType:0-���ݺ�,1-Ʊ�ݺ�
    Dim strAbc As String, str123 As String
    Dim str1 As String, str2 As String
    
    strAbc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    str123 = "0123456789"
    str1 = Mid(objThis.Text, 1, 1): str1 = IIf(str1 = "", "��", str1)
    str2 = Mid(objThis.Text, 2, 1): str2 = IIf(str2 = "", "��", str2)
        
    If BytType = 0 Then
        Call zlControl.TxtCheckKeyPress(objThis, KeyAscii, m�ı�ʽ)
    Else
        If objThis.Text = "" Or objThis.SelLength = Len(objThis.Text) Or _
            objThis.SelStart = 0 And (objThis.SelLength > 0 Or InStr(strAbc, str1) = 0 Or InStr(strAbc, str1) > 0 And InStr(strAbc, str2) = 0) Or _
            objThis.SelStart = 1 And (objThis.SelLength > 0 Or InStr(strAbc, str1) > 0 And InStr(strAbc, str2) = 0) Then
            
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            
            '������������ĸ���Һ���������,ѡ�е�һ����ĸʱ,ֻ������ĸ
            '����һ����ĸ,λ���ڵ�һ����ĸ֮ǰʱ,ֻ������ĸ
            If objThis.SelStart = 0 And objThis.SelLength = 1 And InStr(strAbc, str2) > 0 And objThis.SelLength <> Len(objThis.Text) Or _
               objThis.SelStart = 0 And objThis.SelLength = 0 And InStr(strAbc, str1) > 0 And InStr(strAbc, str2) = 0 Then
                If InStr(strAbc & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            Else
                If InStr(str123 & strAbc & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        Else
            '����������ĸ,λ���ڵ�һ����ĸ֮ǰ��������ĸ֮��ʱ,����������
            If (objThis.SelStart = 0 Or objThis.SelStart = 1) And objThis.SelLength = 0 And InStr(strAbc, str1) > 0 And InStr(strAbc, str2) > 0 Then
                If objThis.SelStart = 1 Then    '����ɾ����һ����ĸ
                    If Chr(8) <> Chr(KeyAscii) Then KeyAscii = 0: Beep: Exit Sub
                Else
                    KeyAscii = 0: Beep: Exit Sub
                End If
            Else
                If InStr(str123 & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        End If
    End If
End Sub

Public Function ActualMoney(str�ѱ� As String, ByVal lng������ĿID As Long, ByVal curӦ�ս�� As Currency, _
    Optional ByVal lng�շ�ϸĿID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal dbl���� As Double, Optional ByVal dbl�Ӱ�Ӽ��� As Double) As Currency
'���ܣ������շ�ϸĿID��������ĿID(ǰ������),Ӧ�ս��,���ѱ����õķֶα������۹������ʵ�ս�
'       ���ҩƷ���ɱ����ձ����������ʵ�ս��
'������str�ѱ�=���˷ѱ�����ǰ���̬�ѱ�,�����ʽΪ"���˷ѱ�,��̬�ѱ�1,��̬�ѱ�2,..."
'      lng�ⷿID,dbl����,��ҩƷ����Ŀ���ɱ��ۼ��մ���ʱ����Ҫ����
'      dbl����=�����������ڵ��ۼ�����
'      dbl�Ӱ�Ӽ���=С������,�����Ӧ�ս���Ѱ��Ӱ�Ӽۼ���ʱ��Ҫ�����ڻ�ԭ������
'���أ������۹���ͱ��������ʵ�ս��,����Ƕ�̬�ѱ�,��"str�ѱ�"�������Żݷѱ�(ע�����δ���ۼ���,����ԭ������,Ҳ���ܷ��ص�һ��)
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
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str�ѱ�, lng�շ�ϸĿID, lng������ĿID, curӦ�ս�� / (1 + dbl�Ӱ�Ӽ���), dbl����, lng�ⷿID)
        
    str�ѱ� = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl�Ӱ�Ӽ���), gstrDec)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetActualMoney(ByVal str�ѱ� As String, ByVal lng����ID As Long, ByVal curӦ�� As Currency, ByVal lng�շ�ϸĿID As Long) As Currency
'���ܣ�����ָ���ķѱ��������Ŀ���շ���Ŀ,����ָ������ʵ���տ���
'������
'   str�ѱ�   ���ѱ�
'   lng����ID  ��������ĿID
'   curӦ�գ�Ӧ�ս��ֵ
'���أ�ʵ��Ӧ�յĽ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select ʵ�ձ���" & vbNewLine & _
            "From �ѱ���ϸ" & vbNewLine & _
            "Where �ѱ� = [1] And �շ�ϸĿid = [3] And Abs([4]) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select ʵ�ձ���" & vbNewLine & _
            "From �ѱ���ϸ A" & vbNewLine & _
            "Where �ѱ� = [1] And ������Ŀid = [2] And Abs([4]) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And Not Exists" & vbNewLine & _
            " (Select 1 From �ѱ���ϸ C Where C.�ѱ� = A.�ѱ� And C.�շ�ϸĿid = [3])"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str�ѱ�, lng����ID, lng�շ�ϸĿID, curӦ��)
    If rsTmp.EOF Then
        GetActualMoney = curӦ��
    Else
        GetActualMoney = curӦ�� * rsTmp!ʵ�ձ��� / 100
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function ReturnMovedExes(ByVal strNO As String, ByVal BytType As Byte, Optional ByVal strFormCaption As String) As Boolean
'����:�����û�ѡ���ѡ�����ݱ��е����ݵ���ǰ���ݱ���
'����:bytType��ʾ��������,ֵ::1-�շ�,2-����,3-�Զ�����,4-�Һ�,5-���￨,6-Ԥ��,7-���ʣ�
'����:�û�ѡ��ȡ������,���߳�ѡ����ת��ʧ��,�򷵻�False

    MsgBox "��ǰ�����ĵ���" & strNO & "�ں����ݱ���!" & vbCrLf _
        & "����ϵͳ����Ա��ϵ,ת�뵽�������ݱ��ٲ���!", vbInformation, gstrSysName
    ReturnMovedExes = False

'�����ǳ�ѡ�������ݵĹ��̣��ݴ棬���ڽ���͸������ʱ����
    
'    If MsgBox("��ǰ��������" & strNO & "�ں����ݱ���,ϵͳ��Ҫ�Ȱ���˵�����ص�����ת�뵽�������ݱ���ܼ���!" & vbCrLf & _
'                             "ȷ��Ҫ���д˲�����?", vbInformation + vbYesNo, gstrSysName) = vbNo Then
'        ReturnMovedExes = False     '�˾��ʡ
'        Exit Function
'    End If
'
'    If zlDatabase.ReturnMovedExes(strNO, bytType, strFormCaption) Then
'        ReturnMovedExes = True
'    Else
'        '��ϸ������֮ǰ��ִ�й��̳���ʱ����
'        MsgBox "��ϵͳ����,��õ�����ص�����δ��ת�뵽�������ݱ�." & vbCrLf & "����δ�ɹ�,����ϵͳ����Ա��ϵ!", vbInformation, gstrSysName
'        ReturnMovedExes = False
'    End If
End Function

Public Function OverTime(Curdate As Date) As Boolean
'���ܣ��жϵ�ǰ�Ƿ��ڼӰ�ʱ�䷶Χ��
'���أ���-��ǰ���ڼӰ�ʱ����,��-������
    Dim curTime As Date, DateBegin As Date, DateEnd As Date
    Dim str���� As String, str���� As String
    
    curTime = CDate(Format(Curdate, "HH:MM:SS"))
    
    str���� = zlDatabase.GetPara(1, glngSys)
    If str���� <> "" Then
        DateBegin = CDate(Trim(Split(UCase(str����), "AND")(0)))
        DateEnd = CDate(Trim(Split(UCase(str����), "AND")(1)))
    End If
    
    If Not (curTime >= DateBegin And curTime <= DateEnd) Then
        str���� = zlDatabase.GetPara(2, glngSys)
        If str���� <> "" Then
            DateBegin = CDate(Trim(Split(UCase(str����), "AND")(0)))
            DateEnd = CDate(Trim(Split(UCase(str����), "AND")(1)))
        End If
        
        If Not (curTime >= DateBegin And curTime <= DateEnd) Then OverTime = True
    End If
End Function

Public Function GetInsureName(intInsure As Integer) As String
'���ܣ����ݱ��������Ż�ȡ�����������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ������� Where ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, intInsure)  'һ�����������SQL�����壬����ͬʱ���Ӷ��ҽ��ʱ���е������
    If Not rsTmp.EOF Then
        GetInsureName = "" & rsTmp!����
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStockCheck(ByVal BytType As Byte) As Collection
'���ܣ���ȡҩƷ�����ĳ�����ļ���
'������bytType:0-ҩƷ��1-����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '�������
    
    strSQL = _
        " Select Distinct A.ID,C.��鷽ʽ" & _
        " From ���ű� A,��������˵�� B," & IIf(BytType = 0, "ҩƷ������", "���ϳ�����") & " C" & _
        " Where B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� " & IIf(BytType = 0, "IN('��ҩ��','��ҩ��','��ҩ��')", "='���ϲ���'") & _
        " And C.�ⷿID(+)=A.ID"
        '26046:վ��ȡ��.
        '"   And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlFeeCommon")
    For i = 1 To rsTmp.RecordCount
        colStock.Add nvl(rsTmp!��鷽ʽ, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function Get���㷽ʽ(str���� As String, Optional str���� As String) As ADODB.Recordset
    Dim strSQL As String, strIF As String
    
    On Error GoTo errH
    
    If str���� <> "" Then
        If InStr(1, str����, ",") > 0 Then
            strIF = "And Instr(','||[2]||',',','||B.����||',')>0 "
        Else
            strIF = "And B.���� = [2]"
        End If
    End If
    strSQL = _
        " Select B.����,B.����,Nvl(Nvl(A.ȱʡ��־,B.ȱʡ��־),0) as ȱʡ,Nvl(B.����,1) as ����,Nvl(B.Ӧ����,0) as Ӧ����" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ " & _
        " And (B.����<>7 Or B.����=7 And Exists(Select 1 From һ��ͨĿ¼ C Where C.���㷽ʽ=B.���� And C.����=1))   " & strIF
    If InStr(1, str����, ",9") > 0 Then
        strSQL = strSQL & " Union " & _
                 " Select ����,����,Nvl(ȱʡ��־,0) As ȱʡ,Nvl(����,1) as ����,Nvl(Ӧ����,0) as Ӧ���� " & _
                 " From ���㷽ʽ " & _
                 " Where ����=9 " & _
                 " Order by ����,����"
    Else
        strSQL = strSQL & " Order by ����,lpad(����,3,' ')"
    End If
    Set Get���㷽ʽ = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str����, str����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get����ҽ�Ƹ��ʽ(lng����ID As Long, Optional lng��ҳID As Long) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If lng��ҳID = 0 Then
        strSQL = "Select ҽ�Ƹ��ʽ From ������Ϣ Where ����ID=[1]"
    Else
        strSQL = "Select ҽ�Ƹ��ʽ From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then Get����ҽ�Ƹ��ʽ = "" & rsTmp!ҽ�Ƹ��ʽ
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMedPayMode(ByVal strName As String, ByRef rsMedPayMode As ADODB.Recordset) As Byte
'���ܣ�����ҽ�Ƹ��ʽ���Ʒ��������
    Dim strSQL As String
    
    On Error GoTo errH
    
    If rsMedPayMode Is Nothing Then
        strSQL = "Select ����,����,ȱʡ��־ From ҽ�Ƹ��ʽ"
        Set rsMedPayMode = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    End If
    rsMedPayMode.Filter = "����='" & strName & "'"
    If rsMedPayMode.RecordCount > 0 Then GetMedPayMode = Val(rsMedPayMode!����)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMedPayModeName(ByVal strCode As String) As String
'���ܣ�����ҽ�Ƹ��ʽ���뷵��������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strCode)
        
    If rsTmp.RecordCount > 0 Then GetMedPayModeName = rsTmp!����
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiWarnRange(ByVal lngPatient As Long, ByVal lngPage As Long) As String
'���ܣ���ȡ���˱������÷�Χ,���ڼ��ʱ��б���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select Zl_Patiwarnscheme([1], [2]) As ���ò��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngPatient, lngPage)
        
    If rsTmp.RecordCount > 0 Then GetPatiWarnRange = rsTmp!���ò���
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnitWarn(Optional ByVal str���ò��� As String, Optional ByVal str����ID As String) As ADODB.Recordset
'���ܣ����ز������ʱ�����¼��
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(����ID,0) ����ID,���ò���,Nvl(��������,1) as ��������," & _
            " ����ֵ,������־1,������־2,������־3" & _
            " From ���ʱ����� Where 1=1" & _
            IIf(str���ò��� = "", "", " And ���ò��� = [1]") & _
            IIf(str����ID = "", "", " And Nvl(����ID,0) = [2]")
    Set GetUnitWarn = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str���ò���, str����ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.�������� = "" & rsTmp!������
        UserInfo.���� = "" & rsTmp!����
        UserInfo.���� = "" & rsTmp!����
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPersonnel(str���� As String, Optional blnBaseInfo As Boolean) As ADODB.Recordset
'���ܣ���ȡָ�����ʵ���Ա�б�
    Dim strSQL As String
    On Error GoTo errH
    
    If str���� <> "" Then
        If blnBaseInfo Then
            strSQL = "Select a.id,a.���,a.����,a.���� From ��Ա�� a,��Ա����˵�� b" & _
            " Where a.ID = b.��ԱID And b.��Ա����=[1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by a.����"
        Else
            strSQL = "Select a.Id, a.���, a.����, a.����, a.���֤��, a.��������, a.�Ա�, a.����, a.��������, a.�칫�ҵ绰, a.�����ʼ�, a.ִҵ���, a.ִҵ��Χ, " & _
                    "a.����ְ��, a.רҵ����ְ��, a.Ƹ�μ���ְ��, a.ѧ��, a.��ѧרҵ, a.��ѧʱ��, a.��ѧ����, a.������ѵ, a.���п���, a.���˼��, a.����ʱ��, " & _
                    "a.����ʱ��, a.����ԭ��, a.����, a.վ�� From ��Ա�� a,��Ա����˵�� b" & _
            " Where a.ID = b.��ԱID And b.��Ա����=[1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by a.����"
        End If
    Else
        If blnBaseInfo Then
            strSQL = "Select id,���,����,���� From ��Ա�� A" & _
            " Where (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by ����"
        Else
            strSQL = zlGetFullFieldsTable("��Ա��", 0, "", False) & _
            " Where (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by ����"
        End If
    End If
    Set GetPersonnel = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPersonnelID(str���� As String, Optional ByRef rs��Ա As ADODB.Recordset) As Long
'���ܣ�������Ա��������ID
'˵�����鿴�շѵ�ʱ��������(ҽ��)���������Ѳ���ҽ���ˣ���mrs�������в�����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If str���� = "" Then Exit Function
    
    If Not rs��Ա Is Nothing Then
        rs��Ա.Filter = "����='" & str���� & "'"
        If rs��Ա.RecordCount > 0 Then GetPersonnelID = rs��Ա!ID: Exit Function
    End If
    
    On Error GoTo errH
    strSQL = "Select ID from ��Ա�� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str����)
    If Not rsTmp.EOF Then GetPersonnelID = rsTmp!ID
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDepartments(ByVal str���� As String, _
    ByVal str������� As String, _
    Optional ByVal bln������Ա���� As Boolean = False, _
    Optional ByVal blnCheckվ�� As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ʵĲ����б�
    '���:str����='�ٴ�','����','��ҩ��',...,����Ϊ��
    '     str�������:��,����:��1,3
    '     bln������Ա����-����Ա����������
    '����:
    '����:
    '����:���˺�
    '����:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    str���� = Replace(str����, "'", "")
    If str���� <> "" Then
        If InStr(1, str����, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.��������||',')>0"
        Else
            strSQL = " And B.�������� = [1]"
        End If
    End If
    If bln������Ա���� Then strSQL = strSQL & "  And A.id=C.����ID and C.��Աid =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        " From ���ű� A,��������˵�� B " & IIf(bln������Ա����, ",������Ա C", "") & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And Instr(',' || [2]|| ',',',' || B.������� || ',')>0 " & strSQL & _
         IIf(blnCheckվ��, " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
        " Order by A.����"
    Set GetDepartments = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str����, str�������, UserInfo.ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'�÷����� GetDepartments �������ƣ��÷���ȡ����վ������
Public Function GetDepts(ByVal str���� As String, ByVal str������� As String, Optional ByVal bln������Ա���� As Boolean = False) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ʵĲ����б�
    '���:str����='�ٴ�','����','��ҩ��',...,����Ϊ��
    '     str�������:��,����:��1,3
    '     bln������Ա����-����Ա����������
    '����:
    '����:
    '����:���˺�
    '����:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    str���� = Replace(str����, "'", "")
    If str���� <> "" Then
        If InStr(1, str����, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.��������||',')>0"
        Else
            strSQL = " And B.�������� = [1]"
        End If
    End If
    If bln������Ա���� Then strSQL = strSQL & "  And A.id=C.����ID and C.��Աid =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        " From ���ű� A,��������˵�� B " & IIf(bln������Ա����, ",������Ա C", "") & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And Instr(',' || [2]|| ',',',' || B.������� || ',')>0 " & strSQL & _
        " Order by A.����"
    Set GetDepts = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str����, str�������, UserInfo.ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnitDept() As ADODB.Recordset
'���ܣ���ȡ�������Ҷ�Ӧ��ϵ
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ����ID, ����ID From �������Ҷ�Ӧ"
    Set GetUnitDept = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Public Function GetDeptOrUnit(ByVal BytType As Byte, lngDept As Long, ByVal strServiceRange As String) As ADODB.Recordset
'���ܣ���ȡָ�������Ŀ���,��ָ�����ҵĲ���
'������bytType=0-ָ�������Ŀ���,1-ָ�����ҵĲ���
'      strServiceRange=�������1-���2-סԺ��3�������סԺ
'       lngDept=����ID����ID
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,�������Ҷ�Ӧ B,��������˵�� C " & _
            " Where " & IIf(BytType = 0, "B.����ID=A.ID And B.����ID", "B.����ID=A.ID And B.����ID") & "=[1] " & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And C.����ID=A.ID And Instr(',' || [2]|| ',',',' || C.������� || ',')>0 " & _
            " And C.��������=" & IIf(BytType = 0, "'�ٴ�'", "'����'") & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Set GetDeptOrUnit = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", lngDept, strServiceRange)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function StringDelItem(ByVal strAll As String, ByVal strItem As String, Optional strSplit As String = ",") As String
'���ܣ���ָ�����ַ����б���ɾ��һ��(����ж��ƥ���,ֻ�Ƴ���һ��)
    Dim i As Long, arrTmp As Variant
    
    arrTmp = Split(strAll, strSplit)
    For i = 0 To UBound(arrTmp)
        If arrTmp(i) = strItem Then
            strItem = ""
        Else
            StringDelItem = StringDelItem & "," & arrTmp(i)
        End If
    Next
    StringDelItem = Mid(StringDelItem, 2)
End Function

Public Function GetOneCardBalance(ByVal lng����ID As Long) As ADODB.Recordset
'���ܣ���ȡһ��ͨ�����¼
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select A.��λ�ʺ�, A.�������, B.ҽԺ����, A.��Ԥ�� as ���" & vbNewLine & _
            "From ����Ԥ����¼ A, һ��ͨĿ¼ B" & vbNewLine & _
            "Where A.����id = [1] And A.���㷽ʽ = B.���㷽ʽ"

    Set GetOneCardBalance = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetOneCard() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ��ͨ���ü�¼��
    '����:����һ��ͨ���ü�¼��
    '����:���˺�
    '����:2014-07-04 10:17:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    If Not grsOneCard Is Nothing Then
        If grsOneCard.State = 1 Then
            Set GetOneCard = grsOneCard
            Exit Function
        End If
    End If
    strSQL = "Select ���,����,ҽԺ����,���㷽ʽ From һ��ͨĿ¼ Where ����=1"
    Set grsOneCard = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Set GetOneCard = grsOneCard
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ID(ByVal lng����ID As Long) As Long
'���ܣ����ݿ���ID��ȡ��Ӧ�Ĳ���ID,
'       ����ж������,ȡID��С��һ��,û���ҵ�ʱ����0
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Min(����ID) ����ID From �������Ҷ�Ӧ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    
    If Not rsTmp.EOF Then Get����ID = Val("" & rsTmp!����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnit(ByVal blnLimitUnit As Boolean, _
    ByVal strServiceRange As String, ByVal strType As String, _
    Optional bln���� As Boolean = False, _
    Optional blnNotNode As Boolean = False, _
    Optional blnShowNodeCode As Boolean = False) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���в���������б�
    '���:blnLimitUnit=�Ƿ������в���Ȩ�ޣ�û��ʱ��ֻ��ȡ����Ա�����Ŀ��һ���
    '       blnNotNode-�Ƿ�����վ��:true,������վ��,����վ��
    '       blnShowNodeCode:��ʾվ����
    '����:
    '����:������������ݼ�
    '����:���˺�
    '����:2011-02-28 17:21:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strUnitIDs As String
    Dim strWhere As String
    
    On Error GoTo errH
    If blnLimitUnit Then strUnitIDs = GetUserUnits
    strWhere = ""
    If blnNotNode = False Then strWhere = " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) "
    strSQL = _
         " Select A.ID,A.����,A.���� " & IIf(bln����, ",A.����", "") & IIf(blnShowNodeCode, ",A.վ��", "") & _
         " From ���ű� A,��������˵�� B" & _
         " Where B.����ID = A.ID And B.������� IN(" & strServiceRange & ") And B.�������� = [2]" & _
         " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            strWhere & vbNewLine & _
         IIf(blnLimitUnit, " And Instr(','||[1]||',',','||A.ID||',')>0", "") & _
         " Order by A.����"
    Set GetUnit = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strUnitIDs, strType)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����IDs(lngUnit As Long) As String
'���ܣ����ݲ����������Ӧ�Ŀ���ID��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    strSQL = "Select Distinct ����ID From �������Ҷ�Ӧ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngUnit)
    
    strSQL = "0"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If Not IsNull(rsTmp!����ID) Then
                strSQL = strSQL & "," & rsTmp!����ID
            End If
            rsTmp.MoveNext
        Next
    End If
    Get����IDs = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUserUnits(Optional ByVal blnDept As Boolean) As String
'���ܣ���ȡ��ǰ�û������в���ID����ID
'      �������Ա���ڿ���,�򷵻ؿ���ID��������������ID
'      blnDept:True��ʾ��ȡ����Ա��������,�Լ����������µ����п���,���򷵻ز���Ա��������,�Լ����ڿ��������Ĳ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    'union��ȥ���ظ���
    If blnDept Then
        strSQL = "Select A.����ID ����ID From �������Ҷ�Ӧ A,������Ա B Where A.����ID=B.����ID And B.��ԱID=[1]" & _
            " Union Select ����ID as ����ID From ������Ա Where ��ԱID=[1]"
    Else
        strSQL = "Select A.����ID ����ID From �������Ҷ�Ӧ A,������Ա B Where A.����ID=B.����ID And B.��ԱID=[1]" & _
            " Union Select ����ID as ����ID From ������Ա Where ��ԱID=[1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        GetUserUnits = GetUserUnits & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    
    If GetUserUnits = "" Then
        GetUserUnits = "0"
    Else
        GetUserUnits = Mid(GetUserUnits, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���˿���ID(lng����ID) As Long
'���ܣ���ȡ��Ժ���˵�ǰ���˿���ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.��Ժ����ID From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    If Not rsTmp.EOF Then Get���˿���ID = rsTmp!��ǰ����id
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GET��������(lngDeptID As Long, Optional ByRef rsDept As ADODB.Recordset) As String
'���ܣ���ȡ��������
'������lngDeptID=����ID
'���أ���������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If rsDept Is Nothing Then
        strSQL = "Select ���� from ���ű� Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngDeptID)
    Else
        Set rsTmp = rsDept
        rsTmp.Filter = "ID=" & lngDeptID
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select ���� from ���ű� Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngDeptID)
        End If
    End If
    
    If Not rsTmp.EOF Then GET�������� = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetPersonnelIDCode(ByVal strName As String, Optional ByRef strID As String, Optional ByRef strCode As String)
'����:������Ա������ȡ��ID�ͱ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID,���� From ��Ա�� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", strName)
    
    If Not rsTmp.EOF Then
        strID = rsTmp!ID
        strCode = rsTmp!����
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Function GetDoctorOrNurse(ByVal BytType As Byte, Optional ByVal strUnits As String) As ADODB.Recordset
'���ܣ���ȡҽ����ʿ�б�.
'������bytType=0-ҽ����1-��ʿ
'       strUnits=���һ���ID��,��:18,26,31
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If strUnits <> "" Then
        If InStr(1, strUnits, ",") > 0 Then
            strSQL = " And Instr(','|| [2] || ',',',' || C.����ID || ',')>0"
        Else
            strSQL = " And C.����ID=[2]"
        End If
    End If
    
    strSQL = _
        "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,��Ա����˵�� B,������Ա C,��������˵�� D" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.����ID=D.����ID" & _
        " And B.��Ա����=[1] And D.������� IN(1,2,3) And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & strSQL & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        " Order by ����" '
    Set GetDoctorOrNurse = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", IIf(BytType = 0, "ҽ��", "��ʿ"), strUnits)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function is����(lng����ID As Long, ByRef rs�������� As ADODB.Recordset) As Boolean
'���ܣ��жϿ����Ƿ��ǲ�������
'������lng����ID=ָ������ID
    is���� = Sys.DeptHaveProperty(lng����ID, "����")
End Function

Public Function isMediRoom(lngID As Long) As Boolean
'���ܣ��жϲ����Ƿ�ҩ��
'������lngID=����ID
     isMediRoom = Sys.DeptHaveProperty(lngID, "��ҩ��") Or Sys.DeptHaveProperty(lngID, "��ҩ��") Or Sys.DeptHaveProperty(lngID, "��ҩ��")
End Function

Public Function isCliniOrNurse(ByVal lngDept As Long) As Boolean
'����:���ݲ���ID�ж��Ƿ����ٴ�������
    isCliniOrNurse = Sys.DeptHaveProperty(lngDept, "�ٴ�") Or Sys.DeptHaveProperty(lngDept, "����")
End Function

Public Function GetNORule(ByVal intNo As Integer) As Integer
'����:��ȡָ��NO�ı�Ź���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��Ź��� From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, intNo)
    If Not rsTmp.EOF Then GetNORule = Val("" & rsTmp!��Ź���)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetShareInvoiceGroupID(ByVal bytKind As Byte) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��Ʊ�ֵĹ���Ʊ������
    '����:���˺�
    '����:2011-04-29 10:24:48
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    If bytKind = 1 Or bytKind = 3 Then  '�շѺͽ���
        strSQL = "" & _
        "   Select A.ID,nvl(M.����,' ') as ʹ��������,A.ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B,Ʊ��ʹ����� M" & vbNewLine & _
        "   Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "           And A.ʹ�����=M.����(+) " & _
        "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "   Order by ʹ��������,ʣ������ Desc"
    ElseIf bytKind = 5 Then
        '���￨
        strSQL = "" & _
        "   Select A.ID,nvl(M.����,' ') as ʹ��������,M.ID as ʹ�����ID,M.���� as ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B,ҽ�ƿ���� M" & vbNewLine & _
        "   Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "           And to_number(nvl(A.ʹ�����,'0'))=M.ID(+) " & _
        "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "   Order by ʹ��������,ʣ������ Desc"
    ElseIf bytKind = 2 Then  'Ԥ��
        strSQL = "" & _
        "   Select A.ID,to_number(nvl(A.ʹ�����,'0')) as ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B" & vbNewLine & _
        "   Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "   Order by ʹ�����,ʣ������ Desc"
    Else
        strSQL = "" & _
        "   Select A.ID,A.ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B" & vbNewLine & _
        "   Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "   Order by ʹ�����,ʣ������ Desc"
    End If
    Set GetShareInvoiceGroupID = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, bytKind)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlCheckInvoiceOverplusEnough(ByVal bytKind As Byte, _
    ByVal intNum As Integer, Optional lngʣ������ As Long, _
    Optional lng����ID As Long = 0, Optional strʹ����� As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ʊ�ݵ�ʣ�������Ƿ����
    '���:bytKind-1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '     intNum-��ǰ�Աȵ�����(-1��������)
    '     lng����ID-ֻ��鵱ǰ������Ʊ��(32455)
    '     strʹ�����-ʹ�����
    '����:lngʣ������-���ص�ǰʣ������
    '����:���㷵��true,���򷵻�False
    '����:���˺�
    '����:2009-12-28 17:16:16
    '����:26948
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '-1��������
    If intNum = -1 Then zlCheckInvoiceOverplusEnough = True: Exit Function
    Err = 0: On Error GoTo Errhand:
    
    lngʣ������ = 0
    
    strSQL = "" & _
        "   Select Sum(nvl(ʣ������,0)) as ʣ������ " & vbNewLine & _
        "   From Ʊ�����ü�¼" & vbNewLine & _
        "   Where Ʊ�� = [1]  " & _
        "               And (nvl(ʹ�����,'LXH')=[4] or nvl(ʹ�����,'LXH')='LXH')  " & _
        "               And ������ = [2] And ʹ�÷�ʽ = 1 and nvl(ʣ������,0)>0" & vbNewLine & _
                    IIf(lng����ID = 0, "", "             and ID=[3]") & _
        "   Union ALL " & _
        "   Select Sum(nvl(ʣ������,0)) as ʣ������  " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B" & vbNewLine & _
          " Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "             And (nvl(A.ʹ�����,'LXH')=[4] or nvl(A.ʹ�����,'LXH')='LXH')  " & _
          "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
                       IIf(lng����ID = 0, "", "             and A.ID=[3]") & _
          "  "
    strSQL = "Select sum(ʣ������) as ʣ������ From (" & strSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, UserInfo.����, lng����ID, strʹ�����)
    lngʣ������ = Val(nvl(rsTemp!ʣ������))
    zlCheckInvoiceOverplusEnough = lngʣ������ > intNum
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "") As Long
'���ܣ���ȡ�������ò���ָ��Ʊ��������÷�Χ�ڵ�����ID
'������bytKind      =   Ʊ��
'      intNum       =   Ҫ��ӡ��Ʊ������
'      lngLastUseID =   �ϴ�ʹ�õ�����ID
'      lngShareUseID=   ���ز���ָ���Ĺ���ID
'      strBill      =   ��ǰƱ�ݺţ����ڼ���������ε�Ʊ�ݷ�Χ
'      strUseType-ʹ�����
'���أ�
'      >0   =   �ɹ������õ�����ID
'      =0   =   ʧ��
'      -1   =   û������(����򲻹�����δ����),δ���ù���
'      -2   =   û������(����򲻹�����δ����),���õĹ���������򲻹�
'      -3   =   ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
'      -4   =   ָ�����ε�Ʊ�ݲ�����
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo errH
    '1.�ϴε����������Ƿ���ò�����
    If lngLastUseID > 0 Then
        strSQL = "" & _
        "   Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
        "   From Ʊ�����ü�¼ " & _
        "   Where Ʊ��=[1] And ʣ������>=[2] And ID=[3]  " & _
        "           And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, intNum, lngLastUseID, IIf(Trim(strUseType) = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then    'Ŀǰ��Ʊ�ݺſ��ܺ��ϴβ�ͬ��������Ҫ��鷶Χ
                If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '����û�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
                
            ElseIf intNum > 1 Then  '����ȷ���������ε���ʱ,��ǰƱ�ݺ��������β�����
                GetInvoiceGroupID = -4: Exit Function
            End If
        End With
    End If
    
    '2.�ϴε��������β����û򲻿���ʱ,ȡ������Ĳ������õ�
    '  �ж��������ʹ�õ�����,�ٵ�����,��������
    strSQL = "" & _
    "   Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����" & vbNewLine & _
    "   From Ʊ�����ü�¼" & vbNewLine & _
    "   Where Ʊ�� = [1] And ʣ������ >= [2] And ������ = [3]  " & _
    "           And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL ) " & _
    "           And ʹ�÷�ʽ = 1" & vbNewLine & _
    "   Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� desc, ��ʼ����"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, intNum, UserInfo.����, IIf(strUseType = "", "LXH", strUseType))
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then GetInvoiceGroupID = !ID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
            blnTmp = False
            strPre = "" & !ǰ׺�ı�
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = !ID: Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.û�����õ�,ʹ�ñ��ز���ָ���Ĺ�������
    If lngShareUseID > 0 Then
        strSQL = "" & _
        "   Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
        "   From Ʊ�����ü�¼  " & _
        "   Where Ʊ��=[1] And ʣ������>=[2] And ID=[3] " & _
        "   And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, intNum, lngShareUseID, IIf(strUseType = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    GetInvoiceGroupID = lngReturn   '����δ�ҵ���ԭ�����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function CheckUsedBill(bytKind As Byte, ByVal lng����ID As Long, _
    Optional ByVal strBill As String, _
     Optional ByVal strUseType As String = "") As Long
    '���ܣ���鵱ǰ����Ա�Ƿ��п���Ʊ������(���û���),�����ؿ��õ�����ID
    '������bytKind=Ʊ��
    '      lng����ID=��һ�μ��ʱΪ�������õĹ�������ID,�Ժ�Ϊ�ϴ�ʹ�õ�����ID
    '      strBill=Ҫ��鷶Χ��Ʊ�ݺ�
    '˵����
    '    1.�ڼ�鷶Χʱ,��������ж�������Ʊ��,��ֻҪ������һ��֮�о�����
    '    2.�ڼ�鷶Χʱ,����Ҳ�ڼ�鷶Χ֮�ڡ�
    '    3.���ж�������ʱ,ȱʡ���ٵ�����,��������,"���ʹ�õ�����"ԭ��
    '���أ�
    '      ������Ʊ������ID>0
    '      0=ʧ��
    '      -1:û������(�����δ����)��Ҳû�й���(δ����)
    '      -2:���õĹ���������
    '      -3:ָ��Ʊ�ݺŲ��ڵ�ǰ���÷�Χ��(������������Ʊ�ݵ����)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    
    '����Ա��ʣ�������Ʊ�ݼ�
    strSQL = _
        "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
        "From Ʊ�����ü�¼" & vbNewLine & _
        "Where Ʊ�� = [1] And ʹ�÷�ʽ = 1 And ʣ������ > 0 And ������ = [2] And (Nvl(ʹ�����,'LXH')=[3] or  ʹ����� is NULL)" & vbNewLine & _
        "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� Desc, ��ʼ����"
    Set rsSelf = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, UserInfo.����, IIf(strUseType = "", "LXH", strUseType))
    If lng����ID = 0 Then
        '�����е�һ�μ��,��û�����ñ��ع���
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function 'Ҳû������Ʊ��
        '������Ʊ��,������ԭ�򷵻�
        lngReturn = rsSelf!ID
    Else
        '�ϴ�ʹ�õ�����ID���һ�μ��Ĺ���ID,���ж�����
        strSQL = "Select ID,ʹ�÷�ʽ,ʣ������,ǰ׺�ı�,��ʼ����,��ֹ���� From Ʊ�����ü�¼ Where Ʊ��=[1]  And (Nvl(ʹ�����,'LXH')=[3] or  ʹ����� is NULL) And ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, lng����ID, IIf(strUseType = "", "LXH", strUseType))
        '����26352 by ���ջ� 2009-11-20
        If rsTmp.EOF Then CheckUsedBill = -2: Exit Function
        
        If rsTmp!ʹ�÷�ʽ = 2 Then '����,Ҫ�ȿ���û������
            If Not rsSelf.EOF Then
                '�����õģ�����
                lngReturn = rsSelf!ID
            Else
                'û������ȡ����
                If rsTmp!ʣ������ = 0 Then CheckUsedBill = -2: Exit Function '�����Ѿ�����
                lngReturn = rsTmp!ID
                blnTmp = True
            End If
        Else
            '����Ʊ��
            If rsTmp!ʣ������ > 0 Then
                '��ʣ��
                lngReturn = rsTmp!ID
            Else
                '������ʣ�������
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '��������Ҳû��ʣ��
                lngReturn = rsSelf!ID
            End If
        End If
    End If
    
    '���Ʊ�ŷ�Χ�Ƿ���ȷ
    If strBill <> "" Then
        If blnTmp Then
            '�ڹ��÷�Χ�ڷ�Χ�ж�
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!��ʼ����) And UCase(strBill) <= UCase(rsTmp!��ֹ����) And Len(strBill) = Len(rsTmp!��ʼ����)) Then
                lngReturn = -3
            End If
        Else
            '�ڿ������÷�Χ���ж�
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '����������,�������������м��
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    CheckUsedBill = 0
End Function


Public Function GetNextBill(lng����ID As Long) As String
'���ܣ�������������ID,��ȡ��һ��ʵ��Ʊ�ݺ�
'˵����1.��ȡ������Χ�ڵ���ЧƱ��ʱ,���ؿ����û�����
'      2.�ſ��ѱ���ĺ���
    Dim rsMain As ADODB.Recordset
    Dim rsDelete As ADODB.Recordset
    Dim strSQL As String, strBill As String
    
    On Error GoTo errH
    
    strSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ����" & _
        " From Ʊ�����ü�¼ Where ʣ������>0 And ID=[1]"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, "ȡһ��Ʊ�ݺ�", lng����ID)
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!��ǰ����) Then
        strBill = UCase(rsMain!��ʼ����)
    Else
        strBill = UCase(zlCommFun.IncStr(rsMain!��ǰ����))
    End If
    
     '�����:25448
     '���˺�:ȡ����;����=1 And ԭ��=5 And ���:ԭ���ǿ��ܴ����Ѿ�ʹ���˵�Ʊ��,ʹ���˵�,���ų�
     'Ʊ��: 1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
     '����:1-����(ԭ����1��3��5��������)��2-�ջ�(ԭ����2��4��������)
     'ԭ��:1-��������Ʊ�ݣ�2-�����ջط�Ʊ��3-�ش򷢳�Ʊ�ݣ�4-�ش��ջ�Ʊ�ݣ�5-��������Ʊ��
     
    strSQL = "Select Upper(����) as ���� From Ʊ��ʹ����ϸ" & _
        " Where ����||''>=[1] And ����ID=[2]" & _
        " Order by ����"
        
    Set rsDelete = zlDatabase.OpenSQLRecord(strSQL, "ȡһ��Ʊ�ݺ�", strBill, lng����ID)
    Do While True
        '��鷶Χ
        If Left(strBill, Len("" & rsMain!ǰ׺�ı�)) <> UCase("" & rsMain!ǰ׺�ı�) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!��ʼ����) And strBill <= UCase(rsMain!��ֹ����)) Then
            Exit Function
        End If
                
        '�ſ������
        rsDelete.Filter = "����='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = zlCommFun.IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub UpdateShareID(ByVal lngModule As Long, ByVal strShareIDs As String, _
    Optional bytKind As Byte = 5, Optional strParName As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¹������εĴſ�ID
    '���:strShareIDs:�������������
    '        bytKind=1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '        strParName-������(��ʱ,�Գ��õ�����Ϊ׼)
    '����:���˺�
    '����:2011-07-26 17:09:17
    'Ŀǰ�ݶԾ������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strShare As String, varData As Variant, varTemp As Variant, strSQL As String
    Dim i As Long, strIDs As String, lngID As Long, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '��ʽ:����ID1,Ԥ�����ID1|����IDn,Ԥ�����IDn|...
    varData = Split(strShareIDs, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        If Val(varTemp(0)) <> 0 Then
            strIDs = strIDs & "," & Val(varTemp(0))
        End If
    Next
    If strShare <> "" Then
        strShare = Mid(strShare, 2)
            strSQL = "" & _
            "   Select  /*+ rule */ ID From Ʊ�����ü�¼ A,Table(f_num2list([1])) J  " & _
            "   Where A.ID=J.Column_value  And A.Ʊ��=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������ID", lngID, bytKind)
        strShare = ""
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",", ",")
            lngID = Val(varTemp(0))
            If lngID <> 0 Then
                rsTemp.Filter = "ID=" & lngID
                If rsTemp.RecordCount <> 0 Then
                     strShare = strShare & "|" & lngID & "," & varTemp(1)
                End If
                rsTemp.Filter = 0
            End If
            If Val(varTemp(0)) <> 0 Then
                strIDs = strIDs & "," & Val(varTemp(0))
            End If
        Next
    End If
    If strShare <> "" Then strShare = Mid(strShare, 2)
    Select Case bytKind
    Case 1  '�շ��վ�
    Case 2  ' Ԥ���վ�
    Case 3   ' �����վ�
    Case 4   ' �Һ��վ�
    Case 5   '���￨
        If strParName <> "" Then
            zlDatabase.SetPara strParName, strShare, glngSys, lngModule
        Else
            zlDatabase.SetPara "����ҽ�ƿ�����", strShare, glngSys, lngModule
        End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Public Function ExistBill(lngID As Long, bytKind As Byte) As Boolean
'���ܣ��ж��Ƿ����ָ����Ʊ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select ID From Ʊ�����ü�¼ Where ID=[1] And Ʊ��=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�������ID", lngID, bytKind)
    ExistBill = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function




Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intTYPE As Integer
    Dim dtCurdate As Date, strMaxNO As String
    Dim strYearStr As String
    
    Err = 0: On Error GoTo errH:
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = zlStr.PrefixNO & strNO
        Exit Function
    End If
'    ElseIf intNum = 0 Then
'        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
'        Exit Function
'    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ����,������ From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, intNum)
    dtCurdate = Date
    If Not rsTmp.EOF Then
        intTYPE = Val("" & rsTmp!��Ź���)
        dtCurdate = rsTmp!����
        strMaxNO = nvl(rsTmp!������)
    End If
    strYearStr = zlStr.PrefixNO
    If strMaxNO = "" Then strMaxNO = strYearStr & "000001"
    If intTYPE = 1 Then
        '���ձ��
        strSQL = Format(CDate(Format(dtCurdate, "YYYY-MM-dd")) - CDate(Format(dtCurdate, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = zlStr.PrefixNO & strSQL & Format(Right(strNO, 4), "0000")
        Exit Function
    End If
    '������
    If Len(strNO) = 6 Then
        GetFullNO = Left(strMaxNO, 2) & strNO: Exit Function
    End If
    GetFullNO = Left(strMaxNO, 2) & zlLeftPad(Right(strNO, 6), 6, "0")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillOperCheck(bytNO As Byte, strOperator As String, Datadd As Date, Optional strMessage As String = "����", _
    Optional ByVal strNO As String, Optional ByVal lngPatientID As Long, _
    Optional ByVal bytFlag As Byte = 2, Optional ByVal blnOnlyCheckLimit As Boolean, Optional ByVal blnCheckOperator As Boolean = True, _
    Optional ByVal blnCheckCur As Boolean = True) As Boolean
'���ܣ��жϵ�ǰ��Ա�Ե����Ƿ��в���Ȩ��
'������
'   bytNO��1-�Һŵ���,2-�շѵ�,3-���۵�,4-�������,5-סԺ����,6-Ԥ����,7-���ʵ���,8-���￨
'   strOperator������ʵ�ʵĲ���Ա
'   DatAdd�����ݵĵǼ�ʱ��
'   strNO   �����������ʱ����ȷ������
'   lngPatientID�����������ʱ�����ڼ��ʱ�����ȷ�������еĲ���
'   bytFlag��1-�շѵ�,2-���ʵ�,3-���ʵ�
'   blnOnlyCheckLimit��ֻ���������
'   blnCheckOperator��Ҫ����Ƿ�����������˵���
'   blnCheckCur���Ƿ���������
'���أ��Ƿ��в���Ȩ��
'˵�����������ʾ�ڱ������С�

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
    
    strSQL = "Select Nvl(ʱ������,0) as ʱ������,Nvl(���˵���,0) as ���˵���,Nvl(�������,0) as ������� From ���ݲ������� Where ��ԱID=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, UserInfo.ID, bytNO)
    If rsTmp.EOF Then
        BillOperCheck = True
        Exit Function
    Else
        If Not blnOnlyCheckLimit Then
            If rsTmp!���˵��� = 0 And blnCheckOperator Then
                If strOperator <> UserInfo.���� Then
                    MsgBox "��û��Ȩ�޶�" & strOperator & "�����" & strBill & "����" & strMessage & "��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            If rsTmp!ʱ������ > 0 Then
                If Int(zlDatabase.Currentdate) - Int(CDate(Datadd)) + 1 > rsTmp!ʱ������ Then
                    MsgBox "��ֻ�ܶ� " & rsTmp!ʱ������ & " ���ڴ����" & strBill & "����" & strMessage & "��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If rsTmp!������� > 0 And blnCheckCur Then
            If strNO <> "" Then
                curTmp = GetBillMoney(int��Դ, strNO, lngPatientID, bytFlag)
                If curTmp >= rsTmp!������� Then
                    MsgBox "��ֻ�ܶ� " & rsTmp!������� & " Ԫ���µ�" & strBill & "����" & strMessage & "��" & _
                    vbCrLf & "����[" & strNO & "]��ʵ�ս��ϼ�Ϊ:" & curTmp, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        BillOperCheck = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillMoney(ByVal int��Դ As Integer, strNO As String, Optional lng����ID As Long, Optional ByVal bytFlag As Byte = 2) As Currency
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�ŵ��ݵ�ʵ�ս��ϼ�,��һ�ż��ʱ���ָ�����˵�ʵ�ս��ϼ�
    '��Σ�int��Դ-1-����,2-סԺ
    '      bytFlag-1-�շѵ�,2-���ʵ�,3-���ʵ�(�Զ����ʵ�)
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-02 14:26:50
    '˵����int��Դ�����˲���
    '------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
    
    If lng����ID = 0 Then
        strSQL = "Select Sum(ʵ�ս��) as ��� From  " & IIf(int��Դ = 1, "������ü�¼", " סԺ���ü�¼") & " Where NO=[1] And ��¼����=[2] And ��¼״̬ IN(0,1)"
    Else
        strSQL = "Select Sum(ʵ�ս��) as ��� From " & IIf(int��Դ = 1, "������ü�¼", " סԺ���ü�¼") & " Where NO=[1] And ��¼����=[2] And ��¼״̬ IN(0,1) And ����ID=[3]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag, lng����ID)
    
    If Not rsTmp.EOF Then GetBillMoney = Val("" & rsTmp!���)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadBillInfo(ByVal int��Դ As Integer, ByVal strNO As String, _
    ByVal intFlag As Integer, ByRef strOperator As String, ByRef Datadd As Date, _
    Optional ByRef lng����ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�ŵ��ݵĲ���Ա�͵Ǽ�ʱ��
    '��Σ�int��Դ-1-����,2-סԺ
    '      intFlag:-1=����,-2=Ԥ��,-3=�����㣬����ͬ סԺ���ü�¼��������ü�¼.��¼����(1-�շѼ�¼��2(12)-���ʼ�¼��3(13)-�Զ����ʼ�¼;4-�Һż�¼��5(15)-���￨��¼)
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-02 16:03:22
    '˵��������������Ϻ���BillOperCheckʹ�á�
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
        
    On Error GoTo errH
    
    If intFlag = -1 Then
        strSQL = "Select ����Ա����,�շ�ʱ�� as �Ǽ�ʱ��, ����ID From ���˽��ʼ�¼ Where NO=[1] And ��¼״̬ IN(1,3)"
    ElseIf intFlag = -2 Then
        strSQL = "Select ����Ա����,�տ�ʱ�� as �Ǽ�ʱ��,����ID  From ����Ԥ����¼ Where NO=[1] And ��¼״̬ IN(1,3)"
    ElseIf intFlag = -3 Then
        strSQL = "Select ����Ա����, �Ǽ�ʱ��, ����id From ���ò����¼ Where NO = [1] And ��¼״̬ In (1, 3) And Rownum < 2"
    Else
        strSQL = "Select Nvl(����Ա����,������) as ����Ա����,�Ǽ�ʱ��,����ID  From " & IIf(int��Դ = 1, "������ü�¼", " סԺ���ü�¼") & " Where NO=[1] And ��¼����=[2] And ��¼״̬ IN(0,1,3) And RowNum=1"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, intFlag)
    If Not rsTmp.EOF Then
        strOperator = rsTmp!����Ա����
        Datadd = rsTmp!�Ǽ�ʱ��
        lng����ID = Val(nvl(rsTmp!����ID))
        ReadBillInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(intTYPE As Integer, lng����ID As Long) As Currency
'����:��ȡָ�����˵Ļ��۵����ϼ�
'���:IntType:0-����,1-סԺ,2-����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnAllFee As Boolean, strWhere As String
        
    On Error GoTo errH
    
    '���ʱ�����������סԺ���۷���
    If intTYPE = 1 Then
        blnAllFee = Val(zlDatabase.GetPara("���ʱ�����������סԺ���۷���", glngSys, 1150)) = 1
        If blnAllFee Then
            strWhere = ""
        Else
            strWhere = " And Nvl(��ҳID,0) = (Select Nvl(��ҳID,0) From ������Ϣ Where ����ID = [1])"
        End If
    Else
        strWhere = ""
    End If
    
    If intTYPE = 1 Then
        strSQL = "" & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ�  " & _
        "   From סԺ���ü�¼ " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1] and �����־=2" & strWhere
    Else
        If intTYPE = 2 Then
            strSQL = "Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ� From ������ü�¼ Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1]"
            strSQL = strSQL & " union ALL  Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ� From סԺ���ü�¼ Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1]"
            strSQL = "Select Sum(nvl(���۷��úϼ�,0)) as ���۷��úϼ� From ( " & strSQL & ")"
        Else
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
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ�����˵Ļ����ܶ�", lng����ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!���۷��úϼ�
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    If Not rsTmp.EOF Then
        GetPatiDayMoney = Val("" & rsTmp!���)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function HavedInCost(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '����÷���False,���򷵻�true
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT SUM(ʵ�ս��) ʵ�ս�� FROM סԺ���ü�¼ where ����ID=[1] AND ��ҳID=[2] and ��¼״̬<>0 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ��з���", lng����ID, lng��ҳID)
    If Not rsTemp Is Nothing Then
        If Not rsTemp.EOF Then
            If nvl(rsTemp!ʵ�ս��, 0) <> 0 Then HavedInCost = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HavedDirections(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'����:��鲡�˱���סԺ�Ƿ��Ѿ�����ҽ��
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT 1 FROM ����ҽ����¼ Where ����ID = [1] And ��ҳid = [2] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ǵ���ҽ��", lng����ID, lng��ҳID)
    If Not rsTemp Is Nothing Then
        HavedDirections = rsTemp.EOF = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMoneyInfo(lng����ID As Long, Optional dblModiMoney As Double, _
    Optional blnInsure As Boolean, _
    Optional int���� As Integer = -1, _
    Optional bln������ͳ�� As Boolean = False, _
    Optional bytModiMoneyType As Byte = 0, _
    Optional ByVal blnFamilyMoney As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵�ʣ���
    '���:blnInsure=�Ƿ��ſ�ҽ�����˵�Ԥ�����
    '       curModiMoney=�޸�ʱ,ԭ���ݵĵ�ǰ���˵ķ��úϼ�
    '       int����:����(0-�����סԺ����;1-����;2-סԺ),-1��ʾ����
    '       bytModiMoneyType-�޸ķ��õ����(�ڰ����ͳ��ʱ��Ч)
    '       blnFamilyMoney-�Ƿ��ȡ�������
    '����:
    '����:����ʣ���
    '����:���˺�
    '����:2011-07-21 15:33:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, blnҽ�� As Boolean, lng��ҳID As Long
    Dim strSQL As String
    On Error GoTo errH
    If blnInsure Then
        strSQL = "Select A.����,A.��ҳID From ������ҳ A,������Ϣ B" & _
                " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
                " And B.����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
        If Not rsTmp.EOF Then
            blnҽ�� = Not IsNull(rsTmp!����)
            lng��ҳID = rsTmp!��ҳID
        End If
    End If
    strSQL = "Select " & IIf(bln������ͳ��, "����,", "") & IIf(blnFamilyMoney, "0 As ����,", "") & _
            "       Nvl(�������,0) As �������,Nvl(Ԥ�����,0) As Ԥ�����" & _
            " From �������" & _
            " Where ����=1 And ����ID=[1] " & IIf(int���� = -1, "", " And ����=[4]")
    '79868,��ȡ���˼������
    If blnFamilyMoney Then
        strSQL = strSQL & " Union All " & _
                " Select " & IIf(bln������ͳ��, "a.����,", "") & IIf(blnFamilyMoney, "1 As ����,", "") & _
                "       Nvl(a.�������, 0) As �������, Nvl(a.Ԥ�����, 0) As Ԥ�����" & _
                " From ������� A, ���˼��� B" & _
                " Where a.����id = b.����id And b.����id = [1] And a.���� = 1 " & _
                "       And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) " & _
                IIf(int���� = -1, "", " And ����=[4]")
    End If
  
    If dblModiMoney <> 0 Then   '����Ҫ��Union��ʽ,���ֱ��ȥ��,�ڲ�������޼�¼ʱ,���᷵�ؼ�¼
        strSQL = strSQL & " Union All " & _
                " Select " & IIf(bln������ͳ��, "[4] as ����,", "") & IIf(blnFamilyMoney, "0 As ����,", "") & _
                "       -1*[3] as �������,0 as Ԥ����� From Dual"
    End If
    
    '���Ϊҽ��סԺ���ˣ����ڷ���������ſ�Ԥ���еķ���(���ڱ���)
    If blnInsure And blnҽ�� Then
        strSQL = strSQL & " Union All " & _
        " Select  " & IIf(bln������ͳ��, "Decode(��ҳID,NULL,1,0,1,2) as ����,", "") & IIf(blnFamilyMoney, "0 As ����,", "") & _
        "       -1*Nvl(���,0) as �������,0 as Ԥ�����" & _
        " From ����ģ�����" & _
        " Where ����ID=[1] And ��ҳID=[2] "
    End If
    strSQL = "Select " & IIf(bln������ͳ��, "����,", "") & IIf(blnFamilyMoney, "����,", "") & _
            "       nvl(Sum(�������),0) as �������,nvl(Sum(Ԥ�����),0) as Ԥ����� " & _
            " From (" & strSQL & ")" & vbCrLf & _
            IIf(bln������ͳ�� And blnFamilyMoney, " Group by ����,����", _
                IIf(bln������ͳ��, " Group by ����", IIf(blnFamilyMoney, " Group by ����", "")))
    
    Set GetMoneyInfo = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, lng��ҳID, dblModiMoney, int����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function isYBPati(lng����ID As Long, Optional blnIn As Boolean, Optional int���� As Integer) As Boolean
'���ܣ��ж�һ��סԺ�����Ƿ�ҽ������
'������blnIN=�Ƿ������Ժ
'      int����=��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select A.���� From ������ҳ A,������Ϣ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
        " And B.����ID=[1] " & IIf(blnIn, " And A.��Ժ���� is NULL", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    If Not rsTmp.EOF Then
        isYBPati = Not IsNull(rsTmp!����)
        int���� = nvl(rsTmp!����, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStockSet(ByVal lngҩ��ID As Long, ByVal lngҩƷID As Long) As Recordset
    Dim strSQL As String
    
    If Val(zlDatabase.GetPara(150, glngSys)) = 0 Then '����ҩƷ���ⷽʽ��0-�������Ƚ��ȳ���1-��Ч������ȳ�,Ч����ͬ�����ٰ������Ƚ��ȳ�
        strSQL = "Nvl(����,0)"
    Else
        strSQL = "Ч��,Nvl(����,0)" 'Ч��Ϊ�����������
    End If
    
    'ҩ��������ҩƷ����Ч��(����Ŀⷿһ����ҩ��)
    strSQL = "Select Nvl(����,0) as ����,Nvl(��������,0) as ���," & _
        " Nvl(���ۼ�,Nvl(Decode(Nvl(ʵ������,0),0,0,ʵ�ʽ��/ʵ������),0)) as ʱ��," & _
        " Nvl(ʵ�ʲ��,0) as ʵ�ʲ��,Nvl(ʵ�ʽ��,0) as ʵ�ʽ��" & _
        " From ҩƷ���" & _
        " Where �ⷿID=[1] And ҩƷID=[2] And Nvl(��������,0)>0" & _
        " And ����=1 And (Nvl(����,0)=0 Or Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
        " Order by " & strSQL
        
    On Error GoTo errH
    Set GetStockSet = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngҩ��ID, lngҩƷID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getʱ��ҩƷӦ�ս��(ByVal lngҩ��ID As Long, ByVal lngҩƷID As Long, ByRef dblAllTime As Double, ByVal strDec As String, ByRef dblPriceSingle As Double) As Currency
'���ܣ���ȡ����ʱ��ҩƷ��Ӧ�ս����ݲ�ͬ�ĳ��ⷽʽ�����κϼƣ�
'������
'      strDec-���ý���λ��
'      dblAllTime-����Ϊ����������(�ۼ�����)���������Ϊ0���ʾ����㹻�������ʾ��治��
'      dblPriceSingle-ֻ��һ������ʱ���ظ����εĵ��ۣ�������������ȳ��ٳ�����������������ͬ�������ĵ��۲�ͬ

    Dim rsPrice As ADODB.Recordset
    Dim dblPrice As Double, dblCurTime As Double, i As Long
    
    Set rsPrice = GetStockSet(lngҩ��ID, lngҩƷID)
    'ʱ��=�ܽ��/������
    dblPrice = 0 '������Ӧ�ս��
    For i = 1 To rsPrice.RecordCount
        If dblAllTime = 0 Then Exit For
        'ȡС��
        If dblAllTime <= rsPrice!��� Then
            dblCurTime = dblAllTime
        Else
            dblCurTime = rsPrice!���
        End If
        If i = 1 Then
            dblPriceSingle = Format(rsPrice!ʱ��, gstrFeePrecisionFmt)
        Else
            dblPriceSingle = 0
        End If
        dblPrice = dblPrice + Format(dblCurTime * Format(rsPrice!ʱ��, gstrFeePrecisionFmt), strDec)
        dblAllTime = dblAllTime - dblCurTime
        rsPrice.MoveNext
    Next
    
    Getʱ��ҩƷӦ�ս�� = dblPrice
End Function

Public Function GetAuditRecord(lng����ID As Long, lng��ҳID As Long, Optional lng��Ŀid As Long) As ADODB.Recordset
'���ܣ���ȡָ�����˵ķ���������Ŀ,��δ����ʹ����������������Ϊ��ʱ,��������Ϊ��
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ��ĿId,ʹ������,��������,ʹ������-Nvl(��������,0) �������� From ����������Ŀ " & _
            "Where ����ID=[1] And ��ҳID=[2]" & IIf(lng��Ŀid <> 0, " And ��ĿID=[3]", "")
    Set GetAuditRecord = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, lng��ҳID, lng��Ŀid)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistInsure(ByVal strNO As String, Optional ByVal strTime As String, _
    Optional ByVal blnAuditing As Boolean, Optional ByVal bytFlag As Byte = 2) As Integer
'���ܣ��ж�ָ����סԺ���ʵ����Ƿ��ҽ�����˼ǵ���
'������strNO=���ʵ��ݺ�
'      blnAuditing=�Ƿ����ڼ������,ֻ���δ��˵Ĳ�������
'      bytFlag=2-�˹����ʵ�,3-�Զ����ʵ�
'���أ�������򷵻ز�������
'˵����1.ֻ��סԺҽ������,�������ﲡ�˵�ҽ������
'      2.���ʱ�ֻ���ص�һ�����˵�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.���� From סԺ���ü�¼ A,������ҳ B" & _
        " Where A.��¼����=[2] And A.��¼״̬" & IIf(blnAuditing, "=0", " IN(0,1,3)") & " And B.���� is Not NULL" & _
        " And A.NO=[1] And A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
        IIf(strTime <> "", " And A.�Ǽ�ʱ��=[3]", "")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag)
    End If

    If Not rsTmp.EOF Then BillExistInsure = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub AdjustCpt(lngID As Long)
    '���ܣ�ҩƷ����
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "zl_ҩƷ�շ���¼_Adjust(" & lngID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, App.ProductName)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng�շ�ϸĿID, int����)
    If rsTmp.RecordCount > 0 Then Getҽ������ = rsTmp!����
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��������(lngҩƷID As Long) As Double
'���ܣ���ȡָ��ҩƷ�Ĵ�������,�����۵�λ���ء�
'������lngҩƷID=ҩƷID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(A.��������,0) as �������� From ҩƷ���� A,ҩƷ��� B Where A.ҩ��ID=B.ҩ��ID And B.ҩƷID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngҩƷID)
    If Not rsTmp.EOF Then Get�������� = rsTmp!��������
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ְ��(lngҩƷID As Long) As String
'���ܣ�����ҩƷID��ȡ�䴦��ְ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    Get����ְ�� = "00"
    strSQL = "Select Nvl(B.����ְ��,'00') as ����ְ�� From ҩƷ��� A,ҩƷ���� B Where A.ҩ��ID=B.ҩ��ID And A.ҩƷID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngҩƷID)
    If Not rsTmp.EOF Then Get����ְ�� = rsTmp!����ְ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDiseaseCode(ByRef frmParent As Object, ByRef blnCancel As Boolean, _
    ByVal strInput As String, ByVal strSex As String, ByVal strKind As String, _
    ByVal X As Long, ByVal Y As Long, ByVal txtHeight As Long, Optional ByVal bytSize As Byte) As ADODB.Recordset
'����:����������ַ����ض�Ӧ�ļ��������¼��
'����:strCode-����ֵ,strSex-�Ա�����,strKind-�����������
'     x,y������ѡ��������Ļ����ʾ������λ��,txtHeight-�����ĸ߶�,blnCnacel�Ƿ�ȡ��ѡ��
'     ��"bytSize=?"��ʾ���������С(0-С����,1-������;С����Ϊ9����,������Ϊ12����),Ĭ��С���塣
    Dim strSQL As String, strCode As String
    Dim strLike As String, strWhere As String, lngCodeKind As Long
    
    If Trim(strInput) = "" Then Exit Function
    
    strLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    strCode = strLike & UCase(Trim(strInput)) & "%"
    '����ƥ�䷽ʽ��0-ƴ��,1-���,2-����
    lngCodeKind = Val(zlDatabase.GetPara("���뷽ʽ"))
    
    
    If zlCommFun.IsCharAlpha(strInput) Then
        If lngCodeKind = 0 Then
            strWhere = "(A.���� Like [1] Or A.���� Like [1])"
        ElseIf lngCodeKind = 1 Then
            strWhere = "(A.���� Like [1] Or A.����� Like [1])"
        Else
            strWhere = "(A.���� Like [1] Or A.���� Like [1] Or A.����� Like [1])"
        End If
    ElseIf IsNumeric(strInput) Or zlCommFun.IsNumOrChar(strInput) Then
        strWhere = "A.���� Like [1]"
    ElseIf zlCommFun.IsCharChinese(strInput) Then
        strWhere = "A.���� Like [1]"
    Else
        If lngCodeKind = 0 Then
            strWhere = "(A.���� Like [1] Or A.���� Like [1] Or A.���� Like [1])"
        ElseIf lngCodeKind = 1 Then
            strWhere = "(A.���� Like [1] Or A.���� Like [1] Or A.����� Like [1])"
        Else
            strWhere = "(A.���� Like [1] Or A.���� Like [1] Or A.���� Like [1] Or A.����� Like [1])"
        End If
    End If
    If strSex <> "" Then strWhere = strWhere & " And (A.�Ա�����='" & strSex & "' Or A.�Ա����� is NULL)"
       
       
'    If strKind <> "" Then
'        strSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.�����,A.˵��,A.�Ա�����,B.���" & _
'            " From ��������Ŀ¼ A,����������� B" & _
'            " Where A.���=B.���� And A.���=[2] And Rownum<=100 And " & strWhere & _
'            " Order by A.���,A.����"
'    Else
    '90044ȡ�����Ʒ�������
        strSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.�����,A.˵��,A.�Ա�����" & _
            " From ��������Ŀ¼ A" & _
            " Where A.���=[2] And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) And " & strWhere & _
            " Order by A.����"
            
'    End If
    
    Set GetDiseaseCode = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "������������", 1, "", "��ѡ��", False, False, True, X, Y, txtHeight, blnCancel, False, True, strCode, strKind, "bytSize=" & bytSize)
End Function

Public Function GetDiseaseCodeNew(ByRef frmParent As Object, ByRef blnCancel As Boolean, _
    ByVal strInput As String, ByVal strSex As String, ByVal strKind As String, _
    ByVal X As Long, ByVal Y As Long, ByVal txtHeight As Long, Optional ByVal bytSize As Byte) As ADODB.Recordset
'����:����������ַ����ض�Ӧ�ļ��������¼��,��෵��100����¼
'����:strCode-����ֵ,strSex-�Ա�����,strKind-�����������
'     x,y������ѡ��������Ļ����ʾ������λ��,txtHeight-�����ĸ߶�,blnCnacel�Ƿ�ȡ��ѡ��
'     ��"bytSize=?"��ʾ���������С(0-С����,1-������;С����Ϊ9����,������Ϊ12����),Ĭ��С���塣
    Dim strSQL As String, strCode As String, strRight As String
    Dim strLike As String, strWhere As String, lngCodeKind As Long
    
    If Trim(strInput) = "" Then Exit Function
    
    strLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    strCode = strLike & UCase(Trim(strInput)) & "%"
    strRight = UCase(Trim(strInput)) & "%"
    '����ƥ�䷽ʽ��0-ƴ��,1-���,2-����
    lngCodeKind = Val(zlDatabase.GetPara("���뷽ʽ"))

    If zlCommFun.IsCharChinese(strInput) Then
        strSQL = "���� Like [2] or '('||����||')'||���� Like [2]" '���뺺��ʱֻƥ������
    Else
        strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(lngCodeKind = 0, "����", "�����") & " Like [2]"
    End If
    
'    If strSex <> "" Then strWhere = strWhere & " And (A.�Ա�����='" & strSex & "' Or A.�Ա����� is NULL)"
       
    strSQL = _
                " Select ID,ID as ��ĿID,����,����,����," & IIf(lngCodeKind = 0, "����", "����� as ����") & ",˵��" & _
                " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                IIf(strSex <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order by ����"

'    strSql = "Select A.ID,A.����,A.����,A.����,A.����,A.�����,A.˵��,A.�Ա�����" & _
'             " From  ��������Ŀ¼ A" & _
'             " Where Rownum<=100 And A.���=[3] And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) And " & strWhere & _
'             " Order by A.����"

    
    Set GetDiseaseCodeNew = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "������������", 1, "", "��ѡ��", False, False, True, X, Y, txtHeight, blnCancel, False, True, strRight, strCode, strKind, strSex, "bytSize=" & bytSize)
End Function

Public Function HaveOut(lng����ID As Long) As Boolean
'���ܣ��жϲ��˵�ǰ�Ƿ��Ѿ���Ժ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��Ժ���� From ������Ϣ A,������ҳ B Where A.����ID=B.����ID And A.��ҳID=B.��ҳID and A.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    If rsTmp.EOF Then HaveOut = True: Exit Function 'δ��Ժ���˵�����Ժ����
    If Not IsNull(rsTmp!��Ժ����) Then
        If rsTmp!��Ժ���� <= zlDatabase.Currentdate Then HaveOut = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveExecute(ByVal int��Դ As Integer, ByVal strNO As String, _
    ByVal int��¼���� As Integer, Optional blnAll As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��жϷ��õ����Ƿ������ȫִ�л򲿷�ִ�е�����
    '��Σ�int��Դ-1-����;2-סԺ
    '      strNO=���õ��ݺ�,
    '      int��¼����=��¼����(1-�շ�,2-����)
    '      blnALL=�б𵥾����Ƿ�ȫ��Ϊ��ȫִ�л򲿷�ִ�е�����
    '���أ�����ִ�еģ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-03-02 16:23:05
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String
    On Error GoTo errH
    If int��¼���� = 1 Then
        strWhere = " And mod(��¼����,10)=[2]"
    Else
        strWhere = " And ��¼����=[2]"
    End If
    strWhere = strWhere & " And " & IIf(blnAll, " Not", "") & " ִ��״̬ IN(1,2)"
    
    strSQL = "" & _
    " Select Nvl(Count(ID),0) as ��Ŀ" & _
    " From " & IIf(int��Դ = 1, "������ü�¼", "סԺ���ü�¼") & _
    " Where NO=[1] And ��¼״̬ IN(0,1,3)  " & strWhere

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, int��¼����)
    
    If blnAll Then
        HaveExecute = (rsTemp!��Ŀ = 0)
    Else
        HaveExecute = (rsTemp!��Ŀ > 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function HaveBilling(ByVal int��Դ As Integer, ByVal strNO As String, Optional ByVal blnAll As Boolean = True, _
    Optional ByVal strTime As String, Optional ByVal bytFlag As Byte = 2) As Integer
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��ж�һ�ż��ʵ�/���Ƿ��Ѿ�����
    '��Σ�int��Դ-1-����;2-סԺ
    '      strNO=���ʵ��ݺ�,�������ＰסԺ
    '      blnALL=�Ƿ�����ŵ������ݽ����ж�,����ֻ��δ���ʲ��ֽ����ж�
    '���Σ�
    '���أ�0-δ����,1=��ȫ������,2-�Ѳ��ֽ���
    '���ƣ����˺�
    '���ڣ�2010-03-02 16:37:22
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    
    On Error GoTo errH
        
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag)
    End If
    
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

Public Function Check����ʱ��(ByVal varDate As Date, ByVal varPaitOrNO As Variant) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鷢��ʱ���Ƿ�Ϸ�
    '������varDate=����ʱ��
    '      varPaitOrNO=����ID����ʵ���(�����Ƕಡ�˵�)
    '���أ�������ʾ
    '����:���˺�
    '����:2015-07-10 15:47:18
    '˵����1.��鷢��ʱ�䲻�����ڲ��˵���Ժʱ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    If TypeName(varPaitOrNO) = "String" Then
        strSQL = "Select Distinct ����,����ID,��ҳID  From סԺ���ü�¼ Where ��¼����=2 And NO=[1]"
            
        strSQL = "Select A.����,B.��ҳID,B.��Ժ����" & _
            " From (" & strSQL & ") A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID"
    Else
        strSQL = "" & _
        "Select nvl(B.����,A.����) as ����,B.��ҳID,B.��Ժ���� From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[2]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, CStr(varPaitOrNO), Val(varPaitOrNO))
    For i = 1 To rsTmp.RecordCount
        If Format(varDate, "yyyy-MM-dd HH:mm:ss") < Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm:ss") Then
            Check����ʱ�� = "���õķ���ʱ�䲻��С�ڲ���""" & rsTmp!���� & """����Ժʱ��:" & Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm:ss") & "��"
            Exit Function
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnAuditReFee(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ���鲡���Ƿ����δ��׼���˷�����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select 1" & vbNewLine & _
            "From Dual" & vbNewLine & _
            "Where Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From סԺ���ü�¼ A" & vbNewLine & _
            "       Where A.����id = [1] And A.��ҳid = [2] And Exists (Select 1 From ���˷������� B Where B.����id = A.ID And B.״̬ = 0))"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, lng��ҳID)
    GetUnAuditReFee = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get������׼��Ŀ(lng����ID As Long, strField As String, Optional intInsure As Integer) As String
'���ܣ�����ҽ�����˵Ĳ��ֻ�ȡ������׼��Ŀ������,�����������շ�ϸĿ��
'������strField=�����ֶ�,��"C.�շ�ϸĿID"
'˵�����жϲ��ֺ󣬿���ֱ�ӷ���SQL��䣬��Ч�ʲ���
'    IN (
'        Select �շ�ϸĿID From ����֧����Ŀ
'        Where ���� = XXXX
'            And �շ�ϸĿID IN (Select �շ�ϸĿID From ������׼��Ŀ Where Nvl(����,0)=0 And ����=1 And ����ID=XXXX)
'        ) Or 0=(Select Count(*) From ������׼��Ŀ Where Nvl(����,0)=0 And ����=1 And ����ID=XXXX)
'
'    Not IN (
'        Select �շ�ϸĿID From ����֧����Ŀ
'        Where ���� = XXXX
'            And �շ�ϸĿID IN (Select �շ�ϸĿID From ������׼��Ŀ Where Nvl(����,0)=0 And ����=2 And ����ID=XXXX)
'        ) Or 0=(Select Count(*) From ������׼��Ŀ Where Nvl(����,0)=0 And ����=2 And ����ID=XXXX)
'
'    IN (
'        Select �շ�ϸĿID From ����֧����Ŀ
'        Where ���� = XXXX
'            And Nvl(����ID,0) IN (Select �շ�ϸĿID From ������׼��Ŀ Where Nvl(����,0)=1 And ����=1 And ����ID=XXXX)
'        ) Or 0=(Select Count(*) From ������׼��Ŀ Where Nvl(����,0)=1 And ����=1 And ����ID=XXXX)
'
'    Not IN (
'        Select �շ�ϸĿID From ����֧����Ŀ
'        Where ���� = XXXX
'            And Nvl(����ID,0) IN (Select �շ�ϸĿID From ������׼��Ŀ Where Nvl(����,0)=1 And ����=2 And ����ID=XXXX)
'        ) Or 0=(Select Count(*) From ������׼��Ŀ Where Nvl(����,0)=1 And ����=2 And ����ID=XXXX)

    Dim rsTmp As ADODB.Recordset
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
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

Public Function is���տ�(ByVal strNO As String) As Boolean
'����:�ж�һ��Ԥ������Ƿ�����ȡ�Ĵ��տ�
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select NO From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
            "Where A.NO = [1] And A.��¼���� = 1 And A.���㷽ʽ = B.���� And B.���� = 5"
            
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO)
    is���տ� = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckFeeItemAvailable(ByVal lngFeeItemID As Long, ByVal bytFlag As Byte) As Boolean
'����:����շ���Ŀ�Ƿ�δͣ��,���ҷ����ڲ���
'����:bytFlag:�������:1-����,2-סԺ
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From �շ���ĿĿ¼ Where ID = [1] And (����ʱ�� is Null Or ����ʱ�� > Sysdate) And ������� In (" & bytFlag & ",3)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItemID)
    CheckFeeItemAvailable = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckTextLength(strName As String, txtObj As TextBox) As Boolean
'����:��鲢��ʾ�ı������볤���Ƿ���
    CheckTextLength = zlControl.TxtCheckInput(txtObj, strName, , True)
End Function

Public Function ReCalcOld(ByVal DateBir As Date, ByRef cbo���䵥λ As ComboBox, Optional ByVal lng����ID As Long, Optional ByVal blnSetControl As Boolean = True, _
    Optional ByVal datCalc As Date) As String
'����:���ݳ����������¼��㲡�˵�����,�������䵥λ
'����:blnSetControl�Ƿ��������䵥λ�ؼ�
'     datCalc-ָ����������,δָ��ʱ��ϵͳʱ�����
'����:����,���䵥λ
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    If datCalc = CDate(0) Then
        strSQL = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    Else
        strSQL = "Select Zl_Age_Calc([1],[2],[3]) old From Dual"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, DateBir, datCalc)
    If blnSetControl = False Then
        ReCalcOld = Trim(nvl(rsTmp!old))
        Exit Function
    End If
    
    If Not IsNull(rsTmp!old) Then
        If rsTmp!old Like "*��" Or rsTmp!old Like "*��" Or rsTmp!old Like "*��" Then
            strTmp = Mid(rsTmp!old, 1, Len(rsTmp!old) - 1)
            If IsNumeric(strTmp) Then
                Call cbo.Locate(cbo���䵥λ, Mid(rsTmp!old, Len(rsTmp!old), 1))
            Else
                strTmp = rsTmp!old
                cbo���䵥λ.ListIndex = -1
            End If
        Else
            strTmp = rsTmp!old
            If IsNumeric(strTmp) Then
                cbo���䵥λ.ListIndex = 0
            Else
                cbo���䵥λ.ListIndex = -1
            End If
        End If
    End If
    If cbo���䵥λ.ListIndex = -1 Then
        cbo���䵥λ.Visible = False
    Else
        If cbo���䵥λ.Visible = False Then cbo���䵥λ.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReCalcBirth(ByVal strOld As String, ByVal str���䵥λ As String) As String
'����:������������䵥λ���㲡�˵ĳ�������,���䵥λΪ��ʱ,�������ռٶ�Ϊ1��1��,���䵥λΪ��ʱ,�������ڼٶ�Ϊ1��
'����:��������
    Dim strTmp As String, strFormat As String, lngDays As Long
    
    strTmp = "____-__-__"
    If str���䵥λ = "" Then
        strFormat = "YYYY-MM-DD"
        If strOld Like "*��*��" Or strOld Like "*��*����" Then
            strFormat = "YYYY-MM-01"
            lngDays = 365 * Val(strOld) + 30 * Val(Mid(strOld, InStr(1, strOld, "��") + 1))
        ElseIf strOld Like "*��*��" Or strOld Like "*����*��" Then
            lngDays = 30 * Val(strOld) + Val(Mid(strOld, InStr(1, strOld, "��") + 1))
        ElseIf strOld Like "*��" Or IsNumeric(strOld) Then
            strFormat = "YYYY-01-01"
            lngDays = 365 * Val(strOld)
        ElseIf strOld Like "*��" Or strOld Like "*����" Then
            strFormat = "YYYY-MM-01"
            lngDays = 30 * Val(strOld)
        ElseIf strOld Like "*��" Then
            lngDays = Val(strOld)
        End If
        If lngDays <> 0 Then strTmp = Format(DateAdd("d", lngDays * -1, zlDatabase.Currentdate), strFormat)
    ElseIf strOld <> "" Then
        Select Case str���䵥λ
            Case "��"
                If Val(strOld) > 200 Then lngDays = -1
            Case "��"
                If Val(strOld) > 2400 Then lngDays = -1
            Case "��"
                If Val(strOld) > 73000 Then lngDays = -1
        End Select
        
        If lngDays = 0 Then
            strTmp = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
            strTmp = Format(DateAdd(strTmp, Val(strOld) * -1, zlDatabase.Currentdate), "YYYY-MM-DD")
            
            If str���䵥λ = "��" Then
                strTmp = Format(strTmp, "YYYY-01-01")
            ElseIf str���䵥λ = "��" Then
                strTmp = Format(strTmp, "YYYY-MM-01")
            End If
        End If
    End If
    ReCalcBirth = strTmp
End Function

Public Function CheckOldData(ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox) As Boolean
'���ܣ������������ֵ����Ч��
'���أ�
    If Not IsNumeric(txt����.Text) Then CheckOldData = True: Exit Function
    
    Select Case cbo���䵥λ.Text
        Case "��"
            If Val(txt����.Text) > 200 Then
                MsgBox "���䲻�ܴ���200��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txt����.Text) > 2400 Then
                MsgBox "���䲻�ܴ���2400��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txt����.Text) > 73000 Then
                MsgBox "���䲻�ܴ���73000��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function

Public Function GetOldAcademic(ByVal DateBir As Date, ByVal str���䵥λ As String) As Long
'���ܣ����ݵ�ǰ�ĳ������ں����䵥λ�����������ϵ�����ֵ
'���أ�����
    Dim DatCur As Date, lngOld As Long, strInterval As String
    If DateBir = CDate(0) Or InStr(" ������", str���䵥λ) < 2 Then Exit Function
    
    DatCur = zlDatabase.Currentdate
    
    strInterval = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
    lngOld = DateDiff(strInterval, DateBir, DatCur)
    If DateAdd(strInterval, lngOld, DateBir) > DatCur Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function

Public Sub LoadOldData(ByVal strOld As String, ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox)
'����:�����ݿ��б�������䰴�淶�ĸ�ʽ���ص�����,���淶��ԭ����ʾ
    Call zlControl.LoadOldData(strOld, txt����, cbo���䵥λ)
End Sub
Public Function zlGetFeeFields(Optional strTableName As String = "������ü�¼", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ�����ֵ
    '��Σ�strTableName:��:������ü�¼;סԺ���ü�¼;....
    '      blnReadDatabase-�����ݿ��ж�ȡ
    '���Σ�
    '���أ��ֶμ�
    '���ƣ����˺�
    '���ڣ�2010-03-10 10:41:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strFileds As String
    
    Err = 0: On Error GoTo Errhand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "������ü�¼"
        zlGetFeeFields = "" & _
        "Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, " & _
        "����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, " & _
        "�Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, " & _
        "����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, " & _
        "���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
        Exit Function
    Case "סԺ���ü�¼"
        zlGetFeeFields = "" & _
         " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, " & _
         " �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, " & _
         " ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, " & _
         " ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, " & _
         " ����id , ���ʽ��, ���մ���ID, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
         Exit Function
    Case "���˽��ʼ�¼"
        zlGetFeeFields = "Id, No, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������, ��ע"
        Exit Function
    Case "����Ԥ����¼"
        zlGetFeeFields = "" & _
        " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, " & _
        " ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�,Ԥ�����,�����ID,���㿨���,����,������ˮ��,����˵��,������λ,�������,У�Ա�־"
        Exit Function
    Case "��Ա��"
        zlGetFeeFields = "" & _
        "Id, ���, ����, ����, ���֤��, ��������, �Ա�, ����, ��������, �칫�ҵ绰, �����ʼ�, ִҵ���, ִҵ��Χ, " & _
        "����ְ��, רҵ����ְ��, Ƹ�μ���ְ��, ѧ��, ��ѧרҵ, ��ѧʱ��, ��ѧ����, ������ѵ, ���п���, ���˼��, ����ʱ��, " & _
        "����ʱ��, ����ԭ��, ����, վ��"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo Errhand:
    strSQL = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ϣ", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & nvl(!Column_Name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
Errhand:
    zlGetFeeFields = "*"
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetFullFieldsTable(Optional strTableName As String = "������ü�¼", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�����ݱ��е��ֶ�.������Select Id,....
    '��Σ�bytHistory-0-��������ʷ����,1-��������ʷ����,2-����������( select * from tablename Union select * from Htablename)
    '      strWhere-����
    '      blnSubTable-�Ƿ��ӱ�
    '      strAliasName-����
    '���Σ�
    '���أ�select ID ... From tableName Union ALL
    '���ƣ����˺�
    '���ڣ�2010-03-10 11:19:11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '��
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '����ʷ
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '���߶�����
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & " " & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
End Function
Public Function GetServiceDept(str�շ�ϸĿIDs As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    If InStr(1, str�շ�ϸĿIDs, ",") = 0 Then
        strSQL = "" & _
        "   Select  /*+ rule */ Distinct   �շ�ϸĿID,Nvl(��������ID,0) as ��������ID,ִ�п���id " & _
        "   From �շ�ִ�п��� A " & _
        "   Where   A.�շ�ϸĿID  =[2] "
    Else
        strSQL = "" & _
        "   Select  /*+ rule */ Distinct   �շ�ϸĿID,Nvl(��������ID,0) as ��������ID,ִ�п���id " & _
        "   From �շ�ִ�п��� A," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        "   Where   A.�շ�ϸĿID  = j.Column_Value"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡִ�п�����Ϣ", Replace(str�շ�ϸĿIDs, "'", ""), Val(str�շ�ϸĿIDs))
    If Not rsTmp.EOF Then Set GetServiceDept = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlComboxLoadFromSQL(ByVal strSQL As String, cboControl As Variant, Optional ByVal blnID As Boolean = False) As Boolean
'�������Ĺ����Ǵ����ݿ��ж����б�ֵ��װ����������
    Dim rsTemp As New ADODB.Recordset
    Dim intCount As Long
    Dim cmbArray As Variant
    
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡCbo����")
    '����������
    If IsArray(cboControl) Then
        cmbArray = cboControl
    Else
        'ǿ�����һ������
        cmbArray = Array(cboControl)
    End If
    
    For intCount = LBound(cmbArray) To UBound(cmbArray)
        cmbArray(intCount).Clear
        Do Until rsTemp.EOF
            If IsNull(rsTemp("����")) Then
                cmbArray(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("����")
            Else
                cmbArray(intCount).AddItem rsTemp("����") & "." & rsTemp("����")
            End If
            If blnID = True Then cmbArray(intCount).ItemData(cmbArray(intCount).NewIndex) = rsTemp("ID")
            If rsTemp("ȱʡ��־") = 1 Then
                cmbArray(intCount).ListIndex = cmbArray(intCount).NewIndex
                cmbArray(intCount).ItemData(cmbArray(intCount).NewIndex) = 1
            End If
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        If blnID = True Then cmbArray(intCount).ListIndex = 0
    Next
    
    zlComboxLoadFromSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlComboxLoadFromSQL = False
End Function

Public Function zlAddComboItem(cboControl As Control, strItem As String, Optional ByVal cboType As Integer = 1, Optional ByVal cboItemData As Long) As Boolean
    '����cboType  = 1ʱ��ʾ�����������ִ�ͷ��
    '             = 2ʱ��ʾȫ������
    Dim varTemp As Variant
    Dim strTemp As String
    
    '�������б����
    If IsNull(strItem) Or Trim(strItem) = "" Then Exit Function
    For varTemp = 0 To cboControl.ListCount - 1
        If cboType = 1 Then
            strTemp = Mid(cboControl.List(varTemp), InStr(cboControl.List(varTemp), ".") + 1)
            If strItem = strTemp Then
                cboControl.ListIndex = varTemp
                Exit Function
            End If
        ElseIf cboType = 2 Then
            If strItem = cboControl.List(varTemp) Then
                cboControl.ListIndex = varTemp
                Exit Function
            End If
        Else
            If cboItemData = cboControl.ItemData(varTemp) Then
                cboControl.ListIndex = varTemp
                Exit Function
            End If
        End If
    Next
    
    If cboType = 1 Then
        cboControl.AddItem strItem
        cboControl.ListIndex = cboControl.NewIndex
    ElseIf cboType = 2 Then
        cboControl.AddItem strItem
        cboControl.ListIndex = cboControl.NewIndex
    End If
End Function
Public Function zlCboFindItem(ByVal cboObj As Object, ByVal lngFindID As Long, _
    Optional strItem As String = "", Optional blnOnlyFind As Boolean = True, Optional blnFindLocal As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Combox��ItemData���ݽ��ж�λ
    '��Σ�cboObj-Combox����
    '         lngFindID-��Ҫ���ҵ�ID
    '         strItem-��Ҫ���ҵĻ����ӵ�����(��blnOnlyFind=false)ʱ
    '         blnOnlyFind-�Ƿ����.
    '        blnFindLocal-�ҵ���,��λ��
    '���Σ�
    '���أ��ҵ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-04-06 17:28:17
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngLocate As Long
    zlCboFindItem = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If cboObj.ItemData(lngLocate) = lngFindID Then
            If blnFindLocal Then cboObj.ListIndex = lngLocate
            zlCboFindItem = True
            Exit Function
        End If
    Next
    If blnOnlyFind Then Exit Function
    cboObj.AddItem strItem
    cboObj.ItemData(cboObj.NewIndex) = lngFindID
    If blnFindLocal Then cboObj.ListIndex = cboObj.NewIndex
    zlCboFindItem = True
End Function
Public Function zlPatiCardCheck(ByVal byt���ó��� As Byte, lng����ID As Long, str���� As String, bytˢ����ʽ As Byte) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鲡��ˢ����ʽ
    '��Σ�byt���ó���: 1-�Һ�;2-�շ�
    '         lng����ID:����ID(δ������,������)
    '         str����;δˢ��ʱ,Ϊ��
    '         bytˢ����ʽ: 1-����ˢ��;2-ҽ��ˢ��
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-04-27 16:09:08
    '˵����һ�����ŵ����ݲ��ˣ�ʹ�õ�ҽ����ͬʱҲ�Ǿ��￨��ҽԺҪ�������ҽ����ʽ����
    '          �����֤�Һš��շѣ����������Էѷ�ʽֱ��ˢ�����У����Ҫ���ڹҺš��շ�ʱ�����ݲ���ˢ�������������ҽ�������֤��ʽˢ�Ŀ���
    '          ����ֱ��ˢ�Ŀ�������ʾ�������������
    '����:29283
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = " Select Zl_Paticardcheck([1],[2],[3],[4]) as ��ʾ��Ϣ From Dual "
    ' Zl_Paticardcheck
    '  ���ó���_IN NUMBER ,
    '  ����id_In Number,
    '  ����_In   Varchar2,
    '  ˢ����ʽ_In Number:=1
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鲡��ˢ����ʽ�Ƿ�Ϸ�", byt���ó���, lng����ID, str����, bytˢ����ʽ)
    strSQL = nvl(rsTemp!��ʾ��Ϣ)
    If strSQL <> "" Then
        MsgBox strSQL, vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    zlPatiCardCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
    strSQL = "Select nvl(B.����,A.����) As ���� From ������ҳ A,������Ϣ B where a.����id=b.����id and  A.����id=[1] and a.��ҳid=[2] and ��Ŀ���� IS NOT NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ��Ѿ�����", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        zlPatiIS�����ѱ�Ŀ = False
    Else
        zlPatiIS�����ѱ�Ŀ = True
        If blnMsgbox Then
                MsgBox "���ˡ�" & nvl(rsTemp!����) & " ���Ѿ���Ŀ,��������м��ʻ����ʲ���!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Function
        End If
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlCheckIsMzToZY(ByVal strNos As String, ByVal int���� As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ�����תסԺ�����Ƿ��Ѿ����
    '���:strNos-���ݺ�(�ö��ŷ���)
    '        int����-�շѵ�;2-���ʵ�
    '����:�������,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2011-03-02 16:18:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strNO As String
    On Error GoTo errHandle
    strNO = Replace(strNos, "'", "")
     strSQL = "" & _
     "  Select /*+ rule */   1 From ������ü�¼ A,������˼�¼ B,Table(f_Str2list([1])) J" & _
     "  Where  A.NO=J.Column_Value and A.��¼����=[2] and A.ID=B.����ID  " & _
     "                  And  B.����  =1 and Rownum=1"
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ��������תסԺ����", strNO, int����)
    zlCheckIsMzToZY = Not rsTemp.EOF
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_GetInvoicePreperty(ByVal lngModule As Long, _
    ByVal intƱ�� As Integer, Optional strʹ����� As String) As Ty_FactProperty
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ʽ
    '���:intƱ��:1 - �շ��վ�, 2 - Ԥ���վ�, 3 - �����վ�, 4 - �Һ��վ�, 5 - ���￨, 12 - Ԥ����Ʊ
    '����:��Ʊ���������
    '����:���˺�
    '����:2011-07-19 16:43:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Ty_Fact As Ty_FactProperty, strFactType As String, varData As Variant, varTemp As Variant
    Dim strShareTypeUseID As String, lng����Ʊ�� As Long, lngʹ��Ʊ�� As Long
    Dim strFactTypeFormat As String, strFacePrintMode As String
    Dim intPrintMode As Long, intPrintMode1 As Long, lng����ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, lngFormat As Long, lngFormat1 As Long
    
    strFactType = Switch(intƱ�� = 1, "�����շ�Ʊ������", intƱ�� = 2, "����Ԥ��Ʊ������", intƱ�� = 12, "����Ԥ��Ʊ������", intƱ�� = 3, "���ý���Ʊ������", intƱ�� = 4, "���ùҺ�Ʊ������", intƱ�� = 5, "����ҽ�ƿ�����", True, "")
    strFactTypeFormat = Switch(intƱ�� = 1, "�շѷ�Ʊ��ʽ", intƱ�� = 2, "Ԥ����Ʊ��ʽ", intƱ�� = 12, "�˿Ʊ��ʽ", intƱ�� = 3, "���ʷ�Ʊ��ʽ", intƱ�� = 4, "�Һŷ�Ʊ��ʽ", intƱ�� = 5, "ҽ�ƿ���Ʊ��ʽ", True, "")
    strFacePrintMode = Switch(intƱ�� = 1, "�շѷ�Ʊ��ӡ��ʽ", intƱ�� = 2, "Ԥ����Ʊ��ӡ��ʽ", intƱ�� = 12, "Ԥ���˿��ӡ��ʽ", intƱ�� = 3, "���˽��ʴ�ӡ", intƱ�� = 4, "�Һŷ�Ʊ��ӡ��ʽ", intƱ�� = 5, "ҽ�ƿ���Ʊ��ӡ��ʽ", True, "")
    
    If strFactType = "" Then Exit Function
    '78751:���ϴ�,2014/10/20,����Ԥ��Ʊ�ݴ�ӡ��ʽ
    Ty_Fact.strUseType = strʹ�����
    '��ʼ��Ʊ��ʽ
'    If intƱ�� = 2 Then
'        'Ԥ�����޸�
'        Ty_Fact.intInvoiceFormat = 0
'    Else
        strFactTypeFormat = Trim(zlDatabase.GetPara(strFactTypeFormat, glngSys, lngModule, ""))
        '��ʽ:ʹ�����1,��ʽ1|ʹ�����2,��ʽ2...
        varData = Split(strFactTypeFormat, "|")
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",", ",")
            lngFormat = Val(varTemp(1))
            If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
            If Trim(varTemp(0)) = strʹ����� And lngFormat <> 0 Then
                Ty_Fact.intInvoiceFormat = lngFormat: Exit For
            End If
        Next
        If Ty_Fact.intInvoiceFormat = 0 And lngFormat1 <> 0 Then Ty_Fact.intInvoiceFormat = lngFormat
'    End If
    
    '��ӡ��ʽ(0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ)
    '����50656
'    If intƱ�� = 2 Then
'        'Ԥ����Ϊ�Զ���ӡ
'        Ty_Fact.intInvoicePrint = 1
'    Else
        '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
        strFacePrintMode = Trim(zlDatabase.GetPara(strFacePrintMode, glngSys, lngModule, ""))
        Ty_Fact.intInvoicePrint = -1
        '��ʽ:ʹ�����1,��ӡ��ʽ1|ʹ�����2,��ӡ��ʽ2...
        varData = Split(strFacePrintMode, "|")
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",,", ",")
            intPrintMode = Val(varTemp(1))
            If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
            If Trim(varTemp(0)) = strʹ����� Then
                Ty_Fact.intInvoicePrint = intPrintMode: Exit For
            End If
        Next
        If Ty_Fact.intInvoicePrint < 0 Then Ty_Fact.intInvoicePrint = intPrintMode1
'    End If
    '��������
    
    '��ʽ:����ID1,ʹ�����1|....
    strShareTypeUseID = Trim(zlDatabase.GetPara(strFactType, glngSys, lngModule, "0"))
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng����ID = Val(varTemp(0))
        If intƱ�� = 2 Or intƱ�� = 12 Or intƱ�� = 5 Then
            If Val(varTemp(1)) = 0 Then lng����Ʊ�� = lng����ID    '���õ�.
            If Val(varTemp(1)) = Val(strʹ�����) And lng����ID <> 0 Then
                lngʹ��Ʊ�� = lng����ID
            End If
        Else
            If Trim(varTemp(1)) = "" Then lng����Ʊ�� = lng����ID    '���õ�.
            If Trim(varTemp(1)) = strʹ����� And lng����ID <> 0 Then
                lngʹ��Ʊ�� = lng����ID
            End If
        End If
    Next
    
    On Error GoTo errHandle
    '����˳��
    '1.��ʹ��
    '2.ʹ��������ֵ�
    '3.����ʹ������
    strSQL = _
    "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
    "From Ʊ�����ü�¼" & vbNewLine & _
    "Where (ID =[1] or ID =[2]) And ʣ������ > 0   " & vbNewLine & _
    "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� Desc, ��ʼ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", lng����Ʊ��, lngʹ��Ʊ��)
    If rsTemp.EOF = False Then
        Ty_Fact.lngShareUseID = Val(nvl(rsTemp!ID)) '���õ�����ID
    End If
    zl_GetInvoicePreperty = Ty_Fact
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zl_GetInvoiceUserType(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ʹ�����
    '����:��Ʊ��ʹ�����
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select  Zl_Billclass([1],[2],[3]) as ʹ����� From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡƱ��ʹ�����", lng����ID, lng��ҳID, intInsure)
    zl_GetInvoiceUserType = nvl(rsTemp!ʹ�����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zl_GetInvoiceShareID(ByVal lngModule As Long, Optional strʹ����� As String = "") As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ�Ĺ���Ʊ��ID
    '����:���������ID
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeUseID As String
    Dim lng����ID As Long '���������ID
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lng����Ʊ�� As Long, lngʹ��Ʊ�� As Long
    
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    If lngModule = 1137 Then
        strShareTypeUseID = Trim(zlDatabase.GetPara("���ý���Ʊ������", glngSys, lngModule, "0"))
        '��ʽ:����ID1,ʹ�����1|....
    Else
        strShareTypeUseID = Trim(zlDatabase.GetPara("�����շ�Ʊ������", glngSys, lngModule, "0"))
        '��ʽ:����ID1,ʹ�����1|....
    End If
    
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng����ID = Val(varTemp(0))
        If Trim(varTemp(1)) = "" Then lng����Ʊ�� = lng����ID    '���õ�.
        If Trim(varTemp(1)) = strʹ����� And lng����ID <> 0 Then
            lngʹ��Ʊ�� = lng����ID
        End If
    Next
    On Error GoTo errHandle
    '����˳��
    '1.��ʹ��
    '2.ʹ��������ֵ�
    '3.����ʹ������
    strSQL = _
    "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
    "From Ʊ�����ü�¼" & vbNewLine & _
    "Where (ID =[1] or ID =[2]) And ʣ������ > 0   " & vbNewLine & _
    "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� Desc, ��ʼ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", lng����Ʊ��, lngʹ��Ʊ��)
    If rsTemp.EOF = False Then
        zl_GetInvoiceShareID = Val(nvl(rsTemp!ID))
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlStartFactUseType(ByVal intƱ�� As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ʹ����ʹ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-10 16:11:47
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select  1 as ���� From Ʊ�����ü�¼ where Ʊ��=[1] and nvl(ʹ�����,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���Ʊ���Ƿ�������ʹ������", intƱ��)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetSpecialItemFee(strClass As String, Optional ByVal strPriceGrade As String, Optional ByVal lng�շ�ϸĿID As Long) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ѡ����￨����дסԺ���ü�¼ʱ�ı�����Ϣ(�շ����,�շ�ϸĿID,���㵥λ,������ĿID,������Ŀ,�վݷ�Ŀ,ԭ��,�ּ�,�Ƿ���,���ұ�־)
    '���:
    '   strClass=�����ѡ����￨��������
    '   strPriceGrade ��ͨ�۸�ȼ�
    '����:ָ�����������ķ��ü�
    '����:���˺�
    '����:2011-07-07 02:17:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strWherePriceGrade As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    
    If lng�շ�ϸĿID = 0 Then
        strSQL = _
            "Select a.��� As �շ����, a.Id As �շ�ϸĿid, a.���㵥λ, c.Id As ������Ŀid, Nvl(a.���ηѱ�, 0) As ���ηѱ�, c.���� As ������Ŀ, c.�վݷ�Ŀ, b.ԭ��, b.�ּ�," & vbNewLine & _
            "       Nvl(b.ȱʡ�۸�, 0) ȱʡ�۸�, Nvl(a.�Ƿ���, 0) As �Ƿ���, Nvl(a.ִ�п���, 0) As ���ұ�־" & vbNewLine & _
            "From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շ��ض���Ŀ D" & vbNewLine & _
            "Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And d.�շ�ϸĿid = a.Id And d.�ض���Ŀ = [1]" & vbNewLine & _
            "      And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "      And Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    Else
        strSQL = _
            "Select a.��� As �շ����, a.Id As �շ�ϸĿid, a.���㵥λ, c.Id As ������Ŀid, Nvl(a.���ηѱ�, 0) As ���ηѱ�, c.���� As ������Ŀ, c.�վݷ�Ŀ, b.ԭ��, b.�ּ�," & vbNewLine & _
            "       Nvl(b.ȱʡ�۸�, 0) ȱʡ�۸�, Nvl(a.�Ƿ���, 0) As �Ƿ���, Nvl(a.ִ�п���, 0) As ���ұ�־" & vbNewLine & _
            "From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C " & vbNewLine & _
            "Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And A.ID = [3]" & vbNewLine & _
            "      And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "      And Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ض���Ŀ�ķ��ü�", strClass, strPriceGrade, lng�շ�ϸĿID)
    If Not rsTmp.EOF Then Set zlGetSpecialItemFee = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlGetUnitID(bytFlag As Byte, lngID As Long) As Long
'���ܣ������շ��ض���Ŀ��ִ�п���
'������bytFlag=ִ�п��ұ�־,lngID=�շ�ϸĿID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '����ȷ����
            zlGetUnitID = UserInfo.����ID 'ȡ����Ա���ڿ���
        Case 4 'ָ������
            strSQL = "Select B.ִ�п���ID From �շ���ĿĿ¼ A,�շ�ִ�п��� B Where B.�շ�ϸĿID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                zlGetUnitID = rsTmp!ִ�п���ID 'Ĭ��ȡ��һ��(���ж��)
            Else
                zlGetUnitID = UserInfo.����ID '��û��ָ������ȡ����Ա���ڿ���
            End If
        Case 1, 2, 3 '���˿���,����Ա����
            zlGetUnitID = UserInfo.����ID '��ȡ����Ա����
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlGetSaveCardFeeSQL(ByVal lngCardTypeID As Long, bytStyle As Byte, strNO As String, lng����ID As Long, lng��ҳID As Long, _
        lng���˲���ID As Long, lng���˿���ID As Long, lng��ʶ�� As Long, str�ѱ� As String, _
        strԭ���� As String, str���� As String, str�Ա� As String, str���� As String, str���� As String, str���� As String, _
        str�䶯ԭ�� As String, curӦ�ս�� As Double, curʵ�ս�� As Double, str���㷽ʽ As String, dt����ʱ�� As Date, lng����ID As Long, rsMoney As ADODB.Recordset, _
        ByVal strICCard As String, _
        Optional lngˢ�����ID As Long, Optional bln���ѿ� As Boolean, Optional strˢ������ As String, Optional lng����ID As Long, Optional strժҪ As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ��ҽ�ƿ����ü�¼SQL���
    '���:bytStyle=0-����,1-����,2-����
    '       cur���=���￨���
    '       str���㷽ʽ=���Ϊ��,��ʾ����,�����ֽ�
    '       rsMoney:�������￨�շ���Ϣ�ļ�¼��
    '       strԭ����=������ʱ��
    '       lng����ID=��ǰ���õľ��￨����ID
    '       str����-�������oracle�ĵ����Ż�Ϊ��
    '       strICCard=IC����,ͨ����IC����ʽ����ʱ,ͬʱ��д������Ϣ��IC���ֶ�
    '����:ҽ�ƿ����ü�¼SQL���
    '����:���˺�
    '����:2011-07-08 01:08:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngUnitID As Long, strSQL As String
    
    '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
    Select Case rsMoney!���ұ�־
        Case 4 'ָ������
            lngUnitID = zlGetUnitID(rsMoney!���ұ�־, rsMoney!�շ�ϸĿID)
        Case 1, 2 '���˿���
            If lng���˿���ID <> 0 Then
                lngUnitID = lng���˿���ID
            Else
                lngUnitID = UserInfo.����ID
            End If
        Case 0, 3, 5, 6
            lngUnitID = UserInfo.����ID
    End Select
  'Zl_ҽ�ƿ���¼_Insert
    strSQL = "Zl_ҽ�ƿ���¼_Insert("
    '  --��������������=0-����,1-����,2-����(�൱���ش�)
    '  --      ����ʱ,���ݺ�_IN�������ԭ��/�����ĵ��ݺš�
    '  --      ����/������,�ٻ���ʱ�������һ�ο���Ϊ׼��
    '  ��������_In   Number,
    strSQL = strSQL & "" & bytStyle & ","
    '  ���ݺ�_In     סԺ���ü�¼.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  ����id_In     סԺ���ü�¼.����id%Type,
    strSQL = strSQL & "'" & lng����ID & "',"
    '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type,
    strSQL = strSQL & "" & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & ","
    '  ��ʶ��_In     סԺ���ü�¼.��ʶ��%Type,
    strSQL = strSQL & "" & IIf(lng��ʶ�� = 0, "NULL", lng��ʶ��) & ","
    '  �ѱ�_In       סԺ���ü�¼.�ѱ�%Type,
    strSQL = strSQL & "'" & str�ѱ� & "',"
    '  �����id_In   ҽ�ƿ����.ID%Type,
    strSQL = strSQL & "" & lngCardTypeID & ","
    '  ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & IIf(strԭ���� = "", "NULL", "'" & strԭ���� & "'") & ","
    '  ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & IIf(str���� = "", "NULL", "'" & str���� & "'") & ","
    '  �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
    strSQL = strSQL & IIf(str�䶯ԭ�� = "", "NULL", "'" & str�䶯ԭ�� & "'") & ","
    '  ����_In       ������Ϣ.����֤��%Type,
    strSQL = strSQL & IIf(str���� = "", "NULL", "'" & str���� & "'") & ","
    '  ����_In       סԺ���ü�¼.����%Type,
    strSQL = strSQL & IIf(str���� = "", "NULL", "'" & str���� & "'") & ","
    '  �Ա�_In       סԺ���ü�¼.�Ա�%Type,
    strSQL = strSQL & IIf(str�Ա� = "", "NULL", "'" & str�Ա� & "'") & ","
    '  ����_In       סԺ���ü�¼.����%Type,
    strSQL = strSQL & IIf(str���� = "", "NULL", "'" & str���� & "'") & ","
    '  ���˲���id_In סԺ���ü�¼.���˲���id%Type,
    strSQL = strSQL & "" & lng���˲���ID & ","
    '  ���˿���id_In סԺ���ü�¼.���˿���id%Type,
    strSQL = strSQL & "" & lng���˿���ID & ","
    '  �շ�ϸĿid_In סԺ���ü�¼.�շ�ϸĿid%Type,
    strSQL = strSQL & "" & rsMoney!�շ�ϸĿID & ","
    '  �շ����_In   סԺ���ü�¼.�շ����%Type,
    strSQL = strSQL & "'" & rsMoney!�շ���� & "',"
    '  ���㵥λ_In   סԺ���ü�¼.���㵥λ%Type,
    strSQL = strSQL & "'" & nvl(rsMoney!���㵥λ) & "',"
    '  ������Ŀid_In סԺ���ü�¼.������Ŀid%Type,
    strSQL = strSQL & "" & rsMoney!������ĿID & ","
    '  �վݷ�Ŀ_In   סԺ���ü�¼.�վݷ�Ŀ%Type,
    strSQL = strSQL & "'" & nvl(rsMoney!�վݷ�Ŀ) & "',"
    '  ��׼����_In   סԺ���ü�¼.��׼����%Type,
    strSQL = strSQL & "" & curӦ�ս�� & ","
    '  ִ�в���id_In סԺ���ü�¼.ִ�в���id%Type,
    strSQL = strSQL & "" & lngUnitID & ","
    '  ��������id_In סԺ���ü�¼.��������id%Type,
    strSQL = strSQL & "" & UserInfo.����ID & ","
    '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �Ӱ��־_In   סԺ���ü�¼.�Ӱ��־%Type,
    strSQL = strSQL & "" & IIf(OverTime(dt����ʱ��), "1", "0") & ","
    '  ����ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(dt����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
    '  Ic����_In     ������Ϣ.Ic����%Type := Null,
    strSQL = strSQL & "'" & strICCard & "',"
    '  Ӧ�ս��_In   סԺ���ü�¼.Ӧ�ս��%Type,
    strSQL = strSQL & "" & curӦ�ս�� & ","
    '  ʵ�ս��_In   סԺ���ü�¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & curʵ�ս�� & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    strSQL = strSQL & "" & IIf(str���㷽ʽ = "", "NULL", "'" & str���㷽ʽ & "'") & ","
    '  ˢ�����id_In ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "" & IIf(lngˢ�����ID = 0, "NULL", lngˢ�����ID) & ","
    '  ���ѿ�_In     Integer := 0,
    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
    '  ˢ������_In   ����ҽ�ƿ���Ϣ.����%Type
    strSQL = strSQL & "'" & strˢ������ & "',"
    '  ����ID_IN
    strSQL = strSQL & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
    '  ������ˮ��_In
    strSQL = strSQL & "NULL,"
    '  ����˵��_In
    strSQL = strSQL & "NULL,"
    '  ������λ_In
    strSQL = strSQL & "NULL,"
    '  ժҪ_In   סԺ���ü�¼.ժҪ%Type,
    strSQL = strSQL & "" & IIf(strժҪ = "", "NULL", "'" & strժҪ & "'") & ")"
    
    zlGetSaveCardFeeSQL = strSQL
End Function
Public Function zlAddUpdateSwapSQL(ByVal blnԤ�� As Boolean, ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    str���� As String, str������ˮ�� As String, str����˵�� As String, _
    ByRef cllPro As Collection, Optional intУ�Ա�־ As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������ˮ�ź���ˮ˵��
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    '����:cllPro-����SQL��
    '����:���˺�
    '����:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "Zl_�����ӿڸ���_Update("
    '  �����id_In   ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "" & lng�����ID & ","
    '  ���ѿ�_In     Number,
    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
    '  ����_In       ����Ԥ����¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ����ids_In    Varchar2,
    strSQL = strSQL & "'" & strIDs & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type
    strSQL = strSQL & "'" & str����˵�� & "',"
    'Ԥ����ɿ�_In Number := 0
    strSQL = strSQL & "" & IIf(blnԤ��, 1, 0) & ","
    '�˷ѱ�־ :1-�˷�;0-����
    strSQL = strSQL & "0,"
    'У�Ա�־
    strSQL = strSQL & "" & IIf(intУ�Ա�־ = 0, "NULL", intУ�Ա�־) & ")"
    zlAddArray cllPro, strSQL
End Function

Public Function zlAddThreeSwapSQLToCollection(ByVal blnԤ���� As Boolean, _
    ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal str���� As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    ' ����:cllPro-����SQL��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
    'strExpend:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
    varData = Split(strExpend, "||")
    Dim str������Ϣ As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If zlCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                    str������Ϣ = Mid(str������Ϣ, 3)
                    'Zl_�������㽻��_Insert
                    strSQL = "Zl_�������㽻��_Insert("
                    '�����id_In ����Ԥ����¼.�����id%Type,
                    strSQL = strSQL & "" & lng�����ID & ","
                    '���ѿ�_In   Number,
                    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                    '����_In     ����Ԥ����¼.����%Type,
                    strSQL = strSQL & "'" & str���� & "',"
                    '����ids_In  Varchar2,
                    strSQL = strSQL & "'" & strIDs & "',"
                    '������Ϣ_In Varchar2:������Ŀ|��������||...
                    strSQL = strSQL & "'" & str������Ϣ & "',"
                    'Ԥ����ɿ�_In Number := 0
                    strSQL = strSQL & IIf(blnԤ����, "1", "0") & ")"
                    zlAddArray cllPro, strSQL
                    str������Ϣ = ""
                End If
                str������Ϣ = str������Ϣ & "||" & strTemp
            End If
        End If
    Next
    If str������Ϣ <> "" Then
        str������Ϣ = Mid(str������Ϣ, 3)
        'Zl_�������㽻��_Insert
        strSQL = "Zl_�������㽻��_Insert("
        '�����id_In ����Ԥ����¼.�����id%Type,
        strSQL = strSQL & "" & lng�����ID & ","
        '���ѿ�_In   Number,
        strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
        '����_In     ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '����ids_In  Varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '������Ϣ_In Varchar2:������Ŀ|��������||...
        strSQL = strSQL & "'" & str������Ϣ & "',"
        'Ԥ����ɿ�_In Number := 0
        strSQL = strSQL & IIf(blnԤ����, "1", "0") & ")"
        zlAddArray cllPro, strSQL
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsNotSucceedPrintBill(ByVal BytType As Byte, ByVal strNos As String, ByRef strOutValidNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ��Ѿ�������ӡ
    '���:bytType-1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '       strNos-���δ�ӡƱ�ݵĵ���,�ö��ŷ���
    '����:strOutValidNos-��ӡʧ�ܵĵ��ݺ�
    '����:���ڲ��湦Ʊ�ݵĴ�ӡ,����true,���򷵻�False
    '����:���˺�
    '����:2012-01-16 18:06:01
    '����:44322,44326,44332,44330
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempNos As String, rsTemp As ADODB.Recordset
    Dim strSQL As String, strBillNos As String
    Dim bytBill As Byte
    On Error GoTo errHandle
    strBillNos = Replace(Replace(strNos, "'", ""), " ", "")
    'Ӧȡ���һ�δ�ӡ��������
    strSQL = "" & _
        "Select  /*+ rule */ distinct  B.NO " & _
        " From Ʊ��ʹ����ϸ A,Ʊ�ݴ�ӡ���� B,Table( f_Str2list([2])) J" & _
        " Where A.��ӡID =b.ID And B.��������=[1] And B.No=J.Column_value "
        'And A.Ʊ��=b.��������:�п���ʹ�õ�������Ʊ��:����Һ�ʹ�������շ�Ʊ��
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���Ʊ���Ƿ��ӡ", BytType, strBillNos)
    
    strTempNos = ""
    With rsTemp
        Do While Not .EOF
            If InStr(1, "," & strBillNos & ",", "," & !NO & ",") = 0 Then
                strTempNos = strTempNos & "," & !NO
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then strTempNos = "," & strBillNos
    End With
    If strTempNos <> "" Then strTempNos = Mid(strTempNos, 2)
    rsTemp.Close: Set rsTemp = Nothing
    strOutValidNos = strTempNos
    zlIsNotSucceedPrintBill = strTempNos <> ""
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlIsCheckMedicinePayMode(ByVal strҽ�Ƹ������� As String, _
    Optional ByRef blnҽ�� As Boolean, Optional ByRef bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ�Ƹ��ʽ�Ƿ񹫷ѻ�ҽ��
    '���:strҽ�Ƹ�������-ҽ�Ƹ�������
    '����:blnҽ��-true,��ʾҽ��
    '        bln����-true,��ʾ�ǹ���
    '����:��ҽ���򹫷�ҽ��,����true,���򷵻�False
    '����:���˺�
    '����:2012-01-17 16:25:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "": blnҽ�� = False: bln���� = False
    On Error GoTo errHandle
    If grsҽ�Ƹ��ʽ Is Nothing Then
        strSQL = "Select ����,����,����,ȱʡ��־,�Ƿ�ҽ��,�Ƿ񹫷� From ҽ�Ƹ��ʽ"
    ElseIf grsҽ�Ƹ��ʽ.State <> 1 Then
        strSQL = "Select ����,����,����,ȱʡ��־,�Ƿ�ҽ��,�Ƿ񹫷� From ҽ�Ƹ��ʽ"
    End If
    If strSQL <> "" Then
        Set grsҽ�Ƹ��ʽ = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ�Ƹ��ʽ")
    End If
    grsҽ�Ƹ��ʽ.Find "����='" & strҽ�Ƹ������� & "'", , adSearchForward, 1
    If grsҽ�Ƹ��ʽ.EOF Then Exit Function
    blnҽ�� = Val(nvl(grsҽ�Ƹ��ʽ!�Ƿ�ҽ��)) = 1
    bln���� = Val(nvl(grsҽ�Ƹ��ʽ!�Ƿ񹫷�)) = 1
    zlIsCheckMedicinePayMode = blnҽ�� Or bln����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlLeftPad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���������ƿո�
    '����:�����ִ�
    '����:���˺�
    '����:2012-02-22 17:58:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlLeftPad = zlStr.LPAD(strCode, lngLen, strChar, True)
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
    zlSubstr = zlStr.SubB(strInfor, lngStart, lngLen)
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
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����״̬", strNO, "," & str��� & ",")
    If rsTemp.EOF Then
        rsTemp.Close: Set rsTemp = Nothing: Exit Function
    End If
    str������s = "": str����IDs = ""
    With rsTemp
        Do While Not .EOF
            str����IDs = str����IDs & "," & Val(nvl(rsTemp!ID))
            If InStr(1, str������s & vbCrLf, vbCrLf & nvl(rsTemp!������) & vbCrLf) = 0 Then
                str������s = str������s & vbCrLf & nvl(rsTemp!������)
            End If
            .MoveNext
        Loop
    End With
    If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    zlCheckIsExistsApplied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlExecuteChargeRollingCurtain(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���շ����ʹ���
    '���:frmMain-���õ�������
    '����:���˺�
    '����:2013-10-16 10:15:22
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCargeBill As Object
    Dim strCommon As String
    Dim intAtom As Integer
    Err = 0: On Error Resume Next
    Set objCargeBill = CreateObject("zL9CashBill.clsChargeBill")
    If Err <> 0 Then
        Set objCargeBill = Nothing
        MsgBox "�������ʲ���(zl9CashBill)ʧ��,���Ʋ�����ʧ��,�������Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    '6.1.7.1.    InitOracle:��ʼ������
    '���:
      '     cnMain-���ݿ�����
      '   strDBUser-���ݿ�������
      '     lngSys-ϵͳ��
      
     'ΪͨѶԭ�Ӹ�ֵ
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    '����ͨѶԭ��
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    If objCargeBill.InitOracle(gcnOracle, gstrDBUser, glngSys) = False Then
        Set objCargeBill = Nothing
        Exit Sub
    End If
    Call GlobalDeleteAtom(intAtom)
    'ChargeRollingCurtain(ByVal frmMain As Object)
    If objCargeBill.ChargeRollingCurtain(frmMain) = False Then
        Set objCargeBill = Nothing
        Exit Sub
    End If
    Set objCargeBill = Nothing
End Sub
 
Public Function zlIsPrintBill(ByVal lng����ID As Long, _
    ByVal lng����ID As Long, int���� As Integer, Optional strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ��ӡ��Ʊ�ݵ�
    '���:lng����ID-����ID-ָ���Ĳ���ID
    '       lng����ID-�Һ�ID(0-Ϊ���е�)
    '       int����-1-�շ�;3-����;4-�Һ�
    '       strNo-����=1��4ʱ,Ϊ�Һŵ���
    '����:
    '����:��ӡ��Ʊ�ݵķ���true,���򷵻�False
    '����:���˺�
    '����:2013-11-06 17:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
        
    If int���� = 4 Then
        strSQL = "" & _
        "   Select  1 " & _
        "   From ���˹Һż�¼ A,Ʊ�ݴ�ӡ���� B" & _
        "   Where a.NO=b.NO and B.��������=4 and A.����ID " & IIf(lng����ID <> 0, "+0", "") & " =[1] " & IIf(lng����ID <> 0, " And A.ID=[2]", "") & _
        "            And  Exists (Select 1 From Ʊ��ʹ����ϸ M Where b.Id = m.��ӡid And ���� = 1) And Rownum < 2  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ��ӡƱ��", lng����ID, lng����ID)
        zlIsPrintBill = Not rsTemp.EOF
        Exit Function
    End If
    
    If int���� = 1 Then '�շ�
        If strNO = "" And lng����ID <> 0 Then
            strSQL = "Select NO From ���˹Һż�¼ where ID=[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Һŵ���", lng����ID)
            If rsTemp.EOF Then Exit Function
            strNO = nvl(rsTemp!NO)
        End If
        strSQL = "" & _
        "   Select  1 " & _
        "   From ������ü�¼ A,Ʊ�ݴ�ӡ���� B" & _
        "   Where  a.NO=b.NO and B.��������=1 and a.��¼����=1 and A.����ID=[1] " & _
        IIf(strNO = "", "", "      And Exists(Select 1 From ����ҽ����¼ M Where �Һŵ�=[3]  And M.ID=A.ҽ�����) ") & _
        "      And Exists (Select 1 From Ʊ��ʹ����ϸ M Where b.Id = m.��ӡid And ���� = 1) And Rownum < 2  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ��ӡƱ��", lng����ID, lng����ID, strNO)
        zlIsPrintBill = Not rsTemp.EOF
        Exit Function
    End If
    
    strSQL = "" & _
    "   Select  1 " & _
    "   From ���˽��ʼ�¼ A,Ʊ�ݴ�ӡ���� B " & _
    "   Where  a.NO=b.NO and B.��������=3  and A.����ID=[1] " & _
    "      And Exists (Select 1 From Ʊ��ʹ����ϸ M Where b.Id = m.��ӡid And ���� = 1) And Rownum < 2  "
    If lng����ID <> 0 Then
        strSQL = strSQL & vbCrLf & _
        "        AND exists(SELECT 1 From סԺ���ü�¼ WHERE a.id=����ID  AND ����ID+0=[1] AND ��ҳID+0=[2])"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ��ӡƱ��", lng����ID, lng����ID)
    zlIsPrintBill = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlExistOperationData(ByVal lng����ID As Long, ByVal strNO As String, _
    Optional ByVal lng����ID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ǰ���˷�����ҵ������
    '���:lng����ID-����ID-ָ���Ĳ���ID
    '       strNo-�Һŵ���
    '       lng����ID-�Һ�ID
    '����:
    '����:����ҵ������,����true,���򷵻�False
    '����:���˺�
    '����:2013-11-06 17:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If strNO <> "" Then
        strSQL = "" & _
        "   Select 1 From ����ҽ����¼ A Where ����ID+0=[1] And �Һŵ�=[2]"
    ElseIf lng����ID <> 0 Then
        strSQL = "" & _
        "   Select 1 From ����ҽ����¼ A,���˹Һż�¼ B Where  A.�Һŵ�=B.NO And B.ID=[3] "
    Else
        strSQL = "" & _
        "   Select 1 From ����ҽ����¼ A Where ����ID =[1] AND ROWNUM<2"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж������Ƿ�����ҵ������", lng����ID, strNO, lng����ID)
    zlExistOperationData = rsTemp.EOF = False
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet��������() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-05 12:03:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select ���� From ���㷽ʽ where ����=9 And nvl(�Ƿ�̶�,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������")
    If Not rsTemp.EOF Then
        zlGet�������� = nvl(rsTemp!����)
    Else
        zlGet�������� = "����"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetԭ����ID(ByVal lng����ID As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݳ���ID��ȡԭ����ID
    '���:lng����ID-��ǰ����ID
    '����:
    '����:����ԭ����ID,0-ԭ����IDδ��ȡ��
    '����:���˺�
    '����:2014-06-13 17:26:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = " " & _
    "   Select A.����id,A.�Ǽ�ʱ�� " & _
    "   From ������ü�¼ A, " & _
    "       (   Select Max(NO) as NO,Max(�Ǽ�ʱ��) as �Ǽ�ʱ��   " & _
    "           From  ������ü�¼ Where ����ID=[1] ) B " & _
    "   Where a.No = B.NO And Mod(a.��¼����, 10) = 1 And ��¼״̬ In (1, 3)  " & _
    "         And  a.�Ǽ�ʱ��<= B.�Ǽ�ʱ�� " & _
    "   Order by A.�Ǽ�ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡԭ����ID", lng����ID)
    If Not rsTemp.EOF Then
     zlGetԭ����ID = Val(nvl(rsTemp!����ID))
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlReadBillFormat(ByVal ReportCode As String) As ADODB.Recordset
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ������Ĵ�ӡ��ʽ
    '���:ReportCode-��������
    '����:�����ӡ��ʽ�ļ�¼��
    '����:���ϴ�
    '����:2014-10-20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���='" & ReportCode & "'  " & _
    "   Order by  ���"
    Set zlReadBillFormat = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get������Ϣ�ӱ�(ByVal lng����ID As Long, Optional ByVal str��Ϣ�� As String = "") As ADODB.Recordset
'���ܣ�
'    ��ȡ������Ϣ�ӱ���
'����:
    Dim strSQL As String
    Dim intRet As Integer
    
    intRet = UBound(Split(str��Ϣ��, ","))
    If intRet = -1 Then '��ȡ�������дӱ���Ϣ
        strSQL = "Select ��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID =[1] And ��Ϣֵ is Not Null"
    ElseIf intRet = 0 Then '��ȡָ��ĳ���ӱ���Ϣ
        strSQL = "Select ��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID =[1] And ��Ϣ��='" & Split(str��Ϣ��, ",")(0) & "'" & " And ��Ϣֵ is Not Null "
    ElseIf intRet > 0 Then '��ȡָ���Ķ���ӱ���Ϣֵ
        strSQL = "Select ��Ϣ��, ��Ϣֵ" & vbNewLine & _
            "From ������Ϣ�ӱ�" & vbNewLine & _
            "Where ����id = [1] And" & vbNewLine & _
            "      ��Ϣ�� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) And ��Ϣֵ is Not Null "
    End If
    
    On Error GoTo errH
    Set Get������Ϣ�ӱ� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˴ӱ�", lng����ID, str��Ϣ��)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

