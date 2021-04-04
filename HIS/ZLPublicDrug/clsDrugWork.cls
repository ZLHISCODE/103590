VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugWork"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function zlCheckPriceAdjustBySell(ByVal lngҩƷid As Long, ByVal lngҩ��id As Long) As Boolean
    '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
    '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    '���۳���ʱֻ�ж�ҩ��
    '���أ�True-�����������۳��⣻false-���ܽ������۳���
    
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '���û����ȫ�ֵ����۹������򲻽��к�����飬����true
    If Val(zldatabase.GetPara(274, 100, , 0)) = 0 Then zlCheckPriceAdjustBySell = True: Exit Function
        
    gstrSQL = "Select ҩƷid, ͨ����, ���, 0 As �ⷿid, '' As �ⷿ, ������, '' As ����, ����, ��λ, ҩ���װ, �ۼ�, Sum(�ɱ��� * ʵ������) / Sum(ʵ������) As �ɱ���, �Ƿ�ʱ��" & vbNewLine & _
        " From (Select a.ҩƷid, '[' || c.���� || ']' || c.���� As ͨ����, c.���, c.���� As ������, Null As ����, a.ҩ�ⵥλ As ��λ, a.ҩ���װ, b.�ּ� As �ۼ�," & vbNewLine & _
        "              d.ƽ���ɱ��� As �ɱ���, 0 As �Ƿ�ʱ��, d.ʵ������" & vbNewLine & _
        "       From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D" & vbNewLine & _
        "       Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And" & vbNewLine & _
        "             (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0 And Nvl(a.�Ƿ����۹���, 0) = 1 And" & vbNewLine & _
        "             b.�ּ� <> d.ƽ���ɱ��� And d.�ⷿid In(SELECT Distinct ����id From ��������˵��  WHERE �������� LIKE '%ҩ��') And a.ҩƷid = [1])" & vbNewLine & _
        " Group By ҩƷid, ͨ����, ���, ������, ����, ��λ, ҩ���װ, �ۼ�" & vbNewLine & _
        " Having Sum(ʵ������) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select a.ҩƷid, '[' || c.���� || ']' || c.���� As ͨ����, c.���, d.�ⷿid, e.���� As �ⷿ, d.�ϴβ��� As ������, d.�ϴ����� As ����, d.����," & vbNewLine & _
        "       a.ҩ�ⵥλ As ��λ, a.ҩ���װ, d.���ۼ� As �ۼ�, d.ƽ���ɱ��� As �ɱ���, 1 As �Ƿ�ʱ��" & vbNewLine & _
        " From ҩƷ��� A, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E" & vbNewLine & _
        " Where a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And c.�Ƿ��� = 1 And" & vbNewLine & _
        "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.�Ƿ����۹���, 0) = 1 And d.���ۼ� <> d.ƽ���ɱ���" & vbNewLine & _
        " And d.�ⷿid In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') And a.ҩƷid = [1] And d.�ⷿid = [2] " & _
        " Order By ͨ����,�ⷿid,����"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "zlPublicDrug��zlCheckPriceAdjustBySell", lngҩƷid, lngҩ��id)
    
    'û�ҵ����������۹���Ҫ��ļ�¼������true
    If rsData.EOF Then zlCheckPriceAdjustBySell = True: Exit Function
    
    '�ҵ����������۹���Ҫ��ļ�¼ʱ����ʾ������fals
    MsgBox "��(" & rsData!�ⷿ & ")��ҩƷ" & rsData!ͨ���� & "���������۹�����Ҫ�󣺳ɱ��ۺ��ۼ۲�һ�£��������۳��⡣" & vbCrLf & "����ϵҩ����ҩ���ƽ��е��۴�����", vbInformation, "zlPublicDrug"
    zlCheckPriceAdjustBySell = False
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function zlGetDrugAlllAvailable(ByVal lngҩƷid As Long, ByRef dbl�������� As Double) As Boolean
    '����ָ��ҩƷ������ҩ���Ŀ������������ܿⷿ������
    '1.����ָ��ҩƷ�Ŀ�������
    '2.�ų�Ч�ڹ��ڵĿ��ҩƷ��¼
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Nvl(Sum(��������), 0) As �������� " & _
        " From ҩƷ��� Where ���� = 1 And ҩƷid = [1] And (Ч�� > Sysdate Or Ч�� Is Null) "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAlllAvailable", lngҩƷid)
    
    dbl�������� = rsTemp!��������
    
    zlGetDrugAlllAvailable = True
End Function


Private Function zlGetDrugAvailable(ByVal lng�ⷿid As Long, ByVal lngҩƷid As Long, ByRef dbl�������� As Double) As Boolean
    '����ָ���ⷿ��ҩƷ�Ŀ�����������������
    '1.����ָ���ⷿ��ҩƷ�Ŀ�������
    '2.�ų�Ч�ڹ��ڵĿ��ҩƷ��¼
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Nvl(Sum(��������), 0) As �������� " & _
        " From ҩƷ��� Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And (Ч�� > Sysdate Or Ч�� Is Null) "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAvailable", lng�ⷿid, lngҩƷid)
    
    dbl�������� = rsTemp!��������
    
    zlGetDrugAvailable = True
End Function


Private Function zlCheckDrugDosage(ByVal byt���� As Byte, ByVal strNos As String, ByRef strDosageNos As String) As Boolean
    '�ж�ҩƷ�Ƿ��Ѿ���ҩ
    'ҩƷ�շ���¼�и��ݴ�����������ȡ��Ӧ���շѣ��������͵ķ�ҩ���ݣ����жϵ����Ƿ����Ѿ���ҩ��ҩƷ����������ҩ�ĵ��ݺ�
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select /*+ rule*/ a.No " & _
        " From δ��ҩƷ��¼ A, Table(Cast(f_Str2list([2]) As t_Strlist)) B " & _
        " Where a.���� = [1] And a.��ҩ�� Is Not Null And a.No = b.Column_Value Order By a.NO "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAvailable", IIf(byt���� = 1, 8, 9), strNos)
    
    strDosageNos = ""
    Do While Not rsTemp.EOF
        strDosageNos = IIf(strDosageNos = "", "", strDosageNos & ",") & rsTemp!NO
        
        rsTemp.MoveNext
    Loop
    
    zlCheckDrugDosage = True
End Function

Private Function zlGetDrugCurrentPrice(ByVal lng�ⷿid As Long, ByVal lngҩƷid As Long, ByVal dbl�������� As Double, ByRef dbl������� As Double, ByRef dbl����۸� As Double, ByRef dbl������ As Double) As Boolean
    '��ȡʱ��ҩƷ�ۼۺ�Ӧ�ս��
    '1.����ҩƷID�ͿⷿID��ȡʱ��ҩƷ���ۼۼ������ܽ��(Ӧ�ս��)
    '2.�ۼ۵�ȷ��:�ܽ��/������;���ֻ��һ�����Σ�ֱ��ȡҩƷ�������ۼ�
    '3.����ѡ����Ҫ���ݳ��������д���(��:ϵͳ����:150-ҩƷ���������㷨(0-�������Ƚ��ȳ���1-��Ч������ȳ�))
    '4.�����ǰ��治�㣬ֻ��������������������Ľ��
    '5.�۸񣬽��ľ��ȸ���ϵͳ����ȷ��(���þ���)
    Dim rsTemp As ADODB.Recordset
    Dim dblʣ������ As Double
    Dim dblʵ�ʳ������� As Double
    Dim dbl��ǰ���� As Double
    Dim dbl��ǰ�۸� As Double
    Dim blnOnlyOneRs As Boolean     '�Ƿ�ֻ��һ������¼
    
    dbl������� = 0
    dbl����۸� = 0
    dbl������ = 0
    
    gstrSQL = "Select Nvl(����,0) As ����, Ч��, Nvl(ʵ������, 0) As ʵ������, Nvl(ʵ�ʽ��, 0) As ʵ�ʽ��, Nvl(ʵ�ʲ��, 0) As ʵ������, nvl(���ۼ�,0) As ���ۼ� " & _
        " From ҩƷ��� Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] "
    If gtype_UserSysParms.P150_ҩƷ���������㷨 = 1 Then
        gstrSQL = gstrSQL & " Order by Ч��,���� "
    Else
        gstrSQL = gstrSQL & " Order by ���� "
    End If
        
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "zlGetDrugCurrentPrice", lng�ⷿid, lngҩƷid)
    
    'ʵ���㷨
    '1.��ʵ������������
    '2.Ĭ�����ϸ���ƿ�棬��治�㰴���������������
    '3.ʱ�۷���ҩƷ����ҩƷ���.���ۼۡ�ȡ�۸񣬷��򰴡�ʵ�ʽ��/ʵ������������۸�
    '4.���ʵ������С��0����0ȡֵ�����Ҽ۸�ͽ��ҲΪ0
    '5.�����λ�Ч��˳�����ηֽ��������
    
    If rsTemp.RecordCount = 0 Then
        zlGetDrugCurrentPrice = True
        Exit Function
    ElseIf rsTemp.RecordCount = 1 Then
        blnOnlyOneRs = True
    End If
        
    dblʣ������ = dbl��������
    
    Do While Not rsTemp.EOF
        dbl��ǰ�۸� = 0
        dbl��ǰ���� = 0
        
        'ʱ�����ȡ���ۼۣ����򰴿�����������
        If rsTemp!���� > 0 And rsTemp!���ۼ� > 0 Then
            dbl��ǰ�۸� = Round(rsTemp!���ۼ�, gtype_UserSysParms.P157_���õ��۱���λ��)
        ElseIf rsTemp!ʵ������ > 0 Then
            dbl��ǰ�۸� = Round(rsTemp!ʵ�ʽ�� / rsTemp!ʵ������, gtype_UserSysParms.P157_���õ��۱���λ��)
        End If
        
        '�������ʣ��������������Ҳ��ʵ�������ɹ�����ʱ
        If rsTemp!ʵ������ > 0 And dblʣ������ > 0 Then
            If dblʣ������ - rsTemp!ʵ������ <= 0 Then
                '���ηֽ�͹��ˣ�û��ʣ����
                dbl��ǰ���� = dblʣ������
                dblʣ������ = 0
            Else
                '��ʣ�࣬��Ҫ���¸����ηֽ�
                dbl��ǰ���� = rsTemp!ʵ������
                dblʣ������ = dblʣ������ - rsTemp!ʵ������
            End If
            
            '�������ηֽ�ĳ����������㱾���εĳ�����Ͳ���¼�ۼƳ�������
            '����
            dbl������ = dbl������ + Round(dbl��ǰ�۸� * dbl��ǰ����, gtype_UserSysParms.P9_���ý���λ��)
            
            dblʵ�ʳ������� = dblʵ�ʳ������� + dbl��ǰ����
        End If
        
        '������ʵ�������ϼ�
        '����
        dbl������� = dbl������� + rsTemp!ʵ������
    Loop
    
    If blnOnlyOneRs = True Then
        'ֻ��һ������¼ʱ��ֱ��ȡǰ���Ѽ���ļ۸�
        '����
        dbl����۸� = dbl��ǰ�۸�
    ElseIf dbl������ > 0 And dblʵ�ʳ������� > 0 Then
        '�����ĺϼƳ�����ͳ��������������۸�
        '����
        dbl����۸� = Round(dbl������ / dblʵ�ʳ�������, gtype_UserSysParms.P157_���õ��۱���λ��)
    End If
    
    zlGetDrugCurrentPrice = True
    
End Function

Private Function zlGetDrugReturnAmount(ByVal byt���� As Byte, ByVal strNos As String, ByRef rsData As ADODB.Recordset) As Boolean
    '��ȡҩƷ��ҩ���ݵ�׼������
    '���ݴ���ĵ������ʺ͵��ݺŹ���ҩƷ�����ñ���ѯ
    '���ؼ�¼����NO,����ID,�շ�ϸĿID,ԭʼ��������ִ��������׼������
    
    gstrSQL = "Select /*+ rule*/" & vbNewLine & _
        " NO, ����id, ҩƷid As �շ�ϸĿid, Sum(ԭʼ����) As ԭʼ����, Sum(׼������) As ��ִ������, Sum(׼������) As ׼������" & vbNewLine & _
        "From (Select a.No, a.��¼״̬, a.���, a.ҩƷid, a.����, a.����id, a.ʵ������ * Nvl(a.����, 1) As ԭʼ����, 0 As ׼������" & vbNewLine & _
        "       From ҩƷ�շ���¼ A, Table(Cast(f_Str2list([2]) As t_Strlist)) B" & vbNewLine & _
        "       Where a.���� = [1] And a.No = b.Column_Value And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0) And a.������� Is Not Null" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select a.No, a.��¼״̬, a.���, a.ҩƷid, a.����, a.����id, 0 As ԭʼ����, a.ʵ������ * Nvl(a.����, 1) As ׼������" & vbNewLine & _
        "       From ҩƷ�շ���¼ A, Table(Cast(f_Str2list([2]) As t_Strlist)) B" & vbNewLine & _
        "       Where a.���� = [1] And a.No = b.Column_Value And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0 Or Mod(a.��¼״̬, 3) = 2) And" & vbNewLine & _
        "             a.������� Is Not Null)" & vbNewLine & _
        "Group By NO, ����id, ҩƷid " & vbNewLine & _
        "Order By NO, ����id, ҩƷid "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "zlGetDrugReturnAmount", IIf(byt���� = 1, 8, 9), strNos)
    
End Function

Private Function zlTriggerDrugAjustPrice(ByVal strҩƷIDs As String) As Boolean
    '����ҩƷ����
    '���ݴ��������ж�ÿ��ҩƷ�Ƿ��������Ч���¼۸������������е��۴���
    Dim arrSql As Variant
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    arrSql = Array()
     
    gstrSQL = "Select Distinct n.Id " & _
               " From �շ���ĿĿ¼ I, �շѼ�Ŀ N, ҩƷ��� P " & _
               " Where i.Id = n.�շ�ϸĿid And i.Id = p.ҩƷid And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And " & _
                   " n.�䶯ԭ�� = 0 And Sysdate>n.ִ������ " & _
               " Order By n.id "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "zlTriggerDrugAjustPrice����ѯ������Ч��¼")
    
    If rsTemp.RecordCount = 0 Then Exit Function
    
    For i = 0 To rsTemp.RecordCount - 1
        gstrSQL = "Zl_ҩƷ�շ���¼_Adjust(" & rsTemp!Id & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    Next

    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "zlTriggerDrugAjustPrice��ִ�е���")
    Next
    
    zlTriggerDrugAjustPrice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function zlAutoSendDrug(ByVal lng�ⷿid As Long, ByVal byt���� As Byte, ByVal strNO As String, ByRef strSendSQL As String) As Boolean
    '�Զ���ҩ
    '���ݴ�����֯��ҩ��Ҫ�����ݣ�������֯�õķ�ҩ����SQL���
    
    strSendSQL = "zl_ҩƷ�շ���¼_������ҩ("
    '�ⷿID
    strSendSQL = strSendSQL & lng�ⷿid
    '����
    strSendSQL = strSendSQL & "," & IIf(byt���� = 1, 8, 9)
    'NO
    strSendSQL = strSendSQL & ",'" & strNO & "'"
    '��ҩ��(�����)
    strSendSQL = strSendSQL & ",'" & UserInfo.�û����� & "'"
    '��ҩ��(���뾭����ҩ����ʱ������ҩ�˲���)
    strSendSQL = strSendSQL & ",'" & UserInfo.�û����� & "'"
    'У���ˣ�����ҽ����
    strSendSQL = strSendSQL & ","
    '��ҩ��ʽ
    strSendSQL = strSendSQL & ",1"
    '��ҩʱ��
    strSendSQL = strSendSQL & ","
    '����Ա����
    strSendSQL = strSendSQL & ",'" & UserInfo.�û����� & "'"
    '����Ա����
    strSendSQL = strSendSQL & ",'" & UserInfo.�û����� & "'"
    '����λ��
    strSendSQL = strSendSQL & "," & gtype_UserSysParms.P9_���ý���λ��
    '�Զ���˼��˵�
    strSendSQL = strSendSQL & ","
    '�Ƿ�����
    strSendSQL = strSendSQL & ",1"
    '�˲���
    strSendSQL = strSendSQL & ","
    ''''����
    strSendSQL = strSendSQL & ")"
    
    zlAutoSendDrug = True
    Exit Function
End Function

Private Function zlGetReceiveNo(ByVal lng����id As Long, ByVal str��ʼʱ�� As String, ByVal str����ʱ�� As String, ByRef rsDrawNO As ADODB.Recordset) As Boolean
    '��ȡָ�����ҵ���ҩ��
    '���ݴ��β�ѯʱ�䷶Χ�ڵ�ָ����ҩ���ŵ���ҩ�ţ��������ݼ�����ʽ����
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ��ҩ�� From δ��ҩƷ��¼ Where �Է�����id + 0 = [1] And �������� Between [2] And [3] Order By ��ҩ��"
    Set rsDrawNO = zldatabase.OpenSQLRecord(gstrSQL, "zlGetReceiveNo", lng����id, CDate(Format(str��ʼʱ��, "yyyy-mm-dd hh:mm:ss")), CDate(Format(str����ʱ��, "yyyy-mm-dd hh:mm:ss")))
    
    zlGetReceiveNo = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Class_Initialize()
    Call GetUserInfo
    Call GetSysParms
    Call GetStockCheckRule
End Sub