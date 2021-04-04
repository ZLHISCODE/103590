VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    If zlGetComLib = False Then Exit Function
    
    Call GetUserInfo
'    Call GetSysParms
'    Call GetStockCheckRule
    
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckPriceAdjustBySell(ByVal lngҩƷID As Long, ByVal lngҩ��id As Long, Optional ByVal blnShowMsg As Boolean = True) As Boolean
    '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
    '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    '���۳���ʱֻ�ж�ҩ��
    '������lngҩƷid-ҩƷ���ID��lngҩ��id-���۳����ҩ��ID��blnShowMsg-�Ƿ���ʾ��ʾ(Ĭ��Ϊ��ʾ)
    '���أ�True-�����������۳��⣻false-���ܽ������۳���
    
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '���û����ȫ�ֵ����۹������򲻽��к�����飬����true
    If Val(gobjDatabase.GetPara(275, 100, , 0)) = 0 Then zlCheckPriceAdjustBySell = True: Exit Function
    
    '������޿��
    If lngҩƷID > 0 Then
        If lngҩ��id > 0 Then
            gstrSQL = "Select 1 from ҩƷ��� Where ����=1 and ҩƷid=[1] and �ⷿid=[2] " & _
                " And Not (���� = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0)"
        Else
            gstrSQL = "Select 1 from ҩƷ��� Where ����=1 and ҩƷid=[1] " & _
                " And Not (���� = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0)"
        End If
        Set rsData = gobjDatabase.OpenSQLRecord(gstrSQL, "zlCheckPriceAdjustBySell", lngҩƷID, lngҩ��id)
        
        If rsData.EOF Then
            '�޿��ʱ�����շѼ�Ŀȡ�ۼۣ���ҩƷ���ȡ�ɱ��ۣ����Ƚϼ۸�
            gstrSQL = "Select a.�ɱ���, b.�ּ� As �ۼ� " & _
                " From ҩƷ��� A, �շѼ�Ŀ B " & _
                " Where a.ҩƷid = b.�շ�ϸĿid And (Sysdate Between b.ִ������ And b.��ֹ����) And Nvl(a.�Ƿ����۹���, 0) = 1 " & _
                " And b.�ּ� <> a.�ɱ��� And a.ҩƷid = [1] "
            Set rsData = gobjDatabase.OpenSQLRecord(gstrSQL, "zlCheckPriceAdjustBySell", lngҩƷID)
            
            If rsData.EOF Then
                'û�ҵ���ʾ�۸�һ�»�δ�������۹���
                zlCheckPriceAdjustBySell = True
            Else
                '�ҵ���ʾ�۸�һ��
                zlCheckPriceAdjustBySell = False
            End If
            
            Exit Function
        End If
    End If
    

    gstrSQL = "Select ҩƷid, ͨ����, ���, 0 As �ⷿid, '' As �ⷿ, ������, '' As ����, ����, ��λ, ҩ���װ, �ۼ�, Sum(�ɱ��� * ʵ������) / Sum(ʵ������) As �ɱ���, �Ƿ�ʱ��" & vbNewLine & _
        " From (Select a.ҩƷid, '['|| c.���� || ']'|| c.����||decode(c.����,null,null,'('||c.����||')') ||c.��� As ͨ����, c.���, c.���� As ������, Null As ����, a.ҩ�ⵥλ As ��λ, a.ҩ���װ, b.�ּ� As �ۼ�," & vbNewLine & _
        "              d.ƽ���ɱ��� As �ɱ���, 0 As �Ƿ�ʱ��, d.ʵ������" & vbNewLine & _
        "       From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D" & vbNewLine & _
        "       Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And" & vbNewLine & _
        "             (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0 And Nvl(a.�Ƿ����۹���, 0) = 1 And" & vbNewLine & _
        "             b.�ּ� <> d.ƽ���ɱ��� And d.�ⷿid In(SELECT Distinct ����id From ��������˵��  WHERE �������� LIKE '%ҩ��') And a.ҩƷid = [1] And d.�ⷿid = [2] " & vbNewLine & _
        "  And Not (D.���� = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0))" & vbNewLine & _
        " Group By ҩƷid, ͨ����, ���, ������, ����, ��λ, ҩ���װ, �ۼ�, �Ƿ�ʱ�� " & vbNewLine & _
        " Having Sum(ʵ������) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select a.ҩƷid, '['|| c.���� || ']'|| c.����||decode(c.����,null,null,'('||c.����||')') ||c.��� As ͨ����, c.���, d.�ⷿid, e.���� As �ⷿ, d.�ϴβ��� As ������, d.�ϴ����� As ����, d.����," & vbNewLine & _
        "       a.ҩ�ⵥλ As ��λ, a.ҩ���װ, Nvl(d.���ۼ�,0) As �ۼ�, d.ƽ���ɱ��� As �ɱ���, 1 As �Ƿ�ʱ��" & vbNewLine & _
        " From ҩƷ��� A, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E" & vbNewLine & _
        " Where a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And c.�Ƿ��� = 1 And" & vbNewLine & _
        "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.�Ƿ����۹���, 0) = 1 And Nvl(d.���ۼ�,0) <> d.ƽ���ɱ���" & vbNewLine & _
        " And d.�ⷿid In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') And a.ҩƷid = [1] And d.�ⷿid = [2] " & vbNewLine & _
        " And Not (d.���� = 0 And d.�������� < 0 And d.ʵ������ = 0 And d.ʵ�ʽ�� = 0 And d.ʵ�ʲ�� = 0) " & vbNewLine & _
        " Order By ͨ����,�ⷿid,����"
    Set rsData = gobjDatabase.OpenSQLRecord(gstrSQL, "zlPublicDrug��zlCheckPriceAdjustBySell", lngҩƷID, lngҩ��id)
    
    'û�ҵ����������۹���Ҫ��ļ�¼������true
    If rsData.EOF Then zlCheckPriceAdjustBySell = True: Exit Function
    
    '�ҵ����������۹���Ҫ��ļ�¼ʱ����ʾ������false
    If blnShowMsg = True Then
        MsgBox "��(" & rsData!�ⷿ & ")��ҩƷ" & rsData!ͨ���� & "���������۹�����Ҫ�󣺳ɱ��ۺ��ۼ۲�һ�£��������۳��⡣" & vbCrLf & "����ϵҩ����ҩ���ƽ��е��۴�����", vbInformation, "zlPublicDrug"
    End If
    
    zlCheckPriceAdjustBySell = False
    
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Private Function zlGetDrugAlllAvailable(ByVal lngҩƷID As Long, ByRef dbl�������� As Double) As Boolean
    '����ָ��ҩƷ������ҩ���Ŀ������������ܿⷿ������
    '1.����ָ��ҩƷ�Ŀ�������
    '2.�ų�Ч�ڹ��ڵĿ��ҩƷ��¼
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Nvl(Sum(��������), 0) As �������� " & _
        " From ҩƷ��� Where ���� = 1 And ҩƷid = [1] And (Ч�� > Sysdate Or Ч�� Is Null) "
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAlllAvailable", lngҩƷID)
    
    dbl�������� = rsTemp!��������
    
    zlGetDrugAlllAvailable = True
End Function


Private Function zlGetDrugAvailable(ByVal lng�ⷿid As Long, ByVal lngҩƷID As Long, ByRef dbl�������� As Double) As Boolean
    '����ָ���ⷿ��ҩƷ�Ŀ�����������������
    '1.����ָ���ⷿ��ҩƷ�Ŀ�������
    '2.�ų�Ч�ڹ��ڵĿ��ҩƷ��¼
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Nvl(Sum(��������), 0) As �������� " & _
        " From ҩƷ��� Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And (Ч�� > Sysdate Or Ч�� Is Null) "
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAvailable", lng�ⷿid, lngҩƷID)
    
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
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAvailable", IIf(byt���� = 1, 8, 9), strNos)
    
    strDosageNos = ""
    Do While Not rsTemp.EOF
        strDosageNos = IIf(strDosageNos = "", "", strDosageNos & ",") & rsTemp!NO
        
        rsTemp.MoveNext
    Loop
    
    zlCheckDrugDosage = True
End Function

Private Function zlGetDrugCurrentPrice(ByVal lng�ⷿid As Long, ByVal lngҩƷID As Long, ByVal dbl�������� As Double, ByRef dbl������� As Double, ByRef dbl����۸� As Double, ByRef dbl������ As Double) As Boolean
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
        
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugCurrentPrice", lng�ⷿid, lngҩƷID)
    
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
    Set rsData = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugReturnAmount", IIf(byt���� = 1, 8, 9), strNos)
    
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
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlTriggerDrugAjustPrice����ѯ������Ч��¼")
    
    If rsTemp.RecordCount = 0 Then Exit Function
    
    For i = 0 To rsTemp.RecordCount - 1
        gstrSQL = "Zl_ҩƷ�շ���¼_Adjust(" & rsTemp!ID & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    Next

    For i = 0 To UBound(arrSql)
        Call gobjDatabase.ExecuteProcedure(CStr(arrSql(i)), "zlTriggerDrugAjustPrice��ִ�е���")
    Next
    
    zlTriggerDrugAjustPrice = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zlDrugAdjust(lngҩƷID As Long) As Boolean
    '���ܣ�ҩƷ����
    '���Σ�lngҩƷID-ҩƷid
    '���أ�True-�ɹ�,False-ʧ��
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    
    gstrSQL = _
        "Select ID From �շѼ�Ŀ " & _
        " Where Sysdate Between ִ������ And Nvl(��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
        " And Nvl(�䶯ԭ��,0)=0 And �շ�ϸĿID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "zlDrugAdjust", lngҩƷID)
    
    If Not rsTmp.EOF Then
        gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & rsTmp!ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, App.ProductName)
        rsTmp.MoveNext
    End If
    
    zlDrugAdjust = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
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
    Set rsDrawNO = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetReceiveNo", lng����id, CDate(Format(str��ʼʱ��, "yyyy-mm-dd hh:mm:ss")), CDate(Format(str����ʱ��, "yyyy-mm-dd hh:mm:ss")))
    
    zlGetReceiveNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
