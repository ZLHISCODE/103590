VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaxBill"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim strLeftNo As String         '��Ʊǰ׺
Dim strItems As String          '�վݷ�Ŀ�Ķ����ַ���
Dim aryItem() As String         '�վݷ�Ŀ�Ķ�������

Dim rsTemp As New ADODB.Recordset
Dim intCount As Integer

Public Function zlTaxUseable(ByVal intƱ�� As Integer) As Boolean
    '------------------------------------------------------
    '���ܣ���鱾���Ƿ�ʹ��˰�ش�ӡ
    '������intƱ��=1�������շ�Ʊ�ݣ�intƱ��=2��סԺ��Ʊ��
    '------------------------------------------------------
    If intƱ�� = 1 Then
        zlTaxUseable = Val(GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "����ʹ��˰Ʊ��ӡ", 0)) <> 0
    Else
        zlTaxUseable = Val(GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "סԺʹ��˰Ʊ��ӡ", 0)) <> 0
    End If
End Function

Public Function zlTaxBillSet(ByVal cnMain As ADODB.Connection, intƱ�� As Integer) As Boolean
    '------------------------------------------------------
    '���ܣ�����ָ��Ʊ�ݵı��ز���
    '������intƱ��=1�������շ�Ʊ�ݣ�intƱ��=2��סԺ��Ʊ��
    '------------------------------------------------------
    Set gcnOracle = cnMain
    Select Case intƱ��
        Case 1
            frmOutSet.Show 1
        Case 2
            frmInSet.Show 1
    End Select
End Function

Public Function zlTaxOutPrint(cnMain As ADODB.Connection, strNOs As String, Optional strExpend As String) As String
    '------------------------------------------------------
    '���ܣ������վݴ�ӡ(���������˷Ѻ���ٴ�ӡ)
    '������strNOs�����δ�ӡ�ĵ��ݺ����ַ���
    '      strExpend,��|�߷ָ�,��һλΪ��������(1-�շ� 2-�Һ�),�ڶ�λ����Ժ���չ
    '���أ��㳤���ַ�������ʾ�ɹ���������ʾ���������ԭ��
    '------------------------------------------------------
    Dim strUnions As String, strDecode As String
    Dim bln������ As Boolean, strMoneys As String
    Dim lngReturn As Long, strError As String * 255
    Dim strName As String, strInfo As String, intType As Integer
    
    Set gcnOracle = cnMain
    
    bln������ = Val(GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "����ʹ��������", 0)) <> 0
    strLeftNo = GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "����˰Ʊǰ׺", "")
    strItems = GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "����˰Ʊ��Ŀ", "")
    intType = Val(Split(strExpend & "|", "|")(0))
    If intType = 0 Then intType = 1
    aryItem = Split(strItems, "|")
    
    strDecode = ""
    For intCount = LBound(aryItem) To UBound(aryItem)
        strDecode = strDecode & ",'" & Split(aryItem(intCount), ";")(0) & "','" & Split(aryItem(intCount), ";")(1) & "'"
    Next
    
    strUnions = ""
    For intCount = 1 To 19
        strUnions = strUnions & " union all select '" & Format(intCount, "00") & "' as ��Ʊ��Ŀ,0 as ��� From Dual"
    Next
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select ���� From ������ü�¼ Where ��¼���� = " & IIf(intType = 1, "1", "4") & " And NO IN(" & strNOs & ") And Rownum=1"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "����˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If Not .EOF Then strName = IIf(IsNull(!����), "", !����)
        
        gstrSql = "Select ��Ʊ��Ŀ,sum(���) as ���" & _
                " From (Select decode(�վݷ�Ŀ" & strDecode & ",null) as ��Ʊ��Ŀ,ʵ�ս�� as ���" & _
                "       From ������ü�¼" & _
                "       Where ��¼���� = " & IIf(intType = 1, "1", "4") & " and NO In (" & strNOs & ")" & _
                        strUnions & ")" & _
                " Where ��Ʊ��Ŀ is not null" & _
                " Group By ��Ʊ��Ŀ"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "����˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        strMoneys = ""
        Do While Not .EOF
            strMoneys = strMoneys & ";" & Format(!���, "0.00")
            .MoveNext
        Loop
        strMoneys = Mid(strMoneys, 2)
        strMoneys = strMoneys & ";;0.00;;0.00;;0.00;;0.00;;0.00;" '5���Զ�������Ŀ
        
        gstrSql = "Select Max(ID) as ID From Ʊ�ݴ�ӡ���� Where �������� = " & IIf(intType = 1, "1", "4") & " And NO IN(" & strNOs & ")"
        gstrSql = "Select Distinct ����,ʹ���� From Ʊ��ʹ����ϸ" & _
                " Where ��ӡID=(" & gstrSql & ") Order by ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "����˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxOutPrint = "δ�ҵ���Ҫ���ݵ�Ʊ��": Exit Function
        If .RecordCount > 1 Then zlTaxOutPrint = "��Ҫ���ݵ�Ʊ�ݲ���ͬһƱ��": Exit Function
    
ReTry:
        strError = String(255, " ")
        If bln������ Then
            lngReturn = gobjTax.BJ_Normal_Invoice(2, strLeftNo & !����, strName, !ʹ����, strMoneys, strError)
        Else
            lngReturn = gobjTax.BJ_Normal_Invoice_NoPwd(2, strLeftNo & !����, strName, !ʹ����, strMoneys, strError)
        End If
        If lngReturn <> 0 Then
            strInfo = "˰�����ӿڵ���ʧ��(��Ʊ)��" & vbCrLf & vbCrLf & _
                "�����ţ�" & lngReturn & vbCrLf & "������Ϣ��" & zStr(strError) & vbCrLf & vbCrLf & _
                "�����豸�Ƿ�������������ӡ��ֽ���Ƿ񰲷���ȷ��" & vbCrLf & _
                "���������ԭ�������ѡ�����ԡ�Ҫ������"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxOutPrint = "" '����ʾ,���ó���������ʾ
                Exit Function
            End If
        End If
        zlTaxOutPrint = ""
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlTaxOutErase(cnMain As ADODB.Connection, strNOs As String, Optional strExpend As String) As String
    '------------------------------------------------------
    '���ܣ������վ�����
    '������strNo���������ϵĵ��ݺ����ַ���
    '���أ��㳤���ַ�������ʾ�ɹ���������ʾ���������ԭ��
    '------------------------------------------------------
    Dim lngReturn As Long, strError As String * 255
    Dim strInfo As String, intType As Integer
    
    Set gcnOracle = cnMain
    strLeftNo = GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "����˰Ʊǰ׺", "")
    
    intType = Val(Split(strExpend & "|", "|")(0))
    If intType = 0 Then intType = 1
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
'        gstrSql = "Select Max(ID) as ID From Ʊ�ݴ�ӡ���� Where ��������=1 And NO IN(" & strNOs & ")"
'        gstrSql = "Select Distinct ����,ʹ���� From Ʊ��ʹ����ϸ" & _
'                " Where ��ӡID=(" & gstrSql & ") Order by ����"
        
        '���ڲ����˷��ٴ�ӡ���൥���޸ĵ����,���ܸ����ϵ��������·���Ʊ�ݡ�
        gstrSql = "Select ����,ʹ����" & _
                 " From Ʊ��ʹ����ϸ" & _
                 " Where Id = (" & _
                 "       Select Max(Id)" & _
                 "       From Ʊ��ʹ����ϸ" & _
                 "       Where ���� = 2 And ��ӡid In (" & _
                 "             Select Id From Ʊ�ݴ�ӡ���� Where �������� = " & IIf(intType = 1, "1", "4") & " And No In (" & strNOs & ")))"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "����˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxOutErase = "δ�ҵ���Ҫ���ϵ�Ʊ��": Exit Function
        If .RecordCount > 1 Then zlTaxOutErase = "��Ҫ���ϵ�Ʊ�ݲ���ͬһƱ��": Exit Function
        
ReTry:
        strError = String(255, " ")
        lngReturn = gobjTax.BJ_Other_Invoice(1, 2, strLeftNo & !����, !ʹ����, "", strError)
        If lngReturn <> 0 Then
            strInfo = "˰�����ӿڵ���ʧ��(��Ʊ)��" & vbCrLf & vbCrLf & _
                "�����ţ�" & lngReturn & vbCrLf & "������Ϣ��" & zStr(strError) & vbCrLf & vbCrLf & _
                "�����豸�Ƿ�����������" & vbCrLf & _
                "���������ԭ�������ѡ�����ԡ�Ҫ������"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxOutErase = "" '����ʾ,���ó���������ʾ
                Exit Function
            End If
        End If
        zlTaxOutErase = ""
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlTaxOutReput(cnMain As ADODB.Connection, strNOs As String, Optional strExpend As String) As String
    '------------------------------------------------------
    '���ܣ������վ����´�ӡ
    '������strNo�����δ�ӡ�ĵ��ݺ����ַ���
    '���أ��㳤���ַ�������ʾ�ɹ���������ʾ���������ԭ��
    '------------------------------------------------------
    Dim lngReturn As Long, strError As String * 255
    Dim strOldInvioce As String, strInfo As String, intType As Integer
    
    intType = Val(Split(strExpend & "|", "|")(0))
    If intType = 0 Then intType = 1
    Set gcnOracle = cnMain
    strLeftNo = GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "����˰Ʊǰ׺", "")
    
    '���ڵ���ʱ�������Ѿ������£�������Ҫ��ȡ�ش�ǰ��ʵ��Ʊ�ݺ��룬�������ϲ���
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select ����,ʹ����" & _
                 " From Ʊ��ʹ����ϸ" & _
                 " Where Id = (" & _
                 "       Select Max(Id)" & _
                 "       From Ʊ��ʹ����ϸ" & _
                 "       Where ���� = 2 And ��ӡid In (" & _
                 "             Select Id From Ʊ�ݴ�ӡ���� Where �������� = " & IIf(intType = 1, "1", "4") & " And No In (" & strNOs & ")))"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "����˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxOutReput = "δ�ҵ��ϴεĴ�ӡ��¼": Exit Function
        If .RecordCount > 1 Then zlTaxOutReput = "��Ҫ�ش��Ʊ�ݲ���ͬһƱ��": Exit Function
        
ReTry:
        strError = String(255, " ")
        lngReturn = gobjTax.BJ_Other_Invoice(1, 2, strLeftNo & !����, !ʹ����, "", strError)
        If lngReturn <> 0 Then
            strInfo = "˰�����ӿڵ���ʧ��(��Ʊ)��" & vbCrLf & vbCrLf & _
                "�����ţ�" & lngReturn & vbCrLf & "������Ϣ��" & zStr(strError) & vbCrLf & vbCrLf & _
                "�����豸�Ƿ�����������" & vbCrLf & _
                "���������ԭ�������ѡ�����ԡ�Ҫ������"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxOutReput = "" '����ʾ,���ó���������ʾ
                Exit Function
            End If
        End If
    End With
    
    '���ô�ӡ������ʵ�ַ�Ʊ���ش����
    zlTaxOutReput = zlTaxOutPrint(cnMain, strNOs)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlTaxInPrint(cnMain As ADODB.Connection, lngId As Long) As String
    '------------------------------------------------------
    '���ܣ�סԺ�վݴ�ӡ
    '������lngId�����δ�ӡ�Ľ��ʵ���id
    '���أ��㳤���ַ�������ʾ�ɹ���������ʾ���������ԭ��
    '------------------------------------------------------
    Dim strDecode As String, bln������ As Boolean
    Dim strMoneys As String, intTimes As Integer
    Dim lngReturn As Long, strError As String * 255
    Dim strName As String, strInfo As String
    
    Set gcnOracle = cnMain
    
    bln������ = Val(GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "סԺʹ��������", 0)) <> 0
    strLeftNo = GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "סԺ˰Ʊǰ׺", "")
    strItems = GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "סԺ˰Ʊ��Ŀ", "")
    aryItem = Split(strItems, "|")
    
    strDecode = ""
    For intCount = LBound(aryItem) To UBound(aryItem)
        strDecode = strDecode & ",'" & Split(aryItem(intCount), ";")(0) & "','" & Split(aryItem(intCount), ";")(1) & "'"
    Next
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select B.���� From ������Ϣ B,���˽��ʼ�¼ A Where A.����ID=B.����ID And A.ID=" & lngId
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "����˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If Not .EOF Then strName = IIf(IsNull(!����), "", !����)
    
        gstrSql = "Select ��Ʊ��Ŀ,sum(���) as ���" & _
                " From (Select decode(�վݷ�Ŀ" & strDecode & ",null) as ��Ʊ��Ŀ,���ʽ�� as ���" & _
                "       From סԺ���ü�¼" & _
                "       Where ���ʷ���=1 and ����id=" & lngId & ")" & _
                " Where ��Ʊ��Ŀ is not null" & _
                " Group By ��Ʊ��Ŀ" & _
                " Having sum(���)<>0"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "סԺ˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        strMoneys = "": intTimes = 0
        Do While Not .EOF
            intTimes = intTimes + 1
            strMoneys = strMoneys & ";" & Mid(!��Ʊ��Ŀ, 1, InStr(1, !��Ʊ��Ŀ, "-") - 1) & ";" & _
                        Mid(!��Ʊ��Ŀ, InStr(1, !��Ʊ��Ŀ, "-") + 1) & ";1;" & Format(!���, "0.00") & ";"
            If intTimes = 4 Then Exit Do
            .MoveNext
        Loop
        strMoneys = intTimes & strMoneys
        
        gstrSql = "Select L.ʵ��Ʊ��,I.סԺ��,L.����Ա����" & _
                " From ���˽��ʼ�¼ L,������Ϣ I" & _
                " Where L.����id=I.����id And L.id=" & lngId
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "סԺ˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxInPrint = "δ�ҵ���Ҫ���ݵ�Ʊ��": Exit Function
        If .RecordCount > 1 Then zlTaxInPrint = "��Ҫ���ݵ�Ʊ�ݲ���ͬһƱ��": Exit Function
        
ReTry:
        strError = String(255, " ")
        If bln������ Then
            lngReturn = gobjTax.BJ_Normal_Invoice(1, strLeftNo & !ʵ��Ʊ��, strName, !����Ա����, strMoneys & ";" & !סԺ�� & ";", strError)
        Else
            lngReturn = gobjTax.BJ_Normal_Invoice_NoPwd(1, strLeftNo & !ʵ��Ʊ��, strName, !����Ա����, strMoneys & ";" & !סԺ�� & ";", strError)
        End If
        If lngReturn <> 0 Then
            strInfo = "˰�����ӿڵ���ʧ��(��Ʊ)��" & vbCrLf & vbCrLf & _
                "�����ţ�" & lngReturn & vbCrLf & "������Ϣ��" & zStr(strError) & vbCrLf & vbCrLf & _
                "�����豸�Ƿ�������������ӡ��ֽ���Ƿ񰲷���ȷ��" & vbCrLf & _
                "���������ԭ�������ѡ�����ԡ�Ҫ������"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxInPrint = "" '����ʾ,���ó���������ʾ
                Exit Function
            End If
        End If
        zlTaxInPrint = ""
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlTaxInErase(cnMain As ADODB.Connection, lngId As Long) As String
    '------------------------------------------------------
    '���ܣ�סԺ�վ�����
    '������lngId���������ϵĽ��ʵ���id
    '���أ��㳤���ַ�������ʾ�ɹ���������ʾ���������ԭ��
    '------------------------------------------------------
    Dim lngReturn As Long, strError As String * 255
    Dim strInfo As String
    
    Set gcnOracle = cnMain
    strLeftNo = GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "סԺ˰Ʊǰ׺", "")
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select L.ʵ��Ʊ��,L.����Ա����" & _
                " From ���˽��ʼ�¼ L Where L.id=" & lngId
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "סԺ˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxInErase = "δ�ҵ���Ҫ���ϵ�Ʊ��": Exit Function
        If .RecordCount > 1 Then zlTaxInErase = "��Ҫ���ϵ�Ʊ�ݲ���ͬһƱ��": Exit Function
    
ReTry:
        strError = String(255, " ")
        lngReturn = gobjTax.BJ_Other_Invoice(1, 1, strLeftNo & !ʵ��Ʊ��, !����Ա����, "", strError)
        If lngReturn <> 0 Then
            strInfo = "˰�����ӿڵ���ʧ��(��Ʊ)��" & vbCrLf & vbCrLf & _
                "�����ţ�" & lngReturn & vbCrLf & "������Ϣ��" & zStr(strError) & vbCrLf & vbCrLf & _
                "�����豸�Ƿ�����������" & vbCrLf & _
                "���������ԭ�������ѡ�����ԡ�Ҫ������"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxInErase = "" '����ʾ,���ó���������ʾ
                Exit Function
            End If
        End If
        zlTaxInErase = ""
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlTaxInReput(cnMain As ADODB.Connection, lngId As Long) As String
    '------------------------------------------------------
    '���ܣ�סԺ�վ����´�ӡ
    '������lngId�����δ�ӡ�Ľ��ʵ���id
    '���أ��㳤���ַ�������ʾ�ɹ���������ʾ���������ԭ��
    '------------------------------------------------------
    Dim lngReturn As Long, strError As String * 255
    Dim strOldInvioce As String, strInfo As String
    
    Set gcnOracle = cnMain
    strLeftNo = GetSetting("ZLSOFT", "����ȫ��\˰Ʊ��ӡ", "סԺ˰Ʊǰ׺", "")
    
    '���ڵ���ʱ�������Ѿ������£�������Ҫ��ȡ�ش�ǰ��ʵ��Ʊ�ݺ��룬�������ϲ���
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select ����,ʹ����" & _
                 " From Ʊ��ʹ����ϸ" & _
                 " Where Id In (" & _
                 "       Select Max(Id)" & _
                 "       From Ʊ��ʹ����ϸ" & _
                 "       Where ���� = 2 And ��ӡid In (" & _
                 "             Select Id From Ʊ�ݴ�ӡ���� Where ��������=3 And No In (" & _
                 "             Select No From ���˽��ʼ�¼ Where ID=" & lngId & ")))"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "סԺ˰��", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxInReput = "δ�ҵ��ϴεĴ�ӡ��¼": Exit Function
        If .RecordCount > 1 Then zlTaxInReput = "��Ҫ�ش��Ʊ�ݲ���ͬһƱ��": Exit Function
                
ReTry:
        strError = String(255, " ")
        lngReturn = gobjTax.BJ_Other_Invoice(1, 1, strLeftNo & !����, !ʹ����, "", strError)
        If lngReturn <> 0 Then
            strInfo = "˰�����ӿڵ���ʧ��(��Ʊ)��" & vbCrLf & vbCrLf & _
                "�����ţ�" & lngReturn & vbCrLf & "������Ϣ��" & zStr(strError) & vbCrLf & vbCrLf & _
                "�����豸�Ƿ�����������" & vbCrLf & _
                "���������ԭ�������ѡ�����ԡ�Ҫ������"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxInReput = "" '����ʾ,���ó���������ʾ
                Exit Function
            End If
        End If
    End With
    
    '���ô�ӡ������ʵ�ַ�Ʊ���ش����
    zlTaxInReput = zlTaxInPrint(cnMain, lngId)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Class_Initialize()
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    gstrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
End Sub