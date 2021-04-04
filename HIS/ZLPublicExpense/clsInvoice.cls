VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum gEM_BillType
    EM_�շ��վ� = 1
    EM_Ԥ���վ� = 2
    EM_�����վ� = 3
    EM_�Һ��վ� = 4
    EM_ҽ�ƿ��վ� = 5
    EM_�˷��վ� = 11
    EM_���ʺ�Ʊ = 13
End Enum

Public Function zlGetInvoiceUserType(ByVal lngModule As Long, ByVal bytBillType As gEM_BillType, _
    ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
    ByVal intInsure As Integer, ByRef strOutUseType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�ʹ�����
    '���:lngModule-ģ���
    '     bytBillType-Ʊ��
    '     lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     intInsure-����
    '����:strOutUseType-����ʹ�����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 10:14:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    strOutUseType = ""
'    If Not zlStartFactUseType(bytBillType) Then zlGetInvoiceUserType = True: Exit Function
    
    strSql = "Select  Zl_Billclass([1],[2],[3]) as ʹ����� From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ȡƱ��ʹ�����", lng����ID, lng��ҳId, intInsure)
    strOutUseType = Nvl(rsTemp!ʹ�����)
    zlGetInvoiceUserType = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function GetDrawShareInvoiceRec(ByVal bytBillType As gEM_BillType) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��Ʊ�ֵĹ���Ʊ������
    '���:bytBillType -Ʊ��
    '����:rsOutDraw-���ع���Ʊ�ݼ�(ID,������,�Ǽ�ʱ��,��ʼ����,��ֹ����,ʣ������)
    '����:������
    '����:2014-05-23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errH
    If bytBillType = 1 Or bytBillType = 3 Then  '�շѺͽ���
        strSql = "" & _
        "   Select A.ID,nvl(M.����,' ') as ʹ��������,A.ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B,Ʊ��ʹ����� M" & vbNewLine & _
        "   Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "           And A.ʹ�����=M.����(+) " & _
        "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "   Order by ʹ��������,ʣ������ Desc"
    ElseIf bytBillType = 5 Then
        '���￨
        strSql = "" & _
        "   Select A.ID,nvl(M.����,' ') as ʹ��������,M.ID as ʹ�����ID,M.���� as ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B,ҽ�ƿ���� M" & vbNewLine & _
        "   Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "           And to_number(nvl(A.ʹ�����,'0'))=M.ID(+) " & _
        "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "   Order by ʹ��������,ʣ������ Desc"
    ElseIf bytBillType = 2 Then  'Ԥ��
        strSql = "" & _
        "   Select A.ID,to_number(nvl(A.ʹ�����,'0')) as ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B" & vbNewLine & _
        "   Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "   Order by ʹ�����,ʣ������ Desc"
    Else
        strSql = "" & _
        "   Select A.ID,A.ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
        "   From Ʊ�����ü�¼ A,��Ա�� B" & vbNewLine & _
        "   Where A.Ʊ��=[1] And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
        "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "   Order by ʹ�����,ʣ������ Desc"
    End If
    Set GetDrawShareInvoiceRec = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, bytBillType)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetInvoiceShareID(ByVal lngModule As Long, _
    ByVal bytBillType As gEM_BillType, _
    Optional ByVal strUseType As String = "", _
    Optional ByRef lngShareInvoiceID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ�Ĺ���Ʊ��ID
    '���:lngModule-ģ���
    '     strUseType-ʹ�����
    '     bytBillType-Ʊ��
    '����:lngShareInvoiceID-���ع��õķ�ƱID
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 10:22:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeUseID As String
    Dim lng����ID As Long, rsTemp As ADODB.Recordset, strSql As String
    Dim lng����Ʊ�� As Long, lngʹ��Ʊ�� As Long, trFactType As String
    Dim strFactType As String
    
    lngShareInvoiceID = 0
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    If lngModule = 1137 Then
        strFactType = "" & _
        Switch(bytBillType = EM_�Һ��վ�, "���ùҺ�Ʊ������", _
                bytBillType = EM_�����վ�, "���ý���Ʊ������", _
                bytBillType = EM_�շ��վ�, "���ý���Ʊ������", _
                bytBillType = EM_ҽ�ƿ��վ�, "����ҽ�ƿ�����", _
                bytBillType = EM_Ԥ���վ�, "����Ԥ��Ʊ������", _
                True, "")
    Else
        strFactType = "" & _
        Switch(bytBillType = EM_�Һ��վ�, "���ùҺ�Ʊ������", _
                bytBillType = EM_�����վ�, "���ý���Ʊ������", _
                bytBillType = EM_�շ��վ�, "�����շ�Ʊ������", _
                bytBillType = EM_ҽ�ƿ��վ�, "����ҽ�ƿ�����", _
                bytBillType = EM_Ԥ���վ�, "����Ԥ��Ʊ������", _
                True, "")
    End If
    If strFactType = "" Then Exit Function
    
    '��ʽ:����ID1,ʹ�����1|....
    strShareTypeUseID = Trim(gobjDatabase.GetPara(strFactType, glngSys, lngModule, "0"))
    
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng����ID = Val(varTemp(0))
        If Trim(varTemp(1)) = "" Then lng����Ʊ�� = lng����ID    '���õ�.
        If Trim(varTemp(1)) = strUseType And lng����ID <> 0 Then
            lngʹ��Ʊ�� = lng����ID
        End If
    Next
    
    On Error GoTo errHandle
    '����˳��
    '1.��ʹ��
    '2.ʹ��������ֵ�
    '3.����ʹ������
    strSql = _
    "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
    "From Ʊ�����ü�¼" & vbNewLine & _
    "Where (ID =[1] or ID =[2]) And ʣ������ > 0   " & vbNewLine & _
    "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� Desc, ��ʼ����"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", lng����Ʊ��, lngʹ��Ʊ��)
    
    If Not rsTemp.EOF Then lngShareInvoiceID = Val(Nvl(rsTemp!ID))
    Set rsTemp = Nothing
    zlGetInvoiceShareID = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetInvoicePrintFormat(ByVal lngModule As Long, _
    ByVal bytBillType As gEM_BillType, _
    Optional ByVal strUseType As String = "", _
    Optional ByRef intOutPrintFormat As Integer, _
    Optional ByVal bytInExseKind As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵĴ�ӡ��ʽ
    '���:lngModule-ģ���
    '     strUseType-ʹ�����
    '     bytBillType-Ʊ��
    '����:intOutPrintFormat-����Ʊ�ݴ�ӡ��ʽ
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 10:22:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactTypeFormat  As String, intPrintFormat As Integer, intPrintFormat1 As Integer
    Dim varData As Variant, varTemp As Variant, i As Integer
    On Error GoTo errHandle
    
    strFactTypeFormat = "" & _
    Switch(bytBillType = EM_�Һ��վ�, "�Һŷ�Ʊ��ʽ", _
               bytBillType = EM_�����վ�, "סԺ���ʷ�Ʊ��ʽ", _
               bytBillType = EM_�շ��վ�, "�շѷ�Ʊ��ʽ", _
               bytBillType = EM_ҽ�ƿ��վ�, "ҽ�ƿ���Ʊ��ʽ", _
               bytBillType = EM_Ԥ���վ�, "Ԥ����Ʊ��ʽ", _
               bytBillType = EM_�˷��վ�, "�˷ѷ�Ʊ��ʽ", _
               bytBillType = EM_���ʺ�Ʊ, "���Ϸ�Ʊ��ʽ", _
               True, "")
               
    If bytBillType = EM_�����վ� Or bytBillType = EM_�շ��վ� Then
        If bytInExseKind = 1 Then
            strFactTypeFormat = "������ʷ�Ʊ��ʽ"
        End If
        If bytInExseKind = 2 Then
            strFactTypeFormat = "סԺ���ʷ�Ʊ��ʽ"
        End If
    End If
    
    intOutPrintFormat = 0
    '��ʼ��Ʊ��ʽ
    If bytBillType = EM_Ԥ���վ� Then
        'Ԥ�����޸�ʽ
       intOutPrintFormat = 0: zlGetInvoicePrintFormat = True
       Exit Function
    End If
    
    strFactTypeFormat = Trim(gobjDatabase.GetPara(strFactTypeFormat, glngSys, lngModule, ""))
    '��ʽ:ʹ�����1,��ʽ1|ʹ�����2,��ʽ2...
    varData = Split(strFactTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        intPrintFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intPrintFormat1 = intPrintFormat
        If Trim(varTemp(0)) = strUseType And intPrintFormat <> 0 Then
           intOutPrintFormat = intPrintFormat: Exit For
        End If
    Next
    If intOutPrintFormat = 0 And intPrintFormat1 <> 0 Then intOutPrintFormat = intPrintFormat1
    zlGetInvoicePrintFormat = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetInvoicePrintMode(ByVal lngModule As Long, _
    ByVal bytBillType As gEM_BillType, _
    Optional ByVal strUseType As String = "", _
    Optional ByRef intOutPrintMode As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵĴ�ӡ��ʽ
    '���:lngModule-ģ���
    '     strUseType-ʹ�����
    '     bytBillType-Ʊ��
    '����:intOutPrintMode-����Ʊ�ݴ�ӡ��ʽ
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 10:22:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFacePrintMode  As String, intPrintMode As Integer, intPrintMode1 As Integer
    Dim varData As Variant, varTemp As Variant, i As Integer
    On Error GoTo errHandle
    If lngModule = 1137 Then
        strFacePrintMode = "" & _
        Switch(bytBillType = EM_�Һ��վ�, "�Һŷ�Ʊ��ӡ��ʽ", _
                   bytBillType = EM_�����վ�, "���˽��ʴ�ӡ", _
                   bytBillType = EM_�շ��վ�, "���˽��ʴ�ӡ", _
                   bytBillType = EM_ҽ�ƿ��վ�, "ҽ�ƿ���Ʊ��ӡ��ʽ", _
                   bytBillType = EM_Ԥ���վ�, "Ԥ����Ʊ��ӡ��ʽ", _
                   bytBillType = EM_�˷��վ�, "���Ϸ�Ʊ��ӡ��ʽ", _
                   bytBillType = EM_���ʺ�Ʊ, "���Ϸ�Ʊ��ӡ��ʽ", _
                   True, "")
    Else
        strFacePrintMode = "" & _
        Switch(bytBillType = EM_�Һ��վ�, "�Һŷ�Ʊ��ӡ��ʽ", _
                   bytBillType = EM_�����վ�, "���˽��ʴ�ӡ", _
                   bytBillType = EM_�շ��վ�, "�շѷ�Ʊ��ӡ��ʽ", _
                   bytBillType = EM_ҽ�ƿ��վ�, "ҽ�ƿ���Ʊ��ӡ��ʽ", _
                   bytBillType = EM_Ԥ���վ�, "Ԥ����Ʊ��ӡ��ʽ", _
                   bytBillType = EM_�˷��վ�, "�˷ѷ�Ʊ��ӡ��ʽ", _
                   bytBillType = EM_���ʺ�Ʊ, "���Ϸ�Ʊ��ӡ��ʽ", _
                   True, "")
    End If
    
    intOutPrintMode = 0
    
    strFacePrintMode = Trim(gobjDatabase.GetPara(strFacePrintMode, glngSys, lngModule, ""))
    
    intOutPrintMode = -1
    '��ʽ:ʹ�����1,��ӡ��ʽ1|ʹ�����2,��ӡ��ʽ2...
    varData = Split(strFacePrintMode, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,", ",")
        intPrintMode = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
        If Trim(varTemp(0)) = strUseType Then
            intOutPrintMode = intPrintMode: Exit For
        End If
    Next
    If intOutPrintMode < 0 Then intOutPrintMode = intPrintMode1
    zlGetInvoicePrintMode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetInvoiceNoLength(ByVal bytBillType As gEM_BillType, Optional ByRef lngCardTypeID As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݳ���
    '����:����Ʊ�ݳ���
    '����:���˺�
    '����:2014-09-30 10:52:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim intLen As Long, i As Integer
    Dim varTemp As Variant
    
    intLen = 7
    If bytBillType = EM_ҽ�ƿ��վ� Then
        If lngCardTypeID <> 0 Then
            strSql = "select nvl(�Ƿ��ϸ����,0) as ���ų��� from ҽ�ƿ���� where ID=[1] "
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "ȡҽ�ƿ�Ʊ�ų���", lngCardTypeID)
            If Not rsTemp.EOF Then
                intLen = Val(Nvl(rsTemp!���ų���))
            End If
        End If
        zlGetInvoiceNoLength = intLen
        Exit Function
    End If

    i = Switch(bytBillType = EM_�Һ��վ�, 4, _
        bytBillType = EM_�����վ�, 3, _
        bytBillType = EM_�շ��վ�, 1, _
        bytBillType = EM_ҽ�ƿ��վ�, 0, _
        bytBillType = EM_Ԥ���վ�, 2, _
        True, 0)
    'λ���ֱ�:1-�շ�,2-Ԥ��,3-����,4-�Һ�
    varTemp = Split(gobjDatabase.GetPara(20, glngSys, , "7|7|7|7|7"), "|")
    If UBound(varTemp) >= i - 1 Then
        intLen = Val(varTemp(i - 1))
    End If
    zlGetInvoiceNoLength = intLen
End Function

Public Function zlGetInvoicePreperty(lngModule As Long, _
     ByVal bytBillType As gEM_BillType, _
     ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
     ByVal intInsure As Integer, _
     ByRef objFactProperty As clsFactProperty, _
     Optional ByRef lngCardTypeID As Long, _
     Optional ByVal bln����Ԥ�� As Boolean, _
     Optional ByVal bytInExseKind As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵ķ�Ʊ��Ϣ,�������Ϣ���ظ�ObjFaceProperty����
    '���:lngModule-ģ���
    '     lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     intInsure-����
    '     bytBillType-Ʊ��
    '     lngCardTypeID-�����ID(bytBillType=ҽ�ƿ�ʱ��Ч),��Ҫ��ȡ�ÿ��Ƿ��ϸ����Ʊ��
    '     bln����Ԥ��-Ԥ������
    '     bytInExseKind-��������,1-����;2-סԺ
    '����:objFactProperty-���ط�Ʊ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 11:17:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, i As Integer, strSql As String, rsTemp As ADODB.Recordset
    Dim strUseType As String, lngShareID As Long, intPrintMode As Integer, intPrintFormat As Integer
    
    
    Set objFactProperty = New clsFactProperty
    
    objFactProperty.Ʊ�� = IIf(bytBillType = EM_���ʺ�Ʊ, EM_�����վ�, IIf(bytBillType = EM_�˷��վ�, EM_�շ��վ�, bytBillType))
    On Error GoTo errHandle
    '��ȡʹ�����
    If bytBillType <> 2 Then
        Call zlGetInvoiceUserType(lngModule, objFactProperty.Ʊ��, lng����ID, lng��ҳId, intInsure, strUseType)
        objFactProperty.ʹ����� = strUseType
        objFactProperty.����ʹ����� = zlStartFactUseType(objFactProperty.Ʊ��)
    Else
        objFactProperty.ʹ����� = IIf(bln����Ԥ��, "1", "2")
        objFactProperty.����ʹ����� = True
    End If
    Call zlGetInvoiceShareID(lngModule, objFactProperty.Ʊ��, objFactProperty.ʹ�����, lngShareID)
    objFactProperty.��������ID = lngShareID
    
    Call zlGetInvoicePrintMode(lngModule, bytBillType, objFactProperty.ʹ�����, intPrintMode)
    objFactProperty.��ӡ��ʽ = intPrintMode
    Call zlGetInvoicePrintFormat(lngModule, bytBillType, objFactProperty.ʹ�����, intPrintFormat, bytInExseKind)
    objFactProperty.��ӡ��ʽ = intPrintFormat
    objFactProperty.Ʊ�ų��� = 7
    If bytBillType = EM_ҽ�ƿ��վ� Then
         objFactProperty.�ϸ���� = 0
        If lngCardTypeID <> 0 Then
            strSql = "select nvl(�Ƿ��ϸ����,0) as �ϸ����,���ų��� from ҽ�ƿ���� where ID=[1] "
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "ȡҽ�ƿ��Ƿ��ϸ����", lngCardTypeID)
            If Not rsTemp.EOF Then
                objFactProperty.�ϸ���� = Val(Nvl(rsTemp!�ϸ����)) <> 0
                objFactProperty.Ʊ�ų��� = Val(Nvl(rsTemp!���ų���))
            End If
        End If
    Else
        i = Switch(bytBillType = EM_�Һ��վ�, 4, _
               bytBillType = EM_�����վ�, 3, _
               bytBillType = EM_�շ��վ�, 1, _
               bytBillType = EM_ҽ�ƿ��վ�, 0, _
               bytBillType = EM_Ԥ���վ�, 2, _
               bytBillType = EM_�˷��վ�, 1, _
               bytBillType = EM_���ʺ�Ʊ, 3, _
               True, 0)
        'λ���ֱ�:1-�շ�,2-Ԥ��,3-����,4-�Һ�
        objFactProperty.�ϸ���� = Val(Mid(gobjDatabase.GetPara(24, glngSys, , "00000"), i, 1)) = 1
        objFactProperty.Ʊ�ų��� = zlGetInvoiceNoLength(objFactProperty.Ʊ��, lngCardTypeID)
    End If
    zlGetInvoicePreperty = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlCheckInvoiceIsUsed(lngModule As Long, _
       ByVal lng����ID As Long, ByVal strInvoiceNO As String, ByRef blnOutUsed As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ��ѱ�ʹ��
    '���:lngModule -ģ���
    '     lng����ID-����ID
    '     strInvoiceNo-���ķ�Ʊ��
    '����:blnOutUsed-true:��ʾʹ��;False-��ʾδʹ��
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 11:49:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    blnOutUsed = False
    
    strSql = "Select ���� From Ʊ��ʹ����ϸ Where ����ID=[1]  And ����=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��鷢Ʊ�Ƿ��ѱ�ʹ��", lng����ID, strInvoiceNO)
    blnOutUsed = Not rsTemp.EOF
    
    zlCheckInvoiceIsUsed = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetInvoiceGroupID(ByVal strUserName As String, ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "", _
    Optional lngOutUseID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ò���ָ��Ʊ��������÷�Χ�ڵ�����ID
    '���: bytKind      =   Ʊ��
    '      intNum       =   Ҫ��ӡ��Ʊ������
    '      lngLastUseID =   �ϴ�ʹ�õ�����ID
    '      lngShareUseID=   ���ز���ָ���Ĺ���ID
    '      strBill      =   ��ǰƱ�ݺţ����ڼ���������ε�Ʊ�ݷ�Χ
    '      strUseType-ʹ�����
    '      strUserName-����Ա����
    '����:lngOutUseID-��������ID
    '      >0   =   �ɹ������õ�����ID
    '      =0   =   ʧ��
    '      -1   =   û������(����򲻹�����δ����),δ���ù���
    '      -2   =   û������(����򲻹�����δ����),���õĹ���������򲻹�
    '      -3   =   ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
    '      -4   =   ָ�����ε�Ʊ�ݲ�����
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 12:03:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo errH
    '1.�ϴε����������Ƿ���ò�����
    If lngLastUseID > 0 Then
        strSql = "" & _
        "   Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
        "   From Ʊ�����ü�¼ " & _
        "   Where Ʊ��=[1] And ʣ������>=[2] And ID=[3]  " & _
        "           And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL) "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, intNum, lngLastUseID, IIf(Trim(strUseType) = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then    'Ŀǰ��Ʊ�ݺſ��ܺ��ϴβ�ͬ��������Ҫ��鷶Χ
                If strBill = "" Then lngOutUseID = lngLastUseID: GetInvoiceGroupID = True: Exit Function '����û�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then lngOutUseID = lngLastUseID: GetInvoiceGroupID = True: Exit Function
                
            ElseIf intNum > 1 Then  '����ȷ���������ε���ʱ,��ǰƱ�ݺ��������β�����
                lngOutUseID = -4: Exit Function
            End If
        End With
    End If
    
    '2.�ϴε��������β����û򲻿���ʱ,ȡ������Ĳ������õ�
    '  �ж��������ʹ�õ�����,�ٵ�����,��������
    strSql = "" & _
    "   Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����" & vbNewLine & _
    "   From Ʊ�����ü�¼" & vbNewLine & _
    "   Where Ʊ�� = [1] And ʣ������ >= [2] And ������ = [3]  " & _
    "           And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL ) " & _
    "           And ʹ�÷�ʽ = 1" & vbNewLine & _
    "   Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� desc, ��ʼ����"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, intNum, strUserName, IIf(strUseType = "", "LXH", strUseType))
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then lngOutUseID = Val(Nvl(!ID)):   GetInvoiceGroupID = True: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
            blnTmp = False
            strPre = "" & !ǰ׺�ı�
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                blnTmp = True
            End If
            If Not blnTmp Then lngOutUseID = Val(Nvl(!ID)):   GetInvoiceGroupID = True:  Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.û�����õ�,ʹ�ñ��ز���ָ���Ĺ�������
    If lngShareUseID > 0 Then
        strSql = "" & _
        "   Select ǰ׺�ı�,��ʼ����,��ֹ����" & vbNewLine & _
        "   From Ʊ�����ü�¼  " & _
        "   Where Ʊ��=[1] And ʣ������>=[2] And ID=[3] " & _
        "   And (Nvl(ʹ�����,'LXH')=[4] Or  ʹ����� Is NULL) "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, intNum, lngShareUseID, IIf(strUseType = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then lngOutUseID = lngShareUseID:   GetInvoiceGroupID = True:  Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
                blnTmp = False
                strPre = "" & !ǰ׺�ı�
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then lngOutUseID = lngShareUseID:   GetInvoiceGroupID = True:  Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    lngOutUseID = lngReturn
    GetInvoiceGroupID = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetInvoiceGroupID(ByVal lngModule As Long, _
    ByVal strUserName As String, ByVal bytBillType As gEM_BillType, ByVal strUserType As String, ByVal lngLastUserID As Long, _
    ByVal lngShareUseID As Long, ByRef lngOutUserID As Long, _
    Optional intNum As Integer = 1, Optional strCurInvoiceNo As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ������ID
    '���: lngModule -ģ���
    '    bytBillType -Ʊ��
    '    intNum -��Ҫ�ķ�Ʊ��
    '    lngLastUserID -�ϴ�ʹ�õ�����ID
    '    lngShareUseID -����Ʊ��ID
    '    strCurInvoiceNo -��ǰƱ��
    '    strUserType -ʹ�����
    '    strUserName-����Ա����
    '����:lngOutUserID-��������ID
    '     >0   =   �ɹ������õ�����ID
    '     =0   =   ��ȡʧ��
    '     -1   =   û������(����򲻹�����δ����),δ���ù���
    '     -2   =   û������(����򲻹�����δ����),���õĹ���������򲻹�
    '     -3   =   ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
    '     -4   =   ָ�����ε�Ʊ�ݲ�����
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 11:58:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetInvoiceGroupID = GetInvoiceGroupID(strUserName, bytBillType, intNum, lngLastUserID, lngShareUseID, strCurInvoiceNo, strUserType, lngOutUserID)
End Function

Public Function zlGetNextBill(ByVal lngModule As Long, ByVal lng����ID As Long, _
    ByRef strOutNextInvoiceNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������ID,��ȡ��һ��ʵ��Ʊ�ݺ�
    '���:lngModule-ģ���
    '     lng����ID-����ID
    '����:strOutNextInvoiceNo-һ���ŷ�Ʊ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-27 13:46:43
    '˵����1.��ȡ������Χ�ڵ���ЧƱ��ʱ,���ؿ����û�����
    '      2.�ſ��ѱ���ĺ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMain As ADODB.Recordset
    Dim rsDelete As ADODB.Recordset
    Dim strSql As String, strBill As String
    
    On Error GoTo errH
    strOutNextInvoiceNo = ""
    
    strSql = "Select ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ����" & _
        " From Ʊ�����ü�¼ Where ʣ������>0 And ID=[1]"
    Set rsMain = gobjDatabase.OpenSQLRecord(strSql, "ȡһ��Ʊ�ݺ�", lng����ID)
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!��ǰ����) Then
        strBill = UCase(rsMain!��ʼ����)
    Else
        strBill = UCase(IncStr(rsMain!��ǰ����))
    End If
    strSql = "Select Upper(����) as ���� From Ʊ��ʹ����ϸ" & _
        " Where ����||''>=[1] And ����ID=[2]" & _
        " Order by ����"
        
    Set rsDelete = gobjDatabase.OpenSQLRecord(strSql, "ȡһ��Ʊ�ݺ�", strBill, lng����ID)
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
        strBill = IncStr(strBill)
    Loop
    strOutNextInvoiceNo = strBill
    zlGetNextBill = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlStartFactUseType(ByVal bytBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ʹ����ʹ������
    '���:bytBillType-Ʊ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-10 16:11:47
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    strSql = "Select  1 as ���� From Ʊ�����ü�¼ where Ʊ��=[1] and nvl(ʹ�����,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���Ʊ���Ƿ�������ʹ������", bytBillType)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckInvoiceOverplusEnough(ByVal bytKind As Byte, _
    ByVal intNum As Integer, Optional lngʣ������ As Long, _
    Optional lng����ID As Long = 0, Optional strʹ����� As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ʊ�ݵ�ʣ�������Ƿ����
    '���:bytKind-1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '     intNum-��ǰ�Աȵ�����(-1����������)
    '     lng����ID-ֻ��鵱ǰ������Ʊ��(32455)
    '     strʹ�����-ʹ�����
    '����:lngʣ������-���ص�ǰʣ������
    '����:���㷵��true,���򷵻�False
    '����:���˺�
    '����:2009-12-28 17:16:16
    '����:26948
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    '-1����������
    If intNum = -1 Then zlCheckInvoiceOverplusEnough = True: Exit Function
    Err = 0: On Error GoTo ErrHand:
    
    lngʣ������ = 0
    
    strSql = "" & _
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
    strSql = "Select sum(ʣ������) as ʣ������ From (" & strSql & ")"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����Ʊ������", bytKind, UserInfo.����, lng����ID, strʹ�����)
    lngʣ������ = Val(Nvl(rsTemp!ʣ������))
    zlCheckInvoiceOverplusEnough = lngʣ������ > intNum
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then Resume
End Function

Public Function zlCheckBillNOIsPrintInvoice(ByVal bytƱ�� As gEM_BillType, ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ĵ����Ƿ��Ѿ����ڴ�ӡƱ�ݵ����
    '���:bytƱ��-1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '     strNOs-ָ��Ҫ�ش�ĵ��ݺţ������ţ������Ƕ�����ݺţ�Ϊ"'AAA','BBB',..."����ʽ
    '����:����Ʊ�ݷ���true,���򷵻�False
    '����:���˺�
    '����:2014-09-30 10:29:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "" & _
    "   Select /*+ rule */ Max(A.ID) as ID " & _
    "   From Ʊ�ݴ�ӡ���� A,Table( f_Str2list([1])) J " & _
    "   Where A.��������=[2]   And A.NO=J.Column_Value"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "mdlOutExse", Replace(strNos, "'", ""), bytƱ��)
    If rsTemp.RecordCount <> 0 Then
        zlCheckBillNOIsPrintInvoice = Val(Nvl(rsTemp!ID)) <> 0
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub