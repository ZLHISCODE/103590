Attribute VB_Name = "mdlEInvoice_BS"
Option Explicit
'*********************************************************************************************************************************************
'��˼����Ʊ����ش���
'һ������Ʊ�ݹ����ӿڴ���:
'   1.zlInitIFacePara :��ʼ����˼�ӿ�����
'   2.zlGetҵ���ʶ:�Ա�Hisҵ�񳡺��벩˼ҵ���ʶ
'��������Ʊ�ݽӿڵ���������ݴ���:
'   1.zlGetJson_CreateEInvoice:��˼���ߵ���Ʊ����Ҫ��Json����
'     1.1:zlGetJson_CreateEInvoiceByCharge:�շѵ���Ʊ��
'     1.2:zlGetJson_CreateEInvoiceByDeposit:Ԥ������Ʊ��
'     1.3:zlGetJson_CreateEInvoiceByMzBalance:������ʵ���Ʊ��
'     1.4:zlGetJson_CreateEInvoiceByZyBalance:סԺ���ʵ���Ʊ��
'     1.5:zlGetJson_CreateEInvoiceByRegsit:�Һŵ���Ʊ��
'     1.6:zlGetJson_CreateEInvoiceBySendCard:��������Ʊ��
'   3.zlGetJson_PrintEInvoice:��ȡ��ӡ����Ʊ��Json��ʽ����
'   4.zlGetJson_SendNotice:��ȡ������֪��Json��ʽ����
'   5.zlGetJson_CheckCancelEInvoice:Ʊ�����ϼ���Json
'   6.zlGetJson_CancelEInvoice:Ʊ������Json
'����ֽ��Ʊ����ؽӿ�
'   1.zlGetJson_GetNextInvoiceNo:��ȡֽ��Ʊ��Json��ʽ����
'   2.zlGetJson_TurnPaper:��ȡ����ֽ��Ʊ��Json��ʽ����
'   3.zlGetJson_TurnPaperPrint:��ȡ����ֽ��Ʊ�ݴ�ӡJson��ʽ����
'   4.zlGetJson_CancelPaper:��ȡֽ��Ʊ������Json��ʽ����

'Ŀǰ�漰��˼�Ľӿ�
'1.invoiceEBillOutpatient:�����շѵ���Ʊ��
'2.invEBillHospitalized:סԺ����Ʊ��
'1.getEBillAccountStatus
'2.invoicePayMentVoucher
'����:���ϴ�
'����:2020-03-03 14:11:42
'*********************************************************************************************************************************************
Public Enum BS_Version
    V2_0_3 = 0
    V3_1_0
    V3_2_0
End Enum

Private Type IFaceBs
    URL_Type                As String
    URL_Address             As String
    Ӧ���ʺ�                As String
    ǩ��˽Կ                As String
    ֧�ְ汾                As BS_Version
    ���ݴ��䷽ʽ            As String
    �ַ�����                As String
    ȱʡ�����ID            As Long
    ҽ�ƿ����ͱ��          As String
    ���֤�������ͱ��      As String
    �����޿��Ŀ������    As String
    �����޿��Ŀ���          As String
    ¼����ԭ��            As Boolean
    ���Ѷ��ձ���          As String
    ���Ѷ�������          As String
    ����ÿ�Ʊ              As Boolean
    �շ�ֽ��Ʊ�ݴ���        As String
    �Һ�ֽ��Ʊ�ݴ���        As String
    ����ֽ��Ʊ�ݴ���        As String
    Ԥ��ֽ��Ʊ�ݴ���        As String
End Type
Public gBs_Type As IFaceBs

Private mlngSys As Long
Private mstrOperatorCode As String
Private mstrOperatorName As String
Private mcllJsonKey As Collection
Private mcllJsonFormat As Collection


Public Function zlInitIFacePara(ByVal lngSys As Long, ByVal strOperatorCode As String, ByVal strOperatorName As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ��˼�ӿ�����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/21 15:35
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    mlngSys = lngSys
    mstrOperatorCode = strOperatorCode: mstrOperatorName = strOperatorName
    
    gBs_Type.URL_Type = ""
    gBs_Type.URL_Address = ""
    gBs_Type.Ӧ���ʺ� = ""
    gBs_Type.ǩ��˽Կ = ""
    gBs_Type.֧�ְ汾 = 0
    gBs_Type.���ݴ��䷽ʽ = ""
    gBs_Type.�ַ����� = ""
    gBs_Type.ȱʡ�����ID = 0
    gBs_Type.���֤�������ͱ�� = ""
    gBs_Type.�����޿��Ŀ������ = ""
    gBs_Type.�����޿��Ŀ��� = ""
    gBs_Type.ҽ�ƿ����ͱ�� = "" '��˼֤�����Ͷ���
    gBs_Type.���Ѷ��ձ��� = ""
    gBs_Type.���Ѷ������� = ""
    gBs_Type.����ÿ�Ʊ = False
    gBs_Type.�շ�ֽ��Ʊ�ݴ��� = ""
    gBs_Type.�Һ�ֽ��Ʊ�ݴ��� = ""
    gBs_Type.����ֽ��Ʊ�ݴ��� = ""
    gBs_Type.Ԥ��ֽ��Ʊ�ݴ��� = ""
    
    strSQL = "Select ������, ������, ����ֵ From �����ӿ����� Where �ӿ��� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlInitIFacePara", gobjEinvProvider.�ṩ��)
    With rsTmp
        Do While Not .EOF
            Select Case UCase(Nvl(!������))
                Case "URL_TYPE"
                    gBs_Type.URL_Type = Nvl(!����ֵ)
                Case "URL_ADDRESS"
                    gBs_Type.URL_Address = Nvl(!����ֵ)
                Case "Ӧ���ʺ�"
                    gBs_Type.Ӧ���ʺ� = Nvl(!����ֵ)
                Case "ǩ��˽Կ"
                    gBs_Type.ǩ��˽Կ = Nvl(!����ֵ)
                Case "֧�ְ汾"
                    If Nvl(!����ֵ) = "V3.2.0" Then
                        gBs_Type.֧�ְ汾 = BS_Version.V3_2_0
                    ElseIf Nvl(!����ֵ) = "V3.1.0" Then
                        gBs_Type.֧�ְ汾 = BS_Version.V3_1_0
                    Else
                        gBs_Type.֧�ְ汾 = BS_Version.V2_0_3
                    End If
                    
                Case "���ݴ��䷽ʽ"
                    gBs_Type.���ݴ��䷽ʽ = Nvl(!����ֵ)
                Case "�ַ�����"
                    gBs_Type.�ַ����� = Nvl(!����ֵ)
                Case "ȱʡ�����ID"
                    gBs_Type.ȱʡ�����ID = Val(Nvl(!����ֵ))
                Case "���֤�������ͱ��"
                    gBs_Type.���֤�������ͱ�� = Nvl(!����ֵ)
                Case "�����޿��Ŀ������"
                    gBs_Type.�����޿��Ŀ������ = Nvl(!����ֵ)
                Case "�����޿��Ŀ���"
                    gBs_Type.�����޿��Ŀ��� = Nvl(!����ֵ)
                Case "¼����ԭ��"
                    gBs_Type.¼����ԭ�� = Val(Nvl(!����ֵ)) = 1
                Case "ҽ�ƿ����ͱ��"
                    gBs_Type.ҽ�ƿ����ͱ�� = Nvl(!����ֵ)
                Case "���Ѷ��ձ���"
                    gBs_Type.���Ѷ��ձ��� = Nvl(!����ֵ)
                Case "���Ѷ�������"
                    gBs_Type.���Ѷ������� = Nvl(!����ֵ)
                Case "����ÿ��ߵ���Ʊ��"
                    gBs_Type.����ÿ�Ʊ = Val(Nvl(!����ֵ)) = 1
                Case "�շ�ֽ��Ʊ�ݴ���"
                    gBs_Type.�շ�ֽ��Ʊ�ݴ��� = Nvl(!����ֵ)
                Case "�Һ�ֽ��Ʊ�ݴ���"
                    gBs_Type.�Һ�ֽ��Ʊ�ݴ��� = Nvl(!����ֵ)
                Case "����ֽ��Ʊ�ݴ���"
                    gBs_Type.����ֽ��Ʊ�ݴ��� = Nvl(!����ֵ)
                Case "Ԥ��ֽ��Ʊ�ݴ���"
                    gBs_Type.����ֽ��Ʊ�ݴ��� = Nvl(!����ֵ)
            End Select
            .MoveNext
        Loop
    End With
    
    Call InitVersionDiff
    zlInitIFacePara = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitVersionDiff() As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���ð汾����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/6/2 13:51
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    '1.�ڵ����
    Set mcllJsonKey = New Collection
    If gBs_Type.֧�ְ汾 > V3_1_0 Then
        Call mcllJsonKey.Add("patientCategory", "_�������") '�շѡ��Һ�
        Call mcllJsonKey.Add("patientCategoryCode", "_������ұ���") '�Һ�
    Else
        Call mcllJsonKey.Add("category", "_�������")
        Call mcllJsonKey.Add("patientCategory", "_������ұ���")
    End If
    
    '2.���ݸ�ʽ����
    Set mcllJsonFormat = New Collection
    If gBs_Type.֧�ְ汾 > V3_1_0 Then
        Call mcllJsonFormat.Add(4, "_����С��") '�շѡ��Һš�����
        Call mcllJsonFormat.Add(2, "_����С��") '�շѡ��Һš�����
        Call mcllJsonFormat.Add("yyyyMMdd", "_��������") '��������Ժ���ڣ� �շѡ��Һš�����
    Else
        Call mcllJsonFormat.Add(6, "_����С��")
        Call mcllJsonFormat.Add(6, "_����С��")
        Call mcllJsonFormat.Add("yyyy-MM-dd", "_��������")
    End If
    
    '3.���ݲ���
    '�˴��޷�����
    
    InitVersionDiff = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetVersionDiff(ByVal bytType As Byte, ByVal strKey As String) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�汾����
    ' ��� : bytType:1-�ڵ���죻2-�ڵ��ʽ����
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/6/2 13:51
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    
    If bytType = 1 Then
        GetVersionDiff = mcllJsonKey("_" & strKey)
    Else
        GetVersionDiff = mcllJsonFormat("_" & strKey)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlGetҵ���ʶ(ByVal byt���� As Byte) As String
    'ҵ���ʶ:01  סԺ,02  ����, 03  ����, 04  ����, 05  �������, 06  �Һ�, 07  סԺԤ����, 08  ���Ԥ����
    zlGetҵ���ʶ = Decode(byt����, 1, "02", 2, "07", 3, "01", 4, "06", 5, "02", "02")
End Function

Public Function zlGetJson_CreateEInvoice(ByVal bytInvoiceType As Byte, ByVal lngEInvoiceID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                ByVal strEInvoiceClientCode As String, strServiceCode As String, dblƱ�ݽ��_Out As Double, strJson_Out As String, _
                Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���߷�ƱJson��ʽ����
    ' ��� : bytInvoiceType-���ó���
    '        lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    '        strEInvoiceClientCode-��Ʊ����
    ' ���� : strJson_Out-��Ʊ��Ϣ
    '        strServiceCode-�����ʶ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    Select Case bytInvoiceType
        Case 1
            If Not zlGetJson_CreateEInvoiceByCharge(lngEInvoiceID, lng����ID, lng����ID, strEInvoiceClientCode, dblƱ�ݽ��_Out, strJson_Out, strErrMsg_Out) Then Exit Function
            strServiceCode = "invoiceEBillOutpatient"
        Case 2
            If Not zlGetJson_CreateEInvoiceByDeposit(lngEInvoiceID, lng����ID, lng����ID, strEInvoiceClientCode, dblƱ�ݽ��_Out, strJson_Out, strErrMsg_Out) Then Exit Function
            If lng����ID <> 0 Then
                strServiceCode = "writeOffPayMentVoucher"
            Else
                strServiceCode = "invoicePayMentVoucher"
            End If
        Case 3
            strSQL = "Select Max(��������) As �������� From ���˽��ʼ�¼ Where ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�жϽ�������", lng����ID)
            If rsTmp.RecordCount = 0 Then
                strErrMsg_Out = "δ�ҵ��������ݣ����ܴ�ӡ����Ʊ��": Exit Function
            End If
            If Val(Nvl(rsTmp!��������)) = 1 Then
                If Not zlGetJson_CreateEInvoiceByMzBalance(lngEInvoiceID, lng����ID, lng����ID, strEInvoiceClientCode, dblƱ�ݽ��_Out, strJson_Out, strErrMsg_Out) Then Exit Function
                strServiceCode = "invoiceEBillOutpatient"
            Else
                If Not zlGetJson_CreateEInvoiceByZyBalance(lngEInvoiceID, lng����ID, lng����ID, strEInvoiceClientCode, dblƱ�ݽ��_Out, strJson_Out, strErrMsg_Out) Then Exit Function
                strServiceCode = "invEBillHospitalized"
            End If
        Case 4
            If Not zlGetJson_CreateEInvoiceByRegsit(lngEInvoiceID, lng����ID, lng����ID, strEInvoiceClientCode, dblƱ�ݽ��_Out, strJson_Out, strErrMsg_Out) Then Exit Function
            strServiceCode = "invEBillRegistration"
        Case 5
            If Not zlGetJson_CreateEInvoiceBySendCard(lngEInvoiceID, lng����ID, lng����ID, strEInvoiceClientCode, dblƱ�ݽ��_Out, strJson_Out, strErrMsg_Out) Then Exit Function
            strServiceCode = "invoiceEBillOutpatient"
        Case Else
            strErrMsg_Out = "��Ч��Ӧ�ó���": Exit Function
    End Select
    zlGetJson_CreateEInvoice = True
End Function

Public Function zlGetJson_PrintEInvoice(ByVal lngEInvoiceID As Long, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ��ӡ����Ʊ��Json��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    ' ���� : strJson_Out-��Ʊ��Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strJson_Out = ""
    
    Set rsTmp = GetEInvoiceInfo(lngEInvoiceID, strErrMsg_Out)
    If rsTmp Is Nothing Then Exit Function
    
    With rsTmp
        strJson_Out = GetJsonNodeString("billBatchCode", Nvl(!Ʊ�ݴ���), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", Nvl(!Ʊ�ݺ���), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("random", Nvl(!Ʊ��У����), Json_Text)
        strJson_Out = "{" & strJson_Out & "}"
    End With
    zlGetJson_PrintEInvoice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_SendNotice(ByVal lngEInvoiceID As Long, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������֪��Json��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    ' ���� : strJson_Out-��Ʊ��Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strJsonList As String, strJson As String
    On Error GoTo ErrHand
    strJson_Out = ""
    
    Set rsTmp = GetEInvoiceWithPatiInfo(lngEInvoiceID, strErrMsg_Out)
    If rsTmp Is Nothing Then Exit Function
    
    
    With rsTmp
        strJson_Out = GetJsonNodeString("billBatchCode", Nvl(!Ʊ�ݴ���), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", Nvl(!Ʊ�ݺ���), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("random", Nvl(!Ʊ��У����), Json_Text)
        
        If Nvl(!�ֻ���) <> "" Then
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("noticeType", 1201, Json_Text)
            strJson = strJson & "," & GetJsonNodeString("noticeValue", Nvl(!�ֻ���), Json_Text)
            strJsonList = ",{" & strJson & "}"
        End If
        
        If Nvl(!email) <> "" Then
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("noticeType", 1202, Json_Text)
            strJson = strJson & "," & GetJsonNodeString("noticeValue", Nvl(!email), Json_Text)
            strJsonList = strJsonList & ",{" & strJson & "}"
        End If
        If strJsonList = "" Then Exit Function 'û����Ϣ����;����ֱ���˳�
        
        strJson_Out = strJson_Out & "," & GetNodeString("noticeList") & ":[" & Mid(strJsonList, 2) & "]"
        strJson_Out = "{" & strJson_Out & "}"
    End With
    zlGetJson_SendNotice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_GetNextInvoiceNo(ByVal bytInvoiceType As Byte, ByVal strEInvoiceNodeCode As String, _
                strPaperCode_Out As String, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡֽ��Ʊ��Json��ʽ����
    ' ��� : bytInvoiceType-Ʊ��
    '        strEInvoiceNodeCode-��Ʊ����
    ' ���� : strJson_Out-��Ʊ��Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strJsonList As String, strJson As String
    On Error GoTo ErrHand
    strPaperCode_Out = GetPaperCode(bytInvoiceType)
    
    strJson_Out = ""
    strJson_Out = GetJsonNodeString("placeCode", strEInvoiceNodeCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillBatchCode", strPaperCode_Out, Json_Text)
    '����������Json
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_GetNextInvoiceNo = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_TurnPaper(ByVal bytInvoiceType As Byte, ByVal strEInvoiceNodeCode As String, _
    ByVal strInvoiceNO As String, ByVal lngEInvoiceID As Long, ByVal strEInvoiceCode As String, ByVal strEInvoiceNO As String, _
    ByVal strCreateTime As String, ByVal strOperatorCode As String, ByVal strOperatorName As String, _
    strServiceCode As String, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ����ֽ��Ʊ��Json��ʽ����
    ' ��� : bytInvoiceType-Ʊ��
    '        strEInvoiceNodeCode-��Ʊ����
    '        bytInvoiceType-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '        strInvoiceNO-��Ʊ��
    '        lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '        strEInvoiceCode-����Ʊ�ݴ���
    '        strEInvoiceNO-����Ʊ�ݺ���
    '        strCreateTime-����Ʊ������ʱ��,��ʽ:YYYYMMDDhhmmssSSS
    '        strOperatorCode-����Ա���
    '        strOperatorName-����Ա����
    ' ���� : strJson_Out-��Ʊ��Ϣ
    '        strServiceCode-ҵ���ʶ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strJsonList As String, strJson As String, strPaperCode As String
    On Error GoTo ErrHand
    Set rsTmp = GetEInvoiceWithPatiInfo(lngEInvoiceID, strErrMsg_Out)
    If rsTmp Is Nothing Then Exit Function
    If Val(Nvl(rsTmp!�Ƿ񻻿�)) = 1 Then
        strServiceCode = "reTurnPaper"
    Else
        strServiceCode = "turnPaper"
    End If
    strPaperCode = GetPaperCode(bytInvoiceType)
    strInvoiceNO = Mid(strInvoiceNO, Len(strPaperCode) + 1)
    
    strJson_Out = ""
    strJson_Out = GetJsonNodeString("billBatchCode", strEInvoiceCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", strEInvoiceNO, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillBatchCode", strPaperCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNo", strInvoiceNO, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("busDateTime", strCreateTime, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("placeCode", strEInvoiceNodeCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("operator", strOperatorName, Json_Text)
    '����������Json
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_TurnPaper = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_TurnPaperPrint(ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, _
                strServiceCode As String, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ����ֽ��Ʊ�ݴ�ӡJson��ʽ����
    ' ��� : bytInvoiceType-Ʊ��
    '        strEInvoiceNodeCode-��Ʊ����
    '        bytInvoiceType-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '        strInvoiceNO-��Ʊ��
    '        lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '        strEInvoiceCode-����Ʊ�ݴ���
    '        strEInvoiceNO-����Ʊ�ݺ���
    '        strCreateTime-����Ʊ������ʱ��,��ʽ:YYYYMMDDhhmmssSSS
    '        strOperatorCode-����Ա���
    '        strOperatorName-����Ա����
    ' ���� : strJson_Out-��Ʊ��Ϣ
    '        strServiceCode-ҵ���ʶ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strPaperCode As String
    On Error GoTo ErrHand
    strServiceCode = IIf(bytInvoiceType = 2, "", "printPaperBill")
    
    strPaperCode = GetPaperCode(bytInvoiceType)
    strInvoiceNO = Mid(strInvoiceNO, Len(strPaperCode) + 1)
    
    strJson_Out = ""
    strJson_Out = GetJsonNodeString("pBillBatchCode", strPaperCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNo", strInvoiceNO, Json_Text)
    '����������Json
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_TurnPaperPrint = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_CheckCancelEInvoice(ByVal lngEInvoiceID As Long, strEInvoiceNo_Out As String, strJson_Out As String, Optional strErrMsg_Out As String, _
                Optional ByVal blnCheckAcc As Boolean, Optional strJsonAcc_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ��췢Ʊ���Json��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    ' ���� : strEInvoiceNo_Out-����Ʊ�ݺ�
    '        strJson_Out-Ʊ����Ϣ
    '        strJsonAcc_Out-Ʊ��������Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strJson_Out = ""
    
    Set rsTmp = GetEInvoiceInfo(lngEInvoiceID, strErrMsg_Out)
    If rsTmp Is Nothing Then Exit Function
    With rsTmp
        strEInvoiceNo_Out = Nvl(!Ʊ�ݺ���)
        strJson_Out = GetJsonNodeString("billBatchCode", Nvl(!Ʊ�ݴ���), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", strEInvoiceNo_Out, Json_Text)
        
        If blnCheckAcc Then
            strJsonAcc_Out = strJson_Out
            strJsonAcc_Out = strJsonAcc_Out & "," & GetJsonNodeString("random", Nvl(!Ʊ��У����), Json_Text)
            strJsonAcc_Out = strJsonAcc_Out & "," & GetJsonNodeString("createTime", Nvl(!����ʱ��), Json_Text)
        End If
    End With
    '����������Json
    strJson_Out = "{" & strJson_Out & "}"
    If blnCheckAcc Then strJsonAcc_Out = "{" & strJsonAcc_Out & "}"
    zlGetJson_CheckCancelEInvoice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_CancelEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, ByVal strEInvoiceClientCode As String, ByVal blnNoInputReason As Boolean, _
                strServiceCode As String, strJson_Out As String, strReason_Out As String, Optional strErrMsg_Out As String, _
                Optional ByVal strEInvoiceType As String, Optional ByVal strEInvoiceCode As String, Optional ByVal strEInvoiceNO As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ��췢ƱJson��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    '        strEInvoiceClientCode-��Ʊ����
    '        blnNoInputReason-�Ƿ���������ԭ�򣬿�������ֽ��Ʊ��ʱ��¼��
    '        strEInvoiceType-ҵ�񳡺�
    '        strEInvoiceCode-����Ʊ�ݴ���
    '        strEInvoiceNo-����Ʊ�ݺ���
    ' ���� : strJson_Out-��Ʊ��Ϣ
    '        strҵ���ʶ-ҵ���ʶ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim bytƱ�� As Byte
    On Error GoTo ErrHand
    strJson_Out = ""
    
    If strEInvoiceNO = "" Then
        Set rsTmp = GetEInvoiceInfo(lngEInvoiceID, strErrMsg_Out)
        If rsTmp Is Nothing Then Exit Function
        With rsTmp
            bytƱ�� = Val(Nvl(!Ʊ��))
            strEInvoiceNO = Nvl(!Ʊ�ݺ���)
            strEInvoiceCode = Nvl(!Ʊ�ݴ���)
        End With
    Else
        bytƱ�� = Decode(strEInvoiceType, "02", 1, "07", 2, "01", 3, "06", 4, "02")
    End If
    
    With rsTmp
        If Not blnNoInputReason Then
            If gBs_Type.¼����ԭ�� And mlngSys <> 2600 Then
                If frmInputBox.InputBox(frmMain, "Ʊ�ݳ��", "��¼��Ʊ�ݳ���ԭ��", 30, 1, False, False, strReason_Out) = False Then Exit Function
            Else
                strReason_Out = Decode(bytƱ��, 2, "��Ԥ��", 3, "��������", 4, "�˺�", 5, "�˿�", "�˷�")
            End If
        End If
        strServiceCode = IIf(bytƱ�� = 2, "cancelPayMentVoucherBalance", "writeOffEBill")
        
        strJson_Out = GetJsonNodeString("billBatchCode", strEInvoiceCode, Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", strEInvoiceNO, Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("reason", strReason_Out, Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("operator", mstrOperatorName, Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("busDateTime", Format(zlDatabase.Currentdate, "YYYYMMDDhhmmss000"), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
        strJson_Out = "{" & strJson_Out & "}"
    End With
    zlGetJson_CancelEInvoice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_CancelPaper(ByVal frmMain As Object, ByVal bytInvoiceType As String, ByVal strInvoiceNO As String, ByVal strEInvoiceClientCode As String, ByVal strOperatorName As String, _
                strServiceCode As String, strBusDateTime As String, strJson_Out As String, strReason_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡƱ������Json��ʽ����
    ' ��� : bytInvoiceType -Ʊ�֣��ݴ�ֽ��Ʊ�ݴ���
    '        strInvoiceNO-ֽ��Ʊ�ݺ�
    '        lngEInvoiceID - ����Ʊ��ID
    '        strEInvoiceClientCode-��Ʊ����
    ' ���� : strJson_Out-��Ʊ��Ϣ
    '        strServiceCode-ҵ���ʶ
    '        strBusDateTime-����ʱ��
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strCurCode As String, strPaperCode As String, strPaperNo As String
    
    On Error GoTo ErrHand
    strJson_Out = ""
    strCurCode = GetPaperCode(bytInvoiceType)
    strPaperCode = Left(strInvoiceNO, Len(strCurCode))
    strPaperNo = Mid(strInvoiceNO, Len(strCurCode) + 1)
    
    If gBs_Type.¼����ԭ�� And mlngSys <> 2600 Then
        If frmInputBox.InputBox(frmMain, "Ʊ�ݳ��", "��¼��Ʊ�ݳ���ԭ��", 30, 1, False, False, strReason_Out) = False Then Exit Function
    Else
        strReason_Out = Decode(bytInvoiceType, 2, "��Ԥ��", 3, "��������", 4, "�˺�", 5, "�˿�", "�˷�")
    End If
    strServiceCode = IIf(bytInvoiceType = 2, "invalidPayMentVoucherPaper", "invalidPaper")
    
    strJson_Out = GetJsonNodeString("pBillBatchCode", strPaperCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNo", strPaperNo, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("author", strOperatorName, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("reason", strReason_Out, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("busDateTime", Format(zlDatabase.Currentdate, "YYYYMMDDhhmmss000"), Json_Text)
    '����������Json
    strJson_Out = "{" & strJson_Out & "}"
    
    zlGetJson_CancelPaper = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_CancelBlankInvoice(ByVal strBatchNo As String, ByVal strStartInvoice As String, _
                    ByVal strEndInvoice As String, ByVal strEInvoiceClientCode As String, _
                    ByVal strAuthorName As String, ByVal strReason As String, ByVal strHappenTime As String, _
                    strJson_Out As String, strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡƱ�ݱ���Json��ʽ����
    ' ��� : strBatchNo-���Σ��ݴ�ֽ��Ʊ�ݴ���
    '        strStartInvoice-��ʼֽ��Ʊ�ݺ�
    '        strEndInvoice-��ֹֽ��Ʊ�ݺ�
    '        strAuthorName-������
    '        strReason-����ԭ��
    '        strEInvoiceClientCode-��Ʊ����
    '        strHappenTime-����ʱ��
    ' ���� : strJson_Out-Ʊ�ݱ�����Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    strJson_Out = ""
    strJson_Out = GetJsonNodeString("pBillBatchCode", strBatchNo, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNoStart", strStartInvoice, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNoEnd", strEndInvoice, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("author", strAuthorName, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("reason", strReason, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("busDateTime", Format(strHappenTime, "YYYYMMDDhhmmss000"), Json_Text)
    '����������Json
    strJson_Out = "{" & strJson_Out & "}"
    
    zlGetJson_CancelBlankInvoice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByRegsit(ByVal lngEInvoiceID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                ByVal strEInvoiceClientCode As String, dblƱ���ܽ�� As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�Һŷ�ƱJson��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    '        strEInvoiceClientCode-��Ʊ����
    ' ���� : strJson-�ҺŽ�����Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bytInvoiceType As Byte
    Dim dbl���� As Double
    Dim lng����ID As Long, lng�Һ�ID As Long
    Dim strҵ���ʶ As String, strChargeDetail As String, strListDetail As String
    Dim str����IDs As String, str�Ǽ�ʱ�� As String, strҵ�����Ա As String
    Dim str�������� As String, str�����Ա� As String, str�������� As String, strҽ�Ƹ��ʽ���� As String
    Dim strJsonList As String, strData As String, strValue As String
    Dim strJsonKey_������� As String, strJsonKey_������ұ���
    Dim strJsonFormat_�������� As String
    Dim intJsonFormat_����С�� As Integer, intJsonFormat_����С�� As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 4
    strҵ���ʶ = zlGetҵ���ʶ(bytInvoiceType)
    dblƱ���ܽ�� = 0
    
    '�汾����
    strJsonKey_������� = GetVersionDiff(1, "�������")
    strJsonKey_������ұ��� = GetVersionDiff(1, "������ұ���")
    strJsonFormat_�������� = GetVersionDiff(2, "��������")
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    
    strSQL = "Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ," & vbNewLine & _
            "          Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
            "          Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��," & vbNewLine & _
            "          Max(s.�������) As ҽ����Ŀ����, Max(s.��������) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע," & vbNewLine & _
            "          Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����," & vbNewLine & _
            "          Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��," & vbNewLine & _
            "          Max(a.���ʽ) As ���ʽ����, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����," & vbNewLine & _
            "          Max(B1.Id) As �Һ�id, Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����," & vbNewLine & _
            "          Max(b.���) As ���, Max(q.ҩƷ����) As ҩƷ����" & vbNewLine & _
            "   From ������ü�¼ A, ���˹Һż�¼ B1, �շ���ĿĿ¼ B, �վݷ�Ŀ���� C, �վݷ�Ŀ C1, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S" & vbNewLine & _
            "   Where a.No = B1.No And a.No In (Select Distinct NO From ������ü�¼ Where ����id = [1]) And a.��¼���� = 4 And a.��¼״̬ = 1 And" & vbNewLine & _
            "         a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And a.�վݷ�Ŀ = c1.����(+) And a.�վݷ�Ŀ = c.�վݷ�Ŀ(+) and Decode(c.���ó���(+), 0, 1, c.���ó���(+)) = 1 And a.�շ�ϸĿid = m.ҩƷid(+) And" & vbNewLine & _
            "         m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And t.����(+) = 1 And" & vbNewLine & _
            "         a.���մ���id = s.���մ���id(+)" & vbNewLine & _
            "   Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����" & vbNewLine & _
            IIf(gBs_Type.����ÿ�Ʊ, "", " Having Sum(a.���ʽ��) <> 0") & vbNewLine & _
            "   Order By NO, ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByRegsit", lng����ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ���ϸ���ݣ����ܴ�ӡ����Ʊ��"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        str�������� = Nvl(!����)
        str�����Ա� = Nvl(!�Ա�)
        str�������� = Nvl(!����)
        lng����ID = Val(Nvl(!����ID))
        strҽ�Ƹ��ʽ���� = Nvl(!���ʽ����)
        lng�Һ�ID = Val(Nvl(!�Һ�id))
        str�Ǽ�ʱ�� = Format(Nvl(!�Ǽ�ʱ��), "YYYYMMDDhhmmss000")
        strҵ�����Ա = Nvl(!����Ա����)
        
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!����ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!�վݷ�Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!�վݷ�Ŀ), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!������), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!�������), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!ҩƷ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!���), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!���㵥λ), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!�۸�)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!����)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!ʵ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!�Էѽ��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!Ӧ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!ҽ����������)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!��ע), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!���), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!��������), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dblƱ���ܽ�� = dblƱ���ܽ�� + RoundEx(Nvl(!ʵ�ս��), 6)
            .MoveNext
        Loop
        
        str����IDs = GetBalanceIDs(lng����ID, bytInvoiceType)
        dbl���� = GetBalanceErrorFee(str����IDs)
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
    End With
    
    '������ϸ
    If gBs_Type.���Ѷ��ձ��� <> "" Then
        dblƱ���ܽ�� = dblƱ���ܽ�� - dbl����
    End If
    dblƱ���ܽ�� = RoundEx(dblƱ���ܽ��, 2)
    If Not Get������ϸ(str����IDs, strData, dblƱ���ܽ��, bytInvoiceType, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    'Ʊ����Ϣ
    'ҵ����ˮ��:lngEInvoiceID_lng����ID
    strData = ""
    strData = strData & "" & GetJsonNodeString("busNo", lng����ID & "_" & lngEInvoiceID, Json_Text)
    strData = strData & "," & GetJsonNodeString("busType", strҵ���ʶ, Json_Text)
    strData = strData & "," & GetJsonNodeString("payer", str��������, Json_Text)
    strData = strData & "," & GetJsonNodeString("busDateTime", str�Ǽ�ʱ��, Json_Text)
    strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strData = strData & "," & GetJsonNodeString("payee", strҵ�����Ա, Json_Text)
    strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("totalAmt", dblƱ���ܽ��, Json_num)
    strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl����, 6) <> 0 And gBs_Type.���Ѷ��ձ��� = "", "����" & FormatEx(dbl����, 6) & "�����������", ""), Json_Text)
    strJson_Out = strData
    
    
    '�ƶ�֧��
    If Not Get�ƶ�֧����Ϣ(lng����ID, lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '֪ͨ��Ϣ
    If Not Get֪ͨ��Ϣ(lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '������Ϣ
    Call Getҽ����Ϣ(bytInvoiceType, lng����ID, lng����ID, cllInsureInfo)
    strSQL = "Select To_Char(a.����ʱ��, 'yyyy-mm-dd') As ��������, b.���� As ������ұ���," & vbNewLine & _
            "       b.���� As �����������, a.No As ������" & vbNewLine & _
            "  From ���˹Һż�¼ A, ���ű� B" & vbNewLine & _
            "  Where a.ִ�в���id = b.Id And a.Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByRegsit", lng�Һ�ID)
    
    strData = ""
    strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("ҽ�ƻ�������"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_���ջ�������", "C"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareTypeCode", strҽ�Ƹ��ʽ����, Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalCareType", Getҽ�Ƹ��ʽ����(strҽ�Ƹ��ʽ����), Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_ҽ����", "C"), Json_Text)
    With rsTmp
        If .RecordCount > 0 Then
            strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!��������), strJsonFormat_��������), Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_�������, Nvl(!�����������), Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_������ұ���, Nvl(!������ұ���), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!������), Json_Text)
        Else
            strData = strData & "," & GetJsonNodeString("consultationDate", "", Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_�������, "", Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_������ұ���, "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", lng����ID, Json_Text)
        End If
    End With
    strData = strData & "," & GetJsonNodeString("patientId", lng����ID, Json_Text)
    strData = strData & "," & GetJsonNodeString("sex", str�����Ա�, Json_Text)
    strData = strData & "," & GetJsonNodeString("age", str��������, Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '֧����Ϣ
    If Not Get������Ϣ(str����IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '�ɷ�����
    If Not Get�ɷ�����(str����IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '����ҽ����Ϣ-����
    '������չ��Ϣ-����
    'eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    'isArrears  �Ƿ����ͨ  String  1  ��  0-��1-�ǣ���Ƿ���������ҽԺҵ��Ҫ���Ʊ���Ƿ����ͨ��
    'arrearsReason  ������ͨԭ��  String  200  ��  isArrears=0����д������ͨ��ԭ��
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strData = strData & "," & GetJsonNodeString("isArrears", "1", Json_Text)
    strData = strData & "," & GetJsonNodeString("arrearsReason", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '�շ���Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strChargeDetail
    '�嵥��Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strListDetail
    
    '����������Json��
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByRegsit = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByCharge(ByVal lngEInvoiceID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                ByVal strEInvoiceClientCode As String, dblƱ���ܽ�� As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�Һŷ�ƱJson��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    '        strEInvoiceClientCode-��Ʊ����
    ' ���� : strJson-�ҺŽ�����Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String, rsTmp As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bln������ As Boolean
    Dim bytInvoiceType As Byte
    Dim dbl���� As Double
    Dim lng����ID As Long, lng�Һ�ID As Long, lngҽ����� As Long
    Dim strҵ���ʶ As String, strChargeDetail As String, strListDetail As String
    Dim str����IDs As String, str�Ǽ�ʱ�� As String, strҵ�����Ա As String
    Dim str�������� As String, str�����Ա� As String, str�������� As String, strҽ�Ƹ��ʽ���� As String
    Dim str����� As String
    Dim strJsonList As String, strData As String
    Dim strJsonKey_������� As String
    Dim strJsonFormat_�������� As String
    Dim intJsonFormat_����С�� As Integer, intJsonFormat_����С�� As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 1
    strҵ���ʶ = zlGetҵ���ʶ(bytInvoiceType)
    bln������ = CheckBillExistReplenishData(lng����ID)
    dblƱ���ܽ�� = 0
    
    '�汾����
    strJsonKey_������� = GetVersionDiff(1, "�������")
    strJsonFormat_�������� = GetVersionDiff(2, "��������")
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    
    If bln������ Then
        strWhere = "Select Distinct a.NO From ������ü�¼ a, ���ò����¼ b Where a.����ID = b.�շѽ���ID and b.����id = [1]"
    Else
        strWhere = "Select Distinct NO From ������ü�¼ Where ����id = [1]"
    End If
    strSQL = "Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ," & vbNewLine & _
            "        Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
            "        Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��," & vbNewLine & _
            "        Max(s.�������) As ҽ����Ŀ����, Max(s.��������) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע," & vbNewLine & _
            "        Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����," & vbNewLine & _
            "        Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��," & vbNewLine & _
            "        Max(a.���ʽ) As ���ʽ����, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����," & vbNewLine & _
            "        Max(a.�Һ�id) As �Һ�id, Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����," & vbNewLine & _
            "        Max(b.���) As ���, Max(q.ҩƷ����) As ҩƷ����" & vbNewLine & _
            " From ������ü�¼ A, �շ���ĿĿ¼ B, �վݷ�Ŀ���� C, �վݷ�Ŀ C1, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S" & vbNewLine & _
            " Where a.No In (" & strWhere & ") And Mod(a.��¼����, 10) = 1 And" & vbNewLine & _
            "       a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And a.�վݷ�Ŀ = c1.����(+) And a.�վݷ�Ŀ = c.�վݷ�Ŀ(+) and Decode(c.���ó���(+), 0, 1, c.���ó���(+)) = 1 And a.�շ�ϸĿid = m.ҩƷid(+) And" & vbNewLine & _
            "       m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And t.����(+) = 1 And" & vbNewLine & _
            "       a.���մ���id = s.���մ���id(+)" & vbNewLine & _
            " Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����" & vbNewLine & _
            " Order By NO, ���"
    strSQL = "Select Min(a.����ID) As ����ID, a.No, a.���, a.�շ�ϸĿid, a.���㵥λ, a.�۸�, Sum(a.����) As ����," & vbNewLine & _
            "       Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.�Էѽ��) As �Էѽ��," & vbNewLine & _
            "       Max(a.ҽ����Ŀ����) As ҽ����Ŀ����, Max(a.ҽ����Ŀ����) As ҽ����Ŀ����, Max(a.ҽ����������) As ҽ����������, Max(a.��ע) As ��ע," & vbNewLine & _
            "       Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����," & vbNewLine & _
            "       Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��," & vbNewLine & _
            "       Max(a.���ʽ����) As ���ʽ����, Max(a.�վݷ�Ŀ) As �վݷ�Ŀ, Max(a.�վݷ�Ŀ����) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����," & vbNewLine & _
            "       Max(a.�Һ�id) As �Һ�id, Max(a.������) As ������, Max(a.�������) As �������, Max(a.��Ŀ����) As ��Ŀ����, Max(a.��Ŀ����) As ��Ŀ����," & vbNewLine & _
            "       Max(a.���) As ���, Max(a.ҩƷ����) As ҩƷ����" & vbNewLine & _
            "From (" & strSQL & ") a" & vbNewLine & _
            "Group By a.No, a.���, a.�շ�ϸĿid, a.���㵥λ, a.�۸�" & vbNewLine & _
            IIf(gBs_Type.����ÿ�Ʊ, "", " Having Sum(a.���ʽ��) <> 0")
        
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng����ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ���ϸ���ݣ����ܴ�ӡ����Ʊ��"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        str�������� = Nvl(!����)
        str�����Ա� = Nvl(!�Ա�)
        str�������� = Nvl(!����)
        lng����ID = Val(Nvl(!����ID))
        strҽ�Ƹ��ʽ���� = Nvl(!���ʽ����)
        lng�Һ�ID = Val(Nvl(!�Һ�id))
        lngҽ����� = Val(Nvl(!ҽ�����))
        str�Ǽ�ʱ�� = Format(Nvl(!�Ǽ�ʱ��), "YYYYMMDDhhmmss000")
        strҵ�����Ա = Nvl(!����Ա����)
        
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!����ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!�վݷ�Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!�վݷ�Ŀ), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!������), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!�������), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!ҩƷ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!���), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!���㵥λ), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!�۸�)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!����)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!ʵ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!�Էѽ��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!Ӧ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!ҽ����������)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!��ע), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!���), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!��������), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dblƱ���ܽ�� = dblƱ���ܽ�� + RoundEx(Nvl(!ʵ�ս��), 6)
            .MoveNext
        Loop
        
        str����IDs = GetBalanceIDs(lng����ID, bytInvoiceType)
        dbl���� = GetBalanceErrorFee(str����IDs)
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
    End With
    
    '������ϸ
    If gBs_Type.���Ѷ��ձ��� <> "" Then
        dblƱ���ܽ�� = dblƱ���ܽ�� - dbl����
    End If
    dblƱ���ܽ�� = RoundEx(dblƱ���ܽ��, 2)
    If Not Get������ϸ(str����IDs, strData, dblƱ���ܽ��, bytInvoiceType, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    'Ʊ����Ϣ
    'ҵ����ˮ��:lng����ID_lngEInvoiceID
    strData = ""
    strData = strData & "" & GetJsonNodeString("busNo", lng����ID & "_" & lngEInvoiceID, Json_Text)
    strData = strData & "," & GetJsonNodeString("busType", strҵ���ʶ, Json_Text)
    strData = strData & "," & GetJsonNodeString("payer", str��������, Json_Text)
    strData = strData & "," & GetJsonNodeString("busDateTime", str�Ǽ�ʱ��, Json_Text)
    strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strData = strData & "," & GetJsonNodeString("payee", strҵ�����Ա, Json_Text)
    strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("totalAmt", dblƱ���ܽ��, Json_num)
    strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl����, 6) <> 0 <> 0 And gBs_Type.���Ѷ��ձ��� = "", "����" & FormatEx(dbl����, 6) & "�����������", ""), Json_Text)
    strJson_Out = strData
    
    
    '�ƶ�֧��(һ��)
    If Not Get�ƶ�֧����Ϣ(lng����ID, IIf(bln������, str����IDs, lng����ID), strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '֪ͨ��Ϣ
    If Not Get֪ͨ��Ϣ(lng����ID, strData, str�����) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '������Ϣ
    Call Getҽ����Ϣ(bytInvoiceType, lng����ID, lng����ID, cllInsureInfo)
    Set rsTmp = Nothing
    If lngҽ����� <> 0 Then
        strSQL = "Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')) As ��������, Max(b.����) As ������ұ���," & vbNewLine & _
                "       Max(b.����) As �����������, Max(a.No) As ������, Max(d.����) As ��������" & vbNewLine & _
                "  From ���˹Һż�¼ A, ���ű� B, ������ϼ�¼ C, ��������Ŀ¼ D" & vbNewLine & _
                "  Where a.ִ�в���id = b.Id And " & vbNewLine & _
                "   a.����ID = c.����ID(+) And a.ID = c.��ҳID(+) And c.��ϴ���(+) = 1 and Mod(c.�������(+), 10) = 1 And c.����ID = d.id(+) And " & vbNewLine & _
                "   a.No = (Select Max(�Һŵ�) From ����ҽ����¼ Where ID = [1] Or ���id = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lngҽ�����)
    ElseIf lng�Һ�ID <> 0 Then
        strSQL = "Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')) As ��������, Max(b.����) As ������ұ���," & vbNewLine & _
                "       Max(b.����) As �����������, Max(a.No) As ������, Max(d.����) As ��������" & vbNewLine & _
                "  From ���˹Һż�¼ A, ���ű� B, ������ϼ�¼ C, ��������Ŀ¼ D" & vbNewLine & _
                "  Where a.ִ�в���id = b.Id And a.Id = [1] And " & vbNewLine & _
                "   a.����ID = c.����ID(+) And a.ID = c.��ҳID(+) And c.��ϴ���(+) = 1 and Mod(c.�������(+), 10) = 1 And c.����ID = d.id(+) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng�Һ�ID)
    End If
    If rsTmp Is Nothing Then
        strSQL = "Select To_Char(a.����ʱ��, 'yyyy-mm-dd') As ��������, b.���� As ������ұ���," & vbNewLine & _
                "       b.���� As �����������, a.No As ������, d.���� As ��������" & vbNewLine & _
                "  From ���˹Һż�¼ A, ���ű� B, ������ϼ�¼ C, ��������Ŀ¼ D" & vbNewLine & _
                "  Where a.ִ�в���id = b.Id And " & vbNewLine & _
                "       a.����ID = c.����ID(+) And a.ID = c.��ҳID(+) And c.��ϴ���(+) = 1 and Mod(c.�������(+), 10) = 1 And c.����ID = d.id(+) And " & vbNewLine & _
                "       a.Id = (Select ID" & vbNewLine & _
                "           From (Select ID, ����ʱ�� From ���˹Һż�¼ Where ����id = [1] Order By ����ʱ�� Desc)" & vbNewLine & _
                "           Where Rownum < 2)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng����ID)
    End If
    
    strData = ""
    strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("ҽ�ƻ�������"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_���ջ�������", "C"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareTypeCode", strҽ�Ƹ��ʽ����, Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalCareType", Getҽ�Ƹ��ʽ����(strҽ�Ƹ��ʽ����), Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_ҽ����", "C"), Json_Text)
    With rsTmp
        If .RecordCount > 0 Then
            strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!��������), strJsonFormat_��������), Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_�������, Nvl(!�����������), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientCategoryCode", Nvl(!������ұ���), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!������), Json_Text)
        Else
            strData = strData & "," & GetJsonNodeString("consultationDate", "", Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_�������, "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientCategoryCode", "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", lng����ID, Json_Text)
        End If
    End With
    strData = strData & "," & GetJsonNodeString("patientId", lng����ID, Json_Text)
    strData = strData & "," & GetJsonNodeString("sex", str�����Ա�, Json_Text)
    strData = strData & "," & GetJsonNodeString("age", str��������, Json_Text)
    strData = strData & "," & GetJsonNodeString("caseNumber", str�����, Json_Text)
    strData = strData & "," & GetJsonNodeString("ICD", Nvl(rsTmp!��������), Json_Text)
    strData = strData & "," & GetJsonNodeString("specialDiseasesName", zlGetNodeValueFromCollect(cllInsureInfo, "_��������", "C"), Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '֧����Ϣ
    If Not Get������Ϣ(str����IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '�ɷ�����
    If Not Get�ɷ�����(str����IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '����ҽ����Ϣ-����
    '������չ��Ϣ-����
    'eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    'isArrears  �Ƿ����ͨ  String  1  ��  0-��1-�ǣ���Ƿ���������ҽԺҵ��Ҫ���Ʊ���Ƿ����ͨ��
    'arrearsReason  ������ͨԭ��  String  200  ��  isArrears=0����д������ͨ��ԭ��
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '�շ���Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strChargeDetail
    '�嵥��Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strListDetail
    
    '����������Json��
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByCharge = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByDeposit(ByVal lngEInvoiceID As Long, ByVal lngԤ��ID As Long, ByVal lng����ID As Long, _
                ByVal strEInvoiceClientCode As String, dblƱ���ܶ� As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�Һŷ�ƱJson��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    '        strEInvoiceClientCode-��Ʊ����
    '        lng����ID-����Ԥ��ID
    ' ���� : strJson-�ҺŽ�����Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsDeposit As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim bytInvoiceType As Byte
    Dim dblԤ����� As Double
    Dim lng����ID As Long
    Dim strҵ���ʶ As String, str�Ǽ�ʱ�� As String
    Dim strJsonList As String, strData As String, strChargeDetail As String
    Dim str������ As String, str���� As String, strNO As String
    On Error GoTo ErrHand
    bytInvoiceType = 2
    strҵ���ʶ = zlGetҵ���ʶ(bytInvoiceType)
    
    strSQL = "Select a.No, a.�տ�ʱ��, a.Ԥ�����, a.�����id, a.����id, a.��ҳid, a.����id, a.�ɿλ, a.��λ������, a.��λ�ʺ�, a.ժҪ, a.���㷽ʽ, a.�������, a.����," & vbNewLine & _
            "       a.������ˮ��, a.����˵��, a.������λ, a.���, a.����Ա���, a.����Ա����, Nvl(b.����, c.����) As ����, Nvl(b.�Ա�, c.�Ա�) As �Ա�," & vbNewLine & _
            "       Nvl(b.����, c.����) As ����, c.�����, Nvl(b.סԺ��, c.סԺ��) As סԺ��, c.Email, c.���֤��, c.�ֻ���, 1 As �ɿ�����," & vbNewLine & _
            "       Decode(Nvl(a.Ԥ�����, 0), 1, '07', '07') As ҵ���ʶ, d.���� As ��Ժ���ұ���, d.���� As ��Ժ��������, e.���� As ��Ժ���ұ���," & vbNewLine & _
            "       e.���� As ��Ժ��������, b.��Ժ����, b.��Ժ����, Nvl(b.������, b.סԺ��) As ������, j.���� As ҽ�ƿ�����" & vbNewLine & _
            "From ����Ԥ����¼ A, ������ҳ B, ������Ϣ C, ���ű� D, ���ű� E, ҽ�ƿ���� J" & vbNewLine & _
            "Where a.Id = [1] And a.����id = b.����id(+) And a.��ҳid = b.��ҳid(+) And a.����id = c.����id(+) And b.��Ժ����id = d.Id(+) And" & vbNewLine & _
            "      b.��Ժ����id = e.Id(+) And a.�����id = j.Id(+)"

    Set rsDeposit = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByDeposit", lngԤ��ID)
    If rsDeposit.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ��������ݣ����ܴ�ӡ����Ʊ��"
        Exit Function
    End If
    
    With rsDeposit
        If Nvl(!Ԥ�����) = 1 Then
            strErrMsg_Out = "��˼�ӿڲ�֧������Ԥ��Ʊ�����ɵ���Ʊ�ݣ�"
        End If
        lng����ID = Val(Nvl(!����ID))
        strNO = Nvl(!No)
        str�Ǽ�ʱ�� = Format(Nvl(!�տ�ʱ��), "YYYYMMDDhhmmss000")
    End With
    dblƱ���ܶ� = GetԤ�������ܶ�(strNO)
    dblԤ����� = GetԤ�����(lng����ID, Val(Nvl(rsDeposit!Ԥ�����)))
    
    If lng����ID <> 0 Then
        strSQL = "Select ����, ����, ƾ֤����, ƾ֤����" & vbNewLine & _
                " From ����Ʊ��ʹ�ü�¼" & vbNewLine & _
                " Where ID = [1] And �˿�id Is Null And ��¼״̬ = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByDeposit", lngEInvoiceID)
        If rsTmp.EOF Then
            strErrMsg_Out = "ԭʼԤ����δ���ߵ���Ʊ��ƾ֤�������п����˿�Ʊ�ݣ�"
            Exit Function
        End If
        
        strData = ""
        strData = strData & "" & GetJsonNodeString("busType", strҵ���ʶ, Json_Text)
        strData = strData & "," & GetJsonNodeString("billBatchCode", Nvl(rsTmp!����), Json_Text)
        strData = strData & "," & GetJsonNodeString("reason", "�˿�", Json_Text)
        strData = strData & "," & GetJsonNodeString("operator", mstrOperatorName, Json_Text)
        strData = strData & "," & GetJsonNodeString("busDateTime", str�Ǽ�ʱ��, Json_Text)
        strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
        strData = strData & "," & GetJsonNodeString("voucherBatchCode", Nvl(rsTmp!ƾ֤����), Json_Text)
        strData = strData & "," & GetJsonNodeString("voucherNo", Nvl(rsTmp!ƾ֤����), Json_Text)
        strData = strData & "," & GetJsonNodeString("amt", -1 * dblƱ���ܶ�, Json_num)
        strData = strData & "," & GetJsonNodeString("ownAcBalance", dblԤ�����, Json_num)
        strData = strData & "," & GetJsonNodeString("remark", Nvl(rsDeposit!ժҪ), Json_Text)
        '����������Json��
        strJson_Out = "{" & strJson_Out & "}"
        zlGetJson_CreateEInvoiceByDeposit = True
        Exit Function
    End If

    '�ɷ�����
    Call Get֪ͨ����(lng����ID, Nvl(rsDeposit!���֤��), str������, str����)
    strSQL = "Select c.��������" & vbNewLine & _
            "  From �շ��������� C" & vbNewLine & _
            "  Where c.�����id = [1] And c.���㷽ʽ = [2]" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select c.��������" & vbNewLine & _
            "  From �շ��������� C" & vbNewLine & _
            "  Where c.�����id Is Null And c.���㷽ʽ = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByDeposit", Val(Nvl(rsDeposit!�����ID)), Nvl(rsDeposit!���㷽ʽ))
    strData = ""
    If rsTmp.RecordCount > 0 Then
        strData = Nvl(rsTmp!��������)
    End If
    strData = GetJsonNodeString("payChannelCode", strData, Json_Text)
    strData = strData & "," & GetJsonNodeString("payChannelValue", dblƱ���ܶ�, Json_num)
    strJsonList = "{" & strData & "}"
    strJson_Out = GetNodeString("payChannelDetail") & ":[" & strJsonList & "]"
    
    With rsDeposit
        strData = ""
        strData = strData & "" & GetJsonNodeString("busType", strҵ���ʶ, Json_Text)
        '����ID_����Ʊ��ID
        strData = strData & "," & GetJsonNodeString("busNo", lngԤ��ID & "_" & lngEInvoiceID, Json_Text)
        strData = strData & "," & GetJsonNodeString("payer", Nvl(!����), Json_Text)
        strData = strData & "," & GetJsonNodeString("busDateTime", str�Ǽ�ʱ��, Json_Text)
        strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
        strData = strData & "," & GetJsonNodeString("payee", Nvl(!����Ա����), Json_Text)
        strData = strData & "," & GetJsonNodeString("drawee", Nvl(!����), Json_Text)
        strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
        strData = strData & "," & GetJsonNodeString("tel", Nvl(!�ֻ���), Json_Text)
        strData = strData & "," & GetJsonNodeString("email", Nvl(!email), Json_Text)
        strData = strData & "," & GetJsonNodeString("idCardNo", Nvl(!���֤��), Json_Text)
        strData = strData & "," & GetJsonNodeString("cardType", str������, Json_Text)
        strData = strData & "," & GetJsonNodeString("cardNo", str����, Json_Text)
        strData = strData & "," & GetJsonNodeString("amt", dblƱ���ܶ�, Json_num)
        strData = strData & "," & GetJsonNodeString("ownAcBalance", dblԤ�����, Json_num)
        strData = strData & "," & GetJsonNodeString("category", Nvl(!��Ժ��������), Json_Text)
        strData = strData & "," & GetJsonNodeString("categoryCode", Nvl(!��Ժ���ұ���), Json_Text)
        strData = strData & "," & GetJsonNodeString("inHospitalDate", Format(Nvl(!��Ժ����), "yyyy-MM-dd"), Json_Text)
        strData = strData & "," & GetJsonNodeString("hospitalNo", Nvl(!סԺ��), Json_Text)
        strData = strData & "," & GetJsonNodeString("patientId", Nvl(!����ID), Json_Text)
        strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!��ҳid), Json_Text)
        strData = strData & "," & GetJsonNodeString("caseNumber", Nvl(!������), Json_Text)
        strData = strData & "," & GetJsonNodeString("accountName", Nvl(!ҽ�ƿ�����), Json_Text)
        strData = strData & "," & GetJsonNodeString("accountNo", Nvl(!��λ�ʺ�), Json_Text)
        strData = strData & "," & GetJsonNodeString("accountBank", IIf(Nvl(!ҽ�ƿ�����) <> "", Nvl(!ҽ�ƿ�����), Nvl(!��λ������)), Json_Text)
        strData = strData & "," & GetJsonNodeString("remark", Nvl(!ժҪ), Json_Text)
        If gBs_Type.֧�ְ汾 > BS_Version.V2_0_3 Then
            strData = strData & "," & GetJsonNodeString("workUnit", Nvl(!�ɿλ), Json_Text)
        End If
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '����������Json��
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByDeposit = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByMzBalance(ByVal lngEInvoiceID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                ByVal strEInvoiceClientCode As String, dblƱ���ܽ�� As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�Һŷ�ƱJson��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    '        strEInvoiceClientCode-��Ʊ����
    ' ���� : strJson-�ҺŽ�����Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsBalance As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bytInvoiceType As Byte
    Dim dbl���� As Double
    Dim lng����ID As Long, lng�Һ�ID As Long, lngҽ����� As Long
    Dim str������ As String, str���� As String, strChargeDetail As String, strListDetail As String
    Dim str�������� As String, str�����Ա� As String, str�������� As String, strҽ�Ƹ��ʽ���� As String
    Dim str����� As String
    Dim strJsonList As String, strData As String
    Dim strJsonKey_������� As String
    Dim strJsonFormat_�������� As String
    Dim intJsonFormat_����С�� As Integer, intJsonFormat_����С�� As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 3
    dbl���� = GetBalanceErrorFee(lng����ID)
    dblƱ���ܽ�� = 0
    
    '�汾����
    strJsonKey_������� = GetVersionDiff(1, "�������")
    strJsonFormat_�������� = GetVersionDiff(2, "��������")
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    
    strSQL = "Select a.No, a.�շ�ʱ��, a.��������, a.����Ա���, a.����Ա����, a.����id, a.��ҳid, Decode(Nvl(a.����id, 0), 0, a.ԭ��, c.����) As ����," & vbNewLine & _
            "       '' As �Ա�, '' As ����, c.�����, a.��ע, a.���ʽ��, Decode(Nvl(a.����id, 0), 0, q.�����ʼ�, c.Email) As Email, q.��ϵ��," & vbNewLine & _
            "       Decode(Nvl(a.����id, 0), 0, q.������ô���, c.���֤��) As ���֤��," & vbNewLine & _
            "       Decode(Nvl(a.����id, 0), 0, Nvl(q.�绰, To_Char(j.�ƶ��绰)), c.�ֻ���) As �ֻ���," & vbNewLine & _
            "       Decode(Nvl(a.����id, 0), 0, 2, 1) As �ɿ�����, Decode(Nvl(a.��������, 0), 1, '02', '01') As ҵ���ʶ, c.����� As ������" & vbNewLine & _
            "From ���˽��ʼ�¼ A, ������Ϣ C, ��Լ��λ Q, ��Ա�� J" & vbNewLine & _
            "Where a.Id = [1] And a.����id = c.����id(+) And a.ԭ�� = q.����(+) And q.��ϵ�� = j.����(+)"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByMzBalance", lng����ID)
    If rsBalance.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ��������ݣ����ܴ�ӡ����Ʊ��"
        Exit Function
    End If
    
    strSQL = "      Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ," & vbNewLine & _
            "              Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
            "              Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��," & vbNewLine & _
            "              Max(s.�������) As ҽ����Ŀ����, Max(s.��������) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע," & vbNewLine & _
            "              Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����," & vbNewLine & _
            "              Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��," & vbNewLine & _
            "              Max(a.���ʽ) As ���ʽ����, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����," & vbNewLine & _
            "              Max(a.�Һ�id) As �Һ�id, Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����," & vbNewLine & _
            "              Max(b.���) As ���, Max(q.ҩƷ����) As ҩƷ����" & vbNewLine & _
            "       From ������ü�¼ A, �շ���ĿĿ¼ B, �վݷ�Ŀ���� C, �վݷ�Ŀ C1, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S" & vbNewLine & _
            "       Where a.no In (Select No From ������ü�¼ Where ����ID = [1]) And a.���ʷ��� = 1 And a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And " & vbNewLine & _
            "             a.�վݷ�Ŀ = c1.����(+) And a.�վݷ�Ŀ = c.�վݷ�Ŀ(+) And Decode(c.���ó���(+), 0, 1, c.���ó���(+)) = 1 and" & vbNewLine & _
            "             a.�շ�ϸĿid = m.ҩƷid(+) And m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And" & vbNewLine & _
            "             t.����(+) = 1 And a.���մ���id = s.���մ���id(+)" & vbNewLine & _
            "       Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����" & vbNewLine & _
            "       Order By a.NO, ���"
    
    strSQL = "Select Min(a.����ID) As ����ID, a.No, a.���, a.�շ�ϸĿid, a.���㵥λ, Avg(a.�۸�) As �۸�, Avg(a.����) As ����," & vbNewLine & _
            "       Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.�Էѽ��) As �Էѽ��," & vbNewLine & _
            "       Max(a.ҽ����Ŀ����) As ҽ����Ŀ����, Max(a.ҽ����Ŀ����) As ҽ����Ŀ����, Max(a.ҽ����������) As ҽ����������, Max(a.��ע) As ��ע," & vbNewLine & _
            "       Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����," & vbNewLine & _
            "       Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��," & vbNewLine & _
            "       Max(a.���ʽ����) As ���ʽ����, Max(a.�վݷ�Ŀ) As �վݷ�Ŀ, Max(a.�վݷ�Ŀ����) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����," & vbNewLine & _
            "       Max(a.�Һ�id) As �Һ�id, Max(a.������) As ������, Max(a.�������) As �������, Max(a.��Ŀ����) As ��Ŀ����, Max(a.��Ŀ����) As ��Ŀ����," & vbNewLine & _
            "       Max(a.���) As ���, Max(a.ҩƷ����) As ҩƷ����" & vbNewLine & _
            " From (" & strSQL & ") a" & vbNewLine & _
            " Group By a.No, a.���, a.�շ�ϸĿid, a.���㵥λ" & vbNewLine & _
            IIf(gBs_Type.����ÿ�Ʊ, "", " Having Sum(a.���ʽ��) <> 0")

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByMzBalance", lng����ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ���ϸ���ݣ����ܴ�ӡ����Ʊ��"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        str�������� = Nvl(!����)
        str�����Ա� = Nvl(!�Ա�)
        str�������� = Nvl(!����)
        lng����ID = Val(Nvl(!����ID))
        strҽ�Ƹ��ʽ���� = Nvl(!���ʽ����)
        lng�Һ�ID = Val(Nvl(!�Һ�id))
        lngҽ����� = Val(Nvl(!ҽ�����))
        
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!����ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!�վݷ�Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!�վݷ�Ŀ), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!������), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!�������), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!ҩƷ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!���), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!���㵥λ), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!�۸�)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!����)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!ʵ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!�Էѽ��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!Ӧ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!ҽ����������)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!��ע), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!���), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!��������), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dblƱ���ܽ�� = dblƱ���ܽ�� + RoundEx(Val(Nvl(!ʵ�ս��)), 6)
            .MoveNext
        Loop
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
    End With

    '������ϸ
    If gBs_Type.���Ѷ��ձ��� <> "" Then
        dblƱ���ܽ�� = dblƱ���ܽ�� - dbl����
    End If
    dblƱ���ܽ�� = RoundEx(dblƱ���ܽ��, 2)
    If Not Get������ϸ(lng����ID, strData, dblƱ���ܽ��, 3.1, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    'Ʊ����Ϣ
    'ҵ����ˮ��:lng����ID_lngEInvoiceID
    strData = ""
    strData = strData & "" & GetJsonNodeString("busNo", lng����ID & "_" & lngEInvoiceID, Json_Text)
    strData = strData & "," & GetJsonNodeString("busType", Nvl(rsBalance!ҵ���ʶ), Json_Text)
    If Val(Nvl(rsBalance!����ID)) = 0 Then
        strData = strData & "," & GetJsonNodeString("payer", Nvl(rsBalance!����), Json_Text)
    Else
        strData = strData & "," & GetJsonNodeString("payer", str��������, Json_Text)
    End If
    strData = strData & "," & GetJsonNodeString("busDateTime", Format(Nvl(rsBalance!�շ�ʱ��), "YYYYMMDDhhmmss000"), Json_Text)
    strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strData = strData & "," & GetJsonNodeString("payee", Nvl(rsBalance!����Ա����), Json_Text)
    strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("totalAmt", dblƱ���ܽ��, Json_num)
    strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl����, 6) <> 0 And gBs_Type.���Ѷ��ձ��� = "", "����" & FormatEx(dbl����, 6) & "�����������", Nvl(rsBalance!��ע)), Json_Text)
    strJson_Out = strData
    
    
    '�ƶ�֧��(һ��)
    If Not Get�ƶ�֧����Ϣ(lng����ID, lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '֪ͨ��Ϣ
    With rsBalance
        Call Get֪ͨ����(Val(Nvl(!����ID)), Nvl(!���֤��), str������, str����)
        strData = ""
        strData = strData & "" & GetJsonNodeString("tel", Nvl(!�ֻ���), Json_Text)
        strData = strData & "," & GetJsonNodeString("email", Nvl(!email), Json_Text)
        If gBs_Type.֧�ְ汾 > BS_Version.V2_0_3 Then
            strData = strData & "," & GetJsonNodeString("payerType", Nvl(!�ɿ�����), Json_Text)
        End If
        strData = strData & "," & GetJsonNodeString("idCardNo", Nvl(!���֤��), Json_Text)
        strData = strData & "," & GetJsonNodeString("cardType", str������, Json_Text)
        strData = strData & "," & GetJsonNodeString("cardNo", str����, Json_Text)
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '������Ϣ
    With rsBalance
        Call Getҽ����Ϣ(bytInvoiceType, lng����ID, Val(Nvl(!����ID)), cllInsureInfo)
        Set rsTmp = Nothing
        If lngҽ����� <> 0 Then
            strSQL = "Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')) As ��������, Max(b.����) As ������ұ���," & vbNewLine & _
                    "       Max(b.����) As �����������, Max(a.No) As ������, Max(d.����) As ��������" & vbNewLine & _
                    "  From ���˹Һż�¼ A, ���ű� B, ������ϼ�¼ C, ��������Ŀ¼ D" & vbNewLine & _
                    "  Where a.ִ�в���id = b.Id And " & vbNewLine & _
                    "   a.����ID = c.����ID(+) And a.ID = c.��ҳID(+) And c.��ϴ���(+) = 1 and Mod(c.�������(+), 10) = 1 And c.����ID = d.id(+) And " & vbNewLine & _
                    "   a.No = (Select Max(�Һŵ�) From ����ҽ����¼ Where ID = [1] Or ���id = [1])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lngҽ�����)
        ElseIf lng�Һ�ID <> 0 Then
            strSQL = "Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')) As ��������, Max(b.����) As ������ұ���," & vbNewLine & _
                    "       Max(b.����) As �����������, Max(a.No) As ������, Max(d.����) As ��������" & vbNewLine & _
                    "  From ���˹Һż�¼ A, ���ű� B, ������ϼ�¼ C, ��������Ŀ¼ D" & vbNewLine & _
                    "  Where a.ִ�в���id = b.Id And a.Id = [1] And " & vbNewLine & _
                    "   a.����ID = c.����ID(+) And a.ID = c.��ҳID(+) And c.��ϴ���(+) = 1 and Mod(c.�������(+), 10) = 1 And c.����ID = d.id(+) "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng�Һ�ID)
        End If
        If rsTmp Is Nothing Then
            strSQL = "Select To_Char(a.����ʱ��, 'yyyy-mm-dd') As ��������, b.���� As ������ұ���," & vbNewLine & _
                    "       b.���� As �����������, a.No As ������, d.���� As ��������" & vbNewLine & _
                    "  From ���˹Һż�¼ A, ���ű� B, ������ϼ�¼ C, ��������Ŀ¼ D" & vbNewLine & _
                    "  Where a.ִ�в���id = b.Id And " & vbNewLine & _
                    "       a.����ID = c.����ID(+) And a.ID = c.��ҳID(+) And c.��ϴ���(+) = 1 and Mod(c.�������(+), 10) = 1 And c.����ID = d.id(+) And " & vbNewLine & _
                    "       a.Id = (Select ID" & vbNewLine & _
                    "           From (Select ID, ����ʱ�� From ���˹Һż�¼ Where ����id = [1] Order By ����ʱ�� Desc)" & vbNewLine & _
                    "           Where Rownum < 2)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", Val(Nvl(!����ID)))
        End If
        
        strData = ""
        strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("ҽ�ƻ�������"), Json_Text)
        strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_���ջ�������", "C"), Json_Text)
        strData = strData & "," & GetJsonNodeString("medCareTypeCode", strҽ�Ƹ��ʽ����, Json_Text)
        strData = strData & "," & GetJsonNodeString("medicalCareType", Getҽ�Ƹ��ʽ����(strҽ�Ƹ��ʽ����), Json_Text)
        strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_ҽ����", "C"), Json_Text)
        With rsTmp
            If .RecordCount > 0 Then
                strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!��������), strJsonFormat_��������), Json_Text)
                strData = strData & "," & GetJsonNodeString(strJsonKey_�������, Nvl(!�����������), Json_Text)
                strData = strData & "," & GetJsonNodeString("patientCategoryCode", Nvl(!������ұ���), Json_Text)
                strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!������), Json_Text)
            Else
                strData = strData & "," & GetJsonNodeString("consultationDate", "", Json_Text)
                strData = strData & "," & GetJsonNodeString(strJsonKey_�������, "", Json_Text)
                strData = strData & "," & GetJsonNodeString("patientCategoryCode", "", Json_Text)
                strData = strData & "," & GetJsonNodeString("patientNo", lng����ID, Json_Text)
            End If
        End With
        strData = strData & "," & GetJsonNodeString("patientId", Nvl(!����ID), Json_Text)
        If Val(Nvl(!����ID)) = 0 Then
            strData = strData & "," & GetJsonNodeString("sex", Nvl(!�Ա�), Json_Text)
            strData = strData & "," & GetJsonNodeString("age", Nvl(!����), Json_Text)
        Else
            strData = strData & "," & GetJsonNodeString("sex", str�����Ա�, Json_Text)
            strData = strData & "," & GetJsonNodeString("age", str��������, Json_Text)
        End If
        strData = strData & "," & GetJsonNodeString("caseNumber", Nvl(!������), Json_Text)
        strData = strData & "," & GetJsonNodeString("ICD", Nvl(rsTmp!��������), Json_Text)
        strData = strData & "," & GetJsonNodeString("specialDiseasesName", zlGetNodeValueFromCollect(cllInsureInfo, "_��������", "C"), Json_Text)
        
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '֧����Ϣ
    If Not Get������Ϣ(lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '�ɷ�����
    If Not Get�ɷ�����(lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '����ҽ����Ϣ-����
    '������չ��Ϣ-����
    'eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    'isArrears  �Ƿ����ͨ  String  1  ��  0-��1-�ǣ���Ƿ���������ҽԺҵ��Ҫ���Ʊ���Ƿ����ͨ��
    'arrearsReason  ������ͨԭ��  String  200  ��  isArrears=0����д������ͨ��ԭ��
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '�շ���Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strChargeDetail
    '�嵥��Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strListDetail
    
    '����������Json��
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByMzBalance = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByZyBalance(ByVal lngEInvoiceID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                ByVal strEInvoiceClientCode As String, dblƱ���ܽ�� As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�Һŷ�ƱJson��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    '        strEInvoiceClientCode-��Ʊ����
    ' ���� : strJson-�ҺŽ�����Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsBalance As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bytInvoiceType As Byte
    Dim dbl���� As Double
    Dim strҽ�Ƹ��ʽ���� As String, str������ As String, str���� As String
    Dim str����� As String, strסԺ���� As String
    Dim strJsonList As String, strData As String, strChargeDetail As String, strListDetail As String
    Dim strJsonKey_������� As String
    Dim strJsonFormat_�������� As String
    Dim intJsonFormat_����С�� As Integer, intJsonFormat_����С�� As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 3
    dbl���� = GetBalanceErrorFee(lng����ID)
    dblƱ���ܽ�� = 0
    
    '�汾����
    strJsonKey_������� = GetVersionDiff(1, "�������")
    strJsonFormat_�������� = GetVersionDiff(2, "��������")
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    
    strSQL = "Select a.No, a.�շ�ʱ��, a.��������, a.����Ա���, a.����Ա����, a.����id, a.��ҳid," & vbNewLine & _
            "       Decode(Nvl(a.����id, 0), 0, a.ԭ��, Nvl(b.����, c.����)) As ����, Nvl(b.�Ա�, c.�Ա�) As �Ա�, Nvl(b.����, c.����) As ����, c.�����," & vbNewLine & _
            "       Nvl(b.סԺ��, c.סԺ��) As סԺ��, a.��ʼ����, a.��������, a.��ע, a.���ʽ��, Decode(Nvl(a.����id, 0), 0, q.�����ʼ�, c.Email) As Email," & vbNewLine & _
            "       q.��ϵ��, Decode(Nvl(a.����id, 0), 0, q.������ô���, c.���֤��) As ���֤��," & vbNewLine & _
            "       Decode(Nvl(a.����id, 0), 0, Nvl(q.�绰, To_Char(j.�ƶ��绰)), c.�ֻ���) As �ֻ���," & vbNewLine & _
            "       Decode(Nvl(a.����id, 0), 0, 2, 1) As �ɿ�����, Decode(Nvl(a.��������, 0), 1, '02', '01') As ҵ���ʶ, b.��Ժ����, b.��Ժ����," & vbNewLine & _
            "       m.���� As ��Ժ���ұ���, m.���� As ��Ժ��������, p.���� As ��Ժ���ұ���, p.���� As ��Ժ��������, b.��Ժ���� As ����, t.���� As ��������," & vbNewLine & _
            "       Nvl(b.������, b.סԺ��) As ������, Nvl(b.ҽ�Ƹ��ʽ, c.ҽ�Ƹ��ʽ) As ҽ�Ƹ��ʽ, Nvl(b.��Ժ����, Sysdate) - b.��Ժ���� As סԺ����, f.���� As ��������" & vbNewLine & _
            "From ���˽��ʼ�¼ A, ������ҳ B, ������Ϣ C, ��Լ��λ Q, ��Ա�� J, ���ű� M, ���ű� P, ���ű� T, ������ϼ�¼ E, ��������Ŀ¼ F" & vbNewLine & _
            "Where a.Id = [1] And a.����id = b.����id(+) And a.��ҳid = b.��ҳid(+) And a.����id = c.����id(+) And a.ԭ�� = q.����(+) And" & vbNewLine & _
            "      a.����ID = e.����ID(+) And a.��ҳid = e.��ҳID(+) And e.��ϴ���(+) = 1 And Mod(e.�������(+), 10) = 1 And e.����ID = f.id(+) And " & vbNewLine & _
            "      b.��Ժ����id = m.Id(+) And b.��Ժ����id = p.Id(+) And b.��ǰ����id = t.Id(+)" & vbNewLine & _
            "      And q.��ϵ�� = j.����(+)"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByZyBalance", lng����ID)
    If rsBalance.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ��������ݣ����ܴ�ӡ����Ʊ��"
        Exit Function
    End If
    
    strSQL = "Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ," & vbNewLine & _
            "        Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
            "        Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��," & vbNewLine & _
            "        Max(s.�������) As ҽ����Ŀ����, Max(s.��������) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע," & vbNewLine & _
            "        Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����id) As ����id," & vbNewLine & _
            "        Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ����, Max(a.��ҳid) As ��ҳid," & vbNewLine & _
            "        Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����, Max(b.���) As ���," & vbNewLine & _
            "        Max(q.ҩƷ����) As ҩƷ����" & vbNewLine & _
            " From סԺ���ü�¼ A, �շ���ĿĿ¼ B, �վݷ�Ŀ���� C, �վݷ�Ŀ C1, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S" & vbNewLine & _
            " Where a.No In (Select Distinct NO From סԺ���ü�¼ Where ����id = [1]) And a.���ʷ��� = 1 And a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And" & vbNewLine & _
            "       a.�վݷ�Ŀ = c1.����(+) And a.�վݷ�Ŀ = c.�վݷ�Ŀ(+) And Decode(c.���ó���(+), 0, 2, c.���ó���(+)) = 2 and" & vbNewLine & _
            "       a.�շ�ϸĿid = m.ҩƷid(+) And m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And" & vbNewLine & _
            "       t.����(+) = 1 And a.���մ���id = s.���մ���id(+)" & vbNewLine & _
            " Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����" & vbNewLine & _
            " Order By NO, ���"
    
    strSQL = "Select Min(a.����ID) As ����ID, a.No, a.���, a.�շ�ϸĿid, a.���㵥λ, Avg(a.�۸�) As �۸�, Avg(a.����) As ����," & vbNewLine & _
            "       Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.�Էѽ��) As �Էѽ��," & vbNewLine & _
            "       Max(a.ҽ����Ŀ����) As ҽ����Ŀ����, Max(a.ҽ����Ŀ����) As ҽ����Ŀ����, Max(a.ҽ����������) As ҽ����������, Max(a.��ע) As ��ע," & vbNewLine & _
            "       Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, " & vbNewLine & _
            "       Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(a.�վݷ�Ŀ) As �վݷ�Ŀ, Max(a.�վݷ�Ŀ����) As �վݷ�Ŀ����, Max(a.��ҳid) As ��ҳid," & vbNewLine & _
            "       Max(a.������) As ������, Max(a.�������) As �������, Max(a.��Ŀ����) As ��Ŀ����, Max(a.��Ŀ����) As ��Ŀ����," & vbNewLine & _
            "       Max(a.���) As ���, Max(a.ҩƷ����) As ҩƷ����" & vbNewLine & _
            " From (" & strSQL & ") a" & vbNewLine & _
            " Group By a.No, a.���, a.�շ�ϸĿid, a.���㵥λ" & vbNewLine & _
            IIf(gBs_Type.����ÿ�Ʊ, "", " Having Sum(a.���ʽ��) <> 0")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByZyBalance", lng����ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ���ϸ���ݣ����ܴ�ӡ����Ʊ��"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        Do While Not .EOF
            If InStr(1, strסԺ���� & ",", "," & Nvl(!��ҳid, 0) & ",") = 0 Then
              strסԺ���� = strסԺ���� & "," & Nvl(!��ҳid, 0)
            End If
      
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!����ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!�վݷ�Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!�վݷ�Ŀ), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!������), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!�������), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!ҩƷ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!���), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!���㵥λ), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!�۸�)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!����)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!ʵ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!�Էѽ��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!Ӧ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!ҽ����������)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!��ע), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!���), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!��������), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dblƱ���ܽ�� = dblƱ���ܽ�� + RoundEx(Nvl(!ʵ�ս��), 6)
            .MoveNext
        Loop
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
        strסԺ���� = Mid(strסԺ����, 2)
    End With

    '������ϸ
    If gBs_Type.���Ѷ��ձ��� <> "" Then
        dblƱ���ܽ�� = dblƱ���ܽ�� - dbl����
    End If
    dblƱ���ܽ�� = RoundEx(dblƱ���ܽ��, 2)
    If Not Get������ϸ(lng����ID, strData, dblƱ���ܽ��, 3.2, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    'Ʊ����Ϣ
    With rsBalance
        'ҵ����ˮ��:lng����ID_lngEInvoiceID
        strData = ""
        strData = strData & "" & GetJsonNodeString("busNo", lng����ID & "_" & lngEInvoiceID, Json_Text)
        strData = strData & "," & GetJsonNodeString("busType", Nvl(rsBalance!ҵ���ʶ), Json_Text)
        strData = strData & "," & GetJsonNodeString("payer", Nvl(rsBalance!����), Json_Text)
        strData = strData & "," & GetJsonNodeString("busDateTime", Format(Nvl(rsBalance!�շ�ʱ��), "YYYYMMDDhhmmss000"), Json_Text)
        strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
        strData = strData & "," & GetJsonNodeString("payee", Nvl(rsBalance!����Ա����), Json_Text)
        strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
        strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
        strData = strData & "," & GetJsonNodeString("totalAmt", dblƱ���ܽ��, Json_num)
        strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl����, 6) <> 0 And gBs_Type.���Ѷ��ձ��� = "", "����" & FormatEx(dbl����, 6) & "�����������", Nvl(rsBalance!��ע)), Json_Text)
        strJson_Out = strData
    End With
    
    '�ƶ�֧��(һ��)
    If Not Get�ƶ�֧����Ϣ(Val(Nvl(rsBalance!����ID)), lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '֪ͨ��Ϣ
    With rsBalance
        Call Get֪ͨ����(Val(Nvl(!����ID)), Nvl(!���֤��), str������, str����)
        strData = ""
        strData = strData & "" & GetJsonNodeString("tel", Nvl(!�ֻ���), Json_Text)
        strData = strData & "," & GetJsonNodeString("email", Nvl(!email), Json_Text)
        If gBs_Type.֧�ְ汾 > BS_Version.V2_0_3 Then
            strData = strData & "," & GetJsonNodeString("payerType", Nvl(!�ɿ�����), Json_Text)
        End If
        strData = strData & "," & GetJsonNodeString("idCardNo", Nvl(!���֤��), Json_Text)
        strData = strData & "," & GetJsonNodeString("cardType", str������, Json_Text)
        strData = strData & "," & GetJsonNodeString("cardNo", str����, Json_Text)
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '������Ϣ
    With rsBalance
        Call Getҽ����Ϣ(bytInvoiceType, lng����ID, Val(Nvl(!����ID)), cllInsureInfo)
        strData = ""
        strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("ҽ�ƻ�������"), Json_Text)
        strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_���ջ�������", "C"), Json_Text)
        strData = strData & "," & GetJsonNodeString("medCareTypeCode", Getҽ�Ƹ��ʽ����(Nvl(!ҽ�ƿ����ʽ)), Json_Text)
        strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!ҽ�ƿ����ʽ), Json_Text)
        strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_ҽ����", "C"), Json_Text)
        strData = strData & "," & GetJsonNodeString("category", Nvl(!��Ժ��������), Json_Text)
        strData = strData & "," & GetJsonNodeString("categoryCode", Nvl(!��Ժ���ұ���), Json_Text)
        strData = strData & "," & GetJsonNodeString("leaveCategory", Nvl(!��Ժ��������), Json_Text)
        strData = strData & "," & GetJsonNodeString("leaveCategoryCode", Nvl(!��Ժ���ұ���), Json_Text)
        strData = strData & "," & GetJsonNodeString("hospitalNo", Nvl(!סԺ��), Json_Text)
        strData = strData & "," & GetJsonNodeString("visitNo", Nvl(!סԺ��), Json_Text)
        strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!��Ժ����), strJsonFormat_��������), Json_Text)
        strData = strData & "," & GetJsonNodeString("patientId", Nvl(!����ID), Json_Text)
        strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!��ҳid), Json_Text)
        strData = strData & "," & GetJsonNodeString("sex", Nvl(!�Ա�), Json_Text)
        strData = strData & "," & GetJsonNodeString("age", Nvl(!����), Json_Text)
        strData = strData & "," & GetJsonNodeString("hospitalArea", Nvl(!��������), Json_Text)
        strData = strData & "," & GetJsonNodeString("bedNo", Nvl(!����), Json_Text)
        strData = strData & "," & GetJsonNodeString("caseNumber", Nvl(!������), Json_Text)
        strData = strData & "," & GetJsonNodeString("ICD", Nvl(!��������), Json_Text)
        If InStr(1, strסԺ����, ",") > 0 Then
            strSQL = "Select Min(��Ժ����) As ��Ժ����, Max(��Ժ����) As ��Ժ����, Sum(Nvl(��Ժ����, Sysdate) - ��Ժ����) As סԺ����" & vbNewLine & _
                    "From ������ҳ" & vbNewLine & _
                    "Where ����id = [1] And" & vbNewLine & _
                    "      ��ҳid In (Select /*+cardinality(A,10)*/" & vbNewLine & _
                    "                Column_Value" & vbNewLine & _
                    "               From Table(f_Num2list([2])) a)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByZyBalance", Val(Nvl(!����ID)), strסԺ����)
            With rsTmp
                If Not .EOF Then
                    strData = strData & "," & GetJsonNodeString("inHospitalDate", Format(Nvl(!��Ժ����), strJsonFormat_��������), Json_Text)
                    strData = strData & "," & GetJsonNodeString("outHospitalDate", Format(Nvl(!��Ժ����), strJsonFormat_��������), Json_Text)
                    strData = strData & "," & GetJsonNodeString("hospitalDays", FormatEx(Nvl(!סԺ����), 2), Json_num)
                End If
            End With
        Else
            strData = strData & "," & GetJsonNodeString("inHospitalDate", Format(Nvl(!��Ժ����), strJsonFormat_��������), Json_Text)
            strData = strData & "," & GetJsonNodeString("outHospitalDate", Format(Nvl(!��Ժ����), strJsonFormat_��������), Json_Text)
            strData = strData & "," & GetJsonNodeString("hospitalDays", FormatEx(Nvl(!סԺ����), 2), Json_num)
        End If
        strJson_Out = strJson_Out & "," & strData
    End With
    
    'Ԥ��֧��
    strSQL = "Select q.ƾ֤����, q.ƾ֤����, a.No, Max(a.��Ԥ��) As ��Ԥ��" & vbNewLine & _
            "  From (Select NO, Sum(��Ԥ��) As ��Ԥ��" & vbNewLine & _
            "          From ����Ԥ����¼" & vbNewLine & _
            "          Where ����id = [1] And Mod(��¼����, 10) = 1 Group by No) A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ Q" & vbNewLine & _
            "  Where a.No = b.No And b.��¼���� = 1 And b.Id = q.����id And q.Ʊ�� = 2 And Q.�˿�ID is null" & vbNewLine & _
            " Group by q.ƾ֤����, q.ƾ֤����, a.No"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByZyBalance", lng����ID)
    strJsonList = ""
    With rsTmp
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("voucherBatchCode", Nvl(!ƾ֤����), Json_Text)
            strData = strData & "," & GetJsonNodeString("voucherNo", Nvl(!ƾ֤����), Json_Text)
            strData = strData & "," & GetJsonNodeString("voucherAmt", FormatEx(Nvl(!ƾ֤����, 0), 6), Json_num)
            strJsonList = strJsonList & ",{" & strData & "}"
            .MoveNext
        Loop
        strJson_Out = strJson_Out & "," & GetNodeString("payMentVoucher") & ":[" & Mid(strJsonList, 2) & "]"
    End With
    
    '֧����Ϣ
    If Not Get������Ϣ(lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '�ɷ�����
    If Not Get�ɷ�����(lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '����ҽ����Ϣ-����
    '������չ��Ϣ-����
    'eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    'isArrears  �Ƿ����ͨ  String  1  ��  0-��1-�ǣ���Ƿ���������ҽԺҵ��Ҫ���Ʊ���Ƿ����ͨ��
    'arrearsReason  ������ͨԭ��  String  200  ��  isArrears=0����д������ͨ��ԭ��
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strData = strData & "," & GetJsonNodeString("isArrears", "1", Json_Text)
    strData = strData & "," & GetJsonNodeString("arrearsReason", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '�շ���Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strChargeDetail
    '�嵥��Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strListDetail
    
    '����������Json��
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByZyBalance = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceBySendCard(ByVal lngEInvoiceID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                ByVal strEInvoiceClientCode As String, dblƱ���ܽ�� As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������ƱJson��ʽ����
    ' ��� : lngEInvoiceID -����Ʊ��ʹ�ü�¼.ID
    '        strEInvoiceClientCode-��Ʊ����
    ' ���� : strJson-�ҺŽ�����Ϣ
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bytInvoiceType As Byte
    Dim dbl���� As Double
    Dim lng����ID As Long, lngҽ����� As Long
    Dim strҵ���ʶ As String, strChargeDetail As String, strListDetail As String
    Dim str����IDs As String, str�Ǽ�ʱ�� As String, strҵ�����Ա As String
    Dim str�������� As String, str�����Ա� As String, str�������� As String, strҽ�Ƹ��ʽ���� As String
    Dim strJsonList As String, strData As String
    Dim strJsonKey_������� As String
    Dim strJsonFormat_�������� As String
    Dim intJsonFormat_����С�� As Integer, intJsonFormat_����С�� As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 5
    strҵ���ʶ = zlGetҵ���ʶ(bytInvoiceType)
    dblƱ���ܽ�� = 0
    
    '�汾����
    strJsonKey_������� = GetVersionDiff(1, "�������")
    strJsonFormat_�������� = GetVersionDiff(2, "��������")
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    intJsonFormat_����С�� = Val(GetVersionDiff(2, "����С��"))
    
    strSQL = "Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ," & vbNewLine & _
            "        Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
            "        Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��," & vbNewLine & _
            "        Max(s.�������) As ҽ����Ŀ����, Max(s.��������) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע," & vbNewLine & _
            "        Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����," & vbNewLine & _
            "        Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max('') As ���ʽ����," & vbNewLine & _
            "        Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ, Max(Nvl(c.����, c1.����)) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����, Max(0) As �Һ�id," & vbNewLine & _
            "        Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����, Max(b.���) As ���," & vbNewLine & _
            "        Max(q.ҩƷ����) As ҩƷ����" & vbNewLine & _
            " From סԺ���ü�¼ A, �շ���ĿĿ¼ B, �վݷ�Ŀ���� C, �վݷ�Ŀ C1, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S" & vbNewLine & _
            " Where a.No In (Select Distinct NO From סԺ���ü�¼ Where ����id = [1]) And a.��¼���� = 5 And a.��¼״̬ = 1 And" & vbNewLine & _
            "       a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And a.�վݷ�Ŀ = c1.����(+) And a.�վݷ�Ŀ = c.�վݷ�Ŀ(+) and Decode(c.���ó���(+), 0, 1, c.���ó���(+)) = 1 And a.�շ�ϸĿid = m.ҩƷid(+) And" & vbNewLine & _
            "       m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And t.����(+) = 1 And" & vbNewLine & _
            "       a.���մ���id = s.���մ���id(+)" & vbNewLine & _
            " Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����" & vbNewLine & _
            IIf(gBs_Type.����ÿ�Ʊ, "", " Having Sum(a.���ʽ��) <> 0") & vbNewLine & _
            " Order By NO, ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceBySendCard", lng����ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ���ϸ���ݣ����ܴ�ӡ����Ʊ��"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        str�������� = Nvl(!����)
        str�����Ա� = Nvl(!�Ա�)
        str�������� = Nvl(!����)
        lng����ID = Val(Nvl(!����ID))
        strҽ�Ƹ��ʽ���� = Nvl(!���ʽ����)
        lngҽ����� = Val(Nvl(!ҽ�����))
        str�Ǽ�ʱ�� = Format(Nvl(!�Ǽ�ʱ��), "YYYYMMDDhhmmss000")
        strҵ�����Ա = Nvl(!����Ա����)
        
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!����ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!�վݷ�Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!�վݷ�Ŀ), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!������), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!�������), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!��Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!ҩƷ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!���), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!���㵥λ), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!�۸�)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!����)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!ʵ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!�Էѽ��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!Ӧ�ս��)), intJsonFormat_����С��), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!ҽ����Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!ҽ����������)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!��ע), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!���), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!��������), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dblƱ���ܽ�� = dblƱ���ܽ�� + RoundEx(Nvl(!ʵ�ս��), 6)
            .MoveNext
        Loop
        
        str����IDs = GetBalanceIDs(lng����ID, bytInvoiceType)
        dbl���� = GetBalanceErrorFee(str����IDs)
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
    End With
    
    '������ϸ
    If gBs_Type.���Ѷ��ձ��� <> "" Then
        dblƱ���ܽ�� = dblƱ���ܽ�� - dbl����
    End If
    dblƱ���ܽ�� = RoundEx(dblƱ���ܽ��, 2)
    If Not Get������ϸ(str����IDs, strData, dblƱ���ܽ��, bytInvoiceType, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    'Ʊ����Ϣ
    'ҵ����ˮ��:lngEInvoiceID_lng����ID
    strData = ""
    strData = strData & "" & GetJsonNodeString("busNo", lng����ID & "_" & lngEInvoiceID, Json_Text)
    strData = strData & "," & GetJsonNodeString("busType", strҵ���ʶ, Json_Text)
    strData = strData & "," & GetJsonNodeString("payer", str��������, Json_Text)
    strData = strData & "," & GetJsonNodeString("busDateTime", str�Ǽ�ʱ��, Json_Text)
    strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strData = strData & "," & GetJsonNodeString("payee", strҵ�����Ա, Json_Text)
    strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("totalAmt", dblƱ���ܽ��, Json_num)
    strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl����, 6) <> 0 And gBs_Type.���Ѷ��ձ��� = "", "����" & FormatEx(dbl����, 6) & "�����������", ""), Json_Text)
    strJson_Out = strData
    
    
    '�ƶ�֧��
    If Not Get�ƶ�֧����Ϣ(lng����ID, lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '֪ͨ��Ϣ
    If Not Get֪ͨ��Ϣ(lng����ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '������Ϣ
    Call Getҽ����Ϣ(bytInvoiceType, lng����ID, lng����ID, cllInsureInfo)
    Set rsTmp = Nothing
    If lngҽ����� <> 0 Then
        strSQL = "Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')) As ��������, Max(b.����) As ������ұ���," & vbNewLine & _
                "       Max(b.����) As �����������, Max(a.No) As ������" & vbNewLine & _
                "  From ���˹Һż�¼ A, ���ű� B" & vbNewLine & _
                "  Where a.ִ�в���id = b.Id And " & vbNewLine & _
                "   a.No = (Select Max(�Һŵ�) From ����ҽ����¼ Where ID = [1] Or ���id = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lngҽ�����)
        If rsTmp.RecordCount = 0 Then Set rsTmp = Nothing
    End If
    If rsTmp Is Nothing Then
        strSQL = "Select To_Char(a.����ʱ��, 'yyyy-mm-dd') As ��������, b.���� As ������ұ���," & vbNewLine & _
                "       b.���� As �����������, a.No As ������" & vbNewLine & _
                "  From ���˹Һż�¼ A, ���ű� B" & vbNewLine & _
                "  Where a.ִ�в���id = b.Id And " & vbNewLine & _
                "       a.Id = (Select ID" & vbNewLine & _
                "           From (Select ID, ����ʱ�� From ���˹Һż�¼ Where ����id = [1] Order By ����ʱ�� Desc)" & vbNewLine & _
                "           Where Rownum < 2)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng����ID)
    End If
    
    strData = ""
    strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("ҽ�ƻ�������"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_���ջ�������", "C"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareTypeCode", strҽ�Ƹ��ʽ����, Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalCareType", Getҽ�Ƹ��ʽ����(strҽ�Ƹ��ʽ����), Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_ҽ����", "C"), Json_Text)
    With rsTmp
        If .RecordCount > 0 Then
            strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!��������), strJsonFormat_��������), Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_�������, Nvl(!�����������), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientCategoryCode", Nvl(!������ұ���), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!������), Json_Text)
        Else
            strData = strData & "," & GetJsonNodeString("consultationDate", "", Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_�������, "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientCategoryCode", "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", lng����ID, Json_Text)
        End If
    End With
    strData = strData & "," & GetJsonNodeString("patientId", lng����ID, Json_Text)
    strData = strData & "," & GetJsonNodeString("sex", str�����Ա�, Json_Text)
    strData = strData & "," & GetJsonNodeString("age", str��������, Json_Text)
    
    strJson_Out = strJson_Out & "," & strData
    
    '֧����Ϣ
    If Not Get������Ϣ(str����IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '�ɷ�����
    If Not Get�ɷ�����(str����IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '����ҽ����Ϣ-����
    '������չ��Ϣ-����
    'eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    'isArrears  �Ƿ����ͨ  String  1  ��  0-��1-�ǣ���Ƿ���������ҽԺҵ��Ҫ���Ʊ���Ƿ����ͨ��
    'arrearsReason  ������ͨԭ��  String  200  ��  isArrears=0����д������ͨ��ԭ��
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '�շ���Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strChargeDetail
    '�嵥��Ŀ��ϸ
    strJson_Out = strJson_Out & "," & strListDetail
    
    '����������Json��
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceBySendCard = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function Get������ϸ(ByVal str����IDs As String, strChargeDetail As String, _
                ByVal dblƱ���ܽ�� As Double, _
                Optional ByVal sngҵ������ As Single, Optional strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������ϸ
    ' ��� : sngҵ������:1-�շ�;2-Ԥ��;3.1-�������;3.2-סԺ����,4-�Һ�;5-����
    ' ���� :
    ' ���� : chargeDetail�ڵ�����
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lng��� As Long
    Dim dbl�ϼƽ�� As Double, dbl���� As Double
    Dim strJsonList As String, strData As String, strTable As String
    On Error GoTo ErrHand
    
    dbl���� = 0
    '��˼��֧���շ���Ŀ���Ϊ0��Ҳ��֧���շ���ĿΪ��ʱ���ߵ���Ʊ��
    strTable = IIf(sngҵ������ = 3.2 Or sngҵ������ = 5, "סԺ���ü�¼", "������ü�¼")
    
    strSQL = "Select Rownum As ���, �վݷ�Ŀ����, �վݷ�Ŀ����, ����, ���㵥λ, Round(����, 2) As ����, Round(���ʽ��, 2) As ���ʽ��," & vbNewLine & _
            "        Round(�Էѽ��, 2) As �Էѽ��, ��ע, ���ʽ�� - Round(���ʽ��, 2) As ����" & vbNewLine & _
            "  From (Select /*+cardinality(b,10)*/" & vbNewLine & _
            "         Nvl(c.����, c1.����) As �վݷ�Ŀ����, Nvl(c.����, c1.����) As �վݷ�Ŀ����, 1 As ����, '' As ���㵥λ, Sum(a.���ʽ��) As ����, a.�վݷ�Ŀ," & vbNewLine & _
            "         Sum(a.���ʽ��) As ���ʽ��, Sum(a.���ʽ��) - Sum(a.ͳ����) As �Էѽ��, '' As ��ע" & vbNewLine & _
            "        From " & strTable & " A, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B, �վݷ�Ŀ���� C, �վݷ�Ŀ C1" & vbNewLine & _
            "        Where a.����id = b.Column_Value And a.�վݷ�Ŀ = c1.����(+) And a.�վݷ�Ŀ = c.�վݷ�Ŀ(+) and Decode(c.���ó���(+), 0, [2], c.���ó���(+)) = [2]" & vbNewLine & _
            "        Group By c.����, c1.����, c.����, c1.����, a.�վݷ�Ŀ" & _
            IIf(gBs_Type.����ÿ�Ʊ, "", " Having Sum(a.���ʽ��) <> 0") & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get������ϸ", str����IDs, IIf(sngҵ������ = 3.2, 2, 1))
    If rsTmp.EOF Then Exit Function
    
    strJsonList = ""
    With rsTmp
        Do While Not .EOF
            strData = ""
            lng��� = Val(Nvl(!���, 1))
            strData = strData & "" & GetJsonNodeString("sortNo", Nvl(!���, 1), Json_num)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!�վݷ�Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!�վݷ�Ŀ����), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!���㵥λ), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!����)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!����)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!���ʽ��)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!�Էѽ��)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!��ע), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dbl�ϼƽ�� = dbl�ϼƽ�� + RoundEx(Nvl(!���ʽ��), 2)
            .MoveNext
        Loop
        
        dbl���� = dblƱ���ܽ�� - dbl�ϼƽ��
        If RoundEx(dbl����, 6) <> 0 And gBs_Type.���Ѷ��ձ��� <> "" Then
            strData = ""
            strData = strData & "" & GetJsonNodeString("sortNo", lng��� + 1, Json_num)
            strData = strData & "," & GetJsonNodeString("chargeCode", gBs_Type.���Ѷ��ձ���, Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", gBs_Type.���Ѷ�������, Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(dbl����, 2), Json_num)
            strData = strData & "," & GetJsonNodeString("number", 1, Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(dbl����, 2), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(dbl����, 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", "", Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
        End If
        
        strChargeDetail = Mid(strJsonList, 2)
    End With
    
    Get������ϸ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get֪ͨ��Ϣ(ByVal lng����ID As Long, strNotice As String, Optional str����� As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ֪ͨ��Ϣ
    ' ��� :
    ' ���� :
    ' ���� : Collect����Ա(����,ҽ����,���ջ�������,��������)
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    
    If lng����ID <> 0 Then
        strSQL = "Select Max(a.����id) As ����id, Max(a.����) As ����, Max(a.�ֻ���) As �ֻ���, Max(a.Email) As Email, Max(1) As �ɿ�����," & vbNewLine & _
                "      Max(a.���֤��) As ���֤��, Max(m.����) As ����, Max(a.�����) As �����" & vbNewLine & _
                "From ������Ϣ A," & vbNewLine & _
                "    (" & vbNewLine & _
                "      Select ����id, ����" & vbNewLine & _
                "      From (Select b.����id, b.����" & vbNewLine & _
                "              From ����ҽ�ƿ���Ϣ B" & vbNewLine & _
                "              Where b.�����id = [2] And b.����id = [1])" & vbNewLine & _
                "      Where Rownum < 2) M" & vbNewLine & _
                "Where a.����id = m.����id(+) And a.����id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get֪ͨ��Ϣ", lng����ID, gBs_Type.ȱʡ�����ID)
        
        strNotice = ""
        If rsTmp.RecordCount > 0 Then
            With rsTmp
                strNotice = strNotice & "" & GetJsonNodeString("tel", Nvl(!�ֻ���), Json_Text)
                strNotice = strNotice & "," & GetJsonNodeString("email", Nvl(!email), Json_Text)
                If gBs_Type.֧�ְ汾 > BS_Version.V2_0_3 Then
                    strNotice = strNotice & "," & GetJsonNodeString("payerType", Nvl(!�ɿ�����), Json_Text)
                End If
                strNotice = strNotice & "," & GetJsonNodeString("idCardNo", Nvl(!���֤��), Json_Text)
                If Nvl(!����) <> "" Then
                    strNotice = strNotice & "," & GetJsonNodeString("cardType", gBs_Type.ҽ�ƿ����ͱ��, Json_Text)
                    strNotice = strNotice & "," & GetJsonNodeString("cardNo", Nvl(!����), Json_Text)
                ElseIf Nvl(!���֤��) <> "" And gBs_Type.���֤�������ͱ�� <> "" Then
                    strNotice = strNotice & "," & GetJsonNodeString("cardType", gBs_Type.���֤�������ͱ��, Json_Text)
                    strNotice = strNotice & "," & GetJsonNodeString("cardNo", Nvl(!���֤��), Json_Text)
                Else
                    strNotice = strNotice & "," & GetJsonNodeString("cardType", gBs_Type.�����޿��Ŀ������, Json_Text)
                    strNotice = strNotice & "," & GetJsonNodeString("cardNo", gBs_Type.�����޿��Ŀ���, Json_Text)
                End If
                str����� = Nvl(!�����)
            End With
        End If
    End If
    Get֪ͨ��Ϣ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get֪ͨ����(ByVal lng����ID As Long, ByVal str���֤�� As String, str������ As String, str���� As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ֪ͨ����
    ' ��� :
    ' ���� :
    ' ���� : strM_Payment���ƶ�֧����Ϣ
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    
    If lng����ID <> 0 Then
        strSQL = "Select ����" & vbNewLine & _
                "From (Select b.����id, b.����" & vbNewLine & _
                "       From ����ҽ�ƿ���Ϣ B " & vbNewLine & _
                "       Where b.�����id = [2] And b.����id = [1])" & vbNewLine & _
                "Where Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get֪ͨ����", lng����ID, gBs_Type.ȱʡ�����ID)
        If rsTmp.RecordCount > 0 Then
            str������ = gBs_Type.ҽ�ƿ����ͱ��
            str���� = Nvl(rsTmp!����)
        End If
    End If
    If str������ = "" Then
        If str���֤�� <> "" And gBs_Type.���֤�������ͱ�� <> "" Then
            str������ = gBs_Type.���֤�������ͱ��
            str���� = str���֤��
        Else
            str������ = gBs_Type.�����޿��Ŀ������
            str���� = gBs_Type.�����޿��Ŀ���
        End If
    End If
    Get֪ͨ���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get�ƶ�֧����Ϣ(ByVal lng����ID As Long, ByVal str����IDs As String, strM_Payment As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�ƶ�֧����Ϣ
    ' ��� :
    ' ���� :
    ' ���� : strM_Payment���ƶ�֧����Ϣ
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    
    strSQL = "Select Max(Decode(��Ϣ��, '������', ��Ϣֵ, '֧��������', ��Ϣֵ, '')) As ֧��������," & vbNewLine & _
            "        Max(Decode(��Ϣ��, 'ҽ��֧��������', ��Ϣֵ, 'ҽ��������', ��Ϣֵ, '')) As ҽ��֧��������," & vbNewLine & _
            "        Max(Decode(Upper(��Ϣ��), '֧�������ں�USERID', ��Ϣֵ, '')) As ֧�������ں�userid," & vbNewLine & _
            "        Max(Decode(Upper(��Ϣ��), '֧����С����USERID', ��Ϣֵ, '')) As ֧����С����userid," & vbNewLine & _
            "        Max(Decode(Upper(��Ϣ��), '΢�Ź��ں�OPENID', ��Ϣֵ, '')) As ΢�Ź��ں�openid," & vbNewLine & _
            "        Max(Decode(Upper(��Ϣ��), '΢��С����OPENID', ��Ϣֵ, '')) As ΢��С����openid" & vbNewLine & _
            " From (Select ��Ϣ��, ��Ϣֵ" & vbNewLine & _
            "        From ������Ϣ�ӱ�" & vbNewLine & _
            "        Where ����id = [1] And ��Ϣ�� In ('֧�������ں�USERID', '֧����С����USERID', '΢�Ź��ں�OPENID', '΢��С����OPENID')" & vbNewLine & _
            "        Union All" & vbNewLine & _
            "        Select ������Ŀ, ��������" & vbNewLine & _
            "        From �������㽻��" & vbNewLine & _
            "        Where ����id In (Select ID From ����Ԥ����¼ a, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B Where a.����id = b.Column_Value) And ������Ŀ Like '%������')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get�ƶ�֧����Ϣ", lng����ID, str����IDs)
    
    strM_Payment = ""
    If rsTmp.RecordCount > 0 Then
        With rsTmp
            strM_Payment = strM_Payment & "" & GetJsonNodeString("alipayCode", Nvl(!֧�������ں�userid), Json_Text)
            strM_Payment = strM_Payment & "," & GetJsonNodeString("weChatOrderNo", Nvl(!֧��������), Json_Text)
            If gBs_Type.֧�ְ汾 > BS_Version.V2_0_3 Then
                strM_Payment = strM_Payment & "," & GetJsonNodeString("weChatMedTransNo", Nvl(!ҽ��֧��������), Json_Text)
            End If
            If Nvl(!΢�Ź��ں�openid) <> "" Then
                strM_Payment = strM_Payment & "," & GetJsonNodeString("openID", Nvl(!΢�Ź��ں�openid), Json_Text)
            Else
                strM_Payment = strM_Payment & "," & GetJsonNodeString("openID", Nvl(!΢��С����openid), Json_Text)
            End If
        End With
    End If
    Get�ƶ�֧����Ϣ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Getҽ����Ϣ(ByVal byt���� As Byte, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                cllInsureInfo_Out As Collection, Optional blnסԺ���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡҽ����Ϣ
    ' ��� :
    ' ���� :
    ' ���� : Collect����Ա(����,ҽ����,���ջ�������,��������)
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    Set cllInsureInfo_Out = New Collection
    
    strSQL = "Select Max(a.����) As ����, Max(b.���ջ�������) As ���ջ�������, Max(Nvl(a.��������, c.����)) As ��������" & vbNewLine & _
            "  From ���ս����¼ A, ������� B, ���ղ��� C" & vbNewLine & _
            "  Where a.���� = b.��� And a.����id = c.Id(+) And a.��¼id = [2] And a.���� = Decode([1], 2, 3, 3, 2, 1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Getҽ����Ϣ", byt����, lng����ID)
    With rsTmp
        If .RecordCount > 0 Then
            cllInsureInfo_Out.Add Nvl(!����), "_����"
            cllInsureInfo_Out.Add Nvl(!���ջ�������), "_���ջ�������"
            cllInsureInfo_Out.Add Nvl(!��������), "_��������"
        End If
    End With
    
    If cllInsureInfo_Out.Count > 0 Then
        If Val(cllInsureInfo_Out("_����")) <> 0 Then
            strSQL = "Select Max(ҽ����) As ҽ���� From �����ʻ� Where ����id = [1] And ���� = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Getҽ����Ϣ", lng����ID, Val(cllInsureInfo_Out("_����")))
            With rsTmp
                If .RecordCount > 0 Then
                    cllInsureInfo_Out.Add Nvl(!ҽ����), "_ҽ����"
                End If
            End With
            
            If cllInsureInfo_Out("_��������") = "" And Not blnסԺ���� Then
                strSQL = "Select Max(��������) As ��������" & vbNewLine & _
                        "      From (Select Distinct a.���� As ��������" & vbNewLine & _
                        "             From ���ղ��� A, ������׼��Ŀ B" & vbNewLine & _
                        "             Where a.���� = [1] And a.Id = b.����id And" & vbNewLine & _
                        "                   b.�շ�ϸĿid In (Select Distinct �շ�ϸĿid From ������ü�¼ Where ����id = [2])" & vbNewLine & _
                        "             Union All" & vbNewLine & _
                        "             Select Distinct a.���� As ��������" & vbNewLine & _
                        "             From ���ղ��� A, ������׼��Ŀ B" & vbNewLine & _
                        "             Where a.���� = [1] And a.Id = b.����id And" & vbNewLine & _
                        "                   b.���� In (Select Distinct ���մ���id From ������ü�¼ Where ����id = [2]))" & vbNewLine & _
                        "      Where Rownum < 2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Getҽ����Ϣ", Val(cllInsureInfo_Out("_����")), lng����ID)
                If rsTmp.RecordCount > 0 Then
                    cllInsureInfo_Out.Remove "_��������"
                    cllInsureInfo_Out.Add Nvl(rsTmp!��������), "_��������"
                End If
            End If
        End If
    End If
    Getҽ����Ϣ = True
    Exit Function
ErrHand:
    Set cllInsureInfo_Out = New Collection
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get������Ϣ(ByVal str����IDs As String, strPayment As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������Ϣ
    ' ��� :
    ' ���� :
    ' ���� : strM_Payment���ƶ�֧����Ϣ
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strPayment = ""
    
    '�����֧�����Ѵ��������ۼƵ������ֽ�֧��
    strSQL = "Select �ֽ�Ԥ��, ֧ƱԤ��, ת��Ԥ��, �����ʻ�֧��, ҽ��ͳ�����֧��, ����ҽ��֧��, �����ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, �ֽ�֧��, 0) As �ֽ��˿�," & vbNewLine & _
            "      Decode(Sign(֧Ʊ֧��), -1, ֧Ʊ֧��, 0) As ֧Ʊ�˿�, Decode(Sign(ת��֧��), -1, ת��֧��, 0) As ת���˿�," & vbNewLine & _
            "      Decode(Sign(�ֽ�֧��), -1, 0, �ֽ�֧��) As �ֽ�֧��, Decode(Sign(֧Ʊ֧��), -1, 0, ֧Ʊ֧��) As ֧Ʊ֧��," & vbNewLine & _
            "      Decode(Sign(ת��֧��), -1, 0, ת��֧��) As ת��֧��," & vbNewLine & _
            "      Nvl(�����ʻ�֧��, 0) + Nvl(ҽ��ͳ�����֧��, 0) + Nvl(����ҽ��֧��, 0) As �����ܶ�," & vbNewLine & _
            "      Nvl(�����ܶ�, 0) - Nvl(�����ʻ�֧��, 0) - Nvl(ҽ��ͳ�����֧��, 0) - Nvl(����ҽ��֧��, 0) As �Էѽ��, �����ܶ�, ҽ���������," & vbNewLine & _
            "      0 As �����ʻ����" & vbNewLine & _
            "From (Select /*+cardinality(b,10)*/" & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '�ֽ�', 1, 0), 0) * a.��Ԥ��) As �ֽ�Ԥ��," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0), 0) * a.��Ԥ��) As ֧ƱԤ��," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 0, '�ֽ�', 0, 1), 0) * a.��Ԥ��) As ת��Ԥ��," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '�����˻�֧��', 1, 0)) * a.��Ԥ��) As �����ʻ�֧��," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, 'ҽ��ͳ�����֧��', 1, 0)) * a.��Ԥ��) As ҽ��ͳ�����֧��," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 1, 0)) * a.��Ԥ��) As ����ҽ��֧��," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 0, '�����˻�֧��', 0, 'ҽ��ͳ�����֧��', 0, 1)) *" & vbNewLine & _
            IIf(gBs_Type.���Ѷ��ձ��� = "", "", " Decode(D.����, 9, 0, 1) * ") & " a.��Ԥ��) As �����ֽ�֧��," & vbNewLine & _
            "       Max(Decode(Mod(a.��¼����, 10), 1, 0," & vbNewLine & _
            "                   Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', �������, '�����˻�֧��', �������, 'ҽ��ͳ�����֧��', �������, ''))) As ҽ���������," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, Null, Decode(a.���㷽ʽ, '�ֽ�', 1, 0), 0)) * a.��Ԥ��) As �ֽ�֧��," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, Null, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0), 0)) * a.��Ԥ��) As ֧Ʊ֧��," & vbNewLine & _
            "       Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, Null, Decode(a.���㷽ʽ, '�ֽ�', 0, '֧Ʊ', 0, 1), 0)) * Decode(D.����, 9, 0, 1) * a.��Ԥ��) As ת��֧��," & vbNewLine & _
            "       Sum(��Ԥ��) As �����ܶ�" & vbNewLine & _
            "      From ����Ԥ����¼ A, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B, ��Ʊ������� C, ���㷽ʽ D" & vbNewLine & _
            "      Where a.����id = b.Column_Value And a.���㷽ʽ = c.���㷽ʽ(+) and a.���㷽ʽ = d.����(+))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get������Ϣ", str����IDs)
    
    With rsTmp
        If .RecordCount > 0 Then
            strPayment = strPayment & "" & GetJsonNodeString("accountPay", FormatEx(Val(Nvl(!�����ʻ�֧��)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("fundPay", FormatEx(Val(Nvl(!ҽ��ͳ�����֧��)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("otherfundPay", FormatEx(Val(Nvl(!����ҽ��֧��)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("ownPay", FormatEx(Val(Nvl(!�Էѽ��)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("selfConceitedAmt", 0, Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("selfPayAmt", 0, Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("selfCashPay", FormatEx(Val(Nvl(!�����ֽ�֧��)), 6), Json_num)
            If gBs_Type.֧�ְ汾 > V3_1_0 Then
                strPayment = strPayment & "," & GetJsonNodeString("cashPay", FormatEx(Val(Nvl(!�ֽ�Ԥ��)) + Val(Nvl(!֧ƱԤ��)) + Val(Nvl(!ת��Ԥ��)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("cashRecharge", FormatEx(Val(Nvl(!�ֽ�֧��)) + Val(Nvl(!֧Ʊ֧��)) + Val(Nvl(!ת��֧��)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("cashRefund", FormatEx(Val(Nvl(!�ֽ��˿�)) + Val(Nvl(!֧Ʊ�˿�)) + Val(Nvl(!ת���˿�)), 6), Json_num)
            Else
                strPayment = strPayment & "," & GetJsonNodeString("cashPay", FormatEx(Val(Nvl(!�ֽ�Ԥ��)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("chequePay", FormatEx(Val(Nvl(!֧ƱԤ��)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("transferAccountPay", FormatEx(Val(Nvl(!ת��Ԥ��)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("cashRecharge", FormatEx(Val(Nvl(!�ֽ�֧��)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("chequeRecharge", FormatEx(Val(Nvl(!֧Ʊ֧��)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("transferRecharge", FormatEx(Val(Nvl(!ת��֧��)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("cashRefund", FormatEx(Val(Nvl(!�ֽ��˿�)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("chequeRefund", FormatEx(Val(Nvl(!֧Ʊ�˿�)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("transferRefund", FormatEx(Val(Nvl(!ת���˿�)), 6), Json_num)
            End If
            strPayment = strPayment & "," & GetJsonNodeString("ownAcBalance", FormatEx(Val(Nvl(!�����ʻ����)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("reimbursementAmt", FormatEx(Val(Nvl(!�����ܶ�)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("balancedNumber", Nvl(!ҽ���������), Json_Text)
        End If
    End With
    Get������Ϣ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get�ɷ�����(ByVal str����IDs As String, strPayChannelInfo As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �ɷ�������Ϣ
    ' ��� :
    ' ���� :
    ' ���� : strPayChannelInfo���ɷ�������Ϣ
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:19
    ' ˵�� : Ӧ�ó��ϣ��շѡ��Һ�
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strJsonList As String, strData As String
    On Error GoTo ErrHand
    strPayChannelInfo = ""
    
    strSQL = "Select /*+cardinality(b,10)*/" & vbNewLine & _
            "      Nvl(c.��������, Nvl(d.��������, '-')) As ��������, Sum(��Ԥ��) As �����ܶ�" & vbNewLine & _
            "     From ����Ԥ����¼ A, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B, �շ��������� C," & vbNewLine & _
            "          (Select ���㷽ʽ, �������� From �շ��������� D Where �����id Is Null) D" & vbNewLine & _
            "     Where a.����id = b.Column_Value And a.�����id = c.�����id(+) And a.���㷽ʽ = c.���㷽ʽ(+) And a.���㷽ʽ = d.���㷽ʽ(+)" & vbNewLine & _
            "     Group By Nvl(c.��������, Nvl(d.��������, '-'))" & vbNewLine & _
            "     Order By ��������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get�ɷ�����", str����IDs)
    
    With rsTmp
        Do While Not .EOF
            strData = ""
            If Nvl(!��������) <> "-" Then
                strData = strData & "" & GetJsonNodeString("payChannelCode", Nvl(!��������), Json_Text)
                strData = strData & "," & GetJsonNodeString("payChannelValue", FormatEx(Nvl(!�����ܶ�), 6), Json_num)
                strJsonList = strJsonList & ",{" & strData & "}"
            End If
            .MoveNext
        Loop
        strPayChannelInfo = Mid(strJsonList, 2)
    End With
    
    Get�ɷ����� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Getҽ�Ƹ��ʽ����(ByVal strҽ�Ƹ��ʽ���� As String) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ����ҽ�Ƹ��ʽ�����ȡ����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    If strҽ�Ƹ��ʽ���� = "" Then Exit Function
    strSQL = "Select Max(����) as ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Getҽ�Ƹ��ʽ����", strҽ�Ƹ��ʽ����)
    If rsTmp.RecordCount > 0 Then
        Getҽ�Ƹ��ʽ���� = Nvl(rsTmp!����)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Getҽ�Ƹ��ʽ����(ByVal strҽ�Ƹ��ʽ As String) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ����ҽ�Ƹ��ʽ�����ȡ����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    If strҽ�Ƹ��ʽ = "" Then Exit Function
    strSQL = "Select Max(����) as ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Getҽ�Ƹ��ʽ����", strҽ�Ƹ��ʽ)
    If rsTmp.RecordCount > 0 Then
        Getҽ�Ƹ��ʽ���� = Nvl(rsTmp!����)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetԤ�������ܶ�(ByVal strNO As String) As Double
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ����Ԥ�����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    strSQL = "Select Sum(���) As Ʊ���ܽ��" & vbNewLine & _
            "  From (Select Sum(���) As ���" & vbNewLine & _
            "         From ����Ԥ����¼" & vbNewLine & _
            "         Where NO = [1] And ��¼���� = 1" & vbNewLine & _
            "         Union All" & vbNewLine & _
            "         Select Sum(��Ԥ��) As ���" & vbNewLine & _
            "         From ����Ԥ����¼" & vbNewLine & _
            "         Where ����id In (Select Distinct ����id From ����Ԥ����¼ Where NO = [1] And Mod(��¼����, 10) = 1) And" & vbNewLine & _
            "               Nvl(���, 0) < 0 And Mod(��¼����, 10) = 1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetԤ�����", strNO)
    If rsTmp.RecordCount > 0 Then
        GetԤ�������ܶ� = Val(Nvl(rsTmp!Ʊ���ܽ��))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetԤ�����(ByVal lng����ID As Long, ByVal intԤ������ As Integer) As Double
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ����Ԥ�����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    strSQL = "Select Max(Ԥ�����) As Ԥ����� From ������� " & _
            " Where ����id = [1] And ���� = 1 And ���� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetԤ�����", lng����ID, intԤ������)
    If rsTmp.RecordCount > 0 Then
        GetԤ����� = Val(Nvl(rsTmp!Ԥ�����))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetEInvoiceInfo(ByVal lngEInvoiceID As Long, strErrMsg_Out As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    strSQL = "Select ��¼״̬, Ʊ��, ���� As Ʊ�ݴ���, ���� As Ʊ�ݺ���, ������ As Ʊ��У����, ����ʱ�� From ����Ʊ��ʹ�ü�¼ Where Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetEInvoiceInfo", lngEInvoiceID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ�����Ʊ��ʹ�ü�¼�����顣": Exit Function
    End If
    Set GetEInvoiceInfo = rsTmp
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function GetEInvoiceWithPatiInfo(ByVal lngEInvoiceID As Long, strErrMsg_Out As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    strSQL = "Select a.��¼״̬, a.Ʊ��, a.���� As Ʊ�ݴ���, a.���� As Ʊ�ݺ���, a.������ As Ʊ��У����, b.�ֻ���, b.email," & vbNewLine & _
            "        a.�Ƿ񻻿� " & vbNewLine & _
            "From ����Ʊ��ʹ�ü�¼ a, ������Ϣ b" & vbNewLine & _
            "Where a.Id =[1] And a.����id = b.����id(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetEInvoiceInfo", lngEInvoiceID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ�����Ʊ��ʹ�ü�¼�����顣": Exit Function
    End If
    Set GetEInvoiceWithPatiInfo = rsTmp
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function CheckBillExistReplenishData(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ǲ��ǲ��������
    '����:True-�ǲ����� False-��֮
    '���:lng����ID-����ID
    '����:���ϴ�
    '����:2020-4-30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strSQL = "Select 1 From ���ò����¼ A where ����ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", lng����ID)
    If rsTmp.EOF Then
        CheckBillExistReplenishData = False
    Else
        CheckBillExistReplenishData = True
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceErrorFee(ByVal str����ID As String) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '����:
    '���:str����ID-����IDs
    '����:���ϴ�
    '����:2020-4-30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strSQL = "Select /*+cardinality(c,10)*/ Sum(a.��Ԥ��) as ��Ԥ�� From ����Ԥ����¼ A, ���㷽ʽ B, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C " & _
            " where a.����id = c.Column_Value and a.���㷽ʽ = b.����(+) and b.���� = 9"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", str����ID)
    If Not rsTmp.EOF Then
        GetBalanceErrorFee = Val(Nvl(rsTmp!��Ԥ��))
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceIDs(ByVal str����ID As String, Optional ByVal intҵ������ As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ԭ����ID��ȡ�����漰�����н���ID�ͳ���ID
    '����:
    '���:lng����ID-����ID,1-�շ�;4-�Һ�;5-����
    '����:���ϴ�
    '����:2020-4-30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str����IDs As String, strTable As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    
    strTable = IIf(intҵ������ = 5, "סԺ���ü�¼", "������ü�¼")
    strSQL = "Select Distinct ����ID From " & strTable & _
            " Where (no, ��¼����) In (Select No, ��¼���� From " & strTable & " Where ����ID = [1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", str����ID)
    With rsTmp
        Do While Not .EOF
            str����IDs = str����IDs & "," & Nvl(!����ID)
            .MoveNext
        Loop
    End With
    GetBalanceIDs = Mid(str����IDs, 2)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

    
Public Function GetPaperCode(ByVal bytInvoiceType As Byte) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡֽ��Ʊ�ݴ���
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/6/28 17:41
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    GetPaperCode = Decode(bytInvoiceType, 2, gBs_Type.Ԥ��ֽ��Ʊ�ݴ���, 4, gBs_Type.�Һ�ֽ��Ʊ�ݴ���, 3, gBs_Type.����ֽ��Ʊ�ݴ���, gBs_Type.�շ�ֽ��Ʊ�ݴ���)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
