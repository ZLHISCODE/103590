Attribute VB_Name = "mdlProcessData"
Option Explicit

Public Function ProcDrugInfo(ByVal strDrugType As String, ByVal objDevice As clsDevice) As ADODB.Recordset
'���ܣ���ȡHISҩƷ������Ϣ
'������
'  strDrugType�����ʹ�
'  objDevice���豸����
'���أ��Ѹ�ʽ���ļ�¼��

    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsUpload As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '��HIS����
    Set rsData = mdlDefine.GetHisRecord_DrugInf(1, strDrugType)
    
    '��ʽ��Ҫ�ϴ�������
    Set rsUpload = BuildDrugInfo(rsData, objDevice)
    
    If Not rsUpload Is Nothing Then
        Set ProcDrugInfo = rsUpload
    End If
    
    Exit Function

errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Function

Public Function ProcDrugStock(ByVal lngDeptID As Long, ByVal objDevice As clsDevice) As ADODB.Recordset
'���ܣ���ȡHISҩƷ�����Ϣ
'������
'  lngDeptID��ҩ��ID
'  objDevice���豸����
'���أ��Ѹ�ʽ���ļ�¼��
    
    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsUpload As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '��HISҩƷ�������
    Set rsData = mdlDefine.GetHisRecord_DrugStock(lngDeptID)
    
    '��ʽ��Ҫ�ϴ�������
    Set rsUpload = BuildDrugStock(rsData, objDevice)
    
    If Not rsUpload Is Nothing Then
        Set ProcDrugStock = rsUpload
    End If
    
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Function

Public Function SetUpload(ByVal bytType As Byte, ByVal varKey As Variant, ByVal lngModule As Long) As ADODB.Recordset
'���ܣ���ȡHIS����ϴ���Ϣ
'������
'   bytType��
'       1: ���ﴦ���ϴ� (��ҩ)
'       2: ���﷢ҩ֪ͨ (��ҩ)
'       3: סԺҩƷҽ���ϴ� (�䡢��ҩ)
'   varKey��
'       ��bytType=1ʱ��varKey��ʾ������;�ⷿID;NO����
'       ��ʽ��������;�ⷿID;NO[|����;�ⷿID;NO][|...]��
'       ��bytType=2ʱ��ͬbytType=1
'       ��bytType=3ʱ��varKey��ʾҩƷ�շ�ID��
'       ��ʽ����ҩƷ�շ�ID[,ҩƷ�շ�ID][,...]��
'  lngModule��HISҵ��ģ���
'���أ��Ѹ�ʽ���ļ�¼��

    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsUpload As ADODB.Recordset
    'Dim arrBill As Variant
    'Dim i As Integer

    '��HIS����
    Select Case bytType
    Case 1
        '���ﴦ����ϸ
        Set rsData = mdlDefine.GetHisRecord_ReceipInf(varKey)
        '��ʽ��Ҫ�ϴ�������
        Set rsUpload = BuildReceipDetail(rsData, lngModule)
        
    Case 2
        '���﷢ҩ֪ͨ
        Set rsData = mdlDefine.GetHisRecord_ReceipList(varKey)
        '��ʽ��Ҫ�ϴ�������
        Set rsUpload = BuildReceipList(rsData, lngModule)
        
    Case 3
        'סԺҩƷҽ��
        Set rsData = mdlDefine.GetHisRecord_AdviceInf(varKey)
        '��ʽ��Ҫ�ϴ�������
        Set rsUpload = BuildReceipAdviceInf(rsData, lngModule)
        
    End Select
    
    If Not rsData Is Nothing Then
        Set SetUpload = rsUpload
    End If
End Function

Private Function BuildDrugInfo(ByVal rsDrugInfo As ADODB.Recordset, ByVal objDevice As clsDevice) As ADODB.Recordset
'���ܣ���������ҩƷ��Ϣ�ϴ����ݽṹ�ļ�¼������
'������
'  rsDrugInfo��HISҩƷ��Ϣ��¼������
'  objDevice���豸����

    Const STR_ROOT = "ROOT"
    Const STR_NODE = "CONSIS_BASIC_DRUGSVW"
    
    Dim i As Integer
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    
    If rsDrugInfo Is Nothing Then Exit Function
    
    '��ʼ���ڴ��¼������
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "SN", adInteger, 10, adFldIsNullable
        .Fields.Append "Drug", adVarChar, 100, adFldIsNullable
        .Fields.Append "Content", adVarChar, 3000, adFldIsNullable
        .Open
    End With
    
    With rsDrugInfo
        If .State <> adStateOpen Then .Open
        i = 1
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            '��ʽ����Ҫ�ϴ������ݸ�ʽ
            Select Case objDevice.LinkType
            Case enuLinkType.DB
                strTmp = "delete atf_his_druginfo where drug_code='" & !���� & "' and drugname='" & !ͨ���� & "' " & Chr(13) _
                       & "insert into atf_his_druginfo (drug_code,drugname,specification,drug_type," _
                       & "dosage,dos_unit,pack_amount,pack_name,manufactory,py_code,manu_no) " & Chr(13)
                strTmp = strTmp _
                       & "select '" & !���� & "'," _
                       & "'" & !ͨ���� & "'," _
                       & "'" & NVL(!���) & "'," _
                       & "'" & !���� & "'," _
                       & CDbl(!����ϵ��) & "," _
                       & "'" & !������λ & "'," _
                       & CDbl(!סԺ��װ) & "," _
                       & "'" & !סԺ��λ & "'," _
                       & "'" & IIf(IsNull(!����������), "", !����������) & "'," _
                       & "'" & !ƴ������ & "'," _
                       & "'" & IIf(IsNull(!�����̱��), "", !�����̱��) & "' "
            
            Case enuLinkType.WEBServices
                strTmp = "<" & STR_ROOT & ">"
                
                strTmp = strTmp & vbCrLf & "<" & STR_NODE
                strTmp = strTmp & vbCrLf & "DRUG_CODE = """ & SpecialChar(!���) & """"
                strTmp = strTmp & vbCrLf & "DRUG_NAME = """ & SpecialChar(!ͨ����) & """"
                strTmp = strTmp & vbCrLf & "TRADE_NAME = """ & SpecialChar(!��Ʒ��) & """"
                strTmp = strTmp & vbCrLf & "DRUG_SPEC = """ & SpecialChar(!���) & """"
                strTmp = strTmp & vbCrLf & "DRUG_PACKAGE = """ & NVL(!�����װ) & """"
                strTmp = strTmp & vbCrLf & "DRUG_UNIT = """ & SpecialChar(!���ﵥλ) & """"
                strTmp = strTmp & vbCrLf & "FIRM_ID = """ & SpecialChar(!����������) & """"
                If objDevice.ServiceObject = 1 Then
                    '����
                    strTmp = strTmp & vbCrLf & "DRUG_PRICE = """ & NVL(!�ۼ�) * NVL(!�����װ) & """"
                    strTmp = strTmp & vbCrLf & "DRUG_CONVERTATION = """ & Round(NVL(!ҩ���װ) / NVL(!�����װ), 2) & """"
                Else
                    'סԺ
                    strTmp = strTmp & vbCrLf & "DRUG_PRICE = """ & NVL(!�ۼ�) * NVL(!סԺ��װ) & """"
                    strTmp = strTmp & vbCrLf & "DRUG_CONVERTATION = """ & Round(NVL(!ҩ���װ) / NVL(!סԺ��װ), 2) & """"
                End If
                strTmp = strTmp & vbCrLf & "DRUG_FORM = """ & SpecialChar(!����) & """"
                strTmp = strTmp & vbCrLf & "DRUG_SORT = """ & SpecialChar(!�������) & """"
                strTmp = strTmp & vbCrLf & "BARCODE = """""
                strTmp = strTmp & vbCrLf & "LAST_DATE = """ & Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
                strTmp = strTmp & vbCrLf & "PINYIN = """ & SpecialChar(!ƴ������) & """"
                strTmp = strTmp & ">"
                
                strTmp = strTmp & "</" & STR_NODE & ">"
                strTmp = strTmp & "</" & STR_ROOT & ">"
                
            Case enuLinkType.Directory
                strTmp = ""
            End Select
            
            '�����ڴ��¼��
            If strTmp <> "" Then
                rsData.AddNew
                rsData!SN = i
                rsData!Drug = !���� & "��" & !ͨ���� & "��" & NVL(!���)
                rsData!Content = strTmp
                rsData.Update
                i = i + 1
            End If
            
            .MoveNext
        Loop
        .Close
        
    End With
    Set BuildDrugInfo = rsData
    
End Function

Private Function BuildReceipDetail(ByVal rsVal As ADODB.Recordset, ByVal lngModule As Long) As ADODB.Recordset
'���ܣ������������ﴦ����ϸ(��ҩ)�ϴ����ݽṹ�ļ�¼������
'������
'  rsVal��HIS���ﴦ����ϸ��¼������
'  lngModule��HISҵ��ģ���
    
    Const STR_ROOT = "ROOT"
    Const STR_NODE_T = "CONSIS_PRESC_MSTVW"
    Const STR_NODE_D = "CONSIS_PRESC_DTLVW"
    
    Dim rsData As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strTitle As String, strDetail As String
    Dim lng�ⷿID As Long
    Dim int���� As Integer
    Dim strNO As String
    Dim curӦ�ս�� As Currency, curʵ�ս�� As Currency
    Dim lngDeviceID As Long
    Dim strTmp As String
    Dim bytType As Byte
    
    If rsVal Is Nothing Then Exit Function
    
    '��ʼ���ڴ��¼������
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "DeviceID", adBigInt, , adFldIsNullable
        .Fields.Append "NO", adVarChar, 20, adFldIsNullable
        .Fields.Append "����", adInteger, , adFldIsNullable
        .Fields.Append "Content", adLongVarChar, 20000, adFldIsNullable
        .Open
    End With
    
    With rsTmp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "NO", adVarChar, 20, adFldIsNullable
        .Fields.Append "����", adInteger, , adFldIsNullable
        .Fields.Append "�ⷿID", adBigInt, , adFldIsNullable
        .Fields.Append "DeviceID", adBigInt, , adFldIsNullable
        .Fields.Append "Type", adInteger, 1, adFldIsNullable
        .Fields.Append "Content", adVarChar, 2000, adFldIsNullable
        .Fields.Append "Ӧ�ս��", adCurrency, , adFldIsNullable
        .Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
        .Open
    End With
    
    With rsVal
        If .State <> adStateOpen Then .Open
        i = 1: curӦ�ս�� = 0: curʵ�ս�� = 0: strDetail = ""
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False

            '�����������豸ID
            lngDeviceID = GetDevice(1, !��ҩҩ��id, !ҩƷ����)
            
            If lngDeviceID <= 0 Then GoTo makLoop
            
            bytType = GetDeviceType(lngDeviceID)
            
            '��ϸ��Ϣ
            strDetail = ""
            Select Case bytType
            Case enuLinkType.WEBServices
                strDetail = "<" & STR_NODE_D
                strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
                strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(!NO) & """"
                strDetail = strDetail & vbCrLf & "ITEM_NO = """ & i & """"
                strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(!ҩƷid) & """"
                strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(!ҩƷ��Ʒ��) & """"
                strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(!ҩƷ���) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(!ҩƷ���) & """"
                strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(!���ﵥλ) & """"
                strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(!������) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(!�ۼ�) & """"
                strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(!����) & """"
                strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(!Ӧ�ս��) & """"
                strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(!ʵ�ս��) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(!����) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(!������λ) & """"
                strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(!ҩƷ�÷�) & """"
                strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(!Ƶ��) & """"
                strDetail = strDetail & vbCrLf & ">"
                strDetail = strDetail & vbCrLf & "</" & STR_NODE_D & ">"
                
            End Select
            
            'д�룬��rsData��¼����ʹ��
            If strDetail <> "" Then
                rsTmp.AddNew
                rsTmp!NO = !NO
                rsTmp!���� = !����
                rsTmp!�ⷿID = !��ҩҩ��id
                rsTmp!DeviceID = lngDeviceID
                rsTmp!Type = bytType
                rsTmp!Content = strDetail
                rsTmp!Ӧ�ս�� = NVL(!Ӧ�ս��, 0)
                rsTmp!ʵ�ս�� = NVL(!ʵ�ս��, 0)
                rsTmp.Update
            End If
            
            i = i + 1
            int���� = !����: strNO = !NO: lng�ⷿID = !��ҩҩ��id
            
            .MoveNext
            If .EOF Then
                GoTo makCommon1
            ElseIf int���� <> !���� And strNO <> !NO And lng�ⷿID <> !��ҩҩ��id Then
makCommon1:
                .MovePrevious
                i = 1
            End If
            
makLoop:
            .MoveNext
        Loop
    End With
    
    '���������ı��������¼��
    With rsTmp
        curӦ�ս�� = 0
        curʵ�ս�� = 0
        strDetail = ""
        .Sort = "DeviceID,NO"
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            strDetail = strDetail & !Content
            lngDeviceID = !DeviceID
            strNO = !NO
            curӦ�ս�� = curӦ�ս�� & !Ӧ�ս��
            curʵ�ս�� = curʵ�ս�� & !ʵ�ս��
            
            .MoveNext
            If .EOF Then
                GoTo makCommon
            ElseIf lngDeviceID <> !DeviceID And strNO <> !NO Then
makCommon:
                .MovePrevious
                
                Select Case NVL(!Type, 0)
                Case enuLinkType.DB
                Case enuLinkType.WEBServices
                    '��ͷ
                    strTitle = "<" & STR_ROOT & ">"
                    '����
                    strTmp = GetTitleContent(rsVal, !Type, curӦ�ս��, curʵ�ս��, !����, !NO, !�ⷿID)
                    If strTmp <> "" Then
                        strTitle = strTitle & vbCrLf & strTmp
                        '��ϸ
                        strTitle = strTitle & vbCrLf & strDetail
                        '��β
                        strTitle = strTitle & vbCrLf & "</" & STR_NODE_T & ">" & vbCrLf & "</" & STR_ROOT & ">"
                    End If
                    
                Case enuLinkType.Directory
                End Select
                
                '����rsData��¼��
                rsData.AddNew
                rsData!DeviceID = lngDeviceID
                rsData!NO = strNO
                rsData!���� = !����
                rsData!Content = strTitle
                rsData.Update
                
                strDetail = ""
                curӦ�ս�� = 0
                curʵ�ս�� = 0
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    Set rsTmp = Nothing

    Set BuildReceipDetail = rsData

End Function

Private Function GetTitleContent(ByVal rsTitle As ADODB.Recordset, ByVal bytType As Byte, _
    ByVal curӦ�ս�� As Currency, ByVal curʵ�ս�� As Currency, ByVal int���� As Integer, _
    ByVal strNO As String, ByVal lng�ⷿID As Long) As String
'���ܣ���ȡ����ͷ��Ϣ
    
    Const STR_NODE_T = "CONSIS_PRESC_MSTVW"
    Dim strTitle As String

    If rsTitle Is Nothing Then Exit Function

    rsTitle.Filter = "����=" & int���� & " and NO='" & strNO & "' and ��ҩҩ��id=" & lng�ⷿID
    With rsTitle
        If .EOF = False Then
            '������Ϣ
            Select Case bytType
            
            Case enuLinkType.WEBServices
                strTitle = "<" & STR_NODE_T
                strTitle = strTitle & vbCrLf & "PRESC_DATE = """ & Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
                strTitle = strTitle & vbCrLf & "PRESC_NO = """ & SpecialChar(!NO) & """"
                strTitle = strTitle & vbCrLf & "DISPENSARY = """ & NVL(!��ҩҩ��id) & """"
                strTitle = strTitle & vbCrLf & "PATIENT_ID = """ & NVL(!����ID) & """"
                strTitle = strTitle & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!����) & """"
                strTitle = strTitle & vbCrLf & "PATIENT_TYPE = """ & IIf(NVL(!���ȼ�) = "1", "01", "00") & """"
                strTitle = strTitle & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!��������), "yyyy-MM-DDThh:mm:ss") & """"
                strTitle = strTitle & vbCrLf & "SEX = """ & SpecialChar(!�Ա�) & """"
                strTitle = strTitle & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!���) & """"
                strTitle = strTitle & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!ҽ������) & """"
                strTitle = strTitle & vbCrLf & "PRESC_ATTR = """""
                strTitle = strTitle & vbCrLf & "PRESC_INFO = """""
                
                '�������SQL����д���ú���������ȡ
                strTitle = strTitle & vbCrLf & "RCPT_INFO = """ & SpecialChar(GetRCPT_INFO(NVL(!NO))) & """"
                
                strTitle = strTitle & vbCrLf & "RCPT_REMARK = """""
                strTitle = strTitle & vbCrLf & "REPETITION = ""1"""
                strTitle = strTitle & vbCrLf & "COSTS = """ & curӦ�ս�� & """"
                strTitle = strTitle & vbCrLf & "PAYMENTS = """ & curʵ�ս�� & """"
                strTitle = strTitle & vbCrLf & "ORDERED_BY = """ & NVL(!��������id) & """"
                strTitle = strTitle & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!����ҽ��) & """"
                strTitle = strTitle & vbCrLf & "ENTERED_BY = """ & SpecialChar(!����ҽ��) & """"
                strTitle = strTitle & vbCrLf & "DISPENSE_PRI = """ & NVL(!���ȼ�) & """"
                strTitle = strTitle & vbCrLf & ">"
                
            End Select
            
            If strTitle <> "" Then
                GetTitleContent = strTitle
            End If
        End If
    End With
    
End Function

Private Function BuildReceipList(ByVal rsVal As ADODB.Recordset, ByVal lngModule As Long) As ADODB.Recordset
'���ܣ������������﷢ҩ�ϴ����ݽṹ�ļ�¼������
'������
'  rsVal��HIS���ݼ�
'  lngModule��HISҵ��ģ���

    Const STR_ROOT = "ROOT"
    Const STR_NODE = "CONSIS_PRESC_MSTVW"
    
    Dim rsData As New ADODB.Recordset
    Dim strBill As String
    Dim lngDeviceID As Long
    Dim arrDeviceID As Variant
    Dim i As Integer
    
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "DeviceID", adBigInt, , adFldIsNullable
        .Fields.Append "NO", adVarChar, 20, adFldIsNullable
        .Fields.Append "Content", adVarChar, 1000, adFldIsNullable
        .Open
    End With
    
    With rsVal
        If .State <> adStateOpen Then .Open
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            
            arrDeviceID = GetDevices(NVL(!��ҩҩ��id, 0))
            
            strBill = "<" & STR_ROOT & ">"
            strBill = strBill & vbCrLf & "<" & STR_NODE
            strBill = strBill & vbCrLf & "PRESC_DATE = """ & Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "PRESC_NO = """ & NVL(!NO) & """"
            strBill = strBill & ">" & vbCrLf & "</" & STR_NODE & vbCrLf & ">"
            strBill = strBill & vbCrLf & "</" & STR_ROOT & ">"
            
            '��ͬ�ķ�ҩҩ��������������
            For i = LBound(arrDeviceID) To UBound(arrDeviceID)
                rsData.AddNew
                rsData!DeviceID = arrDeviceID(i)
                rsData!NO = !NO
                rsData!Content = strBill
                rsData.Update
            Next
            Set arrDeviceID = Nothing
            
            .MoveNext
        Loop
        .Close
    End With
    
    rsData.Sort = "NO,DeviceID"
    
    Set BuildReceipList = rsData
    
End Function

Private Function BuildReceipAdviceInf(ByVal rsVal As ADODB.Recordset, ByVal lngModule As Long) As ADODB.Recordset
'���ܣ���������סԺҽ����ҩ�ϴ����ݽṹ�ļ�¼������
'������
'  rsVal��HIS���ݼ�
'  lngModule��HISҵ��ģ���
    
    Dim rsData As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim lngDeviceID As Long
    Dim strTmp As String, strDataA As String, strDataB As String
    Dim intCount As Integer, i As Integer
    Dim strNextTime As String
    
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "DeviceID", adBigInt, , adFldIsNullable
        .Fields.Append "Title", adVarChar, 1000, adFldIsNullable
        .Fields.Append "Detail", adLongVarChar, 10000, adFldIsNullable
        .Fields.Append "��ҩ����ID", adBigInt, , adFldIsNullable
        .Open
    End With
    
    With rsVal
        If .State <> adStateOpen Then .Open
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            lngDeviceID = GetDevice(2, !��ҩҩ��id, !ҩƷ����)
            
            If lngDeviceID <= 0 Then GoTo makLoop
            
            'Ƶ�ʴ���
            '�������������������װ�������򲻷��͵���ҩ��
            If Not (!����װ = 0 Or !ҽ������ = "����") Then GoTo makLoop
            
            If Val(NVL(!Ƶ�ʼ��)) = 0 Or NVL(!�����λ) = "" Or NVL(!ִ��ʱ�䷽��) = "" Or !ҽ������ = "����" Then
                intCount = 1
            Else
                intCount = Val(NVL(!Ƶ�ʴ���))
                If intCount = 0 Then
                    strTmp = "Select Zl_Gettransexenumber([1],[2],[3],[4],[5],[6]) ִ�д��� From Dual "
                    On Error GoTo errHandle
                    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "ȡִ�д���", _
                                CDate(!��ʼִ��ʱ��), CDate(!�״�ʱ��), CDate(!ĩ��ʱ��), Val(!Ƶ�ʼ��), !�����λ, !ִ��ʱ�䷽��)
                    If Not rsTmp.EOF Then
                        intCount = Val(rsTmp.Fields(0).Value)
                    End If
                    rsTmp.Close
                    If intCount = 0 Then
                        intCount = 1
                    End If
                    On Error GoTo 0
                End If
            End If
            
            '��ϸ�ű�
            'ҽ��ҩƷ��Ϣ
            strDataA = "select "
            strDataA = strDataA & !�շ�id & " �շ�ID,"
            strDataA = strDataA & NVL(!סԺ��, "0") & " סԺ��,"
            strDataA = strDataA & !����ID & " ����ID,"
            strDataA = strDataA & "'" & !���� & "' ����,"
            strDataA = strDataA & IIf(NVL(!���˲�������) = "", "null", "'" & !���˲������� & "'") & " ���˲�������,"
            strDataA = strDataA & IIf(NVL(!���˲�������) = "", "null", "'" & !���˲������� & "'") & " ���˲�������,"
            strDataA = strDataA & IIf(NVL(!����ҽ��) = "", "null", "'" & !����ҽ�� & "'") & " ����ҽ��,"
            strDataA = strDataA & IIf(NVL(!����) = "", "null", "'" & !���� & "'") & " ����,"
            strDataA = strDataA & IIf(NVL(!ҩƷ�÷�) = "", "null", "'" & !ҩƷ�÷� & "'") & " ҩƷ�÷�,"
            strDataA = strDataA & "null ����ʱ��,"
            strDataA = strDataA & "'" & !ҩƷ���� & "' ҩƷ����,"
            strDataA = strDataA & "'" & !ҩƷ���� & "' ҩƷ����,"
            strDataA = strDataA & "'" & !��� & "' ���,"
            strDataA = strDataA & !����ϵ�� & " ����ϵ��,"
            strDataA = strDataA & "'" & !������λ & "' ������λ,"
            strDataA = strDataA & "1 �豸���,"
            strDataA = strDataA & "0 ���ȱ��,"
            strDataA = strDataA & IIf(NVL(!ҽ������) = "����", "1", "0") & " ����,"
            strDataA = strDataA & IIf(NVL(!�����) = "", "null", "'" & !����� & "'") & " �����"
            strDataA = strDataA & vbCrLf
            
            '��ֵ��η�������
            On Error GoTo errHandle
            strNextTime = Format(!�״�ʱ��, "YYYY-MM-DD HH:MM:SS")
            strDataB = ""
            For i = 1 To intCount
                strDataB = strDataB & "select "
                strDataB = strDataB & !�շ�id & " �շ�ID,"
                strDataB = strDataB & IIf(intCount = 1, !����, !���� / !����ϵ��) & " ��������,"
                strDataB = strDataB & "'" & strNextTime & "'" & " ִ��ʱ�� "
                
                If i < intCount Then
                    strDataB = strDataB & " union all " & vbCrLf
                    
                    gstrSQL = "Select Zl_Gettransexetime([1],[2],[3],[4],[5]) �´�ִ��ʱ�� From Dual "
                    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "ȡ�´�ִ��ʱ��", _
                                CDate(!��ʼִ��ʱ��), CDate(strNextTime), Val(!Ƶ�ʼ��), !�����λ, !ִ��ʱ�䷽��)
                    If rsTmp.EOF = False Then
                        strNextTime = Format(rsTmp.Fields(0).Value, "YYYY-MM-DD HH:MM:SS")
                    End If
                    rsTmp.Close
                End If
            Next
            On Error GoTo 0
            
            strDataB = "select a.*, b.��������, b.ִ��ʱ�� " & _
                       "from (" & strDataA & ") A, (" & strDataB & ") B " & _
                       "where a.�շ�ID=b.�շ�ID "
            
            '���ݽű�
            strDataA = "select "
            strDataA = strDataA & vbCrLf & "'" & !��ҩ���ű��� & "' ��ҩ���ű���,"
            strDataA = strDataA & vbCrLf & "'" & !��ҩҩ����� & "' ��ҩҩ�����,"
            strDataA = strDataA & vbCrLf & "1 �豸���,"
            strDataA = strDataA & vbCrLf & "getdate() ����ʱ��"
            
            '����
            rsData.AddNew
            rsData!DeviceID = lngDeviceID
            rsData!Title = strDataA
            rsData!Detail = strDataB
            rsData!��ҩ����ID = !��ҩ����ID
            rsData.Update
            
makLoop:
            .MoveNext
        Loop
        .Close
    End With
    
    Set BuildReceipAdviceInf = rsData
    Exit Function
    
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Function

Private Function BuildDrugStock(ByVal rsDrugStock As ADODB.Recordset, ByVal objDevice As clsDevice) As ADODB.Recordset
'���ܣ����������ϴ����ݽṹ��ҩƷ����¼������
'������
'  rsDrugStock��HISҩƷ����¼������
'  objDevice���豸����

    Const STR_ROOT = "ROOT"
    Const STR_NODE = "CONSIS_PHC_STORAGEVW"
    
    Dim strTmp As String
    Dim rsData As New ADODB.Recordset
    Dim i As Integer

    '��ʼ���ڴ��¼������
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "SN", adBigInt, , adFldIsNullable
        .Fields.Append "Drug", adVarChar, 100, adFldIsNullable
        .Fields.Append "Content", adVarChar, 3000, adFldIsNullable
        .Open
    End With

    With rsDrugStock
        If .State <> adStateOpen Then .Open
        i = 1
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
        
            '��ʽ����Ҫ�ϴ������ݸ�ʽ
            Select Case objDevice.LinkType
            Case enuLinkType.DB
            
            Case enuLinkType.WEBServices
                strTmp = "<" & STR_ROOT & ">"
                strTmp = strTmp & vbCrLf & "<" & STR_NODE
                strTmp = strTmp & vbCrLf & "DRUG_CODE = """ & SpecialChar(!����) & """"
                strTmp = strTmp & vbCrLf & "DISPENSARY = """ & !�ⷿID & """"
                strTmp = strTmp & vbCrLf & "DRUG_QUANTITY = """ & NVL(!ʵ������, 0) / NVL(!�����װ, 1) & """"
                strTmp = strTmp & vbCrLf & "LOCATIONINFO = """ & SpecialChar(NVL(!�ⷿ��λ)) & """"
                strTmp = strTmp & vbCrLf & ">"
                strTmp = strTmp & vbCrLf & "</" & STR_NODE & ">"
                strTmp = strTmp & vbCrLf & "</" & STR_ROOT & ">"
            Case enuLinkType.Directory
            End Select
            
            '�����ڴ��¼��
            If strTmp <> "" Then
                rsData.AddNew
                rsData!SN = i
                rsData!Drug = !���� & "��" & !ͨ���� & "��" & NVL(!���)
                rsData!Content = strTmp
                rsData.Update
                
                i = i + 1
            End If
            
            .MoveNext
        Loop
        .Close
        
    End With
            
    Set BuildDrugStock = rsData

End Function
