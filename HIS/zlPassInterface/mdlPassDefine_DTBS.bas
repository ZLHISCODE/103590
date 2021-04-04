Attribute VB_Name = "mdlPassDefine_DTBS"
Option Explicit

'��ͨBS��ӿڶ���
Public Declare Function CRMS_UI Lib "CRMS_UI.dll" (ByVal lngFunc As Long, ByVal strBaseXml As String, ByVal strDetailsXml As String, ByRef strResults As String) As Long
'������lngFunc(���ܱ�ʶ)
'     strBaseXml(������ϢXMl)
'     strDetailsXml����ϸ��ϢXML��
'����
'    strResults��his���ؽ��XML��

Public Function DTBS_StrToXML(ByVal strValue As String) As String
'����:�������ַ����滻�ɹ涨�ַ�
    strValue = Replace(strValue, "&", "&amp;")
    strValue = Replace(strValue, ">", "&gt;")
    strValue = Replace(strValue, "<", "&lt;")
    strValue = Replace(strValue, "'", "&apos;")
    DTBS_StrToXML = Replace(strValue, """", "&quot;")
End Function

Public Function DTBS_MakeBASEXML(ByRef xmlbase As DTBS_BASE) As String
'���ܣ�����BASE XML�ַ���
    Dim strXML As String
    Dim strTab1 As String, strTab2 As String
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    
    With xmlbase
        strXML = "<base_xml>" & _
                    strTab1 & "<source>" & .strHIS & "</source>" & _
                    strTab1 & "<hosp_code>" & .strҽԺ���� & "</hosp_code>" & _
                    strTab1 & "<dept_code>" & .str���Ҵ��� & "</dept_code>" & _
                    strTab1 & "<dept_name>" & .str�������� & "</dept_name>" & _
                    strTab1 & "<doct>" & _
                        strTab2 & "<code>" & .strҽ������ & "</code>" & _
                        strTab2 & "<name>" & .strҽ������ & "</name>" & _
                        strTab2 & "<type>" & .strҽ��������� & "</type>" & _
                        strTab2 & "<type_name>" & .strҽ���������� & "</type_name>" & _
                    strTab1 & "</doct>" & vbCrLf & _
                    "</base_xml>"
                
    End With
    DTBS_MakeBASEXML = strXML
End Function

Public Function DTBS_MakeDetailXML(ByVal bytFunc As DTBS_Func_NUM, Optional ByVal strDoctPWD As String) As String
'���ܣ�����details XML�ַ���
    Dim strXML As String
    
        Select Case bytFunc
        Case DTBS_��¼
            strXML = "<details_xml>" & vbCrLf & _
                        "<doct_pwd>" & strDoctPWD & "</doct_pwd>" & vbCrLf & _
                    "</details_xml>"
        Case DTBS_�˳�
            strXML = "" & _
            "<details_xml>" & vbCrLf & _
                vbTab & "<details_info></details_info>" & vbCrLf & _
            "</details_xml>"
        Case DTBS_��ʼUI
            strXML = "<details_xml></details_xml>"
        End Select
    
    DTBS_MakeDetailXML = strXML
End Function

Public Function DTBS_MakeMedicXML(ByRef xmldetails As DTBS_DETAILS) As String
'���ܣ�HIS���� ��5����
    Dim strXML As String
    Dim strTab1 As String, strTab2 As String
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    
    With xmldetails
        strXML = "<details_xml>" & _
                    strTab1 & "<hosp_flag>" & .str����סԺ��ʶ & "</hosp_flag>" & _
                    strTab1 & "<medicine>" & _
                        strTab2 & "<his_code>" & .strҩƷ���� & "</his_code>" & _
                        strTab2 & "<his_name>" & .strҩƷ���� & "</his_name>" & _
                    strTab1 & "</medicine>" & vbCrLf & _
                "</details_xml>"
    End With
'    Debug.Print strXML
    DTBS_MakeMedicXML = strXML
End Function

Public Function DTBS_MakePresXML(ByRef xmldetails As DTBS_DETAILS) As String
'���ܣ�'HIS���� ��6
    Dim strXML As String, strTmp As String, strSub As String, strPres As String
    Dim strTab1 As String, strTab2 As String, strTab3 As String, strTab4 As String, strTab5 As String
    Dim udt����Դ As DTBS_ALLERGIC
    Dim udt��� As DTBS_DIAGNOSE
    Dim udt������Ϣ As DTBS_PRESCRIPTION
    Dim udtLISFORM As DTBS_LISFORM
    Dim udtLISITEM As DTBS_LISITEM
    Dim udtҩƷ As DTBS_MEDICINE
    
    Dim i As Long, j As Long
    
    
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    strTab3 = vbCrLf & vbTab & vbTab & vbTab
    strTab4 = vbCrLf & vbTab & vbTab & vbTab & vbTab
    strTab5 = vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
    
    With xmldetails
        strXML = "<details_xml  is_upload =""" & .str�Ƿ��ϴ� & """>" & _
                    strTab1 & "<his_time>" & .strHISϵͳʱ�� & "</his_time>" & _
                    strTab1 & "<hosp_flag>" & .str����סԺ��ʶ & "</hosp_flag>" & _
                    strTab1 & "<treat_type>" & .str�������� & "</treat_type>" & _
                    strTab1 & "<treat_code>" & .str����� & "</treat_code>" & _
                    strTab1 & "<lis_adm_no>" & .str�������� & "</lis_adm_no>" & _
                    strTab1 & "<bed_no>" & .str��λ�� & "</bed_no>" & _
                    strTab1 & "<area_code>" & .str������ & "</area_code>"
        With .udt������Ϣ
            strXML = strXML & _
            strTab1 & "<patient>" & _
                strTab2 & "<name>" & .str���� & "</name>" & _
                strTab2 & "<is_infant>" & .str�Ƿ�Ӥ�� & "</is_infant>" & _
                strTab2 & "<birth>" & .str�������� & "</birth>" & _
                strTab2 & "<sex>" & .str�Ա� & "</sex>" & _
                strTab2 & "<weight>" & .str���� & "</weight>" & _
                strTab2 & "<height>" & .str��� & "</height>" & _
                strTab2 & "<id_card>" & .str���֤�� & "</id_card>" & _
                strTab2 & "<card_type>" & .str������ & "</card_type>" & _
                strTab2 & "<card_code>" & .str���� & "</card_code>" & _
                strTab2 & "<pregnant_unit>" & .str����ʱ�䵥λ & "</pregnant_unit>" & _
                strTab2 & "<pregnant>" & .str����ʱ�� & "</pregnant>"
            '����Դ
            strTmp = ""
            If Not .col����Դs Is Nothing Then
                For i = 1 To .col����Դs.Count
                    udt����Դ = .col����Դs(i)
                    With udt����Դ
                        strTmp = strTmp & _
                        strTab3 & "<allergic>" & _
                            strTab4 & "<type>" & .str�������� & "</type>" & _
                            strTab4 & "<name>" & .str����Դ���� & "</name>" & _
                            strTab4 & "<code>" & .str����Դ���� & "</code>" & _
                        strTab3 & "</allergic>"
                    End With
                Next
            End If
            strXML = strXML & strTab2 & "<allergic_data>" & strTmp & strTab2 & "</allergic_data>"
            
            '���
            strTmp = ""
            If Not .col���s Is Nothing Then
                For i = 1 To .col���s.Count
                    udt��� = .col���s(i)
                    With udt���
                        strTmp = strTmp & _
                        strTab3 & "<diagnose>" & _
                            strTab4 & "<type>" & .str������� & "</type>" & _
                            strTab4 & "<name>" & .str������� & "</name>" & _
                            strTab4 & "<code>" & .str��ϴ��� & "</code>" & _
                        strTab3 & "</diagnose>"
                    End With
                Next
            End If
            strXML = strXML & strTab2 & "<diagnose_data>" & strTmp & strTab2 & "</diagnose_data>"
            '�����ⵥ�ڵ�
            strTmp = ""
            If Not .col������ Is Nothing Then
                For i = 1 To .col������.Count
                    udtLISFORM = .col������(i)
                    With udtLISFORM
                        strTmp = strTmp & _
                        strTab3 & "<form>" & _
                            strTab4 & "<no>" & .str���� & "</no>" & _
                            strTab4 & "<project_name>" & .str��Ŀ���� & "</project_name>" & _
                            strTab4 & "<lis_flag>" & .str��� & "</lis_flag>" & _
                            strTab4 & "<result_date>" & .str�������ʱ�� & "</result_date>" & _
                            strTab4 & "<sample_code>" & .str������������ & "</sample_code>" & _
                            strTab4 & "<sample_name>" & .str������������ & "</sample_name>" & _
                            strTab4 & "<mac_flag>" & .str΢�����ͼ��ʶ & "</mac_flag>" & _
                        strTab3 & "</form>"
                        
                        strSub = ""
                        If Not .col��Ŀ�ڵ� Is Nothing Then
                            For j = 1 To .col��Ŀ�ڵ�.Count
                                udtLISITEM = .col��Ŀ�ڵ�(i)
                                With udtLISITEM
                                    strSub = strSub & _
                                    strTab4 & "<item>" & _
                                        strTab5 & "<code>" & .str���� & "</code>" & _
                                        strTab5 & "<name>" & .str���� & "</name>" & _
                                        strTab5 & "<value>" & .str��� & "</value>" & _
                                        strTab5 & "<uom>" & .str���ֵ��λ & "</uom>" & _
                                        strTab5 & "<upper>" & .str�ο���Χ���� & "</upper>" & _
                                        strTab5 & "<lower>" & .str�ο���Χ���� & "</lower>"
                                End With
                            Next
                        End If
                        strTmp = strTmp & strSub
                    End With
                Next
            End If
            strXML = strXML & strTab2 & "<lis_data>" & strTmp & strTab2 & "</lis_data>"
        End With
        strXML = strXML & strTab1 & "</patient>"
        '������Ϣ
        If Not .col������Ϣ Is Nothing Then
            strPres = ""
            For j = 1 To .col������Ϣ.Count
                udt������Ϣ = .col������Ϣ(j)
                With udt������Ϣ
                    strPres = strPres & strTab2 & "<prescription>" & _
                    strTab3 & "<id>" & .str������ & "</id>" & _
                    strTab3 & "<reason>" & .str�������� & "</reason>" & _
                    strTab3 & "<is_urgent>" & .str�Ƿ�������� & "</is_urgent>" & _
                    strTab3 & "<is_new>" & .str�Ƿ��¿����� & "</is_new>" & _
                    strTab3 & "<is_current>" & .str�Ƿ�ǰ���� & "</is_current>" & _
                    strTab3 & "<doct_code>" & .str����ҽ������ & "</doct_code>" & _
                    strTab3 & "<doct_name>" & .str����ҽ������ & "</doct_name>" & _
                    strTab3 & "<dept_code>" & .str�������Ҵ��� & "</dept_code>" & _
                    strTab3 & "<dept_name>" & .str������������ & "</dept_name>" & _
                    strTab3 & "<pres_type>" & .Strҽ������ & "</pres_type>" & _
                    strTab3 & "<pres_time>" & .str����ʱ�� & "</pres_time>"
                    'ҩƷ��Ϣ
                    If .colҩƷ��Ϣ Is Nothing Then
                        Set .colҩƷ��Ϣ = New Collection
                        .colҩƷ��Ϣ.Add udtҩƷ, "_1"
                    End If
                    
                    strTmp = ""
                    For i = 1 To .colҩƷ��Ϣ.Count
                        udtҩƷ = .colҩƷ��Ϣ(i)
                        With udtҩƷ
                            strTmp = strTmp & _
                            strTab4 & "<medicine>" & _
                                strTab5 & "<name>" & .str��Ʒ�� & "</name>" & _
                                strTab5 & "<his_code>" & .strҽԺҩƷ���� & "</his_code>" & _
                                strTab5 & "<insur_code>" & .strҽ������ & "</insur_code>" & _
                                strTab5 & "<pyd_code>" & .str��Һ���� & "</pyd_code>" & _
                                strTab5 & "<link_group>" & .str��Һ����� & "</link_group>" & _
                                strTab5 & "<spec>" & .str��� & "</spec>" & _
                                strTab5 & "<group>" & .str��� & "</group>" & _
                                strTab5 & "<reason>" & .str��ҩ���� & "</reason>" & _
                                strTab5 & "<dose_unit>" & .str��������λ & "</dose_unit>" & _
                                strTab5 & "<dose>" & .str������ & "</dose>" & _
                                strTab5 & "<freq>" & .strƵ�δ��� & "</freq>" & _
                                strTab5 & "<administer>" & .str��ҩ;������ & "</administer>" & _
                                strTab5 & "<begin_time>" & .str��ҩ��ʼʱ�� & "</begin_time>" & _
                                strTab5 & "<end_time>" & .str��ҩ����ʱ�� & "</end_time>" & _
                                strTab5 & "<days>" & .str��ҩ���� & "</days>" & _
                                strTab5 & "<preventiveflag>" & .str�Ƿ�Ԥ����ҩ & "</preventiveflag>" & _
                                strTab5 & "<otno>" & .str�������� & "</otno>" & _
                                strTab5 & "<signer_code>" & .strǩ��ҽʦ���� & "</signer_code>" & _
                                strTab5 & "<accredit_date>" & .str��Ȩʱ�� & "</accredit_date>" & _
                                strTab5 & "<accredit_hours>" & .str������ҩʱ�� & "</accredit_hours>" & _
                                strTab5 & "<accredit_times>" & .str������ҩ���� & "</accredit_times>" & _
                            strTab4 & "</medicine>"
                        End With
                    Next
                    strPres = strPres & strTab3 & "<medicine_data>" & strTmp & strTab3 & "</medicine_data>" & strTab2 & "</prescription>"
                End With
            Next
            strXML = strXML & strTab1 & "<prescription_data>" & strPres & strTab1 & "</prescription_data>"
        End If
        
        strXML = strXML & vbCrLf & "</details_xml>"
    End With
'    Debug.Print strXML
    
    DTBS_MakePresXML = strXML
End Function

Public Function DTBS_GetTreatType(ByVal bytFunc As Byte, ByVal lng�Һ�ID As Long, Optional lng��ҳID As Long) As String
'����:��ȡ��������
'����:bytFunc =1 ����,bytFunc=2 סԺ
'     lng�Һ�ID =���� �Һ�ID,סԺ =����ID
'100=��ͨ����
'101=ר������
'102=ר������
'200=����
'300=����۲�
'400=��ͨסԺ
'401=����סԺ
'500=�Ҵ�
'999=����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strRet As String
    
    If bytFunc = 1 Then
        strSQL = "Select Nvl(a.����,0) as ����,b.���� From ���˹Һż�¼ A, �ҺŰ��� B Where a.Id = [1] And a.�ű� = b.����"
    Else
        strSQL = "Select ��������, ��Ժ���� From ������ҳ Where ����id = [1] And ��ҳid = [2] And ��Ժ���� Is Null"

    End If
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng�Һ�ID, lng��ҳID)
    
    If rsTmp.RecordCount > 0 Then
        If bytFunc = 1 Then
            If rsTmp!���� = 1 Then
                strRet = "200"
            Else
                If rsTmp!���� & "" = "��ͨ" Then
                    strRet = "100"
                ElseIf rsTmp!���� & "" = "ר��" Then
                    strRet = "101"
                ElseIf rsTmp!���� & "" = "ר��" Then
                    strRet = "102"
                Else
                    strRet = "999"
                End If
            End If
        ElseIf bytFunc = 2 Then
            If rsTmp!��Ժ���� & "" = "" Then
                strRet = "500"   '��ͥ����
            ElseIf rsTmp!�������� = 0 Then
                strRet = "400"
            ElseIf rsTmp!�������� = 1 Or rsTmp!�������� = 2 Then
                strRet = "300"
            Else
                strRet = "999"
            End If
            
        End If
    End If
    DTBS_GetTreatType = strRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

