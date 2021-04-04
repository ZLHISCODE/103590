Attribute VB_Name = "mdlPassDefine_YWS"
Option Explicit

Public gstrBaseXml As String     '����BaseXML


Public Function YWS_MakeBASEXML(ByRef xmlbase As YWS_BASE) As String
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
                        strTab2 & "<type_name>" & .strҽ���������� & "</type_name >" & _
                    strTab1 & "</doct>" & vbCrLf & _
                    "</base_xml>"
                
    End With
    YWS_MakeBASEXML = strXML
End Function

Public Function YWS_MakeMedicXML(ByRef xmldetails As YWS_DETAILS) As String
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
    YWS_MakeMedicXML = strXML
End Function

Public Function YWS_MakePresXML(ByRef xmldetails As YWS_DETAILS) As String
'���ܣ�'HIS���� ��6��8��9
    Dim strXML As String, strTmp As String
    Dim strTab1 As String, strTab2 As String, strTab3 As String, strTab4 As String
    Dim udt����Դ As YWS_ALLERGIC
    Dim udt��� As YWS_DIAGNOSE
    Dim udtҩƷ As YWS_MEDICINE
    
    Dim i As Long
    
    
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    strTab3 = vbCrLf & vbTab & vbTab & vbTab
    strTab4 = vbCrLf & vbTab & vbTab & vbTab & vbTab
    With xmldetails
        strXML = "<details_xml>" & _
                    strTab1 & "<his_time>" & .strHISϵͳʱ�� & "</his_time>" & _
                    strTab1 & "<hosp_flag>" & .str����סԺ��ʶ & "</hosp_flag>" & _
                    strTab1 & "<treat_type>" & .str�������� & "</treat_type>" & _
                    strTab1 & "<treat_code>" & .str����� & "</treat_code>" & _
                    strTab1 & "<bed_no>" & .str��λ�� & "</bed_no>"
        With .udt������Ϣ
            strXML = strXML & _
            strTab1 & "<patient>" & _
                strTab2 & "<name>" & .str���� & "</name>" & _
                strTab2 & "<birth>" & .str�������� & "</birth>" & _
                strTab2 & "<sex>" & .str�Ա� & "</sex>" & _
                strTab2 & "<weight>" & .str���� & "</weight>" & _
                strTab2 & "<height>" & .str��� & "</height>" & _
                strTab2 & "<id_card>" & .str���֤�� & "</id_card>" & _
                strTab2 & "<medical_record>" & .str�������� & "</medical_record>" & _
                strTab2 & "<card_type>" & .str������ & "</card_type>" & _
                strTab2 & "<card_code>" & .str���� & "</card_code>" & _
                strTab2 & "<pregnant_unit>" & .str����ʱ�䵥λ & "</pregnant_unit>" & _
                strTab2 & "<pregnant >" & .str����ʱ�� & "</pregnant>"
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
        End With
        strXML = strXML & strTab1 & "</patient>"
        '������Ϣ
        strXML = strXML & strTab1 & "<prescription_data>" & strTab2 & "<prescription>"
        With .udt������Ϣ
            strXML = strXML & _
            strTab3 & "<id>" & .str������ & "</id>" & _
            strTab3 & "<reason>" & .str�������� & "</reason>" & _
            strTab3 & "<is_current>" & .str�Ƿ�ǰ���� & "</is_current>" & _
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
                    strTab3 & "<medicine>" & _
                        strTab4 & "<zxy_type>" & .strҩƷ���� & "</zxy_type>" & _
                        strTab4 & "<oeridid>" & .str������ & "</oeridid>" & _
                        strTab4 & "<pres_type>" & .Strҽ������ & "</pres_type>" & _
                        strTab4 & "<pres_time>" & .str����ʱ�� & "</pres_time>" & _
                        strTab4 & "<name>" & .str��Ʒ�� & "</name>" & _
                        strTab4 & "<his_code>" & .strҽԺҩƷ���� & "</his_code>" & _
                        strTab4 & "<insur_code>" & .strҽ������ & "</insur_code>" & _
                        strTab4 & "<approval>" & .str��׼�ĺ� & "</approval>" & _
                        strTab4 & "<spec>" & .str��� & "</spec>" & _
                        strTab4 & "<group>" & .str��� & "</group>" & _
                        strTab4 & "<reason>" & .str��ҩ���� & "</reason>" & _
                        strTab4 & "<dose_unit>" & .str��������λ & "</dose_unit>" & _
                        strTab4 & "<dose>" & .str������ & "</dose>" & _
                        strTab4 & "<freq>" & .strƵ�δ��� & "</freq>" & _
                        strTab4 & "<administer>" & .str��ҩ;������ & "</administer>" & _
                        strTab4 & "<begin_time>" & .str��ҩ��ʼʱ�� & "</begin_time>" & _
                        strTab4 & "<end_time>" & .str��ҩ����ʱ�� & "</end_time>" & _
                        strTab4 & "<days>" & .str��ҩ���� & "</days>" & _
                    strTab3 & "</medicine>"
                End With
            Next
            strXML = strXML & strTab2 & "<medicine_data>" & strTmp & strTab2 & "</medicine_data>"
        End With
        strXML = strXML & strTab2 & "</prescription>" & strTab1 & "</prescription_data>"
        strXML = strXML & vbCrLf & "</details_xml>"
    End With
'    Debug.Print strXML
    
    YWS_MakePresXML = strXML
End Function

Public Function YWS_StrToXML(ByVal strValue As String) As String
'����:�������ַ����滻�ɹ涨�ַ�
    strValue = Replace(strValue, "&", "&amp;")
    strValue = Replace(strValue, ">", "&gt;")
    strValue = Replace(strValue, "<", "&lt;")
    strValue = Replace(strValue, "'", "&apos;")
    YWS_StrToXML = Replace(strValue, """", "&quot;")
End Function

Public Function YWS_MakeDetailXML(ByVal bytFunc As YWS_Func_NUM) As String
'���ܣ�����details XML�ַ���
    Dim strXML As String
    
        Select Case bytFunc
        
        Case YWS_�˳�
            strXML = "" & _
            "<details_xml>" & vbCrLf & _
                vbTab & "<details_info></details_info>" & vbCrLf & _
            "</details_xml>"
        Case YWS_��ʼ�ͻ���
            strXML = "<details_xml></details_xml>"
        End Select
    
    YWS_MakeDetailXML = strXML
End Function

Public Function YWS_GetTreatType(ByVal bytFunc As Byte, ByVal lng�Һ�ID As Long, Optional lng��ҳID As Long) As String
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
    YWS_GetTreatType = strRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function YWS_ReturnRS(ByVal strXML As String, Optional ByVal bytFunc As Byte) As ADODB.Recordset
    '���ܣ���ҩ��ʿ����������н����������������������ҩ����
    'bytFunc=0 ������ʾ��   =1 ���������
    '<?xml version='1.0' encoding='utf-8'?>
    '<ui_results_xml>
    '  <result_data>
    '    <result>
    '      <oeridid>30064��9911</oeridid>
    '      <result_type>3</result_type>
    '      <result_code>1</result_code>
    '      <result_title>�������</result_title>
    '      <title>�����ע��Һ�ԡ� ά���أ�ע��Һ�����������</title>
    '      <detail>�����ע��������н��ɣ�Ӧ���⣡�����ƣ�ά����Cע��Һ�Ĳ�Ʒ����˵����ά����C���������ҩ���簱�����Һ���飬����Ӱ����Ч������������ʾ��ά����Cע��Һ(Ũ��Ϊ12.5%��pHֵΪ5.7��7.0)�백���ע��Һ(Ũ
    '      ��Ϊ25mg/ml��pHֵΪ8.6��9.3)��Ϻ���Һ�����������[1]��</detail>
    '      <reference>�ο����ף�</reference>
    '      <mediA_hiscode>30064</mediA_hiscode>
    '      <mediA_ywscode></mediA_ywscode>
    '      <mediA_name>�����ע��Һ</mediA_name>
    '      <mediB_hiscode>9911</mediB_hiscode>
    '      <mediB_ywscode></mediB_ywscode>
    '      <mediB_name>ά���أ�ע��Һ</mediB_name>
    '    </result>
    '  </result_data>
    '</ui_results_xml>
        Dim xmlDoc As DOMDocument
        Dim xmlRoot As IXMLDOMElement
        Dim xmlNode As IXMLDOMNode
        Dim xmlNodes As IXMLDOMNodeList
        Dim rsRet As New ADODB.Recordset
        Dim arrTmp As Variant
    
        Dim str��ʾֵ As String
        Dim strҽ��ID As String, strҽ���� As String
    
        Dim i As Long
    
        On Error GoTo errH
    
100     Set xmlDoc = New DOMDocument
102     xmlDoc.loadXML (strXML)
104     If bytFunc = 0 Then
106         Set rsRet = InitAdviceRS(FUN_�����)
        Else
108         Set rsRet = InitAdviceRS(FUN_�����_YWS)
        End If
        '����������κ�Ԫ�أ����˳�
110     If xmlDoc.documentElement Is Nothing Then
112         Set xmlDoc = Nothing
114         Set YWS_ReturnRS = rsRet
            Exit Function
        End If
    
        '��ȡXML����
116     Set xmlRoot = xmlDoc.selectSingleNode("ui_results_xml/result_data")
118     Set xmlNodes = xmlRoot.selectNodes("result")
120     If bytFunc = 0 Then
122         If Not xmlNodes Is Nothing Then
124             For Each xmlNode In xmlNodes
126                 str��ʾֵ = xmlNode.selectSingleNode("result_type").Text
128                 If Val(str��ʾֵ) > 0 Then
130                     strҽ���� = xmlNode.selectSingleNode("oeridid").Text   '30064��9911
132                     arrTmp = Split(strҽ����, "��")
134                     For i = LBound(arrTmp) To UBound(arrTmp)
136                         strҽ��ID = Val(arrTmp(i))
138                         If Val(arrTmp(i)) <> 0 Then
140                             rsRet.Filter = "ҽ��ID ='" & strҽ��ID & "'"
142                             If Not rsRet.EOF Then
144                                 If Val(rsRet!��ʾֵ & "") < Val(str��ʾֵ) Then
146                                     rsRet!��ʾֵ = Val(str��ʾֵ)
                                    End If
                                Else
148                                 rsRet.AddNew
150                                 rsRet!��ʾֵ = Val(str��ʾֵ)
152                                 rsRet!ҽ��ID = Val(strҽ��ID)
154                                 rsRet.Update
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Else
156         If Not xmlNodes Is Nothing Then
158             For Each xmlNode In xmlNodes
160                 rsRet.AddNew
162                 rsRet!Title = xmlNode.selectSingleNode("title").Text
164                 rsRet!Detail = xmlNode.selectSingleNode("detail").Text
166                 rsRet.Update
                Next
            End If
        End If
168     If rsRet.RecordCount > 0 Then rsRet.Filter = ""
    
170     Set YWS_ReturnRS = rsRet
        Exit Function
errH:
172     MsgBox "YWS_ReturnRS �����:" & Err.Number & "������:" & Erl() & " ��������:" & Err.Description, vbOKOnly, gstrSysName
End Function

Public Function YWS_GetDrugType(ByVal strType As String) As String
'����:����ҩƷ����
'   5-��ҩ/6-�г�ҩ/7-��ҩ
    Dim strRet As String
    
    Select Case strType
    
    Case "5"
        strRet = "��ҩ"
    Case "6"
        strRet = "�г�ҩ"
    Case "7"
        strRet = "��ҩ"
    End Select
    YWS_GetDrugType = strRet
End Function
