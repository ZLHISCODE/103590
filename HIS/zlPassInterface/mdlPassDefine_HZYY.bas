Attribute VB_Name = "mdlPassDefine_HZYY"
Option Explicit

Private Function GetҩƷ��Ϣ_HZYY(ByVal strDrugIDs As String) As ADODB.Recordset
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errH
    strSQL = "Select a.ҩƷid, a.���ﵥλ, a.�����װ, a.סԺ��λ, a.סԺ��װ, d.ҩƷ����, e.���,e.���㵥λ , f.�ּ�,B.��������,C.���� as ���ұ��� " & vbNewLine & _
            "From ҩƷ��� A, ҩƷ���� D, �շ���ĿĿ¼ E, �շѼ�Ŀ F, ҩƷ�����̶��� B, ҩƷ������ C" & vbNewLine & _
            "Where a.ҩ��id = d.ҩ��id And a.ҩƷid = e.Id And a.ҩƷid = f.�շ�ϸĿid(+) And a.ҩƷid = b.ҩƷid(+) And b.�������� = c.����(+) And" & vbNewLine & _
            "      Nvl(f.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) > Sysdate And" & vbNewLine & _
            "      a.ҩƷid In (Select /*+cardinality(A,10)*/" & vbNewLine & _
            "                  *" & vbNewLine & _
            "                 From Table(f_Num2list([1])) A)"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_HZYY", strDrugIDs)
    Set GetҩƷ��Ϣ_HZYY = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HZYY_MakeBASEXML(ByRef xmlbase As HZYYBASE) As String
'���ܣ�����BASE XML�ַ���
    Dim strXML As String
    Dim strTab1 As String
    strTab1 = vbCrLf & vbTab

    With xmlbase
        strXML = "<base>" & _
                    strTab1 & "<hospital_code><![CDATA[" & .strHospCode & "]]></hospital_code>" & _
                    strTab1 & "<event_no><![CDATA[" & .strEventNO & "]]></event_no>" & _
                    strTab1 & "<patient_id><![CDATA[" & .strPatiID & "]]></patient_id>" & _
                    strTab1 & "<source><![CDATA[" & .strSource & "]]></source>" & vbCrLf & _
                "</base>"
                
    End With
    HZYY_MakeBASEXML = strXML
End Function

Public Function HZYY_MakeOPSPatient(ByRef xmlPati As OPTPATIENT) As String
    Dim strXML As String
    Dim strTab1 As String
    strTab1 = vbCrLf & vbTab
    With xmlPati
        strXML = "<opt_patient>" & _
                    strTab1 & "<sex><![CDATA[" & .strSex & "]]></sex>" & _
                    strTab1 & "<name><![CDATA[" & .strName & "]]></name>" & _
                    strTab1 & "<id_type><![CDATA[" & .strIDType & "]]></id_type>" & _
                    strTab1 & "<id_no><![CDATA[" & .strIDNO & "]]></id_no>" & _
                    strTab1 & "<birth_weight><![CDATA[" & .strBirthWeight & "]]></birth_weight>" & _
                    strTab1 & "<birthday><![CDATA[" & .strBirthDay & "]]></birthday>" & _
                    strTab1 & "<ethnic_group><![CDATA[" & .strEthnicGroup & "]]></ethnic_group>" & _
                    strTab1 & "<native_place><![CDATA[" & .strNativePlace & "]]></native_place>" & _
                    strTab1 & "<race><![CDATA[" & .strRace & "]]></race>" & _
                    strTab1 & "<med_card_no><![CDATA[" & .strMedCardNO & "]]></med_card_no>" & _
                    strTab1 & "<event_time><![CDATA[" & .strEventTime & "]]></event_time>" & _
                    strTab1 & "<dept_id><![CDATA[" & .strDeptID & "]]></dept_id>" & _
                    strTab1 & "<dept_name><![CDATA[" & .strDeptName & "]]></dept_name>" & _
                    strTab1 & "<pay_type><![CDATA[" & .strPayType & "]]></pay_type>" & _
                    strTab1 & "<pregnancy><![CDATA[" & .strPregnancy & "]]></pregnancy>" & _
                    strTab1 & "<time_of_preg><![CDATA[" & .strTimeOfPreg & "]]></time_of_preg>" & _
                    strTab1 & "<breast_feeding><![CDATA[" & .strBreastFeeding & "]]></breast_feeding>" & _
                    strTab1 & "<height><![CDATA[" & .strHeight & "]]></height>" & _
                    strTab1 & "<weight><![CDATA[" & .strWeight & "]]></weight>" & _
                    strTab1 & "<address><![CDATA[" & .strAddress & "]]></address>"
        strXML = strXML & _
                    strTab1 & "<phone_no><![CDATA[" & .strPhoneNo & "]]></phone_no>" & _
                    strTab1 & "<dialysis><![CDATA[" & .strDialysis & "]]></dialysis>" & _
                    strTab1 & "<marital><![CDATA[" & .strmarital & "]]></marital>" & _
                    strTab1 & "<occupation><![CDATA[" & .strOccupation & "]]></occupation>" & _
                    strTab1 & "<special_constitution><![CDATA[" & .strSpecialConstitution & "]]></special_constitution>" & _
                    strTab1 & "<visit_type><![CDATA[" & .strVisitType & "]]></visit_type>" & _
                    strTab1 & "<patient_condition><![CDATA[" & .strPatiCondition & "]]></patient_condition>" & _
                    vbCrLf & _
                "</opt_patient>"
    End With

    HZYY_MakeOPSPatient = strXML
End Function

Public Function HZYY_MakeIPSPatient(ByRef xmlPati As IPTPATIENT) As String
    Dim strXML As String
    Dim strTab1 As String
    strTab1 = vbCrLf & vbTab

    With xmlPati
        strXML = "<ipt_patient>" & _
            strTab1 & "<sex><![CDATA[" & .strSex & "]]></sex>" & _
            strTab1 & "<name><![CDATA[" & .strName & "]]></name>" & _
            strTab1 & "<id_type><![CDATA[" & .strIDType & "]]></id_type>" & _
            strTab1 & "<id_no><![CDATA[" & .strIDNO & "]]></id_no>" & _
            strTab1 & "<birth_weight><![CDATA[" & .strBirthWeight & "]]></birth_weight>" & _
            strTab1 & "<birthday><![CDATA[" & .strBirthDay & "]]></birthday>" & _
            strTab1 & "<ethnic_group><![CDATA[" & .strEthnicGroup & "]]></ethnic_group>" & _
            strTab1 & "<native_place><![CDATA[" & .strNativePlace & "]]></native_place>" & _
            strTab1 & "<race><![CDATA[" & .strRace & "]]></race>" & _
            strTab1 & "<med_card_no><![CDATA[" & .strMedCardNO & "]]></med_card_no>" & _
            strTab1 & "<pay_type><![CDATA[" & .strPayType & "]]></pay_type>" & _
            strTab1 & "<marital><![CDATA[" & .strmarital & "]]></marital>" & _
            strTab1 & "<occupation><![CDATA[" & .strOccupation & "]]></occupation>" & _
            strTab1 & "<pregnancy><![CDATA[" & .strPregnancy & "]]></pregnancy>" & _
            strTab1 & "<time_of_preg><![CDATA[" & .strTimeOfPreg & "]]></time_of_preg>" & _
            strTab1 & "<breast_feeding><![CDATA[" & .strBreastFeeding & "]]></breast_feeding>" & _
            strTab1 & "<height><![CDATA[" & .strHeight & "]]></height>" & _
            strTab1 & "<weight><![CDATA[" & .strWeight & "]]></weight>" & _
            strTab1 & "<dialysis><![CDATA[" & .strDialysis & "]]></dialysis>" & _
            strTab1 & "<address><![CDATA[" & .strAddress & "]]></address>"
        strXML = strXML & _
            strTab1 & "<phone_no><![CDATA[" & .strPhoneNo & "]]></phone_no>" & _
            strTab1 & "<special_constitution><![CDATA[" & .strSpecialConstitution & "]]></special_constitution>" & _
            strTab1 & "<in_dept_id><![CDATA[" & .strINDeptId & "]]></in_dept_id>" & _
            strTab1 & "<in_dept_name><![CDATA[" & .strINDeptName & "]]></in_dept_name>" & _
            strTab1 & "<hospitalized_time><![CDATA[" & .strHospitalTime & "]]></hospitalized_time>" & _
            strTab1 & "<in_ward_id><![CDATA[" & .strInWardID & "]]></in_ward_id>" & _
            strTab1 & "<in_ward_name><![CDATA[" & .strInWardName & "]]></in_ward_name>" & _
            strTab1 & "<in_ward_bed_no><![CDATA[" & .strInWardBedNo & "]]></in_ward_bed_no>" & _
            strTab1 & "<in_condition><![CDATA[" & .strInConditon & "]]></in_condition>" & _
            strTab1 & "<weight_of_baby><![CDATA[" & .strWeight & "]]></weight_of_baby>" & _
            strTab1 & "<patient_condition><![CDATA[" & .strPatientConditon & "]]></patient_condition>" & _
            vbCrLf & _
        "</ipt_patient>"
    End With

    HZYY_MakeIPSPatient = strXML
End Function

Private Function HZYY_GetOPTPres(ByRef colPres As Collection, ByVal bytFunc As Byte) As String
'bytFunc  =0-�������;1-ɾ������
    Dim strXML      As String
    Dim strTab1     As String '
    Dim strTab2     As String
    Dim udtPres     As OptPrescription
    Dim udtItem     As OptPRESCRIPTIONSITEM
    Dim i           As Long
    Dim j           As Long
    
    
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    
    strXML = "<opt_prescriptions>" & strTab1 & "<opt_prescription>"
    For i = 1 To colPres.Count
        udtPres = colPres(i)
        strXML = strXML & HZYY_GetOPTPresInfo(udtPres.udtOptPresInfo, strTab2, bytFunc)
        If bytFunc = 0 Then
            For j = 1 To udtPres.colPresItem.Count
                udtItem = udtPres.colPresItem(j)
                strXML = strXML & HZYY_GetOPTPresItem(udtItem, strTab2)
            Next
        End If
    Next
    strXML = strXML & _
            strTab1 & "</opt_prescription>" & strTab1 & "</opt_prescriptions>"

    HZYY_GetOPTPres = strXML
End Function

Private Function HZYY_GetOPTPresInfo(ByRef udtInfo As OPTPRESCRIPTIONSINFO, ByVal strTab As String, Optional bytFunc As Byte) As String
'bytFunc  =0-�������;1-ɾ������
    Dim strXML As String
    
    With udtInfo
    If bytFunc = 0 Then
    strXML = strXML & _
            strTab & "<opt_prescription_info>" & _
            strTab & "<recipe_id><![CDATA[" & .strRecipeId & "]]></recipe_id>" & _
            strTab & "<recipe_no><![CDATA[" & .strRecipeNo & "]]></recipe_no>" & _
            strTab & "<recipe_source><![CDATA[" & .strRecipeSource & "]]></recipe_source>" & _
            strTab & "<recipe_category><![CDATA[" & .strRecipeCategory & "]]></recipe_category>" & _
            strTab & "<recipe_type><![CDATA[" & .strRecipeType & "]]></recipe_type>" & _
            strTab & "<dept_id><![CDATA[" & .strDeptID & "]]></dept_id>" & _
            strTab & "<dept_name><![CDATA[" & .strDeptName & "]]></dept_name>" & _
            strTab & "<recipe_doc_title><![CDATA[" & .strRecipeDocTitle & "]]></recipe_doc_title>" & _
            strTab & "<recipe_doc_id><![CDATA[" & .strRecipeDocId & "]]></recipe_doc_id>" & _
            strTab & "<recipe_doc_name><![CDATA[" & .strRecipeDocName & "]]></recipe_doc_name>" & _
            strTab & "<recipe_time><![CDATA[" & .strRecipeTime & "]]></recipe_time>" & _
            strTab & "<herb_unit_price><![CDATA[" & .strHerbUnitPrice & "]]></herb_unit_price>" & _
            strTab & "<herb_packet_count><![CDATA[" & .strHerbPacketCount & "]]></herb_packet_count>" & _
            strTab & "<is_cream><![CDATA[" & .strIsCream & "]]></is_cream>" & _
            strTab & "<recipe_fee_total><![CDATA[" & .strRecipeFeeTotal & "]]></recipe_fee_total>" & _
            strTab & "<original_recipe_id><![CDATA[" & .strOriginalRecipeId & "]]></original_recipe_id>" & _
            strTab & "<recipe_status><![CDATA[" & .strRecipeStatus & "]]></recipe_status>" & _
            strTab & "<urgent_flag><![CDATA[" & .strUrgentFlag & "]]></urgent_flag>"
    strXML = strXML & _
            strTab & "<review_pharm_id><![CDATA[" & .strCheckPharmID & "]]></review_pharm_id>" & _
            strTab & "<review_pharm_name><![CDATA[" & .strCheckPharmName & "]]></review_pharm_name>" & _
            strTab & "<review_pharm_title><![CDATA[" & .strCheckPharmTitle & "]]></review_pharm_title>" & _
            strTab & "<prep_pharm_id><![CDATA[" & .strPrepPharmId & "]]></prep_pharm_id>" & _
            strTab & "<prep_pharm_name><![CDATA[" & .strPrepPharmName & "]]></prep_pharm_name>" & _
            strTab & "<prep_pharm_title><![CDATA[" & .strPrepPharmTitle & "]]></prep_pharm_title>" & _
            strTab & "<check_pharm_id><![CDATA[" & .strCheckPharmID & "]]></check_pharm_id>" & _
            strTab & "<check_pharm_name><![CDATA[" & .strCheckPharmName & "]]></check_pharm_name>" & _
            strTab & "<check_pharm_title><![CDATA[" & .strCheckPharmTitle & "]]></check_pharm_title>" & _
            strTab & "<despensing_pharm_id><![CDATA[" & .strDespensingPharmId & "]]></despensing_pharm_id>" & _
            strTab & "<despensing_pharm_name><![CDATA[" & .strDespensingPharmName & "]]></despensing_pharm_name>" & _
            strTab & "<despensing_pharm_title><![CDATA[" & .strDespensingPharmTitle & "]]></despensing_pharm_title>" & _
            strTab & "</opt_prescription_info>"
    Else
        strXML = strXML & _
                strTab & "<opt_prescription_info>" & _
                "<recipe_id><![CDATA[" & .strRecipeId & "]]></recipe_id>" & _
                "<recipe_no><![CDATA[" & .strRecipeNo & "]]></recipe_no>" & _
                strTab & "</opt_prescription_info>"
    End If
    End With
    
    HZYY_GetOPTPresInfo = strXML
End Function

Private Function HZYY_GetOPTPresItem(ByRef udtItem As OptPRESCRIPTIONSITEM, ByVal strTab As String) As String
    Dim strXML As String

    With udtItem
        strXML = strXML & _
                strTab & "<opt_prescription_item>" & _
                strTab & "<recipe_item_id><![CDATA[" & .strRecipeItemId & "]]></recipe_item_id>" & _
                strTab & "<recipe_id><![CDATA[" & .strRecipeId & "]]></recipe_id>" & _
                strTab & "<drug_purpose><![CDATA[" & .strDrugPurpose & "]]></drug_purpose>" & _
                strTab & "<group_no><![CDATA[" & .strGroupNO & "]]></group_no>" & _
                strTab & "<drug_id><![CDATA[" & .strDrugID & "]]></drug_id>" & _
                strTab & "<drug_name><![CDATA[" & .strDrugName & "]]></drug_name>" & _
                strTab & "<count_unit><![CDATA[" & .strCountUnit & "]]></count_unit>" & _
                strTab & "<pack_unit><![CDATA[" & .strPackUnit & "]]></pack_unit>" & _
                strTab & "<manufacturer_id><![CDATA[" & .strManufacturerID & "]]></manufacturer_id>" & _
                strTab & "<manufacturer_name><![CDATA[" & .strManufacturerName & "]]></manufacturer_name>" & _
                strTab & "<drug_dose><![CDATA[" & .strDrugdose & "]]></drug_dose>" & _
                strTab & "<drug_admin_route_name><![CDATA[" & .strDrugadminRouteName & "]]></drug_admin_route_name>" & _
                strTab & "<drug_using_freq><![CDATA[" & .strDrugUsingFreq & "]]></drug_using_freq>" & _
                strTab & "<drug_using_time_point><![CDATA[" & .strDrugUsingTimePoint & "]]></drug_using_time_point>" & _
                strTab & "<drug_using_aim><![CDATA[" & .strDrugUsingAim & "]]></drug_using_aim>" & _
                strTab & "<drug_using_area><![CDATA[" & .strDrugUsingArea & "]]></drug_using_area>" & _
                strTab & "<duration><![CDATA[" & .strDuration & "]]></duration>" & _
                strTab & "<preparation><![CDATA[" & .strPreparation & "]]></preparation>"
        strXML = strXML & _
                strTab & "<specification><![CDATA[" & .strSpecification & "]]></specification>" & _
                strTab & "<unit_price><![CDATA[" & .strUnitPrice & "]]></unit_price>" & _
                strTab & "<despensing_num><![CDATA[" & .strDespensingNum & "]]></despensing_num>" & _
                strTab & "<fee_total><![CDATA[" & .strFeeTotal & "]]></fee_total>" & _
                strTab & "<start_time><![CDATA[" & .strStartTime & "]]></start_time>" & _
                strTab & "<end_time><![CDATA[" & .strEndTime & "]]></end_time>" & _
                strTab & "<special_prompt><![CDATA[" & .strSpecialPrompt & "]]></special_prompt>" & _
                strTab & "<skin_test_flag><![CDATA[" & .strSkinTestFlag & "]]></skin_test_flag>" & _
                strTab & "<skin_test_result><![CDATA[" & .strSkinTestResult & "]]></skin_test_result>" & _
                strTab & "<skin_test_time><![CDATA[" & .strSkinTestTime & "]]></skin_test_time>" & _
                strTab & "<drug_source><![CDATA[" & .strDrugSource & "]]></drug_source>" & _
                strTab & "<drug_return_flag><![CDATA[" & .strdrugReturnFlag & "]]></drug_return_flag>" & _
                strTab & "<ouvas_flag><![CDATA[" & .strOuvasFlag & "]]></ouvas_flag>" & _
                strTab & "<dripping_speed><![CDATA[" & .strDrippingSpeed & "]]></dripping_speed>" & _
                strTab & "<limit_time><![CDATA[" & .strLimitTime & "]]></limit_time>" & _
                strTab & "<therapeutic_regimen><![CDATA[" & .strTherapeuticRegimen & "]]></therapeutic_regimen>" & _
                strTab & "<dispensing_window><![CDATA[" & .strDispensingWindow & "]]></dispensing_window>" & _
                strTab & "<drug_store_area><![CDATA[" & .strDrugstoreArea & "]]></drug_store_area>" & _
                strTab & "</opt_prescription_item>"
    End With

    HZYY_GetOPTPresItem = strXML
End Function

Private Function HZYY_GetOrder(ByRef udtOrder As Order, Optional ByVal bytType As Byte) As String
'����:��ȡҽ����ϢXML
'bytType=0-��Ԥ;1-ɾ��

    Dim strXML      As String
    Dim strTab1     As String '
    Dim strTab2     As String
    Dim i           As Long
    Dim j           As Long
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    With udtOrder
        strXML = "<orders>"
        strXML = strXML & HZYY_GetOrderNonMedical(.colNonMedical, strTab2, bytType)
        strXML = strXML & HZYY_GetOrderMedical(.colMedical, strTab2, bytType)
        strXML = strXML & HZYY_GetOrderHerbMedical(.colHerbMedical, strTab2, bytType)
    strXML = strXML & vbCrLf & "</orders>"
    End With
    HZYY_GetOrder = strXML
End Function

Private Function HZYY_GetOrderNonMedical(ByRef colNonMed As Collection, ByVal strTab As String, Optional bytFunc As Byte) As String
    Dim strXML          As String
    Dim udtNonMed       As NonMedicalOrderItem
    Dim i               As Long

    If colNonMed Is Nothing Then Exit Function
    With udtNonMed
        If bytFunc = 0 Then
            For i = 1 To colNonMed.Count
            strXML = strXML & _
                strTab & "<non_medical_order_item>" & _
                strTab & "<order_id><![CDATA[" & .strOrderId & "]]></order_id>" & _
                strTab & "<order_time><![CDATA[" & .strOrderTime & "]]></order_time>" & _
                strTab & "<order_dept_id><![CDATA[" & .strOrderDeptID & "]]></order_dept_id>" & _
                strTab & "<order_dept_name><![CDATA[" & .strOrderDeptName & "]]></order_dept_name>" & _
                strTab & "<doc_group><![CDATA[" & .strDocGroup & "]]></doc_group>" & _
                strTab & "<order_doc_name><![CDATA[" & .strOrderDocName & "]]></order_doc_name>" & _
                strTab & "<order_doc_id><![CDATA[" & .strOrderDocID & "]]></order_doc_id>" & _
                strTab & "<order_doc_title><![CDATA[" & .strOrderDocTitle & "]]></order_doc_title>" & _
                strTab & "<order_type><![CDATA[" & .strOrderType & "]]></order_type>" & _
                strTab & "<order_code><![CDATA[" & .strOrderCode & "]]></order_code>" & _
                strTab & "<order_name><![CDATA[" & .strOrderName & "]]></order_name>" & _
                strTab & "<order_category><![CDATA[" & .strOrderCategory & "]]></order_category>" & _
                strTab & "<order_freq><![CDATA[" & .strOrderFreq & "]]></order_freq>" & _
                strTab & "<order_valid_time><![CDATA[" & .strOrderValidTime & "]]></order_valid_time>" & _
                strTab & "<order_invalid_time><![CDATA[" & .strOrderInvalidTime & "]]></order_invalid_time>" & _
                strTab & "<duration><![CDATA[" & .strDuration & "]]></duration>" & _
                strTab & "<check_time><![CDATA[" & .strCheckTime & "]]></check_time>" & _
                strTab & "<check_nurse_id><![CDATA[" & .strCheckNurseID & "]]></check_nurse_id>" & _
                strTab & "<check_nurse_name><![CDATA[" & .strCheckNurseName & "]]></check_nurse_name>" & _
                strTab & "<stop_flag><![CDATA[" & .strStopFlag & "]]></stop_flag>" & _
                strTab & "</non_medical_order_item>"
            Next
        Else
            For i = 1 To colNonMed.Count
            strXML = strXML & _
                strTab & "<non_medical_order_item>" & _
                strTab & "<order_id><![CDATA[" & .strOrderId & "]]></order_id>" & _
                strTab & "</non_medical_order_item>"
            Next
        End If
    End With

    HZYY_GetOrderNonMedical = strXML
End Function

Private Function HZYY_GetOrderMedical(ByRef colMed As Collection, ByVal strTab As String, Optional bytFunc As Byte) As String
'����:bytFunc=0 ҽ�����;1-ɾ��
    Dim strXML      As String
    Dim udtMed   As MedicalOrderItem
    Dim i           As Long
    
    If colMed Is Nothing Then Exit Function
    For i = 1 To colMed.Count
        udtMed = colMed(i)
        With udtMed
        If bytFunc = 0 Then
            strXML = strXML & _
                strTab & "<medical_order_item>" & _
                strTab & "<order_id><![CDATA[" & .strOrderId & "]]></order_id>" & _
                strTab & "<order_time><![CDATA[" & .strOrderTime & "]]></order_time>" & _
                strTab & "<order_dept_id><![CDATA[" & .strOrderDeptID & "]]></order_dept_id>" & _
                strTab & "<order_dept_name><![CDATA[" & .strOrderDeptName & "]]></order_dept_name>" & _
                strTab & "<doc_group><![CDATA[" & .strDocGroup & "]]></doc_group>" & _
                strTab & "<order_doc_id><![CDATA[" & .strOrderDocID & "]]></order_doc_id>" & _
                strTab & "<order_doc_name><![CDATA[" & .strOrderDocName & "]]></order_doc_name>" & _
                strTab & "<order_doc_title><![CDATA[" & .strOrderDocTitle & "]]></order_doc_title>" & _
                strTab & "<order_type><![CDATA[" & .strOrderType & "]]></order_type>" & _
                strTab & "<drug_purpose><![CDATA[" & .strDrugPurpose & "]]></drug_purpose>" & _
                strTab & "<group_no><![CDATA[" & .strGroupNO & "]]></group_no>" & _
                strTab & "<drug_id><![CDATA[" & .strDrugID & "]]></drug_id>" & _
                strTab & "<drug_name><![CDATA[" & .strDrugName & "]]></drug_name>" & _
                strTab & "<count_unit><![CDATA[" & .strCountUnit & "]]></count_unit>" & _
                strTab & "<pack_unit><![CDATA[" & .strPackUnit & "]]></pack_unit>" & _
                strTab & "<manufacturer_id><![CDATA[" & .strManufacturerID & "]]></manufacturer_id>" & _
                strTab & "<manufacturer_name><![CDATA[" & .strManufacturerName & "]]></manufacturer_name>" & _
                strTab & "<drug_dose><![CDATA[" & .strDrugdose & "]]></drug_dose>" & _
                strTab & "<drug_admin_route_name><![CDATA[" & .strDrugadminRouteName & "]]></drug_admin_route_name>" & _
                strTab & "<drug_using_freq><![CDATA[" & .strDrugUsingFreq & "]]></drug_using_freq>"
            strXML = strXML & _
                strTab & "<drug_using_time_point><![CDATA[" & .strDrugUsingTimePoint & "]]></drug_using_time_point>" & _
                strTab & "<drug_using_aim><![CDATA[" & .strDrugUsingAim & "]]></drug_using_aim>" & _
                strTab & "<drug_using_area><![CDATA[" & .strDrugUsingArea & "]]></drug_using_area>" & _
                strTab & "<drug_source><![CDATA[" & .strDrugSource & "]]></drug_source>" & _
                strTab & "<duration><![CDATA[" & .strDuration & "]]></duration>" & _
                strTab & "<preparation><![CDATA[" & .strPreparation & "]]></preparation>" & _
                strTab & "<specifications><![CDATA[" & .strSpecifications & "]]></specifications>" & _
                strTab & "<unit_price><![CDATA[" & .strUnitPrice & "]]></unit_price>" & _
                strTab & "<despensing_num><![CDATA[" & .strDespensingNum & "]]></despensing_num>" & _
                strTab & "<fee_total><![CDATA[" & .strFeeTotal & "]]></fee_total>" & _
                strTab & "<check_time><![CDATA[" & .strCheckTime & "]]></check_time>" & _
                strTab & "<check_nurse_id><![CDATA[" & .strCheckNurseID & "]]></check_nurse_id>"
            strXML = strXML & _
                strTab & "<check_nurse_name><![CDATA[" & .strCheckNurseName & "]]></check_nurse_name>" & _
                strTab & "<order_valid_time><![CDATA[" & .strOrderValidTime & "]]></order_valid_time>" & _
                strTab & "<order_invalid_time><![CDATA[" & .strOrderInvalidTime & "]]></order_invalid_time>" & _
                strTab & "<special_prompt><![CDATA[" & .strSpecialPrompt & "]]></special_prompt>" & _
                strTab & "<skin_test_time><![CDATA[" & .strSkinTestTime & "]]></skin_test_time>" & _
                strTab & "<skin_test_flag><![CDATA[" & .strSkinTestFlag & "]]></skin_test_flag>" & _
                strTab & "<skin_test_result><![CDATA[" & .strSkinTestResult & "]]></skin_test_result>" & _
                strTab & "<drug_return_flag><![CDATA[" & .strdrugReturnFlag & "]]></drug_return_flag>" & _
                strTab & "<stop_flag><![CDATA[" & .strStopFlag & "]]></stop_flag>" & _
                strTab & "<pivas_flag><![CDATA[" & .strPivasFlag & "]]></pivas_flag>" & _
                strTab & "<urgent_flag><![CDATA[" & .strUrgentFlag & "]]></urgent_flag>" & _
                strTab & "<dripping_speed><![CDATA[" & .strDrippingSpeed & "]]></dripping_speed>" & _
                strTab & "<limit_time><![CDATA[" & .strLimitTime & "]]></limit_time>" & _
                strTab & "<therapeutic_regimen><![CDATA[" & .strTherapeuticRegimen & "]]></therapeutic_regimen>" & _
                strTab & "<exe_dept_id><![CDATA[" & .strExeDeptID & "]]></exe_dept_id>" & _
                strTab & "<exe_dept_name><![CDATA[" & .strExeDeptName & "]]></exe_dept_name>" & _
                strTab & "<dispensing_window><![CDATA[" & .strDispensingWindow & "]]></dispensing_window>" & _
                strTab & "<drug_store_area><![CDATA[" & .strDrugstoreArea & "]]></drug_store_area>" & _
                strTab & "</medical_order_item>"
                
            Else
                'ɾ��ҽ��
                strXML = strXML & _
                    strTab & "<medical_order_item>" & _
                    strTab & "<order_id><![CDATA[" & .strOrderId & "]]></order_id>" & _
                    strTab & "<group_no><![CDATA[" & .strGroupNO & "]]></group_no>" & _
                    strTab & "</medical_order_item>"
            End If
        End With
    Next

    HZYY_GetOrderMedical = strXML
End Function

Private Function HZYY_GetOrderHerbMedical(ByRef colHerb As Collection, ByVal strTab As String, Optional bytFunc As Byte) As String
'����:bytFunc=0 ҽ�����;1-ɾ��
    Dim strXML          As String
    Dim strTemp         As String
    Dim udtHerb         As HerbMedicalOrder
    Dim udtHerbItem     As HerbMedicalOrderItem
    Dim i               As Long
    Dim j               As Long
    If colHerb Is Nothing Then Exit Function
    For i = 1 To colHerb.Count
        strXML = strXML & "<herb_medical_order>"
        udtHerb = colHerb(i)
        If bytFunc = 0 Then
            With udtHerb.udtHerbInfo
                strXML = strXML & _
                strTab & "<herb_medical_order_info>" & _
                strTab & "<order_id><![CDATA[" & .strOrderId & "]]></order_id>" & _
                strTab & "<order_time><![CDATA[" & .strOrderTime & "]]></order_time>" & _
                strTab & "<order_dept_id><![CDATA[" & .strOrderDeptID & "]]></order_dept_id>" & _
                strTab & "<order_dept_name><![CDATA[" & .strOrderDeptName & "]]></order_dept_name>" & _
                strTab & "<doc_group><![CDATA[" & .strDocGroup & "]]></doc_group>" & _
                strTab & "<order_doc_id><![CDATA[" & .strOrderDocID & "]]></order_doc_id>" & _
                strTab & "<order_doc_name><![CDATA[" & .strOrderDocName & "]]></order_doc_name>" & _
                strTab & "<order_doc_title><![CDATA[" & .strOrderDocTitle & "]]></order_doc_title>" & _
                strTab & "<order_type><![CDATA[" & .strOrderType & "]]></order_type>" & _
                strTab & "<herb_unit_price><![CDATA[" & .strHerbUnitPrice & "]]></herb_unit_price>" & _
                strTab & "<herb_packet_count><![CDATA[" & .strHerbPacketCount & "]]></herb_packet_count>" & _
                strTab & "<is_cream><![CDATA[" & .strIsCream & "]]></is_cream>" & _
                strTab & "<check_time><![CDATA[" & .strCheckTime & "]]></check_time>" & _
                strTab & "<check_nurse_id><![CDATA[" & .strCheckNurseID & "]]></check_nurse_id>" & _
                strTab & "<check_nurse_name><![CDATA[" & .strCheckNurseName & "]]></check_nurse_name>" & _
                strTab & "<order_valid_time><![CDATA[" & .strOrderValidTime & "]]></order_valid_time>" & _
                strTab & "<order_invalid_time><![CDATA[" & .strOrderInvalidTime & "]]></order_invalid_time>" & _
                strTab & "<drug_return_flag><![CDATA[" & .strdrugReturnFlag & "]]></drug_return_flag>" & _
                strTab & "<stop_flag><![CDATA[" & .strStopFlag & "]]></stop_flag>" & _
                strTab & "<urgent_flag><![CDATA[" & .strUrgentFlag & "]]></urgent_flag>" & _
                strTab & "<exe_dept_id><![CDATA[" & .strExeDeptID & "]]></exe_dept_id>" & _
                strTab & "<exe_dept_name><![CDATA[" & .strExeDeptName & "]]></exe_dept_name>" & _
                strTab & "</herb_medical_order_info>"
            End With
            
            For j = 1 To udtHerb.colItemHerb.Count
                udtHerbItem = udtHerb.colItemHerb(j)
                With udtHerbItem
                    strTemp = "<herb_medical_order_item>" & _
                    strTab & "<order_id><![CDATA[" & .strOrderId & "]]></order_id>" & _
                    strTab & "<order_item_id><![CDATA[" & .strOrderitemID & "]]></order_item_id>" & _
                    strTab & "<group_no><![CDATA[" & .strGroupNO & "]]></group_no>" & _
                    strTab & "<drug_id><![CDATA[" & .strDrugID & "]]></drug_id>" & _
                    strTab & "<drug_name><![CDATA[" & .strDrugName & "]]></drug_name>" & _
                    strTab & "<manufacturer_id><![CDATA[" & .strManufacturerID & "]]></manufacturer_id>" & _
                    strTab & "<manufacturer_name><![CDATA[" & .strManufacturerName & "]]></manufacturer_name>" & _
                    strTab & "<drug_dose><![CDATA[" & .strDrugdose & "]]></drug_dose>" & _
                    strTab & "<drug_admin_route_name><![CDATA[" & .strDrugadminRouteName & "]]></drug_admin_route_name>" & _
                    strTab & "<drug_using_freq><![CDATA[" & .strDrugUsingFreq & "]]></drug_using_freq>" & _
                    strTab & "<preparation><![CDATA[" & .strPreparation & "]]></preparation>" & _
                    strTab & "<specifications><![CDATA[" & .strSpecifications & "]]></specifications>" & _
                    strTab & "<unit_price><![CDATA[" & .strUnitPrice & "]]></unit_price>" & _
                    strTab & "<despensing_num><![CDATA[" & .strDespensingNum & "]]></despensing_num>" & _
                    strTab & "<fee_total><![CDATA[" & .strFeeTotal & "]]></fee_total>" & _
                    strTab & "<special_prompt><![CDATA[" & .strSpecialPrompt & "]]></special_prompt>" & _
                    "</herb_medical_order_item>"
                End With
               strXML = strXML & vbCrLf & strTemp
            Next
            strXML = strXML & vbTab & "</herb_medical_order>"
        Else
            With udtHerb.udtHerbInfo
                strXML = strXML & _
                strTab & "<herb_medical_order_info>" & _
                strTab & "<order_id><![CDATA[" & .strOrderId & "]]></order_id>" & _
                strTab & "</herb_medical_order_info>"
            End With
            strXML = strXML & vbTab & "</herb_medical_order>"
        End If
    Next
    HZYY_GetOrderHerbMedical = strXML
End Function

Public Function HZYY_GetDiag(ByRef colDiag As Collection, Optional bytFunc As Byte) As String
'����:��ȡ�����ϢXML
'����: bytFunc=0  �������;=1 סԺ���
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtDiag     As Diagnosis
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = IIf(bytFunc = 0, "<opt_diagnoses>", "<ipt_diagnoses>")
    For i = 1 To colDiag.Count
        udtDiag = colDiag(i)
        With udtDiag
            strXML = strXML & IIf(bytFunc = 0, "<opt_diagnosis>", "<ipt_diagnosis>") & _
            strTab1 & "<diag_id><![CDATA[" & .strDiagID & "]]></diag_id>" & _
            strTab1 & "<diag_dept_id><![CDATA[" & .strDiagDeptID & "]]></diag_dept_id>" & _
            strTab1 & "<diag_dept_name><![CDATA[" & .strDiagDeptName & "]]></diag_dept_name>" & _
            strTab1 & "<diag_doc_id><![CDATA[" & .strDiagDocID & "]]></diag_doc_id>" & _
            strTab1 & "<diag_doc_name><![CDATA[" & .strDiagDocName & "]]></diag_doc_name>" & _
            strTab1 & "<diag_doc_title><![CDATA[" & .strDiagDocTitle & "]]></diag_doc_title>" & _
            strTab1 & "<diag_date><![CDATA[" & .strDiagDate & "]]></diag_date>" & _
            strTab1 & "<diag_category><![CDATA[" & .strDiagCategory & "]]></diag_category>" & _
            strTab1 & "<diag_type><![CDATA[" & .strDiagType & "]]></diag_type>" & _
            strTab1 & "<diag_name><![CDATA[" & .strDiagName & "]]></diag_name>" & _
            strTab1 & "<diag_code><![CDATA[" & .strDiagCode & "]]></diag_code>" & _
            strTab1 & "<diag_code_type><![CDATA[" & .strDiagCodeType & "]]></diag_code_type>" & _
            strTab1 & "<disease_classification><![CDATA[" & .strDiseaseClassification & "]]></disease_classification>" & _
            strTab1 & "<disease_staging><![CDATA[" & .strDiseaseStaging & "]]></disease_staging>" & _
            strTab1 & "<disease_score><![CDATA[" & .strDiseaseScore & "]]></disease_score>"
            strXML = strXML & IIf(bytFunc = 0, "</opt_diagnosis>", "</ipt_diagnosis>")
        End With
    Next
    strXML = strXML & vbCrLf & IIf(bytFunc = 0, "</opt_diagnoses>", "</ipt_diagnoses>")
    HZYY_GetDiag = strXML
End Function

Public Function HZYY_GetAllergies(ByRef colAllergy As Collection, Optional bytFunc As Byte) As String
'����:��ȡ������ϢXML
'����: bytFunc=0  �������;=1 סԺ���
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtAllergy     As Allergy
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = IIf(bytFunc = 0, "<opt_allergies>", "<ipt_allergies>")
    For i = 1 To colAllergy.Count
        udtAllergy = colAllergy(i)
        With udtAllergy
            strXML = strXML & IIf(bytFunc = 0, "<opt_allergy>", "<ipt_allergy>") & _
            strTab1 & "<allergy_id><![CDATA[" & .strAllergyID & "]]></allergy_id>" & _
            strTab1 & "<allergy_drug><![CDATA[" & .strAllergyDrug & "]]></allergy_drug>" & _
            strTab1 & "<anaphylaxis><![CDATA[" & .strAnaphylaxis & "]]></anaphylaxis>" & _
            strTab1 & "<record_time><![CDATA[" & .strRecordTime & "]]></record_time>"
            strXML = strXML & IIf(bytFunc = 0, "</opt_allergy>", "</ipt_allergy>")
        End With
    Next
    strXML = strXML & vbCrLf & IIf(bytFunc = 0, "</opt_allergies>", "</ipt_allergies>")
    HZYY_GetAllergies = strXML
End Function

Public Function HZYY_GetOperations(ByRef colOper As Collection, Optional bytFunc As Byte) As String
'����:��ȡ������ϢXML
'����: bytFunc=0  �������;=1 סԺ���
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtOper     As Operation
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = IIf(bytFunc = 0, "<opt_operations>", "<ipt_operations>")
    For i = 1 To colOper.Count
        udtOper = colOper(i)
        With udtOper
            strXML = strXML & IIf(bytFunc = 0, "<opt_operation>", "<ipt_operation>") & _
            strTab1 & "<operation_id><![CDATA[" & .strOperationID & "]]></operation_id>" & _
            strTab1 & "<operation_code><![CDATA[" & .strOperationCode & "]]></operation_code>" & _
            strTab1 & "<operation_name><![CDATA[" & .strOperationName & "]]></operation_name>" & _
            strTab1 & "<dept_id><![CDATA[" & .strDeptID & "]]></dept_id>" & _
            strTab1 & "<dept_name><![CDATA[" & .strDeptName & "]]></dept_name>" & _
            strTab1 & "<operation_start_time><![CDATA[" & .strOperationStartTime & "]]></operation_start_time>" & _
            strTab1 & "<operation_end_time><![CDATA[" & .strOperationEndTime & "]]></operation_end_time>" & _
            strTab1 & "<operation_incision_type><![CDATA[" & .strOperationIncisionType & "]]></operation_incision_type>" & _
            strTab1 & "<anesthesia_code><![CDATA[" & .strAnesthesiaCode & "]]></anesthesia_code>" & _
            strTab1 & "<asa><![CDATA[" & .strAsa & "]]></asa>" & _
            strTab1 & "<anesthesia_end_time><![CDATA[" & .strAnesthesiaEndTime & "]]></anesthesia_end_time>" & _
            strTab1 & "<anesthesia_start_time><![CDATA[" & .strAnesthesiaStartTime & "]]></anesthesia_start_time>" & _
            strTab1 & "<is_implant><![CDATA[" & .strIsImplant & "]]></is_implant>" & _
            strTab1 & "<implant_no><![CDATA[" & .strImplantNO & "]]></implant_no>" & _
            strTab1 & "<implant_name><![CDATA[" & .strImplantName & "]]></implant_name>" & _
            strTab1 & "<is_reoperation><![CDATA[" & .strIsReOperation & "]]></is_reoperation>" & _
            strTab1 & "<operation_doc_id><![CDATA[" & .strOperationDocID & "]]></operation_doc_id>" & _
            strTab1 & "<operation_doc_name><![CDATA[" & .strOperationDocName & "]]></operation_doc_name>" & _
            strTab1 & "<operation_level><![CDATA[" & .strOperationlevel & "]]></operation_level>" & _
            strTab1 & "<operation_site_code><![CDATA[" & .strOperationSiteCode & "]]></operation_site_code>" & _
            strTab1 & "<hemorrhage_volume><![CDATA[" & .strhemorrhageVolume & "]]></hemorrhage_volume>" & _
            strTab1 & "<operation_source><![CDATA[" & .strOperationSource & "]]></operation_source>" & _
            strTab1 & "<pre_op_diag_code><![CDATA[" & .strpreOPDiagCode & "]]></pre_op_diag_code>" & _
            strTab1 & "<pre_op_diag_name><![CDATA[" & .strpreOPDiagName & "]]></pre_op_diag_name>"
            strXML = strXML & _
            strTab1 & "<post_op_diag_code><![CDATA[" & .strpostOPDiagCode & "]]></post_op_diag_code>" & _
            strTab1 & "<post_op_diag_name><![CDATA[" & .strpostOPDiagName & "]]></post_op_diag_name>" & _
            strTab1 & "<nnis><![CDATA[" & .strNnis & "]]></nnis>" & _
            strTab1 & "<is_selective_operation><![CDATA[" & .strisSelectiveOperation & "]]></is_selective_operation>"
            
            strXML = strXML & IIf(bytFunc = 0, "</opt_operation>", "</ipt_operation>")
        End With
    Next
    strXML = strXML & vbCrLf & IIf(bytFunc = 0, "</opt_operations>", "</ipt_operations>")
    HZYY_GetOperations = strXML
End Function

Public Function HZYY_GetExams(ByRef colExams As Collection, Optional bytFunc As Byte) As String
'����:��ȡ������ϢXML
'����: bytFunc=0  �������;=1 סԺ���
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtExam     As HZYYExam
    Dim udtItem     As ExamItem
    Dim i           As Long
    Dim j           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = IIf(bytFunc = 0, "<opt_exams>", "<ipt_exams>")
    For i = 1 To colExams.Count
        udtExam = colExams(i)
        With udtExam.udtInfo
            strXML = strXML & IIf(bytFunc = 0, "<opt_exam_info>", "<ipt_exam_info>") & _
            strTab1 & "<report_id><![CDATA[" & .strReportID & "]]></report_id>" & _
            strTab1 & "<exam_item_code><![CDATA[" & .strExamItemCode & "]]></exam_item_code>" & _
            strTab1 & "<exam_item_name><![CDATA[" & .strExamItemName & "]]></exam_item_name>" & _
            strTab1 & "<sample_collect_time><![CDATA[" & .strSampleCollectTime & "]]></sample_collect_time>" & _
            strTab1 & "<sample_code><![CDATA[" & .strSampleCode & "]]></sample_code>" & _
            strTab1 & "<sample_name><![CDATA[" & .strSampleName & "]]></sample_name>" & _
            strTab1 & "<sample_collect_opporunity><![CDATA[" & .strSampleCollectOpporunity & "]]></sample_collect_opporunity>" & _
            strTab1 & "<apply_no><![CDATA[" & .strApplyNO & "]]></apply_no>" & _
            strTab1 & "<applicant_id><![CDATA[" & .strApplicantID & "]]></applicant_id>" & _
            strTab1 & "<applicant_name><![CDATA[" & .strApplicantName & "]]></applicant_name>" & _
            strTab1 & "<applicant_dept_id><![CDATA[" & .strApplicantDeptID & "]]></applicant_dept_id>" & _
            strTab1 & "<applicant_dept_name><![CDATA[" & .strApplicantDeptName & "]]></applicant_dept_name>" & _
            strTab1 & "<reporter_id><![CDATA[" & .strReporterID & "]]></reporter_id>" & _
            strTab1 & "<reporter_name><![CDATA[" & .strReporterName & "]]></reporter_name>" & _
            strTab1 & "<report_time><![CDATA[" & .strReportTime & "]]></report_time>" & _
            strTab1 & "<exam_aim><![CDATA[" & .strExamAim & "]]></exam_aim>"
            strXML = strXML & IIf(bytFunc = 0, "</opt_exam_info>", "</ipt_exam_info>")
        End With
        
        For j = 1 To udtExam.colExamItem.Count
            udtItem = udtExam.colExamItem(j)
            With udtItem
                strXML = strXML & IIf(bytFunc = 0, "<opt_exam_item>", "<ipt_exam_item>") & _
                    strTab1 & "<report_id><![CDATA[" & .strReportID & "]]></report_id>" & _
                    strTab1 & "<report_item_id><![CDATA[" & .strReportItemID & "]]></report_item_id>" & _
                    strTab1 & "<indicator_code><![CDATA[" & .strindicatorCode & "]]></indicator_code>" & _
                    strTab1 & "<indicator_name><![CDATA[" & .strindicatorName & "]]></indicator_name>" & _
                    strTab1 & "<indicator_ename><![CDATA[" & .strindicatorename & "]]></indicator_ename>" & _
                    strTab1 & "<exam_result><![CDATA[" & .strExamResult & "]]></exam_result>" & _
                    strTab1 & "<exam_result_unit><![CDATA[" & .strExamResultUnit & "]]></exam_result_unit>" & _
                    strTab1 & "<reference_result><![CDATA[" & .strreferenceResult & "]]></reference_result>" & _
                    strTab1 & "<upper_limit><![CDATA[" & .strupperlimit & "]]></upper_limit>" & _
                    strTab1 & "<lower_limit><![CDATA[" & .strlowerlimit & "]]></lower_limit>" & _
                    strTab1 & "<critical_flag><![CDATA[" & .strcriticalFlag & "]]></critical_flag>"
                strXML = strXML & IIf(bytFunc = 0, "</opt_exam_item>", "</ipt_exam_item>")
            End With
        Next
    Next
    strXML = strXML & vbCrLf & IIf(bytFunc = 0, "</opt_exams>", "</ipt_exams>")
    HZYY_GetExams = strXML
End Function

Public Function HZYY_GetImageInfo(ByRef colImageInfo As Collection, Optional bytFunc As Byte) As String
'����:��ȡӰ����ϢXML
'����: bytFunc=0  �������;=1 סԺ���
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtImage     As ImageInfo
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = IIf(bytFunc = 0, "<opt_image_infos>", "<ipt_image_infos>")
    For i = 1 To colImageInfo.Count
        udtImage = colImageInfo(i)
        With udtImage
            strXML = strXML & IIf(bytFunc = 0, "<opt_image_info>", "<ipt_image_info>") & _
            strTab1 & "<image_id><![CDATA[" & .strImageID & "]]></image_id>" & _
            strTab1 & "<image_code><![CDATA[" & .strImageCode & "]]></image_code>" & _
            strTab1 & "<image_name><![CDATA[" & .strImageName & "]]></image_name>" & _
            strTab1 & "<perform_method><![CDATA[" & .strperformMethod & "]]></perform_method>" & _
            strTab1 & "<perform_site><![CDATA[" & .strperformSite & "]]></perform_site>" & _
            strTab1 & "<imaging_position><![CDATA[" & .strimagingPosition & "]]></imaging_position>" & _
            strTab1 & "<imaging_diagnosis><![CDATA[" & .strimagingDiagnosis & "]]></imaging_diagnosis>" & _
            strTab1 & "<apply_no><![CDATA[" & .strApplyNO & "]]></apply_no>" & _
            strTab1 & "<applicant_id><![CDATA[" & .strApplicantID & "]]></applicant_id>" & _
            strTab1 & "<applicant_name><![CDATA[" & .strApplicantName & "]]></applicant_name>" & _
            strTab1 & "<applicant_dept_id><![CDATA[" & .strApplicantDeptID & "]]></applicant_dept_id>" & _
            strTab1 & "<applicant_dept_name><![CDATA[" & .strApplicantDeptName & "]]></applicant_dept_name>" & _
            strTab1 & "<perform_time><![CDATA[" & .strPerformTime & "]]></perform_time>" & _
            strTab1 & "<reporter_id><![CDATA[" & .strReporterID & "]]></reporter_id>" & _
            strTab1 & "<reporter_name><![CDATA[" & .strReporterName & "]]></reporter_name>" & _
            strTab1 & "<report_time><![CDATA[" & .strReportTime & "]]></report_time>" & _
            strTab1 & "<perform_aim><![CDATA[" & .strperformAim & "]]></perform_aim>"
            strXML = strXML & IIf(bytFunc = 0, "</opt_image_info>", "</ipt_image_info>")
        End With
    Next
    strXML = strXML & vbCrLf & IIf(bytFunc = 0, "</opt_image_infos>", "</ipt_image_infos>")
    HZYY_GetImageInfo = strXML
End Function

Public Function HZYY_SpecialExams(ByRef colSpExam As Collection, Optional bytFunc As Byte) As String
'����:��������Ŀ��ǩ
'����: bytFunc=0  �������;=1 סԺ���
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtItem     As SpecialExam
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = IIf(bytFunc = 0, "<opt_special_exams>", "<ipt_special_exams>")
    For i = 1 To colSpExam.Count
        udtItem = colSpExam(i)
        With udtItem
            strXML = strXML & IIf(bytFunc = 0, "<opt_special_exam>", "<ipt_special_exam>") & _
            strTab1 & "<report_id><![CDATA[" & .strReportID & "]]></report_id>" & _
            strTab1 & "<exam_item_code><![CDATA[" & .strExamItemCode & "]]></exam_item_code>" & _
            strTab1 & "<exam_item_name><![CDATA[" & .strExamItemName & "]]></exam_item_name>" & _
            strTab1 & "<exam_conclusion><![CDATA[" & .strExamConclusion & "]]></exam_conclusion>" & _
            strTab1 & "<apply_no><![CDATA[" & .strApplyNO & "]]></apply_no>" & _
            strTab1 & "<applicant_id><![CDATA[" & .strApplicantID & "]]></applicant_id>" & _
            strTab1 & "<applicant_name><![CDATA[" & .strApplicantName & "]]></applicant_name>" & _
            strTab1 & "<applicant_dept_id><![CDATA[" & .strApplicantDeptID & "]]></applicant_dept_id>" & _
            strTab1 & "<applicant_dept_name><![CDATA[" & .strApplicantDeptName & "]]></applicant_dept_name>" & _
            strTab1 & "<perform_time><![CDATA[" & .strPerformTime & "]]></perform_time>" & _
            strTab1 & "<reporter_id><![CDATA[" & .strReporterID & "]]></reporter_id>" & _
            strTab1 & "<reporter_name><![CDATA[" & .strReporterName & "]]></reporter_name>" & _
            strTab1 & "<report_time><![CDATA[" & .strReportTime & "]]></report_time>" & _
            strTab1 & "<perform_aim><![CDATA[" & .strperformAim & "]]></perform_aim>"
            strXML = strXML & IIf(bytFunc = 0, "</opt_special_exam>", "</ipt_special_exam>")
        End With
    Next
    strXML = strXML & vbCrLf & IIf(bytFunc = 0, "</opt_special_exams>", "</ipt_special_exams>")
    HZYY_SpecialExams = strXML
End Function

Public Function HZYY_GetElectronicMedical(ByRef udtEletMed As ElectronicMedical) As String
'����:��ȡ������Ӳ�����ǩ
    Dim strXML      As String
    Dim strTab1     As String
    strTab1 = vbCrLf & vbTab
   
    With udtEletMed
        strXML = strXML & "<electronic_medical>" & _
        strTab1 & "<electronic_medical_id><![CDATA[" & .strElectronicMedicalID & "]]></electronic_medical_id>" & _
        strTab1 & "<chief_complaint><![CDATA[" & .strChiefComplaint & "]]></chief_complaint>" & _
        strTab1 & "<medical_history><![CDATA[" & .strMedicalHistory & "]]></medical_history>" & _
        strTab1 & "<past_history><![CDATA[" & .strPastHistory & "]]></past_history>" & _
        strTab1 & "<personal_history><![CDATA[" & .strPersonalHistory & "]]></personal_history>" & _
        strTab1 & "<family_disease_history><![CDATA[" & .strFamilyDiseaseHistory & "]]></family_disease_history>" & _
        strTab1 & "<menstrual_history><![CDATA[" & .strMenstrualHistory & "]]></menstrual_history>" & _
        strTab1 & "<obsterical_history><![CDATA[" & .strObstericalHistory & "]]></obsterical_history>" & _
        strTab1 & "<record_doc_id><![CDATA[" & .strRecordDocID & "]]></record_doc_id>" & _
        strTab1 & "<record_doc_name><![CDATA[" & .strRecordDocName & "]]></record_doc_name>" & _
        strTab1 & "<record_time><![CDATA[" & .strRecordTime & "]]></record_time>"
        strXML = strXML & "</electronic_medical>"
    End With
    HZYY_GetElectronicMedical = strXML
End Function

Public Function HZYY_GetAdmissionRecord(ByRef udtAdmission As AdmissionRecord) As String
'����:��ȡ��Ժ��¼
    Dim strXML      As String
    Dim strTab1     As String
    strTab1 = vbCrLf & vbTab
   
    With udtAdmission
        strXML = strXML & "<admission_record>" & _
            strTab1 & "<admission_record_id><![CDATA[" & .strAdmissionRecordID & "]]></admission_record_id>" & _
            strTab1 & "<admission_record_type><![CDATA[" & .strAdmissionRecordType & "]]></admission_record_type>" & _
            strTab1 & "<chief_complaint><![CDATA[" & .strChiefComplaint & "]]></chief_complaint>" & _
            strTab1 & "<medical_history><![CDATA[" & .strMedicalHistory & "]]></medical_history>" & _
            strTab1 & "<past_history><![CDATA[" & .strPastHistory & "]]></past_history>" & _
            strTab1 & "<personal_history><![CDATA[" & .strPersonalHistory & "]]></personal_history>" & _
            strTab1 & "<family_disease_history><![CDATA[" & .strFamilyDiseaseHistory & "]]></family_disease_history>" & _
            strTab1 & "<menstrual_history><![CDATA[" & .strMenstrualHistory & "]]></menstrual_history>" & _
            strTab1 & "<obsterical_history><![CDATA[" & .strObstericalHistory & "]]></obsterical_history>" & _
            strTab1 & "<operation_history><![CDATA[" & .strOperationHistory & "]]></operation_history>" & _
            strTab1 & "<transfusion_history><![CDATA[" & .strTransfusionHistory & "]]></transfusion_history>" & _
            strTab1 & "<infection_history><![CDATA[" & .strInfectionHistory & "]]></infection_history>" & _
            strTab1 & "<vaccination_history><![CDATA[" & .strVaccinationHistory & "]]></vaccination_history>" & _
            strTab1 & "<physical_exam><![CDATA[" & .strPhysicalExam & "]]></physical_exam>" & _
            strTab1 & "<special_exam><![CDATA[" & .strSpecialExam & "]]></special_exam>" & _
            strTab1 & "<auxiliary_exam><![CDATA[" & .strAuxiliaryExam & "]]></auxiliary_exam>" & _
            strTab1 & "<record_doc_id><![CDATA[" & .strRecordDocID & "]]></record_doc_id>" & _
            strTab1 & "<record_doc_name><![CDATA[" & .strRecordDocName & "]]></record_doc_name>" & _
            strTab1 & "<record_time><![CDATA[" & .strRecordTime & "]]></record_time>"
        strXML = strXML & "</admission_record>"
    End With
    HZYY_GetAdmissionRecord = strXML
End Function

Public Function HZYY_GetProgressNotes(ByRef colItem As Collection) As String
'����:��ȡ����¼��ǩ
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtProg     As HZYYProgressNote
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = strXML & "<progress_notes>"
    For i = 1 To colItem.Count
        udtProg = colItem(i)
        With udtProg
            strXML = strXML & "<progress_note>" & _
            strTab1 & "<progress_note_id><![CDATA[" & .strProgressNoteID & "]]></progress_note_id>" & _
            strTab1 & "<progress_note_type><![CDATA[" & .strProgressNoteType & "]]></progress_note_type>" & _
            strTab1 & "<progress_note_content><![CDATA[" & .strProgressNoteContent & "]]></progress_note_content>" & _
            strTab1 & "<record_doc_id><![CDATA[" & .strRecordDocID & "]]></record_doc_id>" & _
            strTab1 & "<record_doc_name><![CDATA[" & .strRecordDocName & "]]></record_doc_name>" & _
            strTab1 & "<record_time><![CDATA[" & .strRecordTime & "]]></record_time>"
            strXML = strXML & "</progress_note>"
        End With
    Next
    strXML = strXML & "</progress_notes>"

    HZYY_GetProgressNotes = strXML
End Function

Public Function HZYY_GetVitalSigns(ByRef colItem As Collection) As String
'����:��ȡ����������ǩ
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtVital     As VitalSign
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = strXML & "<vital_signs>"
    For i = 1 To colItem.Count
        udtVital = colItem(i)
        With udtVital
            strXML = strXML & "<vital_sign>" & _
            strTab1 & "<vital_sign_no><![CDATA[" & .strVitalSignNO & "]]></vital_sign_no>" & _
            strTab1 & "<temperature><![CDATA[" & .strTemperature & "]]></temperature>  " & _
            strTab1 & "<sbp><![CDATA[" & .strSbp & "]]></sbp>" & _
            strTab1 & "<dbp><![CDATA[" & .strDbp & "]]></dbp>" & _
            strTab1 & "<breathing_rate><![CDATA[" & .strBreathingRate & "]]></breathing_rate>  " & _
            strTab1 & "<pulse_rate><![CDATA[" & .strPulseRate & "]]></pulse_rate>" & _
            strTab1 & "<heart_rate><![CDATA[" & .strHeartRate & "]]></heart_rate>" & _
            strTab1 & "<pain_score><![CDATA[" & .strPainScore & "]]></pain_score>" & _
            strTab1 & "<hour24_amount_in><![CDATA[" & .strHour24Amountin & "]]></hour24_amount_in>" & _
            strTab1 & "<hour24_amount_out><![CDATA[" & .strHour24Amountout & "]]></hour24_amount_out>" & _
            strTab1 & "<test_time><![CDATA[" & .strTestTime & "]]></test_time>"
            strXML = strXML & "</vital_sign>"
        End With
    Next
    strXML = strXML & "</vital_signs>"

    HZYY_GetVitalSigns = strXML
End Function

Public Function HZYY_GetPathological(ByRef colItem As Collection) As String
'����:��ȡ������Ϣ��ǩ
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtPath     As PathologicalExam
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = strXML & "<pathological_exams>"
    For i = 1 To colItem.Count
        udtPath = colItem(i)
        With udtPath
            strXML = strXML & "<pathological_exam>" & _
            strTab1 & "<report_id><![CDATA[" & .strReportID & "]]></report_id>" & _
            strTab1 & "<exam_item_code><![CDATA[" & .strExamItemCode & "]]></exam_item_code>" & _
            strTab1 & "<exam_item_name><![CDATA[" & .strExamItemName & "]]></exam_item_name>" & _
            strTab1 & "<sample_name><![CDATA[" & .strSampleName & "]]></sample_name>" & _
            strTab1 & "<pathologic_diagnosis><![CDATA[" & .strPathologicDiagnosis & "]]></pathologic_diagnosis>" & _
            strTab1 & "<apply_no><![CDATA[" & .strApplyNO & "]]></apply_no>  " & _
            strTab1 & "<applicant_id><![CDATA[" & .strApplicantID & "]]></applicant_id>" & _
            strTab1 & "<applicant_name><![CDATA[" & .strApplicantName & "]]></applicant_name>  " & _
            strTab1 & "<applicant_dept_id><![CDATA[" & .strApplicantDeptID & "]]></applicant_dept_id>" & _
            strTab1 & "<applicant_dept_name><![CDATA[" & .strApplicantDeptName & "]]></applicant_dept_name>" & _
            strTab1 & "<perform_time><![CDATA[" & .strPerformTime & "]]></perform_time>  " & _
            strTab1 & "<reporter_id><![CDATA[" & .strReporterID & "]]></reporter_id> " & _
            strTab1 & "<reporter_name><![CDATA[" & .strReporterName & "]]></reporter_name>" & _
            strTab1 & "<report_time><![CDATA[" & .strReportTime & "]]></report_time>" & _
            strTab1 & "<exam_aim><![CDATA[" & .strExamAim & "]]></exam_aim>"
            strXML = strXML & "</pathological_exam>"
        End With
    Next
    strXML = strXML & "</pathological_exams>"

    HZYY_GetPathological = strXML
End Function

Public Function HZYY_GetBacterialReports(ByRef colItem As Collection) As String
'����:��ȡϸ�����������ǩ
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtItem      As BacterialReportItem
    Dim udtBact     As BacterialReport
    Dim j           As Long
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = strXML & "<bacterial_reports>"
    For i = 1 To colItem.Count
        udtBact = colItem(i)
        With udtBact
            strXML = strXML & "<bacterial_report>"
            With udtBact.udtInfo
                strXML = strXML & "<bacterial_report_info>" & _
                    strTab1 & "<report_id><![CDATA[" & .strReportID & "]]></report_id>  " & _
                    strTab1 & "<exam_item_code><![CDATA[" & .strExamItemCode & "]]></exam_item_code>  " & _
                    strTab1 & "<exam_item_name><![CDATA[" & .strExamItemName & "]]></exam_item_name>  " & _
                    strTab1 & "<sample_collect_time><![CDATA[" & .strSampleCollectTime & "]]></sample_collect_time>  " & _
                    strTab1 & "<sample_code><![CDATA[" & .strSampleCode & "]]></sample_code>  " & _
                    strTab1 & "<sample_name><![CDATA[" & .strSampleName & "]]></sample_name>  " & _
                    strTab1 & "<sample_collect_opporunity><![CDATA[" & .strSampleCollectOpporunity & "]]></sample_collect_opporunity>  " & _
                    strTab1 & "<apply_no><![CDATA[" & .strApplyNO & "]]></apply_no>  " & _
                    strTab1 & "<applicant_id><![CDATA[" & .strApplicantID & "]]></applicant_id>  " & _
                    strTab1 & "<applicant_name><![CDATA[" & .strApplicantName & "]]></applicant_name>  " & _
                    strTab1 & "<applicant_dept_id><![CDATA[" & .strApplicantDeptID & "]]></applicant_dept_id>  " & _
                    strTab1 & "<applicant_dept_name><![CDATA[" & .strApplicantDeptName & "]]></applicant_dept_name>  " & _
                    strTab1 & "<reporter_id><![CDATA[" & .strReporterID & "]]></reporter_id>  " & _
                    strTab1 & "<reporter_name><![CDATA[" & .strReporterName & "]]></reporter_name>  " & _
                    strTab1 & "<report_time><![CDATA[" & .strReportTime & "]]></report_time>"
                strXML = strXML & "</bacterial_report_info>"
            End With
            For j = 1 To .colItem.Count
                udtItem = .colItem(i)
                With udtItem
                    strXML = strXML & "<bacterial_report_item>" & _
                        strTab1 & "<report_id><![CDATA[" & .strReportID & "]]></report_id>  " & _
                        strTab1 & "<report_item_id><![CDATA[" & .strReportItemID & "]]></report_item_id>  " & _
                        strTab1 & "<exam_item_result><![CDATA[" & .strExamItemResult & "]]></exam_item_result>"
                    strXML = strXML & "</bacterial_report_item>"
                End With
            Next
            strXML = strXML & "</bacterial_report>"
        End With
    Next
    strXML = strXML & "</bacterial_reports>"

    HZYY_GetBacterialReports = strXML
End Function

Public Function HZYY_GetDrugSensitives(ByRef colItem As Collection) As String
'����:��ȡҩ������������Ϣ��ǩ
    Dim strXML      As String
    Dim strTab1     As String
    Dim udtItem      As DrugSensitiveItem
    Dim udtDrug     As DrugSensitive
    Dim j           As Long
    Dim i           As Long
    strTab1 = vbCrLf & vbTab
   
    strXML = strXML & "<drug_sensitives>"
    For i = 1 To colItem.Count
        udtDrug = colItem(i)
        strXML = strXML & "<drug_sensitive>"
        With udtDrug.udtInfo
            strXML = strXML & "<drug_sensitive_info>" & _
            strTab1 & "<report_id><![CDATA[" & .strReportID & "]]></report_id>  " & _
            strTab1 & "<exam_item_code><![CDATA[" & .strExamItemCode & "]]></exam_item_code>  " & _
            strTab1 & "<exam_item_name><![CDATA[" & .strExamItemName & "]]></exam_item_name>  " & _
            strTab1 & "<sample_collect_time><![CDATA[" & .strSampleCollectTime & "]]></sample_collect_time>  " & _
            strTab1 & "<sample_code><![CDATA[" & .strSampleCode & "]]></sample_code>  " & _
            strTab1 & "<sample_name><![CDATA[" & .strSampleName & "]]></sample_name>  " & _
            strTab1 & "<sample_collect_opporunity><![CDATA[" & .strSampleCollectOpporunity & "]]></sample_collect_opporunity>  " & _
            strTab1 & "<germ_id><![CDATA[" & .strgermID & "]]></germ_id>  " & _
            strTab1 & "<germ_name><![CDATA[" & .strgermName & "]]></germ_name>  " & _
            strTab1 & "<apply_no><![CDATA[" & .strApplyNO & "]]></apply_no>  " & _
            strTab1 & "<applicant_id><![CDATA[" & .strApplicantID & "]]></applicant_id>  " & _
            strTab1 & "<applicant_name><![CDATA[" & .strApplicantName & "]]></applicant_name>  " & _
            strTab1 & "<applicant_dept_id><![CDATA[" & .strApplicantDeptID & "]]></applicant_dept_id>  " & _
            strTab1 & "<applicant_dept_name><![CDATA[" & .strApplicantDeptName & "]]></applicant_dept_name>  " & _
            strTab1 & "<reporter_id><![CDATA[" & .strReporterID & "]]></reporter_id>  " & _
            strTab1 & "<reporter_name><![CDATA[" & .strReporterName & "]]></reporter_name>  " & _
            strTab1 & "<report_time><![CDATA[" & .strReportTime & "]]></report_time>"
            strXML = strXML & "</drug_sensitive_info>"
        End With
        For j = 1 To udtDrug.colItem.Count
            udtItem = udtDrug.colItem(i)
            With udtItem
                strXML = strXML & "<drug_sensitive_item>" & _
                strTab1 & "<report_id><![CDATA[" & .strReportID & "]]></report_id>  " & _
                strTab1 & "<report_item_id><![CDATA[" & .strReportItemID & "]]></report_item_id>  " & _
                strTab1 & "<antibiotic_id><![CDATA[" & .strantibioticID & "]]></antibiotic_id>  " & _
                strTab1 & "<antibiotic_name><![CDATA[" & .strantibioticName & "]]></antibiotic_name>  " & _
                strTab1 & "<sensitivity><![CDATA[" & .strsensitivity & "]]></sensitivity>  " & _
                strTab1 & "<mic><![CDATA[" & .strmic & "]]></mic>"
                strXML = strXML & "</drug_sensitive_item>"
            End With
        Next
        strXML = strXML & "</drug_sensitive>"

    Next
    strXML = strXML & "</drug_sensitives>"

    HZYY_GetDrugSensitives = strXML
End Function

Public Function HZYY_GetRootXML(colItem As Collection, Optional ByVal bytFunc As Byte, Optional ByVal bytType As Byte) As String
'����:bytFunc=0-����
'     bytType=0-��Ԥ; 1=ɾ��ҽ��
    Dim strRet          As String
    Dim i               As Long
    Dim udtBase         As HZYYBASE
    Dim udtPati         As OPTPATIENT
    Dim udtInPati       As IPTPATIENT
    Dim udtOptPres      As OptPrescription
    Dim udtPresInfo     As OPTPRESCRIPTIONSINFO
    Dim udtPresItem     As OptPRESCRIPTIONSITEM
    Dim udtDiag         As Diagnosis
    Dim udtAllergy      As Allergy
    Dim udtOpera        As Operation
    Dim udtOrder        As Order
    Dim colTmp          As Collection
    
    strRet = strRet & "<root>"
    For i = 1 To colItem.Count
        Select Case UCase(TypeName(colItem(i)))
        
        Case UCase(TypeName(udtBase))
             udtBase = colItem(i)
            strRet = strRet & HZYY_MakeBASEXML(udtBase)
        Case UCase(TypeName(udtPati))
            udtPati = colItem(i)
            strRet = strRet & HZYY_MakeOPSPatient(udtPati)
        Case UCase(TypeName(udtInPati))
            udtInPati = colItem(i)
            strRet = strRet & HZYY_MakeIPSPatient(udtInPati)
        Case UCase(TypeName(udtOrder))
            udtOrder = colItem(i)
            strRet = strRet & HZYY_GetOrder(udtOrder, bytType)
        Case UCase("Collection")
            If colItem(i).Count > 0 Then
                If UCase(TypeName(colItem(i)(1))) = UCase(TypeName(udtOptPres)) Then
                    strRet = strRet & HZYY_GetOPTPres(colItem(i), bytType)
                ElseIf UCase(TypeName(colItem(i)(1))) = UCase(TypeName(udtDiag)) Then
                    strRet = strRet & HZYY_GetDiag(colItem(i), bytFunc)
                ElseIf UCase(TypeName(colItem(i)(1))) = UCase(TypeName(udtAllergy)) Then
                    strRet = strRet & HZYY_GetAllergies(colItem(i), bytFunc)
                ElseIf UCase(TypeName(colItem(i)(1))) = UCase(TypeName(udtOpera)) Then
                    strRet = strRet & HZYY_GetOperations(colItem(i), bytFunc)
                End If
            End If
        Case Else
        End Select
    Next
    strRet = strRet & "</root>"
    
    HZYY_GetRootXML = strRet
    WriteLog "" & glngModel, "HZYY_GetRootXML", strRet
End Function

Public Function HZYY_GetSex(ByVal strPara As String) As String
'0   δ֪���Ա�
'1   ����
'2   Ů��
'9   δ˵�����Ա�
    Dim strRet As String
    If strPara = "" Then
        strRet = "9"
    ElseIf InStr(strPara, "��") > 0 Then
        strRet = "1"
    ElseIf InStr(strPara, "Ů") > 0 Then
        strRet = "2"
    Else
        strRet = "0"
    End If
    HZYY_GetSex = strRet
End Function

Public Function HZYY_GetMarital(ByVal strPara As String) As String
'10  δ��
'20  �ѻ�
'21  ����
'22  �ٻ�
'23  ����
'30  ɥż
'40  ���
'90  δ˵���Ļ���״��
    Dim strRet As String
    If strPara = "" Then
        strRet = "90"
    ElseIf InStr(strPara, "δ��") > 0 Then
        strRet = "10"
    ElseIf InStr(strPara, "�ѻ�") > 0 Then
        strRet = "20"
    ElseIf InStr(strPara, "����") > 0 Then
        strRet = "21"
    ElseIf InStr(strPara, "�ٻ�") > 0 Then
        strRet = "22"
    ElseIf InStr(strPara, "����") > 0 Then
        strRet = "23"
    ElseIf InStr(strPara, "ɥż") > 0 Then
        strRet = "30"
    ElseIf InStr(strPara, "���") > 0 Then
        strRet = "40"
    Else
        strRet = ""
    End If
    HZYY_GetMarital = strRet
End Function

Public Sub HZYY_DrugInstructions(Optional ByVal lngDrugID As Long)
'���ܣ���������ҩƷ˵��
'����:lngDrugID-ҩƷID
    Dim strUrl As String
    Dim lngRet As Long
    
    If lngDrugID = 0 And (Not gobjAdvice Is Nothing) And (glngModel = PM_����༭ Or glngModel = PM_����ҽ���嵥 Or _
        glngModel = PM_סԺ�༭ Or glngModel = PM_סԺҽ���嵥) Then
        With gobjAdvice
            If InStr(",5,6,7,", .TextMatrix(.Row, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(.Row, gobjCOL.intCOL�շ�ϸĿID)) <> 0 Then
                '��ȡ��ѡҽ����ҩƷ����
                lngDrugID = Val(.TextMatrix(.Row, gobjCOL.intCOL�շ�ϸĿID))
            End If
        End With
    End If
    'strUrl = "http://118.31.246.211:8080/zlcx/data_detail.action?webHisId=11221&hospitalCode=cqzl123"
    If gbytType = 0 Then
        strUrl = "http://" & gstrIP & ":" & gstrPort & "/zlcx/data_detail.action?webHisId=" & lngDrugID
    Else
        '��Ʒ�ǹ���
        strUrl = "http://" & gstrIP & ":" & gstrPort & "/zlcx/data_detail.action?webHisId=" & lngDrugID & "&hospitalCode=" & gstrHOSCODE
    End If
    lngRet = ShellExecute(0, "open", strUrl, "", "", SW_SHOWNORMAL)
End Sub

Public Function HZYY_GetPara() As Boolean
        Dim arrList As Variant
        Dim strPara As String
        
        On Error GoTo errH
100     strPara = zlDatabase.GetPara(90001, glngSys, , "") '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
        '��ʽ������IP&&�������˿ں�
102     If strPara = "" Then strPara = "118.31.246.211" & G_STR_SPLIT & "8080" & G_STR_SPLIT & "cqzl123" & G_STR_SPLIT & "6000" & G_STR_SPLIT & "0"
104     arrList = Split(strPara, G_STR_SPLIT)
106     If UBound(arrList) >= 4 Then
            gstrIP = arrList(0)
            gstrPort = arrList(1)
            gstrHOSCODE = arrList(2)     'ҽԺ����
            gstrPortPlus = arrList(3)
            gbytType = Val(arrList(4) & "")
        Else
            gstrIP = "118.31.246.211"
            gstrPort = "8080"
            gstrHOSCODE = "cqzl123"
            gstrPortPlus = "6000"
            gbytType = 0
        End If
        Exit Function
errH:
146     MsgBox "��ȡ����ʧ�ܣ�" & vbNewLine & "HZYY_GetPara:��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HZYY_SetPara() As String
    HZYY_SetPara = gstrIP & G_STR_SPLIT & gstrPort & G_STR_SPLIT & gstrHOSCODE & G_STR_SPLIT & gstrPortPlus & G_STR_SPLIT & gbytType
End Function

Public Function AdviceCheckWarn_HZYY(ByVal bytFunc As Byte, ByVal lngPatiID As Long, ByVal str�Һŵ� As String, _
    ByVal lng��ҳID As String, Optional ByRef rsOut As ADODB.Recordset, Optional ByVal strҽ��IDs As String) As Boolean
'���ܣ����ú���������ҩ���ϵͳ(BS��)��ҽ�����к�����ҩ������ع���
'
'������
'bytFunc=0-ҽ������;1-ҩ�����;2-ɾ������;3-ɾ��ҽ��;4-�ϴ���Ч����\��Чҽ��
'����ֵ:
'   True -������
    Dim udtBase     As HZYYBASE
    Dim udtOptPati  As OPTPATIENT
    Dim udtIptPati  As IPTPATIENT
    Dim udtOptPres  As OptPrescription
    Dim udtPresInfo As OPTPRESCRIPTIONSINFO
    Dim udtPresItem As OptPRESCRIPTIONSITEM
    Dim udtOrderItem    As MedicalOrderItem
    Dim udtHerbInfo     As HerbMedicalOrderInfo
    Dim udtHerbItem     As HerbMedicalOrderItem
    Dim udtHerbOrder    As HerbMedicalOrder
    Dim udtOrd
    Dim udtOrder        As Order
    Dim udtDiag     As Diagnosis
    Dim udtAllergy      As Allergy
    Dim udtOpera  As Operation
    Dim rsPati As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim rs���  As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    
    Dim str��� As String
    Dim str���� As String
    Dim strDrugID As String
    
    Dim colTmp As Collection
    Dim colRecipe As Collection
    Dim colXML As Collection
    
    Dim strSQL As String
    Dim lng�Һ�ID As Long
    Dim lngDeptID As Long
    Dim i As Long, k As Long, lngCount As Long
    Dim strDeptName As String
    Dim str��ϱ��� As String, str������� As String
    Dim strTmp As String
    Dim strUrl As String
    Dim strXML As String
    Dim strRet As String
    Dim blnDo As Boolean
    Dim curDate As Date
    Dim strRecipeNo As String
    Dim strҩƷIDs As String
    Dim byt���� As Byte
    Dim bytRet  As Byte
    Dim blnIsHaveOut As Boolean
    Dim arrTemp As Variant
    
    On Error GoTo errH
    
    Set colXML = New Collection
    Set rsPati = GetPatiInfo_YF(lngPatiID, str�Һŵ�, lng��ҳID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    If str�Һŵ� <> "" Then
        lng�Һ�ID = rsPati!����Id
        strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                        "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                        "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
                        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, lng�Һ�ID)
        rsTmp.Filter = "��Ŀ����='���'"
        If rsTmp.RecordCount <> 0 Then str��� = NVL(rsTmp!��¼����)
        rsTmp.Filter = "��Ŀ����='����'"
        If rsTmp.RecordCount <> 0 Then str���� = NVL(rsTmp!��¼����)
        lngDeptID = CLng(rsPati!�Һſ���ID & "")
        strDeptName = rsPati!�Һſ��� & ""
    Else
        str��� = rsPati!��� & ""
        str���� = rsPati!���� & ""
        lngDeptID = CLng(rsPati!��Ժ����ID)
        strDeptName = rsPati!��Ժ���� & ""
    End If
    
    'base��ǩ
    With udtBase
        .strHospCode = gstrHOSCODE
        .strPatiID = lngPatiID
        If str�Һŵ� <> "" Then
            .strSource = IIf(NVL(rsPati!����, 0) = 1, "����", "����")
            .strEventNO = lng�Һ�ID
        Else
            .strSource = "סԺ" '����|סԺ|����
            .strEventNO = lngPatiID & "_" & lng��ҳID
        End If
    End With
    colXML.Add udtBase
    
    Select Case bytFunc
    Case 0, 1, 4
        curDate = zlDatabase.Currentdate
        If str�Һŵ� <> "" Then
            'opt_patient���ﻼ�߾����ǩ
            With udtOptPati
                .strSex = HZYY_GetSex(rsPati!�Ա� & "")
                .strName = rsPati!���� & ""
                .strIDType = "01"                               '�������֤
                .strIDNO = rsPati!���֤�� & ""
                .strBirthWeight = ""                            '��������
                .strBirthDay = Format(rsPati!�������� & "", "YYYY-MM-DD HH:mm:ss")            '��������
                .strEthnicGroup = rsPati!���� & ""              '����
                .strNativePlace = rsPati!���� & ""              '����
                .strRace = ""                                       '����
                .strMedCardNO = ""                '���￨��
                .strEventTime = Format(rsPati!����ʱ�� & "", "YYYY-MM-DD HH:mm:ss")                 '����ʱ��
                .strDeptID = rsPati!�Һſ���ID & ""               '�Һſ���ID
                .strDeptName = rsPati!�Һſ��� & ""                 '�Һſ�������
                .strPayType = rsPati!ҽ�Ƹ��ʽ & ""                    '��������
                .strPregnancy = ""                  '�Ƿ���
                .strTimeOfPreg = ""                 '����
                .strBreastFeeding = ""              '�Ƿ���
                .strHeight = str���                       '���
                .strWeight = str����                        '����
                .strAddress = rsPati!��ͥ��ַ & ""                          '���˵�ַ
                .strPhoneNo = rsPati!�ֻ��� & ""                         '���˵绰
                .strDialysis = ""                                   '�Ƿ�͸��
                .strmarital = HZYY_GetMarital(rsPati!����״�� & "")                            '����״��
                .strOccupation = rsPati!ְҵ & ""                 'ְҵ
                .strSpecialConstitution = "" '��������
                .strVisitType = "����"                   '�������  ����|����
                .strPatiCondition = ""
            End With
            colXML.Add udtOptPati
        Else
            'סԺ������Ϣ
            With udtIptPati
                .strSex = HZYY_GetSex(rsPati!�Ա� & "")
                .strName = rsPati!���� & ""
                .strIDType = "01"
                .strIDNO = rsPati!���֤�� & ""
                .strBirthWeight = ""                    '��������
                .strBirthDay = Format(rsPati!�������� & "", "YYYY-MM-DD HH:mm:ss")                        '��������
                .strEthnicGroup = rsPati!���� & ""                    '����
                .strNativePlace = rsPati!���� & ""                  '����
                .strRace = ""                           '����
                .strMedCardNO = ""                      '���￨��
                .strPayType = rsPati!ҽ�Ƹ��ʽ & ""                         '��������
                .strPregnancy = ""                      '�Ƿ���
                .strTimeOfPreg = ""                     '����
                .strBreastFeeding = ""                  '�Ƿ���
                .strHeight = str���                         '���
                .strWeight = str����                         '����
                .strAddress = ""                        '���˵�ַ
                .strPhoneNo = ""                        '���˵绰
                .strDialysis = ""                       '�Ƿ�͸��
                .strmarital = ""                        '����״��
                .strOccupation = ""                     'ְҵ
                .strSpecialConstitution = ""            '��������
                .strINDeptId = rsPati!��Ժ����ID        '��Ժ����ID
                .strINDeptName = rsPati!��Ժ���� & ""   '��Ժ��������
                .strHospitalTime = Format(rsPati!��Ժ���� & "", "YYYY-MM-DD HH:mm:ss")                    '��Ժʱ��
                .strInWardID = rsPati!��Ժ����ID & ""                        '��Ժ����ID
                .strInWardName = rsPati!��Ժ���� & ""                     '��Ժ��������
                .strInWardBedNo = rsPati!��Ժ���� & ""                    '��Ժ������
                .strInConditon = ""                      '��Ժ����
                .strWeightOfBaby = ""                    '��������Ժ����
                .strPatientConditon = ""               '����״̬�磺��ͨ���ˡ�Σ�ز���
            End With
            colXML.Add udtIptPati
        End If
        'opt_prescriptions�����ʹ�����ϸ��Ϣ��ǩ
        'ҩƷ��Ϣ
        Select Case glngModel
        Case PM_����༭, PM_����ҽ���嵥, PM_סԺ�༭, PM_סԺҽ���嵥
            Set rsAdvice = CreateAdviceRS_HZYY(rsOut, rs���)
            If glngModel = PM_����༭ Or glngModel = PM_����ҽ���嵥 Then
                Set colRecipe = New Collection
                strRecipeNo = ""
                rsAdvice.Filter = "����ID>0"
                rsAdvice.Sort = "����ID"
                For i = 1 To rsAdvice.RecordCount
                    If Val(strRecipeNo) <> Val(rsAdvice!����ID & "") Then
                        If i <> 1 Then
                            Set udtOptPres.colPresItem = colTmp
                            colRecipe.Add udtOptPres
                        End If
                        strRecipeNo = rsAdvice!����ID & ""
                        With udtPresInfo
                            .strRecipeId = strRecipeNo
                            .strRecipeNo = strRecipeNo
                            .strRecipeSource = "����"                  '������Դ����|����|����
                            .strRecipeCategory = ""                     '��ͨ���������ƴ���,������,���ﴦ����
                            .strRecipeType = IIf(rsAdvice!������� & "" = "5", "��ҩ��", IIf(rsAdvice!������� & "" = "6", "�г�ҩ��", "��ҩ��")) '�������� ��ҩ��|�г�ҩ��|��ҩ��
                            .strDeptID = rsAdvice!��������id & ""
                            .strDeptName = rsAdvice!�������� & ""
                            .strRecipeDocId = rsAdvice!����ҽ��ID & ""                        '����ҽ������
                            .strRecipeDocName = rsAdvice!����ҽ�� & ""                     '����ҽ������
                            .strRecipeTime = Format(rsAdvice!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS")
                            .strHerbPacketCount = ""                    '��Ƭ���� '�������ҩ��Ƭ�������ⲿ����Ϣ��Ҫ��д����ҩ����Ϊ��
                            .strHerbUnitPrice = ""                      '��Ƭ�����۸�
                            .strRecipeFeeTotal = "0"                     '�������
                            '������his�е�״̬��0 ����������1 ��ҩ������2 ���ϴ�����ֻ�����롰2|1|0��
                            .strRecipeStatus = "0"
                            .strUrgentFlag = "0"            '������־ ���ڴ�����������ˣ�1 �ǣ�0 ��,���ṩ��Ϊ0��ֻ�����롰1|0��
                        End With
                        udtOptPres.udtOptPresInfo = udtPresInfo
                        Set colTmp = New Collection
                    End If
                    With udtPresItem
                        .strRecipeId = strRecipeNo                   '����ID
                        .strRecipeItemId = rsAdvice!ҽ��id & ""         '������ϸ���
                        .strGroupNO = rsAdvice!���ID & ""
                        .strDrugID = rsAdvice!ҩƷID & ""
                        .strDrugName = rsAdvice!ҩƷ���� & ""
                        .strDrugUsingAim = rsAdvice!��ҩĿ�� & ""          '��ҩĿ��
                        .strManufacturerID = ""                 '��������ID
                        .strManufacturerName = ""               '������������
                        .strDrugdose = rsAdvice!�������� & rsAdvice!������λ                       '����ʹ�õ�ҩƷ��������λ���磺0.5g ��200ml��
                        .strDrugadminRouteName = rsAdvice!�÷� & ""             'ҩƷʹ��;�����磺������ע��
                        .strDrugUsingFreq = rsAdvice!Ƶ�� & ""                   'ҩƷ��ҩƵ�Σ��磺qd��bid��ÿ��2�ε�
                        .strDuration = rsAdvice!���� & ""                   '�Ƴ� ����
                        rs���.Filter = "ҩƷID=" & rsAdvice!ҩƷID
                        If Not rs���.EOF Then
                            .strPreparation = rs���!ҩƷ���� & ""                    'ҩƷ��������
                            .strSpecification = rs���!��� & ""                   'ҩƷ�������
                            .strUnitPrice = rs���!�ּ� & ""                         '����
                            .strCountUnit = rs���!�����װ & ""                      '��װ������� ���һ��ҩƷ12Ƭ����ҩ��λΪ��ʱ����װ�������Ϊ12����ҩ��λΪƬʱ����װ�������Ϊ1
                            .strPackUnit = rs���!���ﵥλ & ""              '��װ���λ
                            .strFeeTotal = Val(rsAdvice!���� & "") * Val(rs���!�ּ� & "")       'ҩƷ�ĵ���*��ҩ������������λΪԪ
                        Else
                            .strPreparation = ""                   'ҩƷ��������
                            .strSpecification = ""                  'ҩƷ�������
                            .strUnitPrice = ""                      '����
                        End If
                        .strDespensingNum = FormatEx(rsAdvice!���� & "", 5)                     '��ҩ����    ��ҩ��������2�С�10Ƭ��
                        .strSkinTestFlag = "0"                   'Ƥ�Ա�־    1 Ƥ�ԣ�0 ��Ƥ�ԣ�ֻ�����롰1|0��
                        .strdrugReturnFlag = "0"                '�Ƿ���ҩ��־    1 �ǣ�0 ��,���ṩ��Ϊ0��ֻ�����롰1|0��
                        .strOuvasFlag = Val(rsAdvice!��Һ & "")                    '���ﾲ���־    1 �ǣ�0 ��,���ṩ��Ϊ0��ֻ�����롰1|0��
                        .strDrippingSpeed = rsAdvice!���� & ""             '���� ������ҺҩƷ��עʱ����ٶȵ��������磺1Сʱ��20��/���ӣ���ͬʱ�������ֺ͵�λ����λΪСʱ���/����
                        If Val(rsAdvice!��־ & "") = 1 And Val(udtOptPres.udtOptPresInfo.strUrgentFlag) <> 1 Then
                            udtOptPres.udtOptPresInfo.strUrgentFlag = "1"
                        End If
                    End With
                    colTmp.Add udtPresItem, "_" & colTmp.Count + 1
                
                    If i = rsAdvice.RecordCount Then
                        Set udtOptPres.colPresItem = colTmp
                        colRecipe.Add udtOptPres, "_" & colRecipe.Count + 1
                    End If
                    rsAdvice.MoveNext
                Next
                If colRecipe.Count = 0 Then AdviceCheckWarn_HZYY = True: Exit Function     'ҽ���´����û���´�ҩƷʱ������
                colXML.Add colRecipe, "_" & colXML.Count + 1
            ElseIf PM_סԺ�༭ = glngModel Or glngModel = PM_סԺҽ���嵥 Then
                'ordersҽ����Ϣ��ǩ
                Set udtOrder.colMedical = New Collection
                rsAdvice.Filter = "�������='5' OR �������='6'"
                For i = 1 To rsAdvice.RecordCount
                    If PM_סԺ�༭ And rsAdvice!��Ժ��ҩ = 1 Then blnIsHaveOut = True
                    With udtOrderItem
                        .strOrderId = rsAdvice!ҽ��id & ""                      'ҽ��id
                        .strOrderTime = Format(rsAdvice!����ʱ�� & "", "YYYY-MM-DD HH:mm:ss")      'ҽ��ʱ��
                        .strOrderDeptID = rsAdvice!��������id & ""                    'ҽ�����Ҵ���
                        .strOrderDeptName = rsAdvice!�������� & ""                 'ҽ����������
                        .strDocGroup = ""                        'ҽ��ҽ��������
                        .strOrderDocID = rsAdvice!����ҽ��ID & ""                    'ҽ��ҽ��id
                        .strOrderDocName = rsAdvice!����ҽ�� & ""                  'ҽ��ҽ������
                        .strOrderDocTitle = rsAdvice!רҵ����ְ�� & ""                'ҽ��ҽ��ְ��
                        .strOrderType = rsAdvice!ҽ����Ч & ""
                        .strDrugPurpose = ""                      'ҩ��Ŀ��
                        .strGroupNO = rsAdvice!���ID & ""                        '���
                        .strDrugID = rsAdvice!ҩƷID & ""                          'ҩƷID
                        .strDrugName = rsAdvice!ҩƷ���� & ""                       'ҩƷͨ����
                        rs���.Filter = "ҩƷID=" & rsAdvice!ҩƷID
                        If Not rs���.EOF Then
                            .strPreparation = rs���!ҩƷ���� & ""                    'ҩƷ��������
                            .strSpecifications = rs���!��� & ""                    'ҩƷ�������
                            .strUnitPrice = rs���!�ּ� & ""                         '����
                            .strManufacturerID = rs���!�������� & ""                    '��������id
                            .strManufacturerName = rs���!���ұ��� & ""                 '������������
                            .strCountUnit = rs���!סԺ��װ & ""                    '��װ�������
                            .strPackUnit = rs���!סԺ��λ & ""                        '��װ���λ
                            .strFeeTotal = Val(rsAdvice!���� & "") * Val(rs���!�ּ� & "")                 '�ܼ�
                        Else
                            .strPreparation = ""                   'ҩƷ��������
                            .strSpecifications = ""                  'ҩƷ�������
                            .strUnitPrice = ""                      '����
                            .strManufacturerID = ""                      '��������id
                            .strManufacturerName = ""                   '������������
                        End If
                        .strDrugdose = rsAdvice!�������� & rsAdvice!������λ                         'ÿ�θ�ҩ����
                        .strDrugadminRouteName = rsAdvice!�÷� & ""           '��ҩ;��
                        .strDrugUsingFreq = rsAdvice!Ƶ�� & ""                  '��ҩƵ��
                        .strDrugUsingTimePoint = ""            '��ҩʱ��
                        .strDrugUsingAim = rsAdvice!��ҩĿ�� & ""                   '��ҩĿ��
                        .strDrugUsingArea = ""                  '��ҩ��λ
                        .strDrugSource = ""                      'ҩƷ��Դ
                        .strDuration = ""                         '�Ƴ�
                        .strDespensingNum = FormatEx(rsAdvice!���� & "", 5)                     '��ҩ����
                        .strCheckTime = ""                       '����ʱ��
                        .strCheckNurseID = ""                   '���˻�ʿid
                        .strCheckNurseName = ""                '���˻�ʿ����
                        .strOrderValidTime = Format(rsAdvice!��ʼʱ�� & "", "YYYY-MM-DD HH:mm:ss")                 'ҽ����Чʱ��
                        .strOrderInvalidTime = Format(rsAdvice!����ʱ�� & "", "YYYY-MM-DD HH:mm:ss")               'ҽ��ʧЧʱ��
                        .strSpecialPrompt = ""                   '����Ҫ��
                        .strSkinTestTime = ""                   'Ƥ��ʱ��
                        .strSkinTestFlag = "0"                   'Ƥ�Ա�־
                        .strSkinTestResult = ""                 'Ƥ�Խ��
                        .strdrugReturnFlag = ""                 '�Ƿ���ҩ��־
                        .strStopFlag = ""                        '�Ƿ�ͣҩ��־
                        .strPivasFlag = Val(rsAdvice!��Һ & "")                      'סԺ�����־
                        .strUrgentFlag = IIf(rsAdvice!��־ & "" = "1", "1", "0")                       '������־
                        .strDrippingSpeed = rsAdvice!���� & ""                  '����
                        .strLimitTime = ""                       '����ʱ��
                        .strTherapeuticRegimen = ""              '��ҩ����
                        .strExeDeptID = ""                      'ҽ��ִ�п���id
                        .strExeDeptName = ""                   'ҽ��ִ�п�������
                        .strDispensingWindow = ""                '��ҩ���ں�
                        .strDrugstoreArea = ""                  '��Ʒ���ܺ�
                    End With
                    
                    udtOrder.colMedical.Add udtOrderItem
                    rsAdvice.MoveNext
                Next
                
                rsAdvice.Filter = "�������='7'"
                strRecipeNo = ""
                Set udtOrder.colHerbMedical = New Collection
                For i = 1 To rsAdvice.RecordCount
                    If rsAdvice!��Ժ��ҩ = 1 And glngModel = PM_סԺ�༭ Then blnIsHaveOut = True
                    If rsAdvice!���ID <> strRecipeNo Then
                        If i <> 1 Then
                            udtOrder.colHerbMedical.Add udtHerbOrder
                        End If
                        strRecipeNo = rsAdvice!���ID & ""
                        With udtHerbInfo
                            .strOrderId = rsAdvice!���ID & ""                           'ҽ��id
                            .strOrderTime = Format(rsAdvice!����ʱ�� & "", "YYYY-MM-DD HH:mm:ss")                          'ҽ��ʱ��
                            .strOrderDeptID = rsAdvice!��������id & ""                         'ҽ�����Ҵ���
                            .strOrderDeptName = rsAdvice!�������� & ""                      'ҽ����������
                            .strDocGroup = ""                           'ҽ��ҽ��������
                            .strOrderDocID = rsAdvice!����ҽ��ID & ""                          'ҽ��ҽ��id
                            .strOrderDocName = rsAdvice!����ҽ�� & ""                        'ҽ��ҽ������
                            .strOrderDocTitle = rsAdvice!רҵ����ְ�� & ""                       'ҽ��ҽ��ְ��
                            .strOrderType = rsAdvice!ҽ����Ч & ""                          'ҽ������
                            .strHerbUnitPrice = ""                      '��Ƭ�����۸�
                            .strHerbPacketCount = ""                    '��Ƭ����
                            .strIsCream = ""                            '�෽
                            .strCheckTime = ""                          '����ʱ��
                            .strCheckNurseID = ""                       '���˻�ʿid
                            .strCheckNurseName = ""                     '���˻�ʿ����
                            .strOrderValidTime = Format(rsAdvice!��ʼʱ�� & "", "YYYY-MM-DD HH:mm:ss")                     'ҽ����Чʱ��
                            .strOrderInvalidTime = Format(rsAdvice!����ʱ�� & "", "YYYY-MM-DD HH:mm:ss")                       'ҽ��ʧЧʱ��
                            .strdrugReturnFlag = ""                     '�Ƿ���ҩ��־
                            .strStopFlag = ""                           '�Ƿ�ͣҩ��־
                            .strUrgentFlag = IIf(rsAdvice!��־ & "" = "1", "1", "0")                        '������־
                            .strExeDeptID = ""                          'ҽ��ִ�п���id
                            .strExeDeptName = ""                        'ҽ��ִ�п�������
                        End With
                        udtHerbOrder.udtHerbInfo = udtHerbInfo
                        Set udtHerbOrder.colItemHerb = New Collection
                    End If
                    With udtHerbItem
                        .strOrderId = rsAdvice!ҽ��id & ""                   'ҽ��id
                        .strOrderitemID = rsAdvice!ҽ��id & ""               'ҽ����ϸ
                        .strGroupNO = rsAdvice!���ID & ""                   '���
                        .strDrugID = rsAdvice!ҩƷID & ""                    'ҩƷID
                        .strDrugName = rsAdvice!ҩƷ���� & ""                       'ҩƷͨ����
                        .strDrugdose = rsAdvice!�������� & rsAdvice!������λ & ""   'ÿ�θ�ҩ����
                        .strDrugadminRouteName = rsAdvice!�÷� & ""                 '��ҩ;��
                        .strDrugUsingFreq = rsAdvice!Ƶ�� & ""                      '��ҩƵ��
                        rs���.Filter = "ҩƷID=" & rsAdvice!ҩƷID
                        If Not rs���.EOF Then
                            .strPreparation = rs���!ҩƷ���� & ""                      'ҩƷ��������
                            .strSpecifications = rs���!��� & ""                       'ҩƷ�������
                            .strUnitPrice = rs���!�ּ� & ""                            '����
                            .strManufacturerID = rs���!�������� & ""                   '��������id
                            .strManufacturerName = rs���!���ұ��� & ""
                        Else
                            .strPreparation = ""                    'ҩƷ��������
                            .strSpecifications = ""                 'ҩƷ�������
                            .strUnitPrice = ""                      '����
                            .strManufacturerID = ""                 '��������id
                            .strManufacturerName = ""               '������������
                        End If
                        .strDespensingNum = ""                                                           '��ҩ����
                        .strFeeTotal = ""                                                                '�ܼ�
                        .strSpecialPrompt = ""                                                           '����Ҫ��
                    End With
                    udtHerbOrder.colItemHerb.Add udtHerbItem
                    If i = rsAdvice.RecordCount Then
                        udtOrder.colHerbMedical.Add udtHerbOrder
                    End If
                    rsAdvice.MoveNext
                Next
                
                Set udtOrder.colNonMedical = New Collection
                
                If udtOrder.colMedical.Count = 0 And udtOrder.colHerbMedical.Count = 0 Then AdviceCheckWarn_HZYY = True: Exit Function     'ҽ���´����û���´�ҩƷʱ������
                colXML.Add udtOrder, "_" & colXML.Count + 1
            End If
        Case PM_���ŷ�ҩ, PM_������ҩ, PM_PIVA����
            Set rsTmp = CreateAdviceRS_HZYY(rsOut, rs���, strҽ��IDs)
            byt���� = 1
        End Select
        If str�Һŵ� <> "" And PM_����༭ = glngModel Then
            'diagnoses�����Ϣ��ǩ
            Set colTmp = New Collection
            If Not gobjDiags Is Nothing Then
                For i = 1 To gobjDiags.Count
                    With udtDiag
                        If gobjDiags.Item(i).str������� <> "" Then
                            If gobjDiags.Item(i).str�������� <> "" Then
                                .strDiagID = gobjDiags.Item(i).str����ID
                                .strDiagDeptID = lngDeptID
                                .strDiagDeptName = strDeptName
                                .strDiagDate = gobjDiags.Item(i).str���ʱ��
                                .strDiagName = gobjDiags.Item(i).str�������
                                .strDiagCode = gobjDiags.Item(i).str��ϱ���
                            Else
                                .strDiagID = gobjDiags.Item(i).str����ID
                                .strDiagDeptID = lngDeptID
                                .strDiagDeptName = strDeptName
                                .strDiagDate = gobjDiags.Item(i).str���ʱ��
                                .strDiagName = gobjDiags.Item(i).str�������
                                .strDiagCode = gobjDiags.Item(i).str��ϱ���
                            End If
                        End If
                    End With
                    colTmp.Add udtDiag, "_" & colTmp.Count + 1
                Next
            End If
            colXML.Add colTmp
        Else
            Set colTmp = New Collection
            Set rsTmp = Get������ϼ�¼(lngPatiID, IIf(str�Һŵ� <> "", lng�Һ�ID, lng��ҳID), IIf(str�Һŵ� <> "", "1,11", "2,12"))
            For i = 1 To rsTmp.RecordCount
                With udtDiag
                    .strDiagID = rsTmp!id
                    .strDiagDeptID = lngDeptID
                    .strDiagDeptName = strDeptName
                    .strDiagDate = Format(rsTmp!��¼���� & "", "YYYY-MM-DD HH:MM:SS")
                    .strDiagName = rsTmp!���� & ""
                    .strDiagCode = rsTmp!���� & ""
                End With
                colTmp.Add udtDiag, "_" & colTmp.Count + 1
                rsTmp.MoveNext
            Next
            colXML.Add colTmp
        End If
        '������¼ ����ȡһ��ҩƷID����
        'allergies������Ϣ��ǩ
        Set rsTmp = Get���˹�����¼(lngPatiID, IIf(str�Һŵ� <> "", 0, lng��ҳID))
        Set colTmp = New Collection
        For i = 1 To rsTmp.RecordCount
            strDrugID = ""
            If rsTmp!ҩ��ID & "" <> "" Then
                strSQL = "select ҩƷID from ҩƷ��� where ҩ��id=[1] and rownum <2"
                Set rs��� = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, rsTmp!ҩ��ID)
                If Not rs���.EOF Then strDrugID = rs���!ҩƷID & ""
            End If
            With udtAllergy
                .strAllergyID = strDrugID
                .strAllergyDrug = rsTmp!ҩ���� & ""
                .strAnaphylaxis = rsTmp!������Ӧ & ""
                .strRecordTime = rsTmp!��¼ʱ�� & ""
            End With
            colTmp.Add udtAllergy, "_" & colTmp.Count + 1
            rsTmp.MoveNext
        Next
        colXML.Add colTmp
        '���벡��������¼operations������Ϣ��ǩ
        Set colTmp = New Collection
        Set rsTmp = GetPatiOperation(lngPatiID, lng��ҳID, str�Һŵ�)
        For i = 1 To rsTmp.RecordCount
            With udtOpera
                .strOperationID = rsTmp!id & ""
                .strOperationCode = rsTmp!���� & ""
                .strOperationName = rsTmp!���� & ""
                .strOperationStartTime = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS")
            End With
            colTmp.Add udtOpera, "_" & colTmp.Count + 1
            rsTmp.MoveNext
        Next
        colXML.Add colTmp
        If bytFunc = 4 Then
            strUrl = "http://" & gstrIP & ":" & gstrPortPlus & "/v4/valid"
        Else
            strUrl = "http://" & gstrIP & ":" & gstrPortPlus & "/v4/engineAsync"
        End If
        strXML = HZYY_GetRootXML(colXML, IIf(str�Һŵ� <> "", 0, 1))
        strXML = Replace(strXML, "<![CDATA[]]>", "")
        strXML = "charset=utf-8&post_type=1&xml=" & strXML  'ҽ����ҩ
        WriteLog "" & glngModel, "HttpPost", "����ֵ:" & strXML
        strRet = HttpPost(strUrl, strXML, ResponseText, "application/x-www-form-urlencoded; charset=utf-8")
        strRet = Replace(strRet, "<![CDATA[]]>", "")
        WriteLog "" & glngModel, "HttpPost", "����ֵ:" & strRet
        
        If bytFunc = 4 Then AdviceCheckWarn_HZYY = True: Exit Function
        
        If strRet = "����ת����Ԥʧ��" Then
            MsgBox "������ҩ���:" & strRet, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        ElseIf strRet = "" Then
            MsgBox "��ǰ�����˺�����ҩϵͳ�������ں�����ҩ�ӿ�(���������ؿ�)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation + vbOKOnly, gstrSysName
        Else
            Set rsTmp = HZYY_ParseXML(strRet, strTmp, IIf(str�Һŵ� <> "", 0, 1))
            If Not rsTmp Is Nothing Then
                If rsTmp.RecordCount > 0 Then
                    strTmp = Replace(strTmp, "����:XX", "����:" & IIf(str�Һŵ� <> "", udtOptPati.strName, udtIptPati.strName))
                    frmPassResult.ShowMe gfrmMain, rsTmp, strTmp, bytRet, bytFunc, blnIsHaveOut
                End If
            End If
            If bytRet = 1 Then
                Exit Function
            End If
        End If
    Case 2, 3   'ɾ��������ҽ��
        If bytFunc = 2 Then
            strҽ��IDs = Replace(strҽ��IDs, "����ҩ��", "")
            strҽ��IDs = Replace(strҽ��IDs, "����ҩ��", "")
            If strҽ��IDs <> "" Then
                arrTemp = Split(strҽ��IDs, ",")
                Set colRecipe = New Collection
                For i = LBound(arrTemp) To UBound(arrTemp)
                    With udtOptPres
                        .udtOptPresInfo.strRecipeId = arrTemp(i)
                        .udtOptPresInfo.strRecipeNo = arrTemp(i)
                    End With
                    colRecipe.Add udtOptPres
                Next
                colXML.Add colRecipe
            End If
        ElseIf bytFunc = 3 Then
            'ҽ��ID1,ҽ��ID2|��ID
            If strҽ��IDs <> "" Then
                If InStr(strҽ��IDs, "����ҩ��") > 0 Then
                    strҽ��IDs = Replace(strҽ��IDs, "����ҩ��", "")
                    strTmp = Split(strҽ��IDs, "|")(1)
                    arrTemp = Split(Split(strҽ��IDs, "|")(0), ",")
                    Set udtOrder.colMedical = New Collection
                    For i = LBound(arrTemp) To UBound(arrTemp)
                        udtOrderItem.strOrderId = arrTemp(i)
                        udtOrderItem.strGroupNO = strTmp
                        udtOrder.colMedical.Add udtOrderItem
                    Next
                ElseIf InStr(strҽ��IDs, "����ҩ��") > 0 Then
                    strҽ��IDs = Replace(strҽ��IDs, "����ҩ��", "")
                    strTmp = Split(strҽ��IDs, "|")(1)
                    arrTemp = Split(Split(strҽ��IDs, "|")(0), ",")
                    Set udtOrder.colHerbMedical = New Collection
                    For i = LBound(arrTemp) To UBound(arrTemp)
                        udtHerbOrder.udtHerbInfo.strOrderId = arrTemp(i)
                        udtOrder.colHerbMedical.Add udtHerbOrder
                    Next
                End If
                colXML.Add udtOrder
            End If
        End If
        'ɾ��ҽ��������v
        strUrl = "http://" & gstrIP & ":" & gstrPortPlus & "/v4/invalid"
        strXML = HZYY_GetRootXML(colXML, IIf(str�Һŵ� <> "", 0, 1), 1)
        strXML = Replace(strXML, "<![CDATA[]]>", "")
        strXML = "charset=utf-8&post_type=1&xml=" & strXML  'ҽ����ҩ
        WriteLog "" & glngModel, "HttpPost", "����ֵ:" & strXML
        strRet = HttpPost(strUrl, strXML, ResponseText, "application/x-www-form-urlencoded; charset=utf-8")
        strRet = Replace(strRet, "<![CDATA[]]>", "")
        WriteLog "" & glngModel, "HttpPost", "����ֵ:" & strRet
    End Select
    
    AdviceCheckWarn_HZYY = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    AdviceCheckWarn_HZYY = False
End Function

Public Function CreateAdviceRS_HZYY(Optional ByRef rsOut As ADODB.Recordset, Optional ByRef rsDrug As ADODB.Recordset, _
    Optional ByVal strҽ��IDs As String, Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal str�Һŵ� As String) As ADODB.Recordset
'����;����ҽ����¼��
    Dim i As Long, k As Long, lngCount As Long, lngPos As Long
    Dim blnDo As Boolean, blnIsHaveOut As Boolean
    Dim strҩƷ As String, strҽ��ID As String, str���ID As String
    Dim str����ʱ�� As String
    Dim str��Ч As String, str���� As String, str������λ As String, strƵ�� As String
    Dim str��ҩ;�� As String, strƵ�ʱ��� As String, str�÷� As String, str�÷�ID As String, str��ʼʱ�� As String, str����ʱ�� As String
    Dim str��������Tag As String, str��������ID As String, str������ĿIDs As String, strҩƷID As String
    Dim str����ҽ��Tag As String, str����ҽ�� As String, str��ҩĿ�� As String
    Dim str���� As String, str������λ As String, strType As String
    Dim str����ID, str�շ�ϸĿID As String, str���� As String, str��Һ As String
    Dim str������   As String
    
    Dim rsAdvice As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsƵ�� As ADODB.Recordset
    Dim rs����ҽ�� As ADODB.Recordset
    Dim rs��������  As ADODB.Recordset
    Dim rsҩƷ As ADODB.Recordset
    Dim rsְ�� As ADODB.Recordset
    
    Dim curDate As Date
    
    On Error GoTo errH
    
    curDate = zlDatabase.Currentdate
    Set rsAdvice = InitAdviceRS(FUN_ҽ����Ϣ_HZYY)

    Select Case glngModel
    Case PM_����༭, PM_סԺ�༭
        '�����˽���ҩƷ˵������;����Ϊ����༭\סԺ�༭;��鹦��
        If (glngModel = PM_����༭ Or glngModel = PM_סԺ�༭) And gbytReason = 1 Then
            Set rsOut = InitAdviceRS(FUN_�������)
        End If
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_����༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                            And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-DD") = Format(curDate, "yyyy-MM-DD")
                ElseIf glngModel = PM_סԺ�༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                    If blnDo Then
                        blnDo = (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0 _
                                Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                                And .TextMatrix(i, gobjCOL.intCOL״̬) <> "4")
                    End If
                End If

                If blnDo Then
                    str����ID = .TextMatrix(i, gobjCOL.intCOL������ĿID)
                    If InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                        If InStr("," & str������ĿIDs & ",", "," & str����ID & ",") = 0 Then
                            str������ĿIDs = str������ĿIDs & "," & str����ID
                        End If
                    End If
                    strҽ��ID = CStr(.RowData(i))

                    'ȡҩƷ����
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                        strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                    Else
                        strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                    End If
                    'ȡҩƷ��ҩ;��
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str�÷� = ""    'һ����ҩ���ظ�ȡ

                    If str�÷� = "" Then
                        str���� = "": str��Һ = "0"
                        If glngModel = PM_����༭ Or glngModel = PM_סԺ�༭ Then
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                            If k <> -1 Then
                                If .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                                    str�÷� = .TextMatrix(k, gobjCOL.intCOL�÷�)
                                Else
                                    str�÷� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                    If InStr(.TextMatrix(k, gobjCOL.intcolҽ������), "��/����") > 0 Or InStr(.TextMatrix(k, gobjCOL.intcolҽ������), "����/Сʱ") > 0 Then
                                        str���� = .TextMatrix(k, gobjCOL.intcolҽ������)
                                    End If
                                    If Val(.TextMatrix(k, gobjCOL.intColִ�з���)) = 1 Then
                                        str��Һ = "1"
                                    Else
                                        str��Һ = "0"
                                    End If
                                End If
                                str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                            End If
                        End If
                    End If
                    '������������
                    str��������ID = .TextMatrix(i, gobjCOL.intCOL��������ID)
                    If InStr("," & str��������Tag & ",", "," & str��������ID & ",") = 0 Then
                        str��������Tag = str��������Tag & "," & str��������ID
                    End If

                    '����ҽ��
                    str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                    If InStr("," & str����ҽ��Tag & ",", "," & str����ҽ�� & ",") = 0 Then
                        str����ҽ��Tag = str����ҽ��Tag & "," & str����ҽ��
                    End If

                    str��ʼʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:MM:SS")
'
                    str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd HH:mm:ss")         '����ʱ�䣨YYYY-MM-DD HH:mm:SS��
                    '������������λ
                    str���� = .TextMatrix(i, gobjCOL.intCOL����)
                    str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                    str���� = .TextMatrix(i, gobjCOL.intCOL����)
                    str������λ = .TextMatrix(i, gobjCOL.intcol������λ)

                    strҩƷID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)

                    If glngModel = PM_����༭ Then
                        str����ʱ�� = ""
                        str��Ч = "2" '2-��ʱҽ��
                        strType = .TextMatrix(i, gobjCOL.intCOL״̬)
                        If strType = "4" Then
                            strType = "2"       '���ϴ���
                        ElseIf strType = "1" Then
                            strType = "0"       '��������
                        End If
                        str������ = .TextMatrix(i, gobjCOL.intCol������)
                    ElseIf glngModel = PM_סԺ�༭ Then
                        str��Ч = IIf(.TextMatrix(i, gobjCOL.intCOL��Ч) = "����", 1, 2)
                        str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ֹʱ��), "yyyy-MM-dd HH:MM:SS")
                        '�ж��Ƿ���Ժ��ִ�е�ҩƷ
                        If Val(.TextMatrix(i, gobjCOL.intCOLִ������)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID))), gobjCOL.intCOLִ������)) = 5 Then
                            blnIsHaveOut = True
                            str��Ч = "3"
                        End If
                        str������ = ""
                    End If

                    '����˵��
                    If Not rsOut Is Nothing Then
                        If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                        '��ҩ,�г�ҩ
                            rsOut.AddNew
                            rsOut!ҽ��id = CLng(strҽ��ID)
                            rsOut!����ҩƷ˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                            rsOut!״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                            rsOut!ҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������)
                            rsOut.Update
                        ElseIf Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                        '��ҩ�䷽  ����˵����������ҩ������
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                            If k <> -1 Then
                                rsOut.AddNew
                                rsOut!ҽ��id = CLng(.RowData(k) & "")
                                rsOut!����ҩƷ˵�� = .TextMatrix(k, gobjCOL.intCol����ҩƷ˵��)
                                rsOut!״̬ = .TextMatrix(k, gobjCOL.intCOL״̬)
                                rsOut!ҩƷ���� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                rsOut.Update
                            End If
                        End If
                    End If
                    str��ҩĿ�� = .TextMatrix(i, gobjCOL.intcol��ҩĿ��)
                    If str��ҩĿ�� = "1" Then
                        str��ҩĿ�� = "Ԥ����ҩ"
                    ElseIf str��ҩĿ�� = "2" Then
                        str��ҩĿ�� = "������ҩ"
                    Else
                        str��ҩĿ�� = ""
                    End If
                    '----------------------------------------------------------
                    rsAdvice.AddNew
                    rsAdvice!����ID = Val(str������)
                    rsAdvice!ҽ��id = strҽ��ID
                    rsAdvice!���ID = .TextMatrix(i, gobjCOL.intCOL���ID)
                    rsAdvice!ҽ����Ч = str��Ч
                    rsAdvice!ҽ����� = .TextMatrix(i, gobjCOL.intCOL���)
                    rsAdvice!��������id = str��������ID
                    rsAdvice!����ҽ�� = str����ҽ��
                    rsAdvice!������ĿID = str����ID
                    rsAdvice!ҩƷID = Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID))
                    rsAdvice!ҩƷ���� = strҩƷ
                    rsAdvice!ҽ��״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                    rsAdvice!�������� = str����
                    rsAdvice!������λ = str������λ
                    rsAdvice!Ƶ�� = .TextMatrix(i, gobjCOL.intCOLƵ��)
                    rsAdvice!�÷� = str�÷�
                    rsAdvice!�÷�ID = str��ҩ;��
                    rsAdvice!����ʱ�� = str����ʱ��
                    rsAdvice!��ʼʱ�� = str��ʼʱ��
                    rsAdvice!����ʱ�� = str����ʱ��
                    rsAdvice!���� = str����
                    rsAdvice!������λ = str������λ
                    rsAdvice!���� = .TextMatrix(i, gobjCOL.intCOL����)
                    rsAdvice!ҽ������ = .TextMatrix(i, gobjCOL.intcolҽ������)
                    rsAdvice!��ҩĿ�� = str��ҩĿ��
                    rsAdvice!��ҩ���� = .TextMatrix(i, gobjCOL.intcol��ҩ����)
                    rsAdvice!������� = .TextMatrix(i, gobjCOL.intCOL�������)
                    rsAdvice!��־ = .TextMatrix(i, gobjCOL.intCol��־)
                    rsAdvice!���� = str����
                    rsAdvice!��Һ = str��Һ
                    rsAdvice!��Ժ��ҩ = IIf(blnIsHaveOut, 1, 0)
                    rsAdvice.Update
                    '----------------------------------------------------------------------------
                End If
            Next
        End With
    Case PM_����ҽ���嵥, PM_סԺҽ���嵥
        Set rsTmp = GetAdviceInfo_YF(gobjPati.lng����ID, gobjPati.lng��ҳID, gobjPati.str�Һŵ�, , 1)
        With rsTmp
            If rsTmp.RecordCount = 0 Then Set CreateAdviceRS_HZYY = rsAdvice: Exit Function
            For i = 1 To .RecordCount
                If glngModel = PM_����ҽ���嵥 Then
                    blnDo = InStr(",5,6,7,", "," & !������� & ",") > 0 And Val(!�շ�ϸĿid & "") <> 0 And Format(!����ʱ�� & "", "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                ElseIf glngModel = PM_סԺҽ���嵥 Then
                    blnDo = InStr(",5,6,7,", "," & !������� & ",") > 0 And Not InStr(",4,8,9,", "," & !ҽ��״̬ & ",") > 0
                End If
                If blnDo Then

                    If InStr(",5,6,7,", "," & !������� & ",") > 0 And Not InStr(",4,8,9,", "," & !ҽ��״̬ & ",") > 0 And Val(!�շ�ϸĿid & "") = 0 Then
                        If InStr("," & str������ĿIDs & ",", "," & !������ĿID & ",") = 0 Then
                            str������ĿIDs = str������ĿIDs & "," & !������ĿID
                        End If
                    End If
                    '����ҽ��
                    str����ҽ�� = !����ҽ�� & ""
                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                    If InStr("," & str����ҽ��Tag & ",", "," & str����ҽ�� & ",") = 0 Then
                        str����ҽ��Tag = str����ҽ��Tag & "," & str����ҽ��
                    End If

                    If gobjPati.str�Һŵ� <> "" Then
                        str������λ = !���ﵥλ & ""
                    Else
                        str������λ = !סԺ��λ & ""
                    End If

                    If InStr(";" & strƵ�� & ";", ";" & !Ƶ�� & "," & IIf(!������� & "" = "7", 2, 1) & ";") = 0 Then
                        strƵ�� = strƵ�� & ";" & !Ƶ�� & "," & IIf(!������� = "7", 2, 1)
                    End If

                    rsAdvice.AddNew
                    rsAdvice!����ID = Val(!����ID & "")
                    rsAdvice!ҽ��id = !ҽ��id & ""
                    rsAdvice!���ID = !���ID & ""
                    rsAdvice!ҽ����Ч = IIf(Val(!ҽ����Ч & "") = 0, 1, 2) 'HZYY 1����;2����;3��Ժ��ҩ
                    If !Aִ������ & "" <> "5" And !Bִ������ & "" = "5" Then
                        rsAdvice!ҽ����Ч = "3"
                    End If
                    rsAdvice!ҽ����� = lngCount + 1
                    rsAdvice!��������id = !��������id & ""
                    rsAdvice!�������� = !�������� & ""
                    rsAdvice!����ҽ�� = str����ҽ��
                    rsAdvice!������ĿID = !������ĿID & ""
                    rsAdvice!ҩƷID = !�շ�ϸĿid & ""
                    rsAdvice!ҩƷ���� = !ҩƷ���� & ""
                    rsAdvice!ҽ��״̬ = !ҽ��״̬ & ""
                    rsAdvice!�������� = !�������� & ""
                    rsAdvice!������λ = !������λ & ""
                    rsAdvice!Ƶ�� = !Ƶ�� & ""
                    rsAdvice!�÷� = !�÷� & ""
                    rsAdvice!�÷�ID = !�÷�ID & ""
                    rsAdvice!����ʱ�� = !����ʱ�� & ""
                    rsAdvice!��ʼʱ�� = !��ʼʱ�� & ""
                    rsAdvice!����ʱ�� = !����ʱ�� & ""
                    rsAdvice!���� = !���� & ""
                    rsAdvice!������λ = str������λ
                    rsAdvice!���� = !���� & ""
                    rsAdvice!ҽ������ = !ҽ������ & ""
                    
                    If Val(!��ҩĿ�� & "") = 1 Then
                        rsAdvice!��ҩĿ�� = "Ԥ����ҩ"
                    ElseIf Val(!��ҩĿ�� & "") = 2 Then
                        rsAdvice!��ҩĿ�� = "������ҩ"
                    End If
                    
                    rsAdvice!��ҩ���� = !��ҩ���� & ""
                    rsAdvice!������� = !������� & ""
                    rsAdvice!��� = !��� & ""
                    rsAdvice!��־ = !��־ & ""
                    If !��� & "_" & !�������� & "_" & !ִ�з��� = "E_2_1" Then
                        rsAdvice!��Һ = "1"
                    Else
                        rsAdvice!��Һ = "0"
                    End If
                    rsAdvice.Update
                End If
                .MoveNext
            Next
        End With
    Case PM_PIVA����, PM_���ŷ�ҩ, PM_������ҩ
        Set rsTmp = GetAdviceInfo_YF(lng����ID, lng��ҳID, str�Һŵ�)
        With rsTmp
            If rsTmp.RecordCount = 0 Then Set CreateAdviceRS_HZYY = rsAdvice: Exit Function
            For i = 1 To .RecordCount

                If Val(!�շ�ϸĿid & "") = 0 Then
                    If InStr("," & str������ĿIDs & ",", "," & !������ĿID & ",") = 0 Then
                        str������ĿIDs = str������ĿIDs & "," & !������ĿID
                    End If
                End If
                '����ҽ��
                str����ҽ�� = !����ҽ�� & ""
                If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                If InStr("," & str����ҽ��Tag & ",", "," & str����ҽ�� & ",") = 0 Then
                    str����ҽ��Tag = str����ҽ��Tag & "," & str����ҽ��
                End If

                If str�Һŵ� <> "" Then
                    str������λ = !���ﵥλ & ""
                Else
                    str������λ = !סԺ��λ & ""
                End If

                If InStr(";" & strƵ�� & ";", ";" & !Ƶ�� & "," & IIf(!������� & "" = "7", 2, 1) & ";") = 0 Then
                    strƵ�� = strƵ�� & ";" & !Ƶ�� & "," & IIf(!������� = "7", 2, 1)
                End If

                rsAdvice.AddNew
                rsAdvice!ҽ��id = !ҽ��id & ""
                rsAdvice!���ID = !���ID & ""
                rsAdvice!ҽ����Ч = !ҽ����Ч & ""
                rsAdvice!ҽ����� = lngCount + 1
                rsAdvice!��������id = !��������id & ""
                rsAdvice!�������� = !�������� & ""
                rsAdvice!����ҽ�� = str����ҽ��
                rsAdvice!������ĿID = !������ĿID & ""
                rsAdvice!ҩƷID = !�շ�ϸĿid & ""
                rsAdvice!ҩƷ���� = !ҩƷ���� & ""
                rsAdvice!ҽ��״̬ = !ҽ��״̬ & ""
                rsAdvice!�������� = !�������� & ""
                rsAdvice!������λ = !������λ & ""
                rsAdvice!Ƶ�� = !Ƶ�� & ""
                rsAdvice!�÷� = !�÷� & ""
                rsAdvice!�÷�ID = !�÷�ID & ""
                rsAdvice!����ʱ�� = !����ʱ�� & ""
                rsAdvice!��ʼʱ�� = !��ʼʱ�� & ""
                rsAdvice!����ʱ�� = !����ʱ�� & ""
                rsAdvice!���� = !���� & ""
                rsAdvice!������λ = str������λ
                rsAdvice!���� = !���� & ""
                rsAdvice!ҽ������ = !ҽ������ & ""
                rsAdvice!��ҩĿ�� = !��ҩĿ�� & ""
                rsAdvice!��ҩ���� = !��ҩ���� & ""
                rsAdvice!������� = !������� & ""
                rsAdvice!��� = !��� & ""
                rsAdvice!��־ = !��־ & ""
                rsAdvice.Update

                .MoveNext
            Next
        End With
    End Select

    '����������ȡ
    If rsAdvice.RecordCount > 0 Then

        rsAdvice.MoveFirst
        Select Case glngModel

        Case PM_����༭, PM_����ҽ���嵥, PM_סԺ�༭, PM_סԺҽ���嵥, PM_PIVA����, PM_���ŷ�ҩ, PM_������ҩ
            If str������ĿIDs <> "" Then
                str������ĿIDs = Mid(str������ĿIDs, 2)
                Set rsҩƷ = GetRS("ҩƷ���", "ҩ��id,ҩƷid", str������ĿIDs, "ҩ��id")
            End If
            If strƵ�� <> "" Then Set rsƵ�� = GetRS("����Ƶ����Ŀ", "����, ����, ���÷�Χ", strƵ��, "����, ���÷�Χ", 1, 2)
            If str��������Tag <> "" Then Set rs�������� = GetRS("���ű�", "ID,����", str��������Tag)
            If str����ҽ��Tag <> "" Then Set rs����ҽ�� = GetRS("��Ա�� A,רҵ����ְ�� B", "A.ID,A.����,A.רҵ����ְ��,B.����", str����ҽ��Tag, " A.רҵ����ְ��=B.���� And A.����", 0, 1)

            For i = 1 To rsAdvice.RecordCount
                 '����ҽ����Ʒ���´�ʱ,����ȡһ��ҩƷId
                If Val(rsAdvice!ҩƷID & "") = 0 And Val(rsAdvice!ҽ����Ч & "") = 0 Then
                    If Not rsҩƷ Is Nothing Then
                        rsҩƷ.Filter = "ҩ��ID =" & rsAdvice!������ĿID
                        If Not rsҩƷ.EOF Then rsAdvice!ҩƷID = rsҩƷ!ҩƷID & ""
                    End If
                End If

                If InStr("," & str�շ�ϸĿID & ",", "," & rsAdvice!ҩƷID & ",") = 0 Then
                    str�շ�ϸĿID = str�շ�ϸĿID & "," & rsAdvice!ҩƷID
                End If

                If Not rsƵ�� Is Nothing Then
                    rsƵ��.Filter = "���� ='" & rsAdvice!Ƶ�� & "' And ���÷�Χ=" & IIf(rsAdvice!������� & "" = "7", 2, 1)
                    If Not rsƵ��.EOF Then rsAdvice!Ƶ�ʱ��� = rsƵ��!���� & ""
                End If

                If Not rs����ҽ�� Is Nothing Then
                    rs����ҽ��.Filter = "����='" & rsAdvice!����ҽ�� & "'"
                    If Not rs����ҽ��.EOF Then
                        rsAdvice!����ҽ��ID = rs����ҽ��!id & ""
                        rsAdvice!רҵ����ְ�� = rs����ҽ��!���� & ""
                    End If
                End If
                If Not rs�������� Is Nothing Then
                    rs��������.Filter = "ID =" & rsAdvice!��������id
                    If Not rs��������.EOF Then rsAdvice!�������� = rs��������!���� & ""
                End If
                rsAdvice.MoveNext
            Next

            If str�շ�ϸĿID <> "" Then
                str�շ�ϸĿID = Mid(str�շ�ϸĿID, 2)
                Set rsDrug = GetҩƷ��Ϣ_HZYY(str�շ�ϸĿID)
            End If
        End Select
        rsAdvice.MoveFirst
    End If
    Set CreateAdviceRS_HZYY = rsAdvice
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function HZYY_ParseXML(ByVal strData As String, ByRef strPati As String, ByVal bytFunc As Byte) As ADODB.Recordset
    '����:����XML�ַ���
    'bytFunc=0-����;1-סԺ
    '<root>
    '  <base>
    '    <hospital_code><![CDATA[ҽԺCode(����)]]></hospital_code>
    '    <event_no><![CDATA[������ˮ��]]></event_no>
    '    <patient_id><![CDATA[���˺�(��Ҫ����Ψһ��ʶ)]]></patient_id>
    '    <source><![CDATA[��Դ]]></source>
    '  </base>
    '  <pharm_chk_id><![CDATA[���ҩʦ����]]></pharm_chk_id> ----�󷽲��д˱�ǩ
    '  <pharm_chk_name><![CDATA[���ҩʦ����]]></pharm_chk_name> ----�󷽲��д˱�ǩ
    '  <btnStatus><![CDATA[��ť����ֵ(1�޸Ĵ�����2����)]]></btnStatus>----��Ԥ���д˱�ǩ
    '  <message>
    '    <recipe_id><![CDATA[����id]]></recipe_id>
    '    <is_success><![CDATA[�ɹ���ʶ(0��˲�ͨ����1���ͨ��)]]></is_success>----�󷽲��д˱�ǩ
    '    <infos>
    '      <info>
    '        <info_id><![CDATA[һ����ʾ��Ϣ��Ψһid]]></info_id>
    '        <group_no><![CDATA[���]]></group_no>
    '        <drug_id><![CDATA[ҩƷid]]></drug_id>
    '        <drug_name><![CDATA[ҩƷ����]]></drug_name>
    '        <error_info><![CDATA[������Ϣ]]></error_info>
    '        <advice><![CDATA[����]]></advice>
    '        <source><![CDATA[��Դ]]></source>
    '        <rt><![CDATA[��Ϣ�Ĺ�������]]></rt>
    '        <source_id><![CDATA[��Դid]]></source_id>
    '        <severity><![CDATA[����ȼ�]]></severity>
    '        <message_id><![CDATA[������Ϣid]]></message_id>
    '        <type><![CDATA[��ʾ��Ϣ����]]></type>
    '        <analysis_type><![CDATA[��������]]></analysis_type>
    '        <analysis_result_type><![CDATA[��ʾ����]]></analysis_result_type>
    '        <status><![CDATA[״̬:1��Ҫ˫ǩ��ȷ��0����Ҫ˫ǩ��ȷ��]]></status>----�󷽲��д˱�ǩ
    '      </info>
    '      <info>
    '             һ��<info>��ǩ��һ����ʾ��Ϣ��������ʾ��Ϣ���<info>��ǩ
    '      </info>
    '</infos>
    '  </message>
    '  <message>
    '      ����һ��xml�д�����Ŵ�����������ϸ��Ϣ�����ݴ��뷽ʽ��ϵͳ�践�ض���<message>��ǩ
    '  </message>
    '</root>
        Dim xmlDoc As New DOMDocument
        Dim xRoot As IXMLDOMElement
        Dim xNode As IXMLDOMNode
        Dim xmlInfos As IXMLDOMNodeList
        Dim strValue As String
        Dim strRecipeId As String
        Dim rsRet As ADODB.Recordset
        Dim i As Long, j As Long, lngCount As Long
        On Error GoTo errH
100     Set rsRet = InitAdviceRS(FUN_�����_HZYY)
        '��ȡ������Ӧ���ݣ�XML��ʽ��
'strData = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine & _
'            "<root><base><hospital_code>cqzl123</hospital_code><event_no><![CDATA[Q0000001]]></event_no>" & _
'            "<patient_id>101</patient_id><source><![CDATA[����]]></source></base><btnStatus></btnStatus>" & _
'            "<message><recipe_id><![CDATA[118]]></recipe_id><infos><info><info_id></info_id><group_no></group_no>" & _
'            "<drug_id><![CDATA[10010]]></drug_id><drug_name><![CDATA[��Ī���ֽ���]]></drug_name><error_info>" & _
'            "<![CDATA[��ҩ;�������ʡ�]]></error_info><advice><![CDATA[��Ʒ�˿ڷ���ҩ��]]></advice>" & _
'            "<source><![CDATA[1������]]></source><rt><![CDATA[0]]></rt><source_id><![CDATA[SFDAҩƷ˵���鷶��]]></source_id>" & _
'            "<severity><![CDATA[8]]></severity><message_id><![CDATA[1510194380657]]></message_id><type><![CDATA[��ҩ;��]]></type>" & _
'            "<analysis_type><![CDATA[�����Է���]]></analysis_type><analysis_result_type><![CDATA[��ҩ;��]]>" & _
'            "</analysis_result_type><status></status></info></infos></message><version><![CDATA[V1.0]]></version></root>"
'  strData = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbNewLine & _
'        "<root>" & vbNewLine & _
'        "  <base>" & vbNewLine & _
'        "    <hospital_code>1001</hospital_code>" & vbNewLine & _
'        "    <event_no><![CDATA[17080894]]></event_no>" & vbNewLine & _
'        "    <patient_id>3394518</patient_id>" & vbNewLine & _
'        "    <source><![CDATA[סԺ]]></source>" & vbNewLine & _
'        "  </base>" & vbNewLine & _
'        "  <btnStatus/>"
'סԺ
'strData = strData & "<message>" & vbNewLine & _
'"  <group_no/>" & vbNewLine & _
'"  <infos>" & vbNewLine & _
'"    <info>" & vbNewLine & _
'"      <info_id/>" & vbNewLine & _
'"      <order_id><![CDATA[120721918]]></order_id>" & vbNewLine & _
'"      <order_item_id/>" & vbNewLine & _
'"      <drug_id><![CDATA[61851]]></drug_id>" & vbNewLine & _
'"      <drug_name><![CDATA[��Ī���ֿ���ά���Ƭ(2:1)]]></drug_name>" & vbNewLine & _
'"      <order_type/>" & vbNewLine & _
'"      <pivas_flag/>" & vbNewLine & _
'"      <error_info><![CDATA[��ҩ;�������ʡ�]]></error_info>" & vbNewLine & _
'"      <advice><![CDATA[��ҩ�˿ڷ�θ������ҩ��]]></advice>" & vbNewLine & _
'"      <source><![CDATA[˵����]]></source>" & vbNewLine & _
'"      <rt><![CDATA[0]]></rt>" & vbNewLine & _
'"      <source_id/>" & vbNewLine & _
'"      <severity><![CDATA[5]]></severity>" & vbNewLine & _
'"      <message_id><![CDATA[1383841426395]]></message_id>" & vbNewLine & _
'"      <type><![CDATA[��ҩ;��]]></type>" & vbNewLine & _
'"      <analysis_type><![CDATA[�����Է���]]></analysis_type>" & vbNewLine & _
'"      <analysis_result_type><![CDATA[��ҩ;��]]></analysis_result_type>" & vbNewLine & _
'"      <status/>" & vbNewLine & _
'"    </info>"
'strData = strData & "      <info>" & vbNewLine & _
'"        <info_id/>" & vbNewLine & _
'"        <order_id><![CDATA[120721918]]></order_id>" & vbNewLine & _
'"        <order_item_id/>" & vbNewLine & _
'"        <drug_id><![CDATA[61851]]></drug_id>" & vbNewLine & _
'"        <drug_name><![CDATA[��ҩ;������]]></drug_name>" & vbNewLine & _
'"        <order_type/>" & vbNewLine & _
'"        <pivas_flag/>" & vbNewLine & _
'"        <error_info><![CDATA[���ԣ�����]]></error_info>" & vbNewLine & _
'"        <advice/>" & vbNewLine & _
'"        <source><![CDATA[0������]]></source>" & vbNewLine & _
'"        <rt><![CDATA[0]]></rt>" & vbNewLine & _
'"        <source_id/>" & vbNewLine & _
'"        <severity><![CDATA[8]]></severity>" & vbNewLine & _
'"        <message_id><![CDATA[1513238322112]]></message_id>" & vbNewLine & _
'"        <type><![CDATA[�������������]]></type>" & vbNewLine & _
'"        <analysis_type><![CDATA[��ҩ����]]></analysis_type>" & vbNewLine & _
'"        <analysis_result_type/>" & vbNewLine & _
'"        <status/>" & vbNewLine & _
'"      </info>" & vbNewLine & _
'"    </infos>" & vbNewLine & _
'"  </message>" & vbNewLine & _
'"  <version><![CDATA[V1.0]]></version>" & vbNewLine & _
'"</root>"

 
102     xmlDoc.loadXML (strData)
104     If bytFunc = 0 Then
106         strPati = ""
108         strValue = "NO:" & xmlDoc.selectSingleNode(".//event_no").Text & Space(4)
110         strPati = strPati & strValue
112         strValue = "����ID:" & xmlDoc.selectSingleNode(".//patient_id").Text & Space(4)
114         strPati = strPati & strValue
116         strPati = strPati & "����:XX" & Space(4)
        
118         Set xRoot = xmlDoc.selectSingleNode("root")
120         For Each xNode In xRoot.childNodes
122             If xNode.nodeName = "message" Then
124                 strRecipeId = xNode.selectSingleNode(".//recipe_id").Text
126                 Set xmlInfos = xNode.selectNodes(".//info")
128                 For i = 0 To xmlInfos.length - 1
                        On Error Resume Next
130                     strValue = xmlInfos(i).selectSingleNode(".//info_id").Text
132                     If Err.Number > 0 Then Exit For
134                     Err.Clear: On Error GoTo errH
136                     rsRet.AddNew
138                     rsRet!RecipeId = strRecipeId
140                     rsRet!drugID = xmlInfos(i).selectSingleNode(".//drug_id").Text
142                     rsRet!DrugName = xmlInfos(i).selectSingleNode(".//drug_name").Text   'ҩƷ����
144                     rsRet!message = xmlInfos(i).selectSingleNode(".//error_info").Text      '������Ϣ
146                     rsRet!Advice = xmlInfos(i).selectSingleNode(".//advice").Text        '��ҩ����
148                     rsRet!Source = xmlInfos(i).selectSingleNode(".//source").Text        '��Դ
150                     rsRet!GroupNo = xmlInfos(i).selectSingleNode(".//group_no").Text  '���
152                     rsRet!Type = xmlInfos(i).selectSingleNode(".//type").Text         '��ʾ��Ϣ����
154                     rsRet!Severity = xmlInfos(i).selectSingleNode(".//severity").Text '����ȼ�
156                     rsRet.Update
158                     xmlInfos.nextNode
                    Next
                End If
            Next
        Else
160         strPati = ""
162         strValue = "����ID:" & xmlDoc.selectSingleNode(".//patient_id").Text & Space(4)
164         strPati = strPati & strValue
166         strPati = strPati & "����:XX" & Space(4)
168         Set xRoot = xmlDoc.selectSingleNode("root")
170         For Each xNode In xRoot.childNodes
172             If xNode.nodeName = "message" Then
174                 strRecipeId = xNode.selectSingleNode(".//group_no").Text
176                 Set xmlInfos = xNode.selectNodes(".//info")
178                 For i = 0 To xmlInfos.length - 1
                        On Error Resume Next
180                     strValue = xmlInfos(i).selectSingleNode(".//info_id").Text
182                     If Err.Number > 0 Then Exit For
184                     Err.Clear: On Error GoTo errH
186                     rsRet.AddNew
188                     rsRet!RecipeId = 0
190                     rsRet!drugID = xmlInfos(i).selectSingleNode(".//drug_id").Text
192                     rsRet!DrugName = xmlInfos(i).selectSingleNode(".//drug_name").Text   'ҩƷ����
194                     rsRet!message = xmlInfos(i).selectSingleNode(".//error_info").Text      '������Ϣ
196                     rsRet!Advice = xmlInfos(i).selectSingleNode(".//advice").Text        '��ҩ����
198                     rsRet!Source = xmlInfos(i).selectSingleNode(".//source").Text        '��Դ
200                     rsRet!GroupNo = strRecipeId  '���
202                     rsRet!Type = xmlInfos(i).selectSingleNode(".//type").Text         '��ʾ��Ϣ����
204                     rsRet!Severity = xmlInfos(i).selectSingleNode(".//severity").Text '����ȼ�
206                     rsRet.Update
208                     xmlInfos.nextNode
                    Next
                End If
            Next
        End If
210     Set HZYY_ParseXML = rsRet
        Exit Function
errH:
212     MsgBox Err.Description & vbCrLf & "HZYY_ParseXML" & "�� " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

