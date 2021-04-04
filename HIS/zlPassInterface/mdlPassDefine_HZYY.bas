Attribute VB_Name = "mdlPassDefine_HZYY"
Option Explicit

Private Function Get药品信息_HZYY(ByVal strDrugIDs As String) As ADODB.Recordset
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errH
    strSQL = "Select a.药品id, a.门诊单位, a.门诊包装, a.住院单位, a.住院包装, d.药品剂型, e.规格,e.计算单位 , f.现价,B.厂家名称,C.编码 as 厂家编码 " & vbNewLine & _
            "From 药品规格 A, 药品特性 D, 收费项目目录 E, 收费价目 F, 药品生产商对照 B, 药品生产商 C" & vbNewLine & _
            "Where a.药名id = d.药名id And a.药品id = e.Id And a.药品id = f.收费细目id(+) And a.药品id = b.药品id(+) And b.厂家名称 = c.名称(+) And" & vbNewLine & _
            "      Nvl(f.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) > Sysdate And" & vbNewLine & _
            "      a.药品id In (Select /*+cardinality(A,10)*/" & vbNewLine & _
            "                  *" & vbNewLine & _
            "                 From Table(f_Num2list([1])) A)"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_HZYY", strDrugIDs)
    Set Get药品信息_HZYY = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HZYY_MakeBASEXML(ByRef xmlbase As HZYYBASE) As String
'功能：构造BASE XML字符串
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
'bytFunc  =0-处方审查;1-删除处方
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
'bytFunc  =0-处方审查;1-删除处方
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
'功能:获取医嘱信息XML
'bytType=0-干预;1-删除

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
'参数:bytFunc=0 医嘱审查;1-删除
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
                '删除医嘱
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
'参数:bytFunc=0 医嘱审查;1-删除
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
'功能:获取诊断信息XML
'参数: bytFunc=0  门诊诊断;=1 住院诊断
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
'功能:获取过敏信息XML
'参数: bytFunc=0  门诊诊断;=1 住院诊断
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
'功能:获取手术信息XML
'参数: bytFunc=0  门诊诊断;=1 住院诊断
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
'功能:获取检验信息XML
'参数: bytFunc=0  门诊诊断;=1 住院诊断
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
'功能:获取影像信息XML
'参数: bytFunc=0  门诊诊断;=1 住院诊断
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
'功能:特殊检查项目标签
'参数: bytFunc=0  门诊诊断;=1 住院诊断
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
'功能:获取门诊电子病历标签
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
'功能:获取入院记录
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
'功能:获取病程录标签
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
'功能:获取生命体征标签
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
'功能:获取病理信息标签
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
'功能:获取细菌培养报告标签
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
'功能:获取药物敏感试验信息标签
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
'功能:bytFunc=0-门诊
'     bytType=0-干预; 1=删除医嘱
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
'0   未知的性别
'1   男性
'2   女性
'9   未说明的性别
    Dim strRet As String
    If strPara = "" Then
        strRet = "9"
    ElseIf InStr(strPara, "男") > 0 Then
        strRet = "1"
    ElseIf InStr(strPara, "女") > 0 Then
        strRet = "2"
    Else
        strRet = "0"
    End If
    HZYY_GetSex = strRet
End Function

Public Function HZYY_GetMarital(ByVal strPara As String) As String
'10  未婚
'20  已婚
'21  初婚
'22  再婚
'23  复婚
'30  丧偶
'40  离婚
'90  未说明的婚姻状况
    Dim strRet As String
    If strPara = "" Then
        strRet = "90"
    ElseIf InStr(strPara, "未婚") > 0 Then
        strRet = "10"
    ElseIf InStr(strPara, "已婚") > 0 Then
        strRet = "20"
    ElseIf InStr(strPara, "初婚") > 0 Then
        strRet = "21"
    ElseIf InStr(strPara, "再婚") > 0 Then
        strRet = "22"
    ElseIf InStr(strPara, "复婚") > 0 Then
        strRet = "23"
    ElseIf InStr(strPara, "丧偶") > 0 Then
        strRet = "30"
    ElseIf InStr(strPara, "离婚") > 0 Then
        strRet = "40"
    Else
        strRet = ""
    End If
    HZYY_GetMarital = strRet
End Function

Public Sub HZYY_DrugInstructions(Optional ByVal lngDrugID As Long)
'功能：杭州逸曜药品说明
'参数:lngDrugID-药品ID
    Dim strUrl As String
    Dim lngRet As Long
    
    If lngDrugID = 0 And (Not gobjAdvice Is Nothing) And (glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Or _
        glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单) Then
        With gobjAdvice
            If InStr(",5,6,7,", .TextMatrix(.Row, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(.Row, gobjCOL.intCOL收费细目ID)) <> 0 Then
                '获取所选医嘱的药品编码
                lngDrugID = Val(.TextMatrix(.Row, gobjCOL.intCOL收费细目ID))
            End If
        End With
    End If
    'strUrl = "http://118.31.246.211:8080/zlcx/data_detail.action?webHisId=11221&hospitalCode=cqzl123"
    If gbytType = 0 Then
        strUrl = "http://" & gstrIP & ":" & gstrPort & "/zlcx/data_detail.action?webHisId=" & lngDrugID
    Else
        '产品非共用
        strUrl = "http://" & gstrIP & ":" & gstrPort & "/zlcx/data_detail.action?webHisId=" & lngDrugID & "&hospitalCode=" & gstrHOSCODE
    End If
    lngRet = ShellExecute(0, "open", strUrl, "", "", SW_SHOWNORMAL)
End Sub

Public Function HZYY_GetPara() As Boolean
        Dim arrList As Variant
        Dim strPara As String
        
        On Error GoTo errH
100     strPara = zlDatabase.GetPara(90001, glngSys, , "") '读取URLs 固定读取ZLHIS 系统默认100
        '格式服务器IP&&服务器端口号
102     If strPara = "" Then strPara = "118.31.246.211" & G_STR_SPLIT & "8080" & G_STR_SPLIT & "cqzl123" & G_STR_SPLIT & "6000" & G_STR_SPLIT & "0"
104     arrList = Split(strPara, G_STR_SPLIT)
106     If UBound(arrList) >= 4 Then
            gstrIP = arrList(0)
            gstrPort = arrList(1)
            gstrHOSCODE = arrList(2)     '医院编码
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
146     MsgBox "读取参数失败！" & vbNewLine & "HZYY_GetPara:第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HZYY_SetPara() As String
    HZYY_SetPara = gstrIP & G_STR_SPLIT & gstrPort & G_STR_SPLIT & gstrHOSCODE & G_STR_SPLIT & gstrPortPlus & G_STR_SPLIT & gbytType
End Function

Public Function AdviceCheckWarn_HZYY(ByVal bytFunc As Byte, ByVal lngPatiID As Long, ByVal str挂号单 As String, _
    ByVal lng主页ID As String, Optional ByRef rsOut As ADODB.Recordset, Optional ByVal str医嘱IDs As String) As Boolean
'功能：调用杭州逸曜用药监测系统(BS版)对医嘱进行合理用药审查等相关功能
'
'参数：
'bytFunc=0-医嘱保存;1-药嘱审查;2-删除处方;3-删除医嘱;4-上传有效处方\有效医嘱
'返回值:
'   True -允许保存
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
    Dim rs规格  As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    
    Dim str身高 As String
    Dim str体重 As String
    Dim strDrugID As String
    
    Dim colTmp As Collection
    Dim colRecipe As Collection
    Dim colXML As Collection
    
    Dim strSQL As String
    Dim lng挂号ID As Long
    Dim lngDeptID As Long
    Dim i As Long, k As Long, lngCount As Long
    Dim strDeptName As String
    Dim str诊断编码 As String, str诊断描述 As String
    Dim strTmp As String
    Dim strUrl As String
    Dim strXML As String
    Dim strRet As String
    Dim blnDo As Boolean
    Dim curDate As Date
    Dim strRecipeNo As String
    Dim str药品IDs As String
    Dim byt场合 As Byte
    Dim bytRet  As Byte
    Dim blnIsHaveOut As Boolean
    Dim arrTemp As Variant
    
    On Error GoTo errH
    
    Set colXML = New Collection
    Set rsPati = GetPatiInfo_YF(lngPatiID, str挂号单, lng主页ID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    If str挂号单 <> "" Then
        lng挂号ID = rsPati!就诊Id
        strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                        "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                        "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
                        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, lng挂号ID)
        rsTmp.Filter = "项目名称='身高'"
        If rsTmp.RecordCount <> 0 Then str身高 = NVL(rsTmp!记录内容)
        rsTmp.Filter = "项目名称='体重'"
        If rsTmp.RecordCount <> 0 Then str体重 = NVL(rsTmp!记录内容)
        lngDeptID = CLng(rsPati!挂号科室ID & "")
        strDeptName = rsPati!挂号科室 & ""
    Else
        str身高 = rsPati!身高 & ""
        str体重 = rsPati!体重 & ""
        lngDeptID = CLng(rsPati!入院科室ID)
        strDeptName = rsPati!入院科室 & ""
    End If
    
    'base标签
    With udtBase
        .strHospCode = gstrHOSCODE
        .strPatiID = lngPatiID
        If str挂号单 <> "" Then
            .strSource = IIf(NVL(rsPati!急诊, 0) = 1, "急诊", "门诊")
            .strEventNO = lng挂号ID
        Else
            .strSource = "住院" '门诊|住院|急诊
            .strEventNO = lngPatiID & "_" & lng主页ID
        End If
    End With
    colXML.Add udtBase
    
    Select Case bytFunc
    Case 0, 1, 4
        curDate = zlDatabase.Currentdate
        If str挂号单 <> "" Then
            'opt_patient门诊患者就诊标签
            With udtOptPati
                .strSex = HZYY_GetSex(rsPati!性别 & "")
                .strName = rsPati!姓名 & ""
                .strIDType = "01"                               '居民身份证
                .strIDNO = rsPati!身份证号 & ""
                .strBirthWeight = ""                            '出生体重
                .strBirthDay = Format(rsPati!出生日期 & "", "YYYY-MM-DD HH:mm:ss")            '出生日期
                .strEthnicGroup = rsPati!民族 & ""              '民族
                .strNativePlace = rsPati!籍贯 & ""              '籍贯
                .strRace = ""                                       '人种
                .strMedCardNO = ""                '就诊卡号
                .strEventTime = Format(rsPati!就诊时间 & "", "YYYY-MM-DD HH:mm:ss")                 '就诊时间
                .strDeptID = rsPati!挂号科室ID & ""               '挂号科室ID
                .strDeptName = rsPati!挂号科室 & ""                 '挂号科室名称
                .strPayType = rsPati!医疗付款方式 & ""                    '费用类型
                .strPregnancy = ""                  '是否怀孕
                .strTimeOfPreg = ""                 '孕期
                .strBreastFeeding = ""              '是否哺乳
                .strHeight = str身高                       '身高
                .strWeight = str体重                        '体重
                .strAddress = rsPati!家庭地址 & ""                          '病人地址
                .strPhoneNo = rsPati!手机号 & ""                         '病人电话
                .strDialysis = ""                                   '是否透析
                .strmarital = HZYY_GetMarital(rsPati!婚姻状况 & "")                            '婚姻状况
                .strOccupation = rsPati!职业 & ""                 '职业
                .strSpecialConstitution = "" '特殊体质
                .strVisitType = "门诊"                   '就诊类别  门诊|急诊
                .strPatiCondition = ""
            End With
            colXML.Add udtOptPati
        Else
            '住院病人信息
            With udtIptPati
                .strSex = HZYY_GetSex(rsPati!性别 & "")
                .strName = rsPati!姓名 & ""
                .strIDType = "01"
                .strIDNO = rsPati!身份证号 & ""
                .strBirthWeight = ""                    '出生体重
                .strBirthDay = Format(rsPati!出生日期 & "", "YYYY-MM-DD HH:mm:ss")                        '出生日期
                .strEthnicGroup = rsPati!民族 & ""                    '民族
                .strNativePlace = rsPati!籍贯 & ""                  '籍贯
                .strRace = ""                           '人种
                .strMedCardNO = ""                      '就诊卡号
                .strPayType = rsPati!医疗付款方式 & ""                         '费用类型
                .strPregnancy = ""                      '是否怀孕
                .strTimeOfPreg = ""                     '孕期
                .strBreastFeeding = ""                  '是否哺乳
                .strHeight = str身高                         '身高
                .strWeight = str体重                         '体重
                .strAddress = ""                        '病人地址
                .strPhoneNo = ""                        '病人电话
                .strDialysis = ""                       '是否透析
                .strmarital = ""                        '婚姻状况
                .strOccupation = ""                     '职业
                .strSpecialConstitution = ""            '特殊体质
                .strINDeptId = rsPati!入院科室ID        '入院科室ID
                .strINDeptName = rsPati!入院科室 & ""   '入院科室名称
                .strHospitalTime = Format(rsPati!入院日期 & "", "YYYY-MM-DD HH:mm:ss")                    '入院时间
                .strInWardID = rsPati!入院病区ID & ""                        '入院病区ID
                .strInWardName = rsPati!入院病区 & ""                     '入院病区名称
                .strInWardBedNo = rsPati!入院病床 & ""                    '入院病床号
                .strInConditon = ""                      '入院病情
                .strWeightOfBaby = ""                    '新生儿入院体重
                .strPatientConditon = ""               '患者状态如：普通病人、危重病人
            End With
            colXML.Add udtIptPati
        End If
        'opt_prescriptions处方和处方明细信息标签
        '药品信息
        Select Case glngModel
        Case PM_门诊编辑, PM_门诊医嘱清单, PM_住院编辑, PM_住院医嘱清单
            Set rsAdvice = CreateAdviceRS_HZYY(rsOut, rs规格)
            If glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                Set colRecipe = New Collection
                strRecipeNo = ""
                rsAdvice.Filter = "处方ID>0"
                rsAdvice.Sort = "处方ID"
                For i = 1 To rsAdvice.RecordCount
                    If Val(strRecipeNo) <> Val(rsAdvice!处方ID & "") Then
                        If i <> 1 Then
                            Set udtOptPres.colPresItem = colTmp
                            colRecipe.Add udtOptPres
                        End If
                        strRecipeNo = rsAdvice!处方ID & ""
                        With udtPresInfo
                            .strRecipeId = strRecipeNo
                            .strRecipeNo = strRecipeNo
                            .strRecipeSource = "门诊"                  '处方来源门诊|急诊|其他
                            .strRecipeCategory = ""                     '普通处方，儿科处方,麻醉处方,急诊处方等
                            .strRecipeType = IIf(rsAdvice!诊疗类别 & "" = "5", "西药方", IIf(rsAdvice!诊疗类别 & "" = "6", "中成药方", "草药方")) '处方类型 草药方|中成药方|西药方
                            .strDeptID = rsAdvice!开嘱科室id & ""
                            .strDeptName = rsAdvice!开嘱科室 & ""
                            .strRecipeDocId = rsAdvice!开嘱医生ID & ""                        '开方医生工号
                            .strRecipeDocName = rsAdvice!开嘱医生 & ""                     '开方医生姓名
                            .strRecipeTime = Format(rsAdvice!开嘱时间 & "", "YYYY-MM-DD HH:MM:SS")
                            .strHerbPacketCount = ""                    '饮片帖数 '如果是中药饮片处方，这部分信息需要填写，西药方可为空
                            .strHerbUnitPrice = ""                      '饮片单帖价格
                            .strRecipeFeeTotal = "0"                     '处方金额
                            '处方在his中的状态，0 正常处方，1 退药处方，2 作废处方，只能填入“2|1|0”
                            .strRecipeStatus = "0"
                            .strUrgentFlag = "0"            '紧急标志 用于处方的优先审核，1 是，0 否,不提供则为0，只能填入“1|0”
                        End With
                        udtOptPres.udtOptPresInfo = udtPresInfo
                        Set colTmp = New Collection
                    End If
                    With udtPresItem
                        .strRecipeId = strRecipeNo                   '处方ID
                        .strRecipeItemId = rsAdvice!医嘱id & ""         '处方明细编号
                        .strGroupNO = rsAdvice!相关ID & ""
                        .strDrugID = rsAdvice!药品ID & ""
                        .strDrugName = rsAdvice!药品名称 & ""
                        .strDrugUsingAim = rsAdvice!用药目的 & ""          '用药目的
                        .strManufacturerID = ""                 '生产厂家ID
                        .strManufacturerName = ""               '生产厂家名称
                        .strDrugdose = rsAdvice!单次用量 & rsAdvice!单量单位                       '单次使用的药品剂量及单位，如：0.5g 、200ml等
                        .strDrugadminRouteName = rsAdvice!用法 & ""             '药品使用途径，如：静脉推注等
                        .strDrugUsingFreq = rsAdvice!频率 & ""                   '药品给药频次，如：qd、bid、每天2次等
                        .strDuration = rsAdvice!天数 & ""                   '疗程 天数
                        rs规格.Filter = "药品ID=" & rsAdvice!药品ID
                        If Not rs规格.EOF Then
                            .strPreparation = rs规格!药品剂型 & ""                    '药品剂型名称
                            .strSpecification = rs规格!规格 & ""                   '药品规格描述
                            .strUnitPrice = rs规格!现价 & ""                         '单价
                            .strCountUnit = rs规格!门诊包装 & ""                      '包装规格数量 如果一盒药品12片，发药单位为盒时，包装规格数量为12，发药单位为片时，包装规格数量为1
                            .strPackUnit = rs规格!门诊单位 & ""              '包装规格单位
                            .strFeeTotal = Val(rsAdvice!总量 & "") * Val(rs规格!现价 & "")       '药品的单价*发药数量，计量单位为元
                        Else
                            .strPreparation = ""                   '药品剂型名称
                            .strSpecification = ""                  '药品规格描述
                            .strUnitPrice = ""                      '单价
                        End If
                        .strDespensingNum = FormatEx(rsAdvice!总量 & "", 5)                     '发药数量    发药数量，如2盒、10片等
                        .strSkinTestFlag = "0"                   '皮试标志    1 皮试，0 非皮试，只能填入“1|0”
                        .strdrugReturnFlag = "0"                '是否退药标志    1 是，0 否,不提供则为0，只能填入“1|0”
                        .strOuvasFlag = Val(rsAdvice!输液 & "")                    '门诊静配标志    1 是，0 否,不提供则为0，只能填入“1|0”
                        .strDrippingSpeed = rsAdvice!滴速 & ""             '滴速 静脉输液药品滴注时间和速度的描述，如：1小时、20滴/分钟，需同时传入数字和单位，单位为小时或滴/分钟
                        If Val(rsAdvice!标志 & "") = 1 And Val(udtOptPres.udtOptPresInfo.strUrgentFlag) <> 1 Then
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
                If colRecipe.Count = 0 Then AdviceCheckWarn_HZYY = True: Exit Function     '医嘱下达界面没有下达药品时允许保存
                colXML.Add colRecipe, "_" & colXML.Count + 1
            ElseIf PM_住院编辑 = glngModel Or glngModel = PM_住院医嘱清单 Then
                'orders医嘱信息标签
                Set udtOrder.colMedical = New Collection
                rsAdvice.Filter = "诊疗类别='5' OR 诊疗类别='6'"
                For i = 1 To rsAdvice.RecordCount
                    If PM_住院编辑 And rsAdvice!离院带药 = 1 Then blnIsHaveOut = True
                    With udtOrderItem
                        .strOrderId = rsAdvice!医嘱id & ""                      '医嘱id
                        .strOrderTime = Format(rsAdvice!开嘱时间 & "", "YYYY-MM-DD HH:mm:ss")      '医嘱时间
                        .strOrderDeptID = rsAdvice!开嘱科室id & ""                    '医嘱科室代码
                        .strOrderDeptName = rsAdvice!开嘱科室 & ""                 '医嘱科室名称
                        .strDocGroup = ""                        '医嘱医疗组名称
                        .strOrderDocID = rsAdvice!开嘱医生ID & ""                    '医嘱医生id
                        .strOrderDocName = rsAdvice!开嘱医生 & ""                  '医嘱医生姓名
                        .strOrderDocTitle = rsAdvice!专业技术职务 & ""                '医嘱医生职称
                        .strOrderType = rsAdvice!医嘱期效 & ""
                        .strDrugPurpose = ""                      '药嘱目的
                        .strGroupNO = rsAdvice!相关ID & ""                        '组号
                        .strDrugID = rsAdvice!药品ID & ""                          '药品ID
                        .strDrugName = rsAdvice!药品名称 & ""                       '药品通用名
                        rs规格.Filter = "药品ID=" & rsAdvice!药品ID
                        If Not rs规格.EOF Then
                            .strPreparation = rs规格!药品剂型 & ""                    '药品剂型名称
                            .strSpecifications = rs规格!规格 & ""                    '药品规格描述
                            .strUnitPrice = rs规格!现价 & ""                         '单价
                            .strManufacturerID = rs规格!厂家名称 & ""                    '生产厂家id
                            .strManufacturerName = rs规格!厂家编码 & ""                 '生产厂家名称
                            .strCountUnit = rs规格!住院包装 & ""                    '包装规格数量
                            .strPackUnit = rs规格!住院单位 & ""                        '包装规格单位
                            .strFeeTotal = Val(rsAdvice!总量 & "") * Val(rs规格!现价 & "")                 '总价
                        Else
                            .strPreparation = ""                   '药品剂型名称
                            .strSpecifications = ""                  '药品规格描述
                            .strUnitPrice = ""                      '单价
                            .strManufacturerID = ""                      '生产厂家id
                            .strManufacturerName = ""                   '生产厂家名称
                        End If
                        .strDrugdose = rsAdvice!单次用量 & rsAdvice!单量单位                         '每次给药剂量
                        .strDrugadminRouteName = rsAdvice!用法 & ""           '给药途径
                        .strDrugUsingFreq = rsAdvice!频率 & ""                  '给药频率
                        .strDrugUsingTimePoint = ""            '给药时机
                        .strDrugUsingAim = rsAdvice!用药目的 & ""                   '给药目的
                        .strDrugUsingArea = ""                  '给药部位
                        .strDrugSource = ""                      '药品来源
                        .strDuration = ""                         '疗程
                        .strDespensingNum = FormatEx(rsAdvice!总量 & "", 5)                     '发药数量
                        .strCheckTime = ""                       '复核时间
                        .strCheckNurseID = ""                   '复核护士id
                        .strCheckNurseName = ""                '复核护士姓名
                        .strOrderValidTime = Format(rsAdvice!开始时间 & "", "YYYY-MM-DD HH:mm:ss")                 '医嘱生效时间
                        .strOrderInvalidTime = Format(rsAdvice!结束时间 & "", "YYYY-MM-DD HH:mm:ss")               '医嘱失效时间
                        .strSpecialPrompt = ""                   '特殊要求
                        .strSkinTestTime = ""                   '皮试时间
                        .strSkinTestFlag = "0"                   '皮试标志
                        .strSkinTestResult = ""                 '皮试结果
                        .strdrugReturnFlag = ""                 '是否退药标志
                        .strStopFlag = ""                        '是否停药标志
                        .strPivasFlag = Val(rsAdvice!输液 & "")                      '住院静配标志
                        .strUrgentFlag = IIf(rsAdvice!标志 & "" = "1", "1", "0")                       '紧急标志
                        .strDrippingSpeed = rsAdvice!滴速 & ""                  '滴速
                        .strLimitTime = ""                       '限用时间
                        .strTherapeuticRegimen = ""              '用药方案
                        .strExeDeptID = ""                      '医嘱执行科室id
                        .strExeDeptName = ""                   '医嘱执行科室名称
                        .strDispensingWindow = ""                '发药窗口号
                        .strDrugstoreArea = ""                  '商品货架号
                    End With
                    
                    udtOrder.colMedical.Add udtOrderItem
                    rsAdvice.MoveNext
                Next
                
                rsAdvice.Filter = "诊疗类别='7'"
                strRecipeNo = ""
                Set udtOrder.colHerbMedical = New Collection
                For i = 1 To rsAdvice.RecordCount
                    If rsAdvice!离院带药 = 1 And glngModel = PM_住院编辑 Then blnIsHaveOut = True
                    If rsAdvice!相关ID <> strRecipeNo Then
                        If i <> 1 Then
                            udtOrder.colHerbMedical.Add udtHerbOrder
                        End If
                        strRecipeNo = rsAdvice!相关ID & ""
                        With udtHerbInfo
                            .strOrderId = rsAdvice!相关ID & ""                           '医嘱id
                            .strOrderTime = Format(rsAdvice!开嘱时间 & "", "YYYY-MM-DD HH:mm:ss")                          '医嘱时间
                            .strOrderDeptID = rsAdvice!开嘱科室id & ""                         '医嘱科室代码
                            .strOrderDeptName = rsAdvice!开嘱科室 & ""                      '医嘱科室名称
                            .strDocGroup = ""                           '医嘱医疗组名称
                            .strOrderDocID = rsAdvice!开嘱医生ID & ""                          '医嘱医生id
                            .strOrderDocName = rsAdvice!开嘱医生 & ""                        '医嘱医生姓名
                            .strOrderDocTitle = rsAdvice!专业技术职务 & ""                       '医嘱医生职称
                            .strOrderType = rsAdvice!医嘱期效 & ""                          '医嘱类型
                            .strHerbUnitPrice = ""                      '饮片单帖价格
                            .strHerbPacketCount = ""                    '饮片帖数
                            .strIsCream = ""                            '膏方
                            .strCheckTime = ""                          '复核时间
                            .strCheckNurseID = ""                       '复核护士id
                            .strCheckNurseName = ""                     '复核护士姓名
                            .strOrderValidTime = Format(rsAdvice!开始时间 & "", "YYYY-MM-DD HH:mm:ss")                     '医嘱生效时间
                            .strOrderInvalidTime = Format(rsAdvice!结束时间 & "", "YYYY-MM-DD HH:mm:ss")                       '医嘱失效时间
                            .strdrugReturnFlag = ""                     '是否退药标志
                            .strStopFlag = ""                           '是否停药标志
                            .strUrgentFlag = IIf(rsAdvice!标志 & "" = "1", "1", "0")                        '紧急标志
                            .strExeDeptID = ""                          '医嘱执行科室id
                            .strExeDeptName = ""                        '医嘱执行科室名称
                        End With
                        udtHerbOrder.udtHerbInfo = udtHerbInfo
                        Set udtHerbOrder.colItemHerb = New Collection
                    End If
                    With udtHerbItem
                        .strOrderId = rsAdvice!医嘱id & ""                   '医嘱id
                        .strOrderitemID = rsAdvice!医嘱id & ""               '医嘱明细
                        .strGroupNO = rsAdvice!相关ID & ""                   '组号
                        .strDrugID = rsAdvice!药品ID & ""                    '药品ID
                        .strDrugName = rsAdvice!药品名称 & ""                       '药品通用名
                        .strDrugdose = rsAdvice!单次用量 & rsAdvice!单量单位 & ""   '每次给药剂量
                        .strDrugadminRouteName = rsAdvice!用法 & ""                 '给药途径
                        .strDrugUsingFreq = rsAdvice!频率 & ""                      '给药频率
                        rs规格.Filter = "药品ID=" & rsAdvice!药品ID
                        If Not rs规格.EOF Then
                            .strPreparation = rs规格!药品剂型 & ""                      '药品剂型名称
                            .strSpecifications = rs规格!规格 & ""                       '药品规格描述
                            .strUnitPrice = rs规格!现价 & ""                            '单价
                            .strManufacturerID = rs规格!厂家名称 & ""                   '生产厂家id
                            .strManufacturerName = rs规格!厂家编码 & ""
                        Else
                            .strPreparation = ""                    '药品剂型名称
                            .strSpecifications = ""                 '药品规格描述
                            .strUnitPrice = ""                      '单价
                            .strManufacturerID = ""                 '生产厂家id
                            .strManufacturerName = ""               '生产厂家名称
                        End If
                        .strDespensingNum = ""                                                           '发药数量
                        .strFeeTotal = ""                                                                '总价
                        .strSpecialPrompt = ""                                                           '特殊要求
                    End With
                    udtHerbOrder.colItemHerb.Add udtHerbItem
                    If i = rsAdvice.RecordCount Then
                        udtOrder.colHerbMedical.Add udtHerbOrder
                    End If
                    rsAdvice.MoveNext
                Next
                
                Set udtOrder.colNonMedical = New Collection
                
                If udtOrder.colMedical.Count = 0 And udtOrder.colHerbMedical.Count = 0 Then AdviceCheckWarn_HZYY = True: Exit Function     '医嘱下达界面没有下达药品时允许保存
                colXML.Add udtOrder, "_" & colXML.Count + 1
            End If
        Case PM_部门发药, PM_处方发药, PM_PIVA管理
            Set rsTmp = CreateAdviceRS_HZYY(rsOut, rs规格, str医嘱IDs)
            byt场合 = 1
        End Select
        If str挂号单 <> "" And PM_门诊编辑 = glngModel Then
            'diagnoses诊断信息标签
            Set colTmp = New Collection
            If Not gobjDiags Is Nothing Then
                For i = 1 To gobjDiags.Count
                    With udtDiag
                        If gobjDiags.Item(i).str诊断描述 <> "" Then
                            If gobjDiags.Item(i).str疾病编码 <> "" Then
                                .strDiagID = gobjDiags.Item(i).str疾病ID
                                .strDiagDeptID = lngDeptID
                                .strDiagDeptName = strDeptName
                                .strDiagDate = gobjDiags.Item(i).str诊断时间
                                .strDiagName = gobjDiags.Item(i).str诊断描述
                                .strDiagCode = gobjDiags.Item(i).str诊断编码
                            Else
                                .strDiagID = gobjDiags.Item(i).str疾病ID
                                .strDiagDeptID = lngDeptID
                                .strDiagDeptName = strDeptName
                                .strDiagDate = gobjDiags.Item(i).str诊断时间
                                .strDiagName = gobjDiags.Item(i).str诊断描述
                                .strDiagCode = gobjDiags.Item(i).str诊断编码
                            End If
                        End If
                    End With
                    colTmp.Add udtDiag, "_" & colTmp.Count + 1
                Next
            End If
            colXML.Add colTmp
        Else
            Set colTmp = New Collection
            Set rsTmp = Get病人诊断记录(lngPatiID, IIf(str挂号单 <> "", lng挂号ID, lng主页ID), IIf(str挂号单 <> "", "1,11", "2,12"))
            For i = 1 To rsTmp.RecordCount
                With udtDiag
                    .strDiagID = rsTmp!id
                    .strDiagDeptID = lngDeptID
                    .strDiagDeptName = strDeptName
                    .strDiagDate = Format(rsTmp!记录日期 & "", "YYYY-MM-DD HH:MM:SS")
                    .strDiagName = rsTmp!名称 & ""
                    .strDiagCode = rsTmp!编码 & ""
                End With
                colTmp.Add udtDiag, "_" & colTmp.Count + 1
                rsTmp.MoveNext
            Next
            colXML.Add colTmp
        End If
        '过敏记录 任意取一个药品ID传人
        'allergies过敏信息标签
        Set rsTmp = Get病人过敏记录(lngPatiID, IIf(str挂号单 <> "", 0, lng主页ID))
        Set colTmp = New Collection
        For i = 1 To rsTmp.RecordCount
            strDrugID = ""
            If rsTmp!药物ID & "" <> "" Then
                strSQL = "select 药品ID from 药品规格 where 药名id=[1] and rownum <2"
                Set rs规格 = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, rsTmp!药物ID)
                If Not rs规格.EOF Then strDrugID = rs规格!药品ID & ""
            End If
            With udtAllergy
                .strAllergyID = strDrugID
                .strAllergyDrug = rsTmp!药物名 & ""
                .strAnaphylaxis = rsTmp!过敏反应 & ""
                .strRecordTime = rsTmp!记录时间 & ""
            End With
            colTmp.Add udtAllergy, "_" & colTmp.Count + 1
            rsTmp.MoveNext
        Next
        colXML.Add colTmp
        '传入病人手术记录operations手术信息标签
        Set colTmp = New Collection
        Set rsTmp = GetPatiOperation(lngPatiID, lng主页ID, str挂号单)
        For i = 1 To rsTmp.RecordCount
            With udtOpera
                .strOperationID = rsTmp!id & ""
                .strOperationCode = rsTmp!编码 & ""
                .strOperationName = rsTmp!名称 & ""
                .strOperationStartTime = Format(rsTmp!手术时间 & "", "YYYY-MM-DD HH:MM:SS")
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
        strXML = HZYY_GetRootXML(colXML, IIf(str挂号单 <> "", 0, 1))
        strXML = Replace(strXML, "<![CDATA[]]>", "")
        strXML = "charset=utf-8&post_type=1&xml=" & strXML  '医生开药
        WriteLog "" & glngModel, "HttpPost", "传入值:" & strXML
        strRet = HttpPost(strUrl, strXML, ResponseText, "application/x-www-form-urlencoded; charset=utf-8")
        strRet = Replace(strRet, "<![CDATA[]]>", "")
        WriteLog "" & glngModel, "HttpPost", "返回值:" & strRet
        
        If bytFunc = 4 Then AdviceCheckWarn_HZYY = True: Exit Function
        
        If strRet = "解析转换干预失败" Then
            MsgBox "合理用药监测:" & strRet, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        ElseIf strRet = "" Then
            MsgBox "当前启用了合理用药系统，但由于合理用药接口(服务器返回空)未调用成功，请与系统管理员联系。", vbInformation + vbOKOnly, gstrSysName
        Else
            Set rsTmp = HZYY_ParseXML(strRet, strTmp, IIf(str挂号单 <> "", 0, 1))
            If Not rsTmp Is Nothing Then
                If rsTmp.RecordCount > 0 Then
                    strTmp = Replace(strTmp, "姓名:XX", "姓名:" & IIf(str挂号单 <> "", udtOptPati.strName, udtIptPati.strName))
                    frmPassResult.ShowMe gfrmMain, rsTmp, strTmp, bytRet, bytFunc, blnIsHaveOut
                End If
            End If
            If bytRet = 1 Then
                Exit Function
            End If
        End If
    Case 2, 3   '删除处方及医嘱
        If bytFunc = 2 Then
            str医嘱IDs = Replace(str医嘱IDs, "【西药】", "")
            str医嘱IDs = Replace(str医嘱IDs, "【中药】", "")
            If str医嘱IDs <> "" Then
                arrTemp = Split(str医嘱IDs, ",")
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
            '医嘱ID1,医嘱ID2|组ID
            If str医嘱IDs <> "" Then
                If InStr(str医嘱IDs, "【西药】") > 0 Then
                    str医嘱IDs = Replace(str医嘱IDs, "【西药】", "")
                    strTmp = Split(str医嘱IDs, "|")(1)
                    arrTemp = Split(Split(str医嘱IDs, "|")(0), ",")
                    Set udtOrder.colMedical = New Collection
                    For i = LBound(arrTemp) To UBound(arrTemp)
                        udtOrderItem.strOrderId = arrTemp(i)
                        udtOrderItem.strGroupNO = strTmp
                        udtOrder.colMedical.Add udtOrderItem
                    Next
                ElseIf InStr(str医嘱IDs, "【中药】") > 0 Then
                    str医嘱IDs = Replace(str医嘱IDs, "【中药】", "")
                    strTmp = Split(str医嘱IDs, "|")(1)
                    arrTemp = Split(Split(str医嘱IDs, "|")(0), ",")
                    Set udtOrder.colHerbMedical = New Collection
                    For i = LBound(arrTemp) To UBound(arrTemp)
                        udtHerbOrder.udtHerbInfo.strOrderId = arrTemp(i)
                        udtOrder.colHerbMedical.Add udtHerbOrder
                    Next
                End If
                colXML.Add udtOrder
            End If
        End If
        '删除医嘱、处方v
        strUrl = "http://" & gstrIP & ":" & gstrPortPlus & "/v4/invalid"
        strXML = HZYY_GetRootXML(colXML, IIf(str挂号单 <> "", 0, 1), 1)
        strXML = Replace(strXML, "<![CDATA[]]>", "")
        strXML = "charset=utf-8&post_type=1&xml=" & strXML  '医生开药
        WriteLog "" & glngModel, "HttpPost", "传入值:" & strXML
        strRet = HttpPost(strUrl, strXML, ResponseText, "application/x-www-form-urlencoded; charset=utf-8")
        strRet = Replace(strRet, "<![CDATA[]]>", "")
        WriteLog "" & glngModel, "HttpPost", "返回值:" & strRet
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
    Optional ByVal str医嘱IDs As String, Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal str挂号单 As String) As ADODB.Recordset
'功能;构造医嘱记录集
    Dim i As Long, k As Long, lngCount As Long, lngPos As Long
    Dim blnDo As Boolean, blnIsHaveOut As Boolean
    Dim str药品 As String, str医嘱ID As String, str相关ID As String
    Dim str开嘱时间 As String
    Dim str期效 As String, str单量 As String, str单量单位 As String, str频率 As String
    Dim str给药途径 As String, str频率编码 As String, str用法 As String, str用法ID As String, str开始时间 As String, str结束时间 As String
    Dim str开嘱科室Tag As String, str开嘱科室ID As String, str诊疗项目IDs As String, str药品ID As String
    Dim str开嘱医生Tag As String, str开嘱医生 As String, str用药目的 As String
    Dim str总量 As String, str总量单位 As String, strType As String
    Dim str诊疗ID, str收费细目ID As String, str滴速 As String, str输液 As String
    Dim str处方号   As String
    
    Dim rsAdvice As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rs频率 As ADODB.Recordset
    Dim rs开嘱医生 As ADODB.Recordset
    Dim rs开嘱科室  As ADODB.Recordset
    Dim rs药品 As ADODB.Recordset
    Dim rs职务 As ADODB.Recordset
    
    Dim curDate As Date
    
    On Error GoTo errH
    
    curDate = zlDatabase.Currentdate
    Set rsAdvice = InitAdviceRS(FUN_医嘱信息_HZYY)

    Select Case glngModel
    Case PM_门诊编辑, PM_住院编辑
        '启用了禁忌药品说明参数;场合为门诊编辑\住院编辑;审查功能
        If (glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑) And gbytReason = 1 Then
            Set rsOut = InitAdviceRS(FUN_输出内容)
        End If
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_门诊编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                            And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-DD") = Format(curDate, "yyyy-MM-DD")
                ElseIf glngModel = PM_住院编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                    If blnDo Then
                        blnDo = (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0 _
                                Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                                And .TextMatrix(i, gobjCOL.intCOL状态) <> "4")
                    End If
                End If

                If blnDo Then
                    str诊疗ID = .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                    If InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) = 0 Then
                        If InStr("," & str诊疗项目IDs & ",", "," & str诊疗ID & ",") = 0 Then
                            str诊疗项目IDs = str诊疗项目IDs & "," & str诊疗ID
                        End If
                    End If
                    str医嘱ID = CStr(.RowData(i))

                    '取药品名称
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                        str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                    Else
                        str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                    End If
                    '取药品给药途径
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str用法 = ""    '一并给药不重复取

                    If str用法 = "" Then
                        str滴速 = "": str输液 = "0"
                        If glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Then
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                            If k <> -1 Then
                                If .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                                    str用法 = .TextMatrix(k, gobjCOL.intCOL用法)
                                Else
                                    str用法 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                    If InStr(.TextMatrix(k, gobjCOL.intcol医嘱嘱托), "滴/分钟") > 0 Or InStr(.TextMatrix(k, gobjCOL.intcol医嘱嘱托), "毫升/小时") > 0 Then
                                        str滴速 = .TextMatrix(k, gobjCOL.intcol医嘱嘱托)
                                    End If
                                    If Val(.TextMatrix(k, gobjCOL.intCol执行分类)) = 1 Then
                                        str输液 = "1"
                                    Else
                                        str输液 = "0"
                                    End If
                                End If
                                str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                            End If
                        End If
                    End If
                    '开嘱科室名称
                    str开嘱科室ID = .TextMatrix(i, gobjCOL.intCOL开嘱科室ID)
                    If InStr("," & str开嘱科室Tag & ",", "," & str开嘱科室ID & ",") = 0 Then
                        str开嘱科室Tag = str开嘱科室Tag & "," & str开嘱科室ID
                    End If

                    '开嘱医生
                    str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                    If InStr("," & str开嘱医生Tag & ",", "," & str开嘱医生 & ",") = 0 Then
                        str开嘱医生Tag = str开嘱医生Tag & "," & str开嘱医生
                    End If

                    str开始时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:MM:SS")
'
                    str开嘱时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd HH:mm:ss")         '处方时间（YYYY-MM-DD HH:mm:SS）
                    '单量，单量单位
                    str单量 = .TextMatrix(i, gobjCOL.intCOL单量)
                    str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                    str总量 = .TextMatrix(i, gobjCOL.intCOL总量)
                    str总量单位 = .TextMatrix(i, gobjCOL.intcol总量单位)

                    str药品ID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)

                    If glngModel = PM_门诊编辑 Then
                        str结束时间 = ""
                        str期效 = "2" '2-临时医嘱
                        strType = .TextMatrix(i, gobjCOL.intCOL状态)
                        If strType = "4" Then
                            strType = "2"       '作废处方
                        ElseIf strType = "1" Then
                            strType = "0"       '正常处方
                        End If
                        str处方号 = .TextMatrix(i, gobjCOL.intCol处方号)
                    ElseIf glngModel = PM_住院编辑 Then
                        str期效 = IIf(.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱", 1, 2)
                        str结束时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL终止时间), "yyyy-MM-dd HH:MM:SS")
                        '判断是否是院外执行的药品
                        If Val(.TextMatrix(i, gobjCOL.intCOL执行性质)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))), gobjCOL.intCOL执行性质)) = 5 Then
                            blnIsHaveOut = True
                            str期效 = "3"
                        End If
                        str处方号 = ""
                    End If

                    '禁忌说明
                    If Not rsOut Is Nothing Then
                        If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                        '西药,中成药
                            rsOut.AddNew
                            rsOut!医嘱id = CLng(str医嘱ID)
                            rsOut!禁忌药品说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                            rsOut!状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                            rsOut!药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                            rsOut.Update
                        ElseIf Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                        '中药配方  禁忌说明保存在用药服法上
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                            If k <> -1 Then
                                rsOut.AddNew
                                rsOut!医嘱id = CLng(.RowData(k) & "")
                                rsOut!禁忌药品说明 = .TextMatrix(k, gobjCOL.intCol禁忌药品说明)
                                rsOut!状态 = .TextMatrix(k, gobjCOL.intCOL状态)
                                rsOut!药品名称 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                rsOut.Update
                            End If
                        End If
                    End If
                    str用药目的 = .TextMatrix(i, gobjCOL.intcol用药目的)
                    If str用药目的 = "1" Then
                        str用药目的 = "预防用药"
                    ElseIf str用药目的 = "2" Then
                        str用药目的 = "治疗用药"
                    Else
                        str用药目的 = ""
                    End If
                    '----------------------------------------------------------
                    rsAdvice.AddNew
                    rsAdvice!处方ID = Val(str处方号)
                    rsAdvice!医嘱id = str医嘱ID
                    rsAdvice!相关ID = .TextMatrix(i, gobjCOL.intCOL相关ID)
                    rsAdvice!医嘱期效 = str期效
                    rsAdvice!医嘱序号 = .TextMatrix(i, gobjCOL.intCOL序号)
                    rsAdvice!开嘱科室id = str开嘱科室ID
                    rsAdvice!开嘱医生 = str开嘱医生
                    rsAdvice!诊疗项目ID = str诊疗ID
                    rsAdvice!药品ID = Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID))
                    rsAdvice!药品名称 = str药品
                    rsAdvice!医嘱状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                    rsAdvice!单次用量 = str单量
                    rsAdvice!单量单位 = str单量单位
                    rsAdvice!频率 = .TextMatrix(i, gobjCOL.intCOL频率)
                    rsAdvice!用法 = str用法
                    rsAdvice!用法ID = str给药途径
                    rsAdvice!开嘱时间 = str开嘱时间
                    rsAdvice!开始时间 = str开始时间
                    rsAdvice!结束时间 = str结束时间
                    rsAdvice!总量 = str总量
                    rsAdvice!总量单位 = str总量单位
                    rsAdvice!天数 = .TextMatrix(i, gobjCOL.intCOL天数)
                    rsAdvice!医生嘱托 = .TextMatrix(i, gobjCOL.intcol医嘱嘱托)
                    rsAdvice!用药目的 = str用药目的
                    rsAdvice!用药理由 = .TextMatrix(i, gobjCOL.intcol用药理由)
                    rsAdvice!诊疗类别 = .TextMatrix(i, gobjCOL.intCOL诊疗类别)
                    rsAdvice!标志 = .TextMatrix(i, gobjCOL.intCol标志)
                    rsAdvice!滴速 = str滴速
                    rsAdvice!输液 = str输液
                    rsAdvice!离院带药 = IIf(blnIsHaveOut, 1, 0)
                    rsAdvice.Update
                    '----------------------------------------------------------------------------
                End If
            Next
        End With
    Case PM_门诊医嘱清单, PM_住院医嘱清单
        Set rsTmp = GetAdviceInfo_YF(gobjPati.lng病人ID, gobjPati.lng主页ID, gobjPati.str挂号单, , 1)
        With rsTmp
            If rsTmp.RecordCount = 0 Then Set CreateAdviceRS_HZYY = rsAdvice: Exit Function
            For i = 1 To .RecordCount
                If glngModel = PM_门诊医嘱清单 Then
                    blnDo = InStr(",5,6,7,", "," & !诊疗类别 & ",") > 0 And Val(!收费细目id & "") <> 0 And Format(!开嘱时间 & "", "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                ElseIf glngModel = PM_住院医嘱清单 Then
                    blnDo = InStr(",5,6,7,", "," & !诊疗类别 & ",") > 0 And Not InStr(",4,8,9,", "," & !医嘱状态 & ",") > 0
                End If
                If blnDo Then

                    If InStr(",5,6,7,", "," & !诊疗类别 & ",") > 0 And Not InStr(",4,8,9,", "," & !医嘱状态 & ",") > 0 And Val(!收费细目id & "") = 0 Then
                        If InStr("," & str诊疗项目IDs & ",", "," & !诊疗项目ID & ",") = 0 Then
                            str诊疗项目IDs = str诊疗项目IDs & "," & !诊疗项目ID
                        End If
                    End If
                    '开嘱医生
                    str开嘱医生 = !开嘱医生 & ""
                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                    If InStr("," & str开嘱医生Tag & ",", "," & str开嘱医生 & ",") = 0 Then
                        str开嘱医生Tag = str开嘱医生Tag & "," & str开嘱医生
                    End If

                    If gobjPati.str挂号单 <> "" Then
                        str总量单位 = !门诊单位 & ""
                    Else
                        str总量单位 = !住院单位 & ""
                    End If

                    If InStr(";" & str频率 & ";", ";" & !频率 & "," & IIf(!诊疗类别 & "" = "7", 2, 1) & ";") = 0 Then
                        str频率 = str频率 & ";" & !频率 & "," & IIf(!诊疗类别 = "7", 2, 1)
                    End If

                    rsAdvice.AddNew
                    rsAdvice!处方ID = Val(!处方ID & "")
                    rsAdvice!医嘱id = !医嘱id & ""
                    rsAdvice!相关ID = !相关ID & ""
                    rsAdvice!医嘱期效 = IIf(Val(!医嘱期效 & "") = 0, 1, 2) 'HZYY 1长嘱;2临嘱;3离院带药
                    If !A执行性质 & "" <> "5" And !B执行性质 & "" = "5" Then
                        rsAdvice!医嘱期效 = "3"
                    End If
                    rsAdvice!医嘱序号 = lngCount + 1
                    rsAdvice!开嘱科室id = !开嘱科室id & ""
                    rsAdvice!开嘱科室 = !开嘱科室 & ""
                    rsAdvice!开嘱医生 = str开嘱医生
                    rsAdvice!诊疗项目ID = !诊疗项目ID & ""
                    rsAdvice!药品ID = !收费细目id & ""
                    rsAdvice!药品名称 = !药品名称 & ""
                    rsAdvice!医嘱状态 = !医嘱状态 & ""
                    rsAdvice!单次用量 = !单次用量 & ""
                    rsAdvice!单量单位 = !单量单位 & ""
                    rsAdvice!频率 = !频率 & ""
                    rsAdvice!用法 = !用法 & ""
                    rsAdvice!用法ID = !用法ID & ""
                    rsAdvice!开嘱时间 = !开嘱时间 & ""
                    rsAdvice!开始时间 = !开始时间 & ""
                    rsAdvice!结束时间 = !结束时间 & ""
                    rsAdvice!总量 = !总量 & ""
                    rsAdvice!总量单位 = str总量单位
                    rsAdvice!天数 = !天数 & ""
                    rsAdvice!医生嘱托 = !医生嘱托 & ""
                    
                    If Val(!用药目的 & "") = 1 Then
                        rsAdvice!用药目的 = "预防用药"
                    ElseIf Val(!用药目的 & "") = 2 Then
                        rsAdvice!用药目的 = "治疗用药"
                    End If
                    
                    rsAdvice!用药理由 = !用药理由 & ""
                    rsAdvice!诊疗类别 = !诊疗类别 & ""
                    rsAdvice!规格 = !规格 & ""
                    rsAdvice!标志 = !标志 & ""
                    If !类别 & "_" & !操作类型 & "_" & !执行分类 = "E_2_1" Then
                        rsAdvice!输液 = "1"
                    Else
                        rsAdvice!输液 = "0"
                    End If
                    rsAdvice.Update
                End If
                .MoveNext
            Next
        End With
    Case PM_PIVA管理, PM_部门发药, PM_处方发药
        Set rsTmp = GetAdviceInfo_YF(lng病人ID, lng主页ID, str挂号单)
        With rsTmp
            If rsTmp.RecordCount = 0 Then Set CreateAdviceRS_HZYY = rsAdvice: Exit Function
            For i = 1 To .RecordCount

                If Val(!收费细目id & "") = 0 Then
                    If InStr("," & str诊疗项目IDs & ",", "," & !诊疗项目ID & ",") = 0 Then
                        str诊疗项目IDs = str诊疗项目IDs & "," & !诊疗项目ID
                    End If
                End If
                '开嘱医生
                str开嘱医生 = !开嘱医生 & ""
                If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                If InStr("," & str开嘱医生Tag & ",", "," & str开嘱医生 & ",") = 0 Then
                    str开嘱医生Tag = str开嘱医生Tag & "," & str开嘱医生
                End If

                If str挂号单 <> "" Then
                    str总量单位 = !门诊单位 & ""
                Else
                    str总量单位 = !住院单位 & ""
                End If

                If InStr(";" & str频率 & ";", ";" & !频率 & "," & IIf(!诊疗类别 & "" = "7", 2, 1) & ";") = 0 Then
                    str频率 = str频率 & ";" & !频率 & "," & IIf(!诊疗类别 = "7", 2, 1)
                End If

                rsAdvice.AddNew
                rsAdvice!医嘱id = !医嘱id & ""
                rsAdvice!相关ID = !相关ID & ""
                rsAdvice!医嘱期效 = !医嘱期效 & ""
                rsAdvice!医嘱序号 = lngCount + 1
                rsAdvice!开嘱科室id = !开嘱科室id & ""
                rsAdvice!开嘱科室 = !开嘱科室 & ""
                rsAdvice!开嘱医生 = str开嘱医生
                rsAdvice!诊疗项目ID = !诊疗项目ID & ""
                rsAdvice!药品ID = !收费细目id & ""
                rsAdvice!药品名称 = !药品名称 & ""
                rsAdvice!医嘱状态 = !医嘱状态 & ""
                rsAdvice!单次用量 = !单次用量 & ""
                rsAdvice!单量单位 = !单量单位 & ""
                rsAdvice!频率 = !频率 & ""
                rsAdvice!用法 = !用法 & ""
                rsAdvice!用法ID = !用法ID & ""
                rsAdvice!开嘱时间 = !开嘱时间 & ""
                rsAdvice!开始时间 = !开始时间 & ""
                rsAdvice!结束时间 = !结束时间 & ""
                rsAdvice!总量 = !总量 & ""
                rsAdvice!总量单位 = str总量单位
                rsAdvice!天数 = !天数 & ""
                rsAdvice!医生嘱托 = !医生嘱托 & ""
                rsAdvice!用药目的 = !用药目的 & ""
                rsAdvice!用药理由 = !用药理由 & ""
                rsAdvice!诊疗类别 = !诊疗类别 & ""
                rsAdvice!规格 = !规格 & ""
                rsAdvice!标志 = !标志 & ""
                rsAdvice.Update

                .MoveNext
            Next
        End With
    End Select

    '附加数据提取
    If rsAdvice.RecordCount > 0 Then

        rsAdvice.MoveFirst
        Select Case glngModel

        Case PM_门诊编辑, PM_门诊医嘱清单, PM_住院编辑, PM_住院医嘱清单, PM_PIVA管理, PM_部门发药, PM_处方发药
            If str诊疗项目IDs <> "" Then
                str诊疗项目IDs = Mid(str诊疗项目IDs, 2)
                Set rs药品 = GetRS("药品规格", "药名id,药品id", str诊疗项目IDs, "药名id")
            End If
            If str频率 <> "" Then Set rs频率 = GetRS("诊疗频率项目", "编码, 名称, 适用范围", str频率, "名称, 适用范围", 1, 2)
            If str开嘱科室Tag <> "" Then Set rs开嘱科室 = GetRS("部门表", "ID,名称", str开嘱科室Tag)
            If str开嘱医生Tag <> "" Then Set rs开嘱医生 = GetRS("人员表 A,专业技术职务 B", "A.ID,A.姓名,A.专业技术职务,B.编码", str开嘱医生Tag, " A.专业技术职务=B.名称 And A.姓名", 0, 1)

            For i = 1 To rsAdvice.RecordCount
                 '长期医嘱按品种下达时,任意取一个药品Id
                If Val(rsAdvice!药品ID & "") = 0 And Val(rsAdvice!医嘱期效 & "") = 0 Then
                    If Not rs药品 Is Nothing Then
                        rs药品.Filter = "药名ID =" & rsAdvice!诊疗项目ID
                        If Not rs药品.EOF Then rsAdvice!药品ID = rs药品!药品ID & ""
                    End If
                End If

                If InStr("," & str收费细目ID & ",", "," & rsAdvice!药品ID & ",") = 0 Then
                    str收费细目ID = str收费细目ID & "," & rsAdvice!药品ID
                End If

                If Not rs频率 Is Nothing Then
                    rs频率.Filter = "名称 ='" & rsAdvice!频率 & "' And 适用范围=" & IIf(rsAdvice!诊疗类别 & "" = "7", 2, 1)
                    If Not rs频率.EOF Then rsAdvice!频率编码 = rs频率!编码 & ""
                End If

                If Not rs开嘱医生 Is Nothing Then
                    rs开嘱医生.Filter = "姓名='" & rsAdvice!开嘱医生 & "'"
                    If Not rs开嘱医生.EOF Then
                        rsAdvice!开嘱医生ID = rs开嘱医生!id & ""
                        rsAdvice!专业技术职务 = rs开嘱医生!编码 & ""
                    End If
                End If
                If Not rs开嘱科室 Is Nothing Then
                    rs开嘱科室.Filter = "ID =" & rsAdvice!开嘱科室id
                    If Not rs开嘱科室.EOF Then rsAdvice!开嘱科室 = rs开嘱科室!名称 & ""
                End If
                rsAdvice.MoveNext
            Next

            If str收费细目ID <> "" Then
                str收费细目ID = Mid(str收费细目ID, 2)
                Set rsDrug = Get药品信息_HZYY(str收费细目ID)
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
    '功能:解析XML字符串
    'bytFunc=0-门诊;1-住院
    '<root>
    '  <base>
    '    <hospital_code><![CDATA[医院Code(必填)]]></hospital_code>
    '    <event_no><![CDATA[就诊流水号]]></event_no>
    '    <patient_id><![CDATA[病人号(需要病人唯一标识)]]></patient_id>
    '    <source><![CDATA[来源]]></source>
    '  </base>
    '  <pharm_chk_id><![CDATA[审核药师工号]]></pharm_chk_id> ----审方才有此标签
    '  <pharm_chk_name><![CDATA[审核药师名称]]></pharm_chk_name> ----审方才有此标签
    '  <btnStatus><![CDATA[按钮返回值(1修改处方，2忽略)]]></btnStatus>----干预才有此标签
    '  <message>
    '    <recipe_id><![CDATA[处方id]]></recipe_id>
    '    <is_success><![CDATA[成功标识(0审核不通过，1审核通过)]]></is_success>----审方才有此标签
    '    <infos>
    '      <info>
    '        <info_id><![CDATA[一条警示信息的唯一id]]></info_id>
    '        <group_no><![CDATA[组号]]></group_no>
    '        <drug_id><![CDATA[药品id]]></drug_id>
    '        <drug_name><![CDATA[药品名称]]></drug_name>
    '        <error_info><![CDATA[错误信息]]></error_info>
    '        <advice><![CDATA[建议]]></advice>
    '        <source><![CDATA[来源]]></source>
    '        <rt><![CDATA[消息的规则类型]]></rt>
    '        <source_id><![CDATA[来源id]]></source_id>
    '        <severity><![CDATA[错误等级]]></severity>
    '        <message_id><![CDATA[错误信息id]]></message_id>
    '        <type><![CDATA[警示信息类型]]></type>
    '        <analysis_type><![CDATA[分析类型]]></analysis_type>
    '        <analysis_result_type><![CDATA[提示类型]]></analysis_result_type>
    '        <status><![CDATA[状态:1需要双签名确认0不需要双签名确认]]></status>----审方才有此标签
    '      </info>
    '      <info>
    '             一个<info>标签，一条警示信息，多条警示信息多个<info>标签
    '      </info>
    '</infos>
    '  </message>
    '  <message>
    '      对于一张xml中传入多张处方及处方明细信息的数据传入方式，系统需返回多条<message>标签
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
100     Set rsRet = InitAdviceRS(FUN_审查结果_HZYY)
        '读取网关响应数据（XML格式）
'strData = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine & _
'            "<root><base><hospital_code>cqzl123</hospital_code><event_no><![CDATA[Q0000001]]></event_no>" & _
'            "<patient_id>101</patient_id><source><![CDATA[门诊]]></source></base><btnStatus></btnStatus>" & _
'            "<message><recipe_id><![CDATA[118]]></recipe_id><infos><info><info_id></info_id><group_no></group_no>" & _
'            "<drug_id><![CDATA[10010]]></drug_id><drug_name><![CDATA[阿莫西林胶囊]]></drug_name><error_info>" & _
'            "<![CDATA[给药途径不合适。]]></error_info><advice><![CDATA[本品宜口服给药。]]></advice>" & _
'            "<source><![CDATA[1个依据]]></source><rt><![CDATA[0]]></rt><source_id><![CDATA[SFDA药品说明书范本]]></source_id>" & _
'            "<severity><![CDATA[8]]></severity><message_id><![CDATA[1510194380657]]></message_id><type><![CDATA[给药途径]]></type>" & _
'            "<analysis_type><![CDATA[适宜性分析]]></analysis_type><analysis_result_type><![CDATA[给药途径]]>" & _
'            "</analysis_result_type><status></status></info></infos></message><version><![CDATA[V1.0]]></version></root>"
'  strData = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbNewLine & _
'        "<root>" & vbNewLine & _
'        "  <base>" & vbNewLine & _
'        "    <hospital_code>1001</hospital_code>" & vbNewLine & _
'        "    <event_no><![CDATA[17080894]]></event_no>" & vbNewLine & _
'        "    <patient_id>3394518</patient_id>" & vbNewLine & _
'        "    <source><![CDATA[住院]]></source>" & vbNewLine & _
'        "  </base>" & vbNewLine & _
'        "  <btnStatus/>"
'住院
'strData = strData & "<message>" & vbNewLine & _
'"  <group_no/>" & vbNewLine & _
'"  <infos>" & vbNewLine & _
'"    <info>" & vbNewLine & _
'"      <info_id/>" & vbNewLine & _
'"      <order_id><![CDATA[120721918]]></order_id>" & vbNewLine & _
'"      <order_item_id/>" & vbNewLine & _
'"      <drug_id><![CDATA[61851]]></drug_id>" & vbNewLine & _
'"      <drug_name><![CDATA[阿莫西林克拉维酸钾片(2:1)]]></drug_name>" & vbNewLine & _
'"      <order_type/>" & vbNewLine & _
'"      <pivas_flag/>" & vbNewLine & _
'"      <error_info><![CDATA[给药途径不合适。]]></error_info>" & vbNewLine & _
'"      <advice><![CDATA[本药宜口服胃肠道给药。]]></advice>" & vbNewLine & _
'"      <source><![CDATA[说明书]]></source>" & vbNewLine & _
'"      <rt><![CDATA[0]]></rt>" & vbNewLine & _
'"      <source_id/>" & vbNewLine & _
'"      <severity><![CDATA[5]]></severity>" & vbNewLine & _
'"      <message_id><![CDATA[1383841426395]]></message_id>" & vbNewLine & _
'"      <type><![CDATA[给药途径]]></type>" & vbNewLine & _
'"      <analysis_type><![CDATA[适宜性分析]]></analysis_type>" & vbNewLine & _
'"      <analysis_result_type><![CDATA[给药途径]]></analysis_result_type>" & vbNewLine & _
'"      <status/>" & vbNewLine & _
'"    </info>"
'strData = strData & "      <info>" & vbNewLine & _
'"        <info_id/>" & vbNewLine & _
'"        <order_id><![CDATA[120721918]]></order_id>" & vbNewLine & _
'"        <order_item_id/>" & vbNewLine & _
'"        <drug_id><![CDATA[61851]]></drug_id>" & vbNewLine & _
'"        <drug_name><![CDATA[给药途径管理]]></drug_name>" & vbNewLine & _
'"        <order_type/>" & vbNewLine & _
'"        <pivas_flag/>" & vbNewLine & _
'"        <error_info><![CDATA[测试：弹框]]></error_info>" & vbNewLine & _
'"        <advice/>" & vbNewLine & _
'"        <source><![CDATA[0个依据]]></source>" & vbNewLine & _
'"        <rt><![CDATA[0]]></rt>" & vbNewLine & _
'"        <source_id/>" & vbNewLine & _
'"        <severity><![CDATA[8]]></severity>" & vbNewLine & _
'"        <message_id><![CDATA[1513238322112]]></message_id>" & vbNewLine & _
'"        <type><![CDATA[非甾体剂量控制]]></type>" & vbNewLine & _
'"        <analysis_type><![CDATA[用药建议]]></analysis_type>" & vbNewLine & _
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
112         strValue = "病人ID:" & xmlDoc.selectSingleNode(".//patient_id").Text & Space(4)
114         strPati = strPati & strValue
116         strPati = strPati & "姓名:XX" & Space(4)
        
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
142                     rsRet!DrugName = xmlInfos(i).selectSingleNode(".//drug_name").Text   '药品名称
144                     rsRet!message = xmlInfos(i).selectSingleNode(".//error_info").Text      '错误信息
146                     rsRet!Advice = xmlInfos(i).selectSingleNode(".//advice").Text        '用药建议
148                     rsRet!Source = xmlInfos(i).selectSingleNode(".//source").Text        '来源
150                     rsRet!GroupNo = xmlInfos(i).selectSingleNode(".//group_no").Text  '组号
152                     rsRet!Type = xmlInfos(i).selectSingleNode(".//type").Text         '警示信息类型
154                     rsRet!Severity = xmlInfos(i).selectSingleNode(".//severity").Text '错误等级
156                     rsRet.Update
158                     xmlInfos.nextNode
                    Next
                End If
            Next
        Else
160         strPati = ""
162         strValue = "病人ID:" & xmlDoc.selectSingleNode(".//patient_id").Text & Space(4)
164         strPati = strPati & strValue
166         strPati = strPati & "姓名:XX" & Space(4)
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
192                     rsRet!DrugName = xmlInfos(i).selectSingleNode(".//drug_name").Text   '药品名称
194                     rsRet!message = xmlInfos(i).selectSingleNode(".//error_info").Text      '错误信息
196                     rsRet!Advice = xmlInfos(i).selectSingleNode(".//advice").Text        '用药建议
198                     rsRet!Source = xmlInfos(i).selectSingleNode(".//source").Text        '来源
200                     rsRet!GroupNo = strRecipeId  '组号
202                     rsRet!Type = xmlInfos(i).selectSingleNode(".//type").Text         '警示信息类型
204                     rsRet!Severity = xmlInfos(i).selectSingleNode(".//severity").Text '错误等级
206                     rsRet.Update
208                     xmlInfos.nextNode
                    Next
                End If
            Next
        End If
210     Set HZYY_ParseXML = rsRet
        Exit Function
errH:
212     MsgBox Err.Description & vbCrLf & "HZYY_ParseXML" & "行 " & Erl, vbExclamation + vbOKOnly, gstrSysName
End Function

