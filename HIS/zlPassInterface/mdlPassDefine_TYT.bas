Attribute VB_Name = "mdlPassDefine_TYT"
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
'太元通接口函数
'--------------------------------------------------------------------------------------------------------------------------------------
'* 生成XML代码
Public Function MakePatientOrderXml(ByRef patOrder As PatientOrder) As String
'功能:将传人的自定义数据类型组装成一个XML串
'参数：病人相关信息及医嘱相关信息组合而成的自定义数据类型

    Dim strXML As String, strTmp As String
    Dim udtDrug As PatDrug
    Dim udtDiagnosis As PatDiagnosis    '诊断记录
    Dim udtDrugSensitive As PatDrugSensitive    '过敏记录
    Dim udtSymptom As PatSymptom     '症状记录
    Dim i As Long

    With patOrder
        '病人信息
        strXML = "<PatientOrder><Patient patientID='" & .PatientID & "' name='" & .Pname & "' sex='" & .pSex & _
                 "' dateOfBirth='" & .pdateOfBirth & "'></Patient>" & _
                 "<PatOrderInfoExt isLact='" & .isLact & "'  isPregnant ='" & .isPregnant & "' isLiverWhole='" & .isLiverWhole & "' isKidneyWhole='" & .isKidneyWhole & _
                 "'  height='" & .pHeight & "' weight='" & .pWeight & "'></PatOrderInfoExt>" & _
                 "<PatOrderVisitInfo visitID='" & .PvisitID & "' ></PatOrderVisitInfo>" & _
                 "<PatOrderDrugs></PatOrderDrugs><PatOrderDiagnoses></PatOrderDiagnoses><PatOrderDrugSensitives></PatOrderDrugSensitives>" & _
                 "<PatOrderSymptoms></PatOrderSymptoms>"
        '登录医生信息
        strXML = strXML & "<DoctorDeptID>" & .DoctDeptID & "</DoctorDeptID><DoctorDeptName>" & .DoctDeptName & "</DoctorDeptName>" & _
                 "<DoctorID>" & .DoctID & "</DoctorID><DoctorName>" & .DoctName & "</DoctorName><DoctorTitleID>" & .DoctTitleID & "</DoctorTitleID>" & _
                 "<DoctorTitleName>" & .DoctTitleName & "</DoctorTitleName><SysFlag>" & .SysFlag & "</SysFlag></PatientOrder>"
        '药嘱信息
        strTmp = ""
        For i = LBound(.PatDrugs) To UBound(.PatDrugs)
            udtDrug = .PatDrugs(i)

            With udtDrug
                strTmp = strTmp & "<PatOrderDrug drugID ='" & .drugID & "' drugName='" & .DrugName & "'" & _
                         " recMainNo='" & .recMainNo & "' recSubNo='" & .recSubNo & "' dosage='" & .dosage & "'" & _
                         " doseUnits='" & .doseUnits & "' administrationID='" & .administrationID & "'" & _
                         " performFreqDictID='" & .performFreqDictID & "' performFreqDictText='" & .performFreqDictText & "'" & _
                         " startDateTime= '" & .startDateTime & "' stopDateTime='" & .stopDateTime & "'" & _
                         " doctorDept='" & .doctorDept & "' doctorID='" & .DoctorID & "' doctor='" & .Doctor & "'" & _
                         " isNew='" & .isNew & "'></PatOrderDrug>"
            End With
        Next

        strXML = Replace(strXML, "<PatOrderDrugs>", "<PatOrderDrugs>" & strTmp)

        '诊断信息
        strTmp = ""
        For i = LBound(.PatDiagnoses) To UBound(.PatDiagnoses)
            udtDiagnosis = .PatDiagnoses(i)
            With udtDiagnosis
                strTmp = strTmp & "<PatOrderDiagnosis diagnosisID='" & .diagnosisID & "' diagnosisName='" & .diagnosisName & "'" & _
                         " diagnosisType= '" & .diagnosisType & "'></PatOrderDiagnosis>"
            End With
        Next
        strXML = Replace(strXML, "<PatOrderDiagnoses>", "<PatOrderDiagnoses>" & strTmp)

        '过敏信息
        strTmp = ""
        For i = LBound(.PatDrugSensitives) To UBound(.PatDrugSensitives)
            udtDrugSensitive = .PatDrugSensitives(i)
            With udtDrugSensitive
                strTmp = strTmp & "<PatOrderDrugSensitive patOrderDrugSensitiveID='0' drugAllergenID='" & .drugAllergenID & "'></PatOrderDrugSensitive>"
            End With
        Next
        strXML = Replace(strXML, "<PatOrderDrugSensitives>", "<PatOrderDrugSensitives>" & strTmp)

        '症状信息
        strTmp = ""
        For i = LBound(.PatSymptoms) To UBound(.PatSymptoms)
            udtSymptom = .PatSymptoms(i)
            With udtSymptom
                strTmp = strTmp & "<PatOrderSymptom symptomID='" & .symptomID & "' symptomName='" & .symptomName & "'></PatOrderSymptom>"
            End With
        Next
        strXML = Replace(strXML, "<PatOrderSymptoms>", "<PatOrderSymptoms>" & strTmp)

    End With

    MakePatientOrderXml = strXML

End Function

Public Function AnalyzeReturnXml(ByVal strXML As String) As Collection
'功能：将太元通审查结果集进行解析，分离出存在审核问题的药嘱行
    Dim colAuditResult As Collection
    Dim strSub As String
    Dim lngBegin As Long, lngEnd As Long
    Dim lngSubBegin As Long, lngSubEnd As Long
    Dim udtAuditResult As AuditResult

    Set colAuditResult = New Collection

    Do
        lngBegin = InStr(strXML, "<Table1>") + 8
        lngEnd = InStr(strXML, "</Table1>")
        If lngBegin = 8 Then Exit Do
        strSub = Mid(strXML, lngBegin, lngEnd - lngBegin)  '<Table1>...</Table1>中所包含字符串: ...
        '取recMainNo-组医嘱号
        lngSubBegin = InStr(strSub, "<recMainNo>") + 11
        lngSubEnd = InStr(strSub, "</recMainNo>")
        udtAuditResult.recMainNo = Mid(strSub, lngSubBegin, lngSubEnd - lngSubBegin)
        '取recSubNo-医嘱序号
        lngSubBegin = InStr(strSub, "<recSubNo>") + 10
        lngSubEnd = InStr(strSub, "</recSubNo>")
        udtAuditResult.recSubNo = Mid(strSub, lngSubBegin, lngSubEnd - lngSubBegin)
        '取check_alertLevel-警示级别
        lngSubBegin = InStr(strSub, "<check_alertLevel>") + 18
        lngSubEnd = InStr(strSub, "</check_alertLevel>")
        udtAuditResult.alertLevel = Mid(strSub, lngSubBegin, lngSubEnd - lngSubBegin)
        '取strChecksum-审查结果
        lngSubBegin = InStr(strSub, "<strChecksum>") + 13
        lngSubEnd = InStr(strSub, "</strChecksum>")
        udtAuditResult.strChecksum = Mid(strSub, lngSubBegin, lngSubEnd - lngSubBegin)

        colAuditResult.Add udtAuditResult, udtAuditResult.recMainNo & "_" & udtAuditResult.recSubNo  '取组医嘱号和序号作为关键字，方便迅速定位数据

        strXML = Mid(strXML, lngEnd + 9)

    Loop While lngBegin <> 0

    Set AnalyzeReturnXml = colAuditResult
End Function

