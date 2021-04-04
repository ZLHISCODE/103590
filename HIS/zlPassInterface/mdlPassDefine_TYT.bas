Attribute VB_Name = "mdlPassDefine_TYT"
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
'̫Ԫͨ�ӿں���
'--------------------------------------------------------------------------------------------------------------------------------------
'* ����XML����
Public Function MakePatientOrderXml(ByRef patOrder As PatientOrder) As String
'����:�����˵��Զ�������������װ��һ��XML��
'���������������Ϣ��ҽ�������Ϣ��϶��ɵ��Զ�����������

    Dim strXML As String, strTmp As String
    Dim udtDrug As PatDrug
    Dim udtDiagnosis As PatDiagnosis    '��ϼ�¼
    Dim udtDrugSensitive As PatDrugSensitive    '������¼
    Dim udtSymptom As PatSymptom     '֢״��¼
    Dim i As Long

    With patOrder
        '������Ϣ
        strXML = "<PatientOrder><Patient patientID='" & .PatientID & "' name='" & .Pname & "' sex='" & .pSex & _
                 "' dateOfBirth='" & .pdateOfBirth & "'></Patient>" & _
                 "<PatOrderInfoExt isLact='" & .isLact & "'  isPregnant ='" & .isPregnant & "' isLiverWhole='" & .isLiverWhole & "' isKidneyWhole='" & .isKidneyWhole & _
                 "'  height='" & .pHeight & "' weight='" & .pWeight & "'></PatOrderInfoExt>" & _
                 "<PatOrderVisitInfo visitID='" & .PvisitID & "' ></PatOrderVisitInfo>" & _
                 "<PatOrderDrugs></PatOrderDrugs><PatOrderDiagnoses></PatOrderDiagnoses><PatOrderDrugSensitives></PatOrderDrugSensitives>" & _
                 "<PatOrderSymptoms></PatOrderSymptoms>"
        '��¼ҽ����Ϣ
        strXML = strXML & "<DoctorDeptID>" & .DoctDeptID & "</DoctorDeptID><DoctorDeptName>" & .DoctDeptName & "</DoctorDeptName>" & _
                 "<DoctorID>" & .DoctID & "</DoctorID><DoctorName>" & .DoctName & "</DoctorName><DoctorTitleID>" & .DoctTitleID & "</DoctorTitleID>" & _
                 "<DoctorTitleName>" & .DoctTitleName & "</DoctorTitleName><SysFlag>" & .SysFlag & "</SysFlag></PatientOrder>"
        'ҩ����Ϣ
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

        '�����Ϣ
        strTmp = ""
        For i = LBound(.PatDiagnoses) To UBound(.PatDiagnoses)
            udtDiagnosis = .PatDiagnoses(i)
            With udtDiagnosis
                strTmp = strTmp & "<PatOrderDiagnosis diagnosisID='" & .diagnosisID & "' diagnosisName='" & .diagnosisName & "'" & _
                         " diagnosisType= '" & .diagnosisType & "'></PatOrderDiagnosis>"
            End With
        Next
        strXML = Replace(strXML, "<PatOrderDiagnoses>", "<PatOrderDiagnoses>" & strTmp)

        '������Ϣ
        strTmp = ""
        For i = LBound(.PatDrugSensitives) To UBound(.PatDrugSensitives)
            udtDrugSensitive = .PatDrugSensitives(i)
            With udtDrugSensitive
                strTmp = strTmp & "<PatOrderDrugSensitive patOrderDrugSensitiveID='0' drugAllergenID='" & .drugAllergenID & "'></PatOrderDrugSensitive>"
            End With
        Next
        strXML = Replace(strXML, "<PatOrderDrugSensitives>", "<PatOrderDrugSensitives>" & strTmp)

        '֢״��Ϣ
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
'���ܣ���̫Ԫͨ����������н����������������������ҩ����
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
        strSub = Mid(strXML, lngBegin, lngEnd - lngBegin)  '<Table1>...</Table1>���������ַ���: ...
        'ȡrecMainNo-��ҽ����
        lngSubBegin = InStr(strSub, "<recMainNo>") + 11
        lngSubEnd = InStr(strSub, "</recMainNo>")
        udtAuditResult.recMainNo = Mid(strSub, lngSubBegin, lngSubEnd - lngSubBegin)
        'ȡrecSubNo-ҽ�����
        lngSubBegin = InStr(strSub, "<recSubNo>") + 10
        lngSubEnd = InStr(strSub, "</recSubNo>")
        udtAuditResult.recSubNo = Mid(strSub, lngSubBegin, lngSubEnd - lngSubBegin)
        'ȡcheck_alertLevel-��ʾ����
        lngSubBegin = InStr(strSub, "<check_alertLevel>") + 18
        lngSubEnd = InStr(strSub, "</check_alertLevel>")
        udtAuditResult.alertLevel = Mid(strSub, lngSubBegin, lngSubEnd - lngSubBegin)
        'ȡstrChecksum-�����
        lngSubBegin = InStr(strSub, "<strChecksum>") + 13
        lngSubEnd = InStr(strSub, "</strChecksum>")
        udtAuditResult.strChecksum = Mid(strSub, lngSubBegin, lngSubEnd - lngSubBegin)

        colAuditResult.Add udtAuditResult, udtAuditResult.recMainNo & "_" & udtAuditResult.recSubNo  'ȡ��ҽ���ź������Ϊ�ؼ��֣�����Ѹ�ٶ�λ����

        strXML = Mid(strXML, lngEnd + 9)

    Loop While lngBegin <> 0

    Set AnalyzeReturnXml = colAuditResult
End Function

