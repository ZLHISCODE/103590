Attribute VB_Name = "mdlPassDefine_DT"
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
'��ͨ�ӿڶ���
'--------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function dtywzxUI Lib "dtywzxUI" (ByVal nCode As Long, ByVal lParam As Long, ByVal lpcszBuffer As String) As Long
'nCode:0-��ʼ������ʾ�ĸ�״̬��
'1     =�˳����򣬲��ر�״̬��
'3     =ˢ��״̬�ƣ��ָ�����ʼ״̬
'768   =��¼����Ա����
'12    =���ݲ���Ա��ҩƷ������(�Ƿ�������"�ݲ���ʾ")�������Ƿ���ʾҪ����ʾ
'4108  =������ʾҪ����ʾ
'28676 =ҽ�����������������������������ʾ��
'28685 =ҽ���������������������
Public Declare Function dtywzxUI2 Lib "dtywzxUI" (ByVal nCode As Long, ByVal lParam As Long, ByVal lpcszBuffer As String, ByRef strRetXML As String) As Long
'�½ӿڣ����ڴ�����ҩ�������ļ��
'nCode:0-��ʼ������ʾ�ĸ�״̬��
'1     =�˳����򣬲��ر�״̬��
'3     =ˢ��״̬�ƣ��ָ�����ʼ״̬
'768   =��¼����Ա����
'12    =���ݲ���Ա��ҩƷ������(�Ƿ�������"�ݲ���ʾ")�������Ƿ���ʾҪ����ʾ
'4108  =������ʾҪ����ʾ
'28676 =ҽ�����������������������������ʾ��
'28685 =ҽ���������������������

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'* ��ͨXML�ӿڽṹ
Public Type dt_base
    dDoctCode As String     '* ҽ�����루���ţ�
    dDoctName As String     '* ҽ������
    dDoctType As String     '* ҽ���������
    dDeptCode As String     '* ���Ҵ���
    dDeptName As String     '* ��������
    dInHosCode As String    '* סԺ��
    dBedNo As String        '* סԺ����
    mPresDate As Date       '* ����ʱ��
    pCaseID As String       '* ��������
    pOutID As String        '* ��������(�Һŵ���)
    pWeight As String       '* ��������
    pHeight As String       '* �������
    pBirthday As Date       '* ���˳�������
    pPatiName As String     '* ��������
    pSex As String          '* �����Ա�
    pStatms As String       '* �������
    pEffect As String       '* ������
    pBloodPress As String   '* Ѫѹ
    pLiverClean As String   '* ���������
    pCaseCode1 As String    '* ����Դ����1
    pCaseName1 As String    '* ����Դ����1
    pCaseCode2 As String    '* ����Դ����2
    pCaseName2 As String    '* ����Դ����2
    pCaseCode3 As String    '* ����Դ����3
    pCaseName3 As String    '* ����Դ����3
    pDiagnose1 As String    '* ���1(ICD10)
    pDiagnose2 As String    '* ���2(ICD10)
    pDiagnose3 As String    '* ���3(ICD10)
    pDiagnoseName1 As String    '* ���1(����)
    pDiagnoseName2 As String    '* ���2(����)
    pDiagnoseName3 As String    '* ���3(����)
    pBsl1 As String         '* ���������1
    pBsl2 As String         '* ���������2
    pBsl3 As String         '* ���������3
End Type

'* ��ͨ�������ݽṹ
Public Type dt_Pres
    PresID As String        '* ҽ����,û�оʹ�סԺ��
    PresType As String      '* �������ͣ�ҽ�����ͣ�L--���ڣ�T--��ʱ��
    Current As Long         '* ��ǰ������ǣ�1����ǰ������/0����ʷ������
    GroupNum As Long        '* ��ҩ���
    GeneralName As String   '* ҩƷͨ����
    HosMediCode As String   '* ҽԺҩƷ����
    MediName As String      '* ҩƷ��Ʒ��
    DCL As String           '* ������
    PCDM As String          '* Ƶ�δ���
    Days As String          '* ����
    Unit As String          '* ��λ
    GYTJ As String          '* �ã�����ҩ;��
    BTime As String         '* ��ҩ��ʼʱ��,����ҽ���ĵ�һ��ĵ�һ�ε���ҩʱ��
    ETime As String         '* ��ҩ����ʱ��,����ҽ����ͣ��ʱ�䡣�����ҩ����ʹ�ã�ҽ��û��ͣҩʱ��Ļ�����Ϊ��
    PresTime As String      '* ҽ������ҽ����ʱ��
End Type

'--------------------------------------------------------------------------------------------------------------------------------------
'��ͨ�ӿں���
'--------------------------------------------------------------------------------------------------------------------------------------
'* ����XML����
'* �˺�������Ϊͨ�ú����������޸�
Public Function MakeXML(ByRef xmlbase As dt_base, ByRef arrPres As Variant, bytFun As Byte) As String
'������bytFun:0-������ã�1-סԺ����
    Dim strXML As String, lngIndex As Long
    Dim strTab1 As String, strTab2 As String, strTab3 As String, strPati As String, strAge As String
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    strTab3 = vbCrLf & vbTab & vbTab & vbTab
    
    With xmlbase
        
        '* ҽ����������Ϣ
        strXML = "<safe>" & strTab1 & "<doctor_information job_number='" & .dDoctCode & "' " & _
                 "date='" & Format(.mPresDate, "yyyy-MM-dd HH:mm:ss") & "'/>" & strTab1 & _
                 "<doctor_name>" & .dDoctName & "</doctor_name>" & strTab1 & _
                 "<doctor_type>" & .dDoctType & "</doctor_type>" & strTab1 & _
                 "<department_code>" & .dDeptCode & "</department_code>" & strTab1 & _
                 "<department_name>" & .dDeptName & "</department_name>" & strTab1 & _
                 "<case_id>" & .pCaseID & "</case_id>" & strTab1 & _
                 "<inhos_code>" & IIf(bytFun = 0, .pOutID, .dInHosCode) & "</inhos_code>" & strTab1 & _
                 "<bed_no>" & .dBedNo & "</bed_no>" & strTab1
        
        If bytFun = 0 Then
            strPati = "patrent"
            strAge = "age"
        Else
            strPati = "patient"
            strAge = "birth"
        End If
        '* ������Ϣ
        strXML = strXML & "<patient_information weight='" & .pWeight & "' height='" & .pHeight & "' " & _
                 strAge & "='" & IIf((.pBirthday = vbNull), "", Format(.pBirthday, "yyyy-MM-dd")) & "'>" & strTab2 & _
                 "<" & strPati & "_name>" & .pPatiName & "</" & strPati & "_name>" & strTab2 & _
                 "<" & strPati & "_sex>" & .pSex & "</" & strPati & "_sex>" & strTab2 & _
                 "<physiological_statms>" & .pStatms & "</physiological_statms>" & strTab2 & _
                 "<boacterioscopy_effect>" & .pEffect & "</boacterioscopy_effect>" & strTab2 & _
                 "<bloodpressure>" & .pBloodPress & "</bloodpressure>" & strTab2 & _
                 "<liver_clean>" & .pLiverClean & "</liver_clean>" & strTab2 & _
                 IIf(bytFun = 0, "", "<pregnant></pregnant>" & strTab2 & "<pdw></pdw>" & strTab2)

        
        '* ����Դ
        strXML = strXML & "<allergic_history>" & strTab3 & "<case>" & strTab3 & vbTab & _
                 "<case_code>" & .pCaseCode1 & "</case_code>" & strTab3 & vbTab & _
                 "<case_name>" & .pCaseName1 & "</case_name>" & strTab3 & _
                 "</case>" & strTab3 & "<case>" & strTab3 & vbTab & _
                 "<case_code>" & .pCaseCode2 & "</case_code>" & strTab3 & vbTab & _
                 "<case_name>" & .pCaseName2 & "</case_name>" & strTab3 & _
                 "</case>" & strTab3 & "<case>" & strTab3 & vbTab & _
                 "<case_code>" & .pCaseCode3 & "</case_code>" & strTab3 & vbTab & _
                 "<case_name>" & .pCaseName3 & "</case_name>" & strTab3 & _
                 "</case>" & strTab2 & "</allergic_history>" & strTab2
        
        '* �����Ϣ�����������  --75326
        strXML = strXML & "<diagnoses>" & strTab3 & _
                 "<diagnose type = '0' name='" & StrToXML(.pDiagnoseName1) & "'>" & .pDiagnose1 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '0' name='" & StrToXML(.pDiagnoseName2) & "'>" & .pDiagnose2 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '0' name='" & StrToXML(.pDiagnoseName3) & "'>" & .pDiagnose3 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '1' name='" & .pBsl1 & "'>" & .pBsl1 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '1' name='" & .pBsl2 & "'>" & .pBsl2 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '1' name='" & .pBsl3 & "'>" & .pBsl3 & "</diagnose>" & strTab2 & _
                 "</diagnoses>" & strTab1 & "</patient_information>"

        
        '* ������Ϣ
        strXML = strXML & strTab1 & "<prescriptions>"
        For lngIndex = 0 To UBound(arrPres)
            strXML = strXML & arrPres(lngIndex)
        Next lngIndex
        strXML = strXML & strTab1 & "</prescriptions>" & vbCrLf & "</safe>"
    End With
''    Debug.Print strXML
    MakeXML = strXML
End Function

'* ���ɴ�����ҩ��XML����
'* �˺�������Ϊͨ�ú����������޸�
Public Function MakePresXML(ByRef dtpres As dt_Pres, bytFun As Byte) As String
'������bytFun:0-������ã�1-סԺ����
    Dim strXML As String
    Dim strTab2 As String, strTab3 As String, strTab4 As String
    
    strTab2 = vbCrLf & vbTab & vbTab
    strTab3 = vbCrLf & vbTab & vbTab & vbTab
    strTab4 = vbCrLf & vbTab & vbTab & vbTab & vbTab
    
    With dtpres
        If bytFun = 0 Then
            strXML = strTab2 & "<prescription id='" & .PresID & "' " & _
                 "type='" & .PresType & "' current='" & .Current & "'>" & strTab3 & _
                 "<medicine suspension='false' judge='true'>" & strTab4 & _
                 "<group_number>" & .GroupNum & "</group_number>" & strTab4 & _
                 "<general_name>" & .GeneralName & "</general_name>" & strTab4 & _
                 "<license_number>" & .HosMediCode & "</license_number>" & strTab4 & _
                 "<medicine_name>" & .MediName & "</medicine_name>" & strTab4 & _
                 "<single_dose coef='1'>" & .DCL & "</single_dose>" & strTab4 & _
                 "<times>" & .PCDM & "</times>" & strTab4 & _
                 "<days>" & .Days & "</days>" & strTab4 & _
                 "<unit>" & .Unit & "</unit>" & strTab4 & _
                 "<administer_drugs>" & .GYTJ & "</administer_drugs>" & strTab3 & _
                 "</medicine>" & strTab2 & "</prescription>"
        Else
            strXML = strTab2 & "<prescription id='" & .PresID & "' " & _
                 "type='" & .PresType & "'>" & strTab3 & _
                 "<medicine suspension='false' judge='true'>" & strTab4 & _
                 "<group_number>" & .GroupNum & "</group_number>" & strTab4 & _
                 "<general_name>" & .GeneralName & "</general_name>" & strTab4 & _
                 "<license_number>" & .HosMediCode & "</license_number>" & strTab4 & _
                 "<medicine_name>" & .MediName & "</medicine_name>" & strTab4 & _
                 "<single_dose coef='1'>" & .DCL & "</single_dose>" & strTab4 & _
                 "<frequency>" & .PCDM & "</frequency>" & strTab4 & _
                 "<times></times>" & strTab4 & _
                 "<unit>" & .Unit & "</unit>" & strTab4 & _
                 "<administer_drugs>" & .GYTJ & "</administer_drugs>" & strTab3 & _
                 "<begin_time>" & .BTime & "</begin_time>" & strTab3 & _
                 "<end_time>" & .ETime & "</end_time>" & strTab3 & _
                 "<prescription_time>" & .PresTime & "</prescription_time>" & strTab3 & _
                 "</medicine>" & strTab2 & "</prescription>"
        End If
    End With
    MakePresXML = strXML
End Function

Public Function MakeMediXML(ByVal strҩƷ���� As String, ByVal strҩƷID As String, Optional blnReShow As Boolean) As String
    Dim strXML As String

    strXML = "<general_name>" & strҩƷ���� & "</general_name>" & vbCrLf & _
             "<license_number>" & strҩƷID & "</license_number>"
    If blnReShow Then strXML = "<safe>" & strXML & "</safe>"
    
    MakeMediXML = strXML
End Function

Public Function MakeMediDelXML(ByVal str������ As String, ByVal date�������� As Date) As String
    Dim strXML As String

    strXML = "<prescription id='" & str������ & "' date='" & Format(date��������, "yyyy-MM-dd HH:mm:ss") & "'/>"
    
    MakeMediDelXML = strXML
End Function

'* ���������ı��еļ�����
'* �˺�������Ϊͨ�ú����������������޸�
Public Function StrToXML(ByVal strValue As String) As String
    StrToXML = Replace(Replace(Replace(strValue, "<", ""), ">", ""), "'", "")
End Function

Public Function GetAlertFromXml(ByVal strRetXML As String) As String
'���ܣ�ȡ��XML��Alert�ڵ��µ��ַ���
    If InStr(strRetXML, "<ALERT>") <> 0 And InStr(strRetXML, "</ALERT>") <> 0 Then
        GetAlertFromXml = Mid(strRetXML, InStr(strRetXML, "<ALERT>") + 7, InStr(strRetXML, "</ALERT>") - InStr(strRetXML, "<ALERT>") - 7)
    End If
End Function


