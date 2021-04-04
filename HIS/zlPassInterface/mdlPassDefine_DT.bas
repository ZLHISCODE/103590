Attribute VB_Name = "mdlPassDefine_DT"
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
'大通接口定义
'--------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function dtywzxUI Lib "dtywzxUI" (ByVal nCode As Long, ByVal lParam As Long, ByVal lpcszBuffer As String) As Long
'nCode:0-初始化，显示四个状态灯
'1     =退出程序，并关闭状态灯
'3     =刷新状态灯，恢复到初始状态
'768   =记录操作员工号
'12    =根据操作员对药品的设置(是否设置了"暂不提示")来决定是否显示要点提示
'4108  =总是显示要点提示
'28676 =医嘱配伍分析（不保存分析结果，仅提示）
'28685 =医嘱配伍分析，保存分析结果
Public Declare Function dtywzxUI2 Lib "dtywzxUI" (ByVal nCode As Long, ByVal lParam As Long, ByVal lpcszBuffer As String, ByRef strRetXML As String) As Long
'新接口，用于处理用药量超量的检查
'nCode:0-初始化，显示四个状态灯
'1     =退出程序，并关闭状态灯
'3     =刷新状态灯，恢复到初始状态
'768   =记录操作员工号
'12    =根据操作员对药品的设置(是否设置了"暂不提示")来决定是否显示要点提示
'4108  =总是显示要点提示
'28676 =医嘱配伍分析（不保存分析结果，仅提示）
'28685 =医嘱配伍分析，保存分析结果

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'* 大通XML接口结构
Public Type dt_base
    dDoctCode As String     '* 医生代码（工号）
    dDoctName As String     '* 医生姓名
    dDoctType As String     '* 医生级别代码
    dDeptCode As String     '* 科室代码
    dDeptName As String     '* 科室名称
    dInHosCode As String    '* 住院号
    dBedNo As String        '* 住院床号
    mPresDate As Date       '* 处方时间
    pCaseID As String       '* 病历卡号
    pOutID As String        '* 门诊就诊号(挂号单号)
    pWeight As String       '* 病人体重
    pHeight As String       '* 病人身高
    pBirthday As Date       '* 病人出生日期
    pPatiName As String     '* 病人姓名
    pSex As String          '* 病人性别
    pStatms As String       '* 生理情况
    pEffect As String       '* 菌检结果
    pBloodPress As String   '* 血压
    pLiverClean As String   '* 肌肝清除率
    pCaseCode1 As String    '* 过敏源代码1
    pCaseName1 As String    '* 过敏源名称1
    pCaseCode2 As String    '* 过敏源代码2
    pCaseName2 As String    '* 过敏源名称2
    pCaseCode3 As String    '* 过敏源代码3
    pCaseName3 As String    '* 过敏源名称3
    pDiagnose1 As String    '* 诊断1(ICD10)
    pDiagnose2 As String    '* 诊断2(ICD10)
    pDiagnose3 As String    '* 诊断3(ICD10)
    pDiagnoseName1 As String    '* 诊断1(名称)
    pDiagnoseName2 As String    '* 诊断2(名称)
    pDiagnoseName3 As String    '* 诊断3(名称)
    pBsl1 As String         '* 病生理情况1
    pBsl2 As String         '* 病生理情况2
    pBsl3 As String         '* 病生理情况3
End Type

'* 大通处方数据结构
Public Type dt_Pres
    PresID As String        '* 医嘱号,没有就传住院号
    PresType As String      '* 处方类型：医嘱类型（L--长期，T--临时）
    Current As Long         '* 当前处方标记：1（当前处方）/0（历史处方）
    GroupNum As Long        '* 用药组号
    GeneralName As String   '* 药品通用名
    HosMediCode As String   '* 医院药品代码
    MediName As String      '* 药品商品名
    DCL As String           '* 单次量
    PCDM As String          '* 频次代码
    Days As String          '* 天数
    Unit As String          '* 单位
    GYTJ As String          '* 用（给）药途径
    BTime As String         '* 用药开始时间,长期医嘱的第一天的第一次的用药时间
    ETime As String         '* 用药结束时间,长期医嘱的停嘱时间。如果该药正在使用，医生没设停药时间的话，可为空
    PresTime As String      '* 医生在下医嘱的时间
End Type

'--------------------------------------------------------------------------------------------------------------------------------------
'大通接口函数
'--------------------------------------------------------------------------------------------------------------------------------------
'* 生成XML代码
'* 此函数过程为通用函数，无需修改
Public Function MakeXML(ByRef xmlbase As dt_base, ByRef arrPres As Variant, bytFun As Byte) As String
'参数：bytFun:0-门诊调用，1-住院调用
    Dim strXML As String, lngIndex As Long
    Dim strTab1 As String, strTab2 As String, strTab3 As String, strPati As String, strAge As String
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    strTab3 = vbCrLf & vbTab & vbTab & vbTab
    
    With xmlbase
        
        '* 医生及科室信息
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
        '* 病人信息
        strXML = strXML & "<patient_information weight='" & .pWeight & "' height='" & .pHeight & "' " & _
                 strAge & "='" & IIf((.pBirthday = vbNull), "", Format(.pBirthday, "yyyy-MM-dd")) & "'>" & strTab2 & _
                 "<" & strPati & "_name>" & .pPatiName & "</" & strPati & "_name>" & strTab2 & _
                 "<" & strPati & "_sex>" & .pSex & "</" & strPati & "_sex>" & strTab2 & _
                 "<physiological_statms>" & .pStatms & "</physiological_statms>" & strTab2 & _
                 "<boacterioscopy_effect>" & .pEffect & "</boacterioscopy_effect>" & strTab2 & _
                 "<bloodpressure>" & .pBloodPress & "</bloodpressure>" & strTab2 & _
                 "<liver_clean>" & .pLiverClean & "</liver_clean>" & strTab2 & _
                 IIf(bytFun = 0, "", "<pregnant></pregnant>" & strTab2 & "<pdw></pdw>" & strTab2)

        
        '* 过敏源
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
        
        '* 诊断信息、病生理情况  --75326
        strXML = strXML & "<diagnoses>" & strTab3 & _
                 "<diagnose type = '0' name='" & StrToXML(.pDiagnoseName1) & "'>" & .pDiagnose1 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '0' name='" & StrToXML(.pDiagnoseName2) & "'>" & .pDiagnose2 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '0' name='" & StrToXML(.pDiagnoseName3) & "'>" & .pDiagnose3 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '1' name='" & .pBsl1 & "'>" & .pBsl1 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '1' name='" & .pBsl2 & "'>" & .pBsl2 & "</diagnose>" & strTab3 & _
                 "<diagnose type = '1' name='" & .pBsl3 & "'>" & .pBsl3 & "</diagnose>" & strTab2 & _
                 "</diagnoses>" & strTab1 & "</patient_information>"

        
        '* 处方信息
        strXML = strXML & strTab1 & "<prescriptions>"
        For lngIndex = 0 To UBound(arrPres)
            strXML = strXML & arrPres(lngIndex)
        Next lngIndex
        strXML = strXML & strTab1 & "</prescriptions>" & vbCrLf & "</safe>"
    End With
''    Debug.Print strXML
    MakeXML = strXML
End Function

'* 生成处方用药的XML代码
'* 此函数过程为通用函数，无需修改
Public Function MakePresXML(ByRef dtpres As dt_Pres, bytFun As Byte) As String
'参数：bytFun:0-门诊调用，1-住院调用
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

Public Function MakeMediXML(ByVal str药品名称 As String, ByVal str药品ID As String, Optional blnReShow As Boolean) As String
    Dim strXML As String

    strXML = "<general_name>" & str药品名称 & "</general_name>" & vbCrLf & _
             "<license_number>" & str药品ID & "</license_number>"
    If blnReShow Then strXML = "<safe>" & strXML & "</safe>"
    
    MakeMediXML = strXML
End Function

Public Function MakeMediDelXML(ByVal str处方号 As String, ByVal date作废日期 As Date) As String
    Dim strXML As String

    strXML = "<prescription id='" & str处方号 & "' date='" & Format(date作废日期, "yyyy-MM-dd HH:mm:ss") & "'/>"
    
    MakeMediDelXML = strXML
End Function

'* 忽略输入文本中的尖括号
'* 此函数过程为通用函数，基本上无需修改
Public Function StrToXML(ByVal strValue As String) As String
    StrToXML = Replace(Replace(Replace(strValue, "<", ""), ">", ""), "'", "")
End Function

Public Function GetAlertFromXml(ByVal strRetXML As String) As String
'功能：取得XML中Alert节点下的字符串
    If InStr(strRetXML, "<ALERT>") <> 0 And InStr(strRetXML, "</ALERT>") <> 0 Then
        GetAlertFromXml = Mid(strRetXML, InStr(strRetXML, "<ALERT>") + 7, InStr(strRetXML, "</ALERT>") - InStr(strRetXML, "<ALERT>") - 7)
    End If
End Function


