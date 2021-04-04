Attribute VB_Name = "mdlPrint"
Option Explicit

Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjComlib As Object
Public gobjReport As Object
Public glngSys As Long
Public gcnOracle As New ADODB.Connection

'首页信息;医嘱记录;住院病历;护理记录;护理病历;诊疗报告;疾病证明;知情文件;临床路径
'内部应用模块号定义
Public Enum Enum_Inside_Program
    p电子病历管理 = 2250
    p新版住院病历 = 2252
    p新版门诊病历 = 2251
    p疾病报告填写 = 1249
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p临床路径应用 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p电子病案查阅 = 1259
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    P新版护士站 = 1265
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    p观片工具管理 = 1289
    p病人入出 = 1132
    p住院记帐 = 1133
    p费用查询 = 1139
    p门诊分诊管理 = 1113
    p排队叫号虚拟模块 = 1160
    p抗菌用药审核 = 1266
    p手术审核管理 = 1267
    p电子病案审查 = 1560
    p输血审核管理 = 1268
    p手麻接口 = 2425
    p手术授权管理 = 1080
    p输液配置中心 = 1345
    P门诊路径应用 = 1248
    P病案查阅打印 = 1566
    P体检内部接口 = 2150
End Enum

Public Function GetPatiInfo(ByVal lngPatiID As Long, ByVal lngVisitID As Long) As ADODB.Recordset
'功能:获取病人信息
    Dim strSQL As String
        
    strSQL = "Select 出院科室id,姓名,住院号,数据转出 From 病案主页 Where 病人id = [1] And 主页id = [2]"
    On Error GoTo errH
    Set GetPatiInfo = gobjDatabase.OpenSQLRecord(strSQL, "GetPatiInfo", lngPatiID, lngVisitID)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun存在该函数
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

