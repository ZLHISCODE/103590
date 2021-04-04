Attribute VB_Name = "mdlOperationLog"
Option Explicit
'-- 门诊输液操作日志模块
Public Enum OPERTYPE
    QUEUE = 1       '排队操作日志
    MEDICAL = 2     '医嘱
    CALLS = 3       '呼叫
    SEAT = 4        '坐位
End Enum

Public Sub SaveOperLog(ByVal lngDeptID As Long, ByVal varNO As Variant, ByVal Oper As OPERTYPE, ByVal strLogInfo As String)
'功能：操作日志
'参数：
'  lngDeptID：执行科室ID
'  varNO：0_病人ID_挂号单号（门诊）；1_病人ID_主页ID（门诊留观）或者 病人对象(cPatient)
'  Oper：操作类型
'  strLogInfo：日志内容
'返回：

    Dim strSQL As String, strBillNO As String
    Dim lngID As Long, lngPatiID As Long, lngPageID As Long
    Dim strNO As String
    Dim objPati As cPatient
    
    If UCase(TypeName(varNO)) = "CPATIENT" Then
        Set objPati = varNO
        If Not objPati Is Nothing Then
            If objPati.病人来源 = 1 Then
                strNO = "1_" & objPati.Key
            Else
                strNO = "0_" & objPati.病人ID & "_" & objPati.挂号单
            End If
        End If
        strNO = strNO & "__"
    Else
        strNO = varNO & "__"
    End If
    
    If Val(strNO) = 0 Then
        '门诊
        lngPatiID = Val(Split(strNO, "_")(1))
        strBillNO = Trim(Split(strNO, "_")(2))
    Else
        '门诊留观
        lngPatiID = Val(Split(strNO, "_")(1))
        lngPageID = Val(Split(strNO, "_")(2))
    End If
    
    On Error GoTo hErr
    
    lngID = zldatabase.GetNextId("门诊输液操作日志")
    '--1-排队操作日志 2－医嘱操作日志 3-呼叫操作日志 4-坐位操作日志
    strLogInfo = DelInvalidChar(strLogInfo, "%|""?")
    
    'strSQL = "ZL_门诊输液操作日志_Add(" & lngID & "," & lngDeptID & ",'" & strNO & "'," & Oper & ",'" & strLogInfo & "','" & UserInfo.用户名 & "')"
    strSQL = "ZL_门诊输液操作日志_Add(" & lngID & "," & IIf(lngDeptID = 0, "null", lngDeptID) & "," & _
                IIf(lngPatiID <= 0, "Null", lngPatiID) & "," & _
                IIf(strBillNO = "", "Null", "'" & strBillNO & "'") & "," & _
                IIf(lngPageID <= 0, "Null", lngPageID) & "," & _
                Oper & ",'" & strLogInfo & "','" & UserInfo.用户名 & "')"
    Call zldatabase.ExecuteProcedure(strSQL, "保存操作日志")
    Exit Sub
    
hErr:
    SaveErrLog
End Sub
