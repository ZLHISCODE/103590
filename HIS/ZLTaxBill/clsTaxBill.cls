VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaxBill"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim strLeftNo As String         '发票前缀
Dim strItems As String          '收据费目的对照字符串
Dim aryItem() As String         '收据费目的对照数组

Dim rsTemp As New ADODB.Recordset
Dim intCount As Integer

Public Function zlTaxUseable(ByVal int票种 As Integer) As Boolean
    '------------------------------------------------------
    '功能：检查本机是否使用税控打印
    '参数：int票种=1，门诊收费票据；int票种=2，住院的票据
    '------------------------------------------------------
    If int票种 = 1 Then
        zlTaxUseable = Val(GetSetting("ZLSOFT", "公共全局\税票打印", "门诊使用税票打印", 0)) <> 0
    Else
        zlTaxUseable = Val(GetSetting("ZLSOFT", "公共全局\税票打印", "住院使用税票打印", 0)) <> 0
    End If
End Function

Public Function zlTaxBillSet(ByVal cnMain As ADODB.Connection, int票种 As Integer) As Boolean
    '------------------------------------------------------
    '功能：设置指定票据的本地参数
    '参数：int票种=1，门诊收费票据；int票种=2，住院的票据
    '------------------------------------------------------
    Set gcnOracle = cnMain
    Select Case int票种
        Case 1
            frmOutSet.Show 1
        Case 2
            frmInSet.Show 1
    End Select
End Function

Public Function zlTaxOutPrint(cnMain As ADODB.Connection, strNOs As String, Optional strExpend As String) As String
    '------------------------------------------------------
    '功能：门诊收据打印(包括部分退费后的再打印)
    '参数：strNOs，本次打印的单据号码字符串
    '      strExpend,用|线分隔,第一位为单据性质(1-收费 2-挂号),第二位后待以后扩展
    '返回：零长度字符串，表示成功；其他表示发生错误的原因
    '------------------------------------------------------
    Dim strUnions As String, strDecode As String
    Dim bln密码器 As Boolean, strMoneys As String
    Dim lngReturn As Long, strError As String * 255
    Dim strName As String, strInfo As String, intType As Integer
    
    Set gcnOracle = cnMain
    
    bln密码器 = Val(GetSetting("ZLSOFT", "公共全局\税票打印", "门诊使用密码器", 0)) <> 0
    strLeftNo = GetSetting("ZLSOFT", "公共全局\税票打印", "门诊税票前缀", "")
    strItems = GetSetting("ZLSOFT", "公共全局\税票打印", "门诊税票项目", "")
    intType = Val(Split(strExpend & "|", "|")(0))
    If intType = 0 Then intType = 1
    aryItem = Split(strItems, "|")
    
    strDecode = ""
    For intCount = LBound(aryItem) To UBound(aryItem)
        strDecode = strDecode & ",'" & Split(aryItem(intCount), ";")(0) & "','" & Split(aryItem(intCount), ";")(1) & "'"
    Next
    
    strUnions = ""
    For intCount = 1 To 19
        strUnions = strUnions & " union all select '" & Format(intCount, "00") & "' as 发票项目,0 as 金额 From Dual"
    Next
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select 姓名 From 门诊费用记录 Where 记录性质 = " & IIf(intType = 1, "1", "4") & " And NO IN(" & strNOs & ") And Rownum=1"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "门诊税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If Not .EOF Then strName = IIf(IsNull(!姓名), "", !姓名)
        
        gstrSql = "Select 发票项目,sum(金额) as 金额" & _
                " From (Select decode(收据费目" & strDecode & ",null) as 发票项目,实收金额 as 金额" & _
                "       From 门诊费用记录" & _
                "       Where 记录性质 = " & IIf(intType = 1, "1", "4") & " and NO In (" & strNOs & ")" & _
                        strUnions & ")" & _
                " Where 发票项目 is not null" & _
                " Group By 发票项目"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "门诊税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        strMoneys = ""
        Do While Not .EOF
            strMoneys = strMoneys & ";" & Format(!金额, "0.00")
            .MoveNext
        Loop
        strMoneys = Mid(strMoneys, 2)
        strMoneys = strMoneys & ";;0.00;;0.00;;0.00;;0.00;;0.00;" '5个自定名称项目
        
        gstrSql = "Select Max(ID) as ID From 票据打印内容 Where 数据性质 = " & IIf(intType = 1, "1", "4") & " And NO IN(" & strNOs & ")"
        gstrSql = "Select Distinct 号码,使用人 From 票据使用明细" & _
                " Where 打印ID=(" & gstrSql & ") Order by 号码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "门诊税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxOutPrint = "未找到需要传递的票据": Exit Function
        If .RecordCount > 1 Then zlTaxOutPrint = "需要传递的票据不是同一票据": Exit Function
    
ReTry:
        strError = String(255, " ")
        If bln密码器 Then
            lngReturn = gobjTax.BJ_Normal_Invoice(2, strLeftNo & !号码, strName, !使用人, strMoneys, strError)
        Else
            lngReturn = gobjTax.BJ_Normal_Invoice_NoPwd(2, strLeftNo & !号码, strName, !使用人, strMoneys, strError)
        End If
        If lngReturn <> 0 Then
            strInfo = "税控器接口调用失败(开票)：" & vbCrLf & vbCrLf & _
                "错误编号：" & lngReturn & vbCrLf & "错误信息：" & zStr(strError) & vbCrLf & vbCrLf & _
                "请检查设备是否连接正常，打印机纸张是否安放正确。" & vbCrLf & _
                "如果是上述原因，你可以选择重试。要重试吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxOutPrint = "" '已提示,调用程序不用再提示
                Exit Function
            End If
        End If
        zlTaxOutPrint = ""
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlTaxOutErase(cnMain As ADODB.Connection, strNOs As String, Optional strExpend As String) As String
    '------------------------------------------------------
    '功能：门诊收据作废
    '参数：strNo，本次作废的单据号码字符串
    '返回：零长度字符串，表示成功；其他表示发生错误的原因
    '------------------------------------------------------
    Dim lngReturn As Long, strError As String * 255
    Dim strInfo As String, intType As Integer
    
    Set gcnOracle = cnMain
    strLeftNo = GetSetting("ZLSOFT", "公共全局\税票打印", "门诊税票前缀", "")
    
    intType = Val(Split(strExpend & "|", "|")(0))
    If intType = 0 Then intType = 1
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
'        gstrSql = "Select Max(ID) as ID From 票据打印内容 Where 数据性质=1 And NO IN(" & strNOs & ")"
'        gstrSql = "Select Distinct 号码,使用人 From 票据使用明细" & _
'                " Where 打印ID=(" & gstrSql & ") Order by 号码"
        
        '由于部份退费再打印及多单据修改的情况,可能该作废单据已重新发出票据。
        gstrSql = "Select 号码,使用人" & _
                 " From 票据使用明细" & _
                 " Where Id = (" & _
                 "       Select Max(Id)" & _
                 "       From 票据使用明细" & _
                 "       Where 性质 = 2 And 打印id In (" & _
                 "             Select Id From 票据打印内容 Where 数据性质 = " & IIf(intType = 1, "1", "4") & " And No In (" & strNOs & ")))"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "门诊税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxOutErase = "未找到需要作废的票据": Exit Function
        If .RecordCount > 1 Then zlTaxOutErase = "需要作废的票据不是同一票据": Exit Function
        
ReTry:
        strError = String(255, " ")
        lngReturn = gobjTax.BJ_Other_Invoice(1, 2, strLeftNo & !号码, !使用人, "", strError)
        If lngReturn <> 0 Then
            strInfo = "税控器接口调用失败(退票)：" & vbCrLf & vbCrLf & _
                "错误编号：" & lngReturn & vbCrLf & "错误信息：" & zStr(strError) & vbCrLf & vbCrLf & _
                "请检查设备是否连接正常。" & vbCrLf & _
                "如果是上述原因，你可以选择重试。要重试吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxOutErase = "" '已提示,调用程序不用再提示
                Exit Function
            End If
        End If
        zlTaxOutErase = ""
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlTaxOutReput(cnMain As ADODB.Connection, strNOs As String, Optional strExpend As String) As String
    '------------------------------------------------------
    '功能：门诊收据重新打印
    '参数：strNo，本次打印的单据号码字符串
    '返回：零长度字符串，表示成功；其他表示发生错误的原因
    '------------------------------------------------------
    Dim lngReturn As Long, strError As String * 255
    Dim strOldInvioce As String, strInfo As String, intType As Integer
    
    intType = Val(Split(strExpend & "|", "|")(0))
    If intType = 0 Then intType = 1
    Set gcnOracle = cnMain
    strLeftNo = GetSetting("ZLSOFT", "公共全局\税票打印", "门诊税票前缀", "")
    
    '由于调用时，数据已经被更新，所以需要获取重打前的实际票据号码，进行作废操作
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select 号码,使用人" & _
                 " From 票据使用明细" & _
                 " Where Id = (" & _
                 "       Select Max(Id)" & _
                 "       From 票据使用明细" & _
                 "       Where 性质 = 2 And 打印id In (" & _
                 "             Select Id From 票据打印内容 Where 数据性质 = " & IIf(intType = 1, "1", "4") & " And No In (" & strNOs & ")))"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "门诊税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxOutReput = "未找到上次的打印记录": Exit Function
        If .RecordCount > 1 Then zlTaxOutReput = "需要重打的票据不是同一票据": Exit Function
        
ReTry:
        strError = String(255, " ")
        lngReturn = gobjTax.BJ_Other_Invoice(1, 2, strLeftNo & !号码, !使用人, "", strError)
        If lngReturn <> 0 Then
            strInfo = "税控器接口调用失败(退票)：" & vbCrLf & vbCrLf & _
                "错误编号：" & lngReturn & vbCrLf & "错误信息：" & zStr(strError) & vbCrLf & vbCrLf & _
                "请检查设备是否连接正常。" & vbCrLf & _
                "如果是上述原因，你可以选择重试。要重试吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxOutReput = "" '已提示,调用程序不用再提示
                Exit Function
            End If
        End If
    End With
    
    '调用打印操作，实现发票的重打操作
    zlTaxOutReput = zlTaxOutPrint(cnMain, strNOs)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlTaxInPrint(cnMain As ADODB.Connection, lngId As Long) As String
    '------------------------------------------------------
    '功能：住院收据打印
    '参数：lngId，本次打印的结帐单的id
    '返回：零长度字符串，表示成功；其他表示发生错误的原因
    '------------------------------------------------------
    Dim strDecode As String, bln密码器 As Boolean
    Dim strMoneys As String, intTimes As Integer
    Dim lngReturn As Long, strError As String * 255
    Dim strName As String, strInfo As String
    
    Set gcnOracle = cnMain
    
    bln密码器 = Val(GetSetting("ZLSOFT", "公共全局\税票打印", "住院使用密码器", 0)) <> 0
    strLeftNo = GetSetting("ZLSOFT", "公共全局\税票打印", "住院税票前缀", "")
    strItems = GetSetting("ZLSOFT", "公共全局\税票打印", "住院税票项目", "")
    aryItem = Split(strItems, "|")
    
    strDecode = ""
    For intCount = LBound(aryItem) To UBound(aryItem)
        strDecode = strDecode & ",'" & Split(aryItem(intCount), ";")(0) & "','" & Split(aryItem(intCount), ";")(1) & "'"
    Next
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select B.姓名 From 病人信息 B,病人结帐记录 A Where A.病人ID=B.病人ID And A.ID=" & lngId
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "门诊税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If Not .EOF Then strName = IIf(IsNull(!姓名), "", !姓名)
    
        gstrSql = "Select 发票项目,sum(金额) as 金额" & _
                " From (Select decode(收据费目" & strDecode & ",null) as 发票项目,结帐金额 as 金额" & _
                "       From 住院费用记录" & _
                "       Where 记帐费用=1 and 结帐id=" & lngId & ")" & _
                " Where 发票项目 is not null" & _
                " Group By 发票项目" & _
                " Having sum(金额)<>0"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "住院税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        strMoneys = "": intTimes = 0
        Do While Not .EOF
            intTimes = intTimes + 1
            strMoneys = strMoneys & ";" & Mid(!发票项目, 1, InStr(1, !发票项目, "-") - 1) & ";" & _
                        Mid(!发票项目, InStr(1, !发票项目, "-") + 1) & ";1;" & Format(!金额, "0.00") & ";"
            If intTimes = 4 Then Exit Do
            .MoveNext
        Loop
        strMoneys = intTimes & strMoneys
        
        gstrSql = "Select L.实际票号,I.住院号,L.操作员姓名" & _
                " From 病人结帐记录 L,病人信息 I" & _
                " Where L.病人id=I.病人id And L.id=" & lngId
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "住院税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxInPrint = "未找到需要传递的票据": Exit Function
        If .RecordCount > 1 Then zlTaxInPrint = "需要传递的票据不是同一票据": Exit Function
        
ReTry:
        strError = String(255, " ")
        If bln密码器 Then
            lngReturn = gobjTax.BJ_Normal_Invoice(1, strLeftNo & !实际票号, strName, !操作员姓名, strMoneys & ";" & !住院号 & ";", strError)
        Else
            lngReturn = gobjTax.BJ_Normal_Invoice_NoPwd(1, strLeftNo & !实际票号, strName, !操作员姓名, strMoneys & ";" & !住院号 & ";", strError)
        End If
        If lngReturn <> 0 Then
            strInfo = "税控器接口调用失败(开票)：" & vbCrLf & vbCrLf & _
                "错误编号：" & lngReturn & vbCrLf & "错误信息：" & zStr(strError) & vbCrLf & vbCrLf & _
                "请检查设备是否连接正常，打印机纸张是否安放正确。" & vbCrLf & _
                "如果是上述原因，你可以选择重试。要重试吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxInPrint = "" '已提示,调用程序不用再提示
                Exit Function
            End If
        End If
        zlTaxInPrint = ""
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlTaxInErase(cnMain As ADODB.Connection, lngId As Long) As String
    '------------------------------------------------------
    '功能：住院收据作废
    '参数：lngId，本次作废的结帐单的id
    '返回：零长度字符串，表示成功；其他表示发生错误的原因
    '------------------------------------------------------
    Dim lngReturn As Long, strError As String * 255
    Dim strInfo As String
    
    Set gcnOracle = cnMain
    strLeftNo = GetSetting("ZLSOFT", "公共全局\税票打印", "住院税票前缀", "")
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select L.实际票号,L.操作员姓名" & _
                " From 病人结帐记录 L Where L.id=" & lngId
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "住院税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxInErase = "未找到需要作废的票据": Exit Function
        If .RecordCount > 1 Then zlTaxInErase = "需要作废的票据不是同一票据": Exit Function
    
ReTry:
        strError = String(255, " ")
        lngReturn = gobjTax.BJ_Other_Invoice(1, 1, strLeftNo & !实际票号, !操作员姓名, "", strError)
        If lngReturn <> 0 Then
            strInfo = "税控器接口调用失败(退票)：" & vbCrLf & vbCrLf & _
                "错误编号：" & lngReturn & vbCrLf & "错误信息：" & zStr(strError) & vbCrLf & vbCrLf & _
                "请检查设备是否连接正常。" & vbCrLf & _
                "如果是上述原因，你可以选择重试。要重试吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxInErase = "" '已提示,调用程序不用再提示
                Exit Function
            End If
        End If
        zlTaxInErase = ""
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlTaxInReput(cnMain As ADODB.Connection, lngId As Long) As String
    '------------------------------------------------------
    '功能：住院收据重新打印
    '参数：lngId，本次打印的结帐单的id
    '返回：零长度字符串，表示成功；其他表示发生错误的原因
    '------------------------------------------------------
    Dim lngReturn As Long, strError As String * 255
    Dim strOldInvioce As String, strInfo As String
    
    Set gcnOracle = cnMain
    strLeftNo = GetSetting("ZLSOFT", "公共全局\税票打印", "住院税票前缀", "")
    
    '由于调用时，数据已经被更新，所以需要获取重打前的实际票据号码，进行作废操作
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "Select 号码,使用人" & _
                 " From 票据使用明细" & _
                 " Where Id In (" & _
                 "       Select Max(Id)" & _
                 "       From 票据使用明细" & _
                 "       Where 性质 = 2 And 打印id In (" & _
                 "             Select Id From 票据打印内容 Where 数据性质=3 And No In (" & _
                 "             Select No From 病人结帐记录 Where ID=" & lngId & ")))"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "住院税控", gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        If .RecordCount <= 0 Then zlTaxInReput = "未找到上次的打印记录": Exit Function
        If .RecordCount > 1 Then zlTaxInReput = "需要重打的票据不是同一票据": Exit Function
                
ReTry:
        strError = String(255, " ")
        lngReturn = gobjTax.BJ_Other_Invoice(1, 1, strLeftNo & !号码, !使用人, "", strError)
        If lngReturn <> 0 Then
            strInfo = "税控器接口调用失败(退票)：" & vbCrLf & vbCrLf & _
                "错误编号：" & lngReturn & vbCrLf & "错误信息：" & zStr(strError) & vbCrLf & vbCrLf & _
                "请检查设备是否连接正常。" & vbCrLf & _
                "如果是上述原因，你可以选择重试。要重试吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo ReTry
            Else
                zlTaxInReput = "" '已提示,调用程序不用再提示
                Exit Function
            End If
        End If
    End With
    
    '调用打印操作，实现发票的重打操作
    zlTaxInReput = zlTaxInPrint(cnMain, lngId)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Class_Initialize()
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    gstrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
End Sub
