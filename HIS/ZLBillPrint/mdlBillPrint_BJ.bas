Attribute VB_Name = "mdlBillPrint_BJ"
Option Explicit


Public Function Init() As Boolean
'功能：进行第三方票据打印接口的初始化或登录等调用
'返回：执行成功/失败



    '参考：通过注册表读写临时记录数据
    'Call SaveSetting("ZLSOFT", "公共全局\票据打印", "当前票据号", 'XXXX')
    'Call GetSetting("ZLSOFT", "公共全局\票据打印", "当前票据号", "")
    
    Init = True
End Function

Public Function Term() As Boolean
'功能：完成第三方票据打印接口的资源释放、断开连接等调用
'返回：执行成功/失败
    
    
    Term = True
End Function


Public Function SYSConfigure() As Boolean
'功能：参数设置,在HIS"模块参数设置"(文件/参数设置)中调用，可在本接口中完成第三方票据打印接口的参数设置、配置更改等调用
'返回：执行成功/失败
    
    
    SYSConfigure = True
End Function

Public Function DiscardBill(ByVal lng领用ID As Long, ByVal lng票种 As Long, ByVal str票号前缀 As String, _
    ByVal str开始票号 As String, ByVal str结束票号 As String, ByVal DateAdd As Date, ByVal str报损人 As String) As Boolean
'功能：票据报损
'返回：执行成功/失败

    DiscardBill = True
End Function

Public Function PrintBillOut(ByVal strNOs As String) As Boolean
'功能：门诊收费票据打印
'参数：strNOs=门诊收费：以逗号分隔的带引号的多个单据号(一次打印单张或多张单据):'F0000001','F0000002',...
'返回：执行成功/失败
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
   '参考：读取单据号相关费用数据
   '使用f_Str2list(为了使用绑定变量,zltools下提供的将字符串转换为临时内存表的函数),
   '需要在SQL语句中第一个Select关键字后加入“/*+ Rule*/”提示，因为Cbo下临时内存表没有统计数据，否则会导致大表全表扫描
'   strSQL = "Select/*+ Rule*/ 收据费目 as 发票项目,Sum(实收金额) as 金额" & _
'            " From 门诊费用记录" & _
'            " Where 记录性质=1 and NO In (Select * From Table(f_Str2list([1])))" & _
'            " Group By 收据费目"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "获取打印内容", Replace(strNOs, "'", ""))
'    If rstmp.RecordCount = 0 Then
'        Exit Function
'    End If

    

    PrintBillOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PrintBillIn(ByVal lngBalanceId As Long) As Boolean
'功能：住院结帐票据打印
'参数：lngBalanceId=结帐单ID
'返回：执行成功/失败
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
'   strSQL = "Select/*+ Rule*/ 收据费目 as 发票项目,Sum(结帐金额) as 金额" & _
'            " From 住院费用记录" & _
'            " Where 结帐id=[1]" & _
'            " Group By 收据费目"
'   Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "获取打印内容", lngBalanceId)

'   strSQL = "Select L.实际票号,I.住院号,L.操作员姓名" & _
'            " From 病人结帐记录 L,病人信息 I" & _
'            " Where L.病人id=I.病人id And L.id=[1]"


    PrintBillIn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function RePrintBillOut(ByVal strNOs As String, ByVal strInvoice As String) As Boolean
'功能：重打门诊收费票据
'参数：strNOs=门诊收费：以逗号分隔的带引号的多个单据号(一次打印单张或多张单据):'F0000001','F0000002',...
'      strInvoice=本次重打使用的起始票据号
'返回：执行成功/失败
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
'    strSQL = "Select 号码,使用人" & _
'                 " From 票据使用明细" & _
'                 " Where Id = (" & _
'                 "       Select Max(Id)" & _
'                 "       From 票据使用明细" & _
'                 "       Where 性质 = 2 And 打印id In (" & _
'                 "             Select Id From 票据打印内容 Where 数据性质=1 And No In (Select * From Table(f_Str2list([1])))))"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "获取票据号", replace(strNOs,"'",""))

    RePrintBillOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function RePrintBillIn(ByVal lngBalanceId As Long, ByVal strInvoice As String) As Boolean
'功能：重打住院结帐票据
'参数：lngBalanceId=结帐单ID
'      strInvoice=本次重打使用的起始票据号
'返回：执行成功/失败
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
'    strSQL = "Select 号码,使用人" & _
'                 " From 票据使用明细" & _
'                 " Where Id In (" & _
'                 "       Select Max(Id)" & _
'                 "       From 票据使用明细" & _
'                 "       Where 性质 = 2 And 打印id In (" & _
'                 "             Select Id From 票据打印内容 Where 数据性质=3 And No In (" & _
'                 "             Select No From 病人结帐记录 Where ID=[1])))"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "获取票据号", lngBalanceId)
                 
    RePrintBillIn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function EraseBillOut(ByVal strNOs As String) As Boolean
'功能：作废门诊收费票据
'参数：strNOs=门诊收费：以逗号分隔的带引号的多个单据号(一次打印单张或多张单据):'F0000001','F0000002',...
'返回：执行成功/失败
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH

'   由于部份退费再打印及多单据修改的情况,可能该作废单据已重新发出票据，所以要用max(id)
'    strSQL = "Select/*+ Rule*/ 号码,使用人" & _
'                 " From 票据使用明细" & _
'                 " Where Id = (" & _
'                 "       Select Max(Id)" & _
'                 "       From 票据使用明细" & _
'                 "       Where 性质 = 2 And 打印id In (" & _
'                 "             Select Id From 票据打印内容 Where 数据性质=1 And No In (Select * From Table(f_Str2list([1]))))"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "获取票据号", replace(strNOs,"'",""))
    
    EraseBillOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function EraseBillIn(ByVal lngBalanceId As Long) As Boolean
'功能：作废住院结帐票据
'参数：lngBalanceId=结帐单ID
'返回：执行成功/失败
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
'    strSQL = "Select L.实际票号,L.操作员姓名" & _
'             " From 病人结帐记录 L Where L.id=[1]"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "获取票据号", lngBalanceId)
                
    EraseBillIn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
