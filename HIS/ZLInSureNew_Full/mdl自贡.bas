Attribute VB_Name = "mdl自贡"
Option Explicit
Private Declare Function zg_GetErrToStr Lib "DataExchange" Alias "GetErrToStr" (ByVal lngErr As Long) As String
Private Declare Function zg_ReadICCardInfo Lib "DataExchange" Alias "ReadICCardInfo" (ByVal strPass As String) As String
Private Declare Function zg_ChangePassword Lib "DataExchange" Alias "DllChangePassWord" (ByVal strPass As String, ByVal strNewPass As String) As String
Private Declare Function zg_ClinicCharge Lib "DataExchange" Alias "ClinicCharge" _
    (ByVal strPass As String, ByVal strClinicBalanceNO As String, ByVal blnIsPrepare As Long) As String
Private Declare Function zg_InHosReg Lib "DataExchange" Alias "InHosReg" _
    (ByVal strPass As String, ByVal strInHosRegisterNO As String) As Long
Private Declare Function zg_UnInHosReg Lib "DataExchange" Alias "UnInHosReg" _
    (ByVal strPass As String, ByVal strInHosRegisterNO As String) As Long
Private Declare Function zg_PreInHosBalance Lib "DataExchange" Alias "PreInHosBalance" _
    (ByVal strPass As String, ByVal strInHosRegisterNO As String, ByVal lngBalanceType As Long, _
    Optional ByVal strCheckupKind As String = "3", _
    Optional ByVal strSickKindCode As String = "0", _
    Optional ByVal intAccount As Long = 0) As String
Private Declare Function zg_InHosBalance Lib "DataExchange" Alias "InHosBalance" _
    (ByVal strPass As String, ByVal strInHosBalanceNO As String, ByVal lngBalanceType As Long, _
    Optional ByVal strCheckupKind As String = "3", _
    Optional ByVal strSickKindCode As String = "0", _
    Optional ByVal intAccount As Long = 0) As String
Private Declare Function zg_UnInHosBalance Lib "DataExchange" Alias "UnInHosBalance" _
    (ByVal strPass As String, ByVal strInHosBalanceNO As String, ByVal strUnInHosBalanceNo As String) As String

'Public gobj自贡 As New clsZGYB              '调试用
Public mblnInit As Boolean
Public gcn自贡 As New ADODB.Connection

Public Enum 业务类型_自贡
    读卡
    修改密码
    门诊预算
    门诊结算
    入院登记
    撤销入院登记
    住院预结算
    住院结算
    住院结算冲销
End Enum

Private Type 结算信息_自贡
    病人ID As Long
    总金额 As Currency
    全自费 As Currency
    首先自付 As Currency
    进入统筹 As Currency
    本次起付线 As Currency
    实际起付线 As Currency
    统筹支付 As Currency
    超封顶费用 As Currency
    个人帐户 As Currency
    现金支付 As Currency
End Type
Private cur_结算信息 As 结算信息_自贡
Private gstr结算号 As String    '保存单据号

Public Function 医保初始化_自贡() As Boolean
    If mblnInit Then
        医保初始化_自贡 = True
        Exit Function
    End If
    医保初始化_自贡 = 检查医保服务器_自贡
End Function

Private Function 检查医保服务器_自贡() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If gcn中软.State = adStateOpen Then
        检查医保服务器_自贡 = True
        Exit Function
    End If
    
    '读出连接医保服务器的配置
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", TYPE_四川自贡)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    
    If OraDataOpen(gcn自贡, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接医保前置机！", vbInformation, gstrSysName
        Exit Function
    End If
    
    检查医保服务器_自贡 = True
End Function

Public Function 调用接口_自贡(ByVal intType As Integer, ByVal StrInput As String, strOutput As String) As Boolean
    Dim lngErr As Long
    Dim arrPara             '用来分解入参，当某个函数需要多个入参时
    
    Select Case intType
    Case 业务类型_自贡.读卡
        Call WriteBusinessLOG("ReadICCardInfo", "调用前输出", "")
'        strOutPut = gobj自贡.ReadICCardInfo(strInput)
        strOutput = zg_ReadICCardInfo(StrInput)
        Call WriteBusinessLOG("ReadICCardInfo", "", strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case 业务类型_自贡.修改密码
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("DllChangePassWord", "调用前输出", "")
'        strOutPut = gobj自贡.ChangePass(arrPara(0), arrPara(1))
        strOutput = zg_ChangePassword(arrPara(0), arrPara(1))
        Call WriteBusinessLOG("DllChangePassWord", arrPara(0) & "," & arrPara(1), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case 业务类型_自贡.门诊预算, 业务类型_自贡.门诊结算
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("ClinicCharge", "调用前输出", "")
'        strOutPut = gobj自贡.ClinicCharge(arrPara(0), arrPara(1), (Val(arrPara(2)) = 1))
        strOutput = zg_ClinicCharge(arrPara(0), arrPara(1), Val(arrPara(2)))
        Call WriteBusinessLOG("ClinicCharge", arrPara(0) & "," & arrPara(1) & "," & Val(arrPara(2)), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case 业务类型_自贡.入院登记
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("InHosReg", "调用前输出", "")
'        lngErr = gobj自贡.InHosReg(arrPara(0), arrPara(1))
        lngErr = zg_InHosReg(arrPara(0), arrPara(1))
        Call WriteBusinessLOG("InHosReg", arrPara(0) & "," & arrPara(1), lngErr)
    Case 业务类型_自贡.撤销入院登记
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("UnInHosReg", "调用前输出", "")
'        lngErr = gobj自贡.UnInHosReg(arrPara(0), arrPara(1))
        lngErr = zg_UnInHosReg(arrPara(0), arrPara(1))
        Call WriteBusinessLOG("UnInHosReg", arrPara(0) & "," & arrPara(1), lngErr)
    Case 业务类型_自贡.住院预结算
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("PreInHosBalance", "调用前输出", "")
'        strOutPut = gobj自贡.PreInHosBalance(arrPara(0), arrPara(1), Val(arrPara(2)), arrPara(3), arrPara(4), CLng(arrPara(5)))
        strOutput = zg_PreInHosBalance(arrPara(0), arrPara(1), Val(arrPara(2)), arrPara(3), arrPara(4), CLng(arrPara(5)))
        Call WriteBusinessLOG("PreInHosBalance", arrPara(0) & "," & arrPara(1) & "," & Val(arrPara(2)) & "," & arrPara(3) & "," & arrPara(4) & "," & CLng(arrPara(5)), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case 业务类型_自贡.住院结算
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("InHosBalance", "调用前输出", "")
'        strOutPut = gobj自贡.InHosBalance(arrPara(0), arrPara(1), Val(arrPara(2)), arrPara(3), arrPara(4), CLng(arrPara(5)))
        strOutput = zg_InHosBalance(arrPara(0), arrPara(1), Val(arrPara(2)), arrPara(3), arrPara(4), CLng(arrPara(5)))
        Call WriteBusinessLOG("InHosBalance", arrPara(0) & "," & arrPara(1) & "," & Val(arrPara(2)) & "," & arrPara(3) & "," & arrPara(4) & "," & CLng(arrPara(5)), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case 业务类型_自贡.住院结算冲销
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("UnInHosBalance", "调用前输出", "")
'        strOutPut = gobj自贡.UnInHosBalance(arrPara(0), arrPara(1), arrPara(2))
        strOutput = zg_UnInHosBalance(arrPara(0), arrPara(1), arrPara(2))
        Call WriteBusinessLOG("UnInHosBalance", arrPara(0) & "," & arrPara(1) & "," & arrPara(2), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    End Select
    
    '判断是否发生错误
    If lngErr <> 0 Then
'        MsgBox "中软医保接口返回错误，详细信息如下：" & vbCrLf & _
'            gobj自贡.GetErrToStr(lngErr), vbInformation, gstrSysName
        MsgBox "中软医保接口返回错误，详细信息如下：" & vbCrLf & _
            zg_GetErrToStr(lngErr), vbInformation, gstrSysName
        strOutput = ""
        Exit Function
    End If
    
    If Not (intType = 业务类型_自贡.入院登记 Or intType = 业务类型_自贡.撤销入院登记) Then
        If strOutput = "" Then
            MsgBox "接口返回的数据不正确！返回串为空", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If InStr(1, strOutput, "|*") <> 0 Then
        strOutput = Mid(strOutput, 7)
        strOutput = Mid(strOutput, 1, InStr(1, strOutput, "|*") - 1)
    End If
    调用接口_自贡 = True
End Function

Public Function 身份标识_自贡(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strSelfNO-个人编号，刷卡得到；strSelfPwd-病人密码；bytType-识别类型，0-门诊，1-住院
'返回： 空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim strReturn As String
    On Error GoTo errHandle
    
    strReturn = frmIdentify自贡.GetPatient(bytType, lng病人ID, (bytType <> 2), True)
    If strReturn = "" Then Exit Function
    
    身份标识_自贡 = strReturn
    gstr结算号 = ""          '每次预结算时单独从部门表序列中获取序列值，做为门诊虚拟结算的处方号
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_自贡(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim strReturn As String, str发生日期 As String
    Dim arrBalance
    Dim strInsert As String, strValue As String
    Dim str开单部门 As String, str医生姓名 As String, str病种编码 As String, str病种名称 As String, str病种类型 As String
    Dim str医保项目编码 As String, str医院项目名称 As String
    Dim rsTemp As New ADODB.Recordset
    
    Const int总金额 As Integer = 0
    Const int帐户支付 As Integer = 1
    Const int现金支付 As Integer = 2
    Const int统筹支付 As Integer = 3
    On Error GoTo errHand
    
    cur_结算信息.病人ID = rs明细!病人ID
    str发生日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
    gstr结算号 = zlDatabase.GetNextID("部门表")
    
    '重新插入本次预结算的处方明细
    '处方主表(门诊收据号，划价号，门诊收据号-外部，开单科室，医生姓名，病种编码，病种名称，病种类型代码，借贷标志，发生日期)
    strInsert = " Insert Into ClinicBill " & _
                " (ClinicBillNO,ClinicBalanceNO,InvoiceNO,DepartmentName,DoctorName,SickSerialNO,SickName,SickKindCode,RedBillFlag,OccurDate) " & _
                " Values ("
    
    '提取开单部门名称
    gstrSQL = "Select 名称 From 部门表 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取开单部门名称", CLng(rs明细!开单部门ID))
    str开单部门 = Nvl(rsTemp!名称)
    str医生姓名 = Nvl(rs明细!开单人, "HIS")
    
    'todo:门诊是否需要选择病种?此处用来提取病种信息
'    gstrSQL = ""
'    Call OpenRecordset(rsTemp, "提取病种信息")
'    str病种编码 = "001"
'    str病种名称 = "普通病"
'    str病种类型 = "1"
    
    On Error Resume Next
    strValue = "'" & gstr结算号 & "','" & gstr结算号 & "','" & gstr结算号 & "','" & str开单部门 & "','" & str医生姓名 & "'," & _
        Val(str病种编码) & ",'" & str病种名称 & "','" & str病种类型 & "',1,to_Date('" & str发生日期 & "','yyyy-MM-dd hh24:mi:ss')"
    gstrSQL = strInsert & strValue & ")"
    gcn自贡.Execute gstrSQL
    On Error GoTo errHand
    
    With rs明细
        '处方明细表(门诊收费明细流水号，门诊收据号，收费项目编码，医院收费项目名称，单价，数量，金额)
        strInsert = " Insert Into ClinicBillDetail" & _
                    " (ClinicBillDetailNO,ClinicBillNO,ItemNO,HosItemName,Price,Quantity,Amount)" & _
                    " Values ("
        Do While Not .EOF
            '提取项目的医保编码
            gstrSQL = " Select A.项目编码,B.名称 From 保险支付项目 A,收费细目 B" & _
                      " Where A.收费细目ID=B.ID And A.险类=[1] And 收费细目ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目的医保编码", TYPE_四川自贡, CLng(!收费细目ID))
            str医保项目编码 = Nvl(rsTemp!项目编码)
            str医院项目名称 = Nvl(rsTemp!名称)
            
            '处方明细表
            strValue = "'" & .AbsolutePosition & "','" & gstr结算号 & "','" & str医保项目编码 & "','" & str医院项目名称 & "'," & _
                       Val(Format(!实收金额 / !数量, "#####0.0000")) & "," & !数量 & "," & Nvl(!实收金额, 0)
            gstrSQL = strInsert & strValue & ")"
            gcn自贡.Execute gstrSQL
            .MoveNext
        Loop
    End With
    
    '调用门诊虚拟结算接口，返回值格式：本次门诊总金额|个人帐户支付|现金支付|统筹支付|个人帐户余额|姓名|医保帐号|卡号
    If Not 调用接口_自贡(业务类型_自贡.门诊预算, GetPass(cur_结算信息.病人ID) & "|" & gstr结算号 & "|" & 1, strReturn) Then Exit Function
    
    arrBalance = Split(strReturn, "|")
    cur_结算信息.总金额 = Val(arrBalance(int总金额))
    cur_结算信息.个人帐户 = Val(arrBalance(int帐户支付))
    cur_结算信息.现金支付 = Val(arrBalance(int现金支付))
    cur_结算信息.统筹支付 = Val(arrBalance(int统筹支付))
    
    '返回结算串
    If cur_结算信息.个人帐户 <> 0 Then str结算方式 = str结算方式 & "|个人帐户;" & cur_结算信息.个人帐户 & ";0"
    If cur_结算信息.统筹支付 <> 0 Then str结算方式 = str结算方式 & "|医保基金;" & cur_结算信息.统筹支付 & ";0"
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    门诊虚拟结算_自贡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 就诊登记取消_自贡() As Boolean
    On Error GoTo errHand
    
    就诊登记取消_自贡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_自贡(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    Dim strReturn As String
    Dim arrBalance
    Dim rsTemp As New ADODB.Recordset
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    
    Const int总金额 As Integer = 0
    Const int帐户支付 As Integer = 1
    Const int现金支付 As Integer = 2
    Const int统筹支付 As Integer = 3
    On Error GoTo errHand
    '--医保中心不关心发票号--
'    '提取本次结算发票号（中心并不关心发票号）
'    gstrSQL = "Select 实际票号 AS 发票号 From 病人费用记录 Where 结帐ID=" & lng结帐ID & " And Rownum<2"
'    Call OpenRecordset(rsTemp, "提取本次结算发票号")
'    strBalanceNO = Nvl(rsTemp!发票号)
'    If strBalanceNO = "" Then
'        MsgBox "没有发票号不能进行结算，请指定发票号！", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '因预结算时不知道处方号，此处更新处方主表与明细表的处方号与收据号（中心并不关心发票号，上面取消所以此处也不使用了）
'    gstrSQL = "Update ClinicBill Set ClinicBillNO='" & strBalanceNO & "',ClinicBalanceNO='" & strBalanceNO & "',InvoiceNO='" & strBalanceNO & "' Where ClinicBalanceNO='" & gstr结算号 & "'"
'    gcn自贡.Execute gstrSQL
'    gstrSQL = "Update ClinicBillDetail Set ClinicBillNO='" & strBalanceNO & "' Where ClinicBillNO='" & gstr结算号 & "'"
'    gcn自贡.Execute gstrSQL
    '------------------------
    
    '调用门诊结算接口
    If Not 调用接口_自贡(业务类型_自贡.门诊结算, GetPass(cur_结算信息.病人ID) & "|" & gstr结算号 & "|" & 0, strReturn) Then Exit Function
    
    arrBalance = Split(strReturn, "|")
    cur_结算信息.总金额 = Val(arrBalance(int总金额))
    cur_结算信息.个人帐户 = Val(arrBalance(int帐户支付))
    cur_结算信息.现金支付 = Val(arrBalance(int现金支付))
    cur_结算信息.统筹支付 = Val(arrBalance(int统筹支付))
   
    Call Get帐户信息(TYPE_四川自贡, cur_结算信息.病人ID, Year(zlDatabase.Currentdate()), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & cur_结算信息.病人ID & "," & TYPE_四川自贡 & "," & Year(zlDatabase.Currentdate()) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 + cur_结算信息.进入统筹 & "," & _
        cur统筹报销累计 + cur_结算信息.统筹支付 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    'g结算数据.超限自付金额中保存的是门诊病人就诊类型（急诊、特殊病门诊或普通门诊），结算记录的备注保存的是病种的名称
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)'超限自付金额用于暂时保存，门诊类别
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_四川自贡 & "," & cur_结算信息.病人ID & "," & _
        Year(zlDatabase.Currentdate()) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 + cur_结算信息.进入统筹 & "," & _
        cur统筹报销累计 + cur_结算信息.统筹支付 & "," & int住院次数累计 & ",0,0,0," & cur_结算信息.总金额 & ",0,0," & _
        cur_结算信息.进入统筹 & "," & cur_结算信息.统筹支付 & ",0,0," & cur个人帐户 & ",'" & gstr结算号 & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    门诊结算_自贡 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_自贡(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Dim lng冲销ID As Long
    Dim strReturn As String, strBalanceNO As String
    Dim arrBalance
    Dim rsTemp As New ADODB.Recordset
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    
    Const int总金额 As Integer = 0
    Const int帐户支付 As Integer = 1
    Const int现金支付 As Integer = 2
    Const int统筹支付 As Integer = 3
    On Error GoTo errHand
    
    '提取本次冲销ID
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "自贡医保", lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
    
    '取出上次结算的收据号
    gstrSQL = "Select 支付顺序号 From 保险结算记录 Where 险类=[1] And 性质=1 And 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取上次结算的收据号", TYPE_四川自贡, lng结帐ID)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "没有找到原始的收费记录，无法完成门诊冲销！"
        Exit Function
    End If
    strBalanceNO = Nvl(rsTemp!支付顺序号)
    If strBalanceNO = "" Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "无效的门诊收据号，无法完成门诊冲销！"
        Exit Function
    End If
    
    '仅插入门诊主表就可以了，红票为-1
    On Error Resume Next
    gstrSQL = " Insert Into ClinicBill " & _
                " (ClinicBillNO,ClinicBalanceNO,InvoiceNO,DepartmentName,DoctorName,SickSerialNO,SickName,SickKindCode,RedBillFlag,OccurDate,StrikedBillNO) " & _
                " Select 'HCMZ" & strBalanceNO & "','HCMZ" & strBalanceNO & "',InvoiceNO,DepartmentName,DoctorName,SickSerialNO,SickName,SickKindCode," & _
                " -1,to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'),'" & strBalanceNO & "'" & _
                " From ClinicBill Where ClinicBillNO='" & strBalanceNO & "'"
    gcn自贡.Execute gstrSQL
    On Error GoTo errHand
    
    '调用门诊结算交易完成门诊结算冲销
    If Not 调用接口_自贡(业务类型_自贡.门诊结算, GetPass(lng病人ID) & "|" & "HCMZ" & strBalanceNO & "|" & 0, strReturn) Then Exit Function
    
    arrBalance = Split(strReturn, "|")
    cur_结算信息.总金额 = -1 * Val(arrBalance(int总金额))
    cur_结算信息.个人帐户 = -1 * Val(arrBalance(int帐户支付))
    cur_结算信息.现金支付 = -1 * Val(arrBalance(int现金支付))
    cur_结算信息.统筹支付 = -1 * Val(arrBalance(int统筹支付))
   
    Call Get帐户信息(TYPE_四川自贡, cur_结算信息.病人ID, Year(zlDatabase.Currentdate()), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & cur_结算信息.病人ID & "," & TYPE_四川自贡 & "," & Year(zlDatabase.Currentdate()) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur_结算信息.个人帐户 & "," & _
        cur进入统筹累计 + cur_结算信息.进入统筹 & "," & _
        cur统筹报销累计 + cur_结算信息.统筹支付 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)'超限自付金额用于暂时保存，门诊类别
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_四川自贡 & "," & cur_结算信息.病人ID & "," & _
        Year(zlDatabase.Currentdate()) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + cur_结算信息.个人帐户 & "," & cur进入统筹累计 + cur_结算信息.进入统筹 & "," & _
        cur统筹报销累计 + cur_结算信息.统筹支付 & "," & int住院次数累计 & ",0,0,0," & cur_结算信息.总金额 & ",0,0," & _
        cur_结算信息.进入统筹 & "," & cur_结算信息.统筹支付 & ",0,0," & cur_结算信息.个人帐户 & ",'" & strBalanceNO & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    门诊结算冲销_自贡 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 个人余额_自贡(ByVal lng病人ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Nvl(帐户余额,0) AS 余额 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取帐户余额", TYPE_四川自贡, lng病人ID)
    个人余额_自贡 = rsTemp!余额
End Function

Public Function 入院登记_自贡(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim str住院号 As String, strReturn As String
    Dim str科室 As String, str医生 As String, str入院日期 As String
    Dim rsTemp As New ADODB.Recordset
    
    '以下变量和病种有关
    Dim lng病种 As Long, str病种 As String
    Dim rsSelected As New ADODB.Recordset
    Dim rs病种 As New ADODB.Recordset
    On Error GoTo errHand
    
    '住院要选择病种，以确认一些特殊收费项目
    gstrSQL = " Select A.SickSerialNo AS ID,A.SickNum AS 编码,A.SickName AS 名称,A.SickSpell AS 简码 " & _
            " From SickDefine A Where 1=2"
    Call OpenRecordset_OtherBase(rsSelected, "获取已选择的病种", gstrSQL, gcn自贡)
    gstrSQL = " Select A.SickSerialNo AS ID,A.SickNum AS 编码,A.SickName AS 名称,A.SickSpell AS 简码 " & _
            " From SickDefine A Where 1=1"
    Set rs病种 = New ADODB.Recordset
    Call OpenRecordset_OtherBase(rs病种, "身份验证", gstrSQL, gcn自贡)
    
    If rs病种.RecordCount > 0 Then
VirusSelect:
        If frm多病种选择_自贡.ShowSelect(rs病种, "ID", "医保病种选择", "请选择医保病种：", rsSelected, False, gcn自贡) = True Then
            lng病种 = 0
            str病种 = ""
            With rs病种
                If .RecordCount <> 0 Then .MoveFirst
                lng病种 = rs病种("ID")
                Do While Not .EOF
                    str病种 = str病种 & "|" & rs病种!ID
                    .MoveNext
                Loop
                If str病种 <> "" Then str病种 = Mid(str病种, 2)
            End With
        Else
            MsgBox "必须要选择病种！", vbInformation, gstrSysName
            GoTo VirusSelect
        End If
    End If
    
    str住院号 = Left(lng病人ID & "_" & lng主页ID, 16) & "_" & Mid(CStr(Get序列(lng病人ID)), 1, 3)
    '提取病人的入院科室,医生,入院日期
    gstrSQL = "Select B.名称 As 科室,A.门诊医师 As 医生,A.入院日期 " & _
             " From 病案主页 A,部门表 B " & _
             " Where A.入院科室ID=B.Id And A.病人ID=[1] And A.主页ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人的入院科室、医生与入院日期", lng病人ID, lng主页ID)
    str科室 = Nvl(rsTemp!科室)
    str医生 = Nvl(rsTemp!医生)
    str入院日期 = Format(rsTemp!入院日期, "yyyy-MM-dd HH:mm:ss")
    
    '插入入院记录(入院登记号(系统),入院登记号(外部),住院科室,医生,入院日期)
    On Error Resume Next
    gstrSQL = " Insert Into InHosRegister(INHOSREGISTERNO,INHOSNO,DEPARTMENTNAME,DOCTORNAME,INHOSDATE) " & _
              " Values ('" & str住院号 & "','" & str住院号 & "','" & str科室 & "','" & str医生 & "'," & _
              " to_Date('" & str入院日期 & "','yyyy-MM-dd hh24:mi:ss'))"
    gcn自贡.Execute gstrSQL
    On Error GoTo errHand
    
    '插入病种数据
    Call InsertDisease("RegHosSick", str住院号, str病种)
    
    If Not 调用接口_自贡(业务类型_自贡.入院登记, GetPass(lng病人ID) & "|" & str住院号, strReturn) Then Exit Function
    
    '改变病人当前状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_四川自贡 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    '记录病人的主页ID，也就是顺序号
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_四川自贡 & ",'顺序号','''" & str住院号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新序列")
    
    入院登记_自贡 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_自贡(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim str住院号 As String
    Dim strReturn As String
    On Error GoTo errHand
    str住院号 = Get住院号(lng病人ID)
    
    On Error Resume Next
    '填写出院登记记录
    gstrSQL = " Insert Into UnInHosRegister(InHosRegisterNo,UnRegisterReason) " & _
              " Values ('" & str住院号 & "','撤销入院')"
    gcn自贡.Execute gstrSQL
    On Error GoTo errHand
    
    If Not 调用接口_自贡(业务类型_自贡.撤销入院登记, GetPass(lng病人ID) & "|" & str住院号, strReturn) Then Exit Function
    
    '改变病人当前状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_四川自贡 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    入院登记撤销_自贡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_自贡(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim bln结帐 As Boolean
    Dim str住院号 As String, strReturn As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '判断该病人是否结算过，没有结算过的病人费用为零，说明需要调用就诊登记撤销
        bln结帐 = False
        gstrSQL = "Select 1 From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] And Nvl(结帐ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否该调用就诊登记撤销", lng病人ID, lng主页ID)
        If Not rsTemp.EOF Then
            bln结帐 = True
        End If
        
        If Not bln结帐 Then
            '无费出院以撤销入院方式办理
            str住院号 = Get住院号(lng病人ID)
            
            On Error Resume Next
            '填写出院登记记录
            gstrSQL = " Insert Into UnInHosRegister(InHosRegisterNo,UnRegisterReason) " & _
                      " Values ('" & str住院号 & "','撤销入院')"
            gcn自贡.Execute gstrSQL
            On Error GoTo errHand
            
            If Not 调用接口_自贡(业务类型_自贡.撤销入院登记, GetPass(lng病人ID) & "|" & str住院号, strReturn) Then Exit Function
        End If
    End If
    
    '改变病人当前状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_四川自贡 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    出院登记_自贡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_自贡(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '改变病人当前状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_四川自贡 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    出院登记撤销_自贡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 处方上传_自贡(ByVal str单据号 As String, ByVal int性质 As Integer, ByVal int状态 As Integer, Optional ByVal lng病人ID As Long = 0, Optional ByVal bln结算 As Boolean = False) As Boolean
    '如果lng病人ID不为零，则仅仅上传该病人的处方明细
    '自贡医保允许直接录入负数
    'todo:中软前置机表结构未考虑到多病人单的情况
    Dim strNO As String
    Dim str住院号 As String
    Dim rsTmp   As ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim rsHead As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not bln结算 Then
        '检查是否存在未对码的项目开处方
        gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
        If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then
            gstrSQL = " Select 1 " & _
                  " From 住院费用记录 A,(Select * From 保险支付项目 Where 险类=" & TYPE_四川自贡 & ") B,保险帐户 C,病案主页 D,病人信息 E,收费细目 F,部门表 G" & _
                  " Where A.NO='" & str单据号 & "' And A.记录性质=" & int性质 & " And A.记录状态=" & int状态 & _
                  IIf(lng病人ID = 0, "", " And A.病人ID=[2]") & _
                  " And E.病人ID=D.病人ID And E.主页ID=D.主页ID And A.病人ID=E.病人ID And A.开单部门ID=G.ID(+)" & _
                  " And C.病人ID=A.病人ID And A.收费细目ID=B.收费细目ID(+) And A.收费细目ID=F.ID" & _
                  " And C.险类=[1] And Nvl(A.是否上传,0)=0" & _
                  " And B.项目编码 Is NULL And Rownum<2"
        Else
            gstrSQL = " Select 1 " & _
                  " From 住院费用记录 A,(Select * From 保险支付项目 Where 险类=" & TYPE_四川自贡 & ") B,保险帐户 C,病案主页 D,病人信息 E,收费细目 F,部门表 G" & _
                  " Where A.NO='" & str单据号 & "' And A.记录性质=" & int性质 & " And A.记录状态=" & int状态 & _
                  IIf(lng病人ID = 0, "", " And A.病人ID=[2]") & _
                  " And E.病人ID=D.病人ID And E.住院次数=D.主页ID And A.病人ID=E.病人ID And A.开单部门ID=G.ID(+)" & _
                  " And C.病人ID=A.病人ID And A.收费细目ID=B.收费细目ID(+) And A.收费细目ID=F.ID" & _
                  " And C.险类=[1] And Nvl(A.是否上传,0)=0" & _
                  " And B.项目编码 Is NULL And Rownum<2"
        End If
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未对码的项目开处方", TYPE_四川自贡, lng病人ID)
        If rsCheck.RecordCount <> 0 Then
            MsgBox "该处方中存在未对码的项目，请检查！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '打开处方主表
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then

        gstrSQL = " Select A.NO,A.记录性质,A.记录状态,A.病人ID,A.主页ID,SUM(A.实收金额) AS 金额,G.名称 AS 开单科室,A.开单人" & _
              " From 住院费用记录 A,保险帐户 C,病案主页 D,病人信息 E,部门表 G" & _
              " Where A.NO='" & str单据号 & "' And A.记录性质=" & int性质 & " And A.记录状态=" & int状态 & _
              IIf(lng病人ID = 0, "", " And A.病人ID=[2]") & _
              " And E.病人ID=D.病人ID And E.主页ID=D.主页ID And A.病人ID=E.病人ID And A.开单部门ID=G.ID(+)" & _
              " And C.病人ID=A.病人ID And C.险类=[1] And Nvl(A.是否上传,0)=0" & _
              " Group by A.NO,A.记录性质,A.记录状态,A.病人ID,A.主页ID,G.名称,A.开单人"
    Else
        gstrSQL = " Select A.NO,A.记录性质,A.记录状态,A.病人ID,A.主页ID,SUM(A.实收金额) AS 金额,G.名称 AS 开单科室,A.开单人" & _
              " From 住院费用记录 A,保险帐户 C,病案主页 D,病人信息 E,部门表 G" & _
              " Where A.NO='" & str单据号 & "' And A.记录性质=" & int性质 & " And A.记录状态=" & int状态 & _
              IIf(lng病人ID = 0, "", " And A.病人ID=[2]") & _
              " And E.病人ID=D.病人ID And E.住院次数=D.主页ID And A.病人ID=E.病人ID And A.开单部门ID=G.ID(+)" & _
              " And C.病人ID=A.病人ID And C.险类=[1] And Nvl(A.是否上传,0)=0" & _
              " Group by A.NO,A.记录性质,A.记录状态,A.病人ID,A.主页ID,G.名称,A.开单人"
    End If
    Set rsHead = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次需要上传的处方主表", TYPE_四川自贡, lng病人ID)
    
    With rsHead
        Do While Not .EOF
            '处方主表(记帐单号,入院登记号,科室名称,医生,病种名称,金额,冲票标志)
            'Insert Into InHosBill
            '(InHosBillNo,InHosRegisterNO,DepartmentName,DoctorName,SickName,Amount,RedBillFlag)
            'Values
            '()
            strNO = !NO & !记录性质 & !记录状态
            str住院号 = Get住院号(!病人ID)
            On Error Resume Next
            gstrSQL = " Insert Into InHosBill" & _
                    " (InHosBillNo,InHosRegisterNO,DepartmentName,DoctorName,SickName,Amount,RedBillFlag)" & _
                    " Values" & _
                    "('" & strNO & "','" & str住院号 & "','" & Nvl(!开单科室) & "','" & Nvl(!开单人) & "',''," & Format(!金额, "#####0.00;-#####0.00;0.00") & ",1)"
            gcn自贡.Execute gstrSQL
            
            On Error GoTo errHand
            '打开处方明细表
            gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
            If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then
                gstrSQL = " Select A.NO,A.记录性质,A.记录状态,A.序号,A.收费细目ID,B.项目编码,A.付数*A.数次 AS 数量,A.实收金额,F.名称 AS 项目名称,F.计算单位," & _
                      " G.名称 AS 开单科室,A.开单人" & _
                      " From 住院费用记录 A,保险支付项目 B,保险帐户 C,病案主页 D,病人信息 E,收费细目 F,部门表 G" & _
                      " Where A.NO=[1] And A.记录性质=[2] And A.记录状态=[3] And A.病人ID=[4]" & _
                      " And E.病人ID=D.病人ID And E.主页ID=D.主页ID And A.病人ID=E.病人ID And A.开单部门ID=G.ID(+)" & _
                      " And C.病人ID=A.病人ID And A.收费细目ID=B.收费细目ID And A.收费细目ID=F.ID" & _
                      " And C.险类=B.险类 And C.险类=[5] And Nvl(A.是否上传,0)=0 And Nvl(A.实收金额,0)<>0" & _
                      " Order by A.NO,A.记录性质,A.记录状态,A.序号"
                      
            Else
                gstrSQL = " Select A.NO,A.记录性质,A.记录状态,A.序号,A.收费细目ID,B.项目编码,A.付数*A.数次 AS 数量,A.实收金额,F.名称 AS 项目名称,F.计算单位," & _
                      " G.名称 AS 开单科室,A.开单人" & _
                      " From 住院费用记录 A,保险支付项目 B,保险帐户 C,病案主页 D,病人信息 E,收费细目 F,部门表 G" & _
                      " Where A.NO=[1] And A.记录性质=[2] And A.记录状态=[3] And A.病人ID=[4]" & _
                      " And E.病人ID=D.病人ID And E.住院次数=D.主页ID And A.病人ID=E.病人ID And A.开单部门ID=G.ID(+)" & _
                      " And C.病人ID=A.病人ID And A.收费细目ID=B.收费细目ID And A.收费细目ID=F.ID" & _
                      " And C.险类=B.险类 And C.险类=[5] And Nvl(A.是否上传,0)=0 And Nvl(A.实收金额,0)<>0" & _
                      " Order by A.NO,A.记录性质,A.记录状态,A.序号"
            End If
            Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次需要上传的处方明细", CLng(!NO), CLng(!记录性质), CLng(!记录状态), CLng(!病人ID), TYPE_四川自贡)
            
            With rsDetail
                Do While Not .EOF
                    '处方明细(记帐单号,明细序号,医保编码,医院项目名称,单价,数量,金额,单位)
                    'Insert Into InHosBillDetail
                    '(InHosBillNO,InHosBillDetailNO,ItemNO,HosItemName,Price,Quantity,Amount,Spec)
                    'Values
                    '()
                    
                    On Error Resume Next
                    gstrSQL = " Insert Into InHosBillDetail" & _
                              " (InHosRegisterNO,InHosBillNO,InHosBillDetailNO,ItemNO,HosItemName,Price,Quantity,Amount,Spec)" & _
                              " Values" & _
                              "('" & str住院号 & "','" & strNO & "','" & !序号 & "','" & !项目编码 & "','" & ToVarchar(!项目名称, 100) & "'," & Format(!实收金额 / !数量, "#####0.00000;-#####0.00000;0.00") & "," & _
                              Format(!数量, "#####0.00000;-#####0.00000;0.00") & "," & Format(!实收金额, "#####0.00000;-#####0.00000;0.00") & ",'" & Nvl(!计算单位) & "')"
                    gcn自贡.Execute gstrSQL
                    On Error GoTo errHand
                    
                    gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    .MoveNext
                Loop
            End With
            
            .MoveNext
        Loop
    End With
    
    处方上传_自贡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    处方上传_自贡 = True
End Function

Public Function 住院虚拟结算_自贡(ByVal rsExse As ADODB.Recordset, ByVal lng病人ID As Long) As String
    Dim int结算类型 As Integer
    Dim str住院号 As String
    Dim strReturn As String
    Dim intRecur As Integer         '是否原伤复发
    Dim intCureKindCode As Integer  '诊断情况或出院方式
    Dim bln个人帐户 As Boolean
    Dim arrBalance
    
    Dim cur总费用 As Currency
    Dim rsUpload As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Const int总金额 As Integer = 0
    Const int全自费 As Integer = 1
    Const int首先自负 As Integer = 2
    Const int进入统筹 As Integer = 3
    Const int起付线 As Integer = 4
    Const int本次起付线 As Integer = 5
    Const int统筹支付 As Integer = 6
    Const int统筹自付 As Integer = 7
    Const int超封顶 As Integer = 8
    Const int帐户支付 As Integer = 9
    Const int现金支付 As Integer = 10
    On Error GoTo errHand
    
    '处方明细上传(仅提取出还未上传的处方)
    gstrSQL = " Select NO,记录性质,记录状态,count(*) Records From 住院费用记录 A,病人信息 B " & _
              " Where A.病人ID=[1] And A.病人ID=B.病人ID And A.主页ID=B.住院次数 " & _
              " And Nvl(记录状态,0)<>0 And Nvl(实收金额,0)<>0 And Nvl(是否上传,0)=0" & _
              " Having Count(*)>0" & _
              " Group by NO,记录性质,记录状态"
    Set rsUpload = zlDatabase.OpenSQLRecord(gstrSQL, "处方明细上传", lng病人ID)
    With rsUpload
        Do While Not .EOF
            If Not 处方上传_自贡(!NO, !记录性质, !记录状态, 0, True) Then Exit Function
            .MoveNext
        Loop
    End With
    
    '取本次费用总额
    gstrSQL = " Select Nvl(A.金额,0) AS 费用总额 From 病人未结费用 A,病人信息 B" & _
              " Where A.病人ID = [1] And A.病人ID=B.病人ID And A.主页ID=B.住院次数"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取本次费用总额", lng病人ID)
    If rsTemp.RecordCount <> 0 Then
        cur总费用 = rsTemp!费用总额
    Else
        cur总费用 = 0
    End If
    
    bln个人帐户 = (MsgBox("是否使用个人帐户支付？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    
    '如果病人已出院，则是出院结算；否则为中途结算(结算类型。0结算，1中结)
    gstrSQL = "Select A.病人ID,A.主页ID,出院日期,出院方式 From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 And A.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出院日期", lng病人ID)
    str住院号 = Get住院号(rsTemp!病人ID)
    If IsNull(rsTemp!出院日期) Then
        int结算类型 = 1
        intCureKindCode = 2
    Else
        int结算类型 = 0
        Select Case rsTemp!出院方式
        Case "死亡"
            intCureKindCode = 3
        Case "好转"
            intCureKindCode = 1
        Case Else
            intCureKindCode = 0
        End Select
    End If
    
    '以病人ID做为结算单号、出院收据号进行住院预结算
    gstr结算号 = Left(rsTemp!病人ID & "_" & rsTemp!主页ID, 16) & "_" & Mid(CStr(Get序列(rsTemp!病人ID)), 1, 3)
    '先删除以前未结算的结算记录
'    gstrSQL = "Delete InHosBalance Where InHosBalanceNO='" & gstr结算号 & "' And InHosRegisterNO='" & str住院号 & "'"
'    gcn自贡.Execute gstrSQL
    
    'IsRecru:针对人员身份为二等乙级伤残军人（4），1表示原伤复发
    'CureKindCode:治愈情况0-治愈;1-好转;2-未愈;3-死亡（中结一律为2，需初始化HIS数据为前置机中CureKind）
    intRecur = 0
    gstrSQL = "Select 人员身份 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否伤残军人", lng病人ID, TYPE_四川自贡)
    If Nvl(rsTemp!人员身份, 1) = 4 Then
        '伤残军人
        If MsgBox("该病人是原伤复发住院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            intRecur = 1
        End If
    End If
    
    gstrSQL = " Insert Into InHosBalance(InHosBalanceNO,InHosRegisterNO,InvoiceNO,PayType,RedBillFlag,occurdate,IsRecur,CureKindCode)" & _
              " Values ('" & gstr结算号 & "','" & str住院号 & "','" & gstr结算号 & "'," & int结算类型 & ",1," & _
              " to_date('" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & intRecur & "," & intCureKindCode & ")"
    gcn自贡.Execute gstrSQL
    
    '调用住院预结算接口
    '总金额|全自费|个人首先自负|进入统筹费用(包括起付线) |执行起付线|实际起付线金额|统筹支付|统筹自付|超封顶线费用|个人帐户支付|现金支付
    If Not 调用接口_自贡(业务类型_自贡.住院预结算, GetPass(lng病人ID) & "|" & str住院号 & "|" & int结算类型 & "|3|0|" & IIf(bln个人帐户, 1, 0), strReturn) Then Exit Function
    arrBalance = Split(strReturn, "|")
    cur_结算信息.总金额 = Val(arrBalance(int总金额))
    cur_结算信息.统筹支付 = Val(arrBalance(int统筹支付))
    cur_结算信息.个人帐户 = Val(arrBalance(int帐户支付))
    If Format(cur_结算信息.总金额, "#####0.00") <> Format(cur总费用, "#####0.00") Then
        If MsgBox("HIS费用总额与医保费用总额不等，是否继续结算？" & vbCrLf & _
        "HIS：" & Format(cur总费用, "#####0.00") & Space(10) & "医保：" & Format(cur_结算信息.总金额, "#####0.00"), vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    If cur_结算信息.统筹支付 <> 0 Then 住院虚拟结算_自贡 = 住院虚拟结算_自贡 & "|医保基金;" & cur_结算信息.统筹支付 & ";0"
    If cur_结算信息.个人帐户 <> 0 Then 住院虚拟结算_自贡 = 住院虚拟结算_自贡 & "|个人帐户;" & cur_结算信息.个人帐户 & ";0"
    If 住院虚拟结算_自贡 <> "" Then 住院虚拟结算_自贡 = Mid(住院虚拟结算_自贡, 2)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_自贡(ByVal lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim int结算类型 As Integer
'    Dim str临时结算单号 As String
'    Dim str结算单号 As String
    Dim str住院号 As String
    Dim str发票号 As String
    Dim lng主页ID As Long
    
    Dim strReturn As String
    Dim arrBalance
    
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur总费用_HIS As Currency, cur总费用 As Currency, cur医保基金 As Currency, cur个人帐户 As Currency
    Dim rsTemp As New ADODB.Recordset
    
    '以下变量和病种有关
    Dim lng病种 As Long, str病种 As String
    Dim rsSelected As New ADODB.Recordset
    Dim rs病种 As New ADODB.Recordset
    
    Const int总金额 As Integer = 0
    Const int全自费 As Integer = 1
    Const int首先自负 As Integer = 2
    Const int进入统筹 As Integer = 3
    Const int起付线 As Integer = 4
    Const int本次起付线 As Integer = 5
    Const int统筹支付 As Integer = 6
    Const int统筹自付 As Integer = 7
    Const int超封顶 As Integer = 8
    Const int帐户支付 As Integer = 9
    Const int现金支付 As Integer = 10
    On Error GoTo errHand
    '如果病人已出院，则是出院结算；否则为中途结算(结算类型。0结算，1中结)
    gstrSQL = "Select A.病人ID,A.主页ID,出院日期 From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 And A.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出院日期", lng病人ID)
    lng主页ID = rsTemp!主页ID
    str住院号 = Get住院号(rsTemp!病人ID)
'    str临时结算单号 = "NO" & lng病人ID
    If IsNull(rsTemp!出院日期) Then
        int结算类型 = 1
    Else
        int结算类型 = 0
    End If
    
    '--医保中心不关心发票号--
'    '提取本次结算发票号与结算单号
'    gstrSQL = "Select NO,实际票号 From 病人结帐记录 Where ID=" & lng结帐ID
'    Call OpenRecordset(rsTemp, "提取结算单号")
'    str结算单号 = rsTemp!NO
'    str发票号 = Nvl(rsTemp!实际票号)
'    If str发票号 = "" Then
'        MsgBox "发票号不能为空！", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '将结算单号改为结帐单号，出院收据号改为发票号
'    gstrSQL = "Update InHosBalance Set InHosBalanceNO='" & str结算单号 & "',InvoiceNO='" & str发票号 & "' Where InHosBalanceNO='" & str临时结算单号 & "'"
'    gcn自贡.Execute gstrSQL
    '-------------------------
    
    '每次结算都要选择病种，以确认一些特殊收费项目
    gstrSQL = " Select A.SickSerialNo AS ID,A.SickNum AS 编码,A.SickName AS 名称,A.SickSpell AS 简码 " & _
            " From SickDefine A Where 1=2"
    Call OpenRecordset_OtherBase(rsSelected, "获取已选择的病种", gstrSQL, gcn自贡)
    gstrSQL = " Select A.SickSerialNo AS ID,A.SickNum AS 编码,A.SickName AS 名称,A.SickSpell AS 简码 " & _
            " From SickDefine A Where 1=1"
    Set rs病种 = New ADODB.Recordset
    Call OpenRecordset_OtherBase(rs病种, "身份验证", gstrSQL, gcn自贡)
    
    If rs病种.RecordCount > 0 Then
VirusSelect:
        If frm多病种选择_自贡.ShowSelect(rs病种, "ID", "医保病种选择", "请选择医保病种：", rsSelected, False, gcn自贡) = True Then
            lng病种 = 0
            str病种 = ""
            With rs病种
                If .RecordCount <> 0 Then .MoveFirst
                lng病种 = rs病种("ID")
                Do While Not .EOF
                    str病种 = str病种 & "|" & rs病种!ID
                    .MoveNext
                Loop
                If str病种 <> "" Then str病种 = Mid(str病种, 2)
            End With
        Else
            Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "必须要选择病种！", vbInformation, gstrSysName
            GoTo VirusSelect
        End If
    End If
    
    Call InsertDisease("InHosSick", gstr结算号, str病种)
    
    '调用住院结算接口
    '调用住院预结算接口
    '总金额|全自费|个人首先自负|进入统筹费用(包括起付线) |执行起付线|实际起付线金额|统筹支付|统筹自付|超封顶线费用|个人帐户支付|现金支付
    If Not 调用接口_自贡(业务类型_自贡.住院结算, GetPass(lng病人ID) & "|" & gstr结算号 & "|" & int结算类型 & "|3|0|" & IIf(Val(cur_结算信息.个人帐户) = 0, 0, 1), strReturn) Then Exit Function
    arrBalance = Split(strReturn, "|")
    cur_结算信息.总金额 = Val(arrBalance(int总金额))
    cur_结算信息.全自费 = Val(arrBalance(int全自费))
    cur_结算信息.首先自付 = Val(arrBalance(int首先自负))
    cur_结算信息.进入统筹 = Val(arrBalance(int进入统筹))
    cur_结算信息.本次起付线 = Val(arrBalance(int起付线))
    cur_结算信息.实际起付线 = Val(arrBalance(int本次起付线))
    cur_结算信息.统筹支付 = Val(arrBalance(int统筹支付))
    cur_结算信息.超封顶费用 = Val(arrBalance(int超封顶))
    cur_结算信息.个人帐户 = Val(arrBalance(int帐户支付))
    
    '保存保险结算记录
    Call Get帐户信息(TYPE_四川自贡, cur_结算信息.病人ID, Year(zlDatabase.Currentdate()), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & cur_结算信息.病人ID & "," & TYPE_四川自贡 & "," & Year(zlDatabase.Currentdate()) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur_结算信息.个人帐户 & "," & _
        cur进入统筹累计 + cur_结算信息.进入统筹 & "," & _
        cur统筹报销累计 + cur_结算信息.统筹支付 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    'g结算数据.超限自付金额中保存的是门诊病人就诊类型（急诊、特殊病门诊或普通门诊），结算记录的备注保存的是病种的名称
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)'超限自付金额用于暂时保存，门诊类别
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_四川自贡 & "," & lng病人ID & "," & _
        Year(zlDatabase.Currentdate()) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + cur_结算信息.个人帐户 & "," & cur进入统筹累计 + cur_结算信息.进入统筹 & "," & _
        cur统筹报销累计 + cur_结算信息.统筹支付 & "," & int住院次数累计 & "," & cur_结算信息.本次起付线 & ",0," & cur_结算信息.实际起付线 & "," & cur_结算信息.总金额 & "," & cur_结算信息.全自费 & "," & cur_结算信息.首先自付 & "," & _
        cur_结算信息.进入统筹 & "," & cur_结算信息.统筹支付 & ",0," & cur_结算信息.超封顶费用 & "," & cur_结算信息.个人帐户 & ",'" & gstr结算号 & "'," & lng主页ID & "," & int结算类型 & ",'" & str住院号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    住院结算_自贡 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_自贡(ByVal lng结帐ID As Long) As Boolean
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim lng冲销ID As Long
    Dim int结算类型 As Integer
    Dim lng主页ID As Long, lng病人ID As Long
    Dim str结算单号 As String, str被冲结算单号 As String
    Dim str住院号 As String
    
    Dim strReturn As String
    Dim arrBalance
    Dim rsTemp As New ADODB.Recordset
    
    Const int总金额 As Integer = 0
    Const int全自费 As Integer = 1
    Const int起付线 As Integer = 2
    Const int本次起付线 As Integer = 3
    Const int统筹支付 As Integer = 4
    Const int统筹自付 As Integer = 5
    Const int超封顶 As Integer = 6
    Const int帐户支付 As Integer = 7
    Const int现金支付 As Integer = 8
    On Error GoTo errHand
    
    '获取冲销ID
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '获取结算单号，结算类型
    gstrSQL = "Select 病人ID,主页ID,支付顺序号,中途结帐 From 保险结算记录 Where 性质=2 And 险类=[1] And 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取结算单号", TYPE_四川自贡, lng结帐ID)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "没有找到原始的收费记录，无法完成住院结算冲销！", vbInformation, gstrSysName
        Exit Function
    End If
    lng主页ID = rsTemp!主页ID
    cur_结算信息.病人ID = rsTemp!病人ID
    str住院号 = Get住院号(rsTemp!病人ID)
    int结算类型 = Nvl(rsTemp!中途结帐, 0)
    str被冲结算单号 = Nvl(rsTemp!支付顺序号)
    gstrSQL = "Select NO from 病人结帐记录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取结算单号", lng冲销ID)
    str结算单号 = "HCZY" & rsTemp!NO
    
    On Error Resume Next
    '插入结算记录，但红票标志为-1
    gstrSQL = " Insert Into InHosBalance(InHosBalanceNO,InHosRegisterNO,InvoiceNO,PayType,RedBillFlag,StrikedBillNO,occurdate)" & _
              " Values ('" & str结算单号 & "','" & str住院号 & "','" & str结算单号 & "'," & int结算类型 & ",-1,'" & str被冲结算单号 & "',to_date('" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    gcn自贡.Execute gstrSQL
    On Error GoTo errHand
    
    '调用住院结算冲销接口(总金额|全自费|被冲执行起付线|被冲实际起付线金额|被冲统筹支付|被冲统筹自付|被冲超封顶线费用|被冲个人账户支付|冲正后个人账户金额|被冲现金支付)
    If Not 调用接口_自贡(业务类型_自贡.住院结算冲销, GetPass(cur_结算信息.病人ID) & "|" & str结算单号 & "|" & str被冲结算单号, strReturn) Then Exit Function
    
    arrBalance = Split(strReturn, "|")
    cur_结算信息.总金额 = -1 * Val(arrBalance(int总金额))
    cur_结算信息.全自费 = 0
    cur_结算信息.首先自付 = 0
    cur_结算信息.进入统筹 = 0
    cur_结算信息.本次起付线 = -1 * Val(arrBalance(int起付线))
    cur_结算信息.实际起付线 = -1 * Val(arrBalance(int本次起付线))
    cur_结算信息.统筹支付 = -1 * Val(arrBalance(int统筹支付))
    cur_结算信息.超封顶费用 = -1 * Val(arrBalance(int超封顶))
    cur_结算信息.个人帐户 = -1 * Val(arrBalance(int帐户支付))
    
    '保存保险结算记录
    Call Get帐户信息(TYPE_四川自贡, cur_结算信息.病人ID, Year(zlDatabase.Currentdate()), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & cur_结算信息.病人ID & "," & TYPE_四川自贡 & "," & Year(zlDatabase.Currentdate()) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur_结算信息.个人帐户 & "," & _
        cur进入统筹累计 + cur_结算信息.进入统筹 & "," & _
        cur统筹报销累计 + cur_结算信息.统筹支付 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    
    'g结算数据.超限自付金额中保存的是门诊病人就诊类型（急诊、特殊病门诊或普通门诊），结算记录的备注保存的是病种的名称
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)'超限自付金额用于暂时保存，门诊类别
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_四川自贡 & "," & cur_结算信息.病人ID & "," & _
        Year(zlDatabase.Currentdate()) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + cur_结算信息.个人帐户 & "," & cur进入统筹累计 + cur_结算信息.进入统筹 & "," & _
        cur统筹报销累计 + cur_结算信息.统筹支付 & "," & int住院次数累计 & "," & cur_结算信息.本次起付线 & ",0," & cur_结算信息.实际起付线 & "," & cur_结算信息.总金额 & "," & cur_结算信息.全自费 & "," & cur_结算信息.首先自付 & "," & _
        cur_结算信息.进入统筹 & "," & cur_结算信息.统筹支付 & ",0," & cur_结算信息.超封顶费用 & "," & cur_结算信息.个人帐户 & ",'" & str结算单号 & "'," & lng主页ID & "," & int结算类型 & ",'" & str住院号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "自贡医保")
    住院结算冲销_自贡 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Sub InsertDisease(ByVal strTable As String, ByVal strKey As String, ByVal strDisease As String)
    Dim arrDisease
    Dim intDO As Integer, intCOUNT As Integer
    Dim strInsert As String
    Dim rsTemp As New ADODB.Recordset
    '向指定表中插入病种数据（入院是RegHosSick，结算病种是InHosSick）
    On Error Resume Next
    
    arrDisease = Split(strDisease, "|")
    intCOUNT = UBound(arrDisease)
    
    '打开病种记录集
    gstrSQL = "Select SickSerialNO,SickName,SickKindCode From SickDefine Where SickSerialNO in (" & Replace(strDisease, "|", ",") & ")"
    Call OpenRecordset_OtherBase(rsTemp, "打开病种记录集", gstrSQL, gcn自贡)
    
    '准备插入数据
    gstrSQL = "Insert Into " & strTable & "(" & IIf(UCase(strTable) = "REGHOSSICK", "InHosRegisterNO", "InHosBalanceNO") & _
              ",SickSerialNO,HosSickName,RowNO) Values ('" & strKey & "',"
    For intDO = 0 To intCOUNT
        rsTemp.Filter = "SickSerialNO=" & Val(arrDisease(intDO))
        strInsert = arrDisease(intDO) & ",'" & rsTemp!SickName & "'," & intDO + 1 & ")"
        gcn自贡.Execute gstrSQL & strInsert
    Next
    rsTemp.Filter = 0
End Sub

Private Function Get序列(ByVal lng病人ID As Long) As Integer
    Dim rsTemp As New ADODB.Recordset
    '仅供住院使用，用于产生序列
    gstrSQL = "Select Nvl(退休证号,0) AS 序列 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取序列", TYPE_四川自贡, lng病人ID)
    Get序列 = rsTemp!序列 + 1
    
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_四川自贡 & ",'退休证号','" & Get序列 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新序列")
End Function

Private Function Get住院号(ByVal lng病人ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '提取病人的住院登记流水号
    gstrSQL = "Select 顺序号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人的住院流水号", lng病人ID, TYPE_四川自贡)
    Get住院号 = Nvl(rsTemp!顺序号)
End Function

Private Function GetPass(ByVal lng病人ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '提取病人的密码（不允许空密码）
    gstrSQL = "Select 密码 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取密码", lng病人ID, TYPE_四川自贡)
    GetPass = Nvl(rsTemp!密码)
End Function
