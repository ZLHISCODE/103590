Attribute VB_Name = "mdlExpense"
Option Explicit
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
    g本机公共模块 = 5
    g本机私有模块 = 6
End Enum
Public Const SPI_GETWORKAREA = 48
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Public Function GetFeeKind() As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 编码, 名称, 简码 From 收费项目类别"
    Set GetFeeKind = zlDatabase.OpenSQLRecord(strSql, "获取收费类别")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MoneyOverFlow(objBill As ExpenseBill) As Boolean
'功能：检查单据合计金额是否溢出
'说明：以Currency上限922337203685477为准
    Dim dbl应收 As Double, dbl实收 As Double
    Dim i As Integer, j As Integer
    
    '要用VAL转为Double进行运算
    For i = 1 To objBill.Details.Count
        For j = 1 To objBill.Details(i).InComes.Count
            If Abs(dbl应收 + Val(objBill.Details(i).InComes(j).应收金额)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            If Abs(dbl实收 + Val(objBill.Details(i).InComes(j).实收金额)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            dbl应收 = dbl应收 + Val(objBill.Details(i).InComes(j).应收金额)
            dbl实收 = dbl实收 + Val(objBill.Details(i).InComes(j).实收金额)
        Next
    Next
End Function

Public Function GetBillTotal(objBill As ExpenseBill) As Currency
'功能：获取单据费目合计金额
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    
    For Each objBillDetail In objBill.Details
        For Each objBillIncome In objBillDetail.InComes
            GetBillTotal = GetBillTotal + objBillIncome.实收金额
        Next
    Next
End Function

Public Function GetServiceDept(str收费细目IDs As String) As ADODB.Recordset
'功能:获取多个药房的存储库房与服务科室
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    If InStr(1, str收费细目IDs, ",") = 0 Then
        strSql = "" & _
        "   Select Distinct /*+ Rule*/ 收费细目ID,Nvl(开单科室ID,0) as 开单科室ID,执行科室id " & _
        "   From 收费执行科室 A " & _
        "   Where   A.收费细目ID  =[2] "
    Else
        strSql = "" & _
        "   Select Distinct /*+ Rule*/ 收费细目ID,Nvl(开单科室ID,0) as 开单科室ID,执行科室id " & _
        "   From 收费执行科室 A," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        "   Where   A.收费细目ID  = j.Column_Value"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取执行科室信息", Replace(str收费细目IDs, "'", ""), Val(str收费细目IDs))
    If Not rsTmp.EOF Then Set GetServiceDept = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub LoadPatientBaby(ByRef cboBaby As ComboBox, ByVal lngPatient As Long, lngPatientPage As Long)
    Dim rsTmp As ADODB.Recordset, i As Long
    
    cboBaby.Clear
    cboBaby.AddItem "0-病人本人"
    cboBaby.ItemData(cboBaby.NewIndex) = 0
    Call zlControl.CboSetIndex(cboBaby.hWnd, 0)
    
    If lngPatient <> 0 Then
        Set rsTmp = GetPatientBaby(lngPatient, lngPatientPage)
        With rsTmp
            For i = 1 To .RecordCount
                If Not IsNull(!婴儿姓名) Then
                    cboBaby.AddItem !序号 & "-" & !婴儿姓名
                Else
                    cboBaby.AddItem !序号 & "-第" & !序号 & "个婴儿"
                End If
                cboBaby.ItemData(cboBaby.NewIndex) = !序号
                .MoveNext
            Next
        End With
    End If
End Sub

Public Function GetPatientBaby(ByVal lngPatient As Long, lngPatientPage As Long) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 序号, 婴儿姓名 From 病人新生儿记录 Where 病人id = [1] And 主页ID = [2]"
    On Error GoTo errH
    Set GetPatientBaby = zlDatabase.OpenSQLRecord(strSql, "读取新生儿记录", lngPatient, lngPatientPage)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDrugTotal(ByVal objBill As ExpenseBill, ByVal lng药品ID As Long, ByVal lng药房ID As Long, _
    Optional lng批次 As Long = 0) As Double
'功能：获取单据中指定药品在同一药房多行的数量合
    Dim i As Integer, dblCount As Double
    
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).收费细目ID = lng药品ID _
            And objBill.Details(i).执行部门ID = lng药房ID And objBill.Details(i).Detail.批次 = lng批次 Then
            dblCount = dblCount + objBill.Details(i).付数 * objBill.Details(i).数次
        End If
    Next
    GetDrugTotal = dblCount
End Function

Public Function GetFirstRow(curBill As ExpenseBill, Optional strClass As String) As Integer
'功能：获取当前单据中第一个为药品的收费行号
'参数：strClass=取第一中药或西药行,空为药品
'返回：0=没有药品收费行
    Dim i As Long
    If curBill.Details.Count = 0 Then GetFirstRow = 0
    For i = 1 To curBill.Details.Count
        If strClass = "" Then
            If InStr(",5,6,7,", curBill.Details(i).收费类别) > 0 Then
                GetFirstRow = i: Exit Function
            End If
        Else
            If curBill.Details(i).收费类别 = strClass Then
                GetFirstRow = i: Exit Function
            End If
        End If
    Next
End Function

Public Function Get医保大类(ByVal lng收费细目ID As Long, ByVal int险类 As Integer) As String
'功能：获取指定收费项目的保险大类名称
'参数：
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select N.名称" & _
        " From 保险支付项目 M,保险支付大类 N " & _
        " Where M.收费细目ID=[1] And M.险类=[2] And M.大类ID=N.ID"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng收费细目ID, int险类)
    If rsTmp.RecordCount > 0 Then Get医保大类 = rsTmp!名称
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_Check特准项目(ByVal objclsInsure As Object, ByVal intInsure As Integer, ByVal lng病人ID As Long, Optional ByVal bln门诊 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否医保病人是否需要检查特准项目
    '入参:objInsure-创建的医保部件
    '     intInsure-险类
    '     lng病人ID-病人ID
    '     bln门诊-是否门诊
    '出参:
    '返回:如果需要检查特准项目,则返回true,否则返回False
    '编制:刘兴洪
    '问题:24862
    '日期:2009-08-12 10:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zl_Check特准项目 = False
     If bln门诊 Then
        If objclsInsure.GetCapability(support门诊病人不受特准项目限制, lng病人ID, intInsure) = False Then zl_Check特准项目 = True
        Exit Function
     End If
    If objclsInsure.GetCapability(support住院病人不受特准项目限制, lng病人ID, intInsure) = False Then zl_Check特准项目 = True
End Function
Public Function Get保险特准项目(lng病人ID As Long, strField As String) As String
    Dim rsTmp As New ADODB.Recordset
    Dim lng病种ID As Long, int险类 As Integer, strSql As String
    Dim strA1 As String, strA2 As String, strB1 As String, strB2 As String
    
    On Error GoTo errH
            
    '先取病人病种,是该病种是否有各类特准项目设置
    strSql = _
        " Select A.险类,A.病种ID,Nvl(B.大类,0) as 大类,B.性质,Count(*)" & _
        " From 保险帐户 A,保险特准项目 B" & _
        " Where Nvl(A.病种ID,0)=B.病种ID And Nvl(A.病种ID,0)<>0" & _
        " And B.性质 IN(1,2) And A.病人ID=[1]" & _
        " Group by A.险类,A.病种ID,Nvl(B.大类,0),B.性质"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lng病人ID)
    If rsTmp.EOF Then Exit Function
    
    lng病种ID = rsTmp!病种ID
    int险类 = rsTmp!险类
    
    '允许的收费细目
    rsTmp.Filter = "大类=0 And 性质=1"
    If Not rsTmp.EOF Then
        strA1 = strField & _
            " IN (" & _
            "   Select 收费细目ID From 保险支付项目" & _
            "   Where 险类 = " & int险类 & _
            "   And 收费细目ID IN (" & _
            "       Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=0 And 性质=1 And 病种ID=" & lng病种ID & ")" & _
            ")"
    End If
    
    '允许的保险大类
    rsTmp.Filter = "大类=1 And 性质=1"
    If Not rsTmp.EOF Then
        strA2 = strField & _
            " IN (" & _
            "   Select 收费细目ID From 保险支付项目" & _
            "   Where 险类 = " & int险类 & _
            "   And 大类ID IN (" & _
            "       Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=1 And 性质=1 And 病种ID=" & lng病种ID & ")" & _
            ")"
    End If
    
    '禁止的收费细目
    rsTmp.Filter = "大类=0 And 性质=2"
    If Not rsTmp.EOF Then
        strB1 = strField & _
            " Not IN (" & _
            "   Select 收费细目ID From 保险支付项目" & _
            "   Where 险类 = " & int险类 & _
            "   And 收费细目ID IN (" & _
            "       Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=0 And 性质=2 And 病种ID=" & lng病种ID & ")" & _
            ")"
    End If
    
    '禁止的保险大类
    rsTmp.Filter = "大类=1 And 性质=2"
    If Not rsTmp.EOF Then
        strB2 = strField & _
            " Not IN (" & _
            "   Select 收费细目ID From 保险支付项目" & _
            "   Where 险类 = " & int险类 & _
            "   And 大类ID IN (" & _
            "       Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=1 And 性质=2 And 病种ID=" & lng病种ID & ")" & _
            ")"
    End If
    
    '组合SQL(允许部分要用Or)
    strSql = ""
    If strA1 <> "" And strA2 <> "" Then
        strSql = " And (" & strA1 & " Or " & strA2 & ")"
    Else
        If strA1 <> "" Then strSql = " And " & strA1
        If strA2 <> "" Then strSql = " And " & strA2
    End If
    If strB1 <> "" Then strSql = strSql & " And " & strB1
    If strB2 <> "" Then strSql = strSql & " And " & strB2
        
    Get保险特准项目 = strSql
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get处方职务(lng药品ID As Long) As String
'功能：根据药品ID获取其处方职务
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    Get处方职务 = "00"
    strSql = "Select Nvl(B.处方职务,'00') as 处方职务 From 药品规格 A,药品特性 B Where A.药名ID=B.药名ID And A.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lng药品ID)
    If Not rsTmp.EOF Then Get处方职务 = rsTmp!处方职务
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get处方限量(lngID As Long) As Double
'功能：获取指定药品的处方限量,以零售单位返回。
'参数：lngID=药品ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select Nvl(A.处方限量,0) as 处方限量" & _
        " From 药品特性 A,药品规格 B Where A.药名ID=B.药名ID And B.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lngID)
    If Not rsTmp.EOF Then Get处方限量 = rsTmp!处方限量
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemExistInsure(ByVal lng病人ID As Long, ByVal lng收费细目ID As Long, ByVal int险类 As Integer) As Boolean
'功能：判断收费项目是否设置了保险支付项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    
    If gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, int险类) Then
        ItemExistInsure = True: Exit Function
    End If
    
    strSql = "Select 1 From 保险支付项目 Where 收费细目ID=[1] And 险类=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lng收费细目ID, int险类)
    ItemExistInsure = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckLimit(ByVal objBill As ExpenseBill, Optional ByVal intRow As Integer, Optional ByVal bln药房单位 As Boolean) As Boolean
'功能：费用单据药品处方限量检查
'说明：
'   1.全部没超过限量，返回真；如有超过药品，则在函数内提示，并返回假。
'   2.记帐表是为每个病人单独检查
    Dim rsTmp As New ADODB.Recordset, strSql As String
    Dim tmpDetail As BillDetail, curDetail As BillDetail
    Dim strItemIDs As String, i As Integer, j As Integer
    Dim dblTime As Double, dbl剂量 As Double
    
    CheckLimit = True
    If objBill.Details.Count = 0 Then Exit Function
    
    On Error GoTo errH
    
    '收集病人
    For i = 1 To objBill.Details.Count
        If intRow = 0 Or (intRow > 0 And i = intRow) Then
            With objBill.Details(i)
                '收集药品ID
                If InStr(strItemIDs & ",", "," & .收费细目ID & ",") = 0 And InStr(",5,6,7,", .收费类别) > 0 Then
                    strItemIDs = strItemIDs & "," & .收费细目ID
                End If
            End With
        End If
    Next
    If strItemIDs = "" Then Exit Function
    strItemIDs = Mid(strItemIDs, 2)
    
    strSql = "Select  /*+ RULE */  A.药品ID,A.剂量系数,B.计算单位 as 剂量单位" & _
        " From 药品规格 A,诊疗项目目录 B," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        " Where A.药名ID=B.ID And A.药品ID  = j.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", strItemIDs)
    
    strItemIDs = ""
    For j = 1 To objBill.Details.Count
        If intRow = 0 Or (intRow > 0 And j = intRow) Then
            Set tmpDetail = objBill.Details(j)
            If InStr(",5,6,7,", tmpDetail.收费类别) > 0 And tmpDetail.Detail.处方限量 > 0 Then
                If InStr(strItemIDs, "," & tmpDetail.收费细目ID) = 0 Then
                    dblTime = 0
                    For Each curDetail In objBill.Details
                        If InStr(",5,6,7,", curDetail.收费类别) > 0 And tmpDetail.收费细目ID = curDetail.收费细目ID Then
                            dblTime = dblTime + curDetail.付数 * curDetail.数次
                        End If
                    Next
                    rsTmp.Filter = "药品ID=" & tmpDetail.收费细目ID
                    If Not rsTmp.EOF Then
                        If bln药房单位 Then
                            dbl剂量 = dblTime * tmpDetail.Detail.药房包装 * rsTmp!剂量系数
                        Else
                            dbl剂量 = dblTime * rsTmp!剂量系数
                        End If
                        If dbl剂量 > tmpDetail.Detail.处方限量 Then
                            MsgBox "药品 """ & tmpDetail.Detail.名称 & """ 的总剂量 " & _
                                FormatEx(dbl剂量, 5) & rsTmp!剂量单位 & "(" & FormatEx(dblTime, 5) & IIF(bln药房单位, tmpDetail.Detail.药房单位, tmpDetail.Detail.计算单位) & ") 超过处方限量 " & _
                                FormatEx(tmpDetail.Detail.处方限量, 5) & rsTmp!剂量单位 & " ！", vbInformation, gstrSysName
                            CheckLimit = False: Exit Function
                        End If
                    End If
                    strItemIDs = strItemIDs & "," & tmpDetail.收费细目ID
                End If
            End If
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStockInfo(lng药品ID As Long, bln药房 As Boolean, bln药库 As Boolean, _
    Optional ByVal bln药房单位 As Boolean, Optional str药房包装 As String) As String
'功能：获取药品在各个药房，药库的库存信息
'参数："bln药房/bln药库"至少要有一个设置为真
'返回：描述信息
    Dim strSql As String, strSQL2 As String
    Dim str性质 As String, i As Long
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    If bln药房 And bln药库 Then
        str性质 = "'中药房','西药房','成药房','中药库','西药库','成药库'"
    ElseIf bln药房 Then
        str性质 = "'中药房','西药房','成药房'"
    ElseIf bln药库 Then
        str性质 = "'中药库','西药库','成药库'"
    End If
    
    '排除多个性质的情况,不区分门诊、住院
    strSql = _
        " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And Instr([1],B.工作性质)>0"
    '药房不分批药品不管效期
    strSQL2 = "Select 部门ID From 部门性质说明 Where 工作性质 IN('西药房','成药房','中药房')"
    '不分批或分批药品
    strSql = _
        " Select B.编码,B.名称,A.库房ID," & _
        " Nvl(Sum(A.可用数量),0)" & IIF(bln药房单位, "/Nvl(C." & str药房包装 & ",1)", "") & " as 库存" & _
        " From 药品库存 A,(" & strSql & ") B,药品规格 C" & _
        " Where A.库房ID=B.ID And A.药品ID=C.药品ID" & _
        " And ((A.效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Or (Nvl(C.药房分批,0)=0 And A.库房ID IN(" & strSQL2 & ")))" & _
        " And A.性质=1 And A.药品ID=[2]" & _
        " Group by B.编码,B.名称,A.库房ID,Nvl(C." & str药房包装 & ",1)" & _
        " Having Sum(Nvl(A.可用数量,0))<>0" & _
        " Order By B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", str性质, lng药品ID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = strSql & "," & rsTmp!名称 & ":" & rsTmp!库存
        rsTmp.MoveNext
    Loop
    strSql = Mid(strSql, 2)
    GetStockInfo = strSql
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function OverTime() As Boolean
'功能：判断当前是否处于加班时间范围内
'返回：真-当前处于加班时间内,假-不处于
    Dim str上午 As String, str下午 As String
    Dim DateBegin As Date, DateEnd As Date
    Dim curTime As Date
    
    str上午 = zlDatabase.GetPara(1, glngSys): str下午 = zlDatabase.GetPara(2, glngSys)
    curTime = CDate(Format(zlDatabase.Currentdate, "HH:MM:SS"))
    
    If str上午 <> "" Then
        DateBegin = CDate(Trim(Split(UCase(str上午), "AND")(0)))
        DateEnd = CDate(Trim(Split(UCase(str上午), "AND")(1)))
    End If
    
    If Not (curTime >= DateBegin And curTime <= DateEnd) Then
        If str下午 <> "" Then
            DateBegin = CDate(Trim(Split(UCase(str下午), "AND")(0)))
            DateEnd = CDate(Trim(Split(UCase(str下午), "AND")(1)))
        End If
        If Not (curTime >= DateBegin And curTime <= DateEnd) Then OverTime = True
    End If
End Function

Public Function GetBillRows(str单据号 As String, int记录性质 As Integer, int病人来源 As Integer) As Integer
'功能：获取一张费用单据中未作废的费用行数
'参数：int记录性质=1-收费(划价),2-记帐
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strTab As String
    
    strTab = IIF(int记录性质 = 1 Or (int记录性质 = 2 And int病人来源 = 1), "门诊费用记录", "住院费用记录")

    On Error GoTo errH
    
    
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    strSql = _
        " Select 序号,Sum(数量) as 剩余数量" & _
        " From (" & _
        " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号," & _
        " Avg(Nvl(付数, 1) * 数次) As 数量" & _
        " From " & strTab & _
        " Where NO=[1] And 记录性质=[2]" & _
        " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", str单据号, int记录性质)
    If Not rsTmp.EOF Then GetBillRows = rsTmp.RecordCount
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistInsure(strNO As String) As Integer
'功能：判断指定的住院记帐单据是否对医保病人记的帐
'参数：strNO=记帐单据号
'返回：如果是则返回病人险类
'说明：1.只管住院医保病人,不管门诊病人的医技记帐
'      2.记帐表只返回第一个病人的险类,单据中也应该只有一种险类
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select B.险类 From 住院费用记录 A,病案主页 B" & _
        " Where A.记录性质=2 And A.记录状态 IN(0,1,3) And B.险类 is Not NULL" & _
        " And A.NO=[1] And A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", strNO)
    If Not rsTmp.EOF Then BillExistInsure = rsTmp!险类
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistDelete(strNO As String, int记录性质 As Integer, int病人来源 As Integer) As Boolean
'功能：判断指定单据是否包含(部分)退费或销帐的内容
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strTab As String
    
    strTab = IIF(int记录性质 = 1 Or (int记录性质 = 2 And int病人来源 = 1), "门诊费用记录", "住院费用记录")
    
    On Error GoTo errH
    
    strSql = "Select NO From " & strTab & " Where NO=[1] And 记录性质=[2] And 记录状态=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "BillExistDelete", strNO, int记录性质)
    BillExistDelete = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsureName(intInsure As Integer) As String
'功能：根据保险类别序号获取保险类别名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 名称 From 保险类别 Where 序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", intInsure)
    If Not rsTmp.EOF Then GetInsureName = Nvl(rsTmp!名称)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'功能：获取药品或卫材出库检查的集合
'参数：bytType:0-药品，1-卫材
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '避免出错
    
    strSql = _
        " Select Distinct A.ID,C.检查方式" & _
        " From 部门表 A,部门性质说明 B," & IIF(bytType = 0, "药品出库检查", "材料出库检查") & " C" & _
        " Where B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 " & IIF(bytType = 0, "IN('中药房','西药房','成药房')", "='发料部门'") & _
        " And C.库房ID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetStockCheck")
    For i = 1 To rsTmp.RecordCount
        colStock.Add Nvl(rsTmp!检查方式, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function CheckDisable(objBill As ExpenseBill) As String
'功能：检查单据中的药品的禁忌情况
'返回：药品互相禁忌提示信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strInfo As String
    Dim i As Long, j As Long, k As Long
    Dim strGroup As String, strIDs As String
    Dim blnStop As Boolean
    
    For i = 1 To objBill.Details.Count
        If InStr(",5,6,7,", objBill.Details(i).收费类别) > 0 Then
            strIDs = strIDs & "," & objBill.Details(i).收费细目ID
        End If
    Next
    strIDs = Mid(strIDs, 2)
    If strIDs = "" Or UBound(Split(strIDs, ",")) < 1 Then Exit Function
    
    strSql = _
        " Select /*+ RULE */  A.组编号,Count(Distinct A.项目ID) as 禁忌数" & _
        " From 诊疗互斥项目 A,药品规格 B," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        " Where A.项目ID=B.药名ID And B.药品ID  = j.Column_Value" & _
        " Having Count(Distinct A.项目ID)>1  " & _
        "  Group by A.组编号"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", strIDs)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strGroup = strGroup & "," & rsTmp!组编号
            rsTmp.MoveNext
        Next
        strGroup = Mid(strGroup, 2)
        
        For i = 0 To UBound(Split(strGroup, ","))
            strSql = _
            "Select /*+ RULE */   Distinct C.类型,C.组编号,D.编码,D.名称,D.规格" & _
            " From 药品规格 A,诊疗项目目录 B,诊疗互斥项目 C,收费项目目录 D," & _
            "          (Select Column_Value From Table(Cast(f_num2list([2]) As Zltools.t_Numlist ))) J " & _
            " Where A.药名ID=B.ID And B.ID=C.项目ID And A.药品ID=D.ID" & _
            "           And C.组编号=[1]" & _
            "           And A.药品ID  = j.Column_Value" & _
            " Order by C.类型,C.组编号,D.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", Val(Split(strGroup, ",")(i)), strIDs)
            
            If Not rsTmp.EOF Then
                rsTmp.Filter = "类型=1"
                If rsTmp.RecordCount > 1 Then
                    k = k + 1
                    strInfo = strInfo & vbCrLf & "第 " & k & " 组(互相慎用)：" & vbCrLf
                    For j = 1 To rsTmp.RecordCount
                        strInfo = strInfo & "[" & rsTmp!编码 & "]" & rsTmp!名称 & IIF(IsNull(rsTmp!规格), "", "(" & rsTmp!规格 & ")") & "                 " & vbCrLf
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Filter = "类型=2"
                If rsTmp.RecordCount > 1 Then
                    blnStop = True
                    k = k + 1
                    strInfo = strInfo & vbCrLf & "第 " & k & " 组(互相禁用)：" & vbCrLf
                    For j = 1 To rsTmp.RecordCount
                        strInfo = strInfo & "[" & rsTmp!编码 & "]" & rsTmp!名称 & IIF(IsNull(rsTmp!规格), "", "(" & rsTmp!规格 & ")") & "                 " & vbCrLf
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Filter = 0
            End If
        Next
        If strInfo <> "" Then
            If blnStop Then
                CheckDisable = "发现单据中下列药品互相禁用或慎用：" & vbCrLf & strInfo & vbCrLf & "请修改禁用药品后再继续！"
            Else
                CheckDisable = "发现单据中下列药品互相禁用或慎用：" & vbCrLf & strInfo & vbCrLf & "要继续吗？"
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ImportBill(ByVal int来源 As Integer, ByVal str单据号 As String, _
    ByVal int记录性质 As Integer, Optional ByVal bln零耗 As Boolean) As ExpenseBill
    '功能：读取费用单据到单据对象中(目前忽略从属项目,当作独立项目),用于修改或导入时用
    '参数：int记录性质=1-收费(划价),2-记帐
    '      bln零耗=是否零耗费用登记,实收金额为0
    '返回：存放单据信息的单据对象
    '说明：因为可能现时项目价格信息已作调整,所以费用相关内容重新计算
    '      不管是导入还是修改单据,都不应包含已停用收费细目

    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    Dim intCurNo As Integer, strInfo As String
    Dim int序号 As Integer, blnDo As Boolean, i As Integer
    
    Dim dblAllTime As Double, dblCurTime As Double, dblPriceSingle As Double
    Dim dblPrice As Double, str药房 As String
    
    Dim colSerial As New Collection '用于处理从属父号
    Dim strSql As String
    Dim strTab As String
    
    Dim lng西药房 As Long, lng成药房 As Long, lng中药房 As Long
    Dim bln药房单位 As Boolean, str药房单位 As String, str药房包装 As String
        
    strTab = IIF(int记录性质 = 1 Or (int记录性质 = 2 And int来源 = 1), "门诊费用记录", "住院费用记录")
    
    '缺省药房
    lng西药房 = Val(zlDatabase.GetPara(IIF(int来源 = 2, "住院", "门诊") & "缺省西药房", glngSys, p医嘱附费管理))
    lng成药房 = Val(zlDatabase.GetPara(IIF(int来源 = 2, "住院", "门诊") & "缺省成药房", glngSys, p医嘱附费管理))
    lng中药房 = Val(zlDatabase.GetPara(IIF(int来源 = 2, "住院", "门诊") & "缺省中药房", glngSys, p医嘱附费管理))
    
    '药品单位
    bln药房单位 = Val(zlDatabase.GetPara("药品单位", glngSys, p医嘱附费管理)) <> 0
    If int来源 = 1 Then
        str药房单位 = "门诊单位": str药房包装 = "门诊包装"
    Else
        str药房单位 = "住院单位": str药房包装 = "住院包装"
    End If
        
    '------------------------------------------------------------------------------------------
    '收费价目关联:新计算价格,如果有多个价格,则一个收费细目ID行就会有多条序号相同的记录
    '价格父号 is NULL:只取每个收费细目ID的第一行(药品只有一行),因为要重算价格
        
    '使用指定的药房读取正确的库存
    str药房 = "Decode(A.收费类别,'5'," & IIF(lng西药房 <> 0, lng西药房, "A.执行部门ID") & "," & _
        "'6'," & IIF(lng成药房 <> 0, lng成药房, "A.执行部门ID") & "," & _
        "'7'," & IIF(lng中药房 <> 0, lng中药房, "A.执行部门ID") & ",A.执行部门ID)"
    
    '药房不分批核算的药品不管效期
    strSql = _
    " Select X.药品ID,W.材料ID,W.跟踪在用,A.序号 As 序号,A.从属父号,A.NO,A.记录性质,A.记录状态," & IIF(strTab = "住院费用记录", "A.多病人单", " 0 as 多病人单") & ",A.婴儿费,A.费别,A.姓名,A.性别,A.年龄," & _
            IIF(strTab = "住院费用记录", "A.床号,A.病人病区ID,A.主页ID", "A.付款方式 as 床号,0 as 病人病区ID,0 as 主页ID") & _
    "       ,A.标识号,A.病人ID,A.病人科室ID,A.开单部门ID,A.门诊标志,A.加班标志," & _
    "       A.附加标志,A.收费类别,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
    "       A.标准单价 As 标准单价," & str药房 & " as 执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
    "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(F.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
    "       B.屏蔽费别,B.说明,B.执行科室,Nvl(A.费用类型,B.费用类型) 费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
    "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
    "       Decode(A.收费类别,'4',1,X." & str药房包装 & ") as 药房包装," & _
    "       Decode(A.收费类别,'4',B.计算单位,X." & str药房单位 & ") as 药房单位," & _
    "       Decode(A.收费类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,Nvl(Y.库存,0) As 库存,B.录入限量" & _
    " From " & strTab & " A,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,收费项目别名 F,收费项目别名 E1,材料特性 W,药品规格 X," & _
    "       (Select A.药品ID,A.库房ID,Sum(Nvl(A.可用数量,0)) as 库存 From 药品库存 A" & _
    "        Where A.性质=1 And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
    "               And A.药品ID IN(Select 收费细目ID From " & strTab & " Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1])" & _
    "        Group by A.药品ID,A.库房ID) Y" & _
    " Where A.记录性质=[2] And A.记录状态 IN(0,1,3) And A.NO=[1]" & _
    "       And A.价格父号 Is Null And A.收费细目ID=B.ID And A.收费细目ID=D.收费细目ID" & _
    "       And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is NULL)" & _
    "       And A.收费类别=C.编码 And A.收费细目ID=X.药品ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
    "       And A.收费细目ID=Y.药品ID(+) And " & str药房 & "=Y.库房ID(+)" & _
    "       And A.收费细目ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[3]" & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
    "       And ((Sysdate Between D.执行日期 And D.终止日期) Or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"

    strSql = "Select * From (" & strSql & ") Order by 序号"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", str单据号, int记录性质, IIF(gbyt药品名称显示 = 1, 3, 1))
    
    '没有记录就是空单子
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    If rsTmp.RecordCount <> 0 Then
        With rsTmp
            i = 1
            Do While Not .EOF
                '处理单据主体=====================================================
                If i = 1 Then
                    objBill.NO = !NO
                    objBill.病人ID = Nvl(!病人ID, 0)
                    objBill.主页ID = Nvl(!主页ID, 0)
                    objBill.病区ID = Nvl(!病人病区ID, 0)
                    objBill.科室ID = Nvl(!病人科室id, 0)
                    objBill.姓名 = Nvl(!姓名)
                    objBill.性别 = Nvl(!性别)
                    objBill.年龄 = Nvl(!年龄)
                    objBill.费别 = Nvl(!费别)
                    objBill.标识号 = Nvl(!标识号)
                    objBill.床号 = Nvl(!床号)
                    objBill.费别 = Nvl(!费别)
                    objBill.门诊标志 = Nvl(!门诊标志, 0)
                    objBill.加班标志 = Nvl(!加班标志, 0)
                    objBill.婴儿费 = Nvl(!婴儿费, 0)
                    objBill.开单部门ID = Nvl(!开单部门ID, 0)
                    objBill.划价人 = Nvl(!划价人)
                    objBill.开单人 = Nvl(!开单人)
                    objBill.操作员编号 = Nvl(!操作员编号)
                    objBill.操作员姓名 = Nvl(!操作员姓名)
                    objBill.发生时间 = !发生时间
                    objBill.登记时间 = !登记时间
                    objBill.多病人单 = Nvl(!多病人单, 0) <> 0
                End If
                
                '处理收费细目=====================================================
                Set objBillDetail = New BillDetail
                Set objBillDetail.Detail = New Detail
                            
                '处理序号,从属父号
                intCurNo = intCurNo + 1
                objBillDetail.序号 = intCurNo '实际是行号
                colSerial.Add intCurNo, "_" & !序号 '记录原序号现在的行号
                If Not IsNull(!从属父号) Then
                    objBillDetail.从属父号 = colSerial("_" & !从属父号)
                End If
                                                                    
                '使用原定的动态费别
                objBillDetail.收费类别 = !收费类别
                objBillDetail.收费细目ID = !收费细目ID
                objBillDetail.计算单位 = IIF(IsNull(!计算单位), "", !计算单位)
                
                objBillDetail.付数 = Nvl(!付数, 1)
                If InStr(",5,6,7,", !收费类别) > 0 And bln药房单位 Then
                    objBillDetail.数次 = Nvl(!数次, 0) / Nvl(!药房包装, 1)
                Else
                    objBillDetail.数次 = Nvl(!数次, 0)
                End If
                
                objBillDetail.附加标志 = Nvl(!附加标志, 0)
                objBillDetail.摘要 = Nvl(!摘要)
                objBillDetail.执行部门ID = Nvl(!执行部门ID, 0)
                objBillDetail.发药窗口 = Nvl(!发药窗口)
                objBillDetail.Detail.ID = !收费细目ID
                objBillDetail.Detail.编码 = !编码
                objBillDetail.Detail.变价 = Nvl(!是否变价, 0) = 1
                objBillDetail.Detail.从项数次 = 0 '!!!目前忽略从属项目,当作独立项目
                objBillDetail.Detail.固有从属 = 0 '!!!目前忽略从属项目,当作独立项目
                objBillDetail.Detail.规格 = Nvl(!规格)
                objBillDetail.Detail.计算单位 = Nvl(!计算单位)
                
                objBillDetail.Detail.药房单位 = Nvl(!药房单位)
                objBillDetail.Detail.药房包装 = Nvl(!药房包装, 1)
                If InStr(",5,6,7,", !收费类别) > 0 And bln药房单位 Then
                    objBillDetail.Detail.库存 = Nvl(!库存, 0) / Nvl(!药房包装, 1)
                Else
                    objBillDetail.Detail.库存 = Nvl(!库存, 0)
                End If
                objBillDetail.Detail.录入限量 = Val("" & !录入限量)
                
                objBillDetail.Detail.加班加价 = Nvl(!加班加价, 0) <> 0
                objBillDetail.Detail.类别 = Nvl(!类别)
                objBillDetail.Detail.类别名称 = Nvl(!类别名称)
                objBillDetail.Detail.名称 = Nvl(!名称)
                objBillDetail.Detail.商品名 = Nvl(!商品名)
                objBillDetail.Detail.屏蔽费别 = Nvl(!屏蔽费别, 0) <> 0
                objBillDetail.Detail.说明 = Nvl(!说明)
                objBillDetail.Detail.执行科室 = Nvl(!执行科室, 0)
                objBillDetail.Detail.类型 = Nvl(!费用类型)
                objBillDetail.Detail.处方职务 = Get处方职务(objBillDetail.Detail.ID)
                
                objBillDetail.Detail.药名ID = Nvl(!药名ID, 0)
                objBillDetail.Detail.变价 = Nvl(!是否变价, 0) <> 0
                objBillDetail.Detail.分批 = Nvl(!分批, 0) <> 0
                objBillDetail.Detail.跟踪在用 = Nvl(!跟踪在用, 0) = 1
                objBillDetail.Detail.要求审批 = 0
                
                '处理价格部份=====================================================
                Set objBillDetail.InComes = New BillInComes
                Do
                    '按照现有的价格设置重新计算
                    If !是否变价 = 1 Then
                        If InStr(",5,6,7,", !收费类别) > 0 Or (!收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                            '----------------------------------------------------------------------------------------------
                            '时价药品计算价格(分批可不分批)
                            dblAllTime = !付数 * !数次 '这里是售价数量
                            If dblAllTime <> 0 Then
                                dblPrice = Get时价药品应收金额(objBillDetail.执行部门ID, CLng(!收费细目ID), dblAllTime, gstrDec, dblPriceSingle)
                                If dblAllTime <> 0 Then
                                    If !收费类别 = "4" Then
                                        MsgBox "时价卫生材料""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "时价药品""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    End If
                                    '库存不够,只涉及一个批次时以首批时价为准，否则以第一批或者平均价都不合适
                                    objBillIncome.标准单价 = 0
                                Else
                                    objBillIncome.标准单价 = IIF(dblPriceSingle = 0, Format(dblPrice / (!付数 * !数次), gstrDecPrice), dblPriceSingle) '这里是售价价格
                                End If
                            Else
                                objBillIncome.标准单价 = 0
                            End If
                            '----------------------------------------------------------------------------------------------
                        Else
                            If Abs(!标准单价) > Abs(Nvl(!现价, 0)) Then
                                objBillIncome.标准单价 = Nvl(!缺省价格, 0)
                            Else
                                objBillIncome.标准单价 = !标准单价
                            End If
                        End If
                    Else
                        objBillIncome.标准单价 = !现价
                    End If
                                        
                    If InStr(",5,6,7,", !收费类别) > 0 And bln药房单位 Then
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(!药房包装, 1), gstrDecPrice)
                    Else
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价, gstrDecPrice)
                    End If
                    objBillIncome.现价 = Nvl(!现价, 0) '现价原价对药品变价无用
                    objBillIncome.原价 = Nvl(!原价, 0)
                    objBillIncome.收入项目ID = Nvl(!现收入ID, 0)
                    objBillIncome.收入项目 = Nvl(!收入项目)
                    objBillIncome.收据费目 = Nvl(!现费目)
                    
                    '应收金额=单价*付次*数次
                    If !是否变价 = 1 And (InStr(",5,6,7,", !收费类别) > 0 Or !收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                        objBillIncome.应收金额 = dblPrice '保证应收金额与零售金额没有误差
                    Else
                        objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                    End If
                    
                    '附加手术费率用计算(所有收入项目)
                    If Nvl(!附加标志, 0) = 1 And Nvl(!收费类别) = "F" Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * Nvl(!附术收费率, 100) / 100
                    End If
                    
                    '加班费用率计算
                    If Nvl(!加班标志, 0) = 1 And Nvl(!加班加价, 0) = 1 Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * (1 + Nvl(!加班加价率, 0) / 100)
                    End If
                    objBillIncome.应收金额 = Format(objBillIncome.应收金额, gstrDec)
                    
                    '计算实收金额
                    If bln零耗 Then
                        objBillIncome.实收金额 = 0
                    Else
                        If Nvl(!屏蔽费别, 0) = 1 Then
                            objBillIncome.实收金额 = objBillIncome.应收金额
                        Else
                            '使用原定的动态费别
                            objBillIncome.实收金额 = ActualMoney(objBill.费别, !现收入ID, objBillIncome.应收金额, objBillDetail.收费细目ID, _
                                objBillDetail.执行部门ID, !付数 * !数次, IIF(Nvl(!加班标志, 0) = 1 And Nvl(!加班加价, 0) = 1, Nvl(!加班加价率, 0) / 100, 0))
                        End If
                    End If
                    
                    With objBillIncome
                        objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
                    End With
                    
                    '判断下一条记录是否属于当前行
                    blnDo = False
                    int序号 = !序号
                    .MoveNext
                    If Not .EOF Then blnDo = (int序号 = !序号)
                    i = i + 1
                Loop While blnDo And Not .EOF
               
                With objBillDetail
                    objBill.Details.Add .InComes, .Detail, .收费细目ID, .序号, .从属父号, .收费类别, .计算单位, .付数, .数次, .附加标志, .执行部门ID, .发药窗口, , , , .摘要
                End With
            Loop
        End With
    End If
    
    Set ImportBill = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function ImportStuffBill(ByVal int来源 As Integer, ByVal str单据号 As String, _
    ByVal int记录性质 As Integer, ByVal lng虚拟库房ID As Long) As ExpenseBill
    '功能：读取费用单据到单据对象中(目前忽略从属项目,当作独立项目),用于修改或导入时用
    '参数：int记录性质=1-收费(划价),2-记帐
    '      bln零耗=是否零耗费用登记,实收金额为0
    '返回：存放单据信息的单据对象
    '说明：因为可能现时项目价格信息已作调整,所以费用相关内容重新计算
    '      不管是导入还是修改单据,都不应包含已停用收费细目

    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    Dim intCurNo As Integer, strInfo As String
    Dim int序号 As Integer, blnDo As Boolean, i As Integer
    
    Dim dblAllTime As Double, dblCurTime As Double, dblPriceSingle As Double
    Dim dblPrice As Double, str药房 As String
    
    Dim colSerial As New Collection '用于处理从属父号
    Dim strSql As String, strStock As String
    Dim strTab As String
    
    Dim lng西药房 As Long, lng成药房 As Long, lng中药房 As Long
    Dim bln药房单位 As Boolean, str药房单位 As String, str药房包装 As String
        
    strTab = IIF(int记录性质 = 1 Or (int记录性质 = 2 And int来源 = 1), "门诊费用记录", "住院费用记录")
    '------------------------------------------------------------------------------------------
    '收费价目关联:新计算价格,如果有多个价格,则一个收费细目ID行就会有多条序号相同的记录
    '价格父号 is NULL:只取每个收费细目ID的第一行(药品只有一行),因为要重算价格
        
    strStock = _
        " Select A.费用ID,Max( A.药品ID) as 药品ID,Max(A.批次) as 批次,Max(A.商品条码) as 商品条码 ,Max(A.内部条码) as 内部条码 " & _
        " From 药品收发记录 A" & _
        " Where A.NO=[1]  And  单据 =25 And MOD(A.记录状态,3) in (0,1)" & _
        " Group by A.费用ID "
        
    strStock = "" & _
    "   Select A.费用ID,A.批次,A.商品条码,A.内部条码,sum(b.可用数量) as 可用数量 " & _
    "   From (" & strStock & ") A,药品库存 B " & _
    "   Where A.药品id=b.药品ID(+) And B.库房ID(+)=[4] " & _
    "   Group by A.费用ID,A.批次,A.商品条码,A.内部条码"
    
 
    '药房不分批核算的药品不管效期
    strSql = _
    " Select X.药品ID,W.材料ID,W.跟踪在用,A.序号 As 序号,A.从属父号,A.NO,A.记录性质,A.记录状态," & IIF(strTab = "住院费用记录", "A.多病人单", " 0 as 多病人单") & ",A.婴儿费,A.费别,A.姓名,A.性别,A.年龄," & _
            IIF(strTab = "住院费用记录", "A.床号,A.病人病区ID,A.主页ID", "A.付款方式 as 床号,0 as 病人病区ID,0 as 主页ID") & _
    "       ,A.标识号,A.病人ID,A.病人科室ID,A.开单部门ID,A.门诊标志,A.加班标志," & _
    "       A.附加标志,A.收费类别,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
    "       A.标准单价 As 标准单价,A.执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
    "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(F.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
    "       B.屏蔽费别,B.说明,B.执行科室,Nvl(A.费用类型,B.费用类型) 费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
    "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
    "       1 as 药房包装, B.计算单位 as 药房单位, Nvl(W.在用分批,0) as 分批, nvl(y.批次,0) as 批次,y.商品条码,y.内部条码 ,Nvl(Y.可用数量,0) As 库存,B.录入限量" & _
    " From " & strTab & " A,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,收费项目别名 F,收费项目别名 E1,材料特性 W,药品规格 X," & _
    "       (" & strStock & ") Y" & _
    " Where A.记录性质=[2] And A.记录状态 IN(0,1,3) And A.NO=[1]" & _
    "       And A.价格父号 Is Null And A.收费细目ID=B.ID And A.收费细目ID=D.收费细目ID" & _
    "       And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is NULL)" & _
    "       And A.收费类别=C.编码 And A.收费细目ID=X.药品ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
    "       And A.ID=y.费用ID(+) " & _
    "       And A.收费细目ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[3]" & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
    "       And ((Sysdate Between D.执行日期 And D.终止日期) Or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"

    strSql = "Select * From (" & strSql & ") Order by 序号"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", str单据号, int记录性质, IIF(gbyt药品名称显示 = 1, 3, 1), lng虚拟库房ID)
    
    '没有记录就是空单子
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    If rsTmp.RecordCount <> 0 Then
        With rsTmp
            i = 1
            Do While Not .EOF
                '处理单据主体=====================================================
                If i = 1 Then
                    objBill.NO = !NO
                    objBill.病人ID = Nvl(!病人ID, 0)
                    objBill.主页ID = Nvl(!主页ID, 0)
                    objBill.病区ID = Nvl(!病人病区ID, 0)
                    objBill.科室ID = Nvl(!病人科室id, 0)
                    objBill.姓名 = Nvl(!姓名)
                    objBill.性别 = Nvl(!性别)
                    objBill.年龄 = Nvl(!年龄)
                    objBill.费别 = Nvl(!费别)
                    objBill.标识号 = Nvl(!标识号)
                    objBill.床号 = Nvl(!床号)
                    objBill.费别 = Nvl(!费别)
                    objBill.门诊标志 = Nvl(!门诊标志, 0)
                    objBill.加班标志 = Nvl(!加班标志, 0)
                    objBill.婴儿费 = Nvl(!婴儿费, 0)
                    objBill.开单部门ID = Nvl(!开单部门ID, 0)
                    objBill.划价人 = Nvl(!划价人)
                    objBill.开单人 = Nvl(!开单人)
                    objBill.操作员编号 = Nvl(!操作员编号)
                    objBill.操作员姓名 = Nvl(!操作员姓名)
                    objBill.发生时间 = !发生时间
                    objBill.登记时间 = !登记时间
                    objBill.多病人单 = Nvl(!多病人单, 0) <> 0
                End If
                
                '处理收费细目=====================================================
                Set objBillDetail = New BillDetail
                Set objBillDetail.Detail = New Detail
                            
                '处理序号,从属父号
                intCurNo = intCurNo + 1
                objBillDetail.序号 = intCurNo '实际是行号
                colSerial.Add intCurNo, "_" & !序号 '记录原序号现在的行号
                If Not IsNull(!从属父号) Then
                    objBillDetail.从属父号 = colSerial("_" & !从属父号)
                End If
                                                                    
                '使用原定的动态费别
                objBillDetail.收费类别 = !收费类别
                objBillDetail.收费细目ID = !收费细目ID
                objBillDetail.计算单位 = IIF(IsNull(!计算单位), "", !计算单位)
                
                objBillDetail.付数 = Nvl(!付数, 1)
                If InStr(",5,6,7,", !收费类别) > 0 And bln药房单位 Then
                    objBillDetail.数次 = Nvl(!数次, 0) / Nvl(!药房包装, 1)
                Else
                    objBillDetail.数次 = Nvl(!数次, 0)
                End If
                
                objBillDetail.附加标志 = Nvl(!附加标志, 0)
                objBillDetail.摘要 = Nvl(!摘要)
                objBillDetail.执行部门ID = Nvl(!执行部门ID, 0)
                objBillDetail.发药窗口 = Nvl(!发药窗口)
                objBillDetail.Detail.ID = !收费细目ID
                objBillDetail.Detail.编码 = !编码
                objBillDetail.Detail.变价 = Nvl(!是否变价, 0) = 1
                objBillDetail.Detail.从项数次 = 0 '!!!目前忽略从属项目,当作独立项目
                objBillDetail.Detail.固有从属 = 0 '!!!目前忽略从属项目,当作独立项目
                objBillDetail.Detail.规格 = Nvl(!规格)
                objBillDetail.Detail.计算单位 = Nvl(!计算单位)
                
                objBillDetail.Detail.药房单位 = Nvl(!药房单位)
                objBillDetail.Detail.药房包装 = Nvl(!药房包装, 1)
                objBillDetail.Detail.库存 = Nvl(!库存, 0)
                objBillDetail.Detail.录入限量 = Val("" & !录入限量)
                
                objBillDetail.Detail.加班加价 = Nvl(!加班加价, 0) <> 0
                objBillDetail.Detail.类别 = Nvl(!类别)
                objBillDetail.Detail.类别名称 = Nvl(!类别名称)
                objBillDetail.Detail.名称 = Nvl(!名称)
                objBillDetail.Detail.商品名 = Nvl(!商品名)
                objBillDetail.Detail.屏蔽费别 = Nvl(!屏蔽费别, 0) <> 0
                objBillDetail.Detail.说明 = Nvl(!说明)
                objBillDetail.Detail.执行科室 = Nvl(!执行科室, 0)
                objBillDetail.Detail.类型 = Nvl(!费用类型)
                objBillDetail.Detail.处方职务 = Get处方职务(objBillDetail.Detail.ID)
                
                objBillDetail.Detail.药名ID = Nvl(!药名ID, 0)
                objBillDetail.Detail.变价 = Nvl(!是否变价, 0) <> 0
                objBillDetail.Detail.分批 = Nvl(!分批, 0) <> 0
                objBillDetail.Detail.跟踪在用 = Nvl(!跟踪在用, 0) = 1
                objBillDetail.Detail.要求审批 = 0
                objBillDetail.Detail.批次 = Nvl(!批次, 0)
                objBillDetail.Detail.商品条码 = Nvl(!商品条码)
                objBillDetail.Detail.内部条码 = Nvl(!内部条码)
                '处理价格部份=====================================================
                Set objBillDetail.InComes = New BillInComes
                Do
                    '按照现有的价格设置重新计算
                    If !是否变价 = 1 Then
                        If InStr(",5,6,7,", !收费类别) > 0 Or (!收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                            '----------------------------------------------------------------------------------------------
                            '时价药品计算价格(分批可不分批)
                            dblAllTime = !付数 * !数次 '这里是售价数量
                            If dblAllTime <> 0 Then
                                dblPrice = Get时价材料应收金额(lng虚拟库房ID, CLng(!收费细目ID), Nvl(!批次, 0), dblAllTime, gstrDec, dblPriceSingle, True)
                                If dblAllTime <> 0 Then
                                    
                                    If !收费类别 = "4" Then
                                        MsgBox "时价卫生材料""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "时价药品""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    End If
                                    '库存不够,只涉及一个批次时以首批时价为准，否则以第一批或者平均价都不合适
                                    objBillIncome.标准单价 = 0
                                Else
                                    objBillIncome.标准单价 = IIF(dblPriceSingle = 0, Format(dblPrice / (!付数 * !数次), gstrDecPrice), dblPriceSingle) '这里是售价价格
                                End If
                            Else
                                objBillIncome.标准单价 = 0
                            End If
                            '----------------------------------------------------------------------------------------------
                        Else
                            If Abs(!标准单价) > Abs(Nvl(!现价, 0)) Then
                                objBillIncome.标准单价 = Nvl(!缺省价格, 0)
                            Else
                                objBillIncome.标准单价 = !标准单价
                            End If
                        End If
                    Else
                        objBillIncome.标准单价 = !现价
                    End If
                                        
                    If InStr(",5,6,7,", !收费类别) > 0 And bln药房单位 Then
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(!药房包装, 1), gstrDecPrice)
                    Else
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价, gstrDecPrice)
                    End If
                    objBillIncome.现价 = Nvl(!现价, 0) '现价原价对药品变价无用
                    objBillIncome.原价 = Nvl(!原价, 0)
                    objBillIncome.收入项目ID = Nvl(!现收入ID, 0)
                    objBillIncome.收入项目 = Nvl(!收入项目)
                    objBillIncome.收据费目 = Nvl(!现费目)
                    
                    '应收金额=单价*付次*数次
                    If !是否变价 = 1 And (InStr(",5,6,7,", !收费类别) > 0 Or !收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                        objBillIncome.应收金额 = dblPrice '保证应收金额与零售金额没有误差
                    Else
                        objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                    End If
                    
                    '附加手术费率用计算(所有收入项目)
                    If Nvl(!附加标志, 0) = 1 And Nvl(!收费类别) = "F" Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * Nvl(!附术收费率, 100) / 100
                    End If
                    
                    '加班费用率计算
                    If Nvl(!加班标志, 0) = 1 And Nvl(!加班加价, 0) = 1 Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * (1 + Nvl(!加班加价率, 0) / 100)
                    End If
                    objBillIncome.应收金额 = Format(objBillIncome.应收金额, gstrDec)
                    
                    '计算实收金额
                    If Nvl(!屏蔽费别, 0) = 1 Then
                        objBillIncome.实收金额 = objBillIncome.应收金额
                    Else
                        '使用原定的动态费别
                        objBillIncome.实收金额 = ActualMoney(objBill.费别, !现收入ID, objBillIncome.应收金额, objBillDetail.收费细目ID, _
                            objBillDetail.执行部门ID, !付数 * !数次, IIF(Nvl(!加班标志, 0) = 1 And Nvl(!加班加价, 0) = 1, Nvl(!加班加价率, 0) / 100, 0))
                    End If
                    
                    With objBillIncome
                        objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
                    End With
                    
                    '判断下一条记录是否属于当前行
                    blnDo = False
                    int序号 = !序号
                    .MoveNext
                    If Not .EOF Then blnDo = (int序号 = !序号)
                    i = i + 1
                Loop While blnDo And Not .EOF
               
                With objBillDetail
                    objBill.Details.Add .InComes, .Detail, .收费细目ID, .序号, .从属父号, .收费类别, .计算单位, .付数, .数次, .附加标志, .执行部门ID, .发药窗口, , , , .摘要
                End With
            Loop
        End With
    End If
    
    Set ImportStuffBill = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBillMoney(strNO As String) As Currency
'功能：获取一张正常记帐单的单据金额
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select Sum(实收金额) as 金额 From 住院费用记录 Where NO=[1] And 记录性质=2 And 记录状态 IN(0,1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", strNO)
    If Not rsTmp.EOF Then GetBillMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(lng病人ID As Long, ByVal byt来源 As Byte) As Currency
'功能:获取指定病人的记帐划价单金额合计
'参数:byt来源:1-门诊，2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strTab As String
    strTab = IIF(byt来源 = 1, "门诊费用记录", "住院费用记录")
        
    On Error GoTo errH
    
    strSql = "Select Nvl(Sum(实收金额),0) As 划价费用合计 From " & strTab & " Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]"
    
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lng病人ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!划价费用合计
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetStockSet(ByVal lng药房ID As Long, ByVal lng药品ID As Long) As Recordset
'功能:获取药品库存记录集
    Dim strSql As String
    
    If Val(zlDatabase.GetPara(150, glngSys)) = 0 Then '分批药品出库方式：0-按批次先进先出，1-按效期最近先出,效期相同，则再按批次先进先出
        strSql = "Nvl(批次,0)"
    Else
        strSql = "效期,Nvl(批次,0)" '效期为空则排在最后
    End If
    
    '药房不分批药品不管效期(这里的库房一定是药房)
    strSql = "Select Nvl(批次,0) as 批次,Nvl(可用数量,0) as 库存," & _
        " Nvl(零售价,Nvl(Decode(Nvl(实际数量,0),0,0,实际金额/实际数量),0)) as 时价," & _
        " Nvl(实际差价,0) as 实际差价,Nvl(实际金额,0) as 实际金额" & _
        " From 药品库存" & _
        " Where 库房ID=[1] And 药品ID=[2] And Nvl(可用数量,0)>0" & _
        " And 性质=1 And (Nvl(批次,0)=0 Or 效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Order by " & strSql
        
    On Error GoTo errH
    Set GetStockSet = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng药房ID, lng药品ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get时价药品应收金额(ByVal lng药房ID As Long, ByVal lng药品ID As Long, _
    ByRef dblAllTime As Double, ByVal strDec As String, _
    ByRef dblPriceSingle As Double) As Currency
'功能：获取分批时价药品的应收金额（根据不同的出库方式按批次合计）
'参数：
'      strDec-费用金额保留位数
'      dblAllTime-传入为出库总数量(售价数量)，传出如果为0则表示库存足够，否则表示库存不足
'      dblPriceSingle-只有一个批次时返回该批次的单价，避免根据数量先乘再除后因四舍五入引起不同的数量的单价不同
    Dim rsPrice As ADODB.Recordset
    Dim dblPrice As Double, dblCurTime As Double, i As Long
    
    Set rsPrice = GetStockSet(lng药房ID, lng药品ID)
    '时价=总金额/总数量
    dblPrice = 0 '本笔总应收金额
    
    For i = 1 To rsPrice.RecordCount
        If dblAllTime = 0 Then Exit For
        '取小者
        If dblAllTime <= rsPrice!库存 Then
            dblCurTime = dblAllTime
        Else
            dblCurTime = rsPrice!库存
        End If
        If i = 1 Then
            dblPriceSingle = Format(rsPrice!时价, gstrDecPrice)
        Else
            dblPriceSingle = 0
        End If
       
        dblPrice = dblPrice + Format(dblCurTime * Format(rsPrice!时价, gstrDecPrice), strDec)
        dblAllTime = dblAllTime - dblCurTime
        rsPrice.MoveNext
    Next

    Get时价药品应收金额 = dblPrice
End Function

Public Function GetAuditRecord(lng病人ID As Long, lng主页ID As Long) As ADODB.Recordset
'功能：获取指定病人的费用审批项目
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 项目Id,使用限量,已用数量,使用限量-已用数量 可用数量 From 病人审批项目 Where 病人ID=[1] And 主页ID=[2]"
    Set GetAuditRecord = zlDatabase.OpenSQLRecord(strSql, "mdlInExse", lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMoneyInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional curModiMoney As Currency) As ADODB.Recordset
'功能：获取指定病人的剩余额
'参数：
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    
    strSql = "Select Nvl(费用余额,0) as 费用余额,Nvl(预交余额,0) as 预交余额" & _
            " From 病人余额 Where 性质=1 And 类型 = " & IIF(lng主页ID = 0, 1, 2) & " And 病人ID= [1] "
    
    If curModiMoney <> 0 Then   '必须要用Union方式,如果直接去减,在病人余额无记录时,不会返回记录
        strSql = strSql & " Union All  Select -1* " & curModiMoney & " as 费用余额,0 as 预交余额 From Dual"
        strSql = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSql & ")"
    End If
            
    '如果为医保住院病人，则在费用余额中排开预结中的费用(用于报警)
    If lng主页ID <> 0 Then
        strSql = strSql & " Union All " & _
            " Select -1*Nvl(Sum(金额),0) as 费用余额,0 as 预交余额" & _
            " From 保险模拟结算 Where 病人ID=[1] And 主页ID=[2]"
        strSql = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSql & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiUnit(lngPatiID As Long) As Long
'功能：返回病人所属病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select B.当前病区ID From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lngPatiID)
    If Not rsTmp.EOF Then GetPatiUnit = Nvl(rsTmp!当前病区ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AdjustCpt(lngID As Long)
'功能：药品调价
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String

    On Error GoTo errH
    
    strSql = _
        "Select ID From 收费价目" & _
        " Where ((Sysdate Between 执行日期 and 终止日期) Or (Sysdate>=执行日期 And 终止日期 is NULL))" & _
        " And Nvl(变动原因,0)=0 And 收费细目ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lngID)
    Do While Not rsTmp.EOF
        strSql = "zl_药品收发记录_Adjust(" & rsTmp!ID & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "mdlExpense")
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function BillisZeroLog(ByVal strNO As String, ByVal byt来源 As Byte) As Boolean
'功能：判断指定单据是否属于零耗费用登记
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim strTab As String
    strTab = IIF(byt来源 = 1, "门诊费用记录", "住院费用记录")

    On Error GoTo errH

    strSql = "Select 实收金额 From " & strTab & " Where 记录状态 In(0,1,3) And 记录性质=2 And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", strNO)
    BillisZeroLog = True
    For i = 1 To rsTmp.RecordCount
        If Nvl(rsTmp!实收金额, 0) <> 0 Then
            BillisZeroLog = False: Exit For
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillIdentical(ByVal strNO As String, byt来源 As Byte) As Boolean
'功能：判断指定的记帐单据中的状态是否一致,即是否同时存在审核和未审核的内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strTab As String
    
    strTab = IIF(byt来源 = 1, "门诊费用记录", "住院费用记录")
    BillIdentical = True
    
    On Error GoTo errH
    strSql = _
        " Select Count(Distinct 登记时间) as 时间数," & _
        " Sum(Decode(记录状态,0,1,0)) as 未审核," & _
        " Sum(Decode(记录状态,0,0,1)) as 已审核" & _
        " From " & strTab & _
        " Where 记录状态 IN(0,1,3) And NO=[1] And 记录性质=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", strNO)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!未审核, 0) <> 0 And Nvl(rsTmp!已审核, 0) <> 0 Then
            BillIdentical = False
        ElseIf Nvl(rsTmp!时间数, 0) > 1 Then
            BillIdentical = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckValidity(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    Optional ByVal blnAsk As Boolean = True, Optional lng批次 As Long = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查卫生材料的灭菌效期是否过期
    '入参:blnAsk=表示是否询问是否继续,否则为提醒
    '       lng批次:-1表示不根据批次来检查;>=0根据批次来检查效期
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-14 10:21:54
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As New ADODB.Recordset
    Dim curDate As Date, minDate As Date
    Dim strSql As String, strName As String
    
    CheckValidity = True
    
    '仅一次性材料才判断
    '因为可能各批次灭菌效期不同,检查要用到的批次中最小的效期
    strSql = _
        " Select C.名称,Nvl(B.批次,0) as 批次," & _
        "           B.可用数量 as 库存,B.灭菌效期,Sysdate as 时间" & _
        " From 材料特性 A,药品库存 B,收费项目目录 C" & _
        " Where A.材料ID=B.药品ID And A.材料ID=C.ID And A.一次性材料=1" & _
        "       And B.性质=1 And Nvl(B.可用数量,0)>0 And A.灭菌效期 is Not NULL" & _
        "       And A.材料ID=[1] And B.库房ID=[2] " & IIF(lng批次 >= 0, " And nvl(b.批次,0)=[3] ", "") & _
        " Order by Nvl(B.批次,0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlExpense", lng材料ID, lng库房ID, lng批次)
    If Not rsTmp.EOF Then
        strName = rsTmp!名称
        curDate = rsTmp!时间
        minDate = CDate("3000-01-01")
            
        Do While Not rsTmp.EOF
            If rsTmp!灭菌效期 < minDate Then
                minDate = rsTmp!灭菌效期
            End If
            If Nvl(rsTmp!库存, 0) < dbl数量 Then
                dbl数量 = dbl数量 - Nvl(rsTmp!库存, 0)
            Else
                dbl数量 = 0
            End If
            If dbl数量 = 0 Then Exit Do
            rsTmp.MoveNext
        Loop

        If curDate > minDate Then
            If blnAsk Then
                If MsgBox("卫生材料""" & strName & """的灭菌效期""" & Format(minDate, "yyyy-MM-dd") & """已过期,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckValidity = False
                End If
            Else
                MsgBox "提醒：" & vbCrLf & vbCrLf & "卫生材料""" & strName & """的灭菌效期""" & Format(minDate, "yyyy-MM-dd") & """已过期。", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveExecute(ByVal strNO As String, ByVal intFlag As Integer, ByVal blnALL As Boolean, ByVal byt来源 As Byte) As Boolean
'功能：判断费用单据是否包含完全执行或部分执行的内容
'参数：strNO=费用单据号,intFlag=记录性质
'      blnALL=判别单据中是否全部为完全执行或部分执行的内容
'      byt来源:1-门诊，2-住院
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strTab As String
    strTab = IIF(byt来源 = 1, "门诊费用记录", "住院费用记录")
    
    On Error GoTo errH
    strSql = "Select Nvl(Count(ID),0) as 数目" & _
        " From " & strTab & _
        " Where NO=[1] And 记录性质=[2] And 记录状态 IN(0,1,3) And " & IIF(blnALL, " Not", "") & " 执行状态 IN(1,2)"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "HaveExecute", strNO, intFlag)
    
    If blnALL Then
        HaveExecute = (rsTmp!数目 = 0)
    Else
        HaveExecute = (rsTmp!数目 > 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intType As Integer
    Dim dtCurDate As Date, strMaxNo As String
    Dim strYearStr As String
    
    err = 0: On Error GoTo errH:
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    End If
'    ElseIf intNum = 0 Then
'        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
'        Exit Function
'    End If
    GetFullNO = strNO
    
    strSql = "Select 编号规则,Sysdate as 日期,最大号码 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, intNum)
    dtCurDate = Date
    If Not rsTmp.EOF Then
        intType = Val("" & rsTmp!编号规则)
        dtCurDate = rsTmp!日期
        strMaxNo = Nvl(rsTmp!最大号码)
    End If
    strYearStr = PreFixNO
    If strMaxNo = "" Then strMaxNo = strYearStr & "000001"
    If intType = 1 Then
        '按日编号
        strSql = Format(CDate(Format(dtCurDate, "YYYY-MM-dd")) - CDate(Format(dtCurDate, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSql & Format(Right(strNO, 4), "0000")
        Exit Function
    End If
    '按年编号
    If Len(strNO) = 6 Then
        GetFullNO = Left(strMaxNo, 2) & strNO: Exit Function
    End If
    GetFullNO = Left(strMaxNo, 2) & zlLeftPad(Right(strNO, 6), 6, "0")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlLeftPad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按指定长度填制空格
    '返回:返回字串
    '编制:刘兴洪
    '日期:2012-02-22 17:58:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = zlSubstr(strCode, 1, lngLen)
    End If
    zlLeftPad = Replace(strTmp, Chr(0), strChar)
End Function

Private Function zlSubstr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定字串的值,字串中可以包含汉字
    '入参:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '返回:子串
    '编制:刘兴洪
    '日期:2012-02-22 18:00:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    err = 0: On Error GoTo Errhand:
    zlSubstr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    zlSubstr = Replace(zlSubstr, Chr(0), " ")
    Exit Function
Errhand:
    zlSubstr = ""
End Function

Public Function CheckAdviceDrugSurplus(ByVal lng发送号 As Long, Optional ByVal lng医嘱ID As Long) As String
'功能：检查待回退药品医嘱的数量是否大于当前留存的数量
'参数：lng发送号=要回退的发送号
'      lng医嘱ID=要回退的一组药品医嘱的ID，如果不指定由表示批量回退多条医嘱
'返回：提示信息
'说明：护士不能回退医生的操作，所以只涉及住院费用记录(医生才可能发送临嘱为门诊费用)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strMsg As String
    
    On Error GoTo errH
    
    strSql = _
        " Select C.医嘱内容 as 药品,A.收费细目ID as 药品ID,A.病人病区ID as 病区ID,A.执行部门ID as 库房ID,Sum(A.数次) as 回退数量" & _
        " From 住院费用记录 A,病人医嘱发送 B,病人医嘱记录 C" & _
        " Where A.医嘱序号=B.医嘱ID And A.NO=B.NO And A.记录性质=B.记录性质" & _
        " And B.医嘱ID=C.ID And A.收费类别 In('5','6') And A.价格父号 Is Null" & _
        " And B.发送号=[1] And C.诊疗类别 IN('5','6') And (C.相关ID=[2] Or [2]=0)" & _
        " Group by C.医嘱内容,A.收费细目ID,A.病人病区ID,A.执行部门ID"
    strSql = _
        " Select A.药品,D.名称 as 库房,C.住院包装,C.住院单位,A.回退数量,B.留存数量" & _
        " From (" & strSql & ") A,药品留存计划 B,药品规格 C,部门表 D" & _
        " Where A.库房ID=D.ID And A.药品ID=C.药品ID" & _
        " And A.病区ID=B.部门ID(+) And A.库房ID=B.库房ID(+) And A.药品ID=B.药品ID(+) And B.状态(+)=0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckAdviceDrugSurplus", lng发送号, lng医嘱ID)
    Do While Not rsTmp.EOF
        If Nvl(rsTmp!回退数量, 0) > Nvl(rsTmp!留存数量, 0) And Nvl(rsTmp!留存数量, 0) <> 0 Then
            strMsg = strMsg & vbCrLf & "●[" & rsTmp!药品 & "]从""" & rsTmp!库房 & """的回退数量 " & _
                FormatEx(Nvl(rsTmp!回退数量, 0) / Nvl(rsTmp!住院包装, 1), 5) & rsTmp!住院单位 & "，当前留存数量 " & _
                FormatEx(Nvl(rsTmp!留存数量, 0) / Nvl(rsTmp!住院包装, 1), 5) & rsTmp!住院单位
        End If
        rsTmp.MoveNext
    Loop
    
    If strMsg <> "" Then strMsg = "下列药品的回退数量大于留存数量：" & vbCrLf & strMsg & vbCrLf & vbCrLf & "要继续吗？"
    CheckAdviceDrugSurplus = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBalanceRoll(ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, Optional ByVal blnBat As Boolean) As Boolean
'功能：(住院)对要回退的医嘱对应的费用的结帐情况进行检查(一个病人一次住院的)
'参数：blnBat=是否要进行批量回退
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, intInsure As Integer
    
    On Error GoTo errH
        
    '取要回退的记帐NO
    If blnBat Then
        strSql = "Select Distinct 医嘱ID,NO From 病人医嘱发送 Where 记录性质=2 And 发送号=[1]"
    Else
        strSql = "Select Distinct A.医嘱ID,A.NO From 病人医嘱发送 A,病人医嘱记录 B" & _
            " Where A.医嘱ID=B.ID And A.记录性质=2 And A.发送号=[1] And (B.ID=[2] Or B.相关ID=[2])"
    End If
    '取这些NO的结帐情况(非划价未销帐)
    strSql = "Select A.NO,Nvl(A.价格父号,A.序号) as 序号,Sum(Nvl(A.结帐金额,0)) as 结帐金额" & _
        " From 住院费用记录 A,(" & strSql & ") B Where A.NO=B.NO And A.医嘱序号=B.医嘱ID And A.记录性质 IN(2,12) " & _
        " Group by A.NO,Nvl(A.价格父号,A.序号) Having Sum(Nvl(A.结帐金额,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPublic", lng发送号, lng医嘱ID)
    If Not rsTmp.EOF Then
        strSql = "Select A.病人ID,A.险类 From 病案主页 A,病人医嘱记录 B" & _
            " Where Rownum=1 And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPublic", lng医嘱ID)
        If Not rsTmp.EOF Then intInsure = Nvl(rsTmp!险类, 0)
        If intInsure <> 0 Then '先对医保的限制进行检查
            If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, rsTmp!病人ID, intInsure) Then
                MsgBox "该病人为医保病人，要回退医嘱的发送费用中存在已结帐的费用，不能回退。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If gbytBillOpt <> 0 Then
            If gbytBillOpt = 1 Then
                If MsgBox("要回退医嘱的发送费用中存在已结帐的费用，确实要回退吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf gbytBillOpt = 2 Then
                MsgBox "要回退医嘱的发送费用中存在已结帐的费用，不能回退。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    CheckAdviceBalanceRoll = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBalanceRevoke(ByVal lng医嘱ID As Long) As Boolean
'功能：(门诊)对要作废的医嘱对应的费用的结帐情况进行检查
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lng发送号 As Long
    
    If gbytBillOpt = 0 Then
        CheckAdviceBalanceRevoke = True
        Exit Function
    End If
    
    On Error GoTo errH
    
    '医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
    strSql = "Select Distinct 发送号 From 病人医嘱发送" & _
        " Where 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPublic", lng医嘱ID)
    If rsTmp.EOF Then Exit Function
    lng发送号 = rsTmp!发送号
    
    '部份条件见"ZL_病人医嘱记录_作废"
    strSql = "Select A.NO,Nvl(A.价格父号,A.序号) as 序号,Sum(Nvl(A.结帐金额,0)) as 结帐金额" & _
        " From 门诊费用记录 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 I" & _
        " Where A.NO=B.NO And A.记录性质 IN(2,12) And A.记录状态=1 And A.医嘱序号=B.医嘱ID And B.医嘱ID=C.ID" & _
        " And B.记录性质=2 And C.诊疗项目ID=I.ID And B.发送号=[1] And (C.ID=[2] Or C.相关ID=[2])" & _
        " And (" & _
            " A.收费类别 Not In ('5','6','7','E')" & _
            " Or A.收费类别='E' And I.操作类型 Not In ('2','3','4')" & _
            " Or A.收费类别 In ('5','6','7') And Nvl(A.执行状态,0)=0" & _
            " Or Exists(Select 1 From zlParameters Where 系统=[3] And 模块 is NULL And Nvl(私有,0)=0 And 参数号=68 And Nvl(参数值,'0')='0')" & _
            " )" & _
        " Group by A.NO,Nvl(A.价格父号,A.序号) Having Sum(Nvl(A.结帐金额,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPublic", lng发送号, lng医嘱ID, glngSys)
    If Not rsTmp.EOF Then
        If gbytBillOpt = 1 Then
            If MsgBox("要作废医嘱的对应费用中存在已结帐的费用，确实要作废吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        ElseIf gbytBillOpt = 2 Then
            MsgBox "要作废医嘱的对应费用中存在已结帐的费用，不能作废。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckAdviceBalanceRevoke = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBillingRevoke(ByVal lng医嘱ID As Long) As Boolean
'功能：(门诊)对要作废的医嘱对应的记帐费用的审核情况进行检查
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lng发送号 As Long
    
    On Error GoTo errH
    
    '医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
    strSql = "Select Distinct 发送号 From 病人医嘱发送" & _
        " Where 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPublic", lng医嘱ID)
    If rsTmp.EOF Then Exit Function
    lng发送号 = rsTmp!发送号
    
    '部份条件见"ZL_病人医嘱记录_作废"
    strSql = "Select A.NO,A.序号" & _
        " From 门诊费用记录 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 I" & _
        " Where A.NO=B.NO And A.记录性质 IN(2,12) And A.记录状态=1" & _
        " And A.划价人 Is Not NULL And A.划价人<>A.操作员姓名" & _
        " And A.医嘱序号=B.医嘱ID And B.医嘱ID=C.ID And B.记录性质=2" & _
        " And C.诊疗项目ID=I.ID And B.发送号=[1] And (C.ID=[2] Or C.相关ID=[2])" & _
        " And (" & _
            " A.收费类别 Not In ('5','6','7','E')" & _
            " Or A.收费类别='E' And I.操作类型 Not In ('2','3','4')" & _
            " Or A.收费类别 In ('5','6','7') And Nvl(A.执行状态,0)=0" & _
            " Or Exists(Select 1 From zlParameters Where 系统=[3] And 模块 is NULL And Nvl(私有,0)=0 And 参数号=68 And Nvl(参数值,'0')='0')" & _
            " )"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPublic", lng发送号, lng医嘱ID, glngSys)
    If Not rsTmp.EOF Then Exit Function
    
    CheckAdviceBillingRevoke = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRowByFeeItemID(ByRef ObjBillDetails As BillDetails, ByRef lngItemID As Long) As Long
'功能:根据收费项目ID返回其在单据中的行号,如果有重复的,只返回第一个
    Dim i As Long
    
    For i = 1 To ObjBillDetails.Count
        If lngItemID = ObjBillDetails(i).收费细目ID Then
            GetRowByFeeItemID = i: Exit Function
        End If
    Next
End Function

Public Function CheckFeeItemLimitDept(ByVal lngFeeItem As Long) As Boolean
'功能:检查收费项目,如果是主项,是否适用于当前病人科室或病区
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    
    strSql = "Select 科室id From 收费适用科室 Where 项目id = [1] And (Select Count(从项id) From 收费从属项目 Where 主项id = [1]) > 0"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lngFeeItem)
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!科室ID = UserInfo.部门ID Then
                CheckFeeItemLimitDept = True
                Exit For
            End If
            rsTmp.MoveNext
        Next
    Else
        CheckFeeItemLimitDept = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlPatiIS病案已编目(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional blnMsgbox As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取一张单据的实收金额合计,或一张记帐表中指定病人的实收金额合计
    '返回：已编目,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-12 11:26:28
    '说明：28725
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo Errhand:
    strSql = "Select NVL(A.姓名,b.姓名) 姓名 From 病案主页 A,病人信息 B where a.病人id=b.病人id and  A.病人id=[1] and a.主页id=[2] and 编目日期 IS NOT NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查病案是否已经编码", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        zlPatiIS病案已编目 = False
    Else
        zlPatiIS病案已编目 = True
        If blnMsgbox Then
                MsgBox "病人『" & Nvl(rsTemp!姓名) & " 』已经编目,不允许进行记帐或销帐操作!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Function
        End If
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlIs备货材料(ByVal strNO As String, ByVal lng记录性质 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否备货材料记帐
    '入参:strNO-单据号
    '       lng记帐性质:1-收费;2-记帐;
    '出参:
    '返回:如果是备货材料记帐,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-15 11:01:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If lng记录性质 = 1 Then
        strSql = "Select  /*+ rule*/ 1 From 药品收发记录 A,门诊费用记录 B Where A.费用ID=b.ID and A.单据=21 And b.NO=[1] and Mod(b.记录性质,10)=1 and rownum <=1"
    Else
        strSql = "Select  /*+ rule*/ 1 From 药品收发记录 A,住院费用记录 B Where A.费用ID=b.ID and A.单据=21 And b.NO=[1] and b.记录性质=[2] and rownum <=1"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查是否存在备货材料记帐", strNO, lng记录性质)
    zlIs备货材料 = Not rsTemp.EOF
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_vsGrid_Para_Save(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional bln强制保存 As Boolean = False, Optional blnHaveParaPrivs As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '功能:保存vsFlex的宽度到注册表
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '返回:保存成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If blnSaveToDataBase = False Then
        zl_vsGrid_Para_Save = True
        If bln强制保存 = False Then
            If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
    End If
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIF(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    If blnSaveToDataBase Then
        zlDatabase.SetPara strKey, strCol, glngSys, lngModule, blnHaveParaPrivs
    Else
        Call SaveRegInFor(g私有模块, strCaption, strKey, strCol)
    End If
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:从数据库中恢复网格的宽度等信息
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '     blnSaveToDataBase-是否是往数据库中保存参数(如果是往数据库中保存,则强制保存为true,否则根据是否使用个性化风格来确定)
    '     bln强制恢复保存-决定是否将保存注册表的参数值,进行强制恢复
    '返回:恢复成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        '只有在本地注册表中才会处理个性化设置
        zl_vsGrid_Para_Restore = True
        If bln强制恢复保存 = False Then
            If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g私有模块, strCaption, strKey, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
Errhand:
End Function
 




Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
End Sub
Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    err = 0: On Error GoTo Errhand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
Errhand:
End Function
Public Function GetVsGridBoolColVal(ByVal vsGrid As VSFlexGrid, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '功能:获取bool列的值
    '返回:是该单元格为true,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/28
    '------------------------------------------------------------------------------
    Dim strTemp As String
    err = 0: On Error GoTo Errhand:
    With vsGrid
        strTemp = .TextMatrix(lngRow, lngCol)
    End With
    If UCase(strTemp) = UCase("True") Then
        GetVsGridBoolColVal = True: Exit Function
    End If
    GetVsGridBoolColVal = Val(strTemp) <> 0
    Exit Function
Errhand:
End Function
Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub
Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln负数检查 As Boolean = True, Optional bln零检查 As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     bln负数检查     是否进行负数检查
    '     bln零检查         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
   
    Dim dblValue As Double
    If bln零检查 = True Then
        If strInput = "" Then
            ShowMsgBox str项目 & "未输入，请检查!"
            If hWnd <> 0 Then SetFocusHwnd hWnd
            Exit Function
        End If
    End If
    If strInput = "" Then zlDblIsValid = True: Exit Function
    
    If IsNumeric(strInput) = False Then
        MsgBox str项目 & "不是有效的数字格式。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str项目 & "数值过大，不能超过" & 10 ^ intMax - 1 & "。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    If bln负数检查 = True And dblValue < 0 Then
        MsgBox str项目 & "不能输入负数。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str项目 & "数值过小，不能小于-" & 10 ^ intMax - 1 & "位。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    
    If bln零检查 = True And dblValue = 0 Then
        MsgBox str项目 & "不能输入零。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    zlDblIsValid = True
End Function
Public Function Get时价材料应收金额(ByVal lng虚拟库房ID As Long, ByVal lng材料ID As Long, ByVal lng批次 As Long, _
    ByRef dblAllTime As Double, ByVal strDec As String, _
    ByRef dblPriceSingle As Double, Optional bln实际库存 As Boolean = False) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取分批时价材料的应收金额（根据不同的出库方式按批次合计）
    '入参:lng批次-材料批次
    '      strDec-费用金额保留位数
    '      dblAllTime-传入为出库总数量(售价数量)，传出如果为0则表示库存足够，否则表示库存不足
    '      dblPriceSingle-只有一个批次时返回该批次的单价，避免根据数量先乘再除后因四舍五入引起不同的数量的单价不同
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-17 11:08:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsPrice As ADODB.Recordset
    Dim dblPrice As Double, dblCurTime As Double, i As Long
    
    Set rsPrice = GetStuffStockSet(lng虚拟库房ID, lng材料ID, lng批次, bln实际库存)
    '时价=总金额/总数量
    dblPrice = 0 '本笔总应收金额
    
    For i = 1 To rsPrice.RecordCount
        If dblAllTime = 0 Then Exit For
        '取小者
        If dblAllTime <= rsPrice!库存 + IIF(bln实际库存, dblAllTime, 0) Then '+ IIF(bln实际库存, dblAllTime, 0):主要是应用于改费,原因之一是可能可用数量已经没有了,但改费需要加入此数量
            dblCurTime = dblAllTime
        Else
            dblCurTime = rsPrice!库存
        End If
        If i = 1 Then
            dblPriceSingle = Format(rsPrice!时价, gstrDecPrice)
        Else
            dblPriceSingle = 0
        End If
        dblPrice = dblPrice + Format(dblCurTime * Format(rsPrice!时价, gstrDecPrice), strDec)
        dblAllTime = dblAllTime - dblCurTime
        rsPrice.MoveNext
    Next
    Get时价材料应收金额 = dblPrice
End Function

Public Function GetStuffStockSet(ByVal lng虚拟库房ID As Long, _
    ByVal lng材料ID As Long, Optional ByVal lng批次 As Long = -1, _
    Optional bln实际库存 As Boolean = False) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取材料库存记录集
    '入参:bln实际库存-以实际库存为条件(主要是可能因为记帐后,可用库存为零了,但有实际库存,在改费时,对于实价来说,可以计算出价格
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-17 11:10:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    If Val(zlDatabase.GetPara(150, glngSys)) = 0 Then '分批药品出库方式：0-按批次先进先出，1-按效期最近先出,效期相同，则再按批次先进先出
        strSql = "Nvl(批次,0)"
    Else
        strSql = "效期,Nvl(批次,0)" '效期为空则排在最后
    End If
    '不分批材料不管效期
    strSql = "" & _
    "   Select Nvl(批次,0) as 批次,Nvl(可用数量,0) as 库存," & _
    "           Nvl(零售价,Nvl(Decode(Nvl(实际数量,0),0,0,实际金额/实际数量),0)) as 时价," & _
    "           Nvl(实际差价,0) as 实际差价,Nvl(实际金额,0) as 实际金额" & _
    " From 药品库存" & _
    " Where 库房ID=[1] And 药品ID=[2]  " & IIF(lng批次 >= 0, " And NVL(批次,0)=[3] ", "") & IIF(bln实际库存, " And Nvl(实际数量,0)>0", " And Nvl(可用数量,0)>0") & _
    " And 性质=1 And (Nvl(批次,0)=0 Or 效期 is NULL Or 效期>Trunc(Sysdate))" & _
    " Order by " & strSql
    On Error GoTo errH
    Set GetStuffStockSet = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng虚拟库房ID, lng材料ID, lng批次)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function zlIsAllowFeeChange(lng病人ID As Long, lng主页ID As Long, _
   Optional int状态 As Integer = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许费用变动
    '入参:int状态-(-1表示从数据库中读取审核标志进行判断;>0表示,直接根据该状态进行判断)
    '返回:允许变动返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-05-21 15:44:47
    '问题:49501,51612
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    If gbyt病人审核方式 = 0 And gbln未入科禁止记账 = False Then
        ''保持歉容
        zlIsAllowFeeChange = True: Exit Function
    End If
    
    strSql = "" & _
    " Select Nvl(审核标志,0) as 审核标志,nvl(状态,0) as 状态" & _
    " From 病案主页 " & _
    " Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "mdlInExse", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未找到对应的病人信息,不允许进行记录操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '检查未入科病人不允许记账
    If gbln未入科禁止记账 And Val(Nvl(rsTemp!状态)) = 1 Then
        '51612
        MsgBox "病人未入科(第" & lng主页ID & "次住院) ,不能对该病人进行记账或销账操作。", vbInformation, gstrSysName
        Exit Function
    End If
    '审核相关检查
    If gbyt病人审核方式 = 0 Then zlIsAllowFeeChange = True: Exit Function
    
    If int状态 < 0 Then
        int状态 = Val(Nvl(rsTemp!审核标志))
    End If
    '检查相关状态
    If int状态 = 1 Then
        MsgBox "病人在第" & lng主页ID & "次住院中已经开始审核费用,不能对该病人进行费用变动。", vbInformation, gstrSysName
        Exit Function
    End If
    If int状态 = 2 Then
        MsgBox "已经完成了对病人第" & lng主页ID & "次住院费用的审核,不能对该病人进行费用变动。", vbInformation, gstrSysName
        Exit Function
    End If
    zlIsAllowFeeChange = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



