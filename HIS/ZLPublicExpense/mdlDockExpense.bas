Attribute VB_Name = "mdlDockExpense"
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
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public gobjInExse As Object
Private mlng部门编码平均长度 As Long
Public grs收入项目 As ADODB.Recordset
Public glngMainHwnd As Long

Public Function GetFeeKind() As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select 编码, 名称, 简码 From 收费项目类别"
    Set GetFeeKind = gobjDatabase.OpenSQLRecord(strSQL, "获取收费类别")
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    If InStr(1, str收费细目IDs, ",") = 0 Then
        strSQL = "" & _
        "   Select Distinct /*+ Rule*/ 收费细目ID,Nvl(开单科室ID,0) as 开单科室ID,执行科室id " & _
        "   From 收费执行科室 A " & _
        "   Where   A.收费细目ID  =[2] "
    Else
        strSQL = "" & _
        "   Select Distinct /*+ Rule*/ 收费细目ID,Nvl(开单科室ID,0) as 开单科室ID,执行科室id " & _
        "   From 收费执行科室 A," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        "   Where   A.收费细目ID  = j.Column_Value"
    End If
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "获取执行科室信息", Replace(str收费细目IDs, "'", ""), Val(str收费细目IDs))
    If Not rsTmp.EOF Then Set GetServiceDept = rsTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub LoadPatientBaby(ByRef cboBaby As ComboBox, ByVal lngPatient As Long, lngPatientPage As Long)
    Dim rsTmp As ADODB.Recordset, i As Long
    
    cboBaby.Clear
    cboBaby.AddItem "0-病人本人"
    cboBaby.ItemData(cboBaby.NewIndex) = 0
    Call gobjControl.CboSetIndex(cboBaby.hWnd, 0)
    
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
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 序号, 婴儿姓名 From 病人新生儿记录 Where 病人id = [1] And 主页ID = [2]"
    On Error GoTo errH
    Set GetPatientBaby = gobjDatabase.OpenSQLRecord(strSQL, "读取新生儿记录", lngPatient, lngPatientPage)

    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    Dim strSQL As String
    
    strSQL = "Select N.名称" & _
        " From 保险支付项目 M,保险支付大类 N " & _
        " Where M.收费细目ID=[1] And M.险类=[2] And M.大类ID=N.ID"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lng收费细目ID, int险类)
    If rsTmp.RecordCount > 0 Then Get医保大类 = rsTmp!名称
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    Dim lng病种ID As Long, int险类 As Integer, strSQL As String
    Dim strA1 As String, strA2 As String, strB1 As String, strB2 As String
    
    On Error GoTo errH
            
    '先取病人病种,是该病种是否有各类特准项目设置
    strSQL = _
        " Select A.险类,A.病种ID,Nvl(B.大类,0) as 大类,B.性质,Count(*)" & _
        " From 保险帐户 A,保险特准项目 B" & _
        " Where Nvl(A.病种ID,0)=B.病种ID And Nvl(A.病种ID,0)<>0" & _
        " And B.性质 IN(1,2) And A.病人ID=[1]" & _
        " Group by A.险类,A.病种ID,Nvl(B.大类,0),B.性质"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID)
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
    strSQL = ""
    If strA1 <> "" And strA2 <> "" Then
        strSQL = " And (" & strA1 & " Or " & strA2 & ")"
    Else
        If strA1 <> "" Then strSQL = " And " & strA1
        If strA2 <> "" Then strSQL = " And " & strA2
    End If
    If strB1 <> "" Then strSQL = strSQL & " And " & strB1
    If strB2 <> "" Then strSQL = strSQL & " And " & strB2
        
    Get保险特准项目 = strSQL
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get处方职务(lng药品ID As Long) As String
'功能：根据药品ID获取其处方职务
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    Get处方职务 = "00"
    strSQL = "Select Nvl(B.处方职务,'00') as 处方职务 From 药品规格 A,药品特性 B Where A.药名ID=B.药名ID And A.药品ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng药品ID)
    If Not rsTmp.EOF Then Get处方职务 = rsTmp!处方职务
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get处方限量(lngID As Long) As Double
'功能：获取指定药品的处方限量,以零售单位返回。
'参数：lngID=药品ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(A.处方限量,0) as 处方限量" & _
        " From 药品特性 A,药品规格 B Where A.药名ID=B.药名ID And B.药品ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngID)
    If Not rsTmp.EOF Then Get处方限量 = rsTmp!处方限量
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function ItemExistInsure(ByVal lng病人ID As Long, ByVal lng收费细目ID As Long, ByVal int险类 As Integer) As Boolean
'功能：判断收费项目是否设置了保险支付项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    If gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, int险类) Then
        ItemExistInsure = True: Exit Function
    End If
    
    strSQL = "Select 1 From 保险支付项目 Where 收费细目ID=[1] And 险类=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng收费细目ID, int险类)
    ItemExistInsure = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckLimit(ByVal objBill As ExpenseBill, Optional ByVal intRow As Integer, Optional ByVal bln药房单位 As Boolean) As Boolean
'功能：费用单据药品处方限量检查
'说明：
'   1.全部没超过限量，返回真；如有超过药品，则在函数内提示，并返回假。
'   2.记帐表是为每个病人单独检查
    Dim rsTmp As New ADODB.Recordset, strSQL As String
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
    
    strSQL = "Select  /*+ RULE */  A.药品ID,A.剂量系数,B.计算单位 as 剂量单位" & _
        " From 药品规格 A,诊疗项目目录 B," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        " Where A.药名ID=B.ID And A.药品ID  = j.Column_Value"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strItemIDs)
    
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
                                FormatEx(dbl剂量, 5) & rsTmp!剂量单位 & "(" & FormatEx(dblTime, 5) & IIf(bln药房单位, tmpDetail.Detail.药房单位, tmpDetail.Detail.计算单位) & ") 超过处方限量 " & _
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetStockInfo(lng药品ID As Long, bln药房 As Boolean, bln药库 As Boolean, _
    Optional ByVal bln药房单位 As Boolean, Optional str药房包装 As String) As String
'功能：获取药品在各个药房，药库的库存信息
'参数："bln药房/bln药库"至少要有一个设置为真
'返回：描述信息
    Dim strSQL As String, strSQL2 As String
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
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And Instr([1],B.工作性质)>0"
    '药房不分批药品不管效期
    strSQL2 = "Select 部门ID From 部门性质说明 Where 工作性质 IN('西药房','成药房','中药房')"
    '不分批或分批药品
    strSQL = _
        " Select B.编码,B.名称,A.库房ID," & _
        " Nvl(Sum(A.可用数量),0)" & IIf(bln药房单位, "/Nvl(C." & str药房包装 & ",1)", "") & " as 库存" & _
        " From 药品库存 A,(" & strSQL & ") B,药品规格 C" & _
        " Where A.库房ID=B.ID And A.药品ID=C.药品ID" & _
        " And ((A.效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Or (Nvl(C.药房分批,0)=0 And A.库房ID IN(" & strSQL2 & ")))" & _
        " And A.性质=1 And A.药品ID=[2]" & _
        " Group by B.编码,B.名称,A.库房ID,Nvl(C." & str药房包装 & ",1)" & _
        " Having Sum(Nvl(A.可用数量,0))<>0" & _
        " Order By B.编码"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", str性质, lng药品ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!名称 & ":" & rsTmp!库存
        rsTmp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    GetStockInfo = strSQL
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function OverTime() As Boolean
'功能：判断当前是否处于加班时间范围内
'返回：真-当前处于加班时间内,假-不处于
    Dim str上午 As String, str下午 As String
    Dim DateBegin As Date, DateEnd As Date
    Dim curTime As Date
    
    str上午 = gobjDatabase.GetPara(1, glngSys): str下午 = gobjDatabase.GetPara(2, glngSys)
    curTime = CDate(Format(gobjDatabase.Currentdate, "HH:MM:SS"))
    
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
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIf(int记录性质 = 1 Or (int记录性质 = 2 And int病人来源 = 1), "门诊费用记录", "住院费用记录")

    On Error GoTo errH
    
    
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    strSQL = _
        " Select 序号,Sum(数量) as 剩余数量" & _
        " From (" & _
        " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号," & _
        " Avg(Nvl(付数, 1) * 数次) As 数量" & _
        " From " & strTab & _
        " Where NO=[1] And 记录性质=[2]" & _
        " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", str单据号, int记录性质)
    If Not rsTmp.EOF Then GetBillRows = rsTmp.RecordCount
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function BillExistInsure(strNO As String) As Integer
'功能：判断指定的住院记帐单据是否对医保病人记的帐
'参数：strNO=记帐单据号
'返回：如果是则返回病人险类
'说明：1.只管住院医保病人,不管门诊病人的医技记帐
'      2.记帐表只返回第一个病人的险类,单据中也应该只有一种险类
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.险类 From 住院费用记录 A,病案主页 B" & _
        " Where A.记录性质=2 And A.记录状态 IN(0,1,3) And B.险类 is Not NULL" & _
        " And A.NO=[1] And A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then BillExistInsure = rsTmp!险类
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function BillExistDelete(strNO As String, int记录性质 As Integer, int病人来源 As Integer) As Boolean
'功能：判断指定单据是否包含(部分)退费或销帐的内容
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIf(int记录性质 = 1 Or (int记录性质 = 2 And int病人来源 = 1), "门诊费用记录", "住院费用记录")
    
    On Error GoTo errH
    
    strSQL = "Select NO From " & strTab & " Where NO=[1] And 记录性质=[2] And 记录状态=2"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "BillExistDelete", strNO, int记录性质)
    BillExistDelete = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetInsureName(intInsure As Integer) As String
'功能：根据保险类别序号获取保险类别名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 名称 From 保险类别 Where 序号=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", intInsure)
    If Not rsTmp.EOF Then GetInsureName = Nvl(rsTmp!名称)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'功能：获取药品或卫材出库检查的集合
'参数：bytType:0-药品，1-卫材
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '避免出错
    
    strSQL = _
        " Select Distinct A.ID,C.检查方式" & _
        " From 部门表 A,部门性质说明 B," & IIf(bytType = 0, "药品出库检查", "材料出库检查") & " C" & _
        " Where B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 " & IIf(bytType = 0, "IN('中药房','西药房','成药房')", "='发料部门'") & _
        " And C.库房ID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "GetStockCheck")
    For i = 1 To rsTmp.RecordCount
        colStock.Add Nvl(rsTmp!检查方式, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function CheckDisable(objBill As ExpenseBill) As String
'功能：检查单据中的药品的禁忌情况
'返回：药品互相禁忌提示信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInfo As String
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
    
    strSQL = _
        " Select /*+ RULE */  A.组编号,Count(Distinct A.项目ID) as 禁忌数" & _
        " From 诊疗互斥项目 A,药品规格 B," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        " Where A.项目ID=B.药名ID And B.药品ID  = j.Column_Value" & _
        " Having Count(Distinct A.项目ID)>1  " & _
        "  Group by A.组编号"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strIDs)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strGroup = strGroup & "," & rsTmp!组编号
            rsTmp.MoveNext
        Next
        strGroup = Mid(strGroup, 2)
        
        For i = 0 To UBound(Split(strGroup, ","))
            strSQL = _
            "Select /*+ RULE */   Distinct C.类型,C.组编号,D.编码,D.名称,D.规格" & _
            " From 药品规格 A,诊疗项目目录 B,诊疗互斥项目 C,收费项目目录 D," & _
            "          (Select Column_Value From Table(Cast(f_num2list([2]) As Zltools.t_Numlist ))) J " & _
            " Where A.药名ID=B.ID And B.ID=C.项目ID And A.药品ID=D.ID" & _
            "           And C.组编号=[1]" & _
            "           And A.药品ID  = j.Column_Value" & _
            " Order by C.类型,C.组编号,D.编码"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", Val(Split(strGroup, ",")(i)), strIDs)
            
            If Not rsTmp.EOF Then
                rsTmp.Filter = "类型=1"
                If rsTmp.RecordCount > 1 Then
                    k = k + 1
                    strInfo = strInfo & vbCrLf & "第 " & k & " 组(互相慎用)：" & vbCrLf
                    For j = 1 To rsTmp.RecordCount
                        strInfo = strInfo & "[" & rsTmp!编码 & "]" & rsTmp!名称 & IIf(IsNull(rsTmp!规格), "", "(" & rsTmp!规格 & ")") & "                 " & vbCrLf
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Filter = "类型=2"
                If rsTmp.RecordCount > 1 Then
                    blnStop = True
                    k = k + 1
                    strInfo = strInfo & vbCrLf & "第 " & k & " 组(互相禁用)：" & vbCrLf
                    For j = 1 To rsTmp.RecordCount
                        strInfo = strInfo & "[" & rsTmp!编码 & "]" & rsTmp!名称 & IIf(IsNull(rsTmp!规格), "", "(" & rsTmp!规格 & ")") & "                 " & vbCrLf
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function ImportBill(ByVal int来源 As Integer, ByVal str单据号 As String, _
    ByVal int记录性质 As Integer, Optional ByVal bln零耗 As Boolean, _
    Optional ByVal str药品价格等级 As String, _
    Optional ByVal str卫材价格等级 As String, Optional ByVal str普通价格等级 As String) As ExpenseBill
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
    Dim intCurNo As Integer
    Dim int序号 As Integer, blnDo As Boolean, i As Integer
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dbl剩余数量 As Double
    Dim dblAllTime As Double
    
    Dim colSerial As New Collection '用于处理从属父号
    Dim strSQL As String
    Dim strTab As String
    
    Dim lng西药房 As Long, lng成药房 As Long, lng中药房 As Long, str药房 As String
    Dim bln药房单位 As Boolean, str药房单位 As String, str药房包装 As String
    Dim strWherePriceGrade As String
        
    strTab = IIf(int记录性质 = 1 Or (int记录性质 = 2 And int来源 = 1), "门诊费用记录", "住院费用记录")
    
    '价格等级
    If str药品价格等级 <> "" Or str卫材价格等级 <> "" Or str普通价格等级 <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And d.价格等级 = [4])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.类别 || ';') > 0 And d.价格等级 = [5])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And d.价格等级 = [6])" & vbNewLine & _
            "            Or (d.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From 收费价目" & vbNewLine & _
            "                                Where d.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And 价格等级 = [4])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.类别 || ';') > 0 And 价格等级 = [5])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And 价格等级 = [6])))))"
    Else
        strWherePriceGrade = " And d.价格等级 Is Null "
    End If
    
    '缺省药房
    lng西药房 = Val(gobjDatabase.GetPara(IIf(int来源 = 2, "住院", "门诊") & "缺省西药房", glngSys, p医嘱附费管理))
    lng成药房 = Val(gobjDatabase.GetPara(IIf(int来源 = 2, "住院", "门诊") & "缺省成药房", glngSys, p医嘱附费管理))
    lng中药房 = Val(gobjDatabase.GetPara(IIf(int来源 = 2, "住院", "门诊") & "缺省中药房", glngSys, p医嘱附费管理))
    
    '药品单位
    bln药房单位 = Val(gobjDatabase.GetPara("药品单位", glngSys, p医嘱附费管理)) <> 0
    If int来源 = 1 Then
        str药房单位 = "门诊单位": str药房包装 = "门诊包装"
    Else
        str药房单位 = "住院单位": str药房包装 = "住院包装"
    End If
        
    '------------------------------------------------------------------------------------------
    '收费价目关联:新计算价格,如果有多个价格,则一个收费细目ID行就会有多条序号相同的记录
    '价格父号 is NULL:只取每个收费细目ID的第一行(药品只有一行),因为要重算价格
        
    '使用指定的药房读取正确的库存
    str药房 = "Decode(A.收费类别,'5'," & IIf(lng西药房 <> 0, lng西药房, "A.执行部门ID") & "," & _
        "'6'," & IIf(lng成药房 <> 0, lng成药房, "A.执行部门ID") & "," & _
        "'7'," & IIf(lng中药房 <> 0, lng中药房, "A.执行部门ID") & ",A.执行部门ID)"
    
    '药房不分批核算的药品不管效期
    strSQL = _
    " Select X.药品ID,W.材料ID,W.跟踪在用,A.序号 As 序号,A.从属父号,A.NO,A.记录性质,A.记录状态," & IIf(strTab = "住院费用记录", "A.多病人单", " 0 as 多病人单") & ",A.婴儿费,A.费别,A.姓名,A.性别,A.年龄," & _
            IIf(strTab = "住院费用记录", "A.床号,A.病人病区ID,A.主页ID", "A.付款方式 as 床号,0 as 病人病区ID,0 as 主页ID") & _
    "       ,A.标识号,A.病人ID,A.病人科室ID,A.开单部门ID,A.门诊标志,A.加班标志," & _
    "       A.附加标志,A.收费类别,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
    "       A.标准单价 As 标准单价," & str药房 & " as 执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
    "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(F.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
    "       B.屏蔽费别,B.说明,B.执行科室,B.服务对象,Nvl(A.费用类型,B.费用类型) 费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
    "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
    "       Decode(A.收费类别,'4',1,X." & str药房包装 & ") as 药房包装," & _
    "       Decode(A.收费类别,'4',B.计算单位,X." & str药房单位 & ") as 药房单位," & _
    "       Decode(A.收费类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,Nvl(Y.库存,0) As 库存,B.录入限量" & _
    " From " & strTab & " A,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,收费项目别名 F,收费项目别名 E1,材料特性 W,药品规格 X," & _
    "       (Select A.药品ID,A.库房ID,Sum(Nvl(A.可用数量,0)) as 库存 From 药品库存 A" & _
    "        Where A.性质=1 And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
    "               And A.药品ID IN(Select 收费细目ID From " & strTab & " Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1])" & _
                ""
    strSQL = strSQL & _
    "        Group by A.药品ID,A.库房ID) Y" & _
    " Where A.记录性质=[2] And A.记录状态 IN(0,1,3) And A.NO=[1]" & _
    "       And A.价格父号 Is Null And A.收费细目ID=B.ID And A.收费细目ID=D.收费细目ID" & _
    "       And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is NULL)" & _
    "       And A.收费类别=C.编码 And A.收费细目ID=X.药品ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
    "       And A.收费细目ID=Y.药品ID(+) And " & str药房 & "=Y.库房ID(+)" & _
    "       And A.收费细目ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[3]" & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
    "       And ((Sysdate Between D.执行日期 And D.终止日期) Or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & strWherePriceGrade

    strSQL = "Select * From (" & strSQL & ") Order by 序号"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", str单据号, int记录性质, IIf(gSysPara.byt药品名称显示 = 1, 3, 1), _
        str药品价格等级, str卫材价格等级, str普通价格等级)
    
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
                objBillDetail.计算单位 = IIf(IsNull(!计算单位), "", !计算单位)
                
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
                objBillDetail.Detail.服务对象 = Val(Nvl(!服务对象))
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
                    If InStr(",5,6,7,", !收费类别) > 0 Or (!收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        '时价药品计算价格(分批可不分批)
                        dblAllTime = !付数 * !数次 '这里是售价数量
                        If dblAllTime <> 0 Or Nvl(!是否变价, 0) = 0 Then
                            Set rsPrice = gobjDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                        "获取药品当前售价", CLng(!收费细目ID), objBillDetail.执行部门ID, dblAllTime)
                            If rsPrice.EOF Then
                                '获取价格失败
                                If !收费类别 = "4" Then
                                    MsgBox "卫生材料""" & Nvl(!名称) & """获取价格失败！", vbInformation, gstrSysName
                                Else
                                    MsgBox "药品""" & Nvl(!名称) & """获取价格失败！", vbInformation, gstrSysName
                                End If
                                objBillIncome.标准单价 = 0
                            Else
                                strPrice = Nvl(rsPrice!Price) & "|||"
                                varPrice = Split(strPrice, "|")
                                objBillIncome.标准单价 = Val(varPrice(0))
                                dbl剩余数量 = Val(varPrice(2))
                                
                                If dbl剩余数量 <> 0 And Nvl(!是否变价, 0) = 1 Then
                                    '数量未分解完毕
                                    If !收费类别 = "4" Then
                                        MsgBox "时价卫生材料""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "时价药品""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.标准单价 = 0
                                End If
                            End If
                        Else
                            objBillIncome.标准单价 = 0
                        End If
                    ElseIf Nvl(!是否变价, 0) = 1 Then
                        If Abs(!标准单价) > Abs(Val(Nvl(!现价))) Then
                            objBillIncome.标准单价 = Val(Nvl(!缺省价格))
                        Else
                            objBillIncome.标准单价 = !标准单价
                        End If
                    Else
                        objBillIncome.标准单价 = !现价
                    End If
                                        
                    If InStr(",5,6,7,", !收费类别) > 0 And bln药房单位 Then
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(!药房包装, 1), gSysPara.Price_Decimal.strFormt_VB)
                    Else
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价, gSysPara.Price_Decimal.strFormt_VB)
                    End If
                    objBillIncome.现价 = Nvl(!现价, 0) '现价原价对药品变价无用
                    objBillIncome.原价 = Nvl(!原价, 0)
                    objBillIncome.收入项目ID = Nvl(!现收入ID, 0)
                    objBillIncome.收入项目 = Nvl(!收入项目)
                    objBillIncome.收据费目 = Nvl(!现费目)
                    
                    '应收金额=单价*付次*数次
                    objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                    
                    '附加手术费率用计算(所有收入项目)
                    If Nvl(!附加标志, 0) = 1 And Nvl(!收费类别) = "F" Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * Nvl(!附术收费率, 100) / 100
                    End If
                    
                    '加班费用率计算
                    If Nvl(!加班标志, 0) = 1 And Nvl(!加班加价, 0) = 1 Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * (1 + Nvl(!加班加价率, 0) / 100)
                    End If
                    objBillIncome.应收金额 = Format(objBillIncome.应收金额, gSysPara.Money_Decimal.strFormt_VB)
                    
                    '计算实收金额
                    If bln零耗 Then
                        objBillIncome.实收金额 = 0
                    Else
                        If Nvl(!屏蔽费别, 0) = 1 Then
                            objBillIncome.实收金额 = objBillIncome.应收金额
                        Else
                            '使用原定的动态费别
                            objBillIncome.实收金额 = ActualMoney(objBill.费别, !现收入ID, objBillIncome.应收金额, objBillDetail.收费细目ID, _
                                objBillDetail.执行部门ID, !付数 * !数次, IIf(Nvl(!加班标志, 0) = 1 And Nvl(!加班加价, 0) = 1, Nvl(!加班加价率, 0) / 100, 0))
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    Dim intCurNo As Integer
    Dim int序号 As Integer, blnDo As Boolean, i As Integer
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dbl剩余数量 As Double
    Dim dblAllTime As Double

    Dim colSerial As New Collection '用于处理从属父号
    Dim strSQL As String, strStock As String
    Dim strTab As String
    Dim byt单据 As Byte
    
    strTab = IIf(int记录性质 = 1 Or (int记录性质 = 2 And int来源 = 1), "门诊费用记录", "住院费用记录")
    '------------------------------------------------------------------------------------------
    '收费价目关联:新计算价格,如果有多个价格,则一个收费细目ID行就会有多条序号相同的记录
    '价格父号 is NULL:只取每个收费细目ID的第一行(药品只有一行),因为要重算价格
        
    byt单据 = IIf(int记录性质 = 1, 24, 25)
    strStock = _
        " Select A.费用ID,Max( A.药品ID) as 药品ID,Max(A.批次) as 批次,Max(A.商品条码) as 商品条码 ,Max(A.内部条码) as 内部条码 " & _
        " From 药品收发记录 A" & _
        " Where A.NO=[1]  And  单据 =[5] And MOD(A.记录状态,3) in (0,1)" & _
        " Group by A.费用ID "
        
    strStock = "" & _
    "   Select A.费用ID,A.批次,A.商品条码,A.内部条码,sum(b.可用数量) as 可用数量 " & _
    "   From (" & strStock & ") A,药品库存 B " & _
    "   Where A.药品id=b.药品ID(+) And B.库房ID(+)=[4] " & _
    "   Group by A.费用ID,A.批次,A.商品条码,A.内部条码"
    
 
    '药房不分批核算的药品不管效期
    strSQL = _
    " Select X.药品ID,W.材料ID,W.跟踪在用,A.序号 As 序号,A.从属父号,A.NO,A.记录性质,A.记录状态," & IIf(strTab = "住院费用记录", "A.多病人单", " 0 as 多病人单") & ",A.婴儿费,A.费别,A.姓名,A.性别,A.年龄," & _
            IIf(strTab = "住院费用记录", "A.床号,A.病人病区ID,A.主页ID", "A.付款方式 as 床号,0 as 病人病区ID,0 as 主页ID") & _
    "       ,A.标识号,A.病人ID,A.病人科室ID,A.开单部门ID,A.门诊标志,A.加班标志," & _
    "       A.附加标志,A.收费类别,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
    "       A.标准单价 As 标准单价,A.执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
    "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(F.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
    "       B.屏蔽费别,B.说明,B.执行科室,B.服务对象,Nvl(A.费用类型,B.费用类型) 费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
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

    strSQL = "Select * From (" & strSQL & ") Order by 序号"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", str单据号, int记录性质, IIf(gSysPara.byt药品名称显示 = 1, 3, 1), lng虚拟库房ID, byt单据)
    
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
                objBillDetail.计算单位 = IIf(IsNull(!计算单位), "", !计算单位)
                
                objBillDetail.付数 = Nvl(!付数, 1)
                objBillDetail.数次 = Nvl(!数次, 0)
                
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
                objBillDetail.Detail.服务对象 = Val(Nvl(!服务对象))
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
                    If InStr(",5,6,7,", !收费类别) > 0 Or (!收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        '时价药品计算价格(分批可不分批)
                        dblAllTime = !付数 * !数次 '这里是售价数量
                        If dblAllTime <> 0 Or Nvl(!是否变价, 0) = 0 Then
                            Set rsPrice = gobjDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                        "获取药品当前售价", CLng(!收费细目ID), objBillDetail.执行部门ID, dblAllTime)
                            If rsPrice.EOF Then
                                '获取价格失败
                                MsgBox "卫生材料""" & Nvl(!名称) & """获取价格失败！", vbInformation, gstrSysName
                                objBillIncome.标准单价 = 0
                            Else
                                strPrice = Nvl(rsPrice!Price) & "|||"
                                varPrice = Split(strPrice, "|")
                                objBillIncome.标准单价 = Val(varPrice(0))
                                dbl剩余数量 = Val(varPrice(2))
                                
                                If dbl剩余数量 <> 0 And Nvl(!是否变价, 0) = 1 Then
                                    '数量未分解完毕
                                    MsgBox "时价卫生材料""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    objBillIncome.标准单价 = 0
                                End If
                            End If
                        Else
                            objBillIncome.标准单价 = 0
                        End If
                    ElseIf Nvl(!是否变价, 0) = 1 Then
                        If Abs(!标准单价) > Abs(Val(Nvl(!现价))) Then
                            objBillIncome.标准单价 = Val(Nvl(!缺省价格))
                        Else
                            objBillIncome.标准单价 = !标准单价
                        End If
                    Else
                        objBillIncome.标准单价 = !现价
                    End If
                                        
                    objBillIncome.标准单价 = Format(objBillIncome.标准单价, gSysPara.Price_Decimal.strFormt_VB)
                    
                    objBillIncome.现价 = Nvl(!现价, 0) '现价原价对药品变价无用
                    objBillIncome.原价 = Nvl(!原价, 0)
                    objBillIncome.收入项目ID = Nvl(!现收入ID, 0)
                    objBillIncome.收入项目 = Nvl(!收入项目)
                    objBillIncome.收据费目 = Nvl(!现费目)
                    
                    '应收金额=单价*付次*数次
                    objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                    
                    '附加手术费率用计算(所有收入项目)
                    If Nvl(!附加标志, 0) = 1 And Nvl(!收费类别) = "F" Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * Nvl(!附术收费率, 100) / 100
                    End If
                    
                    '加班费用率计算
                    If Nvl(!加班标志, 0) = 1 And Nvl(!加班加价, 0) = 1 Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * (1 + Nvl(!加班加价率, 0) / 100)
                    End If
                    objBillIncome.应收金额 = Format(objBillIncome.应收金额, gSysPara.Money_Decimal.strFormt_VB)
                    
                    '计算实收金额
                    If Nvl(!屏蔽费别, 0) = 1 Then
                        objBillIncome.实收金额 = objBillIncome.应收金额
                    Else
                        '使用原定的动态费别
                        objBillIncome.实收金额 = ActualMoney(objBill.费别, !现收入ID, objBillIncome.应收金额, objBillDetail.收费细目ID, _
                            objBillDetail.执行部门ID, !付数 * !数次, IIf(Nvl(!加班标志, 0) = 1 And Nvl(!加班加价, 0) = 1, Nvl(!加班加价率, 0) / 100, 0))
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
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function GetBillMoney(strNO As String, Optional ByVal int性质 As Integer = 2, Optional ByVal bln门诊 As Boolean = False, Optional lng病人ID As Long) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一张正常单据金额
    '入参:strNO-单据号
    '     int性质=1-收费单,2-记帐单,3-记帐单(自动记帐单),4-挂号单
    '     lng病人ID=病人ID
    '     bln门诊=true:门诊病人:false-住院病人
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-10-17 17:03:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    If lng病人ID = 0 Then
        strSQL = "Select Sum(实收金额) as 金额 From  " & IIf(bln门诊, "门诊费用记录", " 住院费用记录") & " Where NO=[1] And 记录性质=[2] And 记录状态 IN(0,1)"
    Else
        strSQL = "Select Sum(实收金额) as 金额 From " & IIf(bln门诊, "门诊费用记录", " 住院费用记录") & " Where NO=[1] And 记录性质=[2] And 记录状态 IN(0,1) And 病人ID=[3]"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, int性质, lng病人ID)
    If Not rsTmp.EOF Then GetBillMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
 
 
Public Function GetPriceMoneyTotal(ByVal intType As Byte, lng病人ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的划价单金额合计
    '入参:intType:0-门诊;1-住院
    '返回:返回划价总额
    '编制:刘兴洪
    '日期:2014-03-20 18:08:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String, blnAllFee As Boolean
    '记帐报警包含所有住院划价费用
    If intType = 1 Then
        blnAllFee = Val(gobjDatabase.GetPara("记帐报警包含所有住院划价费用", glngSys, 1150)) = 1
        If blnAllFee Then
            strWhere = ""
        Else
            strWhere = " And Nvl(主页ID,0) = (Select Nvl(主页ID,0) From 病人信息 Where 病人ID = [1])"
        End If
    Else
        strWhere = ""
    End If
        
    On Error GoTo errH
    If intType = 1 Then
        strSQL = "" & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
        "   From 住院费用记录 " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] and 门诊标志=2" & strWhere
    Else
        '78226,冉俊明,2014-9-24,修改SQL语句
        '"   From 住院费用记录 and 门诊标志<>2 " & _
        '"   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]"
        strSQL = "" & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计 " & _
        "   From 门诊费用记录  " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]  and 门诊标志<>2" & _
        "   Union ALL   " & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
        "   From 住院费用记录 " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] and 门诊标志<>2 "
        strSQL = "" & _
        "   Select Sum(nvl(划价费用合计,0)) as 划价费用合计  " & _
        "   From ( " & strSQL & ")"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "获取指定病人的划价总额", lng病人ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!划价费用合计
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAuditRecord(lng病人ID As Long, lng主页ID As Long) As ADODB.Recordset
'功能：获取指定病人的费用审批项目
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 项目Id,使用限量,已用数量,使用限量-已用数量 可用数量 From 病人审批项目 Where 病人ID=[1] And 主页ID=[2]"
    Set GetAuditRecord = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetMoneyInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional curModiMoney As Currency) As ADODB.Recordset
'功能：获取指定病人的剩余额
'参数：
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(费用余额,0) as 费用余额,Nvl(预交余额,0) as 预交余额" & _
            " From 病人余额 Where 性质=1 And 类型 = " & IIf(lng主页ID = 0, 1, 2) & " And 病人ID= [1] "
    
    If curModiMoney <> 0 Then   '必须要用Union方式,如果直接去减,在病人余额无记录时,不会返回记录
        strSQL = strSQL & " Union All  Select -1* " & curModiMoney & " as 费用余额,0 as 预交余额 From Dual"
        strSQL = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSQL & ")"
    End If
            
    '如果为医保住院病人，则在费用余额中排开预结中的费用(用于报警)
    If lng主页ID <> 0 Then
        strSQL = strSQL & " Union All " & _
            " Select -1*Nvl(Sum(金额),0) as 费用余额,0 as 预交余额" & _
            " From 保险模拟结算 Where 病人ID=[1] And 主页ID=[2]"
        strSQL = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSQL & ")"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetPatiUnit(lngPatiID As Long) As Long
'功能：返回病人所属病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.当前病区ID From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngPatiID)
    If Not rsTmp.EOF Then GetPatiUnit = Nvl(rsTmp!当前病区ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub AdjustCpt(lngID As Long)
'功能：药品调价
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "zl_药品收发记录_Adjust(" & lngID & ")"
    Call gobjDatabase.ExecuteProcedure(strSQL, "mdlExpense")
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Public Function BillisZeroLog(ByVal strNO As String, ByVal byt来源 As Byte) As Boolean
'功能：判断指定单据是否属于零耗费用登记
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTab As String
    strTab = IIf(byt来源 = 1, "门诊费用记录", "住院费用记录")

    On Error GoTo errH

    strSQL = "Select 实收金额 From " & strTab & " Where 记录状态 In(0,1,3) And 记录性质=2 And NO=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    BillisZeroLog = True
    For i = 1 To rsTmp.RecordCount
        If Nvl(rsTmp!实收金额, 0) <> 0 Then
            BillisZeroLog = False: Exit For
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function BillIdentical(ByVal strNO As String, byt来源 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定的记帐单据中的状态是否一致,即是否同时存在审核和未审核的内容
    '入参:strNO-单据号
    '     byt来源-病人来源:1-门诊;2-住院
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-09 14:25:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIf(byt来源 = 1, "门诊费用记录", "住院费用记录")
    BillIdentical = True
    
    
    On Error GoTo errHandle
    strSQL = _
        " Select Count(Distinct 登记时间) as 时间数," & _
        " Sum(Decode(记录状态,0,1,0)) as 未审核," & _
        " Sum(Decode(记录状态,0,0,1)) as 已审核" & _
        " From " & strTab & _
        " Where 记录状态 IN(0,1,3) And NO=[1] And 记录性质=2"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!未审核, 0) <> 0 And Nvl(rsTmp!已审核, 0) <> 0 Then
            BillIdentical = False
        ElseIf Nvl(rsTmp!时间数, 0) > 1 Then
            BillIdentical = False
        End If
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
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
    Dim strSQL As String, strName As String
    
    CheckValidity = True
    
    '仅一次性材料才判断
    '因为可能各批次灭菌效期不同,检查要用到的批次中最小的效期
    strSQL = _
        " Select C.名称,Nvl(B.批次,0) as 批次," & _
        "           B.可用数量 as 库存,B.灭菌效期,Sysdate as 时间" & _
        " From 材料特性 A,药品库存 B,收费项目目录 C" & _
        " Where A.材料ID=B.药品ID And A.材料ID=C.ID And A.一次性材料=1" & _
        "       And B.性质=1 And Nvl(B.可用数量,0)>0 And A.灭菌效期 is Not NULL" & _
        "       And A.材料ID=[1] And B.库房ID=[2] " & IIf(lng批次 >= 0, " And nvl(b.批次,0)=[3] ", "") & _
        " Order by Nvl(B.批次,0)"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng材料ID, lng库房ID, lng批次)
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function HaveExecute(ByVal strNO As String, ByVal intFlag As Integer, ByVal blnAll As Boolean, ByVal byt来源 As Byte) As Boolean
'功能：判断费用单据是否包含完全执行或部分执行的内容
'参数：strNO=费用单据号,intFlag=记录性质
'      blnALL=判别单据中是否全部为完全执行或部分执行的内容
'      byt来源:1-门诊，2-住院
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    strTab = IIf(byt来源 = 1, "门诊费用记录", "住院费用记录")
    
    On Error GoTo errH
    strSQL = "Select Nvl(Count(ID),0) as 数目" & _
        " From " & strTab & _
        " Where NO=[1] And 记录性质=[2] And 记录状态 IN(0,1,3) And " & IIf(blnAll, " Not", "") & " 执行状态 IN(1,2)"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "HaveExecute", strNO, intFlag)
    
    If blnAll Then
        HaveExecute = (rsTmp!数目 = 0)
    Else
        HaveExecute = (rsTmp!数目 > 0)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:由用户输入的部份单号，返回全部的单号。
    '入参:intNum=项目序号,为0时固定按年产生
    '出参:
    '返回:返回补全的单据号
    '编制:刘兴洪
    '日期:2014-04-09 14:34:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim dtCurDate As Date, strMaxNo As String
    Dim strYearStr As String
    
    Err = 0: On Error GoTo errH:
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期,最大号码 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, intNum)
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
        strSQL = Format(CDate(Format(dtCurDate, "YYYY-MM-dd")) - CDate(Format(dtCurDate, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
        Exit Function
    End If
    '按年编号
    If Len(strNO) = 6 Then
        GetFullNO = Left(strMaxNo, 2) & strNO: Exit Function
    End If
    GetFullNO = Left(strMaxNo, 2) & zlLeftPad(Right(strNO, 6), 6, "0")
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    Err = 0: On Error GoTo Errhand:
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
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select C.医嘱内容 as 药品,A.收费细目ID as 药品ID,A.病人病区ID as 病区ID,A.执行部门ID as 库房ID,Sum(A.数次) as 回退数量" & _
        " From 住院费用记录 A,病人医嘱发送 B,病人医嘱记录 C" & _
        " Where A.医嘱序号=B.医嘱ID And A.NO=B.NO And A.记录性质=B.记录性质" & _
        " And B.医嘱ID=C.ID And A.收费类别 In('5','6') And A.价格父号 Is Null" & _
        " And B.发送号=[1] And C.诊疗类别 IN('5','6') And (C.相关ID=[2] Or [2]=0)" & _
        " Group by C.医嘱内容,A.收费细目ID,A.病人病区ID,A.执行部门ID"
    strSQL = _
        " Select A.药品,D.名称 as 库房,C.住院包装,C.住院单位,A.回退数量,B.留存数量" & _
        " From (" & strSQL & ") A,药品留存计划 B,药品规格 C,部门表 D" & _
        " Where A.库房ID=D.ID And A.药品ID=C.药品ID" & _
        " And A.病区ID=B.部门ID(+) And A.库房ID=B.库房ID(+) And A.药品ID=B.药品ID(+) And B.状态(+)=0"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "CheckAdviceDrugSurplus", lng发送号, lng医嘱ID)
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckAdviceBillingRevoke(ByVal lng医嘱ID As Long) As Boolean
'功能：(门诊)对要作废的医嘱对应的记帐费用的审核情况进行检查
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng发送号 As Long
    Dim intPara As Integer
    
    On Error GoTo errH
    
    '医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
    strSQL = "Select Distinct 发送号 From 病人医嘱发送" & _
        " Where 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1])"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng医嘱ID)
    If rsTmp.EOF Then Exit Function
    lng发送号 = rsTmp!发送号
    
    intPara = Val(gobjDatabase.GetPara(68, glngSys))

   
    '部份条件见"ZL_病人医嘱记录_作废"
    strSQL = "Select A.NO,A.序号" & _
        " From 门诊费用记录 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 I" & _
        " Where A.NO=B.NO And A.记录性质 IN(2,12) And A.记录状态=1" & _
        " And A.划价人 Is Not NULL And A.划价人<>A.操作员姓名" & _
        " And A.医嘱序号=B.医嘱ID And B.医嘱ID=C.ID And B.记录性质=2" & _
        " And C.诊疗项目ID=I.ID And B.发送号=[1] And (C.ID=[2] Or C.相关ID=[2])" & _
        " And (" & _
            " A.收费类别 Not In ('5','6','7','E')" & _
            " Or A.收费类别='E' And I.操作类型 Not In ('2','3','4')" & _
            " Or A.收费类别 In ('5','6','7') And Nvl(A.执行状态,0)=0" & _
            " Or 0=[3])"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng发送号, lng医嘱ID, intPara)
    If Not rsTmp.EOF Then Exit Function
    
    CheckAdviceBillingRevoke = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    Dim strSQL As String, i As Long
    
    strSQL = "Select 科室id From 收费适用科室 Where 项目id = [1] And (Select Count(从项id) From 收费从属项目 Where 主项id = [1]) > 0"

    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItem)
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPatiIS病案已编目(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional blnMsgbox As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取一张单据的实收金额合计,或一张记帐表中指定病人的实收金额合计
    '返回：已编目,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-12 11:26:28
    '说明：28725
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    strSQL = "Select NVL(A.姓名,b.姓名) 姓名 From 病案主页 A,病人信息 B where a.病人id=b.病人id and  A.病人id=[1] and a.主页id=[2] and 编目日期 IS NOT NULL"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查病案是否已经编码", lng病人ID, lng主页ID)
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
    If gobjComlib.ErrCenter = 1 Then Resume
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
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If lng记录性质 = 1 Then
        strSQL = "Select  /*+ rule*/ 1 From 药品收发记录 A,门诊费用记录 B Where A.费用ID=b.ID and A.单据=21 And b.NO=[1] and b.记录性质=[2] and rownum <=1"
    Else
        strSQL = "Select  /*+ rule*/ 1 From 药品收发记录 A,住院费用记录 B Where A.费用ID=b.ID and A.单据=21 And b.NO=[1] and b.记录性质=[2] and rownum <=1"
    End If
    
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查是否存在备货材料记帐", strNO, lng记录性质)
    zlIs备货材料 = Not rsTemp.EOF
    rsTemp.Close
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
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
            If Val(gobjDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
    End If
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    If blnSaveToDataBase Then
        gobjDatabase.SetPara strKey, strCol, glngSys, lngModule, blnHaveParaPrivs
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
            If Val(gobjDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g私有模块, strCaption, strKey, strParaValue)
    Else
        strParaValue = gobjDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    Err = 0: On Error GoTo Errhand:
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
    Err = 0
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
    Err = 0
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
    GetTaskbarHeight = gobjComlib.OS.TaskbarHeight
End Function
Public Function GetVsGridBoolColVal(ByVal vsGrid As VSFlexGrid, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '功能:获取bool列的值
    '返回:是该单元格为true,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/28
    '------------------------------------------------------------------------------
    GetVsGridBoolColVal = gobjComlib.Grid.BoolVal(vsGrid, lngRow, lngCol)
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
    zlDblIsValid = gobjCommFun.DblIsValid(strInput, intMax, bln负数检查, bln零检查, hWnd, str项目)
End Function

Public Function zlIsAllowFeeChange(lng病人ID As Long, lng主页ID As Long, _
   Optional int状态 As Integer = -1, Optional blnNotMsgBox As Boolean, Optional strOutErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许费用变动
    '入参:int状态-(-1表示从数据库中读取审核标志进行判断;>0表示,直接根据该状态进行判断)
    '    blnNotMsgbox-是否显示消费提示框
    '出参:strOutErrMsg-返回错误信息
    '返回:允许变动返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-05-21 15:44:47
    '问题:49501,51612
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If gSysPara.byt病人审核方式 = 0 And gSysPara.bln未入科禁止记账 = False Then
        ''保持歉容
        zlIsAllowFeeChange = True: Exit Function
    End If
    
    strSQL = "" & _
    " Select Nvl(审核标志,0) as 审核标志,nvl(状态,0) as 状态" & _
    " From 病案主页 " & _
    " Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        strOutErrMsg = "未找到对应的病人信息,不允许进行费用变动操作!"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '检查未入科病人不允许记账
    If gSysPara.bln未入科禁止记账 And Val(Nvl(rsTemp!状态)) = 1 Then
        '51612
        strOutErrMsg = "病人未入科(第" & lng主页ID & "次住院) ,不能对该病人进行记账或销账操作。"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '审核相关检查
    If gSysPara.byt病人审核方式 = 0 Then zlIsAllowFeeChange = True: Exit Function
    
    If int状态 < 0 Then
        int状态 = Val(Nvl(rsTemp!审核标志))
    End If
    '检查相关状态
    If int状态 = 1 Then
        strOutErrMsg = "病人在第" & lng主页ID & "次住院中已经开始审核费用,不能对该病人进行费用变动。"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If int状态 = 2 Then
        strOutErrMsg = "已经完成了对病人第" & lng主页ID & "次住院费用的审核,不能对该病人进行费用变动。"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zlIsAllowFeeChange = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回大写的单据号年前缀
    '返回:年前缀
    '编制:刘兴洪
    '日期:2014-04-09 14:34:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PreFixNO = gobjComlib.zlStr.PreFixNO(curDate)
End Function



Public Function GetPatiUnitID(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Long
'功能：根据病人获取对应的病区ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 当前病区ID as 病区ID From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    GetPatiUnitID = Nvl(rsTmp!病区ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Check上班安排(ByVal bln药房 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医院的科室是否使用了上班安排
    '入参:：bln药房=是检查药房上班还是其它科室
    '出参:
    '返回:启用了上班时间的返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-09 14:52:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Static bln药房Load As Boolean
    Static bln药房Last As Boolean
    Static bln非药Load As Boolean
    Static bln非药Last As Boolean
    
    If bln药房 Then '是否有安排只需读取一次
        If bln药房Load Then Check上班安排 = bln药房Last: Exit Function
    Else
        If bln非药Load Then Check上班安排 = bln非药Last: Exit Function
    End If
    
    On Error GoTo errH
    
    If bln药房 Then
        strSQL = "Select 1 From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 IN('西药房','成药房','中药房') And Rownum<2"
    Else
        strSQL = "Select 1 From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 Not IN('西药房','成药房','中药房') And Rownum<2"
    End If
    Call gobjDatabase.OpenRecordset(rsTmp, strSQL, "Check上班安排")
    Check上班安排 = rsTmp.RecordCount > 0
    
    If bln药房 Then
        bln药房Load = True: bln药房Last = Check上班安排
    Else
        bln非药Load = True: bln非药Last = Check上班安排
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function Get操作员部门ID(ByVal int服务对象 As Integer, Optional ByVal lng默认部门 As Long = 0) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取操作员所属服务对指定对象的部门，缺省部门优先
    '返回:返回操作员的缺省部门ID
    '编制:刘兴洪
    '日期:2014-04-09 14:53:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    
    If blnNew Then
        strSQL = "Select Distinct B.部门ID,Nvl(B.缺省,0) as 缺省,C.服务对象 From 部门人员 B,部门性质说明 C" & _
            " Where B.人员ID = [1] And B.部门ID=C.部门ID" & _
            " Order by 缺省 Desc"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    
    '74794,冉俊明,2014-7-18,护士在记账时候调用成套方案时未使用成套方案内的执行科室
    If lng默认部门 <> 0 Then
        rsTmp.Filter = "(服务对象 = 3 and 部门ID = " & lng默认部门 & ") " & _
                    "or (服务对象 = " & int服务对象 & " and 部门ID = " & lng默认部门 & ")"
        If Not rsTmp.EOF Then Get操作员部门ID = rsTmp!部门ID: Exit Function
    End If
    
    rsTmp.Filter = "服务对象 = 3 or 服务对象 = " & int服务对象
    
    If Not rsTmp.EOF Then
        Get操作员部门ID = rsTmp!部门ID
    Else
        Get操作员部门ID = UserInfo.部门ID
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetPatiDayMoneyDetail(rsMoneyDay As ADODB.Recordset, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt来源 As Byte, _
         Optional ByVal lng诊疗项目ID As Long, Optional ByVal lng收费细目ID As Long, Optional ByVal date首日不收取 As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人当天及之后医嘱产生的费用项目明细
    '入参:lng主页ID=住院病人才使用
    '      byt来源:1-门诊(含住院临嘱发送到门诊)，2-住院
    '      str首次时间=本次医嘱发送，首次执行的时间
    '      date首日不收取=用于添加首日不收取的项目天数，但频率又不是每天一次的，实际是每天一次的，例如隔日一次，每24小时一次等

    '出参:rsMoneyDay，包含"诊疗项目ID,收费项目ID,执行部门ID,执行否,收费时间"字段
    '返回: 如果是发送当天之前的医嘱，则本过程暂时没有考虑这种情况，检查当天是否已执行时会检查不到
    '编制:刘兴洪
    '日期:2014-04-09 14:55:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strToDay As String, strDay As String
        
    On Error GoTo errH
    
    If lng诊疗项目ID = 0 Then
        Set rsMoneyDay = New ADODB.Recordset '用于清除Filter属性
        strToDay = Format(gobjDatabase.Currentdate, "yyyy-MM-dd")
        '执行判断：
        '1.传入的是将填定到费用记录中的执行部门，因此也以费用记录中的执行部门为准判断。
        '2.除和跟踪卫材外，医嘱费用的执行科室与医嘱执行科室相同；以后如果不同了，该函数也可以适应
        '3.医嘱执行时，对应费用的执行状态也会同步标记。
        '4.首次不收的项目，如果频率是一天只收一次，则没有产生费用记录（但有医嘱发送记录）,需要读出来当成已生成的，以便其他首次不收的项目判断
        If byt来源 = 1 Then
            strSQL = "Select A.诊疗项目ID,C.收费细目ID as 收费项目ID,C.执行部门ID,Decode(Nvl(C.执行状态,0),0,0,1) as 执行否,To_Char(C.发生时间,'yyyy-mm-dd') as 收费时间,0 as 收费方式" & _
                " From 病人医嘱记录 A,病人医嘱发送 B,门诊费用记录 C" & _
                " Where A.病人ID=[1] And Nvl(A.主页ID,0) = [2] And a.医嘱期效 = 1 And A.ID=B.医嘱ID And B.记录性质=C.记录性质 And B.NO=C.NO" & _
                " And B.医嘱ID=C.医嘱序号 And C.记录状态 IN(0,1) And C.发生时间>=[3]" & _
                " Union " & _
                " Select A.诊疗项目ID,D.收费细目id,D.执行科室ID as 执行部门ID,0 as 执行否,To_Char(B.首次时间,'yyyy-mm-dd') as 收费时间,-1 as 收费方式" & _
                " From 病人医嘱记录 A,病人医嘱发送 B,病人医嘱计价 D" & _
                " Where A.病人ID=[1] And Nvl(A.主页ID,0) = [2] And a.医嘱期效 = 1 " & _
                " And A.ID=B.医嘱ID And NVL(B.首次时间,a.开始执行时间)>=[3] And A.ID=D.医嘱ID And D.收费方式=7" & vbNewLine & _
                " And Not Exists (Select 1 From 门诊费用记录 C Where c.收费细目id=d.收费细目id  And b.记录性质 = c.记录性质 And b.No = c.No And a.Id = c.医嘱序号)" & vbNewLine & _
                " Order by 诊疗项目ID,收费项目ID"
            Set rsMoneyDay = gobjDatabase.OpenSQLRecord(strSQL, "读取当天及后续的医嘱", lng病人ID, lng主页ID, CDate(strToDay))
            Set rsMoneyDay = gobjDatabase.CopyNewRec(rsMoneyDay)
        Else
            '临嘱：病人医嘱记录.上次执行时间为空
            '长嘱，其他医嘱的相同费用，可能不同时间多次发送,Union去除了重复记录
            '首次不收的项目，如果频率是一天只收一次，则没有产生费用记录（但有医嘱发送记录）,需要读出来当成已生成的，以便其他首次不收的项目判断
            strSQL = "Select a.诊疗项目id, c.收费细目id As 收费项目id, c.执行部门id, Decode(Nvl(c.执行状态, 0), 0, 0, 1) As 执行否," & vbNewLine & _
                "     Decode(a.医嘱期效, 0, b.首次时间, c.发生时间) As 首次时间, Decode(b.首次时间,null, 1,Trunc(b.末次时间) - Trunc(b.首次时间) + 1) As 天数,0 as 收费方式" & vbNewLine & _
                "From 病人医嘱记录 A, 病人医嘱发送 B, 住院费用记录 C" & vbNewLine & _
                "Where a.病人id = [1] And a.主页id = [2] And a.Id = b.医嘱id And b.记录性质 = c.记录性质 And b.No = c.No And b.医嘱id = c.医嘱序号 And" & vbNewLine & _
                "      c.记录状态 In (0, 1) And ((b.首次时间 > [3] Or b.末次时间 > [3]) Or a.医嘱期效 = 1 And C.发生时间 >= [3])" & vbNewLine & _
                " Union " & vbNewLine & _
                "Select a.诊疗项目id, D.收费细目id, D.执行科室ID as 执行部门id, 0 As 执行否," & vbNewLine & _
                "     b.首次时间, Decode(a.医嘱期效, 0, Trunc(b.末次时间) - Trunc(b.首次时间) + 1, 1) As 天数,-1 as 收费方式" & vbNewLine & _
                "From 病人医嘱记录 A, 病人医嘱发送 B, 病人医嘱计价 D" & vbNewLine & _
                "Where a.病人id = [1] And a.主页id = [2]" & vbNewLine & _
                "   And a.Id = b.医嘱id And ((b.首次时间 > [3] Or b.末次时间 > [3]) Or (a.医嘱期效 = 1 And b.首次时间 is null and a.开始执行时间 >= [3]))" & vbNewLine & _
                "   And A.ID=D.医嘱ID And D.收费方式=7" & vbNewLine & _
                " And Not Exists (Select 1 From 住院费用记录 C Where c.收费细目id=d.收费细目id  And b.记录性质 = c.记录性质 And b.No = c.No And a.Id = c.医嘱序号)" & vbNewLine & _
                "Order By 诊疗项目id, 收费项目id"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "读取当天及后续的医嘱", lng病人ID, lng主页ID, CDate(strToDay))
            '根据开始时间和天数，将记录集按执行时间分成多条记录
            Set rsMoneyDay = InitPatiExecDays
                    
            For i = 1 To rsTmp.RecordCount
                For j = 1 To rsTmp!天数
                    If j = 1 Then
                        strDay = Format(rsTmp!首次时间, "yyyy-MM-dd")
                    Else
                        strDay = Format(DateAdd("d", j - 1, CDate(rsTmp!首次时间)), "yyyy-MM-dd")
                    End If
                    If strDay >= strToDay Then
                        rsMoneyDay.Filter = "诊疗项目ID=" & Val("" & rsTmp!诊疗项目ID) & " And 收费项目ID=" & Val("" & rsTmp!收费项目ID) & _
                                            " And 收费时间='" & strDay & "' And 执行否=" & Val("" & rsTmp!执行否) & " And 收费方式=" & Val("" & rsTmp!收费方式)
                        If rsMoneyDay.RecordCount = 0 Then
                            rsMoneyDay.AddNew
                            rsMoneyDay!诊疗项目ID = Val("" & rsTmp!诊疗项目ID)
                            rsMoneyDay!收费项目ID = Val("" & rsTmp!收费项目ID)
                            rsMoneyDay!执行部门ID = Val("" & rsTmp!执行部门ID)
                            rsMoneyDay!执行否 = Val("" & rsTmp!执行否)
                            rsMoneyDay!收费方式 = Val("" & rsTmp!收费方式)
                            rsMoneyDay!收费时间 = strDay
                            rsMoneyDay.Update
                        End If
                    End If
                Next
                rsTmp.MoveNext
            Next
            rsMoneyDay.Filter = ""
        End If
    Else
        '门诊发送时用于判断每天首次不收取的项目当天是否执行次数=1,如果=1且没有收费，说明当天首次已经没有收取了
        strSQL = "Select d.执行科室id As 执行部门id" & vbNewLine & _
                "From 病人医嘱记录 A,病人医嘱发送 B, 病人医嘱计价 D" & vbNewLine & _
                "Where A.病人ID=[1] And Nvl(A.主页ID,0) = [2] And a.Id = b.医嘱id And A.id = d.医嘱id And A.诊疗项目ID = [6] And d.收费方式 = 7 And d.收费细目id = [3] And Not Exists" & vbNewLine & _
                " (Select 1" & vbNewLine & _
                "       From " & IIf(byt来源 = 1, "门诊费用记录", "住院费用记录") & " C" & vbNewLine & _
                "       Where c.收费细目id = d.收费细目id And b.记录性质 = c.记录性质 And b.No = c.No And d.医嘱id = c.医嘱序号) And" & vbNewLine & _
                "      Zl_Adviceexecount(d.医嘱id, [4], [5],1) = 1"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "读取当天及后续的医嘱", lng病人ID, lng主页ID, lng收费细目ID, CDate(Format(date首日不收取, "yyyy-MM-dd")), CDate(Format(date首日不收取, "yyyy-MM-dd 23:59:59")), lng诊疗项目ID)
        If rsTmp.RecordCount > 0 Then
            rsMoneyDay.Filter = "诊疗项目ID=" & lng诊疗项目ID & " And 收费项目ID=" & lng收费细目ID & _
                                " And 收费时间='" & Format(date首日不收取, "yyyy-MM-dd") & "' And 执行否=0" & " And 收费方式=-1"
            If rsMoneyDay.RecordCount = 0 Then
                rsMoneyDay.AddNew
                rsMoneyDay!诊疗项目ID = lng诊疗项目ID
                rsMoneyDay!收费项目ID = lng收费细目ID
                rsMoneyDay!执行部门ID = Val("" & rsTmp!执行部门ID)
                rsMoneyDay!执行否 = 0
                rsMoneyDay!收费方式 = -1
                rsMoneyDay!收费时间 = Format(date首日不收取, "yyyy-MM-dd")
                rsMoneyDay.Update
            End If
        End If
    End If
    
    GetPatiDayMoneyDetail = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Private Function InitPatiExecDays() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医嘱相关费用执行的记录集
    '返回:医嘱相关费用执行的记录集
    '编制:刘兴洪
    '日期:2014-04-09 14:56:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "诊疗项目ID", adBigInt
    rsTmp.Fields.Append "收费项目ID", adBigInt
    rsTmp.Fields.Append "执行部门ID", adBigInt
    rsTmp.Fields.Append "收费方式", adInteger
    rsTmp.Fields.Append "执行否", adInteger
    rsTmp.Fields.Append "收费时间", adVarChar, 10
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set InitPatiExecDays = rsTmp
End Function


Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断输入金额是否在原价和现从限定的范围内
    '入参:varL=原价,varR=现价,varI=输入金额
    '返回:如果不在范围内,则为提示信息,否则为空串
    '编制:刘兴洪
    '日期:2014-04-09 15:44:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '如果数值符号相同,则用绝对值判断
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "输入的价格绝对值不在范围(" & FormatEx(Abs(varL), 5) & "-" & FormatEx(Abs(varR), 5) & ")内."
        End If
    Else
        '如果符号不相同,则用原始范围判断
        If varI < varL Or varI > varR Then
            CheckScope = "输入的价格值不在范围(" & FormatEx(varL, 5) & "-" & FormatEx(varR, 5) & ")内."
        End If
    End If
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查部门信息是否加载编码
    '返回:显示编码,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 13:11:01
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    If mlng部门编码平均长度 = 0 Then
        strSQL = "Select Avg(length(编码)) As 长度 From 部门表"
        On Error GoTo errH
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "取部门编码的平均长度")
        mlng部门编码平均长度 = Val(Nvl(rsTemp!长度))
    End If
    '由于编码长度可能过长,无法显示部门的名称,因此自动显示和不显示编码,当大于5时,不显示.小于5时,显示
   zlIsShowDeptCode = mlng部门编码平均长度 <= 5
      
   Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get收费执行科室ID(ByVal lng病人ID As Long, lng主页ID As Long, _
    ByVal str类别 As String, ByVal lng项目id As Long, ByVal int执行科室 As Integer, _
    ByVal lng病人科室ID As Long, ByVal lng开单科室id As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal lng执行科室ID As Long, _
    Optional ByVal bytMode As Byte, Optional ByVal bytCallBy As Byte, _
    Optional ByVal int调用场合 As Integer = 1, _
    Optional lng成套缺省执行科室 As Long = 0) As Long
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费项目的执行科室
    '入参:int范围=1.门诊,2-住院
    '      lng执行科室ID=指定的缺省执行科室ID(用于药品和卫材)
    '      bytMode=1-要返回缺省值,0-其它
    '      bytCallBy=0-医嘱程序调用,1-附费程序调用
    '      int调用场合=1-门诊,2-住院
    '      lng成套缺省执行科室-缺省执行科室ID
    '出参:
    '返回:返回指定的执行科室ID
    '编制:刘兴洪
    '日期:2014-04-09 13:58:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim str药房 As String, lng药房 As Long
    Dim lng病人病区ID As Long, bytDay As Byte
    
    On Error GoTo errH
    
    If str类别 = "4" Then
        lng药房 = Val(gobjDatabase.GetPara(IIf(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省发料部门", glngSys, _
            IIf(bytCallBy = 1, p医嘱附费管理, IIf(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达))))
        
        '有执行科室设置时
        strSQL = _
            " Select Distinct" & _
            "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[2]   " & _
            "       Or Exists(select 1 From 病区科室对应 M where A.开单科室ID=M.病区ID And M.科室ID=[2] ))" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int范围, lng病人科室ID, lng项目id)
        If Not rsTmp.EOF Then
            If bytMode = 1 Then Get收费执行科室ID = rsTmp!执行科室ID  '如果都没有，则返回第一个可用的执行科室
            
            '1:缺省为指定的(医嘱的)执行科室,不管是否服务于病人科室
            rsTmp.Filter = "执行科室ID=" & lng执行科室ID
            
            '2.缺省为参数指定的缺省科室
            If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng药房
            
            '3:其它可服务于病人科室的执行科室
            If rsTmp.EOF Then
                '2.0 如果成套中存在缺省的执行科室,则缺省为成套指定的缺省科室
                If lng成套缺省执行科室 <> 0 Then
                    rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                    If Not rsTmp.EOF Then
                            Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                    End If
                End If
                '2.1:尝试缺省为病人科室
                If lng执行科室ID <> lng病人科室ID And lng药房 <> lng病人科室ID Then
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人科室ID
                End If
                '3.2:尝试缺省为病人病区
                If rsTmp.EOF And lng主页ID <> 0 Then
                    lng病人病区ID = GetPatiUnitID(lng病人ID, lng主页ID)
                    If lng病人病区ID <> 0 And lng病人病区ID <> lng病人科室ID And lng病人病区ID <> lng执行科室ID And lng病人病区ID <> lng药房 Then
                        rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人病区ID
                    End If
                End If
            End If
            '3.3:可服务于病人科室的一个执行科室
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            
            '3.4可服务于所有科室的当前病人科室执行
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0 And 执行科室ID=" & lng病人科室ID
            
            '4:如果都没有，则返回0用于检查
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        If str类别 = "5" Then
            str药房 = "西药房"
            lng药房 = Val(gobjDatabase.GetPara(IIf(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省西药房", glngSys, _
                IIf(bytCallBy = 1, p医嘱附费管理, IIf(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达))))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(gobjDatabase.GetPara(IIf(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省成药房", glngSys, _
                IIf(bytCallBy = 1, p医嘱附费管理, IIf(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达))))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(gobjDatabase.GetPara(IIf(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省中药房", glngSys, _
                IIf(bytCallBy = 1, p医嘱附费管理, IIf(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达))))
        End If
        
        '药品从系统指定的储备药房中找
        If Not Check上班安排(True) Then
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                " And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(gobjDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                " And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        End If
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str药房, int范围, lng病人科室ID, lng项目id, bytDay)
        If Not rsTmp.EOF Then
            If lng成套缺省执行科室 <> 0 Then
                rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                If Not rsTmp.EOF Then
                        Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                End If
            End If
            Get收费执行科室ID = rsTmp!执行科室ID
            rsTmp.Filter = "执行科室ID=" & lng执行科室ID
            If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng药房
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    Else
        Select Case int执行科室
            Case 0 '0-无明确科室
                '1 成套项目选择且存在缺省的执行科室的 成套项目的执行部门ID
                If lng成套缺省执行科室 <> 0 Then
                    Get收费执行科室ID = lng成套缺省执行科室: Exit Function
                End If
                '101736,手工记帐缺省执行科室
                '2 收费项目.缺省科室(手工记帐缺省执行科室)
                If int范围 = 2 Then
                    strSQL = "Select a.执行科室id" & vbNewLine & _
                            " From 收费执行科室 A, 部门表 C" & vbNewLine & _
                            " Where a.执行科室id + 0 = c.Id And a.收费细目id = [1]" & vbNewLine & _
                            "       And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
                            "       And (c.站点 = '" & gstrNodeNo & "' Or c.站点 Is Null)" & vbNewLine & _
                            "       And a.病人来源 = [2] And a.开单科室id Is Null"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng项目id, 2)
                    If Not rsTmp.EOF Then
                        If Val(Nvl(rsTmp!执行科室ID)) <> 0 Then
                            Get收费执行科室ID = Val(Nvl(rsTmp!执行科室ID)): Exit Function
                        End If
                    End If
                    '3 病人科室
                    If lng病人科室ID <> 0 Then Get收费执行科室ID = lng病人科室ID: Exit Function
                    '4 开单科室
                    If lng开单科室id <> 0 Then Get收费执行科室ID = lng开单科室id: Exit Function
                End If
                '5 操作员所属部门ID
                Get收费执行科室ID = Get操作员部门ID(int范围)
            Case 1 '1-病人所在科室
                Get收费执行科室ID = lng病人科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get收费执行科室ID = lng病人科室ID
                Else
                    Get收费执行科室ID = GetPatiUnitID(lng病人ID, lng主页ID)
                End If
            Case 3 '3-操作员所在科室
                Get收费执行科室ID = Get操作员部门ID(int范围, lng成套缺省执行科室)
            Case 4 '4-指定科室
                strSQL = "Select Distinct Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID,Decode(A.病人来源,Null,2,1) as 排序" & _
                    " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                    " Where A.收费细目ID=[1] And A.执行科室ID=B.部门ID" & _
                    " And B.服务对象 IN([2],3) And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                    " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                    " And A.执行科室ID=C.ID And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                    " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " Order by 排序" '默认科室优先
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id, int范围, lng病人科室ID)
                If Not rsTmp.EOF Then
                    If lng成套缺省执行科室 <> 0 Then
                         rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                         If Not rsTmp.EOF Then
                                 Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                         End If
                     End If
                    Get收费执行科室ID = rsTmp!执行科室ID
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                    If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
                End If
            Case 6 '6-开单人所在科室
                Get收费执行科室ID = lng开单科室id
        End Select
        If Get收费执行科室ID = 0 Then Get收费执行科室ID = Get操作员部门ID(int范围)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function PatiCanBilling(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String, Optional ByVal lngModual As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定病人是否具有相关权限
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     strPrivs-权限串
    '     lngModual-模块号
    '返回:具有相关权限,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-09 14:13:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, ";出院未结强制记帐;") > 0 And InStr(strPrivs, ";出院结清强制记帐;") > 0 Then Exit Function
    
    On Error GoTo errH
    strSQL = "Select NVL(B.姓名,A.姓名) 姓名,B.出院日期,B.状态,X.费用余额" & _
        " From 病人信息 A,病案主页 B,病人余额 X" & _
        " Where A.病人ID=B.病人ID And A.病人ID=X.病人ID(+) And X.类型(+) = 2" & _
        " And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!出院日期) And Nvl(rsTmp!状态, 0) <> 3 Then Exit Function
        If InStr(strPrivs, ";出院未结强制记帐;") = 0 Then
            If Nvl(rsTmp!费用余额, 0) <> 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用未结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If InStr(strPrivs, ";出院结清强制记帐;") = 0 Then
            If Nvl(rsTmp!费用余额, 0) = 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用已结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If lngModual = p医嘱附费管理 Or lngModual = p住院医嘱发送 Or lngModual = p住院医嘱下达 Then
            '68081不允许出院病人处理医嘱费用
            strMsg = """" & rsTmp!姓名 & """已经出院(或预出院)，不能对该病人的医嘱进行发送、超期收回、执行、回退。"
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function



Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal cur金额 As Currency, ByVal str类别 As String, ByVal str类别名 As String) As Boolean
'功能：当执行完成有自动审核的费用时，对病人费用进行记帐报警。
'参数：str类别="CDE..."，报警金额涉及到的收费类别
'      str类别名="检查,检验,..."，对应的类别名用于提示
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSQL As String, intR As Integer, i As Long
    Dim cur当日 As Currency
    
    On Error GoTo errH
    
    If lng主页ID <> 0 Then
        '住院病人报警
        strSQL = _
            " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1] And 类型 = 2" & _
            " Union ALL" & _
            " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
        strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
        
        strSQL = "Select NVL(B.姓名,A.姓名) 姓名, Nvl(B.住院号,A.住院号) As 住院号, Nvl(B.出院病床,A.当前床号) As 床号,zl_PatiWarnScheme(A.病人ID,B.主页ID) as 适用病人,C.剩余款," & _
            " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额" & _
            " From 病人信息 A,病案主页 B,(" & strSQL & ") C" & _
            " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+)" & _
            " And A.病人ID=[1] And B.主页ID=[2]"
        Set rsPati = gobjDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病人ID, lng主页ID)
    Else
        '其他按门诊报警
        strSQL = "Select 病人ID,预交余额,费用余额 From 病人余额 Where 性质=1 And 病人ID=[1] And 类型 = 1"
        strSQL = "Select A.姓名,A.住院号,A.当前床号 As 床号,zl_PatiWarnScheme(A.病人ID) as 适用病人,A.担保额," & _
            " Nvl(B.预交余额,0)-Nvl(B.费用余额,0)+Nvl(E.帐户余额,0) as 剩余款" & _
            " From 病人信息 A,(" & strSQL & ") B,医保病人关联表 D,医保病人档案 E" & _
            " Where A.病人ID=B.病人ID(+) And A.病人id = D.病人id(+) And A.险类=D.险类(+)" & _
            " And D.险类=E.险类(+) And D.中心=E.中心(+) And D.医保号=E.医保号(+) And D.标志(+)=1" & _
            " And A.病人ID=[1]"
        Set rsPati = gobjDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病人ID)
    End If
    
    intWarn = -1 '记帐报警时缺省要提示
    '执行报警:门诊病人病区ID=0
    strSQL = "Select Nvl(报警方法,1) as 报警方法,报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线 Where Nvl(病区ID,0)=[1] And 适用病人=[2]"
    Set rsWarn = gobjDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病区ID, CStr(Nvl(rsPati!适用病人)))
    If Not rsWarn.EOF Then
        If Val(Nvl(rsWarn!报警方法)) = 2 Then cur当日 = GetPatiDayMoney(lng病人ID)
        str类别名 = Mid(str类别名, 2)
        For i = 1 To Len(str类别)
            intR = BillingWarn(frmParent, strPrivs, rsWarn, Nvl(rsPati!姓名) & IIf(Nvl(rsPati!住院号) = "", "", "(住院号:" & Nvl(rsPati!住院号) & " 床号:" & Nvl(rsPati!床号) & ")"), Nvl(rsPati!剩余款, 0), cur当日, cur金额, Nvl(rsPati!担保额, 0), Mid(str类别, i, 1), Split(str类别名, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function





Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有部门 As String = "", _
    Optional blnSendKeys As Boolean = True, Optional blnAddItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参:cboDept-指定的部门部件
    '     rsDept-指定的部门
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有部门-所有部门名称
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim intIndex As Integer
    Dim strIDs As String, str简码 As String, strLike As String
    strLike = IIf(Val(gobjDatabase.GetPara("输入匹配")) = 0, "*", "")
    
    '先复制记录集
    Set rsTemp = gobjDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = strLike & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf gobjCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有部门 <> "" Then
        str简码 = gobjCommFun.SpellCode(str所有部门)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有部门) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编码) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0:  Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp): Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编码) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!名称)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编码)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!名称))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 And lngDeptID <= 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    End Select
    
    '弹出选择器
    If gobjDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "缺省," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    If gobjControl.CboLocate(cboDept, lngDeptID, True) = False Then
        If blnAddItem = True Then
            If rsTemp.RecordCount = 1 Then
                cboDept.RemoveItem cboDept.ListCount - 1
                cboDept.AddItem IIf(zlIsShowDeptCode, rsTemp!编码 & "-", "") & rsTemp!名称
                cboDept.ItemData(cboDept.ListCount - 1) = Val(Nvl(rsTemp!ID))
                intIndex = cboDept.NewIndex
                cboDept.AddItem "其他科室…"
                cboDept.ItemData(cboDept.ListCount - 1) = 0
                cboDept.ListIndex = intIndex
            Else
                cboDept.RemoveItem cboDept.ListCount - 1
                cboDept.AddItem IIf(zlIsShowDeptCode, rsReturn!编码 & "-", "") & rsReturn!名称
                cboDept.ItemData(cboDept.ListCount - 1) = Val(Nvl(rsReturn!ID))
                intIndex = cboDept.NewIndex
                cboDept.AddItem "其他科室…"
                cboDept.ItemData(cboDept.ListCount - 1) = 0
                cboDept.ListIndex = intIndex
            End If
            rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
            zlSelectDept = True
            Exit Function
        Else
            GoTo GoNotSel
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing

    If blnSendKeys Then gobjCommFun.PressKey vbKeyTab
    zlSelectDept = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.TxtSelAll cboDept
End Function


Public Function Get医嘱附项内容(ByVal lng医嘱ID As Long, ByVal str中文名 As String) As String
'功能:根据医嘱ID，元素名称、返回医嘱的对应元素的申请附项内容
'参数:str中文名  诊治所见项目.中文名

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
     
    strSQL = "Select a.内容 From 病人医嘱附件 A, 诊治所见项目 B" & _
        " Where a.要素id = b.Id And a.医嘱id = [1] And b.中文名 = [2]"

    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID, str中文名)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strTmp = IIf(strTmp = "", "", strTmp & ",") & rsTmp!内容
            rsTmp.MoveNext
        Next
    End If
    
    Get医嘱附项内容 = strTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function GetStock(ByVal lng药品ID As Long, Optional ByVal lng库房ID As Long, Optional ByVal int范围 As Integer = 2, _
        Optional ByVal strDepartments As String, Optional ByVal lng总量 As Double, Optional ByVal lng批次 As Long = -1) As Double
'功能：获取指定库房指定药品不分批库存(以门诊或住院单位)
'参数：int范围=1-门诊,2-住院(缺省),0-表示按售价
'      strDepartments可用执行科室字符串，用于批量查询库存
'      lng总量 如果lng总量不为空，则查询是否有库存大于这个总量
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    '获取药品库存(不分批或分批药品),药房不分批药品不管效期
    If int范围 = 0 Or int范围 = 3 Then
        If lng批次 = 0 Or lng批次 = -1 Then
            strSQL = _
                " Select Nvl(Sum(A.可用数量),0) as 库存" & _
                " From 药品库存 A" & _
                " Where A.性质=1" & _
                " And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
                " And A.药品ID=[1] And Instr([2],',' || a.库房id || ',')>0 " & _
                " Group By A.库房ID"
        Else
            strSQL = _
                " Select Nvl(Sum(A.可用数量),0) as 库存" & _
                " From 药品库存 A" & _
                " Where A.性质=1" & _
                " And Nvl(A.批次,0) = [3] And (A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
                " And A.药品ID=[1] And (Instr([2],',' || a.库房id || ',')>0 Or A.库房ID In (Select 虚拟库房id From 虚拟库房对照 Where Instr([2],',' || 科室id || ',')>0 And Rownum < 2)) " & _
                " Group By A.库房ID Order By Sign(Nvl(Sum(A.可用数量),0)) Desc "
        End If
    Else
        strTmp = IIf(int范围 = 1, "门诊", "住院")
        If lng批次 = 0 Or lng批次 = -1 Then
            strSQL = _
                " Select Nvl(Sum(A.可用数量),0)/Nvl(B." & strTmp & "包装,1) as 库存" & _
                " From 药品库存 A,药品规格 B" & _
                " Where A.药品ID=B.药品ID(+) And A.性质=1" & _
                " And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
                " And A.药品ID=[1] And Instr([2],',' || a.库房id || ',')>0" & _
                " Group by Nvl(B." & strTmp & "包装,1),A.库房ID"
        Else
            strSQL = _
                " Select Nvl(Sum(A.可用数量),0)/Nvl(B." & strTmp & "包装,1) as 库存" & _
                " From 药品库存 A,药品规格 B" & _
                " Where A.药品ID=B.药品ID(+) And A.性质=1" & _
                " And Nvl(A.批次,0) = [3] And (A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
                " And A.药品ID=[1] And (Instr([2],',' || a.库房id || ',')>0 Or A.库房ID In (Select 虚拟库房id From 虚拟库房对照 Where Instr([2],',' || 科室id || ',')>0 And Rownum < 2)) " & _
                " Group by Nvl(B." & strTmp & "包装,1),A.库房ID Order By Sign(Nvl(Sum(A.可用数量),0)) Desc"
        End If
    End If
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng药品ID, IIf(strDepartments = "", "," & lng库房ID & ",", "," & strDepartments & ","), lng批次)
    
    Do While Not rsTmp.EOF
    
        If strDepartments = "" Then
            GetStock = Format(rsTmp!库存, "0.00000")
            Exit Function
        Else
            If Val(rsTmp!库存) & "" > lng总量 Then
                GetStock = Format(rsTmp!库存, "0.00000")
                Exit Function
            End If
        End If
        rsTmp.MoveNext
    
    Loop
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function



Public Function Get部门名称(lngID As Long, Optional ByRef rs部门 As ADODB.Recordset) As String
'功能：获取部门名称
'参数：lngID=部门ID
'返回：部门名称
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If rs部门 Is Nothing Then
        strSQL = "Select 名称 from 部门表 Where ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngID)
    Else
        Set rsTmp = rs部门
        rsTmp.Filter = "ID=" & lngID
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select 名称 from 部门表 Where ID=[1]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngID)
        End If
    End If
    If Not rsTmp.EOF Then Get部门名称 = rsTmp!名称
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get项目名称(lng项目id As Long) As String
'功能：返回诊疗项目名称
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 名称 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id)
    If Not rsTmp.EOF Then Get项目名称 = rsTmp!名称
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckFeeItemAvailable(ByVal lngFeeItemID As Long, ByVal bytFlag As Byte) As Boolean
'功能:检查收费项目是否未停用,并且服务于病人
'参数:bytFlag:服务对象:1-门诊,2-住院
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From 收费项目目录 Where ID = [1] And (撤档时间 is Null Or 撤档时间 > Sysdate) And 服务对象 In (" & bytFlag & ",3)"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItemID)
    CheckFeeItemAvailable = rsTmp.RecordCount > 0
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function Get收据费目(ByVal lng收入项目ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收据费目
    '返回:返回收据费目
    '编制:刘兴洪
    '日期:2014-04-11 16:33:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    'On Error GoTo errHandle
    '不处理异常，该函数主要是在保存数据时使用，在主窗体处理异常
    If grs收入项目 Is Nothing Then
        strSQL = "Select ID,编码,名称,收据费目 From 收入项目　Where  (撤档时间 is Null Or 撤档时间 > Sysdate)"
        Set grs收入项目 = gobjDatabase.OpenSQLRecord(strSQL, "获取收入项目")
    ElseIf grs收入项目.State <> 1 Then
        strSQL = "Select ID,编码,名称,收据费目 From 收入项目 　Where  (撤档时间 is Null Or 撤档时间 > Sysdate)"
        Set grs收入项目 = gobjDatabase.OpenSQLRecord(strSQL, "获取收入项目")
    End If
    grs收入项目.Filter = "ID=" & lng收入项目ID
    If grs收入项目.EOF = False Then Get收据费目 = grs收入项目!收据费目
'Get收据费目 = True
'    Exit Function
'errHandle:
'    If gobjComlib.ErrCenter() = 1 Then
'        Resume
'    End If
End Function
Public Function GetPatiInforFromAdvice(ByVal lng医嘱ID As Long, _
    ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByRef lng病人科室ID As Long, _
    ByRef lng病人病区ID As Long, _
    ByRef lng医疗小组ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱获取当时开医嘱的病人相关状态(病人科室ID,当前病区ID,医疗小组ID)
    '入参:lng医嘱ID-医嘱ID
    '     lng病人ID-病人ID
    '     lng主页ID-主页ID
    '出参:lng病人科室ID-返回当时医嘱的病人科室id
    '     lng病人病区ID-返回当时医嘱的病人病区ID
    '     lng医疗小组ID-返回当时医嘱的医疗小组ID
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-21 15:05:19
    '问题:70896
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    '只有住院才会存在
    lng病人科室ID = 0: lng病人病区ID = 0: lng医疗小组ID = 0
    If Not (lng病人ID <> 0 And lng主页ID <> 0) Then GetPatiInforFromAdvice = True: Exit Function
    
    strSQL = " " & _
    " Select * " & _
    " From (Select a.病区id, Nvl(b.病人科室id, a.科室id) As 病人科室id, a.医疗小组id, a.开始时间 " & _
    "        From 病人变动记录 A, (Select 开嘱时间, 病人科室id From 病人医嘱记录 Where ID = [3]) B " & _
    "        Where a.病人id = [1] And 主页id = [2] And b.开嘱时间 Between 开始时间 And Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
    "              Nvl(a.科室id, 0) <> 0 " & _
    "        Order By 开始时间 Desc) " & _
    " Where Rownum < 2"
 
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "根据医嘱获取病区及病人科室Id", lng病人ID, lng主页ID, lng医嘱ID)
    If rsTemp.EOF Then GetPatiInforFromAdvice = True: Exit Function
    lng病人科室ID = Nvl(rsTemp!病人科室id, 0)
    lng病人病区ID = Nvl(rsTemp!病区ID, 0)
    lng医疗小组ID = Nvl(rsTemp!医疗小组id, 0)
    GetPatiInforFromAdvice = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function BillOperCheck(bytNO As Byte, strOperator As String, Datadd As Date, Optional strMessage As String = "操作", _
    Optional ByVal strNO As String, Optional ByVal lngPatientID As Long, _
    Optional ByVal bytFlag As Byte = 2, Optional ByVal blnOnlyCheckLimit As Boolean, Optional ByVal blnCheckOperator As Boolean = True, _
    Optional ByVal blnCheckCur As Boolean = True, Optional blnNotMsgBox As Boolean, Optional strOutErrMsg As String) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断当前人员对单据是否有操作权限
    '入参: bytNO：1-挂号单据,2-收费单,3-划价单,4-门诊记帐,5-住院记帐,6-预交款,7-结帐单据,8-就诊卡
    '   strOperator：单据实际的操作员
    '   DatAdd：单据的登记时间
    '   strNO   ：检查金额上限时用来确定单据
    '   lngPatientID：检查金额上限时，对于记帐表用来确定单据中的病人
    '   bytFlag：1-收费单,2-记帐单,3-记帐单
    '   blnOnlyCheckLimit：只检查金额上限
    '   blnCheckOperator：要检查是否允许操作他人单据
    '   blnCheckCur：是否检查金额上限
    '   blnNotMsgBox：True-不显示消息提示框;False-显示消息提示框
    '出参:strOutErrMsg-返回错误信息
    '返回:是否有操作权限,有返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-10-17 17:13:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
 

    Dim strSQL As String, strBill As String
    Dim rsTmp As ADODB.Recordset
    Dim curTmp As Currency
    Dim int来源 As Integer
    
    If bytNO = 1 Or bytNO = 2 Or (bytNO = 3 And bytFlag = 1) Or bytNO = 4 Then
        int来源 = 1
    Else
        int来源 = 2
    End If
    
    If glngSys Like "8??" Then
        strBill = Switch(bytNO = 1, "挂号单据", bytNO = 2, "收费单据", bytNO = 3, _
            "划价单据", bytNO = 4, "记帐单据", bytNO = 5, "记帐单据", _
            bytNO = 6, "预交款单据", bytNO = 7, "结帐单据", bytNO = 8, "会员卡")
    Else
        strBill = Switch(bytNO = 1, "挂号单据", bytNO = 2, "收费单据", bytNO = 3, _
            "划价单据", bytNO = 4, "记帐单据", bytNO = 5, "记帐单据", _
            bytNO = 6, "预交款单据", bytNO = 7, "结帐单据", bytNO = 8, "就诊卡")
    End If
        
    On Error GoTo errH
    
    strSQL = "" & _
    "   Select Nvl(时间限制,0) as 时间限制,Nvl(他人单据,0) as 他人单据,Nvl(金额上限,0) as 金额上限 " & _
    "   From 单据操作控制 Where 人员ID=[1] And 单据=[2]"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, UserInfo.ID, bytNO)
    If rsTmp.EOF Then
        BillOperCheck = True
        Exit Function
    End If

    If Not blnOnlyCheckLimit Then
        If rsTmp!他人单据 = 0 And blnCheckOperator Then
            If strOperator <> UserInfo.姓名 Then
                strOutErrMsg = "你没有权限对" & strOperator & "处理的" & strBill & "进行" & strMessage & "！"
                If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If rsTmp!时间限制 > 0 Then
            If Int(gobjDatabase.Currentdate) - Int(CDate(Datadd)) + 1 > rsTmp!时间限制 Then
                strOutErrMsg = "你只能对 " & rsTmp!时间限制 & " 天内处理的" & strBill & "进行" & strMessage & "！"
                If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    If rsTmp!金额上限 > 0 And blnCheckCur Then
        If strNO <> "" Then
            curTmp = GetBillMoney(strNO, 2, IIf(int来源 = 1, True, False), lngPatientID)
            If curTmp >= rsTmp!金额上限 Then
                strOutErrMsg = "你只能对 " & rsTmp!金额上限 & " 元以下的" & strBill & "进行" & strMessage & "！" & _
                vbCrLf & "单据[" & strNO & "]的实收金额合计为:" & curTmp
                If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    BillOperCheck = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Check留观病人(ByVal strNO As String, ByVal strPrivs As String, Optional ByVal strTime As String, Optional ByVal bytFlag As Byte = 2, _
    Optional blnNotMsgBox As Boolean, Optional strOutErrMsg As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据是否允许对留观病人人进行记帐,对记帐单/表进行检查
    '     主要用于记帐单/表修改,销帐。对于记帐表,只要存在一个留观病人无权限,则整单禁止
    '入参:
    '   blnNotMsgBox：True-不显示消息提示框;False-显示消息提示框
    '出参:strOutErrMsg-返回错误信息
    '返回:没有权限的留观病人,如"留观病人","门诊留观病人","住院留观病人"
    '编制:刘兴洪
    '日期:2016-10-17 17:44:16
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset
    Dim bln门诊留观 As Boolean
    Dim bln住院留观 As Boolean
    Dim strSQL As String
    
    bln门诊留观 = gSysPara.bln门诊留观记帐 And InStr(strPrivs, ";门诊留观记帐;") > 0
    bln住院留观 = gSysPara.bln住院留观记帐 And InStr(strPrivs, ";住院留观记帐;") > 0
        
    If bln门诊留观 And bln住院留观 Then Exit Function
    
    If Not bln门诊留观 And Not bln住院留观 Then
        strSQL = "1,2"
    ElseIf Not bln门诊留观 Then
        strSQL = "1"
    ElseIf Not bln住院留观 Then
        strSQL = "2"
    End If
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(B.病人性质,0) as 病人性质" & _
        " From 住院费用记录 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
        " And A.NO=[1] And A.记录性质=[2]" & _
        " And Nvl(B.病人性质,0) IN(" & strSQL & ") And A.记录状态 IN(0,1,3)" & _
        IIf(strTime <> "", " And A.登记时间=[3]", "")
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    If Not rsTmp.EOF Then
        If rsTmp.RecordCount = 2 Then
            Check留观病人 = "留观病人"
        ElseIf rsTmp!病人性质 = 1 Then
            Check留观病人 = "门诊留观病人"
        ElseIf rsTmp!病人性质 = 2 Then
            Check留观病人 = "住院留观病人"
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function CheckDelPriv(ByVal strNO As String, ByVal strPrivs As String, Optional ByVal strTime As String, _
        Optional ByVal bytFlag As Byte = 2, Optional ByVal bytMode As Byte = 1, Optional ByVal blnNotMsgBox As Boolean, Optional ByRef strOutErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否权限冲销住院记帐单
    '入参:bytMode,部分权限不足时是否仅提示,1-允许继续,返回真,0-不允继续,返回假
    '   blnNotMsgBox：True-不显示消息提示框;False-显示消息提示框
    '出参:strOutErrMsg=返回错误信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-10-17 17:22:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '只判断未销帐费用行
    strSQL = "Select Nvl(Sum(Decode(收费类别,'5',1,'6',1,'7',1,0)),0) as 药品数," & _
        " Nvl(Sum(Decode(收费类别,'4',1,0)),0) as 卫材数," & _
        " Nvl(Sum(Decode(收费类别,'4',0,'5',0,'6',0,'7',0,1)),0) as 诊疗数" & _
        " From 住院费用记录" & _
        " Where 记录性质=[2] And 记录状态 IN(0,1) And NO=[1]" & _
        IIf(strTime <> "", " And 登记时间=[3]", "")
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytFlag)
    End If
    
    If rsTmp.EOF Then CheckDelPriv = True: Exit Function
    '没有住院销帐权限时,菜单和按钮已设置为不可见
    Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
    Dim strNotPrivs As String, strNote As String
    
    blnYP = InStr(strPrivs, ";药品销帐;") > 0
    blnZL = InStr(strPrivs, ";诊疗销帐;") > 0
    blnWC = InStr(strPrivs, ";卫材销帐;") > 0
    
    If blnYP = False And blnZL = False And blnWC = False Then
        strOutErrMsg = "你没有药品销帐或卫材销帐或诊疗销帐的权限,不能对单据[" & strNO & "]进行销帐！"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    strNotPrivs = ""
    If Not blnYP Then strNotPrivs = strNotPrivs & "和药品销帐"
    If Not blnWC Then strNotPrivs = strNotPrivs & "和卫材销帐"
    If Not blnZL Then strNotPrivs = strNotPrivs & "和诊疗销帐"
    strNotPrivs = Mid(strNotPrivs, 2)
    strNote = ""
    
    If blnYP Then strNote = strNote & "或药品销帐"
    If blnWC Then strNote = strNote & "或卫材销帐"
    If blnZL Then strNote = strNote & "或诊疗销帐"
    strNote = Mid(strNote, 2)
 
    If rsTmp!药品数 > 0 And Not blnYP Then
        strOutErrMsg = "你没有" & strNotPrivs & "权限,只能对单据[" & strNO & "]中的" & strNote & "进行销帐！"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    If rsTmp!卫材数 > 0 And Not blnWC Then
        strOutErrMsg = "你没有" & strNotPrivs & "权限,只能对单据[" & strNO & "]中的" & strNote & "进行销帐！"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    If rsTmp!诊疗数 > 0 And Not blnZL Then
        strOutErrMsg = "你没有" & strNotPrivs & "权限,只能对单据[" & strNO & "]中的" & strNote & "进行销帐！"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    CheckDelPriv = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function




Public Function BillCanBeOperate(ByVal strNO As String, ByVal strPriv As String, _
    ByVal strNote As String, Optional ByVal strTime As String, _
    Optional str病人IDs As String, Optional ByVal bytType As Byte = 2, Optional ByVal blnNotMsgBox As Boolean, _
    Optional ByRef strOutErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据的病人信息判断是否有权限操作该单据
    '入参:strNote=描述操作类型,用于提示。销帐时有特殊处理。
    '     str病人IDs=允许时，返回允许操作的病人ID串,空为所有病人
    '   blnNotMsgBox：True-不显示消息提示框;False-显示消息提示框
    '出参:strOutErrMsg=返回错误信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-10-17 18:07:00
    '说明:主要是病人出院(或预出院)后,如果没有权限,则不允许操作
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnOut As Boolean
    Dim strInfo As String
    
    str病人IDs = ""
    If InStr(strPriv, ";出院未结强制记帐;") > 0 And InStr(strPriv, ";出院结清强制记帐;") > 0 Then
        BillCanBeOperate = True: Exit Function
    End If
    
    On Error GoTo errH
    
    '如果无对应主页,则当作已出院病人(如门诊病人医技记帐)
    If strNote Like "*销帐" Then
        '销帐操作时,只对可以销帐部份内容进行判断
        strSQL = _
            " Select 序号 From 住院费用记录" & _
            " Where 记录性质=[2] And NO=[1] And Nvl(执行状态,0)<>1 And 价格父号 is NULL" & _
            " Group by 序号 Having Nvl(Sum(Nvl(付数,1)*数次),0)<>0"
    ElseIf strNote Like "*审核" Then
        '审核操作时,只对未审核部份内容进行判断
        strSQL = _
            " Select 序号 From 住院费用记录" & _
            " Where 记录性质=2 And 价格父号 is NULL And 记录状态=0 And NO=[1]"
    End If
    strSQL = "Select Distinct 姓名,病人ID,主页ID From 住院费用记录" & _
        " Where 记录性质=[2] And NO=[1] And 记录状态 IN(0,1,3)" & _
        IIf(strTime <> "", " And 登记时间=[3]", "") & _
        IIf(strSQL <> "", " And Nvl(价格父号,序号) IN(" & strSQL & ")", "")

    strSQL = "Select B.病人ID,B.姓名," & _
    " Decode(A.病人ID,NULL,Sysdate,A.出院日期) as 出院日期," & _
    " Nvl(A.状态,0) as 状态,Nvl(C.费用余额,0) as 余额" & _
    " From 病案主页 A,(" & strSQL & ") B,病人余额 C" & _
    " Where B.病人ID=A.病人ID(+) And C.性质(+)=1 And C.类型(+)=2  And B.主页ID=A.主页ID(+) And B.病人ID=C.病人ID(+) And C.性质(+)=1 And C.类型(+)=2 "
        
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytType, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytType)
    End If
    
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!出院日期) Or rsTmp!状态 = 3 Then
            If rsTmp!余额 = 0 And InStr(strPriv, ";出院结清强制记帐;") = 0 Then
                strInfo = strInfo & vbCrLf & "病人""" & rsTmp!姓名 & """已出院(或预出院)且费用已经结清。"
            ElseIf rsTmp!余额 <> 0 And InStr(strPriv, ";出院未结强制记帐;") = 0 Then
                strInfo = strInfo & vbCrLf & "病人""" & rsTmp!姓名 & """已出院(或预出院)且费用尚未结清。"
            Else
                str病人IDs = str病人IDs & "," & rsTmp!病人ID
            End If
        Else
            str病人IDs = str病人IDs & "," & rsTmp!病人ID
        End If
        rsTmp.MoveNext
    Loop
    str病人IDs = Mid(str病人IDs, 2)
        
    '只有记帐表销帐可以部份继续
    If strInfo <> "" Then
        strOutErrMsg = Mid(strInfo, 3) & vbCrLf & "你没有权限对单据""" & strNO & """进行" & strNote & "。"
        If Not blnNotMsgBox Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    BillCanBeOperate = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function



Public Function GetBillingBalanceStatu(ByVal int来源 As Integer, ByVal strNO As String, Optional ByVal blnAll As Boolean = True, _
    Optional ByVal strTime As String, Optional ByVal bytFlag As Byte = 2, Optional ByRef intOutStatu As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一张记帐单/表的结帐状态
    '入参：int来源-1-门诊;2-住院
    '      strNO=记帐单据号,不分门诊及住院
    '      blnALL=是否对整张单据内容进行判断,否则只对未销帐部分进行判断
    '出参:intOutStatu-返回结帐状态：0-未结帐,1=已全部结帐,2-已部分结帐
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-10-18 10:50:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    
    On Error GoTo errH
    intOutStatu = 0
    
    '求未作废的费用行
    strSQL = _
        " Select 序号 From (" & _
            " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号, Avg(Nvl(付数, 1) * 数次) As 数量" & _
            " From " & IIf(int来源 = 1, "门诊费用记录", "住院费用记录") & _
            " Where NO=[1] And 记录性质=[2]" & _
            " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    
    '求每行的结帐情况
    strSQL = _
        "Select Nvl(价格父号,序号) as 序号,Sum(Nvl(结帐金额,0)) as 结帐金额" & _
        " From " & IIf(int来源 = 1, "门诊费用记录", "住院费用记录") & _
        " Where NO=[1] And mod(记录性质,10)= [2]" & _
        IIf(Not blnAll, " And Nvl(价格父号,序号) IN(" & strSQL & ")", "") & _
        IIf(strTime <> "", " And 登记时间=[3]", "") & _
        " Group by Nvl(价格父号,序号)"
    
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlDockExpense", strNO, bytFlag)
    End If
    
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '单据行数
        rsTmp.Filter = "结帐金额<>0"
        If rsTmp.EOF Then
            intOutStatu = 0 '无结帐行
        ElseIf rsTmp.RecordCount = lngTmp Then
            intOutStatu = 1 '全部行已结帐
        ElseIf rsTmp.RecordCount > 0 Then
            intOutStatu = 2 '部分行已结帐
        End If
    End If
    GetBillingBalanceStatu = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
 
Public Function zlCheckIsExistsApplied(ByVal strNO As String, ByVal str序号 As String, _
    ByRef str费用IDs As String, Optional ByRef str申请人s As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查销帐的数据中是否存在销帐申请
    '入参:strNo-单据号
    '       str序号-本次销帐的序号(为空为所有)
    '出参:str费用IDs-申请的费用ID
    '返回:存在销帐申请,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-20 15:51:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct A.ID,B.申请人 " & _
    "   From 住院费用记录 A,病人费用销帐 B  " & _
    "   Where A.ID=B.费用ID and A.NO=[1] and A.记录性质=2 And nvl(B.状态,0)=0 " & IIf(str序号 <> "", " and Instr([2],','||序号||',')>0 ", "")
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取申请状态", strNO, "," & str序号 & ",")
    If rsTemp.EOF Then
        rsTemp.Close: Set rsTemp = Nothing: Exit Function
    End If
    str申请人s = "": str费用IDs = ""
    With rsTemp
        Do While Not .EOF
            str费用IDs = str费用IDs & "," & Val(Nvl(rsTemp!ID))
            If InStr(1, str申请人s & vbCrLf, vbCrLf & Nvl(rsTemp!申请人) & vbCrLf) = 0 Then
                str申请人s = str申请人s & vbCrLf & Nvl(rsTemp!申请人)
            End If
            .MoveNext
        Loop
    End With
    If str费用IDs <> "" Then str费用IDs = Mid(str费用IDs, 2)
    zlCheckIsExistsApplied = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function GetBillInsures(strInsure As String, ByVal strNO As String, _
    Optional ByVal strTime As String, Optional ByVal blnAuditing As Boolean, _
    Optional ByVal blnGetNoneInsure As Boolean, Optional ByVal bytFlag As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取记帐表中的险类串"10,20,30,...",也适用于记帐单
    '入参:：strNO=记帐单据号
    '      blnAuditing=是否用于记帐审核,只检查未审核的部份内容
    '      blnGetNoneInsure=是否将非保险费用返回为0险类
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-10-18 13:44:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strInsure = ""
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(B.险类,0) as 险类" & _
        " From 住院费用记录 A,病案主页 B" & _
        " Where A.记录性质=[2] And A.记录状态" & IIf(blnAuditing, "=0", " IN(0,1,3)") & _
            IIf(blnGetNoneInsure, "", " And B.险类 is Not NULL") & _
        " And A.NO=[1] And A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
        IIf(strTime <> "", " And A.登记时间=[3]", "")
    If strTime <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "获取记帐单的相关保险信息", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "获取记帐单的相关保险信息", strNO, bytFlag)
    End If
    
    Do While Not rsTmp.EOF
        strInsure = strInsure & "," & rsTmp!险类
        rsTmp.MoveNext
    Loop
    strInsure = Mid(strInsure, 2)
    GetBillInsures = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetPriceGradeStartType() As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取价格等级的启用类型
    '返回:
    '   0-未启用
    '   1-只启用了站点
    '   2-只启用了医疗付款方式
    '   3-站点和医疗款方式都启用了
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandler
    GetPriceGradeStartType = 0
    strSQL = _
        " Select Nvl(Max(Decode(b.站点, Null, 0, 1)), 0) As 启用站点," & vbNewLine & _
        "        Nvl(Max(Decode(b.医疗付款方式, Null, 0, 1)), 0) As 启用医疗付款方式" & vbNewLine & _
        " From 收费价格等级 A, 收费价格等级应用 B" & vbNewLine & _
        " Where a.名称 = b.价格等级 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "获取有效的价格等级类型")
    If rsTmp.EOF Then Exit Function
    
    If Val(Nvl(rsTmp!启用站点)) = 1 Then
        If Val(Nvl(rsTmp!启用医疗付款方式)) = 1 Then
            GetPriceGradeStartType = 3 '站点和医疗款方式都启用了
        Else
            GetPriceGradeStartType = 1 '只启用了站点
        End If
    Else
        If Val(Nvl(rsTmp!启用医疗付款方式)) = 1 Then
            GetPriceGradeStartType = 2 '只启用了医疗付款方式
        End If
    End If
    Exit Function
errHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetPriceGrade(ByVal str站点 As String, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    Optional ByVal str医疗付款方式 As String, _
    Optional ByRef str药品价格等级_Out As String, _
    Optional ByRef str卫材价格等级_Out As String, _
    Optional ByRef str普通项目价格等级_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 根据医疗付款方式或站点，获取对应的价格等级
    '入参:str站点-登陆的站点，必须传入，传入NULL时，价格等级为返回空
    '     lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     str医疗付款方式:如果传入非空，则以传的医疗付款方式_In方式来提取价格等级;否则以病人ID_In或主页ID来获取对应的病人的医疗付款方式。
    
    '出参:str药品价格等级_out-返回药品价格等级
    '     str卫材价格等级_out-返回卫材价格等级
    '     str普通项目价格等级_out-返回普通收费项目价格等级
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim varPara As Variant
    
    str药品价格等级_Out = "": str卫材价格等级_Out = "": str普通项目价格等级_out = ""
    On Error GoTo errHandle
    '    Zl_Get_Pricegrade
    '  站点_In         In 收费价格等级对照.站点%Type,
    '  病人id_In       In 病人信息.病人id%Type := Null,
    '  主页id_In       In 病案主页.主页id%Type := Null,
    '  医疗付款方式_In In 收费价格等级对照.病人类型%Type := Null
    
    strSQL = "Select Zl_Get_Pricegrade([1],[2],[3],[4]) as 价格等级 From dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取价格等级", str站点, lng病人ID, lng主页ID, str医疗付款方式)
    If Nvl(rsTemp!价格等级) = "" Then GetPriceGrade = True: Exit Function
    '格式:普通价格等级|药品价格等级|卫生材料价格等级
    varPara = Split(rsTemp!价格等级 & "||||", "|")
    
    str普通项目价格等级_out = varPara(0)
    str药品价格等级_Out = varPara(1)
    str卫材价格等级_Out = varPara(2)
    GetPriceGrade = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetRetailPrice(ByVal lng收费细目ID As Long, _
    ByVal str价格等级 As String, ByRef dbl零售价_out As Double, ByRef dbl未分解数_out As Double, _
    Optional ByVal lng库房ID As Long = 0, _
    Optional ByVal dbl数量 As Double = 0) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 根据价格等级获取指定收费项目的零售价的相关信息
    '入参:lng收费细目id-收费细目ID
    '     str价格等级-收费价格等级
    '     lng库房id-库房ID（药品和卫生材料传入)
    '     dbl数量:当前出库数量(药品和卫生材料传入)。
    '出参:dbl零售价_out-返回零售价格
    '     dbl未分解数_out-针对药品和卫生材料有效，表示根据当前输入的出库数量（参数：dbl数量)进行分解时，未分解完的数量.
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim varPara As Variant
    
    On Error GoTo errHandle
    'Zl_Get_Retailprice
    '  收费细目id_In In 收费项目目录.Id%Type,
    '  价格等级_In   In 收费价格等级.名称%Type,
    '  库房id_In     In 部门表.Id%Type := 0,
    '  数量_In       In Number := 0
    ') Return Varchar2
    '  --      a.药品和卫生材料:零售价|未分解数
    '  --      b.针对普通材料:零售价(实价为缺省价格)|0
  
    strSQL = "Select Zl_Get_Retailprice([1],[2],[3],[4]) as 价格 From dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "根据价格等级获取价格", lng收费细目ID, str价格等级, lng库房ID, dbl数量)
    If Nvl(rsTemp!价格) = "" Then Exit Function
    
    varPara = Split(rsTemp!价格 & "||||", "|")
    dbl零售价_out = FormatEx(varPara(0), 5)
    dbl未分解数_out = FormatEx(varPara(1), 5)
    GetRetailPrice = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
