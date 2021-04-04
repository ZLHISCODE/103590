Attribute VB_Name = "mdlExpense"
Option Explicit

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
        
    strSQL = " Select Distinct 收费细目ID,Nvl(开单科室ID,0) as 开单科室ID,执行科室id From 收费执行科室 Where 收费细目ID In (" & str收费细目IDs & ") "
    
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    If Not rsTmp.EOF Then Set GetServiceDept = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDrugTotal(ByVal objBill As ExpenseBill, ByVal lng药品ID As Long, ByVal lng药房ID As Long) As Double
'功能：获取单据中指定药品在同一药房多行的数量合
    Dim i As Integer, dblCount As Double
    
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).收费细目ID = lng药品ID And objBill.Details(i).执行部门ID = lng药房ID Then
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID)
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get处方职务(lng药品ID As Long) As String
'功能：根据药品ID获取其处方职务
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    Get处方职务 = "00"
    strSQL = "Select Nvl(B.处方职务,'00') as 处方职务 From 药品规格 A,药品特性 B Where A.药名ID=B.药名ID And A.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng药品ID)
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
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(A.处方限量,0) as 处方限量" & _
        " From 药品特性 A,药品规格 B Where A.药名ID=B.药名ID And B.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngID)
    If Not rsTmp.EOF Then Get处方限量 = rsTmp!处方限量
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemExistInsure(ByVal lng收费细目ID As Long, ByVal int险类 As Integer) As Boolean
'功能：判断收费项目是否设置了保险支付项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    If gclsInsure.GetCapability(support允许不设置医保项目, , int险类) Then
        ItemExistInsure = True: Exit Function
    End If
    
    strSQL = "Select * From 保险支付项目 Where 收费细目ID=[1] And 险类=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng收费细目ID, int险类)
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
        
    strSQL = "Select A.药品ID,A.剂量系数,B.计算单位 as 剂量单位" & _
        " From 药品规格 A,诊疗项目目录 B" & _
        " Where A.药名ID=B.ID And A.药品ID IN (" & strItemIDs & ")"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlExpense") 'In
    
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
        " Nvl(Sum(A.可用数量),0)" & IIF(bln药房单位, "/Nvl(C." & str药房包装 & ",1)", "") & " as 库存" & _
        " From 药品库存 A,(" & strSQL & ") B,药品规格 C" & _
        " Where A.库房ID=B.ID And A.药品ID=C.药品ID" & _
        " And ((A.效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Or (Nvl(C.药房分批,0)=0 And A.库房ID IN(" & strSQL2 & ")))" & _
        " And A.性质=1 And A.药品ID=[2]" & _
        " Group by B.编码,B.名称,A.库房ID,Nvl(C." & str药房包装 & ",1)" & _
        " Having Sum(Nvl(A.可用数量,0))<>0" & _
        " Order By B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", str性质, lng药品ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!名称 & ":" & rsTmp!库存
        rsTmp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    GetStockInfo = strSQL
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
    
    str上午 = GetSysParVal(1): str下午 = GetSysParVal(2)
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

Public Function GetBillRows(str单据号 As String, int记录性质 As Integer) As Integer
'功能：获取一张费用单据中未作废的费用行数
'参数：int记录性质=1-收费(划价),2-记帐
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    strSQL = _
        " Select 序号,Sum(数量) as 剩余数量" & _
        " From (" & _
        " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号," & _
        " Avg(Nvl(付数, 1) * 数次) As 数量" & _
        " From 病人费用记录" & _
        " Where NO=[1] And 记录性质=[2]" & _
        " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", str单据号, int记录性质)
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
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.险类 From 病人费用记录 A,病案主页 B" & _
        " Where A.记录性质=2 And A.记录状态 IN(0,1,3) And B.险类 is Not NULL" & _
        " And A.NO=[1] And A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then BillExistInsure = rsTmp!险类
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsureName(intInsure As Integer) As String
'功能：根据保险类别序号获取保险类别名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select * From 保险类别 Where 序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", intInsure)
    If Not rsTmp.EOF Then GetInsureName = Nvl(rsTmp!名称)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        " From 部门表 A,部门性质说明 B," & IIF(bytType = 0, "药品出库检查", "材料出库检查") & " C" & _
        " Where B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 " & IIF(bytType = 0, "IN('中药房','西药房','成药房')", "='发料部门'") & _
        " And C.库房ID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetStockCheck")
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
        " Select A.组编号,Count(Distinct A.项目ID) as 禁忌数" & _
        " From 诊疗互斥项目 A,药品规格 B" & _
        " Where A.项目ID=B.药名ID And B.药品ID IN(" & strIDs & ")" & _
        " Having Count(Distinct A.项目ID)>1 Group by A.组编号"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlExpense") 'In
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strGroup = strGroup & "," & rsTmp!组编号
            rsTmp.MoveNext
        Next
        strGroup = Mid(strGroup, 2)
        
        For i = 0 To UBound(Split(strGroup, ","))
            strSQL = _
                "Select Distinct C.类型,C.组编号,D.编码,D.名称,D.规格" & _
                " From 药品规格 A,诊疗项目目录 B,诊疗互斥项目 C,收费项目目录 D" & _
                " Where A.药名ID=B.ID And B.ID=C.项目ID And A.药品ID=D.ID" & _
                " And C.组编号=" & Split(strGroup, ",")(i) & _
                " And A.药品ID IN(" & strIDs & ")" & _
                " Order by C.类型,C.组编号,D.编码"
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlExpense") 'In
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
    
    Dim dblAllTime As Double, dblCurTime As Double
    Dim dblPrice As Double, str药房 As String
    
    Dim colSerial As New Collection '用于处理从属父号
    Dim strSQL As String
    
    Dim lng西药房 As Long, lng成药房 As Long, lng中药房 As Long
    Dim bln药房单位 As Boolean, str药房单位 As String, str药房包装 As String
    
    '缺省药房
    lng西药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int来源 = 2, "住院", "门诊") & "缺省西药房", 0))
    lng成药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int来源 = 2, "住院", "门诊") & "缺省成药房", 0))
    lng中药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int来源 = 2, "住院", "门诊") & "缺省中药房", 0))
    
    '药品单位
    bln药房单位 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "药品单位", 0)) <> 0
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
    strSQL = _
        " Select X.药品ID,W.材料ID,W.跟踪在用," & _
        " A.序号 As 序号,A.从属父号,A.NO,A.记录性质,A.记录状态,A.多病人单,A.婴儿费,A.费别,A.姓名,A.性别,A.年龄," & _
        " A.床号,A.标识号,A.病人ID,A.主页ID,A.病人病区ID,A.病人科室ID,A.开单部门ID,A.门诊标志,A.加班标志," & _
        " A.附加标志,A.收费类别,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
        " A.标准单价 As 标准单价," & str药房 & " as 执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
        " B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(F.名称,B.名称) as 名称,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
        " B.屏蔽费别,B.说明,B.执行科室,Nvl(A.费用类型,B.费用类型) 费用类型,D.现价,D.原价,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
        " E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
        " Decode(A.收费类别,'4',1,X." & str药房包装 & ") as 药房包装," & _
        " Decode(A.收费类别,'4',B.计算单位,X." & str药房单位 & ") as 药房单位," & _
        " Decode(A.收费类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,Nvl(Y.库存,0) As 库存,B.录入限量" & _
        " From 病人费用记录 A,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,收费项目别名 F,材料特性 W,药品规格 X," & _
        "   (Select A.药品ID,A.库房ID,Sum(Nvl(A.可用数量,0)) as 库存 From 药品库存 A" & _
        "       Where A.性质=1 And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
        "       And A.药品ID IN(Select 收费细目ID From 病人费用记录 Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1])" & _
        "    Group by A.药品ID,A.库房ID) Y" & _
        " Where A.记录性质=[2] And A.记录状态 IN(0,1,3) And A.NO=[1]" & _
        " And A.价格父号 Is Null And A.收费细目ID=B.ID And A.收费细目ID=D.收费细目ID" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is NULL)" & _
        " And A.收费类别=C.编码 And A.收费细目ID=X.药品ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
        " And A.收费细目ID=Y.药品ID(+) And " & str药房 & "=Y.库房ID(+)" & _
        " And A.收费细目ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[3]" & _
        " And ((Sysdate Between D.执行日期 And D.终止日期) Or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"

    strSQL = "Select * From (" & strSQL & ") Order by 序号"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", str单据号, int记录性质, IIF(gbln商品名, 3, 1))
    
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
                    objBill.科室ID = Nvl(!病人科室ID, 0)
                    objBill.姓名 = Nvl(!姓名)
                    objBill.性别 = Nvl(!性别)
                    objBill.年龄 = Nvl(!年龄)
                    objBill.费别 = Nvl(!费别)
                    objBill.标识号 = Nvl(!标识号, 0)
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
                                strSQL = _
                                    " Select Nvl(A.批次,0) as 批次,Nvl(A.可用数量,0) as 库存," & _
                                    "   Nvl(Decode(Nvl(A.实际数量,0),0,0,A.实际金额/A.实际数量),0) as 时价" & _
                                    " From 药品库存 A" & _
                                    " Where A.库房ID=[1] And A.药品ID=[2] And Nvl(A.可用数量,0)>0" & _
                                    "   And A.性质=1 And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
                                    " Order by Nvl(A.批次,0)"
                                Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", objBillDetail.执行部门ID, Val(!收费细目ID))
                                '时价=总金额/总数量
                                dblPrice = 0
                                For i = 1 To rsPrice.RecordCount
                                    If dblAllTime = 0 Then Exit For
                                    '取小者
                                    If dblAllTime <= rsPrice!库存 Then
                                        dblCurTime = dblAllTime
                                    Else
                                        dblCurTime = rsPrice!库存
                                    End If
                                    dblPrice = dblPrice + Format(dblCurTime * Format(rsPrice!时价, "0.00000"), gstrDec)
                                    dblAllTime = Val(dblAllTime) - Val(dblCurTime)
                                    rsPrice.MoveNext
                                Next
                                If dblAllTime <> 0 Then
                                    '数量未分解完毕
                                    If !收费类别 = "4" Then
                                        MsgBox "时价卫生材料""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "时价药品""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.标准单价 = 0
                                Else
                                    objBillIncome.标准单价 = Format(dblPrice / (!付数 * !数次), "0.00000") '这里是售价价格
                                End If
                            Else
                                objBillIncome.标准单价 = 0
                            End If
                            '----------------------------------------------------------------------------------------------
                        Else
                            If Abs(!标准单价) > Abs(Nvl(!现价, 0)) Then
                                objBillIncome.标准单价 = Nvl(!原价, 0)
                            Else
                                objBillIncome.标准单价 = !标准单价
                            End If
                        End If
                    Else
                        objBillIncome.标准单价 = !现价
                    End If
                                        
                    If InStr(",5,6,7,", !收费类别) > 0 And bln药房单位 Then
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(!药房包装, 1), "0.00000")
                    Else
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价, "0.00000")
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

Public Function GetBillMoney(strNO As String) As Currency
'功能：获取一张正常记帐单的单据金额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Sum(实收金额) as 金额 From 病人费用记录 Where NO=[1] And 记录性质=2 And 记录状态 IN(0,1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
    If Not rsTmp.EOF Then GetBillMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(lng病人ID As Long) As Currency
'功能:获取指定病人的记帐划价单金额合计
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(Sum(实收金额),0) As 划价费用合计 From 病人费用记录 Where 记录状态=0 And 记帐费用=1 And 病人ID=" & lng病人ID
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlExpense")
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!划价费用合计
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAuditRecord(lng病人ID As Long, lng主页ID As Long) As ADODB.Recordset
'功能：获取指定病人的费用审批项目
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 项目Id From 病人审批项目 Where 病人ID=[1] And 主页ID=[2]"
    Set GetAuditRecord = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMoneyInfo(lng病人ID As Long, Optional curModiMoney As Currency, Optional blnInsure As Boolean) As ADODB.Recordset
'功能：获取指定病人的剩余额
'参数：blnInsure=是否排开医保病人的预结费用
    Dim rsTmp As New ADODB.Recordset
    Dim bln医保 As Boolean, lng主页ID As Long
    Dim strSQL As String
        
    On Error GoTo errH
    
    If blnInsure Then
        strSQL = "Select A.险类,A.主页ID From 病案主页 A,病人信息 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 And B.病人ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID)
        If Not rsTmp.EOF Then
            bln医保 = Not IsNull(rsTmp!险类)
            lng主页ID = rsTmp!主页ID
        End If
    End If
    
    strSQL = "Select Nvl(费用余额,0) as 费用余额,Nvl(预交余额,0) as 预交余额" & _
            " From 病人余额 Where 性质=1 And 病人ID=" & lng病人ID
    
    If curModiMoney <> 0 Then   '必须要用Union方式,如果直接去减,在病人余额无记录时,不会返回记录
        strSQL = strSQL & " Union All " & " Select -1* " & curModiMoney & " as 费用余额,0 as 预交余额 From Dual"
        strSQL = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSQL & ")"
    End If
            
    '如果为医保住院病人，则在费用余额中排开预结中的费用(用于报警)
    If blnInsure And bln医保 Then
        strSQL = strSQL & " Union All " & _
            " Select -1*Nvl(Sum(金额),0) as 费用余额,0 as 预交余额" & _
            " From 保险模拟结算 Where 病人ID=[1] And 主页ID=[2]"
        strSQL = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSQL & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiUnit(lngPatiID As Long) As Long
'功能：返回病人所属病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.当前病区ID From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.住院次数=B.主页ID And A.病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngPatiID)
    If Not rsTmp.EOF Then GetPatiUnit = Nvl(rsTmp!当前病区ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AdjustCpt(lngID As Long)
'功能：药品调价
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    
    strSQL = _
        "Select ID From 收费价目" & _
        " Where ((Sysdate Between 执行日期 and 终止日期) Or (Sysdate>=执行日期 And 终止日期 is NULL))" & _
        " And Nvl(变动原因,0)=0 And 收费细目ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lngID)
    Do While Not rsTmp.EOF
        strSQL = "zl_药品收发记录_Adjust(" & rsTmp!ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "mdlExpense")
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function BillisZeroLog(ByVal strNO As String) As Boolean
'功能：判断指定单据是否属于零耗费用登记
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long

    On Error GoTo errH

    strSQL = "Select 实收金额 From 病人费用记录 Where 记录状态 In(0,1,3) And 记录性质=2 And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
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

Public Function PatiCanBilling(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
'功能：检查指定病人是否具有相关权限
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, "出院未结强制记帐") > 0 _
        And InStr(strPrivs, "出院结清强制记帐") > 0 Then
        Exit Function
    End If
    
    strSQL = "Select A.姓名,B.出院日期,B.状态,X.费用余额" & _
        " From 病人信息 A,病案主页 B,病人余额 X" & _
        " Where A.病人ID=B.病人ID And A.病人ID=X.病人ID(+)" & _
        " And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!出院日期) And Nvl(rsTmp!状态, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "出院未结强制记帐") = 0 Then
            If Nvl(rsTmp!费用余额, 0) <> 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用未结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If InStr(strPrivs, "出院结清强制记帐") = 0 Then
            If Nvl(rsTmp!费用余额, 0) = 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用已结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillIdentical(ByVal strNO As String) As Boolean
'功能：判断指定的记帐单据中的状态是否一致,即是否同时存在审核和未审核的内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    BillIdentical = True
    
    On Error GoTo errH

    strSQL = _
        " Select Count(Distinct 登记时间) as 时间数," & _
        " Sum(Decode(记录状态,0,1,0)) as 未审核," & _
        " Sum(Decode(记录状态,0,0,1)) as 已审核" & _
        " From 病人费用记录" & _
        " Where 记录状态 IN(0,1,3) And NO=[1] And 记录性质=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO)
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

Public Function CheckValidity(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal dbl数量 As Double, Optional ByVal blnAsk As Boolean = True) As Boolean
'功能：检查卫生材料的灭菌效期是否过期
'说明：blnAsk=表示是否询问是否继续,否则为提醒
    Dim rsTmp As New ADODB.Recordset
    Dim curDate As Date, minDate As Date
    Dim strSQL As String, strName As String
    
    CheckValidity = True
    
    '仅一次性材料才判断
    '因为可能各批次灭菌效期不同,检查要用到的批次中最小的效期
    strSQL = _
        " Select C.名称,Nvl(B.批次,0) as 批次," & _
        " B.可用数量 as 库存,B.灭菌效期,Sysdate as 时间" & _
        " From 材料特性 A,药品库存 B,收费项目目录 C" & _
        " Where A.材料ID=B.药品ID And A.材料ID=C.ID And A.一次性材料=1" & _
        " And B.性质=1 And Nvl(B.可用数量,0)>0 And A.灭菌效期 is Not NULL" & _
        " And A.材料ID=[1] And B.库房ID=[2]" & _
        " Order by Nvl(B.批次,0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng材料ID, lng库房ID)
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

Public Function HaveBilling(ByVal strNO As String, Optional ByVal blnALL As Boolean = True, Optional ByVal strTime As String) As Integer
'功能：判断一张记帐单/表是否已经结帐
'参数：strNO=记帐单据号,不分门诊及住院
'      blnALL=是否对整张单据内容进行判断,否则只对未销帐部分进行判断(销帐时)
'返回：0-未结帐,1=已全部结帐,2-已部分结帐
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    
    On Error GoTo errH
        
    '求未作废的费用行
    strSQL = _
        " Select 序号 From (" & _
        " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号," & _
        " Avg(Nvl(付数, 1) * 数次) As 数量" & _
        " From 病人费用记录" & _
        " Where NO=[1] And 记录性质=2" & _
        " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    
    '求每行的结帐情况
    strSQL = _
        "Select Nvl(价格父号,序号) as 序号,Sum(Nvl(结帐金额,0)) as 结帐金额" & _
        " From 病人费用记录" & _
        " Where NO=[1] And 记录性质 IN(2,12)" & _
        IIF(Not blnALL, " And Nvl(价格父号,序号) IN(" & strSQL & ")", "") & _
        IIF(strTime <> "", " And 登记时间=[2]", "") & _
        " Group by Nvl(价格父号,序号)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", strNO, CDate(IIF(strTime = "", "1990-01-01", strTime)))
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '单据行数
        rsTmp.Filter = "结帐金额<>0"
        If rsTmp.EOF Then
            HaveBilling = 0 '无结帐行
        ElseIf rsTmp.RecordCount = lngTmp Then
            HaveBilling = 1 '全部行已结帐
        ElseIf rsTmp.RecordCount > 0 Then
            HaveBilling = 2 '部分行已结帐
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", intNum)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSQL = Format(CDate(Format(rsTmp!日期, "YYYY-MM-dd")) - CDate(Format(rsTmp!日期, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBalanceRoll(ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, Optional ByVal blnBat As Boolean) As Boolean
'功能：(住院)对要回退的医嘱对应的费用的结帐情况进行检查(一个病人一次住院的)
'参数：blnBat=是否要进行批量回退
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intInsure As Integer
    
    On Error GoTo errH
        
    '取要回退的记帐NO
    If blnBat Then
        strSQL = "Select Distinct NO From 病人医嘱发送 Where 记录性质=2 And 发送号=[1]"
    Else
        strSQL = "Select Distinct A.NO From 病人医嘱发送 A,病人医嘱记录 B" & _
            " Where A.医嘱ID=B.ID And A.记录性质=2 And A.发送号=[1] And (B.ID=[2] Or B.相关ID=[2])"
    End If
    '取这些NO的结帐情况(非划价未销帐)
    strSQL = "Select A.NO,Nvl(A.价格父号,A.序号) as 序号,Sum(Nvl(A.结帐金额,0)) as 结帐金额" & _
        " From 病人费用记录 A,(" & strSQL & ") B Where A.NO=B.NO And A.记录性质 IN(2,12) And A.记录状态=1" & _
        " Group by A.NO,Nvl(A.价格父号,A.序号) Having Sum(Nvl(A.结帐金额,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng发送号, lng医嘱ID)
    If Not rsTmp.EOF Then
        strSQL = "Select A.险类 From 病案主页 A,病人医嘱记录 B" & _
            " Where Rownum=1 And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng医嘱ID)
        If Not rsTmp.EOF Then intInsure = Nvl(rsTmp!险类, 0)
        If intInsure <> 0 Then '先对医保的限制进行检查
            If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , intInsure) Then
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
    Dim strSQL As String, lng发送号 As Long
    
    If gbytBillOpt = 0 Then
        CheckAdviceBalanceRevoke = True
        Exit Function
    End If
    
    On Error GoTo errH
    
    '医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
    strSQL = "Select Distinct 发送号 From 病人医嘱发送" & _
        " Where 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng医嘱ID)
    If rsTmp.EOF Then Exit Function
    lng发送号 = rsTmp!发送号
    
    '部份条件见"ZL_病人医嘱记录_作废"
    strSQL = "Select A.NO,Nvl(A.价格父号,A.序号) as 序号,Sum(Nvl(A.结帐金额,0)) as 结帐金额" & _
        " From 病人费用记录 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 I" & _
        " Where A.NO=B.NO And A.记录性质 IN(2,12) And A.记录状态=1 And B.医嘱ID=C.ID" & _
        " And B.记录性质=2 And C.诊疗项目ID=I.ID And B.发送号=[1] And (C.ID=[2] Or C.相关ID=[2])" & _
        " And (" & _
            " A.收费类别 Not In ('5','6','7','E')" & _
            " Or A.收费类别='E' And I.操作类型 Not In ('2','3','4')" & _
            " Or A.收费类别 In ('5','6','7') And Nvl(A.执行状态,0)=0" & _
            " Or Exists(Select 参数值 From 系统参数表 Where 参数号=68 And Nvl(参数值,0)=0)" & _
            " )" & _
        " Group by A.NO,Nvl(A.价格父号,A.序号) Having Sum(Nvl(A.结帐金额,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng发送号, lng医嘱ID)
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

Public Function GetExamineItem(ByVal strItems As String, ByVal lngMediCareID As Long) As ADODB.Recordset
'功能:返回指定险类的收费项目要求审批的记录集
'参数:strItems-收费细目ID串,例如:"2369,2367,2368"
'     lngMediCareID-险类,例如:901
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select A.收费细目id" & vbNewLine & _
            "From 保险支付项目 A ,Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B" & vbNewLine & _
            "Where A.险类 = [1] And A.要求审批 = 1 And A.收费细目id = B.Column_Value"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngMediCareID, strItems)
    
    Set GetExamineItem = rsTmp
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

Public Function CheckExamine(ByRef ObjBillDetails As BillDetails, ByRef rsMedAudit As ADODB.Recordset, ByRef lngMediCareID As Long) As Boolean
'功能:根据给定的收费项目对象集和病人审批项目记录集检查相应的收费项目是否需要审批
    Dim i As Long, strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    For i = 1 To ObjBillDetails.Count
        strTmp = strTmp & "," & ObjBillDetails(i).收费细目ID
    Next
    Set rsTmp = GetExamineItem(Mid(strTmp, 2), lngMediCareID)
    
    strTmp = ""
    For i = 1 To rsTmp.RecordCount
        rsMedAudit.Filter = "项目ID=" & rsTmp!收费细目ID
        If rsMedAudit.RecordCount = 0 Then strTmp = strTmp & "," & GetRowByFeeItemID(ObjBillDetails, rsTmp!收费细目ID)
        rsTmp.MoveNext
    Next
    
    If strTmp <> "" Then
        MsgBox "第" & Mid(strTmp, 2) & "行收费项目要求审批,当前病人未被批准使用!", vbInformation, gstrSysName
        CheckExamine = False: Exit Function
    End If
    CheckExamine = True
End Function

