VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsureSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstr个帐名称 As String
Private mclsInsure As clsInsure

Private mlng病人ID As Long
Private mlng险类 As Long

Private mbln费用结算分离 As Boolean

Private mbln医保不走票号 As Boolean
Private mbln门诊结算作废 As Boolean
Private mbln多单据分单据结算 As Boolean
Private mbln一次结算分单据退费 As Boolean

Public Property Set Insure(ByVal vData As clsInsure)
    Set mclsInsure = vData
End Property

Public Property Get Insure() As clsInsure
    Set Insure = mclsInsure
End Property

Public Property Let 病人ID(ByVal vData As Long)
    mlng病人ID = vData
End Property

Public Property Get 病人ID() As Long
    病人ID = mlng病人ID
End Property

Public Property Let 个帐名称(ByVal vData As String)
    mstr个帐名称 = vData
End Property

Public Property Get 个帐名称() As String
    个帐名称 = mstr个帐名称
End Property

Public Property Let 险类(ByVal vData As Long)
    mlng险类 = vData
End Property

Public Property Get 险类() As Long
    险类 = mlng险类
End Property

Public Function MakePreSwapData() As ADODB.Recordset
    '创建一个预结算记录结构
    '返回:医保相关数据的数据集结构
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    With rsTmp.Fields
        .Append "单据序号", adBigInt, 50, adFldIsNullable
        .Append "费别", adVarChar, 50, adFldIsNullable
        .Append "NO", adVarChar, 8, adFldIsNullable
        .Append "序号", adBigInt, , adFldIsNullable '问题:42961
        .Append "实际票号", adVarChar, 20, adFldIsNullable
        .Append "结算时间", adDBTimeStamp, , adFldIsNullable
        .Append "病人ID", adBigInt, , adFldIsNullable
        .Append "收费类别", adVarChar, 2, adFldIsNullable
        .Append "收据费目", adVarChar, 20, adFldIsNullable
        .Append "计算单位", adVarChar, 50, adFldIsNullable
        .Append "开单人", adVarChar, 100, adFldIsNullable
        .Append "收费细目ID", adBigInt, , adFldIsNullable
        .Append "数量", adDouble, , adFldIsNullable
        .Append "单价", adDouble, , adFldIsNullable
        .Append "实收金额", adCurrency, , adFldIsNullable
        .Append "统筹金额", adCurrency, , adFldIsNullable
        .Append "保险支付大类ID", adBigInt, , adFldIsNullable
        .Append "是否医保", adBigInt, , adFldIsNullable
        .Append "保险编码", adVarChar, 50, adFldIsNullable
        .Append "摘要", adVarChar, 2000, adFldIsNullable
        .Append "是否急诊", adBigInt, , adFldIsNullable
        .Append "开单部门ID", adBigInt, , adFldIsNullable
        .Append "执行部门ID", adBigInt, , adFldIsNullable
    End With
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set MakePreSwapData = rsTmp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MakePreSwapDataFromDB(ByVal bytMode As Byte, ByVal varValue As Variant, _
    ByVal strInvoice As String, Optional ByVal bln补结算 As Boolean) As ADODB.Recordset
    '根据单据号或结帐ID创建一个记录信息(以售价单位)
    '入参:
    '   bytMode 数据获取方式，0-按单据号获取，1-按结帐ID获取
    '   varValue bytMode=0时，单据号；bytMode=1时，结帐ID
    '   strInvoice 开始发票号
    '   bln补结算 是否补结算
    '出参:
    '返回:医保相关数据的数据集
    Dim strTable  As String, strWhere As String
    Dim p As Integer, strSQL As String, strNos As String
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    Dim lng结帐ID As Long, strNo As String
    
    On Error GoTo ErrHandler
    Set rsTmp = MakePreSwapData()
    
    If bytMode = 0 Then
        strWhere = strWhere & " And a.No=[2]"
        strNo = CStr(varValue)
    Else
        If bln补结算 Then
           strTable = ",(Select distinct 收费结帐ID From 费用补充记录 Where 结算ID=[1]) B"
           strWhere = " And A.结帐ID=b.收费结帐ID"
        Else
            strWhere = " And A.结帐ID=[1]"
        End If
        lng结帐ID = CLng(varValue)
    End If

    strSQL = _
        " Select A.NO,Nvl( A.价格父号, A.序号) as 序号,To_char(max(A.登记时间),'YYYY-MM-DD HH24:MI:SS') as 结算时间," & _
        "       A.病人ID,A.费别,A.收费类别,A.收据费目,A.计算单位,A.开单人," & _
        "       A.收费细目ID,A.保险大类ID As 保险支付大类ID,Nvl(A.保险项目否,0) As 是否医保,A.保险编码," & _
        "       Avg(Nvl(A.付数,0)*A.数次) As 数量,Avg(A.标准单价) As 单价," & _
        "       Sum(A.实收金额) As 实收金额,Sum(A.统筹金额) As 统筹金额,max(A.摘要) as 摘要," & _
        "       nvl(A.加班标志,0) as 是否急诊,A.开单部门ID,A.执行部门ID,A.结帐ID" & _
        " From 门诊费用记录 A" & strTable & _
        " Where Mod(A.记录性质,10)=1 " & strWhere & _
        " Group By A.NO, Nvl(A.价格父号, A.序号),A.病人id, A.费别, A.收费类别, A.收据费目, A.计算单位," & _
        "       A.开单人, A.收费细目id, A.保险大类id, Nvl(A.保险项目否, 0), A.保险编码, A.摘要, Nvl(A.加班标志, 0)," & _
        "       A.开单部门id, A.执行部门id,A.结帐ID"
    
    strSQL = _
        " Select '" & strInvoice & "' as 实际票号,A.NO,A.序号,max(A.结算时间) as 结算时间," & _
        "       A.病人ID,A.费别,A.收费类别,A.收据费目,A.计算单位,A.开单人," & _
        "       A.收费细目ID,A.保险支付大类ID,A.是否医保,A.保险编码," & _
        "       sum(A.数量) as 数量,max(A.单价) As 单价, Sum(A.实收金额) As 实收金额, " & _
        "       Sum(A.统筹金额) As 统筹金额,max(A.摘要) as 摘要," & _
        "       Max(A.是否急诊) as 是否急诊,max(A.开单部门ID) as 开单部门ID,max(A.执行部门ID ) as 执行部门ID " & _
        " From (" & strSQL & ") A" & _
        " Group By A.NO,A.序号,A.病人id, A.费别, A.收费类别, A.收据费目, A.计算单位, A.开单人, A.收费细目id, A.保险支付大类ID, " & _
        "       A.是否医保, A.保险编码" & _
        " Order by NO,序号"
    Set rsNo = zlDatabase.OpenSQLRecord(strSQL, "获取重新收费数据-医保", lng结帐ID, strNo)
    
    With rsNo
        p = 0: strNos = ""
        Do While Not rsNo.EOF
            If InStrEx(strNos, Nvl(!NO)) = False Then
                strNos = strNos & "," & Nvl(!NO)
                p = p + 1
            End If
            
            rsTmp.AddNew
            rsTmp!单据序号 = p
            rsTmp!费别 = !费别
            rsTmp!NO = Nvl(!NO)
            rsTmp!序号 = Val(Nvl(!序号))
            rsTmp!实际票号 = Nvl(!实际票号)
            rsTmp!结算时间 = !结算时间
            rsTmp!病人ID = Val(Nvl(!病人ID))
            rsTmp!收费类别 = Nvl(!收费类别)
            rsTmp!收据费目 = Nvl(!收据费目)
            rsTmp!开单人 = Nvl(!开单人)
            rsTmp!收费细目ID = Val(Nvl(!收费细目ID))
            rsTmp!计算单位 = Nvl(!计算单位)
            rsTmp!数量 = Val(Nvl(!数量))
            rsTmp!单价 = Val(Nvl(!单价))
            rsTmp!实收金额 = Val(Nvl(!实收金额))
            rsTmp!统筹金额 = Val(Nvl(!统筹金额))
            rsTmp!保险支付大类ID = IIf(Val(Nvl(!保险支付大类ID)) = 0, Null, Val(Nvl(!保险支付大类ID)))
            rsTmp!是否医保 = Val(Nvl(!是否医保))
            rsTmp!保险编码 = Nvl(!保险编码)
            rsTmp!摘要 = Nvl(!摘要)
            rsTmp!是否急诊 = Val(Nvl(!是否急诊))
            rsTmp!开单部门ID = Val(Nvl(!开单部门ID))
            rsTmp!执行部门ID = Val(Nvl(!执行部门ID))
            rsTmp.Update
            
            rsNo.MoveNext
        Loop
    End With
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakePreSwapDataFromDB = rsTmp
    
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecutePreSwap( _
    rsRecord As ADODB.Recordset, _
    ByRef colBalance As Collection, _
    ByRef strInsureInvoice As String, _
    Optional ByVal blnReBalance As Boolean, _
    Optional ByVal bln补结算 As Boolean) As Boolean
    '门诊预结算
    '入参：
    '   blnReBalance 是否重新结算
    '出参：
    '   colBalance 预结算结果集(每张单据对应一个BalanceMoneys对象元素),多单据一次结算时存在第一张单据中
    '   strInsureInvoice 医保返回发票号
    Dim rsRecord As ADODB.Recordset
    Dim strBalance As String, strAdvance As String
    Dim varBalance As Variant, varItem As Variant, str结算方式 As String
    Dim p As Long, i As Long
    
    On Error GoTo ErrHandler
    mbln医保不走票号 = False: strInsureInvoice = ""
    
    '2-多单据分单据结算
    If mbln多单据分单据结算 Then
        For p = 1 To objBill.Pages.Count
            strBalance = ""
            If blnErrBill Then
                '检查该张单据是否已成功医保结算
                rsBalance.Filter = "No='" & objBill.Pages(p).NO & "'"
                Do While Not rsBalance.EOF
                    strBalance = strBalance & IIf(strBalance = "", "", "||")
                    strBalance = strBalance & Nvl(rsBalance!结算方式) & "|" & Val(Nvl(rsBalance!金额))
                    rsBalance.MoveNext
                Loop
            End If
            
            If strBalance <> "" Then
                Call SetBalanceVal(colBalance, p, strBalance)
            Else
                strBalance = ""
                strAdvance = IIf(blnReBalance, "1", "")
                
                If Not mclsInsure.ClinicPreSwap(rsRecord, strBalance, mlng险类, strAdvance) Then
                    ShowMsgbox "第 " & p & " 张单据预结算失败。"
                    Exit Function
                End If
                
                '只要有一张单据自动走票号，都要走票号
                If strAdvance <> "" And InStr(strAdvance, "|") = 0 Then    '医保票据号 Then
                    '38821,格式:票据号;是否不走票号(1-不走票号;0-自动走票号)
                    varItem = Split(strAdvance & ";", ";")
                    strInsureInvoice = varItem(0)
                    mbln医保不走票号 = mbln医保不走票号 And Val(varItem(1)) = 1
                End If
                
                '报销方式;金额;是否允许修改|....
                If strBalance <> "" Then
                    strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
                    Call SetBalanceVal(colBalance, p, strBalance)
                End If
            End If
        Next
        ExecutePreSwap = True: Exit Function
    End If
    
    '0-多单据一次结算,1-多单据一次结算分单据退费
    strBalance = ""
    strAdvance = IIf(bln补结算 And blnReBalance, "3", IIf(bln补结算, "2", "1"))
    
    If Not mclsInsure.ClinicPreSwap(rsRecord, strBalance, mlng险类, strAdvance) Then
        ShowMsgbox "单据预结算失败。"
        Exit Function
    End If
    
    If strAdvance <> "" And InStr(strAdvance, "|") = 0 Then
        '38821:strAdvance:发票号;是否不走票据号
        varItem = Split(strAdvance & ";", ";")
        strInsureInvoice = varItem(0)
        mbln医保不走票号 = Val(varItem(1)) = 1
    End If
    
    '报销方式;金额;是否允许修改|....
    If strBalance <> "" Then
        If mbln一次结算分单据退费 Then
            '1-多单据一次结算分单据退费
            '单据序号:结算方式;金额;是否允许修改|...||单据序号:结算方式;金额;是否允许修改|...||...
            varBalance = Split(strBalance, "||")
            For i = 0 To UBound(varBalance)
                If InStr(varBalance(i), ":") = 0 Then
                    ShowMsgbox "单据预结算返回结算结果格式不正确。"
                    Exit Function
                End If
                
                varItem = Split(varBalance(i), ":")
                p = Val(varItem(0)): str结算方式 = varItem(1)
                If p < 1 Or p > colBalance.Count Then
                    ShowMsgbox "单据预结算返回结算结果格式不正确。"
                    Exit Function
                End If
                
                str结算方式 = Replace(Replace(str结算方式, "|", "||"), ";", "|")
                Call SetBalanceVal(colBalance, p, str结算方式)
            Next
        Else
            '0-多单据一次结算
            strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
            Call SetBalanceVal(colBalance, 1, strBalance)
        End If
    End If
    
    ZlExecuteInsurePreSwap = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsCanCancel(ByVal lng结算序号 As Long) As Boolean
    '检查医保结算是否允许作废
    Dim strSQL As String, rsBalance As ADODB.Recordset
    Dim strBalance As String

    On Error GoTo ErrHandler
    If mbln门诊结算作废 = False Then
        ShowMsgbox "医保不支持门诊结算作废，不能作废本次结算！"
        Exit Function
    End If

    strSQL = _
        "Select a.结算方式" & vbNewLine & _
        "From 病人预交记录 A, 结算方式 B" & vbNewLine & _
        "Where a.结算方式 = b.名称 And b.性质 In(3,4) And a.卡类别ID Is Null And a.结算序号 = [1]"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, "获取医保结算数据", lng结算序号)
    Do While Not rsBalance.EOF
        If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mlng险类, Nvl(rsBalance!结算方式)) Then
            strBalance = strBalance & "," & Nvl(rsBalance!结算方式)
        End If
        rsBalance.MoveNext
    Loop

    If strBalance <> "" Then
        ShowMsgbox "医保结算方式【" & Mid(strBalance, 2) & "】不支持作废，不能作废本次结算！"
        Exit Function
    End If
    IsCanCancel = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteInsureInfoUpdate(ByVal lng结帐ID As Long, ByRef objBalanceBills As BalanceBills) As Boolean
    '更新重收记录的保险信息
    '参数:
    '返回:所有重收记录的保险信息更新成功返回True，否则返回False
    Dim strSQL As String, rsReCharge As ADODB.Recordset
    Dim strBXInfo As String, strPreNo As String
    Dim cur实收金额 As Currency, cur统筹金额 As Currency, bln保险项目 As Boolean
    Dim blnTrans As Boolean, cllReChargePro As Collection
    Dim objBalanceBill As BalanceBill
    
    On Error GoTo ErrHandler
    Set objBalanceBills = New BalanceBills
    If mbln费用结算分离 = False Then
        strSQL = _
            "Select a.Id, a.No, a.序号, a.病人id, a.收费细目id, Nvl(a.付数, 1) * Nvl(a.数次, 0) As 数量," & vbNewLine & _
            "       Nvl(a.实收金额, 0) As 实收金额, a.摘要, Nvl(a.保险项目否, 0) As 保险项目否, " & vbNewLine & _
            "       a.保险大类id, Nvl(a.统筹金额, 0) As 统筹金额, a.保险编码, a.费用类型" & vbNewLine & _
            "From 门诊费用记录 A" & vbNewLine & _
            "Where a.记录性质 = 11 And a.结帐id = [1]"
    Else
        strSQL = _
            "Select a.No, a.序号, a.病人id, a.收费细目id, Sum(Nvl(a.付数, 1) * a.数次) As 数量," & vbNewLine & _
            "       Nvl(Sum(a.实收金额), 0) As 实收金额, Max(Decode(a.记录状态,0,null,a.摘要)) As 摘要," & vbNewLine & _
            "       Max(Decode(a.记录状态,0,null,Nvl(a.保险项目否, 0))) As 保险项目否, " & vbNewLine & _
            "       Max(Decode(a.记录状态,0,null,a.保险大类id)) As 保险大类id," & vbNewLine & _
            "       Nvl(Sum(a.统筹金额), 0) As 统筹金额," & vbNewLine & _
            "       Max(Decode(a.记录状态,0,null,a.保险编码)) As 保险编码," & vbNewLine & _
            "       Max(Decode(a.记录状态,0,null,a.费用类型)) As 费用类型," & vbNewLine & _
            "       Max(Decode(a.记录状态,0,null,a.Id)) As ID" & vbNewLine & _
            "From 门诊费用记录 A, 费用结算对照 B" & vbNewLine & _
            "Where a.ID = b.费用ID And b.是否重收 = 1 And b.结帐id = [1]" & vbNewLine & _
            "Group By a.No,a.序号,a.病人id,a.收费细目id" & vbNewLine & _
            "Having Nvl(Sum(Nvl(a.付数, 1) * a.数次),0) <> 0"
    End If
    Set rsReCharge = zlDatabase.OpenSQLRecord(strSQL, "获取重收费用记录", lng结帐ID)
    With rsReCharge
        If .RecordCount > 0 Then
            Set cllReChargePro = New Collection
            .Sort = "NO,序号"
            Do While Not .EOF
                If strPreNo <> Nvl(!NO) Then
                    If strPreNo <> "" Then
                        objBalanceBills.AddItem objBalanceBill
                    End If
                    
                    Set objBalanceBill = New BalanceBill
                    objBalanceBill.NO = Nvl(!NO)
                    strPreNo = Nvl(!NO)
                End If
                
                '保险项目否(0/1);保险大类ID;进入统筹金额;保险项目编码;摘要;费用类型
                strBXInfo = mclsInsure.GetItemInsure(Nvl(!病人ID), Nvl(!收费细目ID), Val(Nvl(!实收金额)), _
                    True, mlng险类, Nvl(!摘要) & "||" & Val(Nvl(!数量)))
                If strBXInfo <> "" Then
                    '  Zl_门诊收费记录_Update
                    strSQL = "Zl_门诊收费记录_Update("
                    '  Id_In         In 门诊费用记录.Id%Type,
                    strSQL = strSQL & Nvl(!ID) & ","
                    '  保险大类id_In In 门诊费用记录.保险大类id%Type,
                    strSQL = strSQL & ZVal(Split(strBXInfo, ";")(1)) & ","
                    '  保险项目否_In In 门诊费用记录.保险项目否%Type,
                    strSQL = strSQL & Val(Split(strBXInfo, ";")(0)) & ","
                    '  保险编码_In   In 门诊费用记录.保险编码%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(3)) & "',"
                    '  费用类型_In   In 门诊费用记录.费用类型%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(5)) & "',"
                    '  统筹金额_In   In 门诊费用记录.统筹金额%Type,
                    strSQL = strSQL & Format(Val(Split(strBXInfo, ";")(2)), gstrDec) & ","
                    '  摘要_In       In 门诊费用记录.摘要%Type
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(4)) & "')"
                    zlAddArray cllReChargePro, strSQL
                    
                    cur统筹金额 = CCur(Val(Split(strBXInfo, ";")(2)))
                    bln保险项目 = Val(Split(strBXInfo, ";")(0)) = 1
                Else
                    cur统筹金额 = Val(Nvl(!统筹金额))
                    bln保险项目 = Val(Nvl(!保险项目否)) = 1
                End If
                
                '统计保险金额
                cur实收金额 = Val(Nvl(!实收金额))
                If cur统筹金额 = 0 Or Not bln保险项目 Then
                    '以原始金额为准,不管分币处理
                    objBalanceBill.全自付 = objBalanceBill.全自付 + cur实收金额
                Else
                    objBalanceBill.进入统筹 = objBalanceBill.进入统筹 + cur统筹金额
                    '以原始金额为准,不管分币处理
                    objBalanceBill.先自付 = objBalanceBill.先自付 + cur实收金额 - cur统筹金额
                End If
                objBalanceBill.实收合计 = objBalanceBill.实收合计 + CCur(Val(Nvl(!实收金额)))
                
                rsReCharge.MoveNext
            Loop
            If strPreNo <> "" Then
                objBalanceBills.AddItem objBalanceBill
            End If
            
            '费用结算分离时不更新原始费用记录
            If mbln费用结算分离 = False Then
                '执行过程
                blnTrans = True
                zlExecuteProcedureArrAy cllReChargePro, "执行保险信息更新", True, True
                blnTrans = False
            End If
        End If
    End With
    ExecuteInsureInfoUpdate = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
