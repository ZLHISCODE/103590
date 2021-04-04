VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinicInsureDelSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'保持属性值的局部变量
Private mlng病人ID As Long '局部复制
Private mstr姓名 As String '局部复制
Private mstr性别 As String '局部复制
Private mstr年龄 As String '局部复制
Private mstr门诊号 As String '局部复制
Private mstr住院号 As String '局部复制
Private mstr医疗付款方式 As String '局部复制
Private mstr个人帐户 As String '局部复制
Private mint险类 As Integer '局部复制
Private mbln门诊结算作废 As Boolean '局部复制
Private mbln退费后打印回单 As Boolean '局部复制
Private mbln补充结算 As Boolean '局部复制
Private mlng结帐ID As Long '局部复制
Private mbln医保接口打印票据 As Boolean '局部复制
Private mlng冲销ID As Long '局部复制
'字段:类型,NO,结算方式,结算性质,冲预交,卡类别ID,结算卡序号,卡号,消费卡ID,交易流水号,交易说明,原结帐ID,关联交易ID
Private mrsBalance As ADODB.Recordset '局部复制
Private mbln分单据退费 As Boolean '局部复制

Public Property Let 医疗付款方式(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医疗付款方式 = 5
    mstr医疗付款方式 = vData
End Property


Public Property Get 医疗付款方式() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医疗付款方式
    医疗付款方式 = mstr医疗付款方式
End Property



Public Property Let 住院号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.住院号 = 5
    mstr住院号 = vData
End Property


Public Property Get 住院号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.住院号
    住院号 = mstr住院号
End Property



Public Property Let 门诊号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊号 = 5
    mstr门诊号 = vData
End Property


Public Property Get 门诊号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊号
    门诊号 = mstr门诊号
End Property



Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mstr年龄 = vData
End Property


Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.年龄
    年龄 = mstr年龄
End Property



Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mstr性别 = vData
End Property


Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.性别
    性别 = mstr性别
End Property



Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mstr姓名 = vData
End Property


Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    姓名 = mstr姓名
End Property




Public Function InitData(ByVal str个人帐户 As String) As Boolean
    '初始化公共基础数据
    mstr个人帐户 = str个人帐户
    InitData = True
End Function

Public Property Get 分单据退费() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.分单据退费
    分单据退费 = mbln分单据退费
End Property

Public Property Set rsBalance(ByVal vData As ADODB.Recordset)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.rsBalance = Form1
    Set mrsBalance = vData
    
    If mrsBalance Is Nothing Then Exit Property
    If mrsBalance.State <> 1 Then Exit Property
    
    '字段:类型,NO,结算方式,结算性质,冲预交,卡类别ID,结算卡序号,卡号,消费卡ID,交易流水号,交易说明,原结帐ID,关联交易ID
    mrsBalance.Filter = "类型=" & gBalanceType.gEM_医保 & " And NO<>null"
    mbln分单据退费 = Not mrsBalance.EOF
End Property

Public Property Get rsBalance() As ADODB.Recordset
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.rsBalance
    Set rsBalance = mrsBalance
End Property

Public Property Let 冲销ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.冲销ID = 5
    mlng冲销ID = vData
End Property

Public Property Get 冲销ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.冲销ID
    冲销ID = mlng冲销ID
End Property

Public Property Get 医保接口打印票据() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医保接口打印票据
    医保接口打印票据 = mbln医保接口打印票据
End Property

Public Property Let 结帐ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐ID = 5
    mlng结帐ID = vData
    
    If gclsInsure Is Nothing Then Exit Property
    mbln医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, mlng病人ID, mint险类, CStr(mlng结帐ID))
End Property

Public Property Get 结帐ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐ID
    结帐ID = mlng结帐ID
End Property

Public Property Let 补充结算(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.补充结算 = 5
    mbln补充结算 = vData
End Property

Public Property Get 补充结算() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.补充结算
    补充结算 = mbln补充结算
End Property

Public Property Get 退费后打印回单() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.退费后打印回单
    退费后打印回单 = mbln退费后打印回单
End Property

Public Property Get 门诊结算作废() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊结算作废
    门诊结算作废 = mbln门诊结算作废
End Property

Public Property Let 险类(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类 = 5
    mint险类 = vData
    
    If gclsInsure Is Nothing Then Exit Property
    mbln退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, mlng病人ID, mint险类)
    mbln门诊结算作废 = gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mint险类)
End Property

Public Property Get 险类() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类
    险类 = mint险类
End Property


Public Property Get 个人帐户() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.个人帐户
    个人帐户 = mstr个人帐户
End Property

Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mlng病人ID = vData
End Property

Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病人ID
    病人ID = mlng病人ID
End Property

Public Function ExecuteClinicDelSwap(cllBillPro As Collection, ByRef blnCommit As Boolean) As Boolean
    '执行退费
    '说明：
    '   由调用者开启事务，执行成功，提交数据；执行失败，回滚数据
    
    blnCommit = False
    If mint险类 = 0 Then
        MsgBox "险类无效！", vbInformation, gstrSysName
        Exit Function
    End If
    If rsBalance Is Nothing Then
        MsgBox "原始结算数据无效！", vbInformation, gstrSysName
        Exit Function
    End If
    If rsBalance.State <> 1 Then
        MsgBox "原始结算数据无效！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mbln补充结算 Then
        If ExecuteReplenishDel(cllBillPro) Then
            blnCommit = True
            ExecuteClinicDelSwap = True
        End If
    Else
        ExecuteClinicDelSwap = ExecuteClinicDel(cllBillPro, blnCommit)
    End If
End Function

Private Function GetAllNos() As String
    '获取多单据分单据结算时的所有单据号，单据号降序排列
    Dim strNos As String
    
    '字段:类型,NO,结算方式,结算性质,冲预交,卡类别ID,结算卡序号,卡号,消费卡ID,交易流水号,交易说明,原结帐ID,关联交易ID
    mrsBalance.Filter = "类型=" & gBalanceType.gEM_医保 & " And NO<>null"
    mrsBalance.Sort = "No Desc"
    Do While Not mrsBalance.EOF
        If InStr("," & strNos & ",", "," & NVL(mrsBalance!NO) & ",") = 0 Then
            strNos = strNos & "," & NVL(mrsBalance!NO)
        End If
        mrsBalance.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetAllNos = strNos
End Function

Private Function ExecuteClinicDel(cllBillPro As Collection, ByRef blnCommit As Boolean) As Boolean
    '门诊收费退费
    '入参：
    '   bytMode 医保退费模式：0-多单据一次结算,1-多单据分单据退费
    '   strSavedAdvance 需要修正的原始医保结算信息
    '   strNos、lng病人ID、intInsure、bln门诊结算作废、str个人帐户 bytMode=1时传入
    '   strNos 按升序排列的单据号，格式：A001,A002,...
    '说明:需要在外层启用事务,正常退费后,该过程已提交,不需要调用者提交
    '     如果失败,则事务将回退(主要是避免弹出界面造成死锁)
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAllAdvance As String, strAdvance As String
    Dim colBalance As Collection, varNos As Variant
    Dim p As Integer, strNO As String, str结算方式 As String
    Dim blnCurrentCommit As Boolean
    Dim rsDelBalance As ADODB.Recordset
    
    On Error GoTo ErrHandler
    blnTrans = True
    zlExecuteProcedureArrAy cllBillPro, "医保结算退费", True, False
    
    strAllAdvance = GetYBBalance(mrsBalance, "", True, True)
    Call SaveInsureBalance(mlng病人ID, mlng冲销ID, strAllAdvance)
    
    If Not mbln分单据退费 Then
        strAdvance = mlng冲销ID & "|" & "0"
        If Not gclsInsure.ClinicDelSwap(mlng结帐ID, True, mint险类, strAdvance) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        blnTransMedicare = True
    Else
        Set colBalance = New Collection
        Set rsDelBalance = GetBalanceDetail(0, mlng冲销ID, 1)
        
        '93337,退费时按单据号倒序进行接口调用
        varNos = Split(GetAllNos(), ",")
        For p = UBound(varNos) To LBound(varNos) Step -1
            colBalance.Add New BalanceMoneys
            strNO = varNos(p)
            
            '检查该张单据是否已成功医保结算
            str结算方式 = GetYBBalance(rsDelBalance, strNO)
            
            If str结算方式 <> "" Then
                Call SetBalanceVal(colBalance, colBalance.Count, str结算方式)
            Else
                str结算方式 = GetYBBalance(mrsBalance, strNO, True, True)
                'str结算方式 为空，表示医保不支持医保作废
                If str结算方式 <> "" Then
                    Call SaveInsureBalanceDetail(mlng冲销ID, strNO, str结算方式)
                    
                    strAdvance = mlng冲销ID & "|" & "0" & "|" & strNO
                    '因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
                    If Not gclsInsure.ClinicDelSwap(mlng结帐ID, True, mint险类, strAdvance) Then
                        gcnOracle.RollbackTrans
                        If blnCurrentCommit Then Call CorrectInsureErrBalance(mlng病人ID, mlng冲销ID)  '医保结算校对
                        Exit Function
                    End If
                    blnTransMedicare = True
                    
                    If InsureCheck(str结算方式, strAdvance) Then
                        Call SaveInsureBalanceDetail(mlng冲销ID, strNO, strAdvance)
                        str结算方式 = strAdvance
                    End If
                    Call SetBalanceVal(colBalance, colBalance.Count, str结算方式)
                    gcnOracle.CommitTrans: blnTrans = False
                    blnCommit = True
                    blnCurrentCommit = True
                    
                    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mint险类)
                    blnTransMedicare = False
                    
                    gcnOracle.BeginTrans: blnTrans = True
                End If
            End If
        Next
        
        '全部成功，返回总的结算方式
        strAdvance = GetBalanceStr(colBalance)
    End If
    
    '校对结算结果
    If InsureCheck(strAllAdvance, strAdvance) Then
        Call SaveInsureBalance(mlng病人ID, mlng冲销ID, strAdvance)
    End If
    
    Call BalanceOver(mlng冲销ID)
    gcnOracle.CommitTrans: blnTrans = False
    blnCommit = True
    
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mint险类)
    ExecuteClinicDel = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
    
    '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mint险类)
    If blnCurrentCommit Then Call CorrectInsureErrBalance(mlng病人ID, mlng冲销ID) '医保结算校对
End Function

Private Function ExecuteReplenishDel(cllBillPro As Collection) As Boolean
    '补充结算退费
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAllAdvance As String, strAdvance As String
    
    On Error GoTo ErrHandler
    blnTrans = True
    zlExecuteProcedureArrAy cllBillPro, "医保补充结算退费", True, False
    
    strAllAdvance = GetYBBalance(mrsBalance, "", True, True)
    Call SaveInsureBalance(mlng病人ID, mlng冲销ID, strAllAdvance, True)
    
    'ClinicDelSwap(lngStlID As Long, Optional ByVal bln退费 As Boolean = True, _
        Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
    '功能:将门诊退费的明细和结算数据转发送医保前置服务器确认
    '入参:lngStlID-将要退的费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '     bln退费 -表明是退费交易还是改费交易在调用本接口
    '     strAdvance:格式:冲销ID|补充结算标志|…,每位|分隔
    '           第一位:传入冲销ID,医保可以根据冲销ID来进行取数
    '           第二位:补充结算标志,1-补充结算调和;0非补充结算调用
    '           第三位:NO:当前结算的NO
    '           第四位后: 待以后扩展
    '     注意：
    '           strAdvance在10.34.0以前(不含补允结算)
    '               多单据一次结算时,传入的是原结帐IDs:结帐ID1,结帐ID2,...
    '               其他，传入格式为:退费单据总张数|当前退第几张单据
    '出参:strAdvance:1.原样退回时，返回空
    '                2.退费结算方式与收费结算方式不一致时，返回格式为：结算方式|金额||结算方式|金额||…（其中，金额为负）
    strAdvance = mlng冲销ID & "|1"
    If Not gclsInsure.ClinicDelSwap(mlng结帐ID, True, mint险类, strAdvance) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    blnTransMedicare = True
    
    If Not InsureCheck(strAllAdvance, strAdvance) Then
        gcnOracle.CommitTrans
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mint险类)
        ExecuteReplenishDel = True: Exit Function
    End If
    gcnOracle.CommitTrans: gcnOracle.BeginTrans
    
    '退费和收费不一致时,需要效对
    Call SaveInsureBalance(mlng病人ID, mlng冲销ID, strAdvance, True)
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mint险类)
    ExecuteReplenishDel = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
    
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mint险类)
End Function

Private Function GetYBBalance(rsBalance As ADODB.Recordset, Optional ByVal strNos As String, _
    Optional ByVal blnDel As Boolean, Optional ByVal blnDelCheck As Boolean) As String
    '按单据获取医保原结算方式和结算金额
    '参数：
    '   strNOs - 单据号,多个用逗号隔开：A0001,A0002,...
    '   blnDel - 金额是否取反
    '   blnDelCheck - 是否检查允许门诊结算作废
    '返回:结算信息,格式:结算方式|结算金额||...
    Dim str结算方式 As String, varNos As Variant, strFilter As String
    Dim i As Integer, p As Integer
    Dim colBalance As Collection, strPreNo As String
    
    On Error GoTo errHandle
    If blnDelCheck And mint险类 = 0 Then Exit Function
    If rsBalance Is Nothing Then Exit Function
    
    '字段:类型,NO,结算方式,结算性质,冲预交,卡类别ID,结算卡序号,卡号,消费卡ID,交易流水号,交易说明,原结帐ID,关联交易ID
    If strNos = "" Then
        strFilter = "类型=" & gBalanceType.gEM_医保
    Else
        varNos = Split(strNos, ",")
        For i = 0 To UBound(varNos)
            strFilter = strFilter & " Or (类型=" & gBalanceType.gEM_医保 & " And No='" & varNos(i) & "')"
        Next
        If strFilter <> "" Then strFilter = Mid(strFilter, 5)
    End If
    rsBalance.Filter = strFilter
    If rsBalance.RecordCount = 0 Then Exit Function
    
    rsBalance.Sort = "No"
    Set colBalance = New Collection
    p = 1: colBalance.Add New BalanceMoneys
    With rsBalance
        strPreNo = NVL(!NO)
        Do While Not .EOF
            If strPreNo <> NVL(!NO) Then
                p = p + 1: colBalance.Add New BalanceMoneys
                strPreNo = NVL(!NO)
            End If
            If blnDelCheck Then
                '如果这种结算方式不支持回退,要退为现金,则不用减去
                If mbln门诊结算作废 Then
                    If gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mint险类, !结算方式) Then
                        str结算方式 = NVL(!结算方式) & "|" & IIf(blnDel, -1, 1) * Val(NVL(!冲预交))
                    End If
                Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                    If !结算方式 <> mstr个人帐户 Then
                        str结算方式 = NVL(!结算方式) & "|" & IIf(blnDel, -1, 1) * Val(NVL(!冲预交))
                    End If
                End If
            Else
                str结算方式 = NVL(!结算方式) & "|" & Val(NVL(!冲预交))
            End If
            
            Call SetBalanceVal(colBalance, p, str结算方式)
            .MoveNext
        Loop
    End With
    GetYBBalance = GetBalanceStr(colBalance)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InsureCheck(ByVal str预结算 As String, ByVal strAdvance As String) As Boolean
    '检查当前的医保是否需要较对
    '入参:
    '   str预结算-保险结算
    '   strAdvance-医保返回的结算
    '说明：
    '   正式结算前后,结算方式和结算金额未发生变化时不校对
    Dim blnFind  As Boolean, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo ErrHandler
    If strAdvance = "" Or str预结算 = strAdvance Then Exit Function
    
    InsureCheck = True
    
    varData = Split(str预结算, "||")
    varData1 = Split(strAdvance, "||")
    If UBound(varData) <> UBound(varData1) Then Exit Function
    
    For i = 0 To UBound(varData)
        blnFind = False
        varTemp = Split(varData(i), "|")
        For j = 0 To UBound(varData1)
            varTemp1 = Split(varData1(j), "|")
            If varTemp(0) = varTemp1(0) Then
                blnFind = True
                If Val(varTemp(1)) <> Val(varTemp1(1)) Then Exit Function
            End If
        Next
        If Not blnFind Then Exit Function
    Next
    InsureCheck = False
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Function GetBalanceStr(colBalance As Collection, Optional ByVal intPage As Integer, _
    Optional ByVal intPageCount As Integer) As String
    '功能：返回保险结算方式串,"结算方式|金额||...."
    '参数：intPage=是否指定单据,否则为所有单据
    '      intPageCount=结算单据总张数
    '说明：该函数以colBalance为准计算,对于医保划价收费也是
    Dim p As Integer
    Dim rsTemp As New ADODB.Recordset, strBalance As String
    Dim intPageStart As Integer, intPageEnd As Integer
    Dim objItem As BalanceMoney
    
    On Error GoTo ErrHander
    rsTemp.Fields.Append "结算方式", adVarChar, 50, adFldIsNullable
    rsTemp.Fields.Append "金额", adCurrency, , adFldIsNullable
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    intPageStart = IIf(intPage = 0, 1, intPage)
    intPageEnd = IIf(intPage = 0, IIf(intPageCount = 0, colBalance.Count, intPageCount), intPage)
    For p = intPageStart To intPageEnd
        For Each objItem In colBalance(p)
            rsTemp.Find "结算方式='" & objItem.结算方式 & "'", , adSearchForward, 1
            If rsTemp.EOF Then rsTemp.AddNew
            rsTemp!结算方式 = objItem.结算方式
            rsTemp!金额 = Val(NVL(rsTemp!金额)) + objItem.有效金额
            rsTemp.Update
        Next
    Next
    
    strBalance = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        strBalance = strBalance & "||" & NVL(rsTemp!结算方式) & "|" & NVL(rsTemp!金额)
        rsTemp.MoveNext
    Loop
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    
    GetBalanceStr = strBalance
    Exit Function
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetBalanceVal(colBalance As Collection, ByVal intPage As Integer, _
    ByVal strBalance As String)
    '功能：设置指定编号单据指定保险结算方式的有效值
    '参数：
    '       strBalance-根据结算方式字符串设置结算方式记录集，格式：结算方式1|金额1||结算方式2|金额2||...
    '说明：该函数以colBalance为准计算,对于医保划价收费也是
    '说明：用于正常医保收费修改保险结算金额；及划价单医保收费设置个人帐户等结算金额
    Dim i As Long
    Dim varBalance As Variant, varTemp As Variant
    Dim blnFind As Boolean
    Dim objItem As BalanceMoney, objBalanceMoneys As BalanceMoneys
    
    If strBalance = "" Then Exit Sub
    
    Set objBalanceMoneys = colBalance(intPage)
    
    '格式：结算方式1|金额1||结算方式2|金额2||...
    varBalance = Split(strBalance, "||")
    For i = 0 To UBound(varBalance)
        varTemp = Split(varBalance(i) & "|||", "|")
        blnFind = False
        For Each objItem In objBalanceMoneys
            If objItem.结算方式 = varTemp(0) Then
                objItem.有效金额 = varTemp(1)
                blnFind = True: Exit For
            End If
        Next
            
        If Not blnFind Then
            Set objItem = New BalanceMoney
            objItem.结算方式 = varTemp(0)
            objItem.原始金额 = varTemp(1)
            objItem.允许修改 = Val(varTemp(2)) = 1
            objItem.有效金额 = varTemp(1)
            objBalanceMoneys.AddItem objItem
        End If
    Next

    colBalance.Remove intPage '集合元素不能直接修改
    If colBalance.Count >= intPage Then
        colBalance.Add objBalanceMoneys, , intPage
    Else
        colBalance.Add objBalanceMoneys
    End If
End Sub

Private Sub SaveInsureBalanceDetail(ByVal lng结帐ID As Long, ByVal strNO As String, _
    ByVal strBalance As String)
    '保存医保结算明细
    Dim strSQL As String
    'Zl_医保结算明细_Insert(
    strSQL = "Zl_医保结算明细_Insert( "
    '  结帐id_In   医保结算明细.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '  No_In       医保结算明细.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  结算方式_In Varchar2,
    strSQL = strSQL & "'" & strBalance & "',"
    '  备注_In     医保结算明细.备注%Type := Null,
    strSQL = strSQL & "" & "NULL" & ")"
    zlDatabase.ExecuteProcedure strSQL, "保存医保结算明细"
End Sub

Private Sub SaveInsureBalance(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal strBalance As String, Optional ByVal bln补充结算 As Boolean)
    '保存医保结算数据
    Dim strSQL As String
    If bln补充结算 Then
        strSQL = "Zl_费用补充结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & 2 & ","
        '  结算id_In     In 费用补充记录.结算id%Type,
        strSQL = strSQL & "" & mlng冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & strBalance & "')"
    Else
        'Zl_门诊退费结算_Modify(
        strSQL = "Zl_门诊退费结算_Modify_S("
        '  操作类型_In      Number,
        strSQL = strSQL & "" & 3 & ","
        '  病人id_In        门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  姓名_In          病人预交记录.姓名%Type,
        strSQL = strSQL & "'" & mstr姓名 & "',"
        '  性别_In          病人预交记录.性别%Type,
        strSQL = strSQL & "'" & mstr性别 & "',"
        '  年龄_In          病人预交记录.年龄%Type,
        strSQL = strSQL & "'" & mstr年龄 & "',"
        '  门诊号_In        病人预交记录.门诊号%Type,
        strSQL = strSQL & "'" & mstr门诊号 & "',"
        '  住院号_In        病人预交记录.住院号%Type,
        strSQL = strSQL & "'" & mstr住院号 & "',"
        '  付款方式名称_In  病人预交记录.付款方式名称%Type,
        strSQL = strSQL & "'" & mstr医疗付款方式 & "',"
        '  冲销id_In        病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算方式_In      Varchar2
        strSQL = strSQL & "'" & strBalance & "')"
    End If
    zlDatabase.ExecuteProcedure strSQL, "保存医保结算数据"
End Sub

Private Sub BalanceOver(ByVal lng结帐ID As Long)
    '医保完成结算，更新校对标志
    Dim strSQL As String
    'Zl_病人门诊收费_医保更新(
    strSQL = "Zl_病人门诊收费_医保更新( "
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '  结算序号_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  保险结算_In Varchar2
    strSQL = strSQL & "" & "NULL" & ")"
    zlDatabase.ExecuteProcedure strSQL, "完成医保结算"
End Sub

Private Function CorrectInsureErrBalance(ByVal lng病人ID As Long, ByVal lng结帐ID As Long) As Boolean
    '多单据分单据结算，中途失败时校对结算数据
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim rsBalance As ADODB.Recordset, strBalance As String
    Dim rsBalanceSaved As ADODB.Recordset, strBalanceSaved As String
    
    On Error GoTo ErrHandler
    strSQL = "Select 1" & _
            " From 病人预交记录 A, 结算方式 B" & _
            " Where a.结算方式 = b.名称 And b.性质 In (3, 4) And 结帐id = [1] And a.卡类别ID Is Null " & _
            "       And Nvl(a.校对标志, 0) = 1 And Rownum < 2"
    strSQL = strSQL & "Union All" & _
            " Select 1" & _
            " From 保险结算记录" & _
            " Where 记录id = [1] " & _
            "       And Not Exists(Select 1 " & _
            "                      From 病人预交记录 A, 结算方式 B" & _
            "                      Where a.结算方式 = b.名称 And a.结帐id = 记录id " & _
            "                            And b.性质 In (3, 4) And a.卡类别ID Is Null)" & _
            "       And 卡类别ID Is Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存在需要校对的医保结算", lng结帐ID)
    If rsTemp.EOF Then CorrectInsureErrBalance = True: Exit Function
    
    '先通过“医保结算明细”进行校对
    Set rsBalance = GetBalanceDetail(0, lng结帐ID, 1)
    strBalance = GetYBBalance(rsBalance)
    
    If strBalance = "" Then
        strSQL = "Select a.结帐ID,a.结算方式,a.金额" & _
            " From 保险结算明细 A ,结算方式 C" & _
            " Where a.结算方式=c.名称 And c.性质 in (3,4) And a.结帐id =[1] And a.标志=1 " & _
            " Order by 结算方式"
        '医保管控的过程固定写入了一条"现金",所以排开非医保类的结算方式
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "保险结算管理", lng结帐ID)
        Do While Not rsTemp.EOF
            strBalance = strBalance & "||" & NVL(rsTemp!结算方式) & "|" & Val(NVL(rsTemp!金额))
            rsTemp.MoveNext
        Loop
        If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    End If
    '没有核对数据,直接返回
    If strBalance = "" Then CorrectInsureErrBalance = True: Exit Function
    
    '检查是否需要校对
    Set rsBalanceSaved = GetChargeBalance(0, lng结帐ID)
    strBalanceSaved = GetYBBalance(rsBalanceSaved, lng结帐ID)
    If InsureCheck(strBalanceSaved, strBalance) Then
        Call SaveInsureBalance(lng病人ID, lng结帐ID, strBalance)
    End If
    
    CorrectInsureErrBalance = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
