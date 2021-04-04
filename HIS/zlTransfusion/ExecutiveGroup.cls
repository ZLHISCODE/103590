VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExecutiveGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Group"
Attribute VB_Ext_KEY = "Member0" ,"Group"
Option Explicit

'局部变量，保存集合
Private mCol As Collection
'保持属性值的局部变量
Private mvar流水号 As Long '局部复制
'保持属性值的局部变量
Private mvar配药人 As String '局部复制
Private mvar执行时间 As Date '局部复制
Private mvar组数 As Integer '局部复制
Private mvar接单人 As String '局部复制
Private mvar滴系数 As Integer '局部复制
Private mvar总耗时 As Long
Private mvar执行分类 As String
 
Public Property Let 滴系数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.滴系数 = 5
    mvar滴系数 = vData
End Property

Public Property Get 滴系数() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.滴系数
    滴系数 = mvar滴系数
End Property

Public Property Let 总耗时(ByVal vData As Long)
'检索属性值时使用，位于赋值语句的右边。
' X.滴系数
    mvar总耗时 = vData
End Property

Public Property Get 总耗时() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.滴系数
    Dim objItem As Group
    mvar总耗时 = 0
    For Each objItem In mCol
     mvar总耗时 = mvar总耗时 + objItem.耗时
    Next
    总耗时 = mvar总耗时
End Property

Public Property Let 接单人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.接单人 = 5
    mvar接单人 = vData
End Property

Public Property Get 接单人() As String
'检索属性值时使用，位于赋值语句的右边。
' X.接单人
    接单人 = mvar接单人
End Property

Public Property Let 执行分类(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.接单人 = 5
    mvar执行分类 = vData
End Property

Public Property Get 执行分类() As String
'检索属性值时使用，位于赋值语句的右边。
' X.接单人
    执行分类 = mvar执行分类
End Property

Public Property Let 组数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.组数 = 5
    mvar组数 = vData
End Property

Public Property Get 组数() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.组数
    组数 = mvar组数
End Property

Public Property Let 执行时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行时间 = 5
    mvar执行时间 = vData
End Property

Public Property Get 执行时间() As Date
'检索属性值时使用，位于赋值语句的右边。
' X.执行时间
    执行时间 = mvar执行时间
End Property

Public Property Let 配药人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.配药人 = 5
    mvar配药人 = vData
End Property

Public Property Get 配药人() As String
'检索属性值时使用，位于赋值语句的右边。
' X.配药人
    配药人 = mvar配药人
End Property

Public Property Let 流水号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.流水号 = 5
    mvar流水号 = vData
End Property

Public Property Get 流水号() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.流水号
    流水号 = mvar流水号
End Property

Public Function Add(执行医嘱ID As Long, 发送号 As Long, 执行状态 As Integer, 执行人 As String, 用法 As String, 频率次数 As Integer, 滴速 As Integer, 执行频次 As String, 皮试结果 As String, 执行分类 As String, 已执行数次 As Integer, 发送数次 As Integer, 本次数次 As Integer, 组次 As Integer, BillsItem As Bills, 上次执行时间 As Date, 要求时间 As Date, 计费状态 As Integer, 记录性质 As Integer, 核对人 As String, Optional sKey As String) As Group
'创建新对象
    Dim objNewMember As Group, rsTmp As New ADODB.Recordset, strSQL As String
    Set objNewMember = New Group
    On Error GoTo hErr
    

    '设置传入方法的属性
    objNewMember.执行医嘱ID = 执行医嘱ID
    objNewMember.发送号 = 发送号
    objNewMember.执行状态 = 执行状态
    objNewMember.执行人 = 执行人
    objNewMember.用法 = 用法
    objNewMember.频率次数 = 频率次数
    objNewMember.滴速 = 滴速
    objNewMember.执行频次 = 执行频次
    objNewMember.皮试结果 = 皮试结果
    objNewMember.执行分类 = 执行分类
    objNewMember.已执行数次 = 已执行数次
    objNewMember.发送数次 = 发送数次
    objNewMember.本次数次 = 本次数次
    objNewMember.组次 = 组次
    objNewMember.计费状态 = 计费状态
    objNewMember.记录性质 = 记录性质
    objNewMember.核对人 = 核对人
    
    '-- 收费金额 2007-07-18
    objNewMember.收费金额 = 0
    '--   2011-09-21 添加执行状态<>9
    strSQL = "Select Sum(Nvl(B.实收金额, 0)) As 金额" & vbNewLine & _
            "From 门诊费用记录 B, 病人医嘱发送 A" & vbNewLine & _
            "Where A.医嘱id = B.医嘱序号 And A.NO = B.NO And B.记录状态<>0 And B.费用状态<>1 And a.发送号 = [2] And A.医嘱id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecutiveGroup.Add", 执行医嘱ID, 发送号)
    Do Until rsTmp.EOF
        objNewMember.收费金额 = IIf(IsNull(rsTmp.Fields("金额")), 0, rsTmp.Fields("金额"))
        rsTmp.MoveNext
    Loop
    
    If IsObject(BillsItem) Then
        Call objNewMember.AddBills(BillsItem, 执行医嘱ID & "_" & 发送号)
    End If
    objNewMember.上次执行时间 = 上次执行时间
    objNewMember.要求时间 = 要求时间
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Property Get Item(vntIndexKey As Variant) As Group
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    ' x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
    On Error Resume Next
End Sub

Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub

Public Sub Undo(ByVal lng流水号 As Long, ByVal lngExcDeptID As Long, Optional ByRef lngErrNo As Long)
    '撤消接单
    'lngExcDeptID :执行科室ID 47323
    Dim strSQL() As String, i As Integer, bln已完成 As Boolean
    Dim strBill() As String
    
    Dim objGroup As Group
    Dim intBills As Integer, intBill As Integer
    Dim objBIlls As Bills, objBIll As Bill, strTmp As String, rsTmp As New ADODB.Recordset, Y As Integer
    Dim blnTrans As Boolean
    On Error GoTo errHandle
    If mvar流水号 <> lng流水号 Then Exit Sub
    
    ReDim strSQL(mCol.Count)
    
    For Each objGroup In mCol
        If objGroup.执行状态 = 1 Then
        
            bln已完成 = True
            strSQL(i) = ""
        Else
            '流水号对应的医嘱，执行时间可能会改变
            strTmp = "Select 执行时间 From 病人医嘱执行 Where 医嘱ID=[1] And 发送号=[2] And 流水号=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "取医嘱执行时间", objGroup.执行医嘱ID, objGroup.发送号, lng流水号)
            If Not rsTmp.EOF Then mvar执行时间 = rsTmp!执行时间
            
            strSQL(i) = "ZL_病人医嘱执行_Delete(" & objGroup.执行医嘱ID & "," & objGroup.发送号 & ",to_date('" & Format(mvar执行时间, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),0,0," & lngExcDeptID & ")"
            '暂存药品撤消
            For intBills = 1 To objGroup.BillsCount
                Set objBIlls = objGroup.BillsItem(intBills)
                For intBill = 1 To objBIlls.Count
                    Set objBIll = objBIlls.Item(intBill)
                    strTmp = "Select NO,序号,to_char(登记时间,'yyyy-MM-dd HH24:MI:SS') as 登记时间 ,入出系数 From  暂存药品记录 Where 入出系数=-1 And 登记时间 = [3] And 医嘱ID=[1] And 发送号=[2] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "ExecutiveGroup.Undo", objBIll.医嘱ID, objGroup.发送号, mvar执行时间)
                    Do Until rsTmp.EOF
                        ReDim Preserve strBill(Y + 1)
                        
                        strBill(Y) = "zl_暂存药品记录_undouse('" & rsTmp.Fields("NO") & "'," & rsTmp.Fields("序号") & ",-1,TO_date('" & rsTmp.Fields("登记时间") & "','yyyy-MM-dd HH24:MI:SS'))"
                        rsTmp.MoveNext
                        Y = Y + 1
                    Loop
                Next
            Next
        End If
        i = i + 1
    Next
    LoadPicture
    If bln已完成 Then
        If MsgBox("待撤消的单据中含有已完成项目，这些项目不能撤消，是否继续？" & vbNewLine & "点[是]，只撤消未完成项目；点[否]，不执行任何操作。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    blnTrans = False
    gcnOracle.BeginTrans
    blnTrans = True
    For i = 0 To mCol.Count - 1
        If strSQL(i) <> "" Then
            If strSQL(i) <> "" Then
                Call zlDatabase.ExecuteProcedure(strSQL(i), "executivegroup.Undo")
            End If
        End If
    Next
    
    If Y > 0 Then
        For i = 0 To UBound(strBill) - 1
            If strBill(i) <> "" Then
                Call zlDatabase.ExecuteProcedure(strBill(i), "executivegroup.Undo")
            End If
        Next
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    Exit Sub
    
errHandle:
    Call ErrCenter
    If blnTrans = True Then gcnOracle.RollbackTrans
    Call SaveErrLog
    lngErrNo = Err.Number
End Sub

Public Function ExecuteFinish(ByVal strGroupKey As String, ByVal lngExcDeptID As Long, Optional strSQLTest As String = "", Optional frmMain As Object, Optional objCardSquare As Object) As Boolean
    '调用医嘱的执行完成功能
    ' strGroupKey: 定位一条医嘱的KEY
    ' frmMain    : 为调用一卡通接口传入的主窗体 对象
    'objCardSquare: 为调用一卡通接口传入的一卡通对象, objCardSquare 为Nothing 则表示部件未初始成功。
    'strSQLTest :因皮试的执行过程与其他项目相同，所以这个过程也可以由皮试完成过程调用,当是皮试过程调用时，此参数不为空。
    'lngExcDeptID : 执行科室ID，47323问题修改
    Dim strSQL As String, rsTmp As New ADODB.Recordset, strInfo As String

    Dim blnTrans As Boolean, intReturn As Integer, strPrivs As String
    Dim lngPatiId As Long
    '是否允许完成未收费病人的项目
    On Error GoTo errHandle
 
    '2012-07-17 取病人ID，后面用于取消座位占用
    
    strSQL = "Select 病人ID,诊疗类别 From 病人医嘱记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "医嘱执行完成", Item(strGroupKey).执行医嘱ID)
    If Not rsTmp.EOF Then lngPatiId = Val("" & rsTmp!病人ID)
    
    
    '调用一卡通处理函数
    intReturn = OneCardCheck(Item(strGroupKey).执行医嘱ID, Item(strGroupKey).发送号, frmMain, objCardSquare)
    
    If intReturn = 2 Then
        '新流程失败,内部有提示，此处直接退出
        Exit Function
    ElseIf intReturn = 0 Then
        '--老流程
        If Val(zlDatabase.GetPara("未收费完成", glngSys, 1264)) = 0 Then
                If Item(strGroupKey).收费金额 = 0 Then
                    '这里只检查收费记录，划价记录因为有执行后审核这种情况，前面已检查
                    
                    strSQL = "Select B.记录状态,A.计费状态,b.费用状态 ,D.收费细目id, a.NO,B.No as 收费NO, B.收据费目, D.医嘱内容, To_Char(A.发送时间, 'yyyy-MM-dd HH24:MI:SS') As 发送时间," & vbNewLine & _
                            "       Sum(Nvl(B.实收金额, 0)) As 金额" & vbNewLine & _
                            "From 病人医嘱记录 D, 门诊费用记录 B, 病人医嘱发送 A" & vbNewLine & _
                            "Where A.医嘱id = D.Id And A.医嘱id = B.医嘱序号(+) And A.No = B.No(+) And  nvl(B.记录状态(+),0)<=1 And A.发送号 = [2] And" & vbNewLine & _
                            "      mod(B.记录性质(+),10) = 1 And A.医嘱id = [1]" & vbNewLine & _
                            "Group By B.记录状态,A.计费状态, b.费用状态, D.收费细目id, A.No,B.NO, B.收据费目, D.医嘱内容, To_Char(A.发送时间, 'yyyy-MM-dd HH24:MI:SS')"
    
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecutiveGroup.ExecuteFinish", Item(strGroupKey).执行医嘱ID, Item(strGroupKey).发送号)
                    With rsTmp
                    Do Until .EOF
                        If !计费状态 = -1 Then
                            '无需计费的项目
                            
                            '2011-09-21 添加  (Val("" & !记录状态 <> 0) And Val("" & !执行状态 = 9))
                        ElseIf Val("" & !记录状态) = 0 Or (Val("" & !记录状态 <> 0) And Val(zlCommFun.NVL(!费用状态, 0) = 1)) Then
                            If strInfo = "" Then
                                strInfo = " NO: " & .Fields("NO") & " 发送时间:" & .Fields("发送时间") & vbNewLine
                            End If
                            strInfo = strInfo & "   医嘱:" & .Fields("医嘱内容")
                            If Trim("" & .Fields("收费细目id")) = "" Then
                                strInfo = strInfo & "   未对应费用明细!"
                            ElseIf Trim("" & .Fields("收费NO")) = "" Then
                                strInfo = strInfo & "   未生成费用明细!"
                            Else
                                strInfo = strInfo & " 金额:" & Format(.Fields("金额"), "0.00") & vbNewLine
                            End If
                        End If
                        .MoveNext
                    Loop
                    End With
                    If strInfo <> "" Then
                        MsgBox "该项目还未收费，不能完成。" & vbNewLine & strInfo, vbInformation, gstrSysName
                        Exit Function
                    End If
            End If
        End If  '模块内权限，“未收费完成”
    End If  '是老流程检查
    
    blnTrans = False
    gcnOracle.BeginTrans
    blnTrans = True
    If strSQLTest <> "" Then
        Call zlDatabase.ExecuteProcedure(strSQLTest, "ExecutiveGroup.ExecuteFinish")
    End If

    '2012-07-17 51194问题 此过程 已有自动发卫材功能 ，系统参数[临床应用]－[执行之后对卫生材料自动发料]需要勾上才生效。
    strSQL = "ZL_病人医嘱执行_Finish(" & Item(strGroupKey).执行医嘱ID & "," & Item(strGroupKey).发送号 & ",Null,0,Null,Null," & lngExcDeptID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.ExecuteFinish")
    gcnOracle.CommitTrans
    blnTrans = False
        
    ExecuteFinish = True
    Exit Function
errHandle:
    
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    
    Call SaveErrLog
End Function

Public Sub Update(ByVal str流水号 As String, ByVal strGroupKey As String, ByVal lngExcDeptID As Long, Optional ByRef lngErrNo As Long)
    '保存修改内容
    'lngExcDeptID :执行科室ID 47323问题
    
    Dim objGroup As Group, strSQL As String, rsTmp As ADODB.Recordset
    Dim date原执行时间 As Date, str原摘要 As String, blnBegin As Boolean
    On Error GoTo errHandle
    
    strSQL = "Select 执行时间,执行摘要 From 病人医嘱执行 Where 流水号=[1] and 医嘱ID=[2] and 发送号=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecutiveGroup.Update", 流水号, Item(strGroupKey).执行医嘱ID, Item(strGroupKey).发送号)
    Do Until rsTmp.EOF
        date原执行时间 = zlCommFun.NVL(rsTmp!执行时间, CDate(0))
        str原摘要 = zlCommFun.NVL(rsTmp!执行摘要)
        rsTmp.MoveNext
    Loop
    
    strSQL = "ZL_病人医嘱执行_Update(to_date('" & Format(date原执行时间, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                            Item(strGroupKey).执行医嘱ID & "," & Item(strGroupKey).发送号 & "," & _
                                            "to_date('" & Format(Item(strGroupKey).要求时间, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                            Item(strGroupKey).本次数次 & ",'" & str原摘要 & "','" & Item(strGroupKey).执行人 & "'," & _
                                            "to_date('" & Format(执行时间, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,Null,0,Null,Null," & lngExcDeptID & ")"
    
    gcnOracle.BeginTrans
    blnBegin = True
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.Update")
    strSQL = "Zl_病人医嘱执行_Modify(" & 流水号 & "," & Item(strGroupKey).执行医嘱ID & "," & Item(strGroupKey).发送号 & "," & _
                                Item(strGroupKey).滴速 & "," & Item(strGroupKey).液体量 & "," & 滴系数 & "," & Item(strGroupKey).耗时 & ",'" & _
                                Item(strGroupKey).说明 & "','" & _
                                UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.Update")
    gcnOracle.CommitTrans
    blnBegin = False
    Exit Sub
errHandle:
    lngErrNo = Err.Number
    Call ErrCenter
    If blnBegin Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Public Function ExecuteTestFinish(ByVal strGroupKey As String, ByVal frmMain As Object, ByVal objCardSquare As Object, ByRef str皮试结果 As String) As Boolean
    '皮试完成
    Dim strSQL As String, strTest As String, strOwner As String
    Dim str结果 As String, lngDeptID As Long
    Dim strContent As String, strLabel As String
    Dim strSelect As String, strSelectInput As String, strTextInput As String
    Dim datInput As Date
    Dim objBIll As Bill
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer, intResult  As Integer
    
    For Each objBIll In Item(strGroupKey).BillsItem(strGroupKey)
        If objBIll.医嘱ID = Val(strGroupKey) Then
            strContent = objBIll.医嘱内容
            Exit For
        End If
    Next
    
    On Error GoTo errHandle
    strSQL = "Select A.诊疗类别,A.皮试结果,B.操作类型,Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 " & _
             "From 病人医嘱记录 A, 诊疗项目目录 B " & _
             "Where A.诊疗项目ID=B.ID And A.ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "皮试内容", Val(strGroupKey))
    '阳性
    For i = 0 To UBound(Split(Split(rsTemp!标本部位 & "", ";")(0), ","))
        strSelect = strSelect & "," & Split(Split(rsTemp!标本部位 & "", ";")(0), ",")(i) & "|0"
    Next
    '阴性
    For i = 0 To UBound(Split(Split(rsTemp!标本部位 & "", ";")(1), ","))
        strSelect = strSelect & "," & Split(Split(rsTemp!标本部位 & "", ";")(1), ",")(i) & "|0|2"
    Next
    strSelect = Mid(strSelect, 2)
    
    str结果 = zlCommFun.ShowMsgBox("皮试结果", _
                  strContent & "：^^请根据过敏试验结果选择相应的按钮操作。", _
                  "确定(&O),?取消(&C)", frmMain, vbQuestion, _
                  "皮试时间", datInput, "yyyy-MM-dd HH:mm", _
                  "皮试结果(&P):" & strSelect, strSelectInput, _
                  "过敏反应(&F)", 100, strTextInput)
    
    If str结果 = "" Then Exit Function
    If strSelectInput = "" Then Exit Function
    
    GetTestLabel rsTemp!标本部位, strSelectInput, strLabel, intResult
    lngDeptID = frmMain.cboDept.ItemData(frmMain.cboDept.ListIndex)
    
    strTest = "ZL_病人医嘱记录_皮试(" & _
              Item(strGroupKey).执行医嘱ID & _
              ",'" & strLabel & "'" & _
              "," & intResult & _
              ",''" & _
              ",to_date('" & datInput & "','YYYY-MM-DD HH24:MI:SS')" & _
              ",'" & strTextInput & "') "

    ExecuteTestFinish = ExecuteFinish(strGroupKey, lngDeptID, strTest, frmMain, objCardSquare)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function ExecCanle(ByVal strGroupKey As String, ByVal bln皮试验证 As Boolean, ByVal lngExcDeptID As Long, ByVal objMain As Object) As Boolean
    '取消完成,皮试项目也调此项
    '
    '门诊不用检查是否可以取消。
    'lngExcDeptID :执行科室ID 47323
    
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strUserName As String
    Dim cnNew As ADODB.Connection
    
    On Error GoTo errHandle
    If Item(strGroupKey).执行状态 <> 1 Then
        MsgBox "该执行项目当前不处于已执行状态，不能取消执行。", vbInformation, gstrSysName
        Exit Function
    End If
    '----- 这里简单判断，如要杜绝，需要zl_病人医嘱执行_Cancel中处理
    If Item(strGroupKey).计费状态 = -3 Then
        MsgBox "已退费，不能执行取消执行"
        Exit Function
    End If
    
    If Val(Item(strGroupKey).执行分类) = 3 Then
        '皮试类
        If bln皮试验证 Then
            Set cnNew = New ADODB.Connection
            strUserName = zlDatabase.UserIdentify(objMain, "撤销皮试结果前，请您先输入用户名和密码进行身份验证。", glngSys, 1263, "取消执行完成", cnNew)
            If strUserName = "" Then Exit Function
        End If
        strSQL = "ZL_病人医嘱执行_Cancel(" & Item(strGroupKey).执行医嘱ID & "," & Item(strGroupKey).发送号 & "," & IIf(bln皮试验证, 1, 0) & ",0," & lngExcDeptID & ")"
    Else
        strSQL = "ZL_病人医嘱执行_Cancel(" & Item(strGroupKey).执行医嘱ID & "," & Item(strGroupKey).发送号 & ",Null,0," & lngExcDeptID & ")"
    End If
    
    'strSQL = "ZL_病人医嘱执行_Cancel(" & Item(strGroupKey).执行医嘱ID & "," & Item(strGroupKey).发送号 & "," & IIf(bln皮试验证, 1, 0) & ",0," & lngExcDeptID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.ExecCanle")
    strSQL = "Select 执行状态 From 病人医嘱发送 Where 发送号+0=[1] and 医嘱ID=[2] And Rownum=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecutiveGroup.ExecCanle", Item(strGroupKey).发送号, Item(strGroupKey).执行医嘱ID)
    Do Until rsTmp.EOF
        Item(strGroupKey).执行状态 = Val(rsTmp.Fields("执行状态"))
        Item(strGroupKey).皮试结果 = ""
        rsTmp.MoveNext
    Loop
    ExecCanle = True
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Public Sub FuncExecRefuse(strGroupKey As String)
''功能：拒绝执行
'    Dim lng医嘱ID As Long, lng发送号 As Long
'    Dim strSQL As String
'
'    '正在执行或已执行不允许拒绝
'    If Item(strGroupKey).执行状态 = 2 Then
'        MsgBox "该执行项目当前已经拒绝执行。", vbInformation, gstrSysName
'        Exit Sub
'    End If
'    If Item(strGroupKey).执行状态 = 3 Then
'        MsgBox "该执行项目当前正在执行，不能拒绝。", vbInformation, gstrSysName
'        Exit Sub
'    End If
'    If Item(strGroupKey).执行状态 = 1 Then
'        MsgBox "该执行项目当前已经执行，不能拒绝。", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    If MsgBox("确实要拒绝执行该项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    lng医嘱ID = Item(strGroupKey).执行医嘱ID
'    lng发送号 = Item(strGroupKey).发送号
'
'    strSQL = "ZL_病人医嘱执行_拒绝执行(" & lng医嘱ID & "," & lng发送号 & ")"
'
'    On Error GoTo errH
'
'    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.FuncExecRefuse")
'
'    On Error GoTo 0
'    Item(strGroupKey).执行状态 = 2
'    Exit Sub
'errH:
'
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

'Public Sub FuncExecRestore(strGroupKey As String)
''功能：取消拒绝执行
'    Dim lng医嘱ID As Long, lng发送号 As Long
'    Dim strSQL As String, i As Long
'
'    If Item(strGroupKey).执行状态 <> 2 Then
'        MsgBox "该执行项目没有被拒绝执行。", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    If MsgBox("确实要取消拒绝执行该项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    lng医嘱ID = Item(strGroupKey).执行医嘱ID
'    lng发送号 = Item(strGroupKey).发送号
'
'    strSQL = "ZL_病人医嘱执行_取消拒绝(" & lng医嘱ID & "," & lng发送号 & ")"
'
'    On Error GoTo errH
'
'    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.FuncExecRestore")
'
'    On Error GoTo 0
'    Item(strGroupKey).执行状态 = 0
'    Exit Sub
'errH:
'
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Public Sub ExecStart(ByVal intUndo As Integer, strGroupKey As String, dateStart As Date, strOper As String)
    '填写开始时间
    Dim strSQL As String
    Dim lngID As Long, lngSendNo As Long, lngSeqNo As Long
    
    On Error GoTo hErr
    lngID = Item(strGroupKey).执行医嘱ID
    lngSendNo = Item(strGroupKey).发送号
    lngSeqNo = Me.流水号
 
    If intUndo = 1 Then
        If Me.执行时间 <> dateStart Then Me.执行时间 = dateStart
        strSQL = "Zl_病人医嘱执行_Start(1," & lngSeqNo & "," & lngID & "," & lngSendNo & ",to_date('" & Format(dateStart, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),'" & strOper & "')"
    Else
        strSQL = "Zl_病人医嘱执行_Start(2," & lngSeqNo & "," & lngID & "," & lngSendNo & ")"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, "填写开始时间")
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub SaveDispenseUser(ByVal intUndo As Integer, dateStart As Date, strOper As String)
    '填写配药人，配药时间
    Dim strSQL As String
    Dim lngID As Long, lngSendNo As Long, lngSeqNo As Long
    Dim Y As Integer
    On Error GoTo hErr
    
    For Y = 1 To Me.Count
        If Item(Y).执行分类 = "1-输液" Then
            lngID = Item(Y).执行医嘱ID
            lngSendNo = Item(Y).发送号
            lngSeqNo = Me.流水号
        
            If intUndo = 1 Then
                If Me.执行时间 <> dateStart Then Me.执行时间 = dateStart
                strSQL = "Zl_病人医嘱执行_Start(3," & lngSeqNo & "," & lngID & "," & lngSendNo & ",to_date('" & Format(dateStart, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),'" & strOper & "')"
            Else
                strSQL = "Zl_病人医嘱执行_Start(4," & lngSeqNo & "," & lngID & "," & lngSendNo & ")"
            End If
            Call zlDatabase.ExecuteProcedure(strSQL, "填写开始时间")
       End If
    Next
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
