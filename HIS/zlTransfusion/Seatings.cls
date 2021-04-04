VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Seatings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Seating"
Attribute VB_Ext_KEY = "Member0" ,"Seating"
Option Explicit

'局部变量，保存集合
Private mCol As Collection
Private mlng科室ID As Long
Private mstr科室 As String

Public Property Get 科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.编号
    科室ID = mlng科室ID
End Property

Public Property Get 科室名称() As String
'检索属性值时使用，位于赋值语句的右边。
' X.编号
    科室名称 = mstr科室
End Property

Public Function Add(序号 As Integer, 病人ID As Long, 门诊号 As String, 姓名 As String, 性别 As String, 编号 As String, _
                    类别 As Integer, 状态 As Integer, 现价 As Currency, 收费细目ID As Long, 收费项目 As String, _
                    备注 As String, 类型 As Long, 分类 As String, 呼叫器 As String, 挂号单 As String, Optional sKey As String) As Seating
    '创建新对象
    Dim objNewMember As Seating
    On Error GoTo errHandle
    If 序号 = 0 Then
        Dim strSQL As String
        '新增的数据，要写到库中
        strSQL = "ZL_座位状况记录_INSERT(" & mlng科室ID & ",'" & 编号 & "'," & Val(类别) & "," & Val(状态) & "," & IIf(Val(收费细目ID) = 0, "Null", Val(收费细目ID)) & ",'" & 备注 & "'," & 类型 & ",'" & 分类 & "','" & 呼叫器 & "')"
        zldatabase.ExecuteProcedure strSQL, "seatings.add"
        Call getSeatings(mlng科室ID)
    Else
        Set objNewMember = New Seating
        '设置传入方法的属性
        objNewMember.序号 = 序号
        objNewMember.病人ID = 病人ID
        objNewMember.门诊号 = 门诊号
        objNewMember.姓名 = 姓名
        objNewMember.性别 = 性别
        objNewMember.编号 = 编号
        objNewMember.类别 = 类别
        objNewMember.状态 = 状态
        objNewMember.现价 = 现价
        objNewMember.收费细目ID = 收费细目ID
        objNewMember.收费项目 = IIf(IsNull(收费项目), "", 收费项目)
        objNewMember.备注 = IIf(IsNull(备注), "", 备注)
        objNewMember.类型 = 类型
        objNewMember.呼叫器编号 = 呼叫器
        objNewMember.分类 = 分类
        objNewMember.挂号单 = 挂号单
        If 挂号单 = "" Then
            objNewMember.PatientKey = "1_" & 病人ID & "_" & GetPatiPageID(mlng科室ID, 病人ID)
        Else
            objNewMember.PatientKey = "0_" & 病人ID & "_" & 挂号单
        End If
        objNewMember.Key = sKey
        
        If Len(sKey) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sKey
        End If
        '返回已创建的对象
        Set Add = objNewMember
        Set objNewMember = Nothing
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetNextNo(ByVal 类别 As Integer) As String
    
    Dim strSQL As String, strMaxNO As String, str前缀 As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select Max(编号) As maxno From 座位状况记录 Where  科室ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.GetNextNO", mlng科室ID)
'    Select Case 类别
'        Case 1
'            str前缀 = "加"
'        Case 2
'            str前缀 = "特"
'        Case 3
'            str前缀 = "贵"
'        Case Else
'            str前缀 = "普"
'    End Select
    str前缀 = ""
    If rsTmp.RecordCount > 0 Then
        strMaxNO = zlCommFun.NVL(rsTmp.Fields("maxno"))
        If strMaxNO = "" Then
            GetNextNo = str前缀 & "001"
        Else
            GetNextNo = zlCommFun.IncStr(strMaxNO)
        End If
    Else
        GetNextNo = str前缀 & "001"
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ClearSeatings()
    '清除集合中数据
    Dim lngIndex As Long
    For lngIndex = 1 To Me.Count
        Remove 1
    Next
End Function

Private Function GetPatiPageID(ByVal lngDeptID As Long, ByVal lngPatiID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Max(主页id) 主页id From 排队记录 Where 科室id = [1] And 病人id = [2]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "获取病人的主页ID", lngDeptID, lngPatiID)
    If rsTemp.EOF = False Then
        GetPatiPageID = Val(zlCommFun.NVL(rsTemp!主页id))
    End If
    rsTemp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Function SetSeating(ByVal lng病人ID As Long, ByVal strRegNo As String, ByVal SeatingKey As String) As Boolean
    '安排座位
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strBillNO As String
    Dim blnTrans As Boolean
    Dim lngPageID As Long
    
    On Error GoTo errHandle
    
    If lng病人ID = 0 Or SeatingKey = "" Then
        MsgBox "座位安排的入参错误，请检查!", vbInformation, gstrSysName
        Exit Function
    End If
    '-- 更新座位信息
    Call Me.getSeatings(Me.科室ID)
    SeatingKey = Replace(SeatingKey, "-", "_")
    If Me.Item(SeatingKey).病人ID = 0 And Me.Item(SeatingKey).状态 = 0 Then
        Me.Item(SeatingKey).病人ID = lng病人ID
        Me.Item(SeatingKey).状态 = 1
        Me.Item(SeatingKey).挂号单 = strRegNo
        
        If (strRegNo = "" Or strRegNo Like "*无*") Then
            '从排队记录取门诊留观病人的主页ID
            lngPageID = GetPatiPageID(Me.科室ID, lng病人ID)
            If lngPageID > 0 Then
                Me.Item(SeatingKey).PatientKey = "1_" & lng病人ID & "_" & lngPageID
            Else
                Me.Item(SeatingKey).PatientKey = "0_" & lng病人ID & "_"
            End If
        Else
            Me.Item(SeatingKey).PatientKey = "0_" & lng病人ID & "_" & strRegNo
        End If
        
        strSQL = "Select 病人ID From 座位状况记录 Where 科室ID=[1] and 类别=[2] and 编号=[3]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "SetSeating", Me.科室ID, Me.Item(SeatingKey).类别, Me.Item(SeatingKey).编号)
        If Not rsTmp.EOF Then
            If zlCommFun.NVL(rsTmp.Fields("病人ID"), 0) = 0 Then
                
                If Me.Item(SeatingKey).收费细目ID <> 0 Then
                    '产生划价单
                    strBillNO = zldatabase.NextNo(13)
                    With Me.Item(SeatingKey)
                        Dim str付款方式 As String, str年龄 As String, str费别 As String, lng病区ID As Long, lng科室ID As Long
                        Dim rsBill As ADODB.Recordset
                        Dim str姓名 As String, str性别 As String
                        strSQL = "Select B.编码, A.医疗付款方式, A.年龄, A.费别, A.当前科室id, A.当前病区id, A.姓名, A.性别" & vbNewLine & _
                                "From 医疗付款方式 B, 病人信息 A" & vbNewLine & _
                                "Where A.医疗付款方式 = B.名称 And A.病人ID=[1]"

                        Set rsBill = zldatabase.OpenSQLRecord(strSQL, "SetSeating", lng病人ID)
                        If Not rsBill.EOF Then
                            str付款方式 = rsBill.Fields("编码")
                            str姓名 = rsBill.Fields("姓名")
                            str性别 = rsBill.Fields("性别")
                            str年龄 = zlCommFun.NVL(rsBill.Fields("年龄"))
                            str费别 = zlCommFun.NVL(rsBill.Fields("费别"))
                            lng病区ID = zlCommFun.NVL(rsBill.Fields("当前病区ID"), 0)
                            lng科室ID = zlCommFun.NVL(rsBill.Fields("当前科室ID"), 0)
                        End If
                        
                        Dim str收费类别 As String, str计算单位 As String, lng收入项目ID As Long, str收据费目 As String, lng价格父号 As Long, lng加班标志 As Long
                        Dim cur单价 As Currency, cur金额 As Currency
                        Dim int序号 As Integer
                        strSQL = "Select a.类别, a.计算单位, a.加班加价, b.现价, b.收入项目id, c.收据费目" & vbNewLine & _
                                "From 收入项目 c," & vbNewLine & _
                                "        (Select 收入项目id, 现价, 收费细目id" & vbNewLine & _
                                "            From 收费价目 a" & vbNewLine & _
                                "            Where 终止日期 Is Null Or 终止日期 = To_Date('3000-01-01', 'yyyy-MM-dd')) b, 收费项目目录 a" & vbNewLine & _
                                "Where b.收入项目id = c.Id And a.Id = b.收费细目id And a.Id = [1]"
                        Set rsBill = zldatabase.OpenSQLRecord(strSQL, "SetSeating", .收费细目ID)
                        
                        gcnOracle.BeginTrans
                        blnTrans = True
                        Do Until rsBill.EOF
                            int序号 = int序号 + 1
                            If int序号 = 1 Then
                                lng价格父号 = rsBill.Fields("收入项目ID")
                            End If
                            str收费类别 = rsBill.Fields("类别")
                            str计算单位 = rsBill.Fields("计算单位")
                            cur单价 = rsBill.Fields("现价")
                            lng收入项目ID = rsBill.Fields("收入项目id")
                            str收据费目 = rsBill.Fields("收据费目")
                            
                            strSQL = "Zl_门诊划价记录_Insert('" & strBillNO & "',1," & lng病人ID & ",Null," & .门诊号 & ",'" & str付款方式 & "','" & _
                                     str姓名 & "','" & str性别 & "','" & str年龄 & "','" & _
                                     str费别 & "'," & lng加班标志 & "," & IIf(lng科室ID = 0, "Null", lng科室ID) & "," & Me.科室ID & ",'" & _
                                     UserInfo.姓名 & "',Null," & .收费细目ID & ",'" & str收费类别 & "','" & str计算单位 & "',Null," & _
                                     "Null,1,Null," & Me.科室ID & "," & IIf(int序号 = 1, "Null", lng价格父号) & "," & lng收入项目ID & ",'" & str收据费目 & "'," & cur单价 & "," & _
                                     cur单价 & "," & cur单价 & ",sysdate,sysdate,Null,'" & UserInfo.姓名 & "')"
                            Call zldatabase.ExecuteProcedure(strSQL, "SetSeating")
                            cur金额 = cur金额 + cur单价
                            rsBill.MoveNext
                        Loop
                        gcnOracle.CommitTrans
                        blnTrans = False
                    End With
                    MsgBox "请提醒[" & str姓名 & "]到收费室去缴费。" & vbNewLine & "单据号：" & strBillNO & "，金额：" & Format(cur金额, "0.00"), vbQuestion, gstrSysName
                End If
                strSQL = "Zl_座位状况记录_Setseating(" & Me.科室ID & "," & Me.Item(SeatingKey).类别 & ",'" & Me.Item(SeatingKey).编号 & "'," & lng病人ID & ",'" & strBillNO & "')"
                Call zldatabase.ExecuteProcedure(strSQL, "Seatings.SetSeating")
                
                SetSeating = True
                Call SaveOperLog(Me.科室ID, Me.Item(SeatingKey).PatientKey, SEAT, "安排坐位" & Me.Item(SeatingKey).编号 & IIf(strBillNO <> "", "产生划价单" & strBillNO & "金额：" & Format(cur金额, "0.00"), ""))
            ElseIf zlCommFun.NVL(rsTmp.Fields("病人ID"), 0) = lng病人ID Then
                SetSeating = True
            Else
                MsgBox "座位已经被其他用户先占用，请另选座位！", vbInformation, gstrSysName
                Exit Function

            End If
        Else
            MsgBox "座位已经不存在，请重新安排！", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        MsgBox "座位已经被其他人占用，请另选座位！", vbInformation, gstrSysName
        Exit Function
    End If
    Exit Function
errHandle:
    
    Call ErrCenter
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Sub getSeatings(ByVal deptID As Long)
    '取座位信息
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsPati As ADODB.Recordset
    Dim str挂号单 As String
    
    On Error GoTo errHandle
    Call ClearSeatings
    
    strSQL = "Select Rownum As 序号, Nvl(A.病人id, 0) As 病人id, Nvl(B.门诊号, 0) 门诊号, Nvl(B.姓名, '') As 姓名, " & vbNewLine & _
            "       Nvl(B.性别, '') 性别, B.年龄 ,A.编号, A.类别, Nvl(A.收费细目id, 0) 收费细目id, A.状态, A.备注, Nvl(A.类型,0) as 类型, Nvl(C.现价, 0) As 现价, a.分类, a.呼叫器编号," & vbNewLine & _
            "       Replace('[' || D.编码 || ']' || D.名称, '[]', '') As 收费项目" & vbNewLine & _
            "From 收费项目目录 D, 收费价目 C, 病人信息 B, 座位状况记录 A " & vbNewLine & _
            "Where A.收费细目id = D.ID(+) And (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "      A.收费细目id = C.收费细目id(+) And A.病人id = B.病人id(+) And A.科室id = [1] " & _
            GetPriceGradeSQL(gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级, "D", "C", "2", "3", "4") & _
            "Order by A.类别,A.编号"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Seatings.getSeatings", deptID, gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级)
    With rsTmp
    Do Until .EOF
        str挂号单 = ""
        If Val(zlCommFun.NVL(!病人ID)) > 0 Then
            strSQL = "Select NO From 病人挂号记录 " & _
                     "Where (病人id, ID) In (Select 病人id, Max(ID) ID From 病人挂号记录 " & _
                     "                       Where 记录性质 = 1 And 记录状态 = 1 And 病人id = [1] And 执行部门id = [2] " & _
                     "                       Group By 病人id)"
            Set rsPati = zldatabase.OpenSQLRecord(strSQL, "获取病人挂号单", !病人ID, deptID)
            If rsPati.EOF = False Then
                str挂号单 = zlCommFun.NVL(rsPati!NO)
            End If
            rsPati.Close
        End If
    
        Add .Fields("序号"), .Fields("病人ID"), .Fields("门诊号"), _
            Trim$("" & .Fields("姓名") & " " & .Fields("年龄")), "" & .Fields("性别"), .Fields("编号"), _
            "" & .Fields("类别"), .Fields("状态"), .Fields("现价"), .Fields("收费细目ID"), _
            IIf(IsNull(.Fields("收费项目")), "", .Fields("收费项目")), _
             IIf(IsNull(.Fields("备注")), "", .Fields("备注")), _
             .Fields("类型"), "" & .Fields("分类"), "" & .Fields("呼叫器编号"), _
             str挂号单, "" & .Fields("类别") & "_" & .Fields("编号")
        .MoveNext
    Loop
    End With
    '取科室名称
    mlng科室ID = deptID
    strSQL = "Select 编码,名称 From 部门表 Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.getSeatings", deptID)
    If rsTmp.RecordCount > 0 Then
        mstr科室 = "[" & rsTmp.Fields("编码") & "] " & rsTmp.Fields("名称")
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As Seating
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
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
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub

Public Function Delete(ByVal strKey As String) As Boolean
    '删除座位
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim objSeating As Seating
    On Error GoTo errHandle
        

    
    Set objSeating = Item(strKey)
    
    strSQL = "Select A.病人ID,B.姓名,A.状态  From 座位状况记录 A,病人信息 B Where  A.病人ID=B.病人ID(+) And A.科室ID=[1] And A.编号=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.Delete", mlng科室ID, CStr(objSeating.编号))
    If rsTmp.EOF Then
        Call Remove(strKey)
        Delete = True
        Exit Function
    Else
        If rsTmp!病人ID <> 0 Or rsTmp!状态 = 1 Then
            MsgBox objSeating.编号 & "被" & rsTmp!姓名 & "占用，不能删除!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    
    If strKey <> "" And objSeating.病人ID = 0 And objSeating.状态 <> 1 Then
        If MsgBox("是否删除[" & objSeating.编号 & "]座位？", vbExclamation + vbYesNo, "提醒") = vbYes Then
            strSQL = "Zl_座位状况记录_Delete(" & mlng科室ID & ",'" & objSeating.编号 & "')"
            Call zldatabase.ExecuteProcedure(strSQL, "seatings.delete")
            Call Remove(strKey)
            Delete = True
        End If
    Else
        If strKey = "" Then
            MsgBox "编号为空的座位，不能删除!", vbInformation, gstrSysName
            Exit Function
        End If
        If objSeating.病人ID <> 0 Or objSeating.状态 = 1 Then
            MsgBox objSeating.编号 & "被" & objSeating.姓名 & "占用，不能删除!", vbInformation, gstrSysName
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Clear(ByVal strKey As String, Optional blnInfo As Boolean = True) As Boolean
    '清除座位
    'blnInfo:是否提醒
    Dim strSQL As String, rsTmp As New ADODB.Recordset, strNo As String
    Dim blnTrans As Boolean, strSwap As String
    On Error GoTo errHandle
    If strKey <> "" And Item(strKey).病人ID <> 0 And Item(strKey).状态 = 1 Then
        
        If blnInfo Then
            If MsgBox("是否清除被[" & Item(strKey).姓名 & "]占用的[" & Item(strKey).编号 & "]座位？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            strSwap = "清除"
        Else
            '换座位时,检查是否有划价单,如果未收费则清除,已收费则提示去退费
            strSQL = "Select 病人ID,NO From 座位状况记录 Where 科室ID=[1] and 类别=[2] and 编号=[3]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.Clear", Me.科室ID, Me.Item(strKey).类别, Me.Item(strKey).编号)
            If Not rsTmp.EOF Then
                strNo = zlCommFun.NVL(rsTmp.Fields("NO"))
                If strNo <> "" Then
                    strSQL = "Select nvl(实收金额,0) as 金额 From 门诊费用记录 Where mod(记录性质,10)=1 And 记录状态=1 and 费用状态<>1 And NO=[1] "
                    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.Clear", strNo)
                    If Not rsTmp.EOF Then
                        MsgBox "请提醒[" & Item(strKey).姓名 & "]到收费室退费" & vbNewLine & "单据号：" & strNo & "，金额：" & Format(Item(strKey).现价, "0.00") & "。", vbInformation, gstrSysName
                    Else
                        strSQL = "Select NO,序号 From 门诊费用记录 Where mod(记录性质,10)=1 And 记录状态=0 And NO=[1] "
                        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "Seatings.Clear", strNo)
                        gcnOracle.BeginTrans
                        blnTrans = True
                        Do Until rsTmp.EOF
                            strSQL = "zl_门诊划价记录_Delete('" & strNo & "'," & rsTmp.Fields("序号") & ")"
                            Call zldatabase.ExecuteProcedure(strSQL, "Seatings.Clear")
                            rsTmp.MoveNext
                        Loop
                        gcnOracle.CommitTrans
                        blnTrans = False
                    End If
                End If
            End If
            strSwap = "换座位－清除"
        End If
        
        strSQL = "ZL_座位状况记录_Clear(" & mlng科室ID & ",'" & Item(strKey).编号 & "')"
        Call zldatabase.ExecuteProcedure(strSQL, "Seatings.Clear")
        
        Clear = True
        Call SaveOperLog(Me.科室ID, Item(strKey).PatientKey, SEAT, strSwap & IIf(Item(strKey).姓名 <> "", "被[" & Item(strKey).姓名 & "]占用的", "") & "[" & Item(strKey).编号 & "]座位")
    
        Item(strKey).病人ID = 0
        Item(strKey).姓名 = ""
        Item(strKey).性别 = ""
        Item(strKey).状态 = 0
        Item(strKey).门诊号 = 0
        Item(strKey).诊断 = ""
        Item(strKey).开始时间 = ""
        Item(strKey).挂号单 = ""
        Item(strKey).PatientKey = ""
        
    End If
    Exit Function
errHandle:
    
    Call ErrCenter
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function SwapSeating(strSourceKey, strObjKey) As Boolean
    '
    Dim lng病人ID As Long, strRegNo As String
    On Error GoTo errHandle
    
    If strSourceKey <> "" And strObjKey <> "" Then
        lng病人ID = Item(strSourceKey).病人ID
        strRegNo = Item(strSourceKey).挂号单
        
        If Not Clear(strSourceKey, False) Then Exit Function
        If Not SetSeating(lng病人ID, strRegNo, strObjKey) Then Exit Function
        
        SwapSeating = True
    End If
    
    Exit Function
errHandle:

    Call ErrCenter
End Function


