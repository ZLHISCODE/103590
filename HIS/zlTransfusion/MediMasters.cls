VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediMasters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MediMaster"
Attribute VB_Ext_KEY = "Member0" ,"MediMaster"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'局部变量，保存集合
Private mCol As Collection

Private Function AddMaster(ByVal objMaster As MediMaster, Optional sKey As String) As Boolean
    '创建新对象
    Dim objNewMember As MediMaster
    
    Set objNewMember = objMaster
    '设置传入方法的属性
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As MediMaster
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

Public Function GetMediMasters(ByVal dateBeging As Date, ByVal DateEnd As Date, ByVal lng科室ID As Long, Optional ByVal lng病人ID As Long)
    '取暂存药品信息,存到MediMasters中
    Dim strSQL As String, rsMain As New ADODB.Recordset, rsBill As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim objMediMaster As MediMaster, objMediBill As MediBill
    Dim i As Integer
    On Error GoTo errHandle
    For i = 1 To mCol.Count
        Remove 1
    Next
    '按入库时间取入库记录,显示在主界面中,明细中包含出库记录
    strSQL = "Select C.编码 as 科室码,C.名称 as 科室名,B.病人ID, A.NO, B.门诊号, B.姓名, B.性别, B.年龄, A.操作员, A.登记时间, A.作废时间, Sum(A.金额 * 入出系数) As 合计, A.摘要" & vbNewLine & _
            "From 部门表 C,病人信息 B, 暂存药品记录 A" & vbNewLine & _
            "Where A.科室ID=C.ID And A.病人id = B.病人id And A.入出系数=1 And A.登记时间 Between [1] And [2] And A.科室id = [3] " & IIf(lng病人ID = 0, "", "And A.病人ID=[4]") & vbNewLine & _
            "Group By C.编码 ,C.名称 ,B.病人ID, A.NO, B.门诊号, B.姓名, B.性别, B.年龄, A.操作员, A.登记时间, A.作废时间, A.摘要"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, "GetMediBills", dateBeging, DateEnd, lng科室ID, lng病人ID)
    Do Until rsMain.EOF
        Set objMediMaster = New MediMaster
        objMediMaster.NO = IIf(IsNull(rsMain.Fields("NO")), "", rsMain.Fields("NO"))
        objMediMaster.门诊号 = IIf(IsNull(rsMain.Fields("门诊号")), 0, rsMain.Fields("门诊号"))
        objMediMaster.姓名 = IIf(IsNull(rsMain.Fields("姓名")), "", rsMain.Fields("姓名"))
        objMediMaster.性别 = IIf(IsNull(rsMain.Fields("性别")), "", rsMain.Fields("性别"))
        objMediMaster.年龄 = IIf(IsNull(rsMain.Fields("年龄")), "", rsMain.Fields("年龄"))
        objMediMaster.操作员 = IIf(IsNull(rsMain.Fields("操作员")), "", rsMain.Fields("操作员"))
        objMediMaster.登记时间 = IIf(IsNull(rsMain.Fields("登记时间")), CDate(0), rsMain.Fields("登记时间"))
        objMediMaster.合计 = IIf(IsNull(rsMain.Fields("合计")), 0, Val(rsMain.Fields("合计")))
        objMediMaster.摘要 = IIf(IsNull(rsMain.Fields("摘要")), "", rsMain.Fields("摘要"))
        objMediMaster.病人ID = IIf(IsNull(rsMain.Fields("病人ID")), 0, rsMain.Fields("病人ID"))
        objMediMaster.科室ID = lng科室ID
        objMediMaster.科室名称 = IIf(IsNull(rsMain.Fields("科室码")), "", "[" & rsMain.Fields("科室码") & "]" & rsMain.Fields("科室名"))
        objMediMaster.作废时间 = IIf(IsNull(rsMain.Fields("作废时间")), CDate(0), rsMain.Fields("作废时间"))
        
        strSQL = "Select a.登记时间, a.单价, a.规格, a.金额, a.单位, a.容量," & vbNewLine & _
                "            a.入出系数, a.使用状态, a.数量, Nvl(b.已用数量, 0) As 已用数量, a.序号, a.药品id, a.药品名称, a.医嘱id, a.发送号," & vbNewLine & _
                "            a.执行分类,a.摘要,a.操作员" & vbNewLine & _
                "From (Select 登记时间, 单价, 规格, 金额, 单位,  容量, 入出系数, 使用状态, 数量," & vbNewLine & _
                "                           序号, 药品id, 药品名称, Nvl(医嘱id, 0) as 医嘱id, Nvl(发送号, 0) as 发送号, 执行分类, 摘要, 操作员" & vbNewLine & _
                "            From 暂存药品记录" & vbNewLine & _
                "            Where No = [1] And 入出系数 = 1 And 作废时间 " & IIf(objMediMaster.作废时间 = CDate(0), " is null", "=to_date('" & Format(objMediMaster.作废时间, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')") & ")  a," & vbNewLine & _
                "        (Select Sum(Nvl(数量, 0)) As 已用数量, 序号, 药品id, 药品名称, Nvl(医嘱id,0) as 医嘱id, nvl(发送号,0) as 发送号, 执行分类" & vbNewLine & _
                "            From 暂存药品记录" & vbNewLine & _
                "            Where No = [1] And 入出系数 = -1 And 作废时间 " & IIf(objMediMaster.作废时间 = CDate(0), " is null", "=to_date('" & Format(objMediMaster.作废时间, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')") & vbNewLine & _
                "            Group By 序号, 药品id, 药品名称, 医嘱id, 发送号, 执行分类) b" & vbNewLine & _
                "Where a.药品id = b.药品id(+) And a.药品名称 = b.药品名称(+) And a.医嘱id = b.医嘱id(+) And a.执行分类 = b.执行分类(+)" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select 登记时间, 单价, 规格, 金额, 单位, 容量, 入出系数, 使用状态, 数量," & vbNewLine & _
                "            0 As 已用数量, 序号, 药品id, 药品名称, nvl(医嘱id,0) as 医嘱id,nvl(发送号,0) as 发送号, 执行分类, 摘要, 操作员" & vbNewLine & _
                "From 暂存药品记录" & vbNewLine & _
                "Where No = [1] And 入出系数 = -1 And 作废时间 " & IIf(objMediMaster.作废时间 = CDate(0), " is null", "=to_date('" & Format(objMediMaster.作废时间, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')")

        Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "GetMediBills", objMediMaster.NO)
        Do Until rsBill.EOF
            Set objMediBill = New MediBill
            objMediBill.单价 = IIf(IsNull(rsBill.Fields("单价")), 0, rsBill.Fields("单价"))
            objMediBill.规格 = IIf(IsNull(rsBill.Fields("规格")), "", rsBill.Fields("规格"))
            objMediBill.金额 = IIf(IsNull(rsBill.Fields("金额")), 0, rsBill.Fields("金额"))
            objMediBill.计算单位 = IIf(IsNull(rsBill.Fields("单位")), "", rsBill.Fields("单位"))
            objMediBill.容量 = IIf(IsNull(rsBill.Fields("容量")), 0, rsBill.Fields("容量"))
            objMediBill.入出系数 = IIf(IsNull(rsBill.Fields("入出系数")), 0, rsBill.Fields("入出系数"))
            objMediBill.使用状态 = IIf(IsNull(rsBill.Fields("使用状态")), "", rsBill.Fields("使用状态"))
            objMediBill.数量 = IIf(IsNull(rsBill.Fields("数量")), 0, rsBill.Fields("数量"))
            objMediBill.已用数量 = IIf(IsNull(rsBill.Fields("已用数量")), 0, rsBill.Fields("已用数量"))
            objMediBill.序号 = IIf(IsNull(rsBill.Fields("序号")), 0, rsBill.Fields("序号"))
            objMediBill.药品ID = IIf(IsNull(rsBill.Fields("药品ID")), 0, rsBill.Fields("药品ID"))
            objMediBill.药品名称 = IIf(IsNull(rsBill.Fields("药品名称")), "", rsBill.Fields("药品名称"))
            objMediBill.医嘱ID = IIf(IsNull(rsBill.Fields("医嘱ID")), 0, rsBill.Fields("医嘱ID"))
            objMediBill.发送号 = IIf(IsNull(rsBill.Fields("发送号")), 0, rsBill.Fields("发送号"))
            objMediBill.执行分类 = IIf(IsNull(rsBill.Fields("执行分类")), 0, rsBill.Fields("执行分类"))
            objMediBill.登记时间 = IIf(IsNull(rsBill.Fields("登记时间")), CDate(0), rsBill.Fields("登记时间"))
            objMediBill.使用摘要 = IIf(IsNull(rsBill.Fields("摘要")), "", rsBill.Fields("摘要"))
            objMediBill.填制人 = IIf(IsNull(rsBill.Fields("操作员")), "", rsBill.Fields("操作员"))
            
            If objMediBill.药品ID > 0 Then
                strSQL = "Select b.剂量单位, a.剂量系数, a.门诊单位, a.门诊包装" & vbNewLine & _
                        "From 药品信息 b, 药品规格 a" & vbNewLine & _
                        "Where a.药名id = b.药名id And a.药品id =[1] "
                        
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get药品规格", objMediBill.药品ID)
                If rsTmp.RecordCount > 0 Then
                    objMediBill.剂量单位 = IIf(IsNull(rsTmp.Fields("剂量单位")), "", rsTmp.Fields("剂量单位"))
                    objMediBill.剂量系数 = IIf(IsNull(rsTmp.Fields("剂量系数")), 0, rsTmp.Fields("剂量系数"))
                    objMediBill.门诊单位 = IIf(IsNull(rsTmp.Fields("门诊单位")), "", rsTmp.Fields("门诊单位"))
                    objMediBill.门诊包装 = IIf(IsNull(rsTmp.Fields("门诊包装")), "", rsTmp.Fields("门诊包装"))
                End If
            Else
                objMediBill.剂量单位 = objMediBill.计算单位
                objMediBill.剂量系数 = 1
                objMediBill.门诊单位 = objMediBill.计算单位
                objMediBill.门诊包装 = 1
            End If
            Call objMediMaster.AddBill(objMediBill, objMediBill.序号 & "_" & objMediBill.入出系数 & "_" & Format(objMediBill.登记时间, "yyMMddhhmmss"))
            rsBill.MoveNext
        Loop
        Call AddMaster(objMediMaster, objMediMaster.NO & "_" & IIf(objMediMaster.作废时间 = CDate(0), "0", Format(objMediMaster.作废时间, "yyMMddhhmmss")))
        rsMain.MoveNext
    Loop
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
