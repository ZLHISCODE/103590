VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MediBill"
Attribute VB_Ext_KEY = "Member0" ,"MediBill"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
Option Explicit

Private mvarNO As String '局部复制
Private mvar挂号单 As String '局部复制
Private mvar门诊号 As String
Private mvar姓名 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar年龄 As String '局部复制
Private mvar操作员 As String '局部复制
Private mvar登记时间 As Date '局部复制
Private mvar作废时间 As Date '局部复制
Private mvar合计 As Currency '局部复制
Private mvar摘要 As String '局部复制
Private mvar病人ID As Long
Private mvar科室ID As Long
Private mvar科室名称 As String

Private mColBills As Collection  '保存Bill

Public Property Let 摘要(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.摘要 = 5
    mvar摘要 = MidUni(vData, 1, 200)
End Property

Public Property Get 摘要() As String
'检索属性值时使用，位于赋值语句的右边。
' X.摘要
    摘要 = mvar摘要
End Property

Public Property Let 合计(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.合计 = 5
    mvar合计 = vData
End Property


Public Property Get 合计() As Currency
'检索属性值时使用，位于赋值语句的右边。
' X.合计
    合计 = mvar合计
End Property


Public Property Let 登记时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
    mvar登记时间 = vData
End Property


Public Property Get 登记时间() As Date
'检索属性值时使用，位于赋值语句的右边。
' X.登记时间
    登记时间 = mvar登记时间
End Property

Public Property Let 作废时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
    mvar作废时间 = vData
End Property


Public Property Get 作废时间() As Date
'检索属性值时使用，位于赋值语句的右边。
' X.登记时间
    作废时间 = mvar作废时间
End Property


Public Property Let 操作员(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.操作员 = 5
    mvar操作员 = vData
End Property


Public Property Get 操作员() As String
'检索属性值时使用，位于赋值语句的右边。
' X.操作员
    操作员 = mvar操作员
End Property



Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar年龄 = vData
End Property


Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
' X.年龄
    年龄 = mvar年龄
End Property


Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mvar性别 = vData
End Property


Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
' X.性别
    性别 = mvar性别
End Property

Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
End Property

Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
' X.姓名
    姓名 = mvar姓名
End Property

Public Property Let 科室名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar科室名称 = vData
End Property

Public Property Get 科室名称() As String
'检索属性值时使用，位于赋值语句的右边。
' X.姓名
    科室名称 = mvar科室名称
End Property

Public Property Let 挂号单(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。

    mvar挂号单 = vData
End Property

Public Property Get 挂号单() As String
'检索属性值时使用，位于赋值语句的右边。

    挂号单 = mvar挂号单
End Property

Public Property Let 门诊号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊号 = 5
    mvar门诊号 = vData
End Property

Public Property Get 门诊号() As String
'检索属性值时使用，位于赋值语句的右边。
' X.门诊号
    门诊号 = mvar门诊号
End Property

Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。

    mvar病人ID = vData
End Property

Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。

    病人ID = mvar病人ID
End Property

Public Property Let 科室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。

    mvar科室ID = vData
End Property

Public Property Get 科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。

    科室ID = mvar科室ID
End Property

Public Property Let NO(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.NO = 5
    mvarNO = vData
End Property

Public Property Get NO() As String
'检索属性值时使用，位于赋值语句的右边。
' X.NO
    NO = mvarNO
End Property

Public Property Get BillItem(vntIndexKey As Variant) As MediBill
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set BillItem = mColBills(vntIndexKey)
End Property

Public Property Get BillCount() As Long
    ' x.Count
    BillCount = mColBills.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    mColBills.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mColBills.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '创建类后创建集合
    Set mColBills = New Collection
End Sub

Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mColBills = Nothing
End Sub

Public Sub AddBill(ByVal objMediBill As MediBill, Optional ByVal StrKey As String)
    '创建新对象
    Dim objNewMember As MediBill
    
    Set objNewMember = objMediBill
    '设置传入方法的属性
    If Len(StrKey) = 0 Then
        mColBills.Add objNewMember
    Else
        mColBills.Add objNewMember, StrKey
    End If
    Set objNewMember = Nothing
    
End Sub

Public Function GetNextNo() As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    GetNextNo = zlDatabase.GetNextNo(19)
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub InsertBill(ByVal strNO As String, ByVal dateCurr As Date)
    Dim strInsSQL As String, strSQL As String, rsTmp As New ADODB.Recordset
    Dim lng相关ID As Long, lng已执行数次 As Long, lng发送数量 As Long, date发送时间 As Date
    Dim objBIll As MediBill
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For Each objBIll In mColBills
        With objBIll
        
            '---- 医嘱增加时,要将已执行的数量同步写入,暂存表
'            If .医嘱ID <> 0 And .药品ID <> 0 And .入出系数 = 1 Then
'                strSQL = "Select A.相关id,B.发送时间, (B.发送数次 / C.剂量系数) As 数量" & vbNewLine & _
'                        "From 药品规格 C, 病人医嘱发送 B, 病人医嘱记录 A" & vbNewLine & _
'                        "Where A.收费细目id = C.药品id And A.ID = B.医嘱id And B.发送号 = [2] And B.医嘱id = [1]"
'                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InsertBill", .医嘱ID, .发送号)
'                If Not rsTmp.EOF Then
'                    lng相关ID = Val(rsTmp.Fields("相关ID"))
'                    lng发送数量 = Val(rsTmp.Fields("数量"))
'                    date发送时间 = CDate(rsTmp.Fields("发送时间"))
'                End If
'
'                If lng相关ID <> 0 And lng发送数量 > 0 Then
'                    strSQL = "Select B.发送数次, Sum(Nvl(A.本次数次, 0)) As 已执行数次" & vbNewLine & _
'                            "From 病人医嘱发送 B, 病人医嘱执行 A" & vbNewLine & _
'                            "Where A.医嘱id(+) = B.医嘱id And A.发送号(+) = B.发送号 And B.医嘱id = [1] And B.发送号 = [2]" & vbNewLine & _
'                            "Group By B.发送数次"
'                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InsertBill", lng相关ID, .发送号)
'
'                    If Not rsTmp.EOF Then
'                        '换算出已用数量
'                        lng已执行数次 = (Val(rsTmp.Fields("已执行数次")) / Val(rsTmp.Fields("发送数次"))) * lng发送数量
'                        If lng已执行数次 > .数量 Then
'                            lng已执行数次 = .数量
'                        End If
'                        If lng已执行数次 > 0 Then
'                            .使用状态 = 1
'                            strInsSQL = "ZL_暂存药品记录_Insert('" & strNO & "'," & .序号 & "," & _
'                                     mvar病人ID & "," & .医嘱ID & "," & .发送号 & "," & .药品ID & ",'" & _
'                                     .药品名称 & "','" & .规格 & "'," & .执行分类 & "," & _
'                                     .使用状态 & ",'" & "自动写入" & "'," & -1 & ",'" & _
'                                     .计算单位 & "'," & .容量 & "," & lng已执行数次 & "," & _
'                                     .单价 & "," & lng已执行数次 * .单价 & ",'" & mvar操作员 & "'," & _
'                                     mvar科室ID & ",to_date('" & Format(date发送时间, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
'                            Call zlDatabase.ExecuteProcedure(strInsSQL, "MediMaster.InsertBill")
'                        End If
'                    End If
'                End If
'            End If
            '-- 写入寄存记录
            strInsSQL = "ZL_暂存药品记录_Insert('" & strNO & "'," & .序号 & "," & _
                     mvar病人ID & "," & IIf(.医嘱ID = 0, "Null", .医嘱ID) & "," & IIf(.发送号 = 0, "Null", .发送号) & "," & .药品ID & ",'" & _
                     .药品名称 & "','" & .规格 & "'," & .执行分类 & "," & _
                     .使用状态 & ",'" & mvar摘要 & "'," & .入出系数 & ",'" & _
                     .计算单位 & "'," & .容量 & "," & .数量 & "," & _
                     .单价 & "," & .金额 & ",'" & mvar操作员 & "'," & _
                     mvar科室ID & ",to_date('" & Format(dateCurr, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strInsSQL, "MediMaster.InsertBill")
        End With
    Next
    gcnOracle.CommitTrans
    Exit Sub
errHandle:
    
    Call ErrCenter
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Public Property Get 使用情况() As String
'检索属性值时使用，位于赋值语句的右边。

    Dim objBIll As MediBill
    Dim int使用 As Integer
    For Each objBIll In mColBills
        If objBIll.入出系数 = -1 Then
            int使用 = int使用 + 1
        End If
    Next
    
    If int使用 > 0 Then
        使用情况 = "已用"
    Else
        使用情况 = "未用"
    End If
    
End Property

Public Sub DeleteBill(ByVal intType As Integer)
    'inttype  : 0 执行删除操作，1-执行UPdate时的删除操作
    Dim strNO As String, strSQL As String
    On Error GoTo errHandle
    strNO = Me.NO
    
    If Me.使用情况 <> "未用" And intType = 0 Then
        If MsgBox("该张单据中的药品已经使用，如果删除，这张单据连同使用记录将一同被清除，不能恢复！" & vbNewLine & " 请确认，是否要执行删除操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    ElseIf Me.使用情况 <> "未用" And intType = 1 Then
        MsgBox "该张单据中的药品已经使用，不能删除！" & vbNewLine & " 请确认，是否要执行删除操作？", vbQuestion, gstrSysName
        Exit Sub
    ElseIf Me.使用情况 = "未用" And intType = 0 Then
        If MsgBox("删除该张单据后，这张单据的数据将永远消失，不能恢复。" & vbNewLine & " 请确认，是否要执行删除操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
     
    strSQL = "Zl_暂存药品记录_Delete('" & strNO & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "DeleteBill")
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub InsertUseBill(ByVal strBillKey As String, ByVal dblUse As Double, ByVal dateUse As Date)

    Dim objBIll As MediBill, strInsSQL As String
    
    On Error GoTo errHandle
    Set objBIll = Me.BillItem(strBillKey)
    If Not objBIll Is Nothing Then
        With objBIll
            If (.数量 - .已用数量) >= dblUse And dblUse > 0 Then
                .使用状态 = 1
                strInsSQL = "ZL_暂存药品记录_Insert('" & mvarNO & "'," & .序号 & "," & _
                         mvar病人ID & "," & IIf(.医嘱ID = 0, "Null", .医嘱ID) & "," & IIf(.发送号 = 0, "Null", .发送号) & "," & .药品ID & ",'" & _
                         .药品名称 & "','" & .规格 & "'," & .执行分类 & "," & _
                         .使用状态 & ",'" & mvar摘要 & "'," & -1 & ",'" & _
                         .计算单位 & "'," & .容量 & "," & dblUse & "," & _
                         .单价 & "," & dblUse * .单价 & ",'" & UserInfo.姓名 & "'," & _
                         mvar科室ID & ",to_date('" & Format(dateUse, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
                Call zlDatabase.ExecuteProcedure(strInsSQL, "MediMaster.InsertUseBill")
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UndoUse(ByVal strBillKey As String)
    Dim objBIll As MediBill, strSQL As String
    On Error GoTo errHandle
    Set objBIll = Me.BillItem(strBillKey)
    If Not objBIll Is Nothing Then
        With objBIll
            If .医嘱ID = 0 Then
                strSQL = "ZL_暂存药品记录_UndoUse('" & mvarNO & "'," & .序号 & "," & .入出系数 & ",To_date('" & Format(.登记时间, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
                Call zlDatabase.ExecuteProcedure(strSQL, "MediMaster.UndoUse")
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
