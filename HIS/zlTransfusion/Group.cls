VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Bills"
Option Explicit

'保持属性值的局部变量
'- 病人医嘱发送
Private mvar执行医嘱ID As Long '局部复制
Private mvar发送号 As Long '局部复制
Private mvar发送数次 As Integer
Private mvar已执行数次  As Integer

Private mvar执行分类 As String '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗
Private mvar皮试结果 As String  '(+) 阴性 (-)阳性
Private mvar执行频次 As String
Private mvar用法 As String
Private mvar频率次数 As Integer
Private mvar滴速 As Integer
Private mvar执行状态 As Integer  '0-未执行;1-完全执行;2-拒绝执行;3-正在执行
Private mvar组次 As Integer
Private mvar液体量 As Long
Private mvar耗时 As Long
Private mvar说明 As String
Private mvar医生嘱托 As String
Private mvar执行人 As String
Private mvar计费状态 As Integer  '-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费(对收费类，实际是否能执行，需要根据对应费用单据是否完成收费判断)
'要求时间,本次数次,执行摘要,执行时间
Private mvar要求时间 As Date
Private mvar本次数次 As Integer
Private mvar上次执行时间 As Date
Private mvar收费金额 As Currency '0 未收费 >0表示已收费
Private mvar记录性质 As Integer  '配合一卡通修改 1-收费记录,2-记帐记录
Private mvar上次组次 As Integer     '上次组次
Private mvar核对人 As String
'--
' 病人医嘱记录
Private mColBills As Collection '-- 医嘱执行的具体内容集合,例 如静脉输液的药品明细


Public Property Let 核对人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar核对人 = vData
End Property

Public Property Get 核对人() As String
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     核对人 = mvar核对人
End Property


Public Property Let 记录性质(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.记录性质 = 5
    mvar记录性质 = vData
End Property


Public Property Get 记录性质() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.记录性质
    记录性质 = mvar记录性质
End Property

Public Property Let 要求时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar要求时间 = vData
End Property

Public Property Get 要求时间() As Date
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     要求时间 = mvar要求时间
End Property

Public Property Let 收费金额(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar收费金额 = vData
End Property

Public Property Get 收费金额() As Currency
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     收费金额 = mvar收费金额
End Property

Public Property Let 上次执行时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar上次执行时间 = vData
End Property

Public Property Get 上次执行时间() As Date
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     上次执行时间 = mvar上次执行时间
End Property

Public Function AddBills(ByVal objBIlls As Bills, Optional sKey As String) As Bills
    Dim objNewMember As Bills
    Set objNewMember = New Bills
    
    Set objNewMember = objBIlls
    If Len(sKey) = 0 Then
        mColBills.Add objNewMember
    Else
        mColBills.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set AddBills = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get BillsItem(vntIndexKey As Variant) As Bills
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set BillsItem = mColBills(vntIndexKey)
End Property

Public Property Get BillsCount() As Long
    ' x.Count
    BillsCount = mColBills.Count
End Property


Public Sub BillsRemove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    mColBills.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mColBills.[_NewEnum]
End Property

Public Property Let 组次(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar组次 = vData
End Property

Public Property Get 组次() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     组次 = mvar组次
End Property
Public Property Let 上次组次(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar上次组次 = vData
End Property

Public Property Get 上次组次() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     上次组次 = mvar上次组次
End Property
Public Property Let 本次数次(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar本次数次 = vData
End Property

Public Property Get 本次数次() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     本次数次 = mvar本次数次
End Property

Public Property Let 发送数次(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar发送数次 = vData
End Property

Public Property Get 发送数次() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     发送数次 = mvar发送数次
End Property

Public Property Let 已执行数次(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar已执行数次 = vData
End Property

Public Property Get 已执行数次() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     已执行数次 = mvar已执行数次
End Property

Public Property Let 执行分类(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar执行分类 = vData
End Property

Public Property Get 执行分类() As String
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     执行分类 = mvar执行分类
End Property

Public Property Let 执行人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar执行人 = vData
End Property

Public Property Get 执行人() As String
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     执行人 = mvar执行人
End Property

Public Property Let 皮试结果(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar皮试结果 = vData
End Property

Public Property Get 皮试结果() As String
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     皮试结果 = mvar皮试结果
End Property

Public Property Let 执行频次(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar执行频次 = vData
End Property

Public Property Get 执行频次() As String
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     执行频次 = mvar执行频次
End Property

Public Property Let 滴速(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar滴速 = vData
End Property

Public Property Get 滴速() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     滴速 = mvar滴速
End Property

Public Property Let 计费状态(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar计费状态 = vData
End Property

Public Property Get 计费状态() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     计费状态 = mvar计费状态
End Property

Public Property Let 频率次数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.频率次数  = 1
    mvar频率次数 = vData
End Property

Public Property Get 频率次数() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.频率次数
     频率次数 = mvar频率次数
End Property

Public Property Let 用法(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行状态  = "0-未执行"
    mvar用法 = vData
End Property

Public Property Get 用法() As String
'检索属性值时使用，位于赋值语句的右边。
' X.执行状态
     用法 = mvar用法
End Property

Public Property Let 执行状态(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行状态  = "0-未执行"
    mvar执行状态 = vData
End Property

Public Property Get 执行状态() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.执行状态
     执行状态 = mvar执行状态
End Property

Public Property Let 发送号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发送号 = 5
    mvar发送号 = vData
End Property


Public Property Get 发送号() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.发送号
    发送号 = mvar发送号
End Property

Public Property Let 执行医嘱ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行医嘱ID = 5
    mvar执行医嘱ID = vData
End Property

Public Property Get 执行医嘱ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.执行医嘱ID
    执行医嘱ID = mvar执行医嘱ID
End Property

Public Property Get 液体量() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.执行医嘱ID
    Dim objBIll As Bill
    mvar液体量 = 0
    For Each objBIll In Me.BillsItem(Me.执行医嘱ID & "_" & Me.发送号)
        mvar液体量 = mvar液体量 + objBIll.容量
    Next
    液体量 = mvar液体量
End Property

Public Property Get 耗时() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.执行医嘱ID
    Dim objBIll As Bill
    mvar耗时 = 0
    For Each objBIll In Me.BillsItem(Me.执行医嘱ID & "_" & Me.发送号)
        mvar耗时 = mvar耗时 + objBIll.时间
    Next
    耗时 = mvar耗时
End Property

Public Property Get 说明() As String
'检索属性值时使用，位于赋值语句的右边。
' X.执行医嘱ID
    Dim objBIll As Bill
    If mvar说明 = "" Then
        mvar医生嘱托 = ""
        For Each objBIll In Me.BillsItem(Me.执行医嘱ID & "_" & Me.发送号)
            mvar医生嘱托 = mvar医生嘱托 & objBIll.医生嘱托
        Next
        说明 = mvar医生嘱托
    Else
        说明 = mvar说明
    End If
End Property

Public Property Let 说明(ByVal vDate As String)
'检索属性值时使用，位于赋值语句的右边。
' X.执行医嘱ID
    mvar说明 = MidUni(vDate, 1, 200)
End Property

Private Sub Class_Terminate()
    Set mColBills = Nothing
End Sub

Private Sub Class_Initialize()
    '创建类后创建集合
    Set mColBills = New Collection
End Sub

Public Function FuncExecRefuse() As Boolean
    '功能：接单时拒绝执行
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String
    
    '正在执行或已执行不允许拒绝
    
    If Me.执行状态 = 2 Then
        MsgBox "该执行项目当前已经拒绝执行。", vbInformation, gstrSysName
        Exit Function
    End If
    If Me.执行状态 = 3 Then
        MsgBox "该执行项目当前正在执行，不能拒绝。", vbInformation, gstrSysName
        Exit Function
    End If
    If Me.执行状态 = 1 Then
        MsgBox "该执行项目当前已经执行，不能拒绝。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If MsgBox("确实要拒绝执行该项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    lng医嘱ID = Me.执行医嘱ID
    lng发送号 = Me.发送号
    
    strSQL = "ZL_病人医嘱执行_拒绝执行(" & lng医嘱ID & "," & lng发送号 & ")"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.FuncExecRefuse")
    On Error GoTo 0
    Me.执行状态 = 2
    
    FuncExecRefuse = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncExecRestore() As Boolean
'功能：取消拒绝执行
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String, i As Long
    
    If Me.执行状态 <> 2 Then
        MsgBox "该执行项目没有被拒绝执行。", vbInformation, gstrSysName
        Exit Function
    End If
        
    If MsgBox("确实要取消拒绝执行该项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    lng医嘱ID = Me.执行医嘱ID
    lng发送号 = Me.发送号
    
    strSQL = "ZL_病人医嘱执行_取消拒绝(" & lng医嘱ID & "," & lng发送号 & ")"
    
    On Error GoTo errH
 
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.FuncExecRestore")
 
    On Error GoTo 0
    Me.执行状态 = 0
    
    FuncExecRestore = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
