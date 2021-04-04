VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Seating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvar序号 As Integer '局部复制
Private mvar病人ID As Long '局部复制
Private mvar门诊号 As String '局部复制
Private mvar姓名 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar编号 As String '局部复制
Private mvar类别 As Integer '局部复制
Private mvar状态 As Integer '局部复制
Private mvar现价 As Currency '局部复制
Private mvar收费细目ID As Long '局部复制
Private mvar收费项目 As String '局部复制
Private mvar备注 As String '局部复制
Private mvar类型  As Long  '局部复制
Private mvarKey As String

'---- 配合移动临床修改
Private mvar呼叫器编号 As String    '
Private mvar开始时间 As String
Private mvar诊断     As String
Private mvar分类    As String
Private mvar挂号单  As String
Private mstrPatientKey As String

Public Property Let Key(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费细目ID = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'检索属性值时使用，位于赋值语句的右边。
' X.收费细目ID
    Key = mvarKey
End Property

'病人的Key；门诊：病人ID_挂号单号；门诊留观：病人ID_主页ID
Public Property Let PatientKey(ByVal strVar As String)
    mstrPatientKey = strVar
End Property
Public Property Get PatientKey() As String
    PatientKey = mstrPatientKey
End Property


Public Property Let 备注(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费细目ID = 5
    mvar备注 = MidUni(vData, 1, 100)
End Property


Public Property Get 备注() As String
'检索属性值时使用，位于赋值语句的右边。
' X.收费细目ID
    备注 = mvar备注
End Property

Public Property Let 收费项目(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费细目ID = 5
    mvar收费项目 = vData
End Property


Public Property Get 收费项目() As String
'检索属性值时使用，位于赋值语句的右边。
' X.收费细目ID
    收费项目 = mvar收费项目
End Property

Public Property Let 收费细目ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费细目ID = 5
    mvar收费细目ID = vData
End Property


Public Property Get 收费细目ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.收费细目ID
    收费细目ID = mvar收费细目ID
End Property



Public Property Let 现价(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.现价 = 5
    mvar现价 = vData
End Property


Public Property Get 现价() As Currency
'检索属性值时使用，位于赋值语句的右边。
' X.现价
    现价 = mvar现价
End Property

Public Property Let 类型(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar类型 = vData
End Property


Public Property Get 类型() As Currency
'检索属性值时使用，位于赋值语句的右边。
' X.类型
    类型 = mvar类型
End Property

Public Property Let 状态(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.状态 = 5
    mvar状态 = vData
End Property


Public Property Get 状态() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.状态
    状态 = mvar状态
End Property



Public Property Let 分类(ByVal vData As String)
    mvar分类 = vData
End Property


Public Property Get 分类() As String
    分类 = mvar分类
End Property
Public Property Let 类别(ByVal vData As Integer)
    mvar类别 = vData
End Property


Public Property Get 类别() As Integer
    类别 = mvar类别
End Property


Public Property Let 编号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.编号 = 5
    mvar编号 = vData
End Property


Public Property Get 编号() As String
'检索属性值时使用，位于赋值语句的右边。
' X.编号
    编号 = mvar编号
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

Public Property Let 挂号单(ByVal vData As String)
    mvar挂号单 = vData
End Property

Public Property Get 挂号单() As String
    挂号单 = mvar挂号单
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

Public Property Let 呼叫器编号(ByVal vData As String)
    mvar呼叫器编号 = vData
End Property

Public Property Get 呼叫器编号() As String
    呼叫器编号 = mvar呼叫器编号
End Property

Public Property Let 开始时间(ByVal vData As String)
    mvar开始时间 = vData
End Property

Public Property Get 开始时间() As String
    开始时间 = mvar开始时间
End Property

Public Property Let 诊断(ByVal vData As String)
    mvar诊断 = vData
End Property

Public Property Get 诊断() As String
    诊断 = mvar诊断
End Property

Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
End Property


Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.病人ID
    病人ID = mvar病人ID
End Property

Public Property Let 序号(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property


Public Property Get 序号() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.序号
    序号 = mvar序号
End Property

Public Function Update(科室ID As Long, 收费细目ID As Long, 状态 As Integer, 收费项目 As String, 现价 As Currency, 备注 As String, 类型 As Long, 呼叫器 As String) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select A.病人ID,B.姓名 From 座位状况记录 A,病人信息 B Where  A.病人ID=B.病人ID(+) And A.科室ID=[1] And A.编号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Seating.Update", 科室ID, CStr(mvar编号))
    If rsTmp.EOF Then
        Update = "座位已被删除，请重新选择座位。"
        Exit Function
    End If
    
    If rsTmp!病人ID <> 0 Then
        '座位有人，不能更改
        Update = "已被" & rsTmp!姓名 & "(" & 病人ID & ") 占用"
    Else
        '安排座位 或 清除座位
       
        strSQL = "ZL_座位状况记录_UPDATE(" & 科室ID & ",'" & mvar编号 & "'," & IIf(收费细目ID = 0, "Null", 收费细目ID) & "," & 状态 & ",'" & DelInvalidChar(备注, "'") & "'," & 类型 & ",'" & 呼叫器 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "seating.Update")
        mvar收费细目ID = 收费细目ID
        mvar收费项目 = 收费项目
        mvar现价 = 现价
        mvar备注 = 备注
        mvar呼叫器编号 = 呼叫器

    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

