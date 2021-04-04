VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlPrintDbGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'局部变量保存属性值
Private mvarAppColor As Long '本地复本
Private mvarAppFont As StdFont '本地复本
Private mvarBodyGrid As Object    '本地复本
Private mvarEmptyDown As Single '本地复本
Private mvarEmptyLeft As Single '本地复本
Private mvarEmptyUp As Single '本地复本
Private mvarTitle As zlTabTitle '本地复本
Private mvarBelowAppRows As zlTabAppRows '本地复本
Private mvarUnderAppRows As zlTabAppRows '本地复本
Private mvarUseGrdColor As Boolean '本地复本
Private mvarDataSource As Object  '本地复本
Private mvarFixRow As Integer '局部复制
Private mvarFixCol As Integer '局部复制
Private mvarEmptyRight As Single '局部复制
Private mvarPageHeader As Single '局部复制
Private mvarPageFooter As Single '局部复制
Private mvarHeader As String '局部复制
Private mvarFooter As String '局部复制
Private mvarPrivileged As Boolean '局部复制

Public Property Get Privileged() As Boolean
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.Privileged
    Privileged = mvarPrivileged
End Property

Public Property Let Privileged(ByVal vData As Boolean)
'当给属性赋值时在参数左边使用。
'Syntax: X.Privileged = True
    mvarPrivileged = vData
End Property
Public Property Let Footer(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Footer = 5
    mvarFooter = vData
End Property


Public Property Get Footer() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Footer
    Footer = mvarFooter
End Property



Public Property Let Header(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Header = 5
    mvarHeader = vData
End Property


Public Property Get Header() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Header
    Header = mvarHeader
End Property



Public Property Let PageFooter(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PageFooter = 5
    mvarPageFooter = vData
End Property


Public Property Get PageFooter() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PageFooter
    PageFooter = mvarPageFooter
End Property



Public Property Let PageHeader(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PageHeader = 5
    mvarPageHeader = vData
End Property


Public Property Get PageHeader() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PageHeader
    PageHeader = mvarPageHeader
End Property



Public Property Let EmptyRight(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.EmptyRight = 5
    mvarEmptyRight = vData
End Property


Public Property Get EmptyRight() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.EmptyRight
    EmptyRight = mvarEmptyRight
End Property



Public Property Let FixCol(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FixCol = 5
    mvarFixCol = vData
End Property


Public Property Get FixCol() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FixCol
    FixCol = mvarFixCol
End Property



Public Property Let FixRow(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FixRow = 5
    mvarFixRow = vData
End Property


Public Property Get FixRow() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FixRow
    FixRow = mvarFixRow
End Property




Public Property Set DataSource(ByVal vData As Object)
'当把对象赋值给属性时在 Set 语句左边使用。
'Syntax: Set x.DataSource = Form1
    Set mvarDataSource = vData
End Property


Public Property Get DataSource() As Object
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.DataSource
    Set DataSource = mvarDataSource
End Property

Public Property Let UseGrdColor(ByVal vData As Boolean)
'当给属性赋值时在参数左边使用。
'Syntax: X.UseGrdColor = 5
    mvarUseGrdColor = vData
End Property


Public Property Get UseGrdColor() As Boolean
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.UseGrdColor
    UseGrdColor = mvarUseGrdColor
End Property


Public Property Set UnderAppRows(ByVal vData As Object)
'当把对象赋值给属性时在 Set 语句左边使用。
'Syntax: Set x.UnderAppRows = Form1
    Set mvarUnderAppRows = vData
End Property


Public Property Get UnderAppRows() As zlTabAppRows
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.UnderAppRows
    Set UnderAppRows = mvarUnderAppRows
End Property



Public Property Set BelowAppRows(ByVal vData As Object)
'当把对象赋值给属性时在 Set 语句左边使用。
'Syntax: Set x.BelowAppRows = Form1
    Set mvarBelowAppRows = vData
End Property


Public Property Get BelowAppRows() As zlTabAppRows
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.BelowAppRows
    Set BelowAppRows = mvarBelowAppRows
End Property



Public Property Set Title(ByVal vData As Object)
'当把对象赋值给属性时在 Set 语句左边使用。
'Syntax: Set x.Title = Form1
    Set mvarTitle = vData
End Property


Public Property Get Title() As zlTabTitle
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.Title
    Set Title = mvarTitle
End Property

Public Property Let EmptyUp(ByVal vData As Single)
'当给属性赋值时在参数左边使用。
'Syntax: X.EmptyUp = 5
    mvarEmptyUp = vData
End Property


Public Property Get EmptyUp() As Single
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.EmptyUp
    EmptyUp = mvarEmptyUp
End Property



Public Property Let EmptyLeft(ByVal vData As Single)
'当给属性赋值时在参数左边使用。
'Syntax: X.EmptyLeft = 5
    mvarEmptyLeft = vData
End Property


Public Property Get EmptyLeft() As Single
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.EmptyLeft
    EmptyLeft = mvarEmptyLeft
End Property



Public Property Let EmptyDown(ByVal vData As Single)
'当给属性赋值时在参数左边使用。
'Syntax: X.EmptyDown = 5
    mvarEmptyDown = vData
End Property


Public Property Get EmptyDown() As Single
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.EmptyDown
    EmptyDown = mvarEmptyDown
End Property



Public Property Set BodyGrid(ByVal vData As Object)
'当把对象赋值给属性时在 Set 语句左边使用。
'Syntax: Set x.BodyGrid = Form1
    Set mvarBodyGrid = vData
End Property


Public Property Get BodyGrid() As Object
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.BodyGrid
    Set BodyGrid = mvarBodyGrid
End Property



Public Property Set AppFont(ByVal vData As Object)
'当把对象赋值给属性时在 Set 语句左边使用。
'Syntax: Set x.AppFont = Form1
    Set mvarAppFont = vData
End Property


Public Property Get AppFont() As StdFont
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.AppFont
    Set AppFont = mvarAppFont
End Property



Public Property Let AppColor(ByVal vData As Long)
'当给属性赋值时在参数左边使用。
'Syntax: X.AppColor = 5
    mvarAppColor = vData
End Property


Public Property Get AppColor() As Long
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.AppColor
    AppColor = mvarAppColor
End Property



Private Sub Class_Initialize()
    '当创建 zlPrint1Grd 类时，创建 mTitle 对象
    Set mvarTitle = New zlTabTitle
    '当创建 zlPrint1Grd 类时，创建 mAppFont属性
    Set mvarAppFont = New StdFont
    mvarAppFont.Name = "宋体"
    mvarAppFont.Size = 9
    mvarAppFont.Bold = False
    mvarAppFont.Italic = False
    '当创建 zlPrint1Grd 类时，创建 mUnderAppRows属性
    Set mvarUnderAppRows = New zlTabAppRows
    '当创建 zlPrint1Grd 类时，创建 mBelowAppRows属性
    Set mvarBelowAppRows = New zlTabAppRows
    '当创建 zlPrint1Grd 类时，创建相关属性
    mvarAppColor = 0
    mvarEmptyUp = 25
    mvarEmptyDown = 25
    mvarEmptyLeft = 19
    mvarEmptyRight = 19
    mvarPageHeader = 13
    mvarPageFooter = 13
    mvarFooter = ";;"
    mvarHeader = ";;"

    mvarUseGrdColor = True
    mvarPrivileged = False
    Set mvarDataSource = New ADODB.Recordset

End Sub


Private Sub Class_Terminate()
    Set mvarBodyGrid = Nothing
'    Set mvarDataSource = Nothing
    Set mvarBelowAppRows = Nothing
    Set mvarUnderAppRows = Nothing
    Set mvarTitle = Nothing
    Set mvarAppFont = Nothing
End Sub



