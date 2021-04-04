VERSION 5.00
Begin VB.UserControl DataEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "DataEditor.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "DataEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#####################################################################################
'##     诊治要素编辑器
'#####################################################################################

Option Explicit

'编码、中文名、英文名、类型、长度、小数、单位、性别域、数值域、正常域、初始值、空值文字
Public Enum DataTypeEnum
    dte文本 = 0
    dte上下 = 1
    dte下拉 = 2
    dte复选 = 3
    dte单选 = 4
    dte指针 = 5
End Enum

'0-无限；1-男；2-女，表示该项目适合的患者性别
Public Enum SexLimitEnum
    sle无限 = 0
    sle男 = 1
    sle女 = 2
End Enum

Private mvar所见项ID As Long
Private mvar编码 As String
Private mvar中文名 As String
Private mvar英文名 As String
Private mvar替换域 As Long
Private mvar类型 As DataTypeEnum
Private mvar长度 As Long
Private mvar小数 As Long
Private mvar单位 As String
Private mvar临床意义 As String
Private mvar表示法 As Long
Private mvar性别域 As SexLimitEnum
Private mvar数值域 As String
Private mvar正常域 As String
Private mvar初始值 As String
Private mvar空值文字 As String
Private mvarID As Long
Private mvar分类ID As Long
Private mfrmDataEditor As New frmDataEditor
Private mvarWidth As Long
Private mvarHeight As Long
Public lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lID As Long, bBeteenKeys As Boolean, bNeeded As Boolean


Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
    PropertyChanged "Width"
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
    PropertyChanged "Height"
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let 分类ID(ByVal vData As Long)
    mvar分类ID = vData
    PropertyChanged "分类ID"
End Property

Public Property Get 分类ID() As Long
    分类ID = mvar分类ID
End Property

Public Property Let ID(ByVal vData As Long)
    mvarID = vData
    PropertyChanged "ID"
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property

Public Property Let 空值文字(ByVal vData As String)
    mvar空值文字 = vData
    PropertyChanged "空值文字"
End Property

Public Property Get 空值文字() As String
    空值文字 = mvar空值文字
End Property

Public Property Let 初始值(ByVal vData As String)
    mvar初始值 = vData
    PropertyChanged "初始值"
End Property

Public Property Get 初始值() As String
    初始值 = mvar初始值
End Property

Public Property Let 正常域(ByVal vData As String)
    mvar正常域 = vData
    PropertyChanged "正常域"
End Property

Public Property Get 正常域() As String
    正常域 = mvar正常域
End Property

Public Property Let 数值域(ByVal vData As String)
    mvar数值域 = vData
    PropertyChanged "数值域"
End Property

Public Property Get 数值域() As String
    数值域 = mvar数值域
End Property

Public Property Let 性别域(ByVal vData As SexLimitEnum)
    mvar性别域 = vData
    PropertyChanged "性别域"
End Property

Public Property Get 性别域() As SexLimitEnum
    性别域 = mvar性别域
End Property

Public Property Let 表示法(ByVal vData As Long)
    mvar表示法 = vData
    PropertyChanged "表示法"
End Property

Public Property Get 表示法() As Long
    表示法 = mvar表示法
End Property

Public Property Let 临床意义(ByVal vData As String)
    mvar临床意义 = vData
    PropertyChanged "临床意义"
End Property

Public Property Get 临床意义() As String
    临床意义 = mvar临床意义
End Property

Public Property Let 单位(ByVal vData As String)
    mvar单位 = vData
    PropertyChanged "单位"
End Property

Public Property Get 单位() As String
    单位 = mvar单位
End Property

Public Property Let 小数(ByVal vData As Long)
    mvar小数 = vData
    PropertyChanged "小数"
End Property

Public Property Get 小数() As Long
    小数 = mvar小数
End Property

Public Property Let 长度(ByVal vData As Long)
    mvar长度 = vData
    PropertyChanged "长度"
End Property

Public Property Get 长度() As Long
    长度 = mvar长度
End Property

Public Property Let 类型(ByVal vData As DataTypeEnum)
    mvar类型 = vData
    PropertyChanged "类型"
End Property

Public Property Get 类型() As DataTypeEnum
    类型 = mvar类型
End Property

Public Property Let 替换域(ByVal vData As Long)
    mvar替换域 = vData
    PropertyChanged "替换域"
End Property

Public Property Get 替换域() As Long
    替换域 = mvar替换域
End Property

Public Property Let 英文名(ByVal vData As String)
    mvar英文名 = vData
    PropertyChanged "英文名"
End Property

Public Property Get 英文名() As String
    英文名 = mvar英文名
End Property

Public Property Let 中文名(ByVal vData As String)
    mvar中文名 = vData
    PropertyChanged "中文名"
End Property

Public Property Get 中文名() As String
    中文名 = mvar中文名
End Property

Public Property Let 编码(ByVal vData As String)
    mvar编码 = vData
    PropertyChanged "编码"
End Property

Public Property Get 编码() As String
    编码 = mvar编码
End Property

Public Property Let 所见项ID(ByVal vData As Long)
    mvar所见项ID = vData
    PropertyChanged "所见项ID"
End Property

Public Property Get 所见项ID() As Long
    所见项ID = mvar所见项ID
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    所见项ID = PropBag.ReadProperty("所见项ID", 0)
    编码 = PropBag.ReadProperty("编码", "")
    中文名 = PropBag.ReadProperty("中文名", "")
    英文名 = PropBag.ReadProperty("英文名", "")
    替换域 = PropBag.ReadProperty("替换域", 0)
    类型 = PropBag.ReadProperty("类型", 0)
    长度 = PropBag.ReadProperty("长度", 0)
    小数 = PropBag.ReadProperty("小数", 0)
    单位 = PropBag.ReadProperty("单位", "")
    临床意义 = PropBag.ReadProperty("临床意义", "")
    表示法 = PropBag.ReadProperty("表示法", 0)
    性别域 = PropBag.ReadProperty("性别域", 0)
    数值域 = PropBag.ReadProperty("数值域", "")
    正常域 = PropBag.ReadProperty("正常域", "")
    初始值 = PropBag.ReadProperty("初始值", "")
    空值文字 = PropBag.ReadProperty("空值文字", "")
    ID = PropBag.ReadProperty("ID", 0)
    分类ID = PropBag.ReadProperty("分类ID", 0)
End Sub

Private Sub UserControl_Resize()
    Width = 500
    Height = 480
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "所见项ID", 所见项ID, 0
    PropBag.WriteProperty "编码", 编码, ""
    PropBag.WriteProperty "中文名", 中文名, ""
    PropBag.WriteProperty "英文名", 英文名, ""
    PropBag.WriteProperty "替换域", 替换域, 0
    PropBag.WriteProperty "类型", 类型, 0
    PropBag.WriteProperty "长度", 长度, 0
    PropBag.WriteProperty "小数", 小数, 0
    PropBag.WriteProperty "单位", 单位, ""
    PropBag.WriteProperty "临床意义", 临床意义, ""
    PropBag.WriteProperty "表示法", 表示法, 0
    PropBag.WriteProperty "性别域", 性别域, 0
    PropBag.WriteProperty "数值域", 数值域, ""
    PropBag.WriteProperty "正常域", 正常域, ""
    PropBag.WriteProperty "初始值", 初始值, ""
    PropBag.WriteProperty "空值文字", 空值文字, ""
    PropBag.WriteProperty "ID", ID, 0
    PropBag.WriteProperty "分类ID", 分类ID, 0
    
    PropertyChanged "所见项ID"
    PropertyChanged "编码"
    PropertyChanged "中文名"
    PropertyChanged "英文名"
    PropertyChanged "替换域"
    PropertyChanged "类型"
    PropertyChanged "长度"
    PropertyChanged "小数"
    PropertyChanged "单位"
    PropertyChanged "临床意义"
    PropertyChanged "表示法"
    PropertyChanged "性别域"
    PropertyChanged "数值域"
    PropertyChanged "正常域"
    PropertyChanged "初始值"
    PropertyChanged "空值文字"
    PropertyChanged "ID"
    PropertyChanged "分类ID"
End Sub

Public Sub ShowEditor(x As Long, y As Long, lWidth As Long, lHeight As Long, eType As DataTypeEnum)
    With mfrmDataEditor
        .lKSS = lKSS
        .lKSE = lKSE
        .lKES = lKES
        .lKEE = lKEE
        .lID = lID
        .bNeeded = bNeeded
        .ShowDataEditor x, y, lWidth, lHeight, UserControl.Parent, eType, mvar数值域, mvar初始值, mvar中文名, _
        mvar英文名, mvar长度, mvar小数, mvar单位, mvar性别域, mvar正常域, mvar空值文字
    End With
End Sub






























