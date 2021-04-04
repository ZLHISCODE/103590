VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisReportLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'报告纸张参数定义
'定义
'######################################################################################################################

Public Enum USERCOLOR
    兰色 = 16711680
    白色 = 16777215
    黑色 = 0
    背景色 = 10191921
    表格线色 = 5880731
    表格背景色 = 15790320
    
    项目标题背景色 = -2147483636
    
    红色 = &HFF&
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
End Enum

Public Type USERRECT
    X0      As Single
    Y0      As Single
    X1      As Single
    Y1      As Single
    R0      As Single                            '行间距
    B0      As Single                            '边界
    Page    As Integer                           '页号
End Type

Public Type USERFONT
    Name            As String
    Size            As Single
    Bold            As Boolean
    Italic          As Boolean
    Underline       As Boolean
    Strikethrough   As Boolean
    ForeColor       As Long
    BackColor       As Long
    LineWidth       As Integer
    LineStyle       As Integer
    HAlignment      As Byte                              '横向对齐
    VAlignment      As Byte                              '纵向对齐
    WordWarp        As Boolean                           '自动换行
    MuliLine        As Boolean                           '多行文本
    AutoFit         As Boolean                           '自动适应
    Rows            As Integer                           '行数
End Type

Public Type USERPAPER
    Width As Single
    Height As Single
    BorderLeft As Single
    BorderRight As Single
    BorderTop As Single
    BorderBottom As Single
    PageHead As Single
    PageFoot As Single
    SpaceTop As Single
    SpaceBottom As Single
    PaperSize As Integer
    Orientation As Byte
End Type

Public Event FormatData()

Public Property Get RECT() As USERRECT
    RECT = gobjRect
End Property

Public Property Let RECT(vData As USERRECT)
    gobjRect = vData
End Property

Public Property Get Font() As USERFONT
    Font = gobjFont
End Property

Public Property Let Font(vData As USERFONT)
    gobjFont = vData
    
    gobjDraw.FontName = gobjFont.Name
    gobjDraw.FontSize = gobjFont.Size
    gobjDraw.FontBold = gobjFont.Bold
    gobjDraw.FontItalic = gobjFont.Italic
    
End Property

Public Property Get Paper() As USERPAPER
    Paper = gobjPaper
End Property

Public Property Let Paper(vData As USERPAPER)
    gobjPaper = vData
End Property

Public Property Get Draw() As Object
    Set Draw = gobjDraw
End Property

Public Function GetTextWidth(ByVal strText As String) As Single
    GetTextWidth = gobjDraw.TextWidth(strText)
End Function

Public Function GetTextHeight(ByVal strText As String) As Single
    GetTextHeight = gobjDraw.TextHeight(strText)
End Function

Public Function GetTwipsX(ByVal sglNumber As Single) As Single
    GetTwipsX = gobjDraw.ScaleX(sglNumber, vbCentimeters, vbTwips)
End Function

Public Function GetTwipsY(ByVal sglNumber As Single) As Single
    GetTwipsY = gobjDraw.ScaleY(sglNumber, vbCentimeters, vbTwips)
End Function

Public Function GetCentimetersX(ByVal sglNumber As Single) As Single
    GetCentimetersX = gobjDraw.ScaleX(sglNumber, vbTwips, vbCentimeters)
End Function

Public Function GetCentimetersY(ByVal sglNumber As Single) As Single
    GetCentimetersY = gobjDraw.ScaleY(sglNumber, vbTwips, vbCentimeters)
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Private Sub Class_Initialize()
    gobjFont.Name = "宋体"
    gobjFont.Size = 12
    gobjFont.Bold = False
    gobjFont.Underline = False
    gobjFont.Strikethrough = False
    gobjFont.Italic = False
End Sub

