VERSION 5.00
Begin VB.UserControl udSeat 
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "udSeat.ctx":0000
   ScaleHeight     =   2325
   ScaleMode       =   0  'User
   ScaleWidth      =   2250
   Begin VB.PictureBox picSeat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2320
      Left            =   0
      ScaleHeight     =   2325
      ScaleWidth      =   2250
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2250
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "诊断"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Top             =   1680
         Width           =   1950
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "护士 03:21"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   1185
         Width           =   1950
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "001"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   630
         TabIndex        =   2
         Top             =   165
         Width           =   825
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "张三 23岁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   630
         TabIndex        =   1
         Top             =   690
         Width           =   1455
      End
      Begin VB.Line lineBorder 
         BorderColor     =   &H00FF0000&
         Index           =   3
         X1              =   2205
         X2              =   2205
         Y1              =   15
         Y2              =   2275
      End
      Begin VB.Line lineBorder 
         BorderColor     =   &H00FF0000&
         Index           =   2
         X1              =   15
         X2              =   2215
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line lineBorder 
         BorderColor     =   &H00FF0000&
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   2275
      End
      Begin VB.Line lineBorder 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   15
         X2              =   2215
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Image imgSeatStat 
         Height          =   480
         Index           =   2
         Left            =   1560
         Picture         =   "udSeat.ctx":018A
         Top             =   45
         Width           =   480
      End
      Begin VB.Image imgSex 
         Height          =   480
         Index           =   0
         Left            =   60
         Picture         =   "udSeat.ctx":0A54
         Top             =   570
         Width           =   480
      End
      Begin VB.Line LineGrid 
         BorderColor     =   &H00FF0000&
         Index           =   2
         X1              =   30
         X2              =   2200
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line LineGrid 
         BorderColor     =   &H00FF0000&
         Index           =   1
         X1              =   30
         X2              =   2200
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Line LineGrid 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   30
         X2              =   2200
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Image imgSeat 
         Height          =   480
         Index           =   0
         Left            =   60
         Picture         =   "udSeat.ctx":171E
         Top             =   45
         Width           =   480
      End
      Begin VB.Image imgSex 
         Height          =   480
         Index           =   1
         Left            =   60
         Picture         =   "udSeat.ctx":1FE8
         Top             =   585
         Width           =   480
      End
      Begin VB.Image imgSeat 
         Height          =   480
         Index           =   1
         Left            =   60
         Picture         =   "udSeat.ctx":2CB2
         Top             =   45
         Width           =   480
      End
   End
End
Attribute VB_Name = "udSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---- 主要用于显示输液人员的静态信息
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private mstrSeatNo As String    '编号
Private mstrPatiName As String  '病人姓名
Private mstrSex As String      '性别 1-男 2-女 0-未知
Private mstrStart As String  '开始时间
Private mstrDiagnosis As String     '诊断

Private mintStat As Integer          '状态 0-空 1－有人  2-在维护
Private mintSeatType As Integer         '座位 0-坐位 1－床位

Private mintWidth As Integer    '宽 2600
Private mintHeight As Integer   '高 3000

Private mlngGridColor As Long   '表格线色
Private mintGridWidth As Integer    '表格线粗细
Private mstrKey As String

'-- 表格线色
Property Get GridColor() As Long
    GridColor = mlngGridColor
End Property
Property Let GridColor(ByVal Value As Long)
    mlngGridColor = Value
    PropertyChanged "GridColor"
    lineBorder(0).BorderColor = mlngGridColor
    lineBorder(1).BorderColor = mlngGridColor
    lineBorder(2).BorderColor = mlngGridColor
    lineBorder(3).BorderColor = mlngGridColor
    LineGrid(0).BorderColor = mlngGridColor
    LineGrid(1).BorderColor = mlngGridColor
    LineGrid(2).BorderColor = mlngGridColor
    
End Property
'--- 表格线粗

Property Get GridWidth() As Integer
    GridWidth = mintGridWidth
End Property
Property Let GridWidth(ByVal Value As Integer)
    mintGridWidth = Value
    PropertyChanged "GridWidth"
    lineBorder(0).BorderWidth = mintGridWidth
    lineBorder(1).BorderWidth = mintGridWidth
    lineBorder(2).BorderWidth = mintGridWidth
    lineBorder(3).BorderWidth = mintGridWidth
    LineGrid(0).BorderWidth = mintGridWidth
    LineGrid(1).BorderWidth = mintGridWidth
    LineGrid(2).BorderWidth = mintGridWidth
    
End Property

'--性别
Property Get Sex() As String
    Sex = mstrSex
End Property

Property Let Sex(ByVal strValue As String)
    mstrSex = strValue
    PropertyChanged "Sex"
    If mstrSex = "男" Then
        imgSex(0).Visible = True
        imgSex(1).Visible = False
    ElseIf mstrSex = "女" Then
        imgSex(0).Visible = False
        imgSex(1).Visible = True
    Else
        imgSex(0).Visible = False
        imgSex(1).Visible = False
        
    End If
End Property

'-- 位子类型
Property Get SeatType() As Integer
    SeatType = mintSeatType
End Property
 
Property Let SeatType(ByVal strValue As Integer)
    mintSeatType = strValue
    PropertyChanged "SeatType"
    If mintSeatType = 0 Then
        imgSeat(0).Visible = True
        imgSeat(1).Visible = False
    ElseIf mintSeatType = 1 Then
        imgSeat(0).Visible = False
        imgSeat(1).Visible = True
    End If
End Property

'-- 状态
Property Get Stat() As Integer
    Stat = mintStat
End Property
 
Property Let Stat(ByVal strValue As Integer)
    mintStat = strValue
    PropertyChanged "Stat"
    If mintStat = 2 Then
        imgSeatStat(2).Visible = True
    Else
        imgSeatStat(2).Visible = False
    End If
End Property

'-- 编号
Property Get SeatNo() As String
    SeatNo = mstrSeatNo
End Property
 
Property Let SeatNo(ByVal strValue As String)
    mstrSeatNo = strValue
    PropertyChanged "SeatNo"
    lbl(0).Caption = strValue
End Property

'-- 姓名
Property Get PatiName() As String
    PatiName = mstrPatiName
End Property
 
Property Let PatiName(ByVal strValue As String)
    mstrPatiName = strValue
    PropertyChanged "PatiName"
    lbl(1).Caption = strValue
End Property

'-- 开始时间
Property Get Start() As String
    Start = mstrStart
End Property
 
Property Let Start(ByVal strValue As String)
    mstrStart = strValue
    PropertyChanged "Start"
    lbl(2).Caption = strValue
End Property

'-- 诊断
Property Get Diagnosis() As String
    Diagnosis = mstrDiagnosis
End Property
 
Property Let Diagnosis(ByVal strValue As String)
    mstrDiagnosis = strValue
    PropertyChanged "Diagnosis"
    lbl(3).Caption = strValue
End Property

'-- Key
Property Get Key() As String
    Key = mstrKey
End Property
 
Property Let Key(ByVal strValue As String)
    mstrKey = strValue
End Property

Private Sub imgSeat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgSeat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

Private Sub imgSeatStat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgSeatStat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub imgSex_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgSex_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picSeat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picSeat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    imgSeat(0).Visible = False
    imgSeat(1).Visible = False
    
    
    imgSeatStat(2).Visible = False
    
    imgSex(0).Visible = False
    imgSex(1).Visible = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
