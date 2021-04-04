VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmPatiConnect 
   Caption         =   "病人身份关联"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatiConnect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12465
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   3735
      TabIndex        =   37
      Top             =   1440
      Width           =   3735
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   660
         Left            =   720
         TabIndex        =   38
         Top             =   480
         Width           =   1320
         _Version        =   589884
         _ExtentX        =   2328
         _ExtentY        =   1164
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   480
         ScaleWidth      =   3615
         TabIndex        =   41
         Top             =   3720
         Width           =   3615
         Begin VB.Label lblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "取消关联"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   2880
            TabIndex        =   44
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "未关联"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   1800
            TabIndex        =   43
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "已关联"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   600
            TabIndex        =   42
            Top             =   120
            Width           =   630
         End
         Begin VB.Image img 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   2
            Left            =   2400
            Picture         =   "frmPatiConnect.frx":6852
            Top             =   0
            Width           =   480
         End
         Begin VB.Image img 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   1
            Left            =   1200
            Picture         =   "frmPatiConnect.frx":711C
            Top             =   0
            Width           =   480
         End
         Begin VB.Image img 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frmPatiConnect.frx":79E6
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   3960
      ScaleHeight     =   6135
      ScaleWidth      =   7815
      TabIndex        =   1
      Top             =   360
      Width           =   7815
      Begin VB.Frame fraInfo 
         BackColor       =   &H8000000E&
         Caption         =   "病人信息"
         Height          =   6045
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7575
         Begin VB.TextBox txtAddTime 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2025
         End
         Begin VB.TextBox txt住院次数 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   4320
            Width           =   2025
         End
         Begin VB.TextBox txt床位 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   3960
            Width           =   2025
         End
         Begin VB.TextBox txt科室 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   3600
            Width           =   2025
         End
         Begin VB.TextBox txt家庭地址 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   3240
            Width           =   2025
         End
         Begin VB.TextBox txt出生地点 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2880
            Width           =   2025
         End
         Begin VB.TextBox txt身份证号 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2520
            Width           =   2025
         End
         Begin VB.TextBox txt身份 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2025
         End
         Begin VB.TextBox txt职业 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2025
         End
         Begin VB.TextBox txt婚姻状况 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2025
         End
         Begin VB.TextBox txt学历 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2025
         End
         Begin VB.TextBox txt民族 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2025
         End
         Begin VB.TextBox txt国籍 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2025
         End
         Begin VB.TextBox txt出生日期 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2025
         End
         Begin VB.TextBox txt性别 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   2025
         End
         Begin VB.TextBox txt姓名 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   720
            Width           =   2025
         End
         Begin VB.TextBox txt住院号 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   2025
         End
         Begin VB.TextBox txt状态 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblAddTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "登记时间"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4035
            TabIndex        =   40
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label lbl住院号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院号"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   4245
            TabIndex        =   36
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lbl状态 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "状态"
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   660
            TabIndex        =   35
            Top             =   360
            Width           =   420
         End
         Begin VB.Label lbl住院次数 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院次数"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   240
            TabIndex        =   34
            Top             =   4320
            Width           =   840
         End
         Begin VB.Label lbl床位 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "床位"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   660
            TabIndex        =   33
            Top             =   3960
            Width           =   420
         End
         Begin VB.Label lbl科室 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "科室"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   660
            TabIndex        =   32
            Top             =   3600
            Width           =   420
         End
         Begin VB.Label lbl家庭地址 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "现住址"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   450
            TabIndex        =   31
            Top             =   3240
            Width           =   630
         End
         Begin VB.Label lbl出生地点 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生地点"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   240
            TabIndex        =   30
            Top             =   2880
            Width           =   840
         End
         Begin VB.Label lbl身份证号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份证号"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   240
            TabIndex        =   29
            Top             =   2520
            Width           =   840
         End
         Begin VB.Label lbl职业 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "职业"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   615
            TabIndex        =   28
            Top             =   1800
            Width           =   420
         End
         Begin VB.Label lbl身份 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4455
            TabIndex        =   27
            Top             =   1800
            Width           =   420
         End
         Begin VB.Label lbl婚姻状况 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "婚姻状况"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   240
            TabIndex        =   26
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label lbl学历 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "学历"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   660
            TabIndex        =   25
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lbl民族 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "民族"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4455
            TabIndex        =   24
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lbl国籍 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "国籍"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4455
            TabIndex        =   23
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label lbl出生日期 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生日期"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   240
            TabIndex        =   22
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lbl性别 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4455
            TabIndex        =   21
            Top             =   720
            Width           =   420
         End
         Begin VB.Label lbl姓名 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   660
            TabIndex        =   20
            Top             =   720
            Width           =   420
         End
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8010
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1588
            MinWidth        =   1587
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18830
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   988
            MinWidth        =   988
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsPati 
      Left            =   1560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":82B0
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":EB12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":15374
            Key             =   "link"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":1BBD6
            Key             =   "linkAdd"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":22438
            Key             =   "linkNew"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":28C9A
            Key             =   "linkdelete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":2F4FC
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":35D5E
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":362F8
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":36892
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   840
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatiConnect.frx":3D0F4
      Left            =   3240
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPatiConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Object
Private mrsPati    As ADODB.Recordset   '"病人ID", "关联ID", "姓名", "性别", "年龄", "出生日期", "身份证号", "家庭地址"

Private mlngPatiId As Long
Private mblnUndo   As Boolean
Private mbytEditState As Byte    '当前编辑状态：0-非编辑状态,1-编辑状态
Private mlngLinkID    As Long       '记录选中行的关联ID
Private mstrPrivs As String
Private mstrFilter As String
Private mbytFunc As Byte            '=1 费用调用：传入病人ID自动查找相同身份病人，由用户决定是否自动关联。

Private Const M_BGK_CORLOR As Long = &HEBFFFF

Private Enum PATI_COLUMN
    COL_图标 = 0
    COL_姓名
    COL_性别
    COL_年龄
    COL_出生日期
    COL_身份证号
    COL_家庭地址
    COL_登记时间
    '隐藏列
    COL_关联ID
    COL_病人Id
    COL_国籍
    COL_民族
    COL_EDIT            '0-原始;1-自动关联;2-增加关联;3-取消关联
End Enum

Private Enum E_EDIT
    E_LINKLOAD = 0
    E_LINKAUTO
    E_LINKADD
    E_LINKCANCEL
End Enum

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Manage_RelatingPatiet * 10# + 1 '自动关联
        If UpdatePati(E_EDIT.E_LINKAUTO) Then mbytEditState = 1
    Case conMenu_Manage_RelatingPatiet * 10# + 2 '增加关联
        If UpdatePati(E_EDIT.E_LINKADD) Then mbytEditState = 1
    Case conMenu_Manage_RelatingPatiet * 10# + 3    '取消关联
        If UpdatePati(E_EDIT.E_LINKCANCEL) Then mbytEditState = 1
    Case conMenu_Edit_Save
        Call SaveData
        mbytEditState = 0
        If mbytFunc = 1 Then Unload Me: Exit Sub
    Case conMenu_Edit_Untread
        Call LoadPati(E_LINKLOAD)
        mbytEditState = 0
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.staThis.Visible Then Bottom = Me.staThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    picLeft.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    staThis.Width = lngRight - lngLeft
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    
    Case conMenu_Manage_RelatingPatiet * 10# + 1, conMenu_Manage_RelatingPatiet * 10# + 2
        Control.Enabled = mbytEditState = 0
    Case conMenu_Manage_RelatingPatiet * 10# + 3
        Control.Enabled = mlngLinkID <> 0 And mbytEditState = 0
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = mbytEditState = 1 Or (Control.ID = conMenu_Edit_Save And mbytFunc = 1)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As XtremeDockingPane.Pane
    Call InitCommandBar
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    
    Set objPane = Me.dkpMain.CreatePane(1, 320, 400, DockLeftOf, Nothing)
    objPane.Title = "关联病人列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Call InitReportColumn
    If mbytFunc = 0 Then
        Call LoadPati(E_LINKLOAD)
    Else
        Call LoadPati(E_LINKAUTO)
    End If
    img(0).Picture = ilsPati.ListImages("link").Picture
    img(1).Picture = ilsPati.ListImages("linkAdd").Picture
    img(2).Picture = ilsPati.ListImages("linkdelete").Picture
    lblNote(0).Caption = "已关联"
    lblNote(1).Caption = "待关联"
    lblNote(2).Caption = "待取消"
    
    Call RestoreWinState(Me, App.ProductName, , True)
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
End Sub


Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl

    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    If mbytFunc = 0 Then
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RelatingPatiet * 10# + 1, "自动关联")
            objControl.IconId = conMenu_Kss_Adjustment
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RelatingPatiet * 10# + 2, "增加关联")
            objControl.IconId = conMenu_Kss_Grant
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RelatingPatiet * 10# + 3, "取消关联")
            objControl.IconId = conMenu_Kss_Cancellation
            
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "确定"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
            Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True
        End With
        
        objBar.EnableDocking xtpFlagHideWrap
        objBar.ContextMenuPresent = False
        For Each objControl In objBar.Controls
            If objControl.type <> xtpControlCustom And objControl.type <> xtpControlLabel Then
                objControl.Style = xtpButtonIconAndCaption
            End If
        Next
        
        With cbsMain.KeyBindings
            .Add FALT, vbKeyQ, conMenu_File_Exit
            .Add FALT, vbKeyS, conMenu_Edit_Save
        End With
    Else
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "确定")
            Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True
        End With
        
        objBar.EnableDocking xtpFlagHideWrap
        objBar.ContextMenuPresent = False
        For Each objControl In objBar.Controls
            If objControl.type <> xtpControlCustom And objControl.type <> xtpControlLabel Then
                objControl.Style = xtpButtonIconAndCaption
            End If
        Next
        
        With cbsMain.KeyBindings
            .Add FALT, vbKeyQ, conMenu_File_Exit
        End With
    End If
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        .Columns.DeleteAll
        Set objCol = .Columns.Add(COL_图标, "", 20, False)
        Set objCol = .Columns.Add(COL_姓名, "姓名", 80, True)
        Set objCol = .Columns.Add(COL_性别, "性别", 45, True)
        Set objCol = .Columns.Add(COL_年龄, "年龄", 45, True)
        Set objCol = .Columns.Add(COL_出生日期, "出生日期", 80, True)
        Set objCol = .Columns.Add(COL_身份证号, "身份证号", 150, True)
        Set objCol = .Columns.Add(COL_家庭地址, "现住址", 180, True)
        Set objCol = .Columns.Add(COL_登记时间, "登记时间", 150, True)
        
        Set objCol = .Columns.Add(COL_关联ID, "关联ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_病人Id, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_国籍, "国籍", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_民族, "民族", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_EDIT, "编辑", 0, False): objCol.Visible = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的关联病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        mblnUndo = True
        .MultipleSelection = False '会引发SelectionChanged事件
         mblnUndo = False
        .ShowItemsInGroups = False
        .SetImageList Me.ilsPati
    End With
End Sub

Public Function ShowMe(ByRef frmParent As Object, ByVal strPrivs As String, ByVal lngPatiID As Long, Optional ByVal bytFunc As Byte = 0) As Boolean
'功能:身份关联
    Dim rsPati As ADODB.Recordset
    
    If lngPatiID = 0 Then Exit Function
    Set mfrmParent = frmParent
    mstrPrivs = strPrivs
    mstrFilter = ""
    mbytFunc = bytFunc
    mlngPatiId = lngPatiID
    If mbytFunc = 1 Then
        Set mrsPati = GetPatiLinked(mlngPatiId)
        Set mrsPati = zlDatabase.CopyNewRec(mrsPati, , , Array("EDIT", adInteger, 2, Empty))
        mrsPati.Filter = "病人ID =" & mlngPatiId
        If Not mrsPati.EOF Then
            mstrFilter = mrsPati!国籍 & "|" & mrsPati!民族 & "|" & mrsPati!性别 & _
            "|" & mrsPati!姓名 & "|" & Format(mrsPati!出生日期 & "", "YYYY-MM-DD") & "|" & mrsPati!身份证号
        End If
        Set rsPati = GetPatiSimilar(mstrFilter)
        If Not AppendPatiSimilar(rsPati, E_LINKAUTO) Then Exit Function
    End If
    Me.Show 1, frmParent
    ShowMe = True
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SaveWinState(Me, App.ProductName)
    Set mrsPati = Nothing
End Sub

Private Sub picLeft_Resize()
    
    On Error Resume Next
    fraInfo.Move 120, 120, picLeft.ScaleWidth - 240, picLeft.ScaleHeight - 240
End Sub

Private Sub picNote_Resize()
    On Error Resume Next
    img(0).Move 120, 120, 480, 480
    img(1).Move img(0).Left + 1380, 120, 480, 480
    img(2).Move img(1).Left + 1380, 120, 480, 480
    lblNote(0).Move img(0).Left + 360, 120
    lblNote(1).Move img(1).Left + 360, 120
    lblNote(2).Move img(2).Left + 360, 120
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    rptPati.Move 0, 0, picPati.ScaleWidth, picPati.ScaleHeight - picNote.Height
    picNote.Move 0, picPati.Height - picNote.Height, picPati.ScaleWidth
End Sub

Private Function LoadPati(ByVal bytFunc As Byte, Optional ByVal lngPatiID As Long) As Boolean
'功能:加载病人身份关联列表
'参数:
'bytFunc=0-原始加载;1-自动关联;2-增加关联;3-取消关联
'lngPatiID-病人ID
'   strSimilar '国籍|民族|性别|姓名|出生日期(To_Date('2015/4/30', 'YYYY-MM-DD'))|身份证号
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim lngSelID As Long
    
    On Error GoTo errH
    If mbytFunc = 0 Then
        If bytFunc = E_LINKLOAD Then
            Set rsPati = GetPatiLinked(mlngPatiId)
        ElseIf bytFunc = E_LINKADD Then
            Set rsPati = GetPatiLinked(lngPatiID)
        ElseIf bytFunc = E_LINKAUTO Then
            Set rsPati = GetPatiSimilar(mstrFilter)
        End If
        
        If bytFunc <> E_LINKCANCEL Then
            If rsPati.EOF Then
                If bytFunc = E_LINKAUTO Then
                    MsgBox "未发现身份相似的病人信息！", vbInformation, gstrSysName
                Else
                    MsgBox "未发现该病人的身份信息！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
        
        If bytFunc = E_LINKLOAD Then
            Set mrsPati = zlDatabase.CopyNewRec(rsPati, , , Array("EDIT", adInteger, 2, Empty))      'COPY用于后续增删改
        ElseIf bytFunc = E_LINKADD Or bytFunc = E_LINKAUTO Then
            If Not AppendPatiSimilar(rsPati, bytFunc) Then Exit Function
        ElseIf bytFunc = E_LINKCANCEL Then
            mrsPati.Filter = ""
        End If
    Else
        mrsPati.Filter = ""
    End If
    mrsPati.Sort = "登记时间 ASC"
    '加载病人列表
    Call ClearPatiInfo
    rptPati.Tag = ""
    rptPati.Records.DeleteAll
    rptPati.SortOrder.DeleteAll
    For i = 1 To mrsPati.RecordCount
        Set objRecord = rptPati.Records.Add()
        Set objItem = objRecord.AddItem("")
        If CLng(mrsPati!关联ID) > 0 And Val(mrsPati!EDIT & "") = E_LINKCANCEL Then
            objItem.Icon = ilsPati.ListImages("linkdelete").Index - 1
        ElseIf CLng(mrsPati!关联ID) > 0 Then
            objItem.Icon = ilsPati.ListImages("link").Index - 1
        ElseIf CLng(mrsPati!关联ID) = 0 Then
            objItem.Icon = ilsPati.ListImages("linkAdd").Index - 1
        End If
        objRecord.AddItem mrsPati!姓名 & ""
        objRecord.AddItem mrsPati!性别 & ""
        objRecord.AddItem mrsPati!年龄 & ""
        
        objRecord.AddItem Format(mrsPati!出生日期, "YYYY-MM-DD")
        objRecord.AddItem mrsPati!身份证号 & ""
        objRecord.AddItem Nvl(mrsPati!家庭地址, "未登记")
        objRecord.AddItem Format(mrsPati!登记时间, "YYYY-MM-DD HH:MM:SS")
        '隐藏列
        objRecord.AddItem CLng(mrsPati!关联ID)
        objRecord.AddItem CLng(mrsPati!病人ID)
        objRecord.AddItem mrsPati!国籍 & ""
        objRecord.AddItem mrsPati!民族 & ""
        objRecord.AddItem Nvl(mrsPati!EDIT, "0")
        If CLng(mrsPati!病人ID) = mlngPatiId And bytFunc = E_LINKLOAD And mstrFilter = "" Then
            mstrFilter = mrsPati!国籍 & "|" & mrsPati!民族 & "|" & mrsPati!性别 & _
            "|" & mrsPati!姓名 & "|" & Format(mrsPati!出生日期 & "", "YYYY-MM-DD") & "|" & mrsPati!身份证号
        End If
        mrsPati.MoveNext
    Next
    
    If bytFunc = E_LINKAUTO Then
        For i = 0 To rptPati.Records.Count - 1
            If Val(rptPati.Records.Record(i).Item(COL_EDIT).Value) = E_LINKAUTO Then
                For j = COL_姓名 To COL_登记时间
                    rptPati.Records.Record(i).Item(j).BackColor = M_BGK_CORLOR
                Next
            End If
        Next
    End If
    '当前病人字体加粗
    For i = 0 To rptPati.Records.Count - 1
        If Val(rptPati.Records.Record(i).Item(COL_病人Id).Value) = mlngPatiId Then
            For j = COL_姓名 To COL_登记时间
                rptPati.Records.Record(i).Item(j).Bold = True
            Next
            Exit For
        End If
    Next
    rptPati.Populate
    
    If bytFunc = E_LINKLOAD Or bytFunc = E_LINKAUTO Then
        lngSelID = mlngPatiId
    ElseIf bytFunc = E_LINKADD Or bytFunc = E_LINKCANCEL Then
        lngSelID = lngPatiID
    End If
    For i = 0 To rptPati.Records.Count - 1
        If Val(rptPati.Records.Record(i).Item(COL_病人Id).Value) = lngSelID Then
            Set rptPati.FocusedRow = rptPati.Rows(i)
            Exit For
        End If
    Next
    LoadPati = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPatiInfo(lngID As Long) As Boolean
'功能：显示一个病人信息
'参数：lngID=病人ID
    Dim rsTmp As New ADODB.Recordset, rsPati As ADODB.Recordset
    Dim strSQL As String
    Dim str住院号 As String, str门诊号 As String
    Dim strJsonAsk As String, strJsonOut As String
    Dim colReturn As Collection
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "Select 病人ID,门诊号,住院号,就诊卡号,姓名,性别,年龄,出生日期,出生地点,身份证号,身份,职业,民族,国籍,区域,学历,婚姻状况,家庭地址,家庭电话,登记时间" & _
             "  From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        
    If rsTmp.EOF Then
        MsgBox "未发现该病人的身份信息！", vbInformation, gstrSysName
        Exit Function
    End If

    txt姓名.Text = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
    txt性别.Text = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
    txt出生日期.Text = Format(IIf(IsNull(rsTmp!出生日期), "", rsTmp!出生日期), "yyyy年MM月dd日")
    txt国籍.Text = IIf(IsNull(rsTmp!国籍), "", rsTmp!国籍)
    txt民族.Text = IIf(IsNull(rsTmp!民族), "", rsTmp!民族)
    txt学历.Text = IIf(IsNull(rsTmp!学历), "", rsTmp!学历)
    txt身份.Text = IIf(IsNull(rsTmp!身份), "", rsTmp!身份)
    txt职业.Text = IIf(IsNull(rsTmp!职业), "", rsTmp!职业)
    txt身份证号.Text = IIf(IsNull(rsTmp!身份证号), "", rsTmp!身份证号)
    txt出生地点.Text = IIf(IsNull(rsTmp!出生地点), "", rsTmp!出生地点)
    txt家庭地址.Text = IIf(IsNull(rsTmp!家庭地址), "", rsTmp!家庭地址)
    txt婚姻状况.Text = IIf(IsNull(rsTmp!婚姻状况), "", rsTmp!婚姻状况)
    txtAddTime.Text = Format(Nvl(rsTmp!登记时间), "YYYY-MM-DD HH:MM:SS")
    str门诊号 = IIf(IsNull(rsTmp!门诊号), "", rsTmp!门诊号)
    str住院号 = IIf(IsNull(rsTmp!住院号), "", rsTmp!住院号)
    strSQL = "Select a.病人id, a.主页id,a.住院次数 " & vbNewLine & _
            "From 病人信息 a " & _
            "Where a.病人id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    Set rsPati = InitRecordset("出院日期||30,住院号||18,出院病床||50,病人id|adInteger|18,主页id|adInteger|18,住院次数|adInteger|3,科室||100")
    
    strJsonAsk = "{""input"":{""query_type"":1,""pati_pageids"":""" & lngID & """}}"
    If CallService("Zl_Cissvr_Getpatipageinfo", strJsonAsk, strJsonOut) Then
        Set colReturn = gobjService.GetJsonListValue("output.page_list")
    End If
    If Not rsTmp.EOF And Not colReturn Is Nothing Then
        For i = 1 To colReturn.Count
            If zval(rsTmp!病人ID & "") = zval(gobjService.GetCollValue(colReturn, i, "_pati_id")) And zval(rsTmp!主页ID & "") = zval(gobjService.GetCollValue(colReturn, i, "_pati_pageid")) Then
                With rsPati
                    .AddNew Array("出院日期", "住院号", "出院病床", "病人id", "主页id", "住院次数", "科室"), Array(Format(gobjService.GetCollValue(colReturn, i, "_adtd_time") & "", "YYYY-MM-DD HH:MM:SS"), gobjService.GetCollValue(colReturn, i, "_inpatient_num"), _
                    gobjService.GetCollValue(colReturn, i, "_pati_bed"), gobjService.GetCollValue(colReturn, i, "_pati_id"), gobjService.GetCollValue(colReturn, i, "_pati_pageid"), _
                    Val(rsTmp!住院次数 & ""), gobjService.GetCollValue(colReturn, i, "_pati_dept_name"))
                    .Update
                End With
            End If
        Next
        If Not rsPati.EOF Then rsPati.MoveFirst
    End If
    Set rsTmp = rsPati.Clone
    If rsTmp.EOF Then
        If glngSys Like "8??" Then
            txt状态.Text = "正常"
        Else
            txt状态.Text = "门诊"
        End If
        lbl住院号.Caption = "门诊号"
        txt住院号.Text = IIf(str门诊号 = "", "", str门诊号)
        txt科室.Text = ""
        txt床位.Text = ""
        txt住院次数.Text = ""
    Else
        txt状态.Text = IIf(IsNull(rsTmp!出院日期), "在院", "出院")
        lbl住院号.Caption = "住院号"
        txt住院号.Text = IIf(str住院号 = "", "", str住院号)
        txt科室.Text = rsTmp!科室
        txt床位.Text = IIf(IsNull(rsTmp!出院病床), "家庭", rsTmp!出院病床)
        txt住院次数.Text = Nvl(rsTmp!住院次数)
    End If
    
    ShowPatiInfo = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearPatiInfo()
'功能：清除一个病人信息
'参数：x=控件索引,0=源病人,1=目标病人
    txt姓名.Text = ""
    txt性别.Text = ""
    txt出生日期.Text = ""
    txt国籍.Text = ""
    txt民族.Text = ""
    txt学历.Text = ""
    txt身份.Text = ""
    txt职业.Text = ""
    txt身份证号.Text = ""
    txt出生地点.Text = ""
    txt家庭地址.Text = ""
    txt婚姻状况.Text = ""
    txt状态.Text = ""
    lbl住院号.Caption = "住院号"
    txt住院号.Text = ""
    txt科室.Text = ""
    txt床位.Text = ""
    txt住院次数.Text = ""
    txtAddTime.Text = ""
End Sub

Private Sub rptPati_SelectionChanged()
'功能:
    If mblnUndo Then Exit Sub
    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '非正常情况
    Me.staThis.Panels(2).Text = ""
    With rptPati.SelectedRows(0)
        If .GroupRow Then
            Call ClearPatiInfo
        Else
            If rptPati.Tag = .Record(COL_病人Id).Value Then Exit Sub
            mlngLinkID = Val(.Record(COL_关联ID).Value & "")
            Call ShowPatiInfo(Val(.Record(COL_病人Id).Value & ""))
            rptPati.Tag = .Record(COL_病人Id).Value
        End If
        If Val(.Record(COL_EDIT).Value) = E_LINKAUTO Then
            Me.staThis.Panels(2).Text = "身份证号相同或姓名|性别|年龄|出生日期|国籍都相同的病人。"
        End If
    End With
End Sub


Private Function UpdatePati(ByVal bytFunc As Byte) As Boolean
'功能:增加关联
'参数:bytFunc 1-自动关联;2-增加关联;3-取消关联
    Dim objFrm As New frmPatiSel
    If bytFunc = E_EDIT.E_LINKADD Then
        objFrm.mstrPrivs = mstrPrivs
        objFrm.Show 1, Me
        If objFrm.mlng病人ID <> 0 Then
            mrsPati.Filter = "病人ID=" & objFrm.mlng病人ID
            If mrsPati.RecordCount > 0 Then
                MsgBox "该病人已经在【关联病人列表】中，无需再增加！", vbInformation + vbOKOnly, gstrSysName: Exit Function
            End If
            UpdatePati = LoadPati(bytFunc, objFrm.mlng病人ID)
        End If
    ElseIf bytFunc = E_EDIT.E_LINKAUTO Then
        UpdatePati = LoadPati(bytFunc)
    ElseIf bytFunc = E_EDIT.E_LINKCANCEL Then
        With rptPati.SelectedRows(0)
            mrsPati.Filter = "病人ID=" & .Record(COL_病人Id).Value
            mrsPati!EDIT = E_EDIT.E_LINKCANCEL
            UpdatePati = LoadPati(bytFunc, Val(.Record(COL_病人Id).Value))
        End With
    End If
End Function

Private Function SaveData() As Boolean
'功能:保存更新
    Dim strTime As String
    Dim strPatiID As String
    Dim lngLinKID As Long
    Dim arrSQL As Variant
    Dim blnTrans As Boolean
    Dim i As Long
    
    On Error GoTo errH
    arrSQL = Array()
    mrsPati.Filter = "EDIT=" & E_LINKCANCEL
    If mrsPati.RecordCount > 0 Then
        If Val(mrsPati!关联ID & "") > 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人身份关联_Update(1," & mrsPati!关联ID & ",'" & mrsPati!病人ID & "')"
        End If
    Else
        mrsPati.Filter = ""
        mrsPati.Sort = "关联ID ASC"
        For i = 1 To mrsPati.RecordCount
            If Val(mrsPati!关联ID & "") <> 0 Then
                If lngLinKID = 0 Then lngLinKID = Val(mrsPati!关联ID & "")
                If Val(mrsPati!关联ID & "") <> lngLinKID Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人身份关联_Update(2," & lngLinKID & ",'" & mrsPati!病人ID & "')"
                End If
            Else
                strPatiID = strPatiID & "," & mrsPati!病人ID
            End If
            mrsPati.MoveNext
        Next
        If strPatiID <> "" Then
            strPatiID = Mid(strPatiID, 2)
            strTime = "TO_DATE('" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人身份关联_Update(0," & lngLinKID & ",'" & strPatiID & "','" & UserInfo.姓名 & "'," & strTime & ")"
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    SaveData = LoadPati(E_LINKLOAD)
    Exit Function
errH:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiSimilar(ByVal strSimilar As String) As ADODB.Recordset
'参数: strSimilar '国籍|民族|性别|姓名|出生日期(To_Date('2015/4/30', 'YYYY-MM-DD'))|身份证号
    Dim arrTmp As Variant
    Dim strSQL As String
    
    arrTmp = Split(strSimilar, "|")
    
    strSQL = "Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.身份证号, a.家庭地址, a.国籍, a.民族, 0 As 关联id, a.登记时间" & vbNewLine & _
            "From 病人信息 A" & vbNewLine & _
            "Where ((a.国籍 = [1] And a.民族 = [2] And a.性别 = [3] And a.姓名 = [4] And a.出生日期 = To_Date([5], 'YYYY-MM-DD')) Or" & vbNewLine & _
            "      a.身份证号 = [6]) And Not Exists (Select * From 病人身份关联 B Where b.病人id = a.病人id)" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.身份证号, a.家庭地址, a.国籍, a.民族, b.关联id, a.登记时间" & vbNewLine & _
            "From 病人信息 A, 病人身份关联 B, 病人身份关联 C, 病人信息 D" & vbNewLine & _
            "Where a.病人id = b.病人id And b.关联id = c.关联id And c.病人id = d.病人id And" & vbNewLine & _
            "      ((d.国籍 = [1] And d.民族 = [2] And d.性别 = [3] And d.姓名 = [4] And d.出生日期 = To_Date([5], 'YYYY-MM-DD')) Or" & vbNewLine & _
            "      d.身份证号 = [6])"
    On Error GoTo errH
    Set GetPatiSimilar = zlDatabase.OpenSQLRecord(strSQL, "GetPatiSimilar", (arrTmp(0)), (arrTmp(1)), (arrTmp(2)), (arrTmp(3)), (arrTmp(4)), (arrTmp(5)))
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiLinked(ByVal lngPatiID As Long) As ADODB.Recordset
'功能:通过病人ID加载已关联的病人
    Dim strSQL As String
    
    strSQL = "Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.身份证号, a.家庭地址, a.国籍, a.民族, 0 As 关联id, a.登记时间" & vbNewLine & _
            "From 病人信息 A" & vbNewLine & _
            "Where 病人id = [1] And Not Exists (Select * From 病人身份关联 Where 病人id = [1])" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.身份证号, a.家庭地址, a.国籍, a.民族, b.关联id, a.登记时间" & vbNewLine & _
            "From 病人信息 A, 病人身份关联 B, 病人身份关联 C" & vbNewLine & _
            "Where a.病人id = b.病人id And b.关联id = c.关联id And c.病人id = [1]"
    On Error GoTo errH
    Set GetPatiLinked = zlDatabase.OpenSQLRecord(strSQL, "GetPatiLinked", lngPatiID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AppendPatiSimilar(ByVal rsPati As ADODB.Recordset, ByVal bytFunc As Byte) As Boolean
'功能:追加相似病人
    Dim blnFind As Boolean
    Dim i As Long
    
    For i = 1 To rsPati.RecordCount
        mrsPati.Filter = "病人ID=" & rsPati!病人ID
        If mrsPati.RecordCount = 0 Then
            mrsPati.AddNew Array("病人ID", "关联ID", "姓名", "性别", "年龄", "出生日期", "身份证号", "家庭地址", "国籍", "民族", "登记时间", "EDIT"), _
            Array(rsPati!病人ID, rsPati!关联ID, rsPati!姓名, rsPati!性别, rsPati!年龄, rsPati!出生日期, rsPati!身份证号, rsPati!家庭地址, _
            rsPati!国籍, rsPati!民族, rsPati!登记时间, bytFunc)
            blnFind = True
        End If
        rsPati.MoveNext
    Next
    mrsPati.Filter = ""
    If Not blnFind And bytFunc = E_LINKAUTO Then
        MsgBox "未发现身份相似同时又未关联身份的病人信息！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    AppendPatiSimilar = True
End Function
                
