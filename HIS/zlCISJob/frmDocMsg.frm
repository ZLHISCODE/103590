VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmDocMsg 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药师审方聊天"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14400
   Icon            =   "frmDocMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   14400
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   12600
      Picture         =   "frmDocMsg.frx":6852
      ScaleHeight     =   330
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   405
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTmp 
      Height          =   915
      Left            =   14400
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   915
      _cx             =   1614
      _cy             =   1614
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      MouseIcon       =   "frmDocMsg.frx":6DDC
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   10000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDocMsg.frx":76B6
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1920
         ScaleHeight     =   240
         ScaleWidth      =   480
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   4320
      ScaleHeight     =   3735
      ScaleWidth      =   9735
      TabIndex        =   17
      Top             =   2160
      Width           =   9735
      Begin zl9CISJob.ucCommandBar cbsChat 
         Height          =   420
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   3015
         _extentx        =   5318
         _extenty        =   741
      End
      Begin VB.PictureBox picChat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   9735
         TabIndex        =   4
         Top             =   480
         Width           =   9735
         Begin VB.PictureBox pic气泡A 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   0
            Left            =   0
            ScaleHeight     =   855
            ScaleWidth      =   3375
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   3375
            Begin VB.TextBox txt气泡A 
               BackColor       =   &H0080FF80&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   960
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   19
               Text            =   "frmDocMsg.frx":7751
               Top             =   400
               Width           =   1335
            End
            Begin VB.Label lbl阅读 
               AutoSize        =   -1  'True
               Caption         =   "已读"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   0
               Left            =   2520
               TabIndex        =   21
               Top             =   600
               Width           =   360
            End
            Begin VB.Label lbl气泡A 
               AutoSize        =   -1  'True
               Caption         =   "管理员  2019-07-31 22:12:03"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   0
               Left            =   840
               TabIndex        =   20
               Top             =   30
               Width           =   2430
            End
            Begin VB.Image img气泡A 
               Appearance      =   0  'Flat
               Height          =   720
               Index           =   0
               Left            =   50
               Picture         =   "frmDocMsg.frx":775E
               Stretch         =   -1  'True
               Top             =   0
               Width           =   720
            End
            Begin VB.Shape shp气泡A 
               BackColor       =   &H0080FF80&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H8000000F&
               DrawMode        =   9  'Not Mask Pen
               FillColor       =   &H0080FF80&
               FillStyle       =   0  'Solid
               Height          =   495
               Index           =   0
               Left            =   840
               Shape           =   4  'Rounded Rectangle
               Top             =   285
               Width           =   1575
            End
         End
      End
      Begin VB.VScrollBar vsBar 
         Height          =   7575
         LargeChange     =   4
         Left            =   9480
         SmallChange     =   4
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8055
      ScaleWidth      =   4095
      TabIndex        =   12
      Top             =   360
      Width           =   4095
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2580
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   3360
         _Version        =   589884
         _ExtentX        =   5927
         _ExtentY        =   4551
         _StockProps     =   0
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   1305
      End
      Begin zl9CISJob.ucCommandBar cbsList 
         Height          =   420
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   3975
         _extentx        =   7011
         _extenty        =   741
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "发送(&S)"
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox picAdvice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   4320
      ScaleHeight     =   1695
      ScaleWidth      =   9735
      TabIndex        =   14
      Top             =   360
      Width           =   9735
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   7275
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   12645
         _cx             =   22304
         _cy             =   12832
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "frmDocMsg.frx":8628
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDocMsg.frx":8F02
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   1
         BackColorFrozen =   14737632
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   24
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin zl9CISJob.ucCommandBar cbsAdvice 
         Height          =   420
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   3975
         _extentx        =   7011
         _extenty        =   741
      End
   End
   Begin VB.PictureBox picIn 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   4320
      ScaleHeight     =   2175
      ScaleWidth      =   9735
      TabIndex        =   13
      Top             =   6240
      Width           =   9735
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   0
         Left            =   7200
         ScaleHeight     =   345
         ScaleWidth      =   1095
         TabIndex        =   9
         Top             =   1680
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "关闭(C)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   230
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   90
            Width           =   705
         End
      End
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   1
         Left            =   8520
         ScaleHeight     =   345
         ScaleWidth      =   1095
         TabIndex        =   10
         Top             =   1680
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "发送(S)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   90
            Width           =   705
         End
      End
      Begin VB.TextBox txtIn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   9735
      End
   End
   Begin VB.Timer TmrIcon 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   13680
      Top             =   240
   End
   Begin VB.PictureBox PicNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   13200
      Picture         =   "frmDocMsg.frx":8F9D
      ScaleHeight     =   330
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   14040
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":9867
            Key             =   "PatiIn"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":9E01
            Key             =   "Meet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":A39B
            Key             =   "Msg"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":10BFD
            Key             =   "PatiOut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":114D7
            Key             =   "msgno"
         EndProperty
      EndProperty
   End
   Begin zlSubclass.Subclass Subclass 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Image img气泡B 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   14280
      Picture         =   "frmDocMsg.frx":11A71
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   720
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   11520
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDocMsg.frx":1293B
   End
End
Attribute VB_Name = "frmDocMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'----------------------------------------------------------------------------------------------------
'-----系统托盘相关声明
'----------------------------------------------------------------------------------------------------
Private Const MAX_TOOLTIP As Integer = 64
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_MOUSEWHEEL = &H20A          '鼠标滚动
Private Const SW_RESTORE = 9
Private Const conCOLOR_BULELIGHT As Long = &HE4B440
Private Const conCOLOR_BULE As Long = &HD48A00

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type

Private Enum colList
    col_消息状态 = 0
    col_发送人 = 1
    col_发送时间 = 2
    col_发送信息 = 3
    col_病人信息 = 4
    col_来源 = 5
    col_姓名 = 6
    col_性别 = 7
    col_年龄 = 8
    COL_标识号 = 9
    col_科室 = 10
    
    '隐藏列
    col_科室ID = 11
    col_病人Id = 12
    col_就诊ID = 13
    COL_医嘱IDs = 14
    col_医嘱简介 = 15
    col_会话状态 = 16
    col_未读ID = 17
    col_会话ID = 18
    col_未读时间 = 19
End Enum

Private Enum COL用药清单
    '隐藏列
    COLB_ID = 1
    COLB_组号 = 2
    COLB_用药来源 = 3
    COLB_诊疗项目ID = 4
    COLB_收费细目ID = 5
    COLB_频率间隔 = 6
    COLB_间隔单位 = 7
    COLB_用法id = 8
    COLB_煎法id = 9
    COLB_终止时间 = 10
    '可见列
    COLB_期效 = 11
    COLB_开始时间 = 12
    colB_药品类别 = 13
    colB_用药内容 = 14
    COLB_用法 = 15
    COLB_单次用量 = 16
    COLB_总给予量 = 17
    COLB_天数 = 18
    COLB_执行频次 = 19
    
    '隐藏列
    COLB_频率次数 = 20
    COLB_配方数据 = 21
End Enum


Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

Private mnfIconData As NOTIFYICONDATA
Private mblnIconShow As Boolean
Private mintPreTime As Integer
Public isUnload As Boolean
Private mfrmParent As Object
Private mintType As Integer
Private mdtBegin As Date, mdtEnd As Date
Private mstr未读会话ids As String
Private mbln消息屏蔽 As Boolean
Private WithEvents mclsNotice As clsNotice
Attribute mclsNotice.VB_VarHelpID = -1


Public Sub SetNotifyIcon(ByVal intType As Integer, ByVal strMsg As String)
    'intType 0-初始化  1-消息 2-闪烁 3-还原
    'strMsg
    On Error Resume Next
    '下面的代码可以将图标添加到系统图标
    If intType = 0 And mnfIconData.hwnd <> 0 Then Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
    mnfIconData.hwnd = Me.hwnd
    mnfIconData.uID = picMsg.Picture '这里确定使用哪个图标
    mnfIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    mnfIconData.uCallbackMessage = WM_MOUSEMOVE
    mnfIconData.hIcon = IIf(intType = 2, PicNo.Picture.Handle, picMsg.Picture.Handle)
    mnfIconData.szTip = strMsg & vbNullChar  '这里是将鼠标移到图标上时，将显示的文字
    mnfIconData.cbSize = Len(mnfIconData)
    Call Shell_NotifyIcon(IIf(intType = 0, NIM_ADD, NIM_MODIFY), mnfIconData)
End Sub

Private Sub StartMsg()
    TmrIcon.Enabled = Not TmrIcon.Enabled
    If TmrIcon.Enabled = False Then Call SetNotifyIcon(3, IIf(mintType = 1, "门诊审方聊天", "住院审方聊天") & vbCrLf & "当前用户：" & UserInfo.姓名)
End Sub


Private Sub cbsList_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 4 '刷新
        Call LoadMsg
        Call ClearChat
        rptList.Tag = ""
    End Select
End Sub


Private Sub Form_Load()
    On Error GoTo errH
    isUnload = False
    '隐藏快捷按钮
    cmdSend.Top = -1000
    cmdClose.Top = -1000
    
    '聊天控件颜色
    Me.BackColor = RGB(247, 247, 247)
    picChat.BackColor = RGB(247, 247, 247)
    picBack.BackColor = RGB(247, 247, 247)
    pic气泡A(0).BackColor = RGB(247, 247, 247)
    lbl气泡A(0).BackColor = RGB(247, 247, 247)
    lbl阅读(0).BackColor = RGB(247, 247, 247)
    txt气泡A(0).BackColor = RGB(129, 246, 129)
    picIn.BackColor = RGB(247, 247, 247)
    
    '滚轮事件初始化
    Subclass.hwnd = Me.hwnd
    Subclass.Messages(WM_MOUSEWHEEL) = True
    
    '菜单初始化
    Call InitCommandBarList
    
    Call InitAdviceTable
    
    Call InitDockPannel '开始拖动布局初始化
    
    Call InitReportColumn
    
    Call SetNotifyIcon(0, IIf(mintType = 1, "门诊审方聊天", "住院审方聊天") & vbCrLf & "当前用户：" & UserInfo.姓名)
    
    mstr未读会话ids = ""
    
    Call LoadMsg

    Call ClearChat

    '消息初始化
    Set mclsNotice = zl9ComLib.GetClsNotice
    
    '检查DCN配置是否启用
    If Not mclsNotice Is Nothing Then
        If mclsNotice.CheckDcnEnable(3) = False Then
            Set mclsNotice = Nothing
        End If
    End If
    
    Me.Caption = IIf(mintType = 1, "门诊审方聊天", "住院审方聊天")
    Call RestoreWinState(Me, App.ProductName)
    rptList.AllowColumnSort = False
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    
    If isUnload = False Then
        Cancel = 1
        If txtIn.Text <> "" Then
            If MsgBox("输入区存在待发出的消息，确认关闭会话窗口?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
        End If
        Me.Hide
        Exit Sub
    End If
    Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
    Subclass.Messages(WM_MOUSEWHEEL) = False
    '卸载消息对象
    If Not mclsNotice Is Nothing Then
        Set mclsNotice = Nothing
    End If
    
    Call SaveWinState(Me, App.ProductName)
    Set mfrmParent = Nothing
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngMsg As Long
    
On Error GoTo errH
    lngMsg = X / Screen.TwipsPerPixelX
    If lngMsg = WM_LBUTTONDBLCLK Then
        If IsWindowEnabled(mfrmParent.hwnd) Then
           Me.Hide
           ShowWindow Me.hwnd, SW_RESTORE
           Call picBack_Resize
            vsBar.Value = vsBar.Max '自动定位到最后
            If rptList.SelectedRows.Count > 0 Then
                If InStr(mstr未读会话ids & ",", "," & rptList.SelectedRows(0).Record(col_会话ID).Value & ",") > 0 Then
                    Call ReadMsg(Val(rptList.SelectedRows(0).Record(col_会话ID).Value & ""), rptList.SelectedRows(0).Record(col_发送人).Value & "")
                End If
            End If
        End If
    End If
    Exit Sub
errH:
    Err.Clear
End Sub

Private Sub picAdvice_Resize()
    On Error Resume Next
    cbsAdvice.Width = picAdvice.Width
    vsAdvice.Top = cbsAdvice.Top + cbsAdvice.Height + 10
    vsAdvice.Width = picAdvice.Width
    vsAdvice.Height = picAdvice.Height - vsAdvice.Top
End Sub

Private Sub picBack_Resize()
    On Error Resume Next
    cbsChat.Width = picBack.Width
    picChat.Height = picBack.Height - cbsChat.Top
    picChat.Width = picBack.Width - 300
    
    vsBar.Top = cbsChat.Height: vsBar.Height = picBack.Height - cbsChat.Height
    vsBar.Left = picBack.Width - vsBar.Width
End Sub

Private Sub picChat_Resize()
    On Error Resume Next
    Call CtlResize(1)
End Sub

Private Sub picIn_Resize()
    On Error Resume Next
    txtIn.Width = picIn.Width - 80
    txtIn.Height = picIn.Height - 600
    
    picBtn(1).Top = txtIn.Height + (picIn.Height - txtIn.Height - picBtn(1).Height) / 2
    picBtn(0).Top = picBtn(1).Top
    
    picBtn(1).Left = picIn.Width - picBtn(1).Width - 100
    picBtn(0).Left = picBtn(1).Left - picBtn(0).Width - 200
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    cbsList.Width = picList.Width
    
    rptList.Left = 0
    rptList.Top = cbsList.Top + cbsList.Height
    rptList.Width = picList.Width
    rptList.Height = picList.Height - rptList.Top
End Sub




Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    '尝试定位病人
    On Error Resume Next
    If rptList.SelectedRows.Count = 0 Then Exit Sub

    If Not mfrmParent Is Nothing And Val(rptList.SelectedRows(0).Record(col_病人Id).Value) <> 0 Then
        Call mfrmParent.LocateMsgPati(Val(rptList.SelectedRows(0).Record(col_病人Id).Value), Val(rptList.SelectedRows(0).Record(col_就诊ID).Value), Val(Split(rptList.SelectedRows(0).Record(COL_医嘱IDs).Value, ",")(0)))
    End If
End Sub

Private Sub rptList_SelectionChanged()
    On Error GoTo errH
    If rptList.SelectedRows.Count = 0 Then Exit Sub          '非正常情况
    If rptList.SelectedRows.Count > 0 Then
        If Val(rptList.SelectedRows(0).Record(col_会话状态).Value & "") = 1 Then
            rptList.PaintManager.HighlightForeColor = &H808080
        Else
            rptList.PaintManager.HighlightForeColor = vbBlack
        End If
    End If
    
    If Val(rptList.Tag) = Val(rptList.SelectedRows(0).Record(col_会话ID).Value) Then Exit Sub
    rptList.Tag = Val(rptList.SelectedRows(0).Record(col_会话ID).Value)

    cbsAdvice.FindControl(999).Caption = "病人信息"
    cbsAdvice.RefreshCtl
    
    cbsChat.FindControl(0).Caption = "请选择聊天会话"
    cbsChat.RefreshCtl
    
    LoadChat
    
    vsBar.Value = vsBar.Max '自动定位到最后
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearChat()
    On Error GoTo errH
    cbsAdvice.FindControl(999).Caption = "病人信息"
    cbsAdvice.RefreshCtl
    
    cbsChat.FindControl(0).Caption = "请选择聊天会话"
    cbsChat.RefreshCtl
    
    '清空处方信息
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Cell(flexcpBackColor, vsAdvice.FixedRows, colB_药品类别, vsAdvice.Rows - 1, colB_药品类别) = &H8000000B      '灰蓝色
    vsAdvice.Cell(flexcpBackColor, vsAdvice.FixedRows, COLB_用法, vsAdvice.Rows - 1, COLB_用法) = &H8000000B      '灰蓝色
    vsAdvice.Cell(flexcpBackColor, vsAdvice.FixedRows, 0, vsAdvice.Rows - 1, 0) = &H8000000B
    
    
    '清空聊天控件
    Call SetCtl(0) '卸载控件
    picChat.Visible = False
    
    '清空聊天内容
    txtIn.Text = ""
    
     Call CtlResize(1)
     
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub TmrIcon_Timer()
    If Replace(mstr未读会话ids, ",", "") = "" Then
        StartMsg
        Exit Sub
    End If
    Call SetNotifyIcon(IIf(mblnIconShow, 1, 2), IIf(mintType = 1, "门诊审方聊天", "住院审方聊天") & vbCrLf & "当前用户：" & UserInfo.姓名 & vbCrLf & "当前有" & UBound(Split(Mid(mstr未读会话ids, 2), ",")) + 1 & "个新的会话信息")
    mblnIconShow = Not mblnIconShow
End Sub


'发送和关闭控件
Private Sub cmdClose_Click()
    lblBtn_Click 0
End Sub

Private Sub cmdSend_Click()
    lblBtn_Click 1
End Sub

Private Sub picBtn_Click(Index As Integer)
    lblBtn_Click Index
End Sub

Private Sub picBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtn(Index).BackColor = conCOLOR_BULELIGHT
End Sub

Private Sub picBtn_Resize(Index As Integer)
    lblBtn(Index).Move picBtn(Index).ScaleWidth / 2 + lblBtn(Index).Width / 2, picBtn(Index).ScaleHeight / 2 - lblBtn(Index) / 2
End Sub

Private Sub lblBtn_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
    Else
        Call SendMsg
    End If
End Sub

Private Sub picIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picBtn(0).BackColor <> conCOLOR_BULE And picBtn(0).Enabled Then picBtn(0).BackColor = conCOLOR_BULE
    If picBtn(1).BackColor <> conCOLOR_BULE And picBtn(1).Enabled Then picBtn(1).BackColor = conCOLOR_BULE
End Sub


Private Function InitCommandBarList() As Boolean
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim curDate As Date
    
    On Error GoTo errH
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsList.ObjCommandBar)
    With cbsList.ObjCommandBar

        Set .Icons = imgList.Icons

        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        .ActiveMenuBar.Visible = False
        .Options.LargeIcons = False

        '消息工具栏
        '------------------------------------------------------------------------------------------------------------------
        Set objBar = .Add("缺省", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched
        Set objControl = NewToolBar(objBar, xtpControlLabel, 1, "消息列表")
        objControl.IconId = 5
        
        Set objControl = objBar.Controls.Add(xtpControlLabel, 999, "时间")   '医嘱时间
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = objBar.Controls.Add(xtpControlCustom, 3, "时间")
            objCustom.Handle = cboTime.hwnd
            objCustom.Flags = xtpFlagRightAlign

        
        Set objControl = NewToolBar(objBar, xtpControlButton, 4, "", , , xtpButtonIcon)
        objControl.IconId = 2
        objControl.Flags = xtpFlagRightAlign


    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsAdvice.ObjCommandBar)
    With cbsAdvice.ObjCommandBar

        Set .Icons = imgList.Icons

        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        .ActiveMenuBar.Visible = False
        .Options.LargeIcons = False

        '病人信息工具栏
        '------------------------------------------------------------------------------------------------------------------
        Set objBar = .Add("缺省", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched
        Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "处方信息")
        objControl.IconId = 4
        
        Set objControl = objBar.Controls.Add(xtpControlLabel, 999, "病人信息")   '医嘱时间
        objControl.Flags = xtpFlagRightAlign

    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsChat.ObjCommandBar)
    With cbsChat.ObjCommandBar

        Set .Icons = imgList.Icons

        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        .ActiveMenuBar.Visible = False
        .Options.LargeIcons = False

        '病人信息工具栏
        '------------------------------------------------------------------------------------------------------------------
        Set objBar = .Add("缺省", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched
        Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "请选择聊天会话")
        objControl.IconId = 1

    End With
    
    
    
    '缺省医嘱时间
    cboTime.Clear
    cboTime.AddItem "所有"
    cboTime.AddItem "今天"
    cboTime.AddItem "昨天"
    cboTime.AddItem "最近三天"
    cboTime.AddItem "最近一周"
    cboTime.AddItem "最近两周"
    cboTime.AddItem "最近一月"
    cboTime.AddItem "[指定..]"
    Call zlControl.CboSetIndex(cboTime.hwnd, 6)
    curDate = zlDatabase.Currentdate
    mdtBegin = Format(curDate - 30, "yyyy-MM-dd 00:00:00")
    mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")

    mintPreTime = 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    On Error GoTo errH
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = imgList.Icons
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    On Error GoTo errH
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    On Error GoTo errH
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = True '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    dkpMain.Options.LockSplitters = True
    DockPannelInit = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'InitDockPannel初始区域划分
Private Sub InitDockPannel()
    Dim objPane As Pane
    On Error GoTo errH
    Set objPane = dkpMain.CreatePane(1, 270, 500, DockLeftOf, objPane)
    objPane.Title = "消息列表"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 500, 150, DockRightOf, objPane)
    objPane.Title = "病人信息"
    objPane.Options = PaneNoCaption
'
    Set objPane = dkpMain.CreatePane(3, 500, 500, DockBottomOf, objPane)
    objPane.Title = "消息内容"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(4, 500, 160, DockBottomOf, objPane)
    objPane.Title = "消息发送"
    objPane.Options = PaneNoCaption

    Call DockPannelInit(dkpMain)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'绑定布局控件
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
  On Error GoTo errH
    Select Case Item.ID
        Case 1
            Item.Handle = picList.hwnd
        Case 2
            Item.Handle = picAdvice.hwnd
        Case 3
            Item.Handle = picBack.hwnd
        Case 4
            Item.Handle = picIn.hwnd
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub InitReportColumn()
    Dim objCol As ReportColumn
    
    On Error GoTo errH
    
    With rptList
        Set objCol = .Columns.Add(col_消息状态, "", 18, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("msgno").Index - 1

        Set objCol = .Columns.Add(col_发送人, "发送药师", 0, False)
        Set objCol = .Columns.Add(col_发送时间, "会话时间", 0, False)
        Set objCol = .Columns.Add(col_发送信息, "发送信息", 190, True)
        Set objCol = .Columns.Add(col_病人信息, "病人信息", 120, True)
        Set objCol = .Columns.Add(col_来源, "来源", 0, False)
        Set objCol = .Columns.Add(col_姓名, "病人姓名", 0, False)
        Set objCol = .Columns.Add(col_性别, "性别", 0, False)
        Set objCol = .Columns.Add(col_年龄, "年龄", 0, False)
        Set objCol = .Columns.Add(COL_标识号, "标识号", 0, False)
        Set objCol = .Columns.Add(col_科室, "科室", 0, False)
        Set objCol = .Columns.Add(col_科室ID, "科室ID", 0, False)
        Set objCol = .Columns.Add(col_病人Id, "病人ID", 0, False)
        Set objCol = .Columns.Add(col_就诊ID, "就诊ID", 0, False)
        Set objCol = .Columns.Add(COL_医嘱IDs, "医嘱IDs", 0, False)
        Set objCol = .Columns.Add(col_医嘱简介, "医嘱简介", 0, False)
        Set objCol = .Columns.Add(col_会话状态, "会话状态", 0, False)
        Set objCol = .Columns.Add(col_未读ID, "未读ID", 0, False)
        Set objCol = .Columns.Add(col_会话ID, "会话ID", 0, False)
        Set objCol = .Columns.Add(col_未读时间, "未读时间", 0, False)

        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnShaded
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridNoLines
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的审方消息..."
            .HighlightBackColor = &HFFEDCA
            .HighlightForeColor = vbBlack
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    On Error GoTo errH
    If cboTime.ListIndex = mintPreTime And mintPreTime <> 7 Then Exit Sub
    
    curDate = zlDatabase.Currentdate
    
    Select Case cboTime.Text
    Case "所有"
        mdtBegin = CDate(0)
        mdtEnd = CDate(0)
    Case "今天"
        mdtBegin = Format(curDate, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "昨天"
        mdtBegin = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
    Case "最近三天"
        mdtBegin = Format(curDate - 2, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近一周"
        mdtBegin = Format(curDate - 7, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近两周"
        mdtBegin = Format(curDate - 14, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近一月"
        mdtBegin = Format(curDate - 30, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "[指定..]"
        If Not frmSelectTime.ShowMe(Me, mdtBegin, mdtEnd, cboTime) Then
            '取消时恢复原来的选择
            Call zlControl.CboSetIndex(cboTime.hwnd, mintPreTime)
            rptList.SetFocus
            Exit Sub
        Else
            rptList.SetFocus
        End If
    End Select
        
    If mdtBegin = CDate(0) Or mdtEnd = CDate(0) Then
        cboTime.ToolTipText = ""
    Else
        cboTime.ToolTipText = "范围：" & Format(mdtBegin, "yyyy-MM-dd HH:mm:ss") & " 至 " & Format(mdtEnd, "yyyy-MM-dd HH:mm:ss")
    End If
    mintPreTime = cboTime.ListIndex
    
    Call LoadMsg
    Call ClearChat
    rptList.Tag = ""
    Me.Refresh
    
    rptList.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadMsg()
    Dim strSQL As String, strFilter As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim j As Long, i As Long
    
    On Error GoTo errH
    If cboTime.Text <> "所有" Then
        strFilter = " AND (m.阅读时间 IS NULL Or m.发送时间 Between [2] And [3]) "
    End If
    
    If mintType = 1 Then '门诊
        strSQL = "Select b.Id, b.对象标识, b.对象内容, b.就诊id, b.创建人, b.创建时间, b.状态, '门诊' As 来源, a.病人id, a.姓名, a.性别, a.年龄, a.执行部门id As 科室," & vbNewLine & _
            "              a.执行时间 As 操作时间, a.门诊号 As 标识号, m.Id As 消息id, m.发送时间 As 发送时间, m.阅读时间,decode(m.阅读时间,null,0,1) as 阅读状态" & vbNewLine & _
            "       From 病人挂号记录 A, 聊天会话表 B, 聊天信息表 M" & vbNewLine & _
            "       Where a.病人id = b.病人id And a.Id = b.就诊id And b.Id = m.会话id(+) And b.接收人 = [1] And b.病人来源 = 1 And m.接收人 = [1]" & strFilter
    Else
        strSQL = " Select b.Id, b.对象标识, b.对象内容, b.就诊id, b.创建人, b.创建时间, b.状态, '住院' As 来源, a.病人id, a.姓名, a.性别, a.年龄, a.出院科室id As 科室," & vbNewLine & _
            "              a.入院日期 As 操作时间, a.住院号 As 标识号, m.Id As 消息id, m.发送时间 As 发送时间, m.阅读时间,decode(m.阅读时间,null,0,1) as 阅读状态" & vbNewLine & _
            "       From 病案主页 A, 聊天会话表 B, 聊天信息表 M" & vbNewLine & _
            "       Where a.病人id = b.病人id And a.主页id = b.就诊id And b.Id = m.会话id(+) And b.接收人 =[1] And b.病人来源 = 2  And m.接收人 = [1]" & strFilter
    End If

    strSQL = "Select d.Id, d.对象标识, d.对象内容, d.就诊id, d.创建人, d.创建时间, d.状态, d.来源, d.病人id, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.科室 As 科室id," & vbNewLine & _
            "       d.操作时间, d.标识号, Max(d.消息id) As 未读id, Max(d.发送时间) As 消息时间," & vbNewLine & _
            "       Min(Nvl(d.阅读时间, To_Date('1900-01-01', 'yyyy-mm-dd'))) As 是否未读,min(d.阅读状态)" & vbNewLine & _
            "From (" & strSQL & ") D, 部门表 G" & vbNewLine & _
            "Where d.科室 =g.Id " & vbNewLine & _
            "Group By d.Id, d.对象标识, d.对象内容, d.就诊id, d.创建人, d.创建时间, d.状态, d.来源, d.病人id, d.姓名, d.性别, d.年龄, g.名称, d.科室, d.操作时间, d.标识号" & vbNewLine & _
            "Order By min(d.阅读状态),Max(d.消息id) desc,d.状态"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, mdtBegin, mdtEnd)
    
    rptList.Records.DeleteAll

    With rptList
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                Set objRecord = .Records.Add()
                Set objItem = objRecord.AddItem("")
                    If Format(rsTmp!是否未读 & "", "yyyy-MM-dd") = "1900-01-01" Then
                        objItem.Icon = img16.ListImages("Msg").Index - 1
                    End If
                Set objItem = objRecord.AddItem(rsTmp!创建人 & "")
                Set objItem = objRecord.AddItem(Format(rsTmp!消息时间 & "", "yyyy-MM-dd hh:mm"))
                
                Set objItem = objRecord.AddItem(rsTmp!创建人 & "  " & Format(rsTmp!消息时间 & "", "yyyy-MM-dd hh:mm"))
                objItem.Bold = True
                objItem.ForeColor = vbRed
                objItem.Icon = img16.ListImages.Item("Meet").Index - 1
                
                Set objItem = objRecord.AddItem(rsTmp!姓名 & "  " & rsTmp!性别 & "  " & rsTmp!年龄 & "  " & rsTmp!科室)
                Set objItem = objRecord.AddItem(rsTmp!来源 & "")
                Set objItem = objRecord.AddItem(rsTmp!姓名 & "")
                Set objItem = objRecord.AddItem(rsTmp!性别 & "")
                Set objItem = objRecord.AddItem(rsTmp!年龄 & "")
                Set objItem = objRecord.AddItem(rsTmp!标识号 & "")
                Set objItem = objRecord.AddItem(rsTmp!科室 & "")

                Set objItem = objRecord.AddItem(rsTmp!科室ID & "")
                Set objItem = objRecord.AddItem(rsTmp!病人ID & "")
                Set objItem = objRecord.AddItem(rsTmp!就诊id & "")
                Set objItem = objRecord.AddItem(rsTmp!对象标识 & "")
                Set objItem = objRecord.AddItem(rsTmp!对象内容 & "")
                Set objItem = objRecord.AddItem(rsTmp!状态 & "")
                Set objItem = objRecord.AddItem(rsTmp!未读ID & "")
                Set objItem = objRecord.AddItem(rsTmp!ID & "")
                Set objItem = objRecord.AddItem(Format(rsTmp!是否未读 & "", "yyyy-MM-dd"))
  
                objRecord.PreviewText = rsTmp!对象内容 & ""

                '已完成的会诊病人用灰色显示
                If Val(rsTmp!状态 & "") = 1 Then
                    For j = 0 To rptList.Columns.Count - 1
                        objRecord.Item(j).ForeColor = &H808080
                    Next
                End If

                rsTmp.MoveNext
            Loop
      
        End If
        .Populate
        
        '初始化未读会话ID
        For i = 0 To rptList.Rows.Count - 1
            With rptList.Rows(i)
                If Not .GroupRow Then
                    If .Record(col_未读时间).Value = "1900-01-01" Then
                        If InStr("," & mstr未读会话ids & ",", "," & Val(.Record(col_会话ID).Value & "") & ",") = 0 Then
                            mstr未读会话ids = mstr未读会话ids & "," & Val(.Record(col_会话ID).Value & "")
                        End If
                    End If
                End If
            End With
        Next
        
        If mstr未读会话ids <> "" And TmrIcon.Enabled = False Then Call StartMsg
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadChat()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng气泡 As Long
    Dim strMsg As String
    
    On Error GoTo errH
    '边界值处理
    Call SetCtl(0) '卸载控件
    If rptList.SelectedRows.Count < 1 Then Exit Sub
    If Val(rptList.SelectedRows(0).Record(col_会话ID).Value & "") = 0 Then Exit Sub
    
    picChat.Visible = True
    '显示病人信息
    cbsAdvice.FindControl(999).Caption = "当前病人：" & rptList.SelectedRows(0).Record(col_姓名).Value & " 性别：" & rptList.SelectedRows(0).Record(col_性别).Value & _
  " 年龄：" & rptList.SelectedRows(0).Record(col_年龄).Value & " 科室：" & rptList.SelectedRows(0).Record(col_科室).Value & IIf(rptList.SelectedRows(0).Record(col_来源).Value = "门诊", " 门诊号", " 住院号") & "：" & rptList.SelectedRows(0).Record(COL_标识号).Value
    cbsAdvice.RefreshCtl
    
    '显示聊天人信息
    cbsChat.FindControl(0).Caption = rptList.SelectedRows(0).Record(col_发送人).Value
    cbsChat.RefreshCtl
    
    '显示处方信息
    Call GetAdvice
    
    strSQL = "select a.id,A.会话ID,A.发送人,A.发送内容,A.发送时间,A.接收人,A.阅读时间 from 聊天信息表 A WHERE A.会话ID=[1] order by A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rptList.SelectedRows(0).Record(col_会话ID).Value & ""))
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            
            '药师发送
                '气泡容器加载
               lng气泡 = lng气泡 + 1
               Load pic气泡A(lng气泡)
               Set pic气泡A(lng气泡).Container = picChat
               pic气泡A(lng气泡).Tag = rsTmp!发送人 & ""
               
               '头像显示
               Load img气泡A(lng气泡)
               Set img气泡A(lng气泡).Container = pic气泡A(lng气泡)
               img气泡A(lng气泡).Tag = rsTmp!ID & ""

               '加载气泡
               Load shp气泡A(lng气泡)
               Set shp气泡A(lng气泡).Container = pic气泡A(lng气泡)

               Load txt气泡A(lng气泡)
               Set txt气泡A(lng气泡).Container = pic气泡A(lng气泡)
               
               
               '处理换行格式不正常的问题
               strMsg = Replace(rsTmp!发送内容 & "", vbCrLf, "[换行处理]")
               strMsg = Replace(strMsg, Chr(10), vbCrLf)
               strMsg = Replace(strMsg, "[换行处理]", vbCrLf)
               txt气泡A(lng气泡).Text = strMsg
               
               '加载气泡信息
               Load lbl气泡A(lng气泡)
               Set lbl气泡A(lng气泡).Container = pic气泡A(lng气泡)

               '加载阅读信息
               Load lbl阅读(lng气泡)
               Set lbl阅读(lng气泡).Container = pic气泡A(lng气泡)
               lbl阅读(lng气泡).Caption = "已读"


               If rsTmp!发送人 & "" = UserInfo.姓名 Then
                    Set img气泡A(lng气泡).Picture = img气泡B.Picture
                    
                    shp气泡A(lng气泡).BackColor = RGB(221, 235, 255)
                    shp气泡A(lng气泡).FillColor = RGB(221, 235, 255)
                    txt气泡A(lng气泡).BackColor = RGB(208, 224, 240)
                    lbl气泡A(lng气泡).Caption = UserInfo.姓名 & "  " & Format(rsTmp!发送时间 & "", "yyyy-MM-dd HH:mm")
                    lbl阅读(lng气泡).Visible = rsTmp!阅读时间 & "" <> ""
                Else
                    lbl气泡A(lng气泡).Caption = rsTmp!发送人 & "  " & Format(rsTmp!发送时间 & "", "yyyy-MM-dd HH:mm")
                    lbl阅读(lng气泡).Visible = False
               End If
            rsTmp.MoveNext
        Loop
    End If
    Call CtlResize(1)
    Call SetCtl(1) '显示控件
    If rptList.SelectedRows(0).Record(col_未读时间).Value = "1900-01-01" Then
        Call ReadMsg(Val(rptList.SelectedRows(0).Record(col_会话ID).Value & ""), rptList.SelectedRows(0).Record(col_发送人).Value & "")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReadMsg(ByVal Lng会话ID As Long, str发送人 As String)
    '标记消息已读
    Dim strSQL As String

    On Error GoTo errH
    If Lng会话ID = 0 Then Exit Sub
    strSQL = "Zl_聊天信息表_Edit(2," & Lng会话ID & ",null,'" & str发送人 & "')"
    
    If InStr("," & mstr未读会话ids & ",", "," & Lng会话ID & ",") > 0 Then
        mstr未读会话ids = Replace(mstr未读会话ids & ",", "," & Lng会话ID & ",", "")
        If mstr未读会话ids <> "" Then mstr未读会话ids = "," & mstr未读会话ids
    End If
    
    rptList.SelectedRows(0).Record(col_消息状态).Icon = 9999
    rptList.SelectedRows(0).Record(col_未读时间).Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetCtl(intType As Integer)
    '卸载界面的控件
    'intType -0 卸载控件 -1显示控件
    Dim obj As Object
    Dim obj数组 As Object
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To 6
        Select Case i
                Case 1
                    Set obj数组 = shp气泡A
                Case 2
                    Set obj数组 = img气泡A
                Case 3
                    Set obj数组 = lbl气泡A
                Case 4
                    Set obj数组 = txt气泡A
                Case 5
                    Set obj数组 = lbl阅读
                Case 6
                    Set obj数组 = pic气泡A
        End Select
        For Each obj In obj数组
            If obj.Index <> 0 Then
                Select Case intType
                    Case 0
                        Unload obj
                    Case 1
                        If i <> 5 Then
                            obj.Visible = True
                        End If
                End Select
                
            End If
        Next
    Next
End Sub


Private Sub CtlResize(lngMax As Long)
    Dim i As Long
    Dim lngTxtW As Long, lngTxtH As Long
    Dim lng最大气泡宽度 As Long
    
    On Error Resume Next
    
    If picChat.Visible = False Then
        picChat.Height = picBack.Height - cbsChat.Height
        vsBar.Visible = False
        vsBar.Max = (picChat.Height - picBack.Height + cbsChat.Height) / 100
        Exit Sub
    End If
    
    lng最大气泡宽度 = IIf(picChat.Width / 2 - 880 > 7000, 7000, picChat.Width / 2 - 880) '获取最大的标签宽度用于显示文本框
    
    For i = lngMax To pic气泡A.Count - 1
        If i = 1 Then
            pic气泡A(i).Top = 100: pic气泡A(i).Left = 0: pic气泡A(i).Width = picChat.Width
        Else
            pic气泡A(i).Top = pic气泡A(i - 1).Top + pic气泡A(i - 1).Height + 100: pic气泡A(i).Left = 0: pic气泡A(i).Width = picChat.Width
        End If
        
        lngTxtW = 0: lngTxtH = 0
        If pic气泡A(i).Tag = UserInfo.姓名 Then
            img气泡A(i).Top = 0: img气泡A(i).Left = picChat.Width - img气泡A(i).Width - 50
            lbl气泡A(i).Top = 30: lbl气泡A(i).Left = picChat.Width - lbl气泡A(i).Width - 840

            Call GetTextHight(txt气泡A(i).Text, lngTxtW, lngTxtH)
            If lngTxtH = 0 And lngTxtW <> 0 Then
                txt气泡A(i).Width = lngTxtW - 100
                txt气泡A(i).Height = 330
            ElseIf lngTxtH <> 0 And lngTxtW = 0 Then
                txt气泡A(i).Width = lng最大气泡宽度 - 100
                txt气泡A(i).Height = lngTxtH
            Else
                txt气泡A(i).Width = lngTxtW - 100
                txt气泡A(i).Height = lngTxtH
            End If
            
            txt气泡A(i).Top = 480: txt气泡A(i).Left = picChat.Width - txt气泡A(i).Width - 960
            If txt气泡A(i).Width > 4700 And txt气泡A(i).Height > 1800 Then
                txt气泡A(i).Top = txt气泡A(i).Top + 100
                txt气泡A(i).Left = txt气泡A(i).Left - 100
            End If
            
            shp气泡A(i).Width = txt气泡A(i).Width + 240 + IIf(txt气泡A(i).Width > 4700 And txt气泡A(i).Height > 1800, 220, 120)
            shp气泡A(i).Height = txt气泡A(i).Top + txt气泡A(i).Height + IIf(txt气泡A(i).Width > 4700 And txt气泡A(i).Height > 1800, 350, IIf(txt气泡A(i).Height < 500, 120, 180)) - shp气泡A(i).Top
            shp气泡A(i).Top = 285: shp气泡A(i).Left = picChat.Width - shp气泡A(i).Width - 840

            
            pic气泡A(i).Height = shp气泡A(i).Top + shp气泡A(i).Height + 75
            lbl阅读(i).Top = shp气泡A(i).Top + shp气泡A(i).Height - lbl阅读(i).Height - 15
            lbl阅读(i).Left = shp气泡A(i).Left - lbl阅读(i).Width - 100
        Else
            img气泡A(i).Top = 0: img气泡A(i).Left = 50
            lbl气泡A(i).Top = 30: lbl气泡A(i).Left = 840
            shp气泡A(i).Top = 285: shp气泡A(i).Left = 840
            txt气泡A(i).Top = 480: txt气泡A(i).Left = 960
            Call GetTextHight(txt气泡A(i).Text, lngTxtW, lngTxtH)
            If lngTxtH = 0 And lngTxtW <> 0 Then
                txt气泡A(i).Width = lngTxtW - 100
                txt气泡A(i).Height = 330
            ElseIf lngTxtH <> 0 And lngTxtW = 0 Then
                txt气泡A(i).Width = lng最大气泡宽度 - 100
                txt气泡A(i).Height = lngTxtH
            Else
                txt气泡A(i).Width = lngTxtW - 100
                txt气泡A(i).Height = lngTxtH
            End If
            
            
            If txt气泡A(i).Width > 4700 And txt气泡A(i).Height > 1800 Then
                txt气泡A(i).Top = txt气泡A(i).Top + 100
                txt气泡A(i).Left = txt气泡A(i).Left + 100
            End If
            
            shp气泡A(i).Width = txt气泡A(i).Left + txt气泡A(i).Width + IIf(txt气泡A(i).Width > 4700 And txt气泡A(i).Height > 1800, 220, 120) - shp气泡A(i).Left
            shp气泡A(i).Height = txt气泡A(i).Top + txt气泡A(i).Height + IIf(txt气泡A(i).Width > 4700 And txt气泡A(i).Height > 1800, 350, IIf(txt气泡A(i).Height < 500, 120, 180)) - shp气泡A(i).Top
            pic气泡A(i).Height = shp气泡A(i).Top + shp气泡A(i).Height + 75
            lbl阅读(i).Top = shp气泡A(i).Top + shp气泡A(i).Height - lbl阅读(i).Height - 15
            lbl阅读(i).Left = shp气泡A(i).Left + shp气泡A(i).Width + 100
        End If
    Next
    picChat.Height = pic气泡A(pic气泡A.Count - 1).Top + pic气泡A(pic气泡A.Count - 1).Height + 100
    vsBar.Visible = picBack.Height - cbsChat.Height < picChat.Height
    vsBar.Max = (picChat.Height - picBack.Height + cbsChat.Height) / 100
End Sub


Private Sub txtIn_GotFocus()
    Call zlControl.TxtSelAll(txtIn)
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    
    If InStr("&'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SendMsg
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub vsBar_Change()
    Dim lngValue As Long
    On Error Resume Next
    
    lngValue = vsBar.Value
    picChat.Top = (-lngValue * 100 + cbsChat.Height)
End Sub


Private Sub Subclass_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    '自定义的消息处理函数
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '鼠标坐标
    Dim intShift As Integer              '鼠标按键
    Dim bWay As Boolean                  '鼠标方向
    Dim bMouseFlag As Boolean            '鼠标事件激活标志
    Dim wzDelta, wKeys As Integer
    On Error Resume Next
    
    If vsBar.Visible = False Then Exit Sub
    Select Case Msg
        Case WM_MOUSEWHEEL   '滚动
            wzDelta = (wParam And &HFFFF0000) \ &H10000 '取出32位值的高16位
            If wzDelta > 0 Then
                vsBar.Value = IIf(vsBar.Value - vsBar.LargeChange < 0, 0, vsBar.Value - vsBar.LargeChange)
            Else
                vsBar.Value = IIf(vsBar.Value + vsBar.LargeChange > vsBar.Max, vsBar.Max, vsBar.Value + vsBar.LargeChange)
            End If
    End Select
End Sub


Private Sub SendMsg()
    Dim strSQL As String
    Dim strSend As String
    Dim lng信息ID As Long
    Dim strDate As String, strDateSQL As String
    
    On Error GoTo errH
    If txtIn.Text = "" Then Exit Sub
    If rptList.SelectedRows.Count < 1 Then Exit Sub
    If rptList.SelectedRows(0) Is Nothing Then Exit Sub
    If rptList.SelectedRows(0).GroupRow Then Exit Sub
    If Val(rptList.SelectedRows(0).Record(col_会话ID).Value & "") = 0 Then Exit Sub
    If rptList.SelectedRows(0).Record(col_发送人).Value & "" = "" Then Exit Sub

    lng信息ID = zlDatabase.GetNextId("聊天信息表")
    strSend = Replace(txtIn.Text, "'", "")
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strDateSQL = "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
    strSQL = "Zl_聊天信息表_Edit(1," & Val(rptList.SelectedRows(0).Record(col_会话ID).Value & "") & "," & lng信息ID & ",'" & UserInfo.姓名 & "','" & strSend & "'," & strDateSQL & ",'" & rptList.SelectedRows(0).Record(col_发送人).Value & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

    txtIn.Text = ""
    Call AddMsg(strSend, strDate, UserInfo.姓名, lng信息ID)
    vsBar.Value = vsBar.Max '自动定位到最后
    Call zlControl.ControlSetFocus(txtIn)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReAddChat(lngMaxID As Long, Lng会话ID As Long)
    '加载当前会话的最新消息
    Dim strMsg As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If Lng会话ID = 0 Then Exit Sub
    strSQL = "select a.id,A.会话ID,A.发送人,A.发送内容,A.发送时间,A.接收人,A.阅读时间 from 聊天信息表 A WHERE A.会话ID=[1] and a.id>[2] order by A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Lng会话ID, lngMaxID)
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
               '处理换行格式不正常的问题
               strMsg = Replace(rsTmp!发送内容 & "", vbCrLf, "[换行处理]")
               strMsg = Replace(strMsg, Chr(10), vbCrLf)
               strMsg = Replace(strMsg, "[换行处理]", vbCrLf)
               Call AddMsg(strMsg, Format(rsTmp!发送时间 & "", "yyyy-MM-dd HH:mm"), rsTmp!发送人 & "", Val(rsTmp!ID & ""))
            rsTmp.MoveNext
        Loop
    End If
    vsBar.Value = vsBar.Max '自动定位到最后
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub AddMsg(ByVal strSend As String, ByVal strDate As String, ByVal str发送人 As String, ByVal lngID As Long)
    Dim lng气泡 As Long

    On Error GoTo errH
     '气泡容器加载
    lng气泡 = pic气泡A.Count
    Load pic气泡A(lng气泡)
    Set pic气泡A(lng气泡).Container = picChat
    pic气泡A(lng气泡).Tag = str发送人
    pic气泡A(lng气泡).Visible = True
    
    '头像显示
    Load img气泡A(lng气泡)
    Set img气泡A(lng气泡).Container = pic气泡A(lng气泡)
    img气泡A(lng气泡).Tag = lngID
    img气泡A(lng气泡).Visible = True
    

    '加载气泡
    Load shp气泡A(lng气泡)
    Set shp气泡A(lng气泡).Container = pic气泡A(lng气泡)
    shp气泡A(lng气泡).Visible = True

    Load txt气泡A(lng气泡)
    Set txt气泡A(lng气泡).Container = pic气泡A(lng气泡)
    txt气泡A(lng气泡).Visible = True
    txt气泡A(lng气泡).Text = strSend
    
    
    '加载气泡信息
    Load lbl气泡A(lng气泡)
    Set lbl气泡A(lng气泡).Container = pic气泡A(lng气泡)
    lbl气泡A(lng气泡).Visible = True

    '加载阅读信息
    Load lbl阅读(lng气泡)
    Set lbl阅读(lng气泡).Container = pic气泡A(lng气泡)
    lbl气泡A(lng气泡).Visible = True

    lbl气泡A(lng气泡).Caption = str发送人 & "  " & Format(strDate, "yyyy-MM-dd HH:mm")
    lbl阅读(lng气泡).Caption = "已读"
    lbl阅读(lng气泡).Visible = False
    
    If str发送人 = UserInfo.姓名 Then
        Set img气泡A(lng气泡).Picture = img气泡B.Picture
        shp气泡A(lng气泡).BackColor = RGB(221, 235, 255)
        shp气泡A(lng气泡).FillColor = RGB(221, 235, 255)
        txt气泡A(lng气泡).BackColor = RGB(208, 224, 240)
    End If
    
    pic气泡A(lng气泡).Top = pic气泡A(lng气泡 - 1).Top + pic气泡A(lng气泡 - 1).Height + 100

    Call CtlResize(lng气泡)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetTextHight(ByVal strText As String, lngWidth As Long, lngHeight As Long)
    '用于计算自适应文本框宽度高度
    Dim lng最大气泡宽度 As Long
    Dim lngW As Long, lngH As Long
    
    On Error Resume Next

    lng最大气泡宽度 = IIf(picChat.Width / 2 - 880 > 7000, 7000, picChat.Width / 2 - 880) '获取最大的标签宽度用于显示文本框
    
    '先计算宽度
    vsTmp.ColWidthMax = 0: vsTmp.ColWidthMin = 0
    vsTmp.RowHeightMin = 255: vsTmp.RowHeightMax = 255
    vsTmp.TextMatrix(0, 0) = strText
    vsTmp.Redraw = True
    vsTmp.AutoSizeMode = flexAutoSizeColWidth
    vsTmp.AutoSize 0
    vsTmp.Redraw = True
    vsTmp.Refresh
    lngW = vsTmp.ColWidth(0) - 20

    vsTmp.RowHeightMin = 0: vsTmp.RowHeightMax = 0
    vsTmp.ColWidthMax = lng最大气泡宽度
    vsTmp.ColWidthMin = lng最大气泡宽度
    vsTmp.TextMatrix(0, 0) = strText
    vsTmp.Redraw = True
    vsTmp.AutoSizeMode = flexAutoSizeRowHeight
    vsTmp.AutoSize 0
    vsTmp.Redraw = True
    vsTmp.Refresh
    lngH = vsTmp.RowHeight(0)

    If lngW > lng最大气泡宽度 And lngH = 255 Then
        lngWidth = lngW
    Else
        lngHeight = lngH
        lngWidth = IIf(lngW < lng最大气泡宽度, lngW, lng最大气泡宽度)
    End If
End Function


Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    On Error GoTo errH
    strHead = "ID;组号;用药来源;诊疗项目ID;收费细目ID;频率间隔;间隔单位;用法ID;煎法ID;终止时间;" & _
                "期效,450,4;开始时间,1000,1;药品类别,850,4;用药内容,2000,1;用法,1000,1;单量,850,4;总量,850,4;天数,450,4;执行频次,1000,4;频率次数;配方数据"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        .Editable = flexEDNone
        .WordWrap = True
        .AutoSize colB_用药内容
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetAdvice()
    '提取病人处方信息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, i As Long
    
    On Error GoTo errH
    With vsAdvice
         .Rows = .FixedRows
        If rptList.SelectedRows(0).Record(COL_医嘱IDs).Value & "" = "" Then
            .Rows = .FixedRows + 1
        Else
            strSQL = "Select a.Id, a.相关id As 组号, a.诊疗类别 As 药品类别, a.医嘱内容 As 用药内容, a.医生嘱托 As 医生嘱托, a.诊疗项目id, a.收费细目id, a.天数, a.开始执行时间 As 开始时间," & vbNewLine & _
                    "       a.执行终止时间 As 终止时间, Decode(a.病人来源, 1, a.总给予量 / e.门诊包装, 2, a.总给予量 / e.住院包装, a.总给予量) As 总给予量, a.单次用量, a.执行频次, a.频率次数," & vbNewLine & _
                    "       a.频率间隔, a.间隔单位, b.诊疗项目id As 用药id, c.计算单位, b.医嘱内容 As 用法, d.名称 As 中药用法, Decode(a.病人来源, 1, e.门诊单位, 2, e.住院单位) As 住院单位,a.医嘱期效" & vbNewLine & _
                    "From 病人医嘱记录 A, 病人医嘱记录 B, 诊疗项目目录 C, 诊疗项目目录 D, 药品规格 E" & vbNewLine & _
                    "Where a.相关id = b.Id And a.诊疗项目id = c.Id And a.收费细目id = e.药品id(+) And b.诊疗项目id = d.Id And a.病人id = [1] And" & vbNewLine & _
                    "      (a.Id In (Select Column_Value As 医嘱id From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) Or" & vbNewLine & _
                    "      a.相关id In (Select Column_Value As 医嘱id From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))))" & vbNewLine & _
                    "Order By a.病人id, a.主页id, a.挂号单, a.序号, a.开始执行时间"
        
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rptList.SelectedRows(0).Record(col_病人Id).Value & ""), rptList.SelectedRows(0).Record(COL_医嘱IDs).Value & "")
            If Not rsTmp.EOF Then
                 .Redraw = flexRDNone
                 If .TextMatrix(.Rows - 1, colB_用药内容) = "" And Val(.TextMatrix(.Rows - 1, COLB_ID)) = 0 Then .Rows = .Rows - 1
                 For i = 1 To rsTmp.RecordCount
                    If (rsTmp!药品类别 & "" = "7" Or rsTmp!药品类别 & "" = "E") And Val(.Cell(flexcpData, .Rows - 1, COLB_组号)) = Val(rsTmp!组号 & "") Then
                        If rsTmp!药品类别 & "" = "7" Then
                            .TextMatrix(.Rows - 1, COLB_配方数据) = .TextMatrix(.Rows - 1, COLB_配方数据) & vbCrLf & rsTmp!用药内容 & " " & FormatEx(NVL(rsTmp!单次用量), 5) & rsTmp!计算单位 & " " & rsTmp!医生嘱托
                        Else
                            .TextMatrix(.Rows - 1, COLB_煎法id) = Val(rsTmp!诊疗项目ID & "")
                        End If
                    Else
                        .Rows = .Rows + 1
                        lngRow = .Rows - 1
                            
                        '隐藏列
                        .TextMatrix(lngRow, COLB_期效) = IIf(Val(rsTmp!医嘱期效 & "") = 1, "临嘱", "长嘱")
                        .ColHidden(COLB_期效) = rptList.SelectedRows(0).Record(col_来源).Value = "门诊"
               
                        .TextMatrix(lngRow, COLB_诊疗项目ID) = Val(rsTmp!诊疗项目ID & "")
                        .TextMatrix(lngRow, COLB_收费细目ID) = Val(rsTmp!收费细目id & "")
                        .TextMatrix(lngRow, COLB_频率间隔) = Val(rsTmp!频率间隔 & "")
                        .TextMatrix(lngRow, COLB_间隔单位) = rsTmp!间隔单位 & ""
                        .TextMatrix(lngRow, COLB_用法id) = Val(rsTmp!用药id & "")
                        .TextMatrix(lngRow, COLB_终止时间) = Format(rsTmp!终止时间 & "", "yyyy-mm-dd hh:mm")
                        .TextMatrix(lngRow, COLB_开始时间) = Format(rsTmp!开始时间 & "", "yyyy-mm-dd hh:mm")
                        .TextMatrix(lngRow, colB_药品类别) = Decode(rsTmp!药品类别 & "", "5", "西成药", "6", "中成药", "中草药")
                        .TextMatrix(lngRow, colB_用药内容) = IIf(.TextMatrix(lngRow, colB_药品类别) = "中草药", rsTmp!用法 & "", rsTmp!用药内容 & "")
                        .TextMatrix(lngRow, COLB_用法) = IIf(.TextMatrix(lngRow, colB_药品类别) = "中草药", rsTmp!中药用法 & "", rsTmp!用法 & "")
                        .TextMatrix(lngRow, COLB_单次用量) = IIf(.TextMatrix(lngRow, colB_药品类别) = "中草药", "", FormatEx(NVL(rsTmp!单次用量), 5)) & IIf(.TextMatrix(lngRow, colB_药品类别) = "中草药", "", rsTmp!计算单位 & "")
                        .TextMatrix(lngRow, COLB_总给予量) = FormatEx(NVL(rsTmp!总给予量), 5) & IIf(Val(rsTmp!总给予量 & "") = 0, "", IIf(.TextMatrix(lngRow, colB_药品类别) = "中草药", "付", rsTmp!住院单位 & ""))
                        .Cell(flexcpData, lngRow, COLB_组号) = Val(rsTmp!组号 & "")
                        
                        If .TextMatrix(lngRow, colB_药品类别) = "中草药" Then
                            .TextMatrix(lngRow, COLB_组号) = ""
                        Else
                            If .Cell(flexcpData, lngRow, COLB_组号) = .Cell(flexcpData, lngRow - 1, COLB_组号) And .Cell(flexcpData, lngRow, COLB_组号) <> "" Then
                                If .TextMatrix(lngRow - 1, COLB_组号) = "" Then .TextMatrix(lngRow - 1, COLB_组号) = -(lngRow - 1)
                                .TextMatrix(lngRow, COLB_组号) = .TextMatrix(lngRow - 1, COLB_组号)
                            End If
                        End If
    
                        If rsTmp!天数 & "" = "" Then
                            If rsTmp!终止时间 & "" <> "" And rsTmp!开始时间 & "" <> "" Then
                                .TextMatrix(lngRow, COLB_天数) = FormatEx(NVL(DateDiff("d", CDate(rsTmp!开始时间 & ""), CDate(rsTmp!终止时间 & ""))), 5)
                            End If
                        Else
                            .TextMatrix(lngRow, COLB_天数) = FormatEx(NVL(rsTmp!天数), 5)
                        End If
                        .TextMatrix(lngRow, COLB_执行频次) = rsTmp!执行频次 & ""
                        
                        If rsTmp!药品类别 & "" = "7" Then
                            .TextMatrix(lngRow, COLB_配方数据) = "配方信息：" & .TextMatrix(lngRow, COLB_配方数据) & vbCrLf & rsTmp!用药内容 & " " & FormatEx(NVL(rsTmp!单次用量), 5) & rsTmp!计算单位 & " " & rsTmp!医生嘱托
                        ElseIf rsTmp!药品类别 & "" = "E" Then
                            .TextMatrix(lngRow, COLB_煎法id) = Val(rsTmp!诊疗项目ID & "")
                        End If
                        
                    End If
                    rsTmp.MoveNext
                 Next
                 .Redraw = flexRDDirect
            Else
                .Rows = .FixedRows + 1
                .TextMatrix(.Rows - 1, colB_用药内容) = "医嘱已删除"
                .Cell(flexcpForeColor, .Rows - 1, colB_用药内容, .Rows - 1, colB_用药内容) = vbRed
                .Cell(flexcpFontBold, .Rows - 1, colB_用药内容, .Rows - 1, colB_用药内容) = True
            End If
    
        End If
        Call SetTag一并给药
        .WordWrap = True
        '自动调整行高
        .AutoSize colB_用药内容
        .Cell(flexcpBackColor, .FixedRows, colB_药品类别, .Rows - 1, colB_药品类别) = &H8000000B      '灰蓝色
        .Cell(flexcpBackColor, .FixedRows, COLB_用法, .Rows - 1, COLB_用法) = &H8000000B      '灰蓝色
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, 0) = &H8000000B
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SetTag一并给药(Optional ByVal lng组号 As Long)
'功能：在一并给药的医嘱前加标志
    Dim i As Long
    Dim lngUpRow As Long

    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If lng组号 = 0 Then .TextMatrix(i, 0) = ""
            If lng组号 <> 0 And Val(.TextMatrix(i, COLB_组号)) = lng组号 Then .TextMatrix(i, 0) = ""
            If Val(.TextMatrix(i, COLB_组号)) <> 0 And ((lng组号 = Val(.TextMatrix(i, COLB_组号)) And lng组号 <> 0) Or lng组号 = 0) And .RowHidden(i) = False Then
                lngUpRow = GetUpRow(i)
                If lngUpRow = 0 Then
                    .TextMatrix(i, 0) = "┏"
                Else
                    If Val(.TextMatrix(i, COLB_组号)) = Val(.TextMatrix(lngUpRow, COLB_组号)) And i <> lngUpRow Then
                        If .TextMatrix(lngUpRow, 0) = "┗" Then
                            .TextMatrix(lngUpRow, 0) = "┃"
                        End If
                        .TextMatrix(i, 0) = "┗"
                    Else
                        .TextMatrix(i, 0) = "┏"
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Function GetUpRow(ByVal lngRow As Long) As Long
'功能：取上一行有效行
    Dim i As Long

    With vsAdvice
        lngRow = lngRow - 1
        For i = lngRow To 1 Step -1
            If .RowHidden(i) = False Then
                GetUpRow = i: Exit For
            End If
        Next
    End With
End Function

Public Function ShowMe(frmParent As Object, intType As Integer)
'      intType 1=门诊 2=住院
    Dim i As Long, bln未读 As Boolean
    
    Set mfrmParent = frmParent
    mintType = intType
    Me.Show , frmParent
    
    '初始化判断是否有未读消息，没有未读消息隐藏窗体
    For i = 0 To rptList.Records.Count - 1
        If rptList.Records(i)(col_未读时间).Value & "" = "1900-01-01" Then
            bln未读 = True
            Exit For
        End If
    Next
    If bln未读 = False Then Me.Hide
End Function


Private Sub mclsNotice_DataArrival(ByVal lngNoticeCode As Long, ByVal intChangeType As Integer, ByVal strTableOwner As String, _
    ByVal TableName As String, ByVal strRowId As String)
    'DCN的数据变动通知
    'lngNoticeCode：消息标识(固定值)，用来区分是哪个消息
    'intChangeType：数据变动类型，1-新增 2-更新 3-删除
    'strTableOwner：注册DCN表所有者
    'TableName：表名
    'strRowid：返回的数据变动rowid
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, bln未读 As Boolean
    
    On Error GoTo errH
    If TableName <> "聊天信息表" Then Exit Sub
    
    If intChangeType = 1 Then '有新的消息过来了
        strSQL = "Select a.Id, a.会话id, a.发送人, a.发送内容, a.发送时间, a.接收人, a.阅读时间, b.对象标识, b.对象内容, b.病人id, b.就诊id, b.病人来源, b.创建人, b.创建时间," & vbNewLine & _
                    "       b.接收人 As 会话接收人, b.状态" & vbNewLine & _
                    "From 聊天信息表 A, 聊天会话表 B" & vbNewLine & _
                    "Where a.会话id = b.Id And a.Rowid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRowId)
        If rptList.SelectedRows.Count > 0 Then
            If Val(rsTmp!会话ID & "") = Val(rptList.SelectedRows(0).Record(col_会话ID).Value) Then
                If Val(rsTmp!ID & "") > Val(img气泡A(img气泡A.Count - 1).Tag) Then
                    If Me.Visible = False Then
                        '新增过滤会话ID
                        Call ReAddChat(Val(img气泡A(img气泡A.Count - 1).Tag), Val(rsTmp!会话ID & ""))
                        If InStr("," & mstr未读会话ids & ",", "," & Val(rsTmp!会话ID & "") & ",") = 0 Then
                            mstr未读会话ids = mstr未读会话ids & "," & Val(rsTmp!会话ID & "")
                        End If
                        If TmrIcon.Enabled = False Then StartMsg
                    Else
                        Call ReAddChat(Val(img气泡A(img气泡A.Count - 1).Tag), Val(rsTmp!会话ID & ""))
                        Call ReadMsg(Val(rsTmp!会话ID & ""), rptList.SelectedRows(0).Record(col_发送人).Value & "")
                    End If
                End If
            Else
                If Val(rsTmp!病人来源 & "") = mintType And rsTmp!接收人 & "" = UserInfo.姓名 Then '区分门诊住院和接收人
                    '新增过滤会话ID
                    If InStr("," & mstr未读会话ids & ",", "," & Val(rsTmp!会话ID & "") & ",") = 0 Then
                        mstr未读会话ids = mstr未读会话ids & "," & Val(rsTmp!会话ID & "")
                    End If
                    '启动闪烁
                    If TmrIcon.Enabled = False Then StartMsg
                    
                    '刷新列表
                    LoadMsg
    
                    '重新定位当前会话
                    If Val(rptList.Tag) <> 0 Then
                        For i = 0 To rptList.Rows.Count - 1
                            With rptList.Rows(i)
                                If Not .GroupRow Then
                                    If .Record(col_会话ID).Value = Val(rptList.Tag) Then
                                        Exit For
                                    End If
                                End If
                            End With
                        Next
                    
                        If i <= rptList.Rows.Count - 1 Then
                            Set rptList.FocusedRow = rptList.Rows(i)
                        End If
                    End If
                End If
            End If
        Else
            If Val(rsTmp!病人来源 & "") = mintType And rsTmp!接收人 & "" = UserInfo.姓名 Then '区分门诊住院和接收人
                '新增过滤会话ID
                If InStr("," & mstr未读会话ids & ",", "," & Val(rsTmp!会话ID & "") & ",") = 0 Then
                    mstr未读会话ids = mstr未读会话ids & "," & Val(rsTmp!会话ID & "")
                End If
                '启动闪烁
                If TmrIcon.Enabled = False Then StartMsg
                
                '刷新列表
                LoadMsg
            End If
        End If
    ElseIf intChangeType = 2 Then '有阅读更新的消息过来了
        If rptList.SelectedRows.Count = 0 Then Exit Sub
        If pic气泡A.Count = 0 Then Exit Sub
        For i = lbl阅读.Count - 1 To 1 Step -1
            If lbl阅读(i).Visible = False And pic气泡A(i).Tag = UserInfo.姓名 Then
                bln未读 = True
                Exit For
            End If
        Next
        If bln未读 = False Then Exit Sub

        strSQL = "select a.id,A.会话ID,A.发送人,A.发送内容,A.发送时间,A.接收人,A.阅读时间 from 聊天信息表 A WHERE a.RowId=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRowId)
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                If Val(rsTmp!会话ID & "") = Val(rptList.SelectedRows(0).Record(col_会话ID).Value) And rsTmp!发送人 & "" = UserInfo.姓名 Then
                    For i = 1 To lbl阅读.Count - 1
                        If pic气泡A(i).Tag = UserInfo.姓名 Then
                            lbl阅读(i).Visible = True
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



