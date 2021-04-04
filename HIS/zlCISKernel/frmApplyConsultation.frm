VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApplyConsultation 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "会诊申请单"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   Icon            =   "frmApplyConsultation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   10500
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   61
      Top             =   2385
      Width           =   3645
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   930
         ScaleHeight     =   285
         ScaleWidth      =   2520
         TabIndex        =   62
         Top             =   15
         Width           =   2520
         Begin VB.ComboBox cboItem 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -45
            TabIndex        =   9
            Top             =   -25
            Width           =   2520
         End
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "会诊项目"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   0
         TabIndex        =   63
         Top             =   45
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   855
         X2              =   3105
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   16
      Left            =   8970
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   15
      Left            =   5325
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   14
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1500
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
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
      Height          =   210
      Index           =   13
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   2000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5235
      Width           =   8550
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   56
      Top             =   2835
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Height          =   210
         Index           =   12
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   75
         Width           =   5175
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   885
         X2              =   6285
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "邀请医院"
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
         Index           =   10
         Left            =   0
         TabIndex        =   57
         Top             =   75
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   1
      Left            =   3150
      Picture         =   "frmApplyConsultation.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "编辑(F4)"
      Top             =   2415
      Width           =   285
   End
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   480
      TabIndex        =   55
      Top             =   3285
      Width           =   9510
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         CausesValidation=   0   'False
         Height          =   1305
         Left            =   15
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   9495
         _cx             =   16757
         _cy             =   2293
         Appearance      =   2
         BorderStyle     =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmApplyConsultation.frx":6948
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
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
      End
   End
   Begin VB.Frame fraIdea 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   2300
      Left            =   -60
      TabIndex        =   47
      Top             =   7515
      Width           =   10700
      Begin VB.CommandButton cmdInfo 
         Caption         =   "…"
         Height          =   270
         Index           =   2
         Left            =   9765
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   1755
         Width           =   270
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "…"
         Height          =   270
         Index           =   1
         Left            =   6810
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   1725
         Width           =   270
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Height          =   210
         Index           =   10
         Left            =   5295
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1500
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
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
         Height          =   210
         Index           =   9
         Left            =   1470
         TabIndex        =   19
         Text            =   "2013-06-20 18:00"
         Top             =   1770
         Width           =   1695
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Index           =   18
         Left            =   1470
         MaxLength       =   4000
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   8565
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Height          =   210
         Index           =   11
         Left            =   8745
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1000
      End
      Begin VB.CommandButton cmdDate 
         Height          =   285
         Index           =   0
         Left            =   3165
         Picture         =   "frmApplyConsultation.frx":69D3
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   1740
         Width           =   285
      End
      Begin VB.Line Line2 
         BorderStyle     =   2  'Dash
         X1              =   -30
         X2              =   11130
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "会诊完成时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   15
         Left            =   165
         TabIndex        =   54
         Top             =   1770
         Width           =   1260
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   1410
         X2              =   3135
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   5205
         X2              =   6810
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "会诊科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   4350
         TabIndex        =   53
         Top             =   1770
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   8670
         X2              =   9750
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "会诊医师"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   7815
         TabIndex        =   52
         Top             =   1770
         Width           =   840
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "会诊意见"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   525
         TabIndex        =   51
         Top             =   195
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   0
      Left            =   9720
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   4770
      Width           =   270
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
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
      Height          =   210
      Index           =   8
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4800
      Width           =   8265
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Index           =   17
      Left            =   1410
      MaxLength       =   4000
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5640
      Width           =   8565
   End
   Begin VB.PictureBox picNo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   8320
      ScaleHeight     =   345
      ScaleWidth      =   1680
      TabIndex        =   26
      Top             =   1110
      Visible         =   0   'False
      Width           =   1680
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   60
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   34
         X1              =   255
         X2              =   1695
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   37
         Left            =   0
         TabIndex        =   28
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   405
      Index           =   0
      Left            =   1440
      TabIndex        =   25
      Top             =   2745
      Width           =   1575
      Begin VB.OptionButton optInfo 
         BackColor       =   &H8000000E&
         Caption         =   "院内"
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
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optInfo 
         BackColor       =   &H8000000E&
         Caption         =   "院外"
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
         Left            =   840
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   7
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1640
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   5
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2025
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   4
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2025
      Width           =   1500
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   3
      Left            =   3930
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1640
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   2
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1640
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   1
      Left            =   3930
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2025
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
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
      Height          =   210
      Index           =   19
      Left            =   1425
      TabIndex        =   4
      Text            =   "2013-06-20 18:00"
      Top             =   2430
      Width           =   1725
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Height          =   210
      Index           =   0
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1640
      Width           =   1000
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Index           =   1
      Left            =   4785
      TabIndex        =   1
      Top             =   2310
      Width           =   1575
      Begin VB.OptionButton optInfo 
         BackColor       =   &H8000000E&
         Caption         =   "普通"
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
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optInfo 
         BackColor       =   &H8000000E&
         Caption         =   "紧急"
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
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   10275
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   251330561
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.Line Line1 
      Index           =   33
      X1              =   1335
      X2              =   2925
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "申请科室"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   36
      Left            =   480
      TabIndex        =   46
      Top             =   7125
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   8865
      X2              =   9990
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "主治医师"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   34
      Left            =   8010
      TabIndex        =   45
      Top             =   7125
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   5220
      X2              =   6330
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "申请医师"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   33
      Left            =   4335
      TabIndex        =   44
      Top             =   7125
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   1350
      X2              =   9990
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "会诊目的"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   17
      Left            =   450
      TabIndex        =   43
      Top             =   5235
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   1350
      X2              =   9705
      Y1              =   5025
      Y2              =   5025
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "病情摘要"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   13
      Left            =   450
      TabIndex        =   42
      Top             =   5640
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "临床诊断"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   19
      Left            =   450
      TabIndex        =   41
      Top             =   4800
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   6405
      X2              =   8040
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   8895
      X2              =   10005
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   6405
      X2              =   7530
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   3825
      X2              =   4935
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "会诊范围"
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
      Index           =   11
      Left            =   480
      TabIndex        =   39
      Top             =   2880
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   1335
      X2              =   2430
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1335
      X2              =   2430
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   3825
      X2              =   4935
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4185
      X2              =   6015
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label lblHead 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "会诊申请单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4170
      TabIndex        =   38
      Top             =   750
      Width           =   1875
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "会诊性质"
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
      Index           =   6
      Left            =   3840
      TabIndex        =   37
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "床    号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   36
      Top             =   2025
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "科    室"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   5550
      TabIndex        =   35
      Top             =   2025
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "姓    名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   34
      Top             =   1640
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "住 院 号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2940
      TabIndex        =   33
      Top             =   2025
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "年    龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   5550
      TabIndex        =   32
      Top             =   1635
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "性    别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2940
      TabIndex        =   31
      Top             =   1635
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "会诊时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   35
      Left            =   480
      TabIndex        =   30
      Top             =   2430
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   1320
      X2              =   3135
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "病    情"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   8055
      TabIndex        =   29
      Top             =   1635
      Width           =   840
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   75
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmApplyConsultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormUnload(Cancel As Integer)


Private Enum mCtlID '界面上的控件索引值
    txtNO = 6
    txt姓名 = 2
    txt性别 = 3
    txt年龄 = 7
    txt病情 = 0
    txt床号 = 5
    txt住院号 = 1
    txt科室 = 4
    
    txt会诊时间 = 19
    cmd会诊时间 = 1
    
    fra会诊性质 = 1
    opt普通 = 3
    opt紧急 = 2
    
    fra会诊项目 = 3
    
    fra会诊范围 = 0
    opt院内 = 1
    opt院外 = 0
    
    fra邀请医院 = 2
    txt邀请医院 = 12
    
    txt临床诊断 = 8
    cmd临床诊断 = 0
    
    txt会诊目的 = 13
    txt病情摘要 = 17
    txt申请科室 = 14
    txt申请医师 = 15
    txt主治医师 = 16
    txt会诊意见 = 18
    txt完成时间 = 9
    cmd完成时间 = 0
    txt会诊科室 = 10
    cmd会诊科室 = 1
    txt会诊医师 = 11
    cmd会诊医师 = 2
End Enum

Private Enum VSDETAIL_COL
    COL_代表科室 = 0
    COL_邀请科室
    COL_医生级别
    COL_邀请医生
    
    COL_邀请科室ID
    COL_邀请医生ID
End Enum

Private mobjVBA As Object
Private mobjScript As clsScript
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mobjReport As Object
Private mclsDiagEdit As zlMedRecPage.clsDiagEdit
Private mclsMipModule As zl9ComLib.clsMipModule '消息平台对象
Private mbln补录 As Boolean
Private mbln填意见 As Boolean

Private mstr上次转科时间 As String
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlngAdviceID As Long '一组会诊申请中的某一条医嘱
Private mlngNo As Long ' 申请序号
Private mlng病人科室id As Long
Private mlng病区ID As Long
Private mlng开单科室ID As Long
Private mstr诊断IDs As String
Private mintPState As Integer
Private mdatTurn As Date
Private mint险类 As Integer
Private mintType As Integer '0－新增，1－修改，2－查看，3－书写会诊意见
Private mintPos As Integer
Private mstr入院时间 As String
Private mrsItem会诊 As ADODB.Recordset
Private mlng项目ID As Long
Private mbln书写要求 As Boolean 'true-只写一份
Private mblnReturn As Boolean
Private mbln意见 As Boolean '是否显示会诊意见
Private mbln提醒对码 As Boolean
Private mlng执行科室ID As Long '会诊病人时的执行科室ID
Private mlng发送号 As Long
Private mstr检查入院诊断 As String
Private mstr摘要 As String '摘要，由 gclsInsure.GetItemInfo 获取
Private mstr费别 As String
Private mlng前提ID As Long
Private mbytBaby As Byte  '婴儿序号
'医嘱编辑界面的卡片控件的选择状态
Private mrsCard As ADODB.Recordset
Private mstr_Ctl_医生嘱托 As String            '主项目改变后变为默认

Public Function ShowMe(ByRef frmParent As Object, Optional ByRef lng医嘱ID As Long, Optional ByRef lngNo As Long, Optional ByVal intType As Integer = 2, _
    Optional ByVal intPos As Integer, Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal lng科室id As Long, _
    Optional ByVal lng开单科室ID As Long, Optional ByVal lng病区ID As Long, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, _
    Optional ByRef objMip As Object, Optional ByVal lng项目id As Long, Optional ByVal rsCard As ADODB.Recordset, Optional ByVal lng前提ID As Long, Optional ByVal bytBaby As Byte) As Boolean
'功能：公共接口方法
'参数：intType 操作类型 0－新增，1－修改，2－查看，3－书写会诊意见，默认为查看状态2
'      strDefine 医嘱内容格式串，intPState 病人状态，intPos调用位置 0－医生站主界面，1－医嘱编译界面
'      lngNo 申请序号
'      lng项目ID-医嘱编辑界面输入会诊项目ID
'      rsCard，医嘱卡片和医嘱表格的信息。只有在医嘱编辑界面点医嘱内容右边的下拉按钮时才会传入。
'两个引用参数说明：lngNo－新增时作为传出参数，其它操作都不变；lng医嘱ID－新增修改时做为传出参数。其它操作不变
    mintType = intType
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mintPos = intPos
    mlngAdviceID = lng医嘱ID
    mlng病人科室id = lng科室id
    mlng病区ID = lng病区ID
    mlng开单科室ID = lng开单科室ID
    mlng前提ID = lng前提ID
    mbytBaby = bytBaby
    mintPState = intPState
    mdatTurn = datTurn
    mlngNo = lngNo
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    mlng项目ID = lng项目id
    Set mrsCard = rsCard
    mblnOK = False
    On Error Resume Next
    
    If intPos = 0 And (intType = 0 Or intType = 1 Or intType = 3) Then
        Me.Show , frmParent
    Else
        Me.Show 1, frmParent
    End If
    
    On Error GoTo 0
    
    If mblnOK Then
        lng医嘱ID = mlngAdviceID
        lngNo = mlngNo
    End If
    
    ShowMe = mblnOK
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Asc("'") = KeyAscii Or Asc(";") = KeyAscii Or Asc("%") = KeyAscii Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("当前申请单已经进行了调整尚未保存，是否要继续退出？", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Me.Tag = IIF(mblnOK, "1", "")
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
    If Not mclsDiagEdit Is Nothing Then Set mclsDiagEdit = Nothing
 
    mbln补录 = False
    mbln填意见 = False
    mstr入院时间 = ""
    mstr上次转科时间 = ""
    mint险类 = 0
    Set mclsMipModule = Nothing
    RaiseEvent FormUnload(Cancel)
End Sub

Private Sub Form_Load()
    
    Call InitCommandBar
    
    mbln书写要求 = Val(zlDatabase.GetPara(237, glngSys)) = 1
    mstr检查入院诊断 = zlDatabase.GetPara("要求输入入院诊断", glngSys, p住院医嘱下达)
    Call InitVSDetail
    
    Call LoadItem会诊
    
    Call LoadData
    
    Call SetFormState
    
    mblnChange = False
End Sub

Private Sub SetFormState()
'功能：设置界面状态
    Select Case mintType
    Case 0, 1
        Me.Height = 7900
        Call SetItemEditable(1, 1, 1, 1, 1, 1, 1, 1, 1)
        picNo.Visible = False
    Case 2
        Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1)
        picNo.Visible = True
    Case 3
        Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1, 1, 1, 1, 1)
        picNo.Visible = True
    Case 4
        Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1)
        picNo.Visible = True
    End Select
End Sub

Private Sub InitVSDetail()
'功能：初始化会诊邀请明细表格
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
    
    
    If mbln书写要求 Then
        strHead = "代表科室,1040,4;邀请科室,4560,1;医生级别,1650,1;邀请医生,1650,1;邀请科室ID;邀请医生ID"
    Else
        strHead = "代表科室;邀请科室,5600,1;医生级别,1650,1;邀请医生,1650,1;邀请科室ID;邀请医生ID"
    End If
    
    arrHead = Split(strHead, ";")
    
    With vsDetail
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColData(COL_医生级别) = "    |住院医师|主治医师|(副)主任医师"
    End With
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub vsDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
    If mintType = 2 Then Exit Sub
    If NewRow < 0 Then Exit Sub
    With vsDetail
        If (NewCol = COL_邀请科室 Or NewCol = COL_邀请医生) Then
            If optInfo(opt院内).value Then
                .Editable = flexEDKbdMouse
                .ComboList = "..."
                .FocusRect = flexFocusLight
            Else
                .Editable = flexEDKbdMouse
                .ComboList = ""
                .FocusRect = flexFocusLight
            End If
        ElseIf NewCol = COL_医生级别 Then
            .Editable = flexEDKbdMouse
            .FocusRect = flexFocusSolid
            .ComboList = .ColData(COL_医生级别)
        Else
            .Editable = flexEDNone
            .FocusRect = flexFocusNone
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call InputDepOrDoc(Row, Col)
End Sub

Private Sub vsDetail_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：删除数据行
    Dim i As Long
    
    If mintType = 2 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        With vsDetail
            If .TextMatrix(.Row, COL_邀请科室) <> "" Then
                If MsgBox("要删除当前行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                .RowData(.Row) = 0
                For i = 0 To .Cols - 1
                    .TextMatrix(.Row, i) = ""
                    .Cell(flexcpData, .Row, i) = ""
                Next
                mblnChange = True
            End If
            If Not (.Rows = .FixedRows + 1 And .Row = .FixedRows) Then .RemoveItem .Row
        End With
    End If
    If KeyCode > 127 Then
        Call vsDetail_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsDetail_DblClick()
    Call vsDetail_KeyPress(32)
End Sub

Private Sub vsDetail_GotFocus()
    Call vsDetail_AfterRowColChange(vsDetail.Row, vsDetail.Col, vsDetail.Row, vsDetail.Col)
End Sub

Private Sub vsDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：非回车确认完后编辑的处理(这里Text:=EditText,但ValidateEdit事件中还没有)
    If Not mblnReturn Then
        If Col = COL_邀请科室 Or Col = COL_邀请医生 Then
            vsDetail.TextMatrix(Row, Col) = CStr(vsDetail.Cell(flexcpData, Row, Col))
            Call vsDetail_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
        ElseIf Col = COL_医生级别 Then
            If Trim(vsDetail.TextMatrix(Row, Col)) <> "" Then
                vsDetail.RowData(Row) = 1
            End If
        End If
    End If
    If COL_代表科室 = Col Then Call Set代表科室(Row)
    mblnChange = True
End Sub

Private Sub InputDepOrDoc(ByVal lngRow As Long, lngCol As Long, Optional ByVal strEditText As String)
'功能：格表项目的录入， 邀请科室   邀请医生 ， strEditText 为空表格直接点的按钮。
    Dim strSQL As String, strLike As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strLikeText As String
    Dim lng科室id As String
    Dim str执业类别 As String
    
    On Error GoTo errH
    
    If Not (lngCol = COL_邀请科室 Or lngCol = COL_邀请医生) Then Exit Sub

    vPoint = zlControl.GetCoordPos(vsDetail.hwnd, vsDetail.CellLeft, vsDetail.CellTop)
    
    If lngCol = COL_邀请科室 Then
        strSQL = "select a.id,a.编码,a.名称,a.简码" & _
            " from 部门表 a,部门性质说明 b where a.id=b.部门id and b.工作性质='临床' and b.服务对象 in (2,3)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)"
    ElseIf lngCol = COL_邀请医生 Then
    
        lng科室id = Val(vsDetail.TextMatrix(lngRow, COL_邀请科室ID))
        str执业类别 = vsDetail.TextMatrix(lngRow, COL_医生级别)
        
        If lng科室id = 0 Then
            strSQL = "Select /*+ rule +*/ distinct a.Id, a.编号, a.姓名, a.简码" & vbNewLine & _
                "From 人员表 A, 部门人员 B, 部门表 C, 部门性质说明 D" & vbNewLine & _
                "Where a.Id = b.人员id And b.部门id = c.Id And c.Id = d.部门id And d.工作性质 = '临床' And d.服务对象 In (2, 3)" & vbNewLine & _
                "And (a.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or a.撤档时间 Is Null) And (c.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or c.撤档时间 Is Null)"
        Else
            strSQL = "select distinct a.id,a.编号,a.姓名,a.简码 from 人员表 a,部门人员 b where a.id=b.人员id And (a.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or a.撤档时间 Is Null) and b.部门id=[1]"
        End If
        
        If str执业类别 = "住院医师" Then
            strSQL = strSQL & " and a.专业技术职务 like '%医师'"
        ElseIf str执业类别 = "主治医师" Then
            strSQL = strSQL & " and a.专业技术职务='主治医师'"
        ElseIf str执业类别 = "(副)主任医师" Then
            strSQL = strSQL & " and a.专业技术职务 like '%主任医师'"
        End If
        
    End If
    
    '用于糊模查找的条件串
    If strEditText <> "" Then
        '优化
        strLike = gstrLike
        If Len(strEditText) < 2 Then strLike = ""
        
        If lngCol = COL_邀请科室 Then
            strLikeText = " And (A.编码 Like [2] Or A.名称 Like [3] or a.简码 like [4]) order by a.编码"
        Else
            strLikeText = " And (A.编号 Like [2] Or A.姓名 Like [3] or a.简码 like [4]) order by a.编号"
        End If
    Else
        If lngCol = COL_邀请科室 Then
            strLikeText = " and a.id<>[5] order by a.编码"
        Else
            strLikeText = "  order by a.编号"
        End If
    End If
    
    strSQL = strSQL & strLikeText
    
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "会诊申请单", False, "", "", False, False, True, vPoint.X, vPoint.Y, vsDetail.CellHeight, blnCancel, False, True, _
         lng科室id, UCase(strEditText) & "%", strLike & UCase(strEditText) & "%", strLike & UCase(strEditText) & "%", mlng开单科室ID)
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "未找到匹配项目！", vbInformation, gstrSysName
        End If
        vsDetail.TextMatrix(lngRow, lngCol) = CStr(vsDetail.Cell(flexcpData, lngRow, lngCol))
        Call vsDetail_AfterRowColChange(-1, -1, lngRow, lngCol) '重新使按钮可见
        Exit Sub
    End If
    
    If lngCol = COL_邀请科室 Then
        Call Set邀请科室(lngRow, rsTmp)
    ElseIf lngCol = COL_邀请医生 Then
        Call Set邀请医生(lngRow, rsTmp)
    End If
    vsDetail.RowData(lngRow) = 1
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Set代表科室(ByVal lngRow As Long)
'功能：明细表格中 代表科室 列复选框设置
    Dim i As Long
    
    With vsDetail
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, i, COL_代表科室) = flexUnchecked
        Next
        .Cell(flexcpChecked, lngRow, COL_代表科室) = flexChecked
    End With
End Sub

Private Sub vsDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strTmp As String
    If Asc("'") = KeyAscii Or Asc(";") = KeyAscii Or Asc("%") = KeyAscii Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        mblnReturn = True '标记是按回车确认编辑
        If optInfo(opt院内).value And (Col = COL_邀请科室 Or Col = COL_邀请医生) Then
            KeyAscii = 0
            strTmp = vsDetail.EditText
            Call InputDepOrDoc(Row, Col, strTmp)
            vsDetail.EditText = vsDetail.TextMatrix(Row, Col)
            If vsDetail.TextMatrix(Row, Col) <> "" Then Call LocatedNextCell(Row, Col)
        End If
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strTmp As String
    If optInfo(opt院内).value And (Col = COL_邀请科室 Or Col = COL_邀请医生) Then
        strTmp = vsDetail.EditText
        Call InputDepOrDoc(Row, Col, strTmp)
        If vsDetail.TextMatrix(Row, Col) <> "" Then
            mblnReturn = False
            Call LocatedNextCell(Row, Col)
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Set邀请科室(ByVal lngRow, ByVal rsTmp As ADODB.Recordset)
'功能：设置邀请科室
    Dim i As Long
    
    With vsDetail
        '检查重复输入
        For i = .FixedRows To .Rows - 1
            If i <> lngRow And Val(.TextMatrix(i, COL_邀请科室ID)) = Val(rsTmp!ID & "") Then
                MsgBox "该科室已经在邀请列表中。", vbInformation, gstrSysName
                .TextMatrix(lngRow, COL_邀请科室) = CStr(.Cell(flexcpData, lngRow, COL_邀请科室))
                Call vsDetail_AfterRowColChange(lngRow, COL_邀请科室, lngRow, COL_邀请科室) '重新使按钮可见
                .SetFocus
                Exit Sub
            End If
        Next
                   
        If Val(.TextMatrix(lngRow, COL_邀请科室ID)) <> Val(rsTmp!ID & "") Then
            .TextMatrix(lngRow, COL_邀请科室ID) = Val(rsTmp!ID & "")
            .TextMatrix(lngRow, COL_邀请科室) = rsTmp!名称 & ""
            .Cell(flexcpData, lngRow, COL_邀请科室) = rsTmp!名称 & ""
            
            .TextMatrix(lngRow, COL_邀请医生) = ""
            .Cell(flexcpData, lngRow, COL_邀请医生) = ""
            .TextMatrix(lngRow, COL_邀请医生ID) = ""
        End If
        .TextMatrix(lngRow, COL_邀请科室) = .Cell(flexcpData, lngRow, COL_邀请科室)
        If .TextMatrix(lngRow, COL_邀请科室) <> "" Then Call LocatedNextCell(lngRow, COL_邀请科室)
    End With
    
End Sub

Private Sub Set邀请医生(ByVal lngRow, ByVal rsTmp As ADODB.Recordset)
'功能：设置邀请医生
    With vsDetail
        If Val(.TextMatrix(lngRow, COL_邀请医生ID)) <> Val(rsTmp!ID & "") Then
            .TextMatrix(lngRow, COL_邀请医生ID) = Val(rsTmp!ID & "")
            .TextMatrix(lngRow, COL_邀请医生) = rsTmp!姓名 & ""
            .Cell(flexcpData, lngRow, COL_邀请医生) = rsTmp!姓名 & ""
        End If
        .TextMatrix(lngRow, COL_邀请医生) = .Cell(flexcpData, lngRow, COL_邀请医生)
        If .TextMatrix(lngRow, COL_邀请医生) <> "" Then Call LocatedNextCell(lngRow, COL_邀请医生)
    End With
End Sub

Private Sub LocatedNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：邀请明细表格处理完一个单元格后定位到下一个单元格
    Dim lngTmpRow As Long, lngTmpCol As Long
    Dim i As Long
    
    lngTmpRow = lngRow
    lngTmpCol = lngCol
    
    With vsDetail
        lngTmpCol = lngTmpCol + 1
        If lngTmpCol < .Cols Then
            If Not .ColHidden(lngTmpCol) Then
                .Col = lngTmpCol
                .Row = lngTmpRow
            Else
                Call LocatedNextCell(lngTmpRow, lngTmpCol)
            End If
        Else
            lngTmpCol = .FixedCols
            lngTmpRow = lngTmpRow + 1
            
            If lngTmpRow > .Rows - 1 Then
                i = .Rows - 1
                If Not (Trim(.TextMatrix(i, COL_邀请科室)) = "" And Trim(.TextMatrix(i, COL_医生级别)) = "" And Trim(.TextMatrix(i, COL_邀请医生)) = "") Then
                    .Rows = .Rows + 1
                Else
                    lngTmpRow = lngTmpRow - 1
                    lngTmpCol = COL_邀请医生
                End If
            End If
            
            .Col = lngTmpCol
            .Row = lngTmpRow
            
            If .ColHidden(lngTmpCol) Then Call LocatedNextCell(lngTmpRow, lngTmpCol)
        End If
        
        
        If .Rows - 1 = .Row Then
            i = .Rows - 1
            If Not (Trim(.TextMatrix(i, COL_邀请科室)) = "" And Trim(.TextMatrix(i, COL_医生级别)) = "" And Trim(.TextMatrix(i, COL_邀请医生)) = "") Then
                .Rows = .Rows + 1
            End If
        End If
        
        Call .ShowCell(.Row, .Col)
        Call vsDetail_AfterRowColChange(.Row, .Col, .Row, .Col)
    End With
End Sub

Private Sub vsDetail_KeyPress(KeyAscii As Integer)
    If mintType = 2 Then Exit Sub
    With vsDetail
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call LocatedNextCell(.Row, .Col)
        Else
            If .Col = COL_邀请科室 Or .Col = COL_邀请医生 Or COL_医生级别 Then
                .Editable = flexEDKbdMouse
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDetail_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
                
                If .Col = COL_医生级别 Then
                    .Editable = flexEDKbdMouse
                    .FocusRect = flexFocusSolid
                    .ComboList = .ColData(COL_医生级别)
                End If
                
            ElseIf .Col = COL_代表科室 Then
                Call Set代表科室(.Row)
            Else
                .Editable = flexEDNone
            End If
        End If
    End With
    mblnChange = True
End Sub

Private Sub vsDetail_LostFocus()
'失去焦点后处理空白行
    Dim i As Long
    With vsDetail
        For i = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(i, COL_邀请科室) = "" Then .RemoveItem i
        Next
        .AddItem ""
    End With
End Sub

Private Sub vsDetail_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDetail.EditSelStart = 0
    vsDetail.EditSelLength = zlCommFun.ActualLen(vsDetail.EditText)
End Sub

Private Sub vsDetail_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：开始编辑时认为没有按下回车
    If mintType = 2 Then
        Cancel = True
        Exit Sub
    End If
    mblnReturn = False
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
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
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置")
        objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        
        If mintType = 3 Or mintType = 4 Then
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MeetFinish, "完成会诊(&F)")
            objControl.IconId = 225
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MeetCancel, "取消完成(&C)")
            objControl.IconId = 3014
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_19", Me)
        Case conMenu_File_Preview: Call PrintApply(1)
        Case conMenu_File_Print: Call PrintApply(2)
        Case conMenu_Edit_Save  '保存
            If CheckData = False Then Exit Sub
            If SaveData() Then
                mblnOK = True
            End If
        Case conMenu_Tool_MeetFinish
            If mblnChange Then
                If CheckData = False Then Exit Sub
                If SaveData() Then
                    mblnOK = True
                End If
            End If
            If ExecuteMeet(0) = True Then
                mblnOK = True
            End If
        Case conMenu_Tool_MeetCancel
            If ExecuteMeet(1) = True Then
                mblnOK = True
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'''''''''
    Dim blnVisible As Boolean
    
    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Save
            Control.Enabled = mblnChange
        Case conMenu_Tool_MeetFinish
            Control.Enabled = mintType = 3
        Case conMenu_Tool_MeetCancel
            Control.Enabled = mintType = 4
    End Select
    Control.Visible = blnVisible
End Sub

Private Sub cmdInfo_Click(Index As Integer)
    Select Case Index
    Case cmd临床诊断
        Call GetPatiDiag
    Case cmd会诊科室
        Call GetItem会诊科室(1)
    Case cmd会诊医师
        Call GetItem会诊医师(1)
    End Select
End Sub

Private Sub GetPatiDiag()
'功能：获取术前诊断
    Dim str诊断 As String
    
    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, p住院医生站, mclsMipModule)
    End If
    
    If mclsDiagEdit.ShowDiagEdit(Me, mlngAdviceID, mlng病人ID, mlng主页ID, 2, mlng病人科室id, mstr诊断IDs, str诊断, 0, mlngAdviceID) Then
        txtInfo(txt临床诊断).Text = str诊断
        Call SeekNextCtl
    End If
End Sub

Private Sub GetDefaultPatiDiag()
'功能：获取病人默认的诊断，医嘱新增时调用
'说明：取来源为3首页填写的、如果是中医科，则先提取中医入院诊断，如果为空则再取西医入院诊断
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim bln中医科 As Boolean, i As Long
    Dim strIDs As String, str诊断 As String
    
    On Error GoTo errH
    
    strSQL = "select a.id,a.诊断类型 as 类型,a.诊断描述 as 内容 from 病人诊断记录 a where a.记录来源=3 And NVL(A.编码序号,1) = 1 and a.病人id=[1] and a.主页id=[2] order by a.诊断类型,a.诊断次序"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    If Not rsTmp.EOF Then
        rsTmp.Filter = "类型>10" '诊断中是否包含了中医诊断
        
        If Not rsTmp.EOF Then
            bln中医科 = Sys.DeptHaveProperty(mlng开单科室ID, "中医科")
            If Not bln中医科 Then rsTmp.Filter = 0
        Else
            rsTmp.Filter = 0
        End If

        For i = 1 To rsTmp.RecordCount
            strIDs = strIDs & "," & rsTmp!ID
            str诊断 = str诊断 & "," & rsTmp!内容
            rsTmp.MoveNext
        Next
        
        mstr诊断IDs = Mid(strIDs, 2)
        txtInfo(txt临床诊断).Text = Mid(str诊断, 2)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SeekNextCtl() As Boolean
'功能：定位到下一个焦点的控件上
    Call zlCommFun.PressKey(vbKeyTab)
    SeekNextCtl = True
End Function

Private Sub LoadPatiInfo()
'功能：提取病人基本信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '读取病人相关信息
    If mbytBaby = 0 Then
        strSQL = "Select c.住院号,C.姓名,c.当前病况 as 病情,C.性别,C.年龄, B.名称 As 科室,c.入院日期, C.出院病床 As 当前床号,c.险类,c.费别" & vbNewLine & _
            "From 部门表 B, 病案主页 C" & vbNewLine & _
            "Where C.出院科室id = B.Id And C.病人id = [1] And C.主页id = [2]"
    Else
        strSQL = "Select c.住院号,Nvl(q.婴儿姓名, c.姓名||'之婴'||q.序号) as 姓名,null as 病情,q.婴儿性别 as 性别,Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间) || '天' as 年龄, B.名称 As 科室,c.入院日期, C.出院病床 As 当前床号,c.险类,c.费别" & vbNewLine & _
            "From 部门表 B, 病案主页 C, 病人新生儿记录 Q" & vbNewLine & _
            "Where C.出院科室id = B.Id  And c.病人id = q.病人id And c.主页id = q.主页id And C.病人id = [1] And C.主页id = [2] And q.序号 = [3]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mbytBaby)
    
    If rsTmp.RecordCount > 0 Then
        txtInfo(txt住院号).Text = rsTmp!住院号 & ""
        txtInfo(txt姓名).Text = rsTmp!姓名 & ""
        txtInfo(txt性别).Text = rsTmp!性别 & ""
        txtInfo(txt科室).Text = rsTmp!科室 & ""
        txtInfo(txt床号).Text = rsTmp!当前床号 & ""
        txtInfo(txt年龄).Text = rsTmp!年龄 & ""
        txtInfo(txt病情).Text = rsTmp!病情 & ""
        mstr入院时间 = Format(rsTmp!入院日期 & "", "YYYY-MM-DD HH:mm")
        mint险类 = Val(rsTmp!险类 & "")
        mstr费别 = rsTmp!费别 & ""
    End If
    
    mstr上次转科时间 = Get上次转科日期
    
    txtInfo(txt申请医师).Text = UserInfo.姓名
    
    strSQL = "select 信息值 as 主治医师 from 病案主页从表 where 病人id=[1] and 主页id=[2] and 信息名=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, "主治医师")
    If Not rsTmp.EOF Then txtInfo(txt主治医师).Text = rsTmp!主治医师 & ""
    
    If mintType = 0 Then Call GetDefaultPatiDiag
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadData()
'功能：加载部份数据和参数
    Dim strTmp As String, strSQL As String
    Dim rsTmp As ADODB.Recordset, rs附项 As ADODB.Recordset, rs科室 As ADODB.Recordset
    Dim strIDs As String, i As Integer, lngRow As Long
    Dim datCur As Date
    Dim str会诊科室IDs As String
    Dim blnCanSave As Boolean
    Dim str生效时间 As String
    Dim str会诊诊断 As String
    
    Dim varArr As Variant
    
    On Error GoTo errH
    
    If Not mrsCard Is Nothing Then
        str会诊科室IDs = mrsCard!会诊科室IDs & ""
        blnCanSave = Val(mrsCard!是否保存 & "") = 1
        str生效时间 = mrsCard!生效时间 & ""
    End If
    datCur = zlDatabase.Currentdate
    txtInfo(txt完成时间).Text = ""
    If mintType > 0 Then
        If mlng病人ID = 0 Or mlng主页ID = 0 Or mlngNo = 0 Then
            strSQL = "select a.病人ID,a.主页ID,a.id,a.申请序号 from 病人医嘱记录 a where a.id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID)
            mlngNo = Val(rsTmp!申请序号 & "")
            mlng病人ID = Val(rsTmp!病人ID & "")
            mlng主页ID = Val(rsTmp!主页ID & "")
        End If
        
        If mintType = 3 Or mintType = 4 Then
            strSQL = "select a.发送号,a.执行部门ID from 病人医嘱发送 a where a.医嘱id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID)
            mlng执行科室ID = Val(rsTmp!执行部门ID & "")
            mlng发送号 = Val(rsTmp!发送号 & "")
        End If
        
        txtInfo(txtNO).Text = mlngNo
        
        Call LoadPatiInfo
        
    Else
    
        Call LoadPatiInfo
        '日期
        If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
            If mdatTurn <> CDate(0) Then datCur = mdatTurn - 1 / 24 / 60
            mbln补录 = True
        End If
        
        txtInfo(txt会诊时间).Text = Format(datCur, "YYYY-MM-DD HH:mm")
        
        txtInfo(txt申请医师).Text = UserInfo.姓名
        
        strSQL = "select 名称 as 申请科室 from 部门表 where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng开单科室ID)
        txtInfo(txt申请科室).Text = rsTmp!申请科室 & ""
        
        Call GetDefaultPatiDiag
        
        If str会诊科室IDs <> "" Then
            '设置会诊科室
            '加载几个默认的会诊邀请科室
            varArr = Split(str会诊科室IDs, ",")
            strSQL = "select 名称,id from 部门表 where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str会诊科室IDs)
            With vsDetail
                .Rows = .FixedRows
                For i = 0 To UBound(varArr)
                    If Val(varArr(i)) <> 0 Then
                        .AddItem ""
                        lngRow = .Rows - 1
                        rsTmp.Filter = "id=" & Val(varArr(i))
                        .TextMatrix(lngRow, COL_邀请科室) = rsTmp!名称 & ""
                        .Cell(flexcpData, lngRow, COL_邀请科室) = rsTmp!名称 & ""
                        .TextMatrix(lngRow, COL_邀请科室ID) = rsTmp!ID & ""
                    End If
                Next
            End With
        End If
        
        Exit Sub
    End If
    
    strSQL = "select a.id,to_char(a.开始执行时间,'yyyy-MM-dd hh24:mi') as 会诊时间,a.紧急标志,a.诊疗项目ID,a.执行科室ID,a.开嘱科室ID,a.开嘱医生" & _
        " from 病人医嘱记录 a where a.申请序号=[1] order by a.序号"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngNo)
    
    mlng项目ID = Val(rsTmp!诊疗项目ID & "")
    Call Cbo.Locate(cboItem, mlng项目ID, True)
    mlng开单科室ID = Val(rsTmp!开嘱科室id & "")
    For i = 1 To rsTmp.RecordCount
        strIDs = strIDs & "," & Val(rsTmp!ID & "")
        rsTmp.MoveNext
    Next
    rsTmp.MoveFirst
    strIDs = Mid(strIDs, 2)
    strSQL = "select 医嘱id,项目,内容 from 病人医嘱附件 where 医嘱id IN (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) order by 医嘱ID"
 
    Set rs附项 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
    rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊邀请医院'"
    If Not rs附项.EOF Then
        If rs附项!内容 & "" <> "" Then
            fraInfo(fra邀请医院).Visible = True
            txtInfo(txt邀请医院).Text = rs附项!内容 & ""
        End If
    End If
    
    optInfo(opt院内).value = rs附项.EOF
    optInfo(opt院外).value = Not rs附项.EOF
    
    With vsDetail
        .Rows = .FixedRows + rsTmp.RecordCount
        lngRow = .FixedRows - 1
        
        For i = 1 To rsTmp.RecordCount
            lngRow = 1 + lngRow
            .RowData(lngRow) = 1
            
            rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊邀请科室'"
            .TextMatrix(lngRow, COL_邀请科室) = rs附项!内容 & ""
            .Cell(flexcpData, lngRow, COL_邀请科室) = rs附项!内容 & ""
            .TextMatrix(lngRow, COL_邀请科室ID) = Val(rsTmp!执行科室ID & "")
            
            rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊医生级别'"
            
            .TextMatrix(lngRow, COL_医生级别) = rs附项!内容 & ""
            
            rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊邀请医生'"
            .TextMatrix(lngRow, COL_邀请医生) = rs附项!内容 & ""
            
            rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊代表科室'"
            If Not rs附项.EOF Then
                If .TextMatrix(lngRow, COL_邀请科室) = rs附项!内容 & "" Then
                    .Cell(flexcpChecked, lngRow, COL_代表科室) = flexChecked
                End If
                If mintType = 3 Then
                    mbln填意见 = Is代表科室(mlngAdviceID)
                End If
            Else
                .Cell(flexcpChecked, lngRow, COL_代表科室) = flexUnchecked
            End If
            
            If Val(rsTmp!ID & "") = mlngAdviceID Then
                optInfo(opt紧急).value = 1 = Val(rsTmp!紧急标志 & "")
                rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊范围'"
                If Not rs附项.EOF Then optInfo(opt院内).value = rs附项!内容 & "" = "院内"
                If optInfo(opt院外).value Then
                    rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊邀请医院'"
                    If Not rs附项.EOF Then txtInfo(txt邀请医院).Text = rs附项!内容 & ""
                End If
                rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊目的'"
                If Not rs附项.EOF Then txtInfo(txt会诊目的).Text = rs附项!内容 & ""
                
                rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='病情摘要'"
                If Not rs附项.EOF Then txtInfo(txt病情摘要).Text = rs附项!内容 & ""
                
                rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊意见'"
                If Not rs附项.EOF Then txtInfo(txt会诊意见).Text = rs附项!内容 & ""
                
                rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊完成时间'"
                If Not rs附项.EOF Then txtInfo(txt完成时间).Text = rs附项!内容 & ""
                
                rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊完成科室'"
                If Not rs附项.EOF Then txtInfo(txt会诊科室).Text = rs附项!内容 & ""
                
                rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊医生'"
                If Not rs附项.EOF Then txtInfo(txt会诊医师).Text = rs附项!内容 & ""
                
                rs附项.Filter = "医嘱ID=" & Val(rsTmp!ID & "") & " and 项目='会诊诊断'"
                If Not rs附项.EOF Then str会诊诊断 = rs附项!内容 & ""
            
                txtInfo(txt会诊时间).Text = rsTmp!会诊时间 & ""
                txtInfo(txt申请医师).Text = rsTmp!开嘱医生 & ""
                
                strSQL = "select 名称 as 申请科室 from 部门表 where id=[1]"
                Set rs科室 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!开嘱科室id & ""))
                txtInfo(txt申请科室).Text = rs科室!申请科室 & ""
                
                If mintType = 3 Then
                    txtInfo(txt完成时间).Text = IIF(txtInfo(txt完成时间).Text = "", Format(datCur, "YYYY-MM-DD HH:mm"), txtInfo(txt完成时间).Text)
                    txtInfo(txt会诊科室).Text = .TextMatrix(lngRow, COL_邀请科室)
                    txtInfo(txt会诊医师).Text = UserInfo.姓名
                End If
            End If
            rsTmp.MoveNext
        Next
    End With
    
    '读取诊断
    strTmp = ""
    mstr诊断IDs = GetAdviceDiag(mlngAdviceID, strTmp)
    txtInfo(txt临床诊断).Text = strTmp
    If str会诊诊断 <> "" Then
        txtInfo(txt临床诊断).Text = str会诊诊断
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Is代表科室(ByVal lng医嘱ID As Long) As Boolean
'功能：书写会诊议见时判断当前的科室是不是代表科室
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "select 1 from 病人医嘱记录 a,部门表 b where a.执行科室id=b.id and a.id=[1] and" & vbNewLine & _
        "exists (select 1 from 病人医嘱附件 c where c.医嘱id =[1] and c.项目='会诊代表科室'and b.名称=c.内容)"
        
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    
    If Not rsTmp.EOF Then
        Is代表科室 = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadItem会诊()
'功能：加载会诊项目到下拉列表，如果只有一个项目则隐藏下拉列表

    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
    
    strSQL = "Select id,名称,执行科室,计价性质 from 诊疗项目目录 where 类别='Z' and 操作类型='7' and (撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL) order by 编码"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Set mrsItem会诊 = zlDatabase.CopyNewRec(rsTmp)
    
    If rsTmp.EOF Then
        MsgBox "未找到会诊项目，请先到诊疗项目管理中创建该项目。", vbInformation, "会诊申请"
        mblnChange = False
        Unload Me
        Exit Sub
    Else
        If rsTmp.RecordCount = 1 Then
            fraInfo(fra会诊项目).Visible = False
            mlng项目ID = Val(rsTmp!ID & "")
        Else
            fraInfo(fra会诊项目).Visible = True
            With cboItem
                .Clear
                For i = 1 To rsTmp.RecordCount
                    .AddItem rsTmp!名称 & ""
                    .ItemData(.ListCount - 1) = Val(rsTmp!ID & "")
                    
                    If Val(rsTmp!ID & "") = mlng项目ID And mlng项目ID <> 0 Then
                        .ListIndex = .ListCount - 1
                    End If
                    rsTmp.MoveNext
                Next
                If .ListIndex = -1 Then .ListIndex = 0
            End With
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboItem_Click()
    mlng项目ID = Val(cboItem.ItemData(cboItem.ListIndex))
    If Visible Then mblnChange = True
End Sub

Private Sub cboItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SeekNextCtl
End Sub

Private Sub optInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextCtl
End Sub

Private Sub GetItem会诊科室(ByVal intType As Integer)
'功能：获取会诊科室
'参数：0 文本框按回车，1 点按钮
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    Dim i As Integer, strDoctor As String
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt会诊科室).Tag = txtInfo(txt会诊科室).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt会诊科室).Text = "" Then  '相当于是清除该项目
            txtInfo(txt会诊科室).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
       
    strDoctor = txtInfo(txt会诊医师).Text
    strInput = Trim(UCase(txtInfo(txt会诊科室).Text))    '传入的值存在前缀空格
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称 as 科室,A.简码 From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) And a.Id = b.部门id" & _
        IIF(intType = 0, " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])", "") & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) And B.工作性质='临床'" & _
        IIF(strDoctor <> "", " And a.id in (select x.部门id from 部门人员 X, 人员表 Y where x.人员id=y.id and y.姓名=[3])", "") & _
        " and a.id<>[4] Order by A.编码"
        
    vRect = zlControl.GetControlRect(txtInfo(txt会诊科室).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "科室", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(txt会诊科室).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", strDoctor, mlng开单科室ID)
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("没有找到匹配的科室!", vbInformation, gstrSysName)
            txtInfo(txt会诊科室).SetFocus
            zlControl.TxtSelAll txtInfo(txt会诊科室)
            Exit Sub
        End If
    Else
        txtInfo(txt会诊科室).Text = rsTmp!科室 & ""
        txtInfo(txt会诊科室).Tag = rsTmp!科室 & ""
        txtInfo(txt会诊科室).SetFocus
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItem会诊医师(ByVal intType As Integer)
'功能：获取会诊医师
'参数：intType 0 输入匹配，1 点击按钮
    Dim strInput As String, intIndex As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    intIndex = txt会诊医师
    
    If intType = 0 Then
        If txtInfo(intIndex).Tag = txtInfo(intIndex).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(intIndex).Text = "" Then '相当于是清除该项目
            txtInfo(intIndex).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    
    strSQL = "Select A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生'" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        IIF(intType = 0, " And (A.编号 Like [1] Or A.姓名 Like [2] Or A.简码 Like [2])", "") & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.编号"
    
    On Error GoTo errH
    
    strInput = Trim(UCase(txtInfo(intIndex).Text))
    vRect = zlControl.GetControlRect(txtInfo(intIndex).hwnd)
    
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医生", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(intIndex).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "未找到匹配的人员。", vbInformation, gstrSysName
        End If
        Call zlControl.TxtSelAll(txtInfo(intIndex))
        txtInfo(intIndex).SetFocus: Exit Sub
    Else
        txtInfo(intIndex).Text = rsTmp!姓名 & ""
        txtInfo(intIndex).Tag = rsTmp!姓名 & ""
        txtInfo(intIndex).SetFocus
        Call SeekNextCtl
        If Visible Then mblnChange = True
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
'按键事件，糊模查找
    If Asc("'") = KeyAscii Or Asc(";") = KeyAscii Or Asc("%") = KeyAscii Then
        KeyAscii = 0
    End If
    
    
    If KeyAscii = 13 Then
        Select Case Index
        Case txt会诊科室
            Call GetItem会诊科室(0)
        Case txt会诊医师
            Call GetItem会诊医师(0)
        End Select
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Select Case Index
        Case txt会诊科室
            Call GetItem会诊科室(1)
        Case txt会诊医师
            Call GetItem会诊医师(1)
        End Select
    End If
    
    Select Case Index
        Case txt会诊时间, txt会诊目的, txt邀请医院, txt临床诊断
            If KeyAscii = vbKeyReturn Then Call SeekNextCtl
    End Select
    
    Select Case Index
        Case txt会诊意见, txt会诊目的, txt病情摘要
            If KeyAscii = Asc("'") Then KeyAscii = 0
    End Select
    
End Sub

Private Sub cmdDate_Click(Index As Integer)
'功能：选择日期
    Dim lngIndex As Long
    
    If Index = cmd会诊时间 Then
        lngIndex = txt会诊时间
        
        dtpDate.Left = txtInfo(lngIndex).Left
        dtpDate.Top = cmdDate(Index).Top + cmdDate(Index).Height
        
    ElseIf Index = cmd完成时间 Then
    
        lngIndex = txt完成时间
        dtpDate.Left = cmdDate(Index).Left + cmdDate(Index).Width - dtpDate.Width
        dtpDate.Top = txtInfo(lngIndex).Top - dtpDate.Height + fraIdea.Top
    End If
    
    If IsDate(txtInfo(lngIndex).Text) Then
        dtpDate.value = CDate(txtInfo(lngIndex).Text)
    Else
        dtpDate.value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = lngIndex
    
    dtpDate.ZOrder 0
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
'日期合法性检查
    Dim strDate As String, intIndex As Integer
    
    intIndex = Val(dtpDate.Tag)
    
    If intIndex = txt会诊时间 Then
        '取值
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '判断时间合法性
        If Not Check开始时间(strDate) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txtInfo(intIndex).Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txtInfo_Validate(intIndex, False) '更新数据
        txtInfo(intIndex).SetFocus
        If Visible Then mblnChange = True
    ElseIf intIndex = txt完成时间 Then
        '取值
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '判断会诊完成时间合法性
        If Not Check完成时间(strDate, txtInfo(txt会诊时间).Text) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txtInfo(intIndex).Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txtInfo_Validate(intIndex, False) '更新数据
        txtInfo(intIndex).SetFocus
        If Visible Then mblnChange = True
    End If
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    Dim intIndex As Integer
    
    If KeyAscii = vbKeyEscape Then
        intIndex = Val(dtpDate.Tag)
        If intIndex >= 0 Then txtInfo(intIndex).SetFocus
        dtpDate.Tag = ""
        dtpDate.Visible = False
    End If
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Select Case Index
        Case txt会诊时间
            Call cmdDate_Click(cmd会诊时间)
        Case txt完成时间
            Call cmdDate_Click(cmd完成时间)
        End Select
    End If
End Sub

Private Function Check开始时间(ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的开始时间是否合法   即界面－－生效时间
'说明：
'1.开始时间不能小于病人的入院时间
'2.开始时间不能小于病人的转科时间
'3.开始时间必须小于终止时间
'4.正常录入时,开始时间不能小于当前时间之前30分钟(从而可能造成开嘱时间大于开始时间30分钟)
'5.补录的医嘱开始时间不能大于当前时间，转科补录不能大于转科开始时间
    Dim strInDate As String, blnOut As Boolean
        
    If Not IsDate(strStart) Then
        MsgBox "输入的会诊时间无效。", vbInformation, gstrSysName
        Exit Function
    End If

    strInDate = mstr入院时间
    If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
        strMsg = "医嘱的会诊时间不能小于病人的入院时间 " & strInDate & " 。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    strInDate = ""
    If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
        strInDate = Format(mdatTurn, "yyyy-MM-dd HH:mm")
    ElseIf IsDate(mstr上次转科时间) Then
        strInDate = mstr上次转科时间
    End If
    If strInDate <> "" Then
        If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
            If Format(strStart, "yyyy-MM-dd HH:mm") >= strInDate Then
                strMsg = "医嘱的会诊时间应小于病人" & IIF(mintPState = ps最近转出, "转出", IIF(mintPState = ps预出, "预出院", "出院")) & "的时间 " & strInDate & " 。"
                If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
                strMsg = "医嘱的会诊时间不能小于病人最近的转科时间 " & strInDate & " 。"
                If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
   
    Check开始时间 = True
End Function

Private Function Check完成时间(ByVal strDate As String, ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查会诊完成时间

    Dim strInDate As String, strDateType As String
    
    If Not IsDate(strDate) Then
        strMsg = "输入的时间无效。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") < Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = "会诊完成时间不能小于医嘱会诊时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check完成时间 = True
End Function

Private Function Get上次转科日期() As String
'功能：读取上次转科时间
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select 开始时间 From 病人变动记录 Where 开始时间 is Not NULL And 开始原因=3 And 病人ID=[1] And 主页ID=[2] Order by 开始时间 desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.RecordCount > 0 Then Get上次转科日期 = Format(rsTmp!开始时间 & "", "YYYY-MM-DD HH:mm")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub PrintApply(ByVal intType As Integer)
'功能打印预览申请单
'参数：intType:1-预览，2-打印
    '判断如果还未保存则先保存再打印
    If mintType <> 2 Then
        If mblnChange Then
            If CheckData = False Then Exit Sub
            If SaveData() Then
                mblnOK = True
            End If
        Else
            '如果不可用，则检查医嘱是否符合
            If CheckData = False Then Exit Sub
        End If
    End If
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_19", Me, "医嘱ID=" & mlngAdviceID, "申请序号=" & mlngNo, intType)
End Sub

Private Function CheckData() As Boolean
'功能：检查数据正确性
    Dim strIDs As String, str医嘱内容 As String, strMsg As String
    Dim lngTmp As Long, i As Integer
    Dim vMsg As VbMsgBoxResult
    Dim intCount As Integer
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim bln中医 As Boolean
    Dim str类型 As String
    Dim strTmp As String
    Dim strTabAdvice As String
    Dim rsPrice As ADODB.Recordset
    
    If mintType < 2 Then
        'Call SeekNextControl  '用这种方式会出问题71290
        '这里采用两次设不同控件的焦点，确保validata事件的执行。
        txtInfo(txt会诊时间).SetFocus
        txtInfo(txt会诊目的).SetFocus
        
        '检查时间合法性
        If Not Check开始时间(txtInfo(txt会诊时间).Text) Then
        If txtInfo(txt会诊时间).Enabled Then txtInfo(txt会诊时间).SetFocus
            Exit Function
        End If
        
        '判断是否是补录医嘱
        If DateDiff("n", CDate(txtInfo(txt会诊时间).Text), CDate(zlDatabase.Currentdate)) > gint补录间隔 _
            Or mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
            mbln补录 = True
        Else
            mbln补录 = False
        End If
        
        If optInfo(opt院外).value And txtInfo(txt邀请医院).Text = "" Then
            MsgBox "您选择的院外会诊，请填写邀请医院。", vbInformation, Me.Caption
            txtInfo(txt邀请医院).SetFocus
            Exit Function
        End If
           
        '会诊代表科室确定
        If mbln书写要求 Then
             lngTmp = 0: strMsg = ""
             With vsDetail
                 For i = .FixedRows To .Rows - 1
                     If .Cell(flexcpChecked, i, COL_代表科室) = flexChecked Then
                         lngTmp = lngTmp + 1
                     End If
                     '如果只有一个科室时，则默认勾上不提示
                     If .TextMatrix(i, COL_邀请科室) <> "" Then
                         intCount = intCount + 1
                     End If
                 Next
             End With
             
             If intCount = 1 Then
                 vsDetail.Cell(flexcpChecked, vsDetail.FixedRows, COL_代表科室) = flexChecked
             Else
                 If lngTmp = 0 Then
                     strMsg = "请确定会诊代表科室！"
                 ElseIf lngTmp > 1 Then
                     strMsg = "会诊代表科室只能有一个！"
                 End If
             End If
             
             If strMsg <> "" Then
                 MsgBox strMsg, vbInformation, Me.Caption
                 vsDetail.SetFocus
                 Exit Function
             End If
        End If
        
        If txtInfo(txt会诊目的).Text = "" Then
            MsgBox "没有填写会诊目的。", vbInformation, Me.Caption
            txtInfo(txt会诊目的).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt病情摘要).Text = "" Then
            MsgBox "没有填写病情摘要。", vbInformation, Me.Caption
            txtInfo(txt病情摘要).SetFocus
            Exit Function
        End If
        
        '诊断检查
        If InStr(mstr检查入院诊断, "Z") > 0 Then
            bln中医 = Sys.DeptHaveProperty(mlng病人科室id, "中医科")
            str类型 = IIF(bln中医, "2,12", "2")
            If Not ExistsDiagNoses(mlng病人ID, mlng主页ID, str类型) Then
                strMsg = "病人的入院诊断还没有输入，请先输入病人的入院诊断再下达会诊申请。"
            End If
            If strMsg <> "" Then
                MsgBox strMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '对码检查
        str医嘱内容 = Get项目名称
        
        strTmp = mlng项目ID & "||2"
        mstr摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", strTmp)
            
        With vsDetail
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_邀请科室ID)) <> 0 And optInfo(opt院内).value Then
                    strIDs = strIDs & "," & mlng项目ID & ":" & Val(.TextMatrix(i, COL_邀请科室ID))
                    
                    strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as 结果 From Dual"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", 2, mlng病人ID, mlng主页ID, mint险类, 1, _
                         "Z", mlng项目ID, mlng开单科室ID, UserInfo.姓名, Val(.TextMatrix(i, COL_邀请科室ID)), 0, 0, 0, mstr摘要)
                    
                    If Not rsTmp.EOF Then
                        strMsg = NVL(rsTmp!结果)
                        If strMsg <> "" Then
                            Select Case Val(Split(strMsg, "|")(0))
                            Case 1 '提示
                                If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    strMsg = "": Exit Function
                                End If
                            Case 2 '禁止
                                MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                                strMsg = "": Exit Function
                            End Select
                            strMsg = ""
                        End If
                    End If
                    
                    strTabAdvice = strTabAdvice & " Union ALL " & _
                        "select " & i & " as ID," & i & " as 序号,-null as 相关ID,'Z' as 诊疗类别," & mlng项目ID & " as 管码项目ID," & _
                            mlng项目ID & " as 诊疗项目ID,1 As 总量, 0 As 单量,null as 标本部位,null As 检查方法," & _
                            "0 as 执行标记,0 as 计价特性, null As 附加手术," & Val(mrsItem会诊!执行科室 & "") & " As 执行性质," & Val(.TextMatrix(i, COL_邀请科室ID)) & " as 执行科室id from dual"
 
                ElseIf .TextMatrix(i, COL_邀请科室) <> "" Then
                    strIDs = strIDs & .TextMatrix(i, COL_邀请科室)
                End If
            Next
            If strIDs = "" Then
                MsgBox "没有填写邀请科室。", vbInformation, Me.Caption
                .SetFocus
                Exit Function
            End If
            If optInfo(opt院外).value Then strIDs = ""
        End With
        
        strIDs = Mid(strIDs, 2)
     
        If gint医保对码 = 2 Then mbln提醒对码 = True
    
        strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, 2, "", strIDs, str医嘱内容)
        
        If strMsg <> "" Then
            If gint医保对码 = 1 Then
                vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", Me)
                If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                If vMsg = vbIgnore Then mbln提醒对码 = False
            ElseIf gint医保对码 = 2 Then
                MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        strTabAdvice = Mid(strTabAdvice, 12)
        '医保管控实时监测
        If mint险类 <> 0 And strTabAdvice <> "" Then
            If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) Then
                If MakePriceRecord申请单("52", mlng病人ID, mlng主页ID, strTabAdvice, strIDs, mstr费别, mlng开单科室ID, rsPrice) Then
                    If Not gclsInsure.CheckItem(mint险类, 1, 0, rsPrice) Then
                        MsgBox "医保监测检查未通(执行Insure.CheckItem接口)，本次下达的会诊申请单不能保存。", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
    ElseIf mintType = 3 Then
        If Not Check完成时间(txtInfo(txt完成时间).Text, txtInfo(txt会诊时间).Text) Then
            If txtInfo(txt完成时间).Enabled Then txtInfo(txt完成时间).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt会诊意见).Text = "" Then
            MsgBox "没有填写会诊意见。", vbInformation, Me.Caption
            txtInfo(txt会诊意见).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt会诊科室).Text = "" Then
            MsgBox "没有填写会诊科室。", vbInformation, Me.Caption
            txtInfo(txt会诊科室).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt会诊医师).Text = "" Then
            MsgBox "没有填写会诊医师。", vbInformation, Me.Caption
            txtInfo(txt会诊医师).SetFocus
            Exit Function
        End If
    End If
    CheckData = True
End Function

Private Function SaveData() As Boolean
'功能：保存数据
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim lng医嘱ID As Long, lng医嘱序号 As Long, lng申请序号 As Long
    Dim i As Long, str内容 As String, int紧急 As Integer, lng执行科室ID As Long, lng要素ID As Long
    Dim int计价特性 As Integer, int执行性质 As Integer, int附项序号 As Integer
    Dim datCur As Date, str开嘱时间 As String, strSQL As String, strSource As String
    Dim rs要素 As ADODB.Recordset, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    arrSQL = Array()
    
    strSQL = "Select b.Id As 要素id, b.中文名 As 名称 from 诊治所见项目 B where b.中文名 in " & _
      " ('会诊范围','会诊邀请医院','会诊邀请科室','会诊医生级别','会诊代表科室','会诊邀请医生'," & _
      " '会诊诊断','会诊目的','会诊意见','会诊完成时间','会诊完成科室','会诊医生','病历摘要')"
    Set rs要素 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
    If mintType = 0 Or mintType = 1 Then
        
        txtInfo(txtNO).Text = mlngNo
        
        If mintType = 1 Then '如果是修改医嘱申请科室和申请医生变成当前的
            txtInfo(txt申请医师).Text = UserInfo.姓名
            strSQL = "select 名称 as 申请科室 from 部门表 where id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng开单科室ID)
            txtInfo(txt申请科室).Text = rsTmp!申请科室 & ""
        End If
        
        If mlngNo <> 0 Then
            strSQL = "select a.id from 病人医嘱记录 a where a.申请序号=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngNo)
            For i = 1 To rsTmp.RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & rsTmp!ID & ",1)"
                rsTmp.MoveNext
            Next
        End If
        
        If mbln补录 Then
            int紧急 = 2
        Else
            If optInfo(opt紧急).value Then int紧急 = 1
        End If
        
        If mlngNo <> 0 Then
            lng申请序号 = mlngNo
        Else
            lng申请序号 = Get申请序号
            mlngNo = lng申请序号
        End If
        
        lng医嘱序号 = GetMaxAdviceNO(mlng病人ID, mlng主页ID, mbytBaby)
        
        mrsItem会诊.Filter = "ID=" & mlng项目ID
        str内容 = mrsItem会诊!名称 & ""
        
        int计价特性 = Val(mrsItem会诊!计价性质 & "")
        int执行性质 = Val(mrsItem会诊!执行科室 & "")
        
        datCur = zlDatabase.Currentdate
        str开嘱时间 = IIF(datCur > CDate(txtInfo(txt会诊时间).Text), txtInfo(txt会诊时间).Text, datCur)
        
        '共3个项目依次为：[ID],[序号],[执行科室ID]
        strSource = "ZL_病人医嘱记录_Insert([0],NULL,[1],2," & mlng病人ID & "," & mlng主页ID & "," & mbytBaby & ",1,1,'Z'," & mlng项目ID & ",NULL,NULL,NULL,1,'" & str内容 & "',NULL," & _
            "NULL,'一次性',NULL,NULL,NULL,NULL," & ZVal(int计价特性) & ",[2]," & ZVal(int执行性质) & "," & int紧急 & "," & _
            "To_Date('" & Format(txtInfo(txt会诊时间).Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
            mlng病人科室id & "," & mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & _
            "To_Date('" & Format(str开嘱时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
            "NULL," & ZVal(mlng前提ID) & ",NULL,0,NULL," & IIF(mstr摘要 = "", "null", "'" & mstr摘要 & "'") & ",'" & UserInfo.姓名 & "',Null,NULL,NULL,NULL," & lng申请序号 & ")"
                
        With vsDetail
            '会诊代表科的确定
            str内容 = ""
            If mbln书写要求 Then
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpChecked, i, COL_代表科室) = flexChecked Then
                        str内容 = .TextMatrix(i, COL_邀请科室)
                        Exit For
                    End If
                Next
            End If
            
            For i = .FixedRows To .Rows - 1
                lng执行科室ID = Val(.TextMatrix(i, COL_邀请科室ID))
                int附项序号 = 0
                If optInfo(opt院外).value Then lng执行科室ID = mlng病人科室id
                If .TextMatrix(i, COL_邀请科室) <> "" Then
                
                    lng医嘱序号 = lng医嘱序号 + 1
                    lng医嘱ID = zlDatabase.GetNextID("病人医嘱记录")
                    
                    strSQL = GetStrExcSQL(strSource, lng医嘱ID, lng医嘱序号, ZVal(lng执行科室ID))
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int附项序号 = int附项序号 + 1
                    lng要素ID = Get会诊要素ID(rs要素, "会诊范围")
                    strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊范围',0," & int附项序号 & "," & lng要素ID & ",'" & IIF(optInfo(opt院外).value, "院外", "院内") & "',1)"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    If optInfo(opt院外).value Then
                        int附项序号 = int附项序号 + 1
                        lng要素ID = Get会诊要素ID(rs要素, "会诊邀请医院")
                        strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊邀请医院',0," & int附项序号 & "," & lng要素ID & ",'" & txtInfo(txt邀请医院).Text & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    End If
                    
                    If str内容 <> "" Then
                        int附项序号 = int附项序号 + 1
                        lng要素ID = Get会诊要素ID(rs要素, "会诊代表科室")
                        strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊代表科室',0," & int附项序号 & "," & lng要素ID & ",'" & str内容 & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    End If
                    
                    int附项序号 = int附项序号 + 1
                    lng要素ID = Get会诊要素ID(rs要素, "会诊诊断")
                    strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊诊断',0," & int附项序号 & "," & lng要素ID & ",'" & txtInfo(txt临床诊断).Text & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int附项序号 = int附项序号 + 1
                    lng要素ID = Get会诊要素ID(rs要素, "会诊目的")
                    strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊目的',0," & int附项序号 & "," & lng要素ID & ",'" & txtInfo(txt会诊目的).Text & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int附项序号 = int附项序号 + 1
                    lng要素ID = Get会诊要素ID(rs要素, "病历摘要")
                    strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'病情摘要',0," & int附项序号 & "," & lng要素ID & ",'" & txtInfo(txt病情摘要).Text & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int附项序号 = int附项序号 + 1
                    lng要素ID = Get会诊要素ID(rs要素, "会诊邀请科室")
                    strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊邀请科室',0," & int附项序号 & "," & lng要素ID & ",'" & .TextMatrix(i, COL_邀请科室) & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int附项序号 = int附项序号 + 1
                    lng要素ID = Get会诊要素ID(rs要素, "会诊医生级别")
                    strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊医生级别',0," & int附项序号 & "," & lng要素ID & ",'" & .TextMatrix(i, COL_医生级别) & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int附项序号 = int附项序号 + 1
                    lng要素ID = Get会诊要素ID(rs要素, "会诊邀请医生")
                    strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊邀请医生',0," & int附项序号 & "," & lng要素ID & ",'" & .TextMatrix(i, COL_邀请医生) & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    '诊断关联信息
                    If mstr诊断IDs <> "" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(" & lng医嘱ID & ",'" & mstr诊断IDs & "')"
                    End If
                End If
            Next
            mlngAdviceID = lng医嘱ID
        End With
    ElseIf mintType = 3 Then '书写会诊意见
        int附项序号 = 0
        lng医嘱ID = mlngAdviceID
        
        strSQL = "select a.项目,a.要素id,a.内容 from 病人医嘱附件 a where a.医嘱id=[1] order by a.排列"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
        
        For i = 1 To rsTmp.RecordCount
            If rsTmp!项目 & "" = "会诊意见" Then Exit For
            int附项序号 = int附项序号 + 1
            strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'" & rsTmp!项目 & "',0," & int附项序号 & "," & rsTmp!要素ID & ",'" & rsTmp!内容 & "'" & IIF(i = 1, ",1)", ")")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
            rsTmp.MoveNext
        Next
        
        int附项序号 = int附项序号 + 1
        lng要素ID = Get会诊要素ID(rs要素, "会诊意见")
        strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊意见',0," & int附项序号 & "," & lng要素ID & ",'" & txtInfo(txt会诊意见).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        int附项序号 = int附项序号 + 1
        lng要素ID = Get会诊要素ID(rs要素, "会诊完成时间")
        strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊完成时间',0," & int附项序号 & "," & lng要素ID & ",'" & txtInfo(txt完成时间).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        int附项序号 = int附项序号 + 1
        lng要素ID = Get会诊要素ID(rs要素, "会诊完成科室")
        strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊完成科室',0," & int附项序号 & "," & lng要素ID & ",'" & txtInfo(txt会诊科室).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        int附项序号 = int附项序号 + 1
        lng要素ID = Get会诊要素ID(rs要素, "会诊医生")
        strSQL = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'会诊医生',0," & int附项序号 & "," & lng要素ID & ",'" & txtInfo(txt会诊医师).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
       
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    SaveData = True
    mblnChange = False
    
    If mintType = 0 Or mintType = 1 Then
        Call ZLHIS_CIS_001(mclsMipModule, mlng病人ID, txtInfo(txt姓名).Text, txtInfo(txt住院号).Text, , 2, _
            mlng主页ID, mlng病区ID, , mlng病人科室id, "", , txtInfo(txt床号).Text, _
            lng医嘱ID, IIF(int紧急 = 1, 1, 0), 1, "Z", "", UserInfo.姓名, Format(str开嘱时间, "yyyy-MM-dd HH:mm:ss"), mlng开单科室ID, "", , , "")
    End If
    
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetStrExcSQL(ByVal strSource As String, ParamArray arrInput() As Variant) As String
'功能：生成 ZL_病人医嘱记录_Insert过程语句，arrInput参数 [ID],[序号],[执行科室ID]
    Dim i As Integer, strTmp As String, strResult As String
    
    strResult = strSource
    For i = 0 To UBound(arrInput)
        strTmp = arrInput(i)
        strResult = Replace(strResult, "[" & i & "]", strTmp)
    Next
    GetStrExcSQL = strResult
End Function
 
Private Function Get项目名称() As String
    mrsItem会诊.Filter = 0
    mrsItem会诊.Filter = "ID=" & mlng项目ID
    Get项目名称 = mrsItem会诊!名称 & ""
End Function

Private Function Get会诊要素ID(ByRef rsIn As ADODB.Recordset, ByVal str名称 As String) As Long
'功能：获取会诊固定要素
    Dim strSQL As String
    
    On Error GoTo errH
    
    rsIn.Filter = "名称='" & str名称 & "'"
    If Not rsIn.EOF Then
        Get会诊要素ID = Val(rsIn!要素ID & "")
    End If
  
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteMeet(ByVal intType As Integer) As Boolean
'功能：完成对当前病的人会诊
'参数：intType 0－完成会诊，1－取消息会诊
    Dim strSQL As String
    
    If MsgBox("确实要" & IIF(intType = 1, "取消", "") & "完成对该""" & txtInfo(txt姓名).Text & """的会诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    If intType = 0 Then
        strSQL = "ZL_病人医嘱执行_Finish(" & mlngAdviceID & "," & mlng发送号 & ",NULL,0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlng执行科室ID & ")"
    Else
        strSQL = "ZL_病人医嘱执行_Cancel(" & mlngAdviceID & "," & mlng发送号 & ",Null,0," & mlng执行科室ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    End If
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    mintType = IIF(mintType = 3, 4, 3)
    Call SetFormState
    ExecuteMeet = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get申请序号() As Long
'功能：获取申请序号
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "Select 病人医嘱记录_申请序号.Nextval as 申请序号 From Dual"
    
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Get申请序号 = Val(rsTmp!申请序号 & "")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub optInfo_Click(Index As Integer)
    Select Case Index
    Case opt院内
        fraInfo(fra邀请医院).Visible = False
        txtInfo(txt邀请医院).Text = ""
        vsDetail.Rows = vsDetail.FixedRows
        vsDetail.Rows = vsDetail.FixedRows + 1
    Case opt院外
        fraInfo(fra邀请医院).Visible = True
        txtInfo(txt邀请医院).Locked = False
        txtInfo(txt邀请医院).TabStop = True
        txtInfo(txt邀请医院).Text = ""
        vsDetail.Rows = vsDetail.FixedRows
        vsDetail.Rows = vsDetail.FixedRows + 1
    End Select
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
'合法性检查和值的恢复
    If mintType = 0 Then Exit Sub
    
    Select Case Index
    Case txt会诊时间
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(Index).Text) Then
                    '恢复人为的清除缺省为开始时间
                    txtInfo(Index).Text = txtInfo(Index).Text
                End If
            End If
        Else
            '检查时间合法性
            If Not Check完成时间(txtInfo(Index).Text, txtInfo(Index).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
        End If
        '判断是否是补录医嘱
        If DateDiff("n", CDate(txtInfo(Index).Text), CDate(zlDatabase.Currentdate)) > gint补录间隔 _
            Or mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
            mbln补录 = True
        Else
            mbln补录 = False
        End If
    Case txt完成时间
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(Index).Tag) Then
                    '恢复人为的清除
                    txtInfo(Index).Text = txtInfo(Index).Tag
                End If
            End If
        Else
            '检查时间合法性
            If Not Check开始时间(txtInfo(Index).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
            
        End If
    
    Case txt会诊科室, txt会诊医师
        If txtInfo(Index).Tag <> txtInfo(Index).Text And txtInfo(Index).Tag <> "" Then txtInfo(Index).Text = txtInfo(Index).Tag
    End Select
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub SetItemEditable(Optional int会诊时间 As Integer, Optional int会诊性质 As Integer, _
    Optional int会诊项目 As Integer, Optional int会诊范围 As Integer, Optional int邀请医院 As Integer, Optional int邀请明细 As Integer, Optional int临床诊断 As Integer, _
    Optional int会诊目的 As Integer, Optional int病情摘要 As Integer, _
    Optional int会诊意见 As Integer, Optional int完成时间 As Integer, _
    Optional int会诊科室 As Integer, Optional int会诊医师 As Integer)
'功能：设置指定编辑项的可用状态
'参数：0-保持不变,-1-锁定,1-允许
        
    If int会诊时间 = 1 Then
        txtInfo(txt会诊时间).Locked = False
        txtInfo(txt会诊时间).TabStop = True
        txtInfo(txt会诊时间).BackColor = vbWindowBackground
        cmdDate(cmd会诊时间).Enabled = True
    ElseIf int会诊时间 = -1 Then
        txtInfo(txt会诊时间).Locked = True
        txtInfo(txt会诊时间).TabStop = False
        txtInfo(txt会诊时间).BackColor = vbButtonFace
        cmdDate(cmd会诊时间).Enabled = False
    End If
    
    If int会诊性质 = 1 Then
        fraInfo(fra会诊性质).Enabled = True
    ElseIf int会诊性质 = -1 Then
        fraInfo(fra会诊性质).Enabled = False
    End If
    
    If int会诊项目 = 1 Then
        cboItem.Locked = False
        cboItem.TabStop = True
        cboItem.BackColor = vbWindowBackground
    ElseIf int会诊项目 = -1 Then
        cboItem.Locked = True
        cboItem.TabStop = False
        cboItem.BackColor = vbButtonFace
    End If
    
    If int会诊范围 = 1 Then
        fraInfo(fra会诊范围).Enabled = True
    ElseIf int会诊性质 = -1 Then
        fraInfo(fra会诊范围).Enabled = False
    End If
    
    If int邀请医院 = 1 Then
        txtInfo(txt邀请医院).Locked = False
        txtInfo(txt邀请医院).TabStop = True
        txtInfo(txt邀请医院).BackColor = vbWindowBackground
    ElseIf int邀请医院 = -1 Then
        txtInfo(txt邀请医院).Locked = True
        txtInfo(txt邀请医院).TabStop = False
        txtInfo(txt邀请医院).BackColor = vbButtonFace
    End If
    
    If int邀请明细 = 1 Then
        vsDetail.TabStop = True
        fraDetail.Enabled = True
    ElseIf int邀请明细 = -1 Then
        vsDetail.TabStop = False
        vsDetail.Editable = flexEDNone
        fraDetail.Enabled = False
    End If
    
    If int临床诊断 = 1 Then
        txtInfo(txt临床诊断).Locked = False
        txtInfo(txt临床诊断).TabStop = True
        txtInfo(txt临床诊断).BackColor = vbWindowBackground
        cmdInfo(cmd临床诊断).Enabled = True
    ElseIf int临床诊断 = -1 Then
        txtInfo(txt临床诊断).Locked = True
        txtInfo(txt临床诊断).TabStop = False
        txtInfo(txt临床诊断).BackColor = vbButtonFace
        cmdInfo(cmd临床诊断).Enabled = False
    End If
    
    If int会诊目的 = 1 Then
        txtInfo(txt会诊目的).Locked = False
        txtInfo(txt会诊目的).TabStop = True
        txtInfo(txt会诊目的).BackColor = vbWindowBackground
    ElseIf int会诊目的 = -1 Then
        txtInfo(txt会诊目的).Locked = True
        txtInfo(txt会诊目的).TabStop = False
        txtInfo(txt会诊目的).BackColor = vbButtonFace
    End If
    
    
    If int会诊意见 = 1 Then
        txtInfo(txt会诊意见).Locked = False
        txtInfo(txt会诊意见).TabStop = True
        txtInfo(txt会诊意见).BackColor = vbWindowBackground
    ElseIf int会诊意见 = -1 Then
        txtInfo(txt会诊意见).Locked = True
        txtInfo(txt会诊意见).TabStop = False
        txtInfo(txt会诊意见).BackColor = vbButtonFace
    End If
    
    If int病情摘要 = 1 Then
        txtInfo(txt病情摘要).Locked = False
        txtInfo(txt病情摘要).TabStop = True
        txtInfo(txt病情摘要).BackColor = vbWindowBackground
    ElseIf int病情摘要 = -1 Then
        txtInfo(txt病情摘要).Locked = True
        txtInfo(txt病情摘要).TabStop = False
        txtInfo(txt病情摘要).BackColor = vbButtonFace
    End If
    
    
    If int完成时间 = 1 Then
        txtInfo(txt完成时间).Locked = False
        txtInfo(txt完成时间).TabStop = True
        txtInfo(txt完成时间).BackColor = vbWindowBackground
        cmdDate(cmd完成时间).Enabled = True
    ElseIf int完成时间 = -1 Then
        txtInfo(txt完成时间).Locked = True
        txtInfo(txt完成时间).TabStop = False
        txtInfo(txt完成时间).BackColor = vbButtonFace
        cmdDate(cmd完成时间).Enabled = False
    End If
    
    
    If int会诊科室 = 1 Then
        txtInfo(txt会诊科室).Locked = False
        txtInfo(txt会诊科室).TabStop = True
        txtInfo(txt会诊科室).BackColor = vbWindowBackground
        cmdInfo(cmd会诊科室).Enabled = True
    ElseIf int会诊科室 = -1 Then
        txtInfo(txt会诊科室).Locked = True
        txtInfo(txt会诊科室).TabStop = False
        txtInfo(txt会诊科室).BackColor = vbButtonFace
        cmdInfo(cmd会诊科室).Enabled = False
    End If
    
    
    If int会诊医师 = 1 Then
        txtInfo(txt会诊医师).Locked = False
        txtInfo(txt会诊医师).TabStop = True
        txtInfo(txt会诊医师).BackColor = vbWindowBackground
        cmdInfo(cmd会诊医师).Enabled = True
    ElseIf int会诊医师 = -1 Then
        txtInfo(txt会诊医师).Locked = True
        txtInfo(txt会诊医师).TabStop = False
        txtInfo(txt会诊医师).BackColor = vbButtonFace
        cmdInfo(cmd会诊医师).Enabled = False
    End If
     
End Sub

