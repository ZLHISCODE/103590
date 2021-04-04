VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApplyOperation 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "手术申请单"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10650
   Icon            =   "frmApplyOperation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   27
      Left            =   5490
      ScaleHeight     =   255
      ScaleWidth      =   1305
      TabIndex        =   72
      Top             =   2565
      Width           =   1300
      Begin VB.ComboBox cboInfo 
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
         Index           =   27
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   -25
         Width           =   1200
      End
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
      Height          =   255
      Index           =   17
      Left            =   5325
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   6885
      Width           =   1260
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   17
      Left            =   6705
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   6885
      Width           =   270
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8715
      ScaleHeight     =   330
      ScaleWidth      =   1770
      TabIndex        =   67
      Top             =   1080
      Width           =   1770
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   0
         Width           =   1330
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
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
         Index           =   0
         Left            =   0
         TabIndex        =   69
         Top             =   45
         Width           =   240
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   240
         X2              =   1680
         Y1              =   285
         Y2              =   285
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   13
      Left            =   5400
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   65
      Top             =   4080
      Width           =   2055
      Begin VB.ComboBox cboInfo 
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
         Index           =   13
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   -25
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   24
      Left            =   10080
      Picture         =   "frmApplyOperation.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   57
      TabStop         =   0   'False
      ToolTipText     =   "编辑(F4)"
      Top             =   10005
      Width           =   285
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
      Height          =   255
      Index           =   21
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   9570
      Width           =   1455
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
      Height          =   255
      Index           =   23
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   9570
      Width           =   1455
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
      Height          =   255
      Index           =   22
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   9570
      Width           =   1455
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
      Height          =   255
      Index           =   11
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3600
      Width           =   855
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
      Height          =   255
      Index           =   20
      Left            =   8445
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1380
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   20
      Left            =   9945
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   7440
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
      Height          =   255
      Index           =   19
      Left            =   4845
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   7440
      Width           =   1500
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   19
      Left            =   6405
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   7440
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
      Height          =   255
      Index           =   18
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   7440
      Width           =   1260
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   18
      Left            =   3165
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   7440
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
      Height          =   255
      Index           =   16
      Left            =   1695
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   6885
      Width           =   1260
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   16
      Left            =   3165
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   6885
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
      Height          =   255
      Index           =   14
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   14
      Left            =   3165
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   6360
      Width           =   270
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   12
      Left            =   1530
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   39
      Top             =   4080
      Width           =   2055
      Begin VB.ComboBox cboInfo 
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
         Index           =   12
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   -25
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   10
      Left            =   7845
      TabIndex        =   36
      ToolTipText     =   "选择(*)"
      Top             =   3600
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
      Height          =   255
      Index           =   10
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3600
      Width           =   6060
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   8
      Left            =   3285
      Picture         =   "frmApplyOperation.frx":6948
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "编辑(F4)"
      Top             =   2640
      Width           =   285
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
      Height          =   255
      Index           =   8
      Left            =   1620
      TabIndex        =   18
      Text            =   "2013-06-20 18:00"
      Top             =   2655
      Width           =   1935
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
      Height          =   255
      Index           =   4
      Left            =   8745
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1545
      Width           =   1330
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
      Height          =   255
      Index           =   6
      Left            =   4365
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2010
      Width           =   1215
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
      Height          =   255
      Index           =   1
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1330
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
      Height          =   255
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1330
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
      Height          =   255
      Index           =   7
      Left            =   6885
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2010
      Width           =   1215
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
      Height          =   255
      Index           =   5
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1335
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
      Height          =   255
      Index           =   3
      Left            =   6450
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1665
      Width           =   1330
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
      Height          =   255
      Index           =   9
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3090
      Width           =   8175
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Index           =   9
      Left            =   9960
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "选择(*)"
      Top             =   3090
      Width           =   270
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   10800
      TabIndex        =   17
      Top             =   5280
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
   Begin VSFlex8Ctl.VSFlexGrid vsOper 
      Height          =   1605
      Left            =   1605
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4560
      Width           =   8760
      _cx             =   15452
      _cy             =   2831
      Appearance      =   0
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
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmApplyOperation.frx":6A3E
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
      Editable        =   0
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
      Height          =   255
      Index           =   24
      Left            =   8400
      TabIndex        =   34
      Text            =   "2013-06-20 18:00"
      Top             =   10020
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid vsOther 
      Height          =   1545
      Left            =   1560
      TabIndex        =   33
      Top             =   7845
      Width           =   8775
      _cx             =   15478
      _cy             =   2725
      Appearance      =   0
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
      SheetBorder     =   -2147483632
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmApplyOperation.frx":6A68
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
      Editable        =   0
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
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   15
      Left            =   5280
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   43
      Top             =   6360
      Width           =   2175
      Begin VB.ComboBox cboInfo 
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
         Index           =   15
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   -25
         Width           =   1995
      End
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   5460
      X2              =   6765
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "再次手术类型"
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
      Index           =   27
      Left            =   4140
      TabIndex        =   71
      Top             =   2595
      Width           =   1260
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   225
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "其它内容"
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
      Index           =   26
      Left            =   645
      TabIndex        =   66
      Top             =   7905
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   8400
      X2              =   10080
      Y1              =   10290
      Y2              =   10290
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   1560
      X2              =   3240
      Y1              =   9840
      Y2              =   9840
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
      Index           =   21
      Left            =   645
      TabIndex        =   64
      Top             =   9600
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "生效时间"
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
      Index           =   24
      Left            =   7485
      TabIndex        =   63
      Top             =   10020
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   8400
      X2              =   10080
      Y1              =   9840
      Y2              =   9840
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
      Index           =   23
      Left            =   7485
      TabIndex        =   62
      Top             =   9600
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   4920
      X2              =   6600
      Y1              =   9840
      Y2              =   9840
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
      Index           =   22
      Left            =   3960
      TabIndex        =   61
      Top             =   9600
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "手术级别"
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
      Index           =   11
      Left            =   8325
      TabIndex        =   56
      Top             =   3615
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "手术科室"
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
      Index           =   12
      Left            =   645
      TabIndex        =   55
      Top             =   4080
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   9285
      X2              =   10320
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "第三助手"
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
      Index           =   20
      Left            =   7455
      TabIndex        =   53
      Top             =   7440
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   8325
      X2              =   10200
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "第二助手"
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
      Index           =   19
      Left            =   3960
      TabIndex        =   51
      Top             =   7440
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   4815
      X2              =   6720
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "第一助手"
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
      Index           =   18
      Left            =   645
      TabIndex        =   49
      Top             =   7440
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   1605
      X2              =   3480
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "主刀医生科室"
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
      Index           =   17
      Left            =   3960
      TabIndex        =   47
      Top             =   6900
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   5280
      X2              =   7200
      Y1              =   7155
      Y2              =   7155
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "主刀医生"
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
      Index           =   16
      Left            =   645
      TabIndex        =   46
      Top             =   6900
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   1605
      X2              =   3480
      Y1              =   7155
      Y2              =   7155
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "麻醉执行科室"
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
      Left            =   3960
      TabIndex        =   44
      Top             =   6360
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   5280
      X2              =   7200
      Y1              =   6630
      Y2              =   6630
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "麻醉方法"
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
      Index           =   14
      Left            =   645
      TabIndex        =   42
      Top             =   6360
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   1605
      X2              =   3480
      Y1              =   6630
      Y2              =   6630
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "手术名称"
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
      Index           =   10
      Left            =   585
      TabIndex        =   40
      Top             =   3600
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   1605
      X2              =   3525
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "手术时间"
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
      Index           =   8
      Left            =   585
      TabIndex        =   38
      Top             =   2655
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "术前诊断"
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
      Index           =   9
      Left            =   645
      TabIndex        =   37
      Top             =   3135
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   1605
      X2              =   8160
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   5280
      X2              =   7200
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "手术情况"
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
      Index           =   13
      Left            =   4320
      TabIndex        =   35
      Top             =   4095
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   1605
      X2              =   3285
      Y1              =   2925
      Y2              =   2925
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   8700
      X2              =   9660
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   6765
      X2              =   8085
      Y1              =   1920
      Y2              =   1920
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
      Index           =   4
      Left            =   7816
      TabIndex        =   16
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "附加手术"
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
      Index           =   25
      Left            =   645
      TabIndex        =   15
      Top             =   4560
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
      Index           =   2
      Left            =   2925
      TabIndex        =   14
      Top             =   1680
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
      Index           =   3
      Left            =   5371
      TabIndex        =   13
      Top             =   1680
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
      Index           =   6
      Left            =   3285
      TabIndex        =   12
      Top             =   2040
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
      Index           =   1
      Left            =   480
      TabIndex        =   11
      Top             =   1640
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
      Index           =   7
      Left            =   5790
      TabIndex        =   10
      Top             =   2040
      Width           =   840
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
      Left            =   645
      TabIndex        =   9
      Top             =   2070
      Width           =   840
   End
   Begin VB.Label lblHead 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "手术申请单"
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
      Left            =   4605
      TabIndex        =   8
      Top             =   750
      Width           =   1875
   End
   Begin VB.Line linHead 
      X1              =   4590
      X2              =   6510
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   4245
      X2              =   5565
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1380
      X2              =   2640
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   4245
      X2              =   5565
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   6765
      X2              =   8085
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1605
      X2              =   3045
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   1605
      X2              =   10200
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "frmApplyOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCtlID
    e_No = 0
    e_姓名 = 1
    e_性别 = 2
    e_年龄 = 3
    e_病情 = 4
    e_床号 = 5
    e_住院号 = 6
    e_科室 = 7
    e_手术时间 = 8
    e_术前诊断 = 9
    e_手术名称 = 10
    e_手术级别 = 11
    e_执行科室 = 12
    e_手术情况 = 13
    e_麻醉方法 = 14
    e_麻醉执行科室 = 15
    e_主刀医生 = 16
    e_主刀医生科室 = 17
    e_第一助手 = 18
    e_第二助手 = 19
    e_第三助手 = 20
    e_申请科室 = 21
    e_申请医师 = 22
    e_主治医师 = 23
    e_生效时间 = 24
    e_附加手术 = 25
    e_其它内容 = 26
    e_再次手术类型 = 27
End Enum

Private Enum mTableCol
    '附加手术表格
    COL_名称 = 0
    COL_手术级别 = 1
    COL_计价性质 = 2
    COL_执行性质 = 3
    COL_摘要 = 4
    COL_必要时 = 5
    
    '扩展项目
    COL_项目名称 = 0
    col_内容 = 1
End Enum

Private mobjVBA As Object
Private mobjScript As clsScript
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mobjReport As Object
Private mblnReturn As Boolean '是否了回车确认

Private mintType As Integer '=0新增，=1修改，=2查看
Private mint场合 As Integer '0 住院医生工作站，1 门诊医生工作站；
Private mint调用场合 As Integer '申请单调用场合，0－医生站调用，1－医嘱编辑界面调用。为1时允许用缓存数据加载界面和保存为缓存数据。
Private mint调用类型 As Integer  '1-门诊,2-住院
Private mint服务对象 As Integer '1-门诊,2-住院

Private mbln提醒对码 As Boolean
Private mint险类 As Integer '当前病人险类
Private mrsAppend As ADODB.Recordset
Private mbln外院医生建档 As Boolean '外院医生必须建档
Private mstr主刀等级 As String   '主刀医生的手术等级
Private mobjEmrInterface As Object           '新版病历申请附项读取部件

Private mlng病人ID As Long
Private mstr挂号单 As String
Private mlng主页ID As Long
Private mlng挂号ID As Long
Private mlng病区ID As Long
Private mlng病人科室id As Long '病人科室id/挂号执行科室id
Private mlng病人性质 As Long   '0-住院，1-门诊
Private mlng开单科室ID As Long
Private mstr开单科室 As String '申请科室名称

Private mlng手术项目ID As Long

Private mlng手术执行科室性质 As Long
Private mlng手术执行科室ID As Long

Private mlng麻醉项目ID As Long
Private mlng麻醉执行科室性质 As Long
Private mlng麻醉执行科室id As Long

Private mstr附手术IDs As String
Private mintPState As Integer
Private mdatTurn As Date
Private mstr入院时间 As String
Private mstr上次转科时间 As String
Private mstrDefine As String
Private mlngUpdateAdvice As Long  '传入的医嘱ID，组医嘱ID
Private mstr诊断IDs As String  '诊断关联
Private mbln补录 As Boolean
Private mclsDiagEdit As zlMedRecPage.clsDiagEdit
Private mclsMipModule As zl9ComLib.clsMipModule '消息平台对象
Private mrsCard As ADODB.Recordset
Private mstr摘要 As String '主手术医嘱摘要 由 gclsInsure.GetItemInfo 获取；lblInfo(e_麻醉执行科室).Tag 麻醉方法项目的摘要
Private mstr费别 As String
Private mlng前提ID As Long
Private mbytBaby As Byte  '婴儿序号

Public Function ShowMe(frmParent As Object, ByVal int场合 As Integer, ByVal intType As Integer, ByVal lng病人ID As Long, ByVal str就诊ID As String, ByVal lng病人性质 As Long, _
    Optional ByRef lng医嘱ID As Long, Optional ByVal lng科室id As Long, Optional ByVal lng开单科室ID As Long, Optional ByVal strDefine As String, _
    Optional ByVal lng病区ID As Long, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, Optional ByVal int调用场合 As Integer, _
    Optional ByRef objMip As Object, Optional ByVal lng项目id As Long, Optional ByRef rsCard As ADODB.Recordset, Optional ByVal lng前提ID As Long, Optional ByVal bytBaby As Byte) As Boolean
'功能：公共接口
'参数：frmParent 父对象窗体；int场合 0 住院医生工作站，1 门诊医生工作站； lng病人性质 0-住院，1-门诊；lng病人ID；
'      str就诊ID 跟据 int场合 判断，主页id/挂号单；
'      intType 操作类型   0-新增，1-修改，2-查看,3-医嘱编辑调用；lng医嘱ID 传入的医嘱ID；
'      lng科室ID  病人科室id/挂号执行科室id；int调用场合 0－工作站界面，1－ 医嘱编辑界面； lng开单科室ID 开嘱科室id；
'      strDefine 医嘱内容格式串；
'      lng病区ID、intPState、datTurn 住院才有；objMip 消息对象用于消息发送 住院才有；
'      lng项目ID-医嘱编辑界面输入手术项目传入的主手术项目ID
'      rsCard，医嘱卡片和医嘱表格的信息。只有在医嘱编辑界面点医嘱内容右边的下拉按钮时才会传入。
    
    mint场合 = int场合
    
    If mint场合 = 0 Then
        mlng主页ID = Val(str就诊ID)
        mint调用类型 = 2
        mint服务对象 = 2
    Else
        mstr挂号单 = str就诊ID
        mint调用类型 = 1
        mint服务对象 = 1
    End If
    mint调用场合 = int调用场合
    mlng病人ID = lng病人ID
    mlng病人性质 = lng病人性质
    mlng病人科室id = lng科室id
    mlng病区ID = lng病区ID
    mlng开单科室ID = lng开单科室ID
    mlng前提ID = lng前提ID
    mbytBaby = bytBaby
    mstrDefine = strDefine
    mlngUpdateAdvice = lng医嘱ID
    mintPState = intPState
    mintType = intType
    mdatTurn = datTurn
    mlng手术项目ID = lng项目id
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Set mrsCard = rsCard
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    If mblnOK Then lng医嘱ID = mlngUpdateAdvice
    
    ShowMe = mblnOK
    Set rsCard = mrsCard
End Function

Private Sub cboInfo_Click(Index As Integer)
    Dim blnCancel As Boolean, intIdx As Integer
    Dim strSQL As String, rsTmp As Recordset
    Dim vRect As RECT
    
    If Index = e_执行科室 Or Index = e_麻醉执行科室 Then
        With cboInfo(Index)
            If .ItemData(.ListIndex) = -1 Then
                '他科执行，弹出选择执行科室
                strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                    " From 部门表 A,部门性质说明 B" & _
                    " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
                    IIF(gstrNodeNo <> "", " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " Order by A.编码"
                vRect = zlControl.GetControlRect(cboInfo(Index).hwnd)
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "执行科室", , , , , , True, vRect.Left, vRect.Top, .Height, blnCancel, , True)
                If Not rsTmp Is Nothing Then
                    intIdx = Cbo.FindIndex(cboInfo(Index), rsTmp!ID)
                    If intIdx <> -1 Then
                        .ListIndex = intIdx
                    Else
                        .AddItem rsTmp!编码 & "-" & rsTmp!名称, .ListCount - 1
                        .ItemData(.NewIndex) = rsTmp!ID
                        .ListIndex = .NewIndex
                    End If
                    If .ListIndex >= 0 Then
                        .Tag = .ItemData(.ListIndex)
                    End If
                Else
                    If Not blnCancel Then
                        MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
                    End If
                    '恢复成现有的科室(不引发Click)
                    If .Tag <> "" Then
                        intIdx = Cbo.FindIndex(cboInfo(Index), Val(.Tag))
                        Call Cbo.SetIndex(cboInfo(Index).hwnd, intIdx)
                    End If
                End If
            Else
                If Index = e_执行科室 Then
                    mlng手术执行科室ID = cboInfo(Index).ItemData(cboInfo(Index).ListIndex)
                ElseIf Index = e_麻醉执行科室 Then
                    mlng麻醉执行科室id = cboInfo(Index).ItemData(cboInfo(Index).ListIndex)
                End If
            End If
        End With
    End If
    
    If Visible Then mblnChange = True
End Sub

Private Sub cboInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SeekNextCtl
End Sub

Private Sub cmdDate_Click(Index As Integer)
'功能：选择日期
    Dim lngIndex As Long
    
    If Index = e_手术时间 Then
        lngIndex = e_手术时间
        dtpDate.Left = txtInfo(lngIndex).Left
        dtpDate.Top = cmdDate(Index).Top + cmdDate(Index).Height
    ElseIf Index = e_生效时间 Then
        lngIndex = e_生效时间
        dtpDate.Left = cmdDate(Index).Left + cmdDate(Index).Width - dtpDate.Width
        dtpDate.Top = txtInfo(lngIndex).Top - dtpDate.Height
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
    If intIndex = e_手术时间 Then
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
    ElseIf intIndex = e_生效时间 Then
        '取值
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '判断时间合法性
        If Not Check安排时间(strDate, txtInfo(e_生效时间).Text) Then
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

Private Sub cmdInfo_Click(Index As Integer)

    Select Case Index
    Case e_术前诊断
        Call GetBeforeOperDiag
    Case e_手术名称
        Call GetItemOper(1)
    Case e_麻醉方法
        Call GetItem麻醉(1)
    Case e_主刀医生
        Call GetItem主刀医生(1)
    Case e_主刀医生科室
        Call GetItem主刀医生科室(1)
    Case e_第一助手
        Call GetItemDoctor(1, e_第一助手)
    Case e_第二助手
        Call GetItemDoctor(1, e_第二助手)
    Case e_第三助手
        Call GetItemDoctor(1, e_第三助手)
    End Select
End Sub

Private Sub GetBeforeOperDiag()
'功能：获取术前诊断
    Dim str诊断 As String
    Dim lng就诊ID As Long
    
    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mlng病人性质 = 1, 1260, 1261), mclsMipModule)
    End If
    lng就诊ID = IIF(mint场合 = 0, mlng主页ID, mlng挂号ID)
    If mclsDiagEdit.ShowDiagEdit(Me, mlngUpdateAdvice, mlng病人ID, lng就诊ID, IIF(mlng病人性质 = 1, 1, 2), mlng病人科室id, mstr诊断IDs, str诊断, 0, mlngUpdateAdvice) Then
        txtInfo(e_术前诊断).Text = str诊断
        Call SeekNextCtl
    End If
End Sub

Private Sub Load缺省手术()
'功能：加载缺省手术项目，主手项目从外面传入时
    Dim strSQL As String, rsTmp As Recordset
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模,A.执行科室 as 执行科室性质ID,a.计价性质 as 计价性质ID" & _
        " From 诊疗项目目录 A where a.id=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng手术项目ID)
    If Not rsTmp.EOF Then
        mlng手术项目ID = Val(rsTmp!ID & "")
        Call Init申请附项
        mlng手术执行科室性质 = Val(rsTmp!执行科室性质ID & "")
        lblInfo(e_手术名称).Tag = Val(rsTmp!计价性质ID & "")
        txtInfo(e_手术级别).Text = GetMax手术等级(mlng手术项目ID)
        txtInfo(e_手术名称).Text = rsTmp!名称 & ""
        txtInfo(e_手术名称).Tag = txtInfo(e_手术名称).Text
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItemOper(ByVal intType As Integer)
'功能：选择手术项目－主手术名称 控件设置
'参数：intType =0 KeyPress调用，=1 下拉按钮调用
    Dim strSQL As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim lngTmp As Long
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模,A.执行科室 as 执行科室性质ID,a.计价性质 as 计价性质ID" & _
        " From 诊疗项目目录 A,诊疗项目别名 B" & _
        " Where A.ID=B.诊疗项目ID And A.类别='F' And A.服务对象 IN(" & IIF(mint服务对象 = 1, 1, 2) & ",3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        IIF(intType = 0, " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2])", "") & _
        " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[4])" & _
        " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
        Decode(gbytCode, 0, " And B.码类 IN([3],3)", 1, " And B.码类 IN([3],3)", "") & _
        " Order by A.编码"
    vRect = zlControl.GetControlRect(txtInfo(e_手术名称).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Me.Caption, False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(e_手术名称).Height, blnCancel, False, True, UCase(txtInfo(e_手术名称).Text) & "%", _
        gstrLike & UCase(txtInfo(e_手术名称).Text) & "%", gbytCode + 1, mlng病人科室id)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "未找到匹配的项目。", vbInformation, gstrSysName
        End If
        Call zlControl.TxtSelAll(txtInfo(e_手术名称))
        txtInfo(e_手术名称).SetFocus: Exit Sub
    Else
        blnDo = True
        '如果参数启用了，检查手术的开单权
        If gbln手术授权管理 Then
            If CheckUserEmpower(Val(rsTmp!ID & "")) = False Then
                blnDo = False
                MsgBox "当前操作员不具备手术""" & rsTmp!名称 & """的开单权，不允许下达。", vbInformation, Me.Caption
                Call zlControl.TxtSelAll(txtInfo(e_手术名称))
                txtInfo(e_手术名称).SetFocus: Exit Sub
            End If
        End If
    End If
    
    If blnDo Then
        mlng手术项目ID = Val(rsTmp!ID & "")
        Call Init申请附项
        mlng手术执行科室性质 = Val(rsTmp!执行科室性质ID & "")
        Call Set手术执行科室(mlng手术执行科室性质)
        lblInfo(e_手术名称).Tag = Val(rsTmp!计价性质ID & "")
        txtInfo(e_手术级别).Text = GetMax手术等级(mlng手术项目ID)
        txtInfo(e_手术名称).Text = rsTmp!名称 & ""
        txtInfo(e_手术名称).Tag = txtInfo(e_手术名称).Text
        txtInfo(e_手术名称).SetFocus
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

Private Sub GetItem麻醉(ByVal intType As Integer)
'功能：选择麻醉项目－麻醉方法 控件设置
'参数：intType =0 KeyPress调用，=1 下拉按钮调用
    Dim strSQL As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim lngTmp As Long
    Dim strSQLItem As String, strLike As String
    
    On Error GoTo errH
    
    cboInfo(e_麻醉执行科室).TabStop = True
    
    If intType = 1 Then
        '输入麻醉项目
        strSQLItem = " From 诊疗项目目录 A Where A.类别='G' And A.服务对象 IN([2],3) And A.ID<>[1]" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[3])" & _
                " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                
        strSQL = "Select 0 as 末级,Max(Level) as 级ID,ID,上级ID,编码,名称,NULL as 单位,NULL as 规模,NULL as 执行科室性质ID,null as 计价性质ID" & _
            " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select 分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID" & _
            " Group by ID,上级ID,编码,名称"
            
        strSQL = strSQL & " Union ALL" & _
            " Select 1 as 末级,1 as 级ID,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模,A.执行科室 as 执行科室性质ID,a.计价性质 as 计价性质ID" & _
            strSQLItem & " Order By 末级,级ID Desc,编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "麻醉项目", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, mlng手术项目ID, mint服务对象, mlng病人科室id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到匹配项目！", vbInformation, gstrSysName
            End If
            cboInfo(e_麻醉执行科室).TabStop = False
            txtInfo(e_麻醉方法).SetFocus: Exit Sub
        End If
    Else
        If txtInfo(e_麻醉方法).Tag = txtInfo(e_麻醉方法).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(e_麻醉方法).Text = "" Then '相当于是清除该项目
            mlng麻醉项目ID = 0
            Call cboInfo(e_麻醉执行科室).Clear
            cboInfo(e_麻醉执行科室).TabStop = False
            txtInfo(e_麻醉方法).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
        
        '优化
        strLike = gstrLike
        If Len(txtInfo(e_麻醉方法).Text) < 2 Then strLike = ""
    
        '输入麻醉项目
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模,A.执行科室 as 执行科室性质ID,a.计价性质 as 计价性质ID" & _
            " From 诊疗项目目录 A,诊疗项目别名 B" & _
            " Where A.ID=B.诊疗项目ID And A.类别='G' And A.服务对象 IN([3],3)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]) And B.码类=[4]" & _
            " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[5])" & _
            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
            " Order by A.编码"
        vRect = zlControl.GetControlRect(txtInfo(e_麻醉方法).hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "麻醉项目", False, "", "", False, False, True, vRect.Left, vRect.Top, txtInfo(e_麻醉方法).Height, blnCancel, False, True, _
            UCase(txtInfo(e_麻醉方法).Text) & "%", strLike & UCase(txtInfo(e_麻醉方法).Text) & "%", mint服务对象, gbytCode + 1, mlng病人科室id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到匹配项目！", vbInformation, gstrSysName
            End If
            txtInfo(e_麻醉方法).Text = cmdInfo(e_麻醉方法).Tag
            zlControl.TxtSelAll txtInfo(e_麻醉方法)
            txtInfo(e_麻醉方法).SetFocus
            Exit Sub
        End If
    End If
    
    mlng麻醉项目ID = Val(rsTmp!ID & "")
    mlng麻醉执行科室性质 = Val(rsTmp!执行科室性质ID & "")
    Call Set麻醉执行科室(mlng麻醉执行科室性质)
    
    lblInfo(e_麻醉方法).Tag = Val(rsTmp!计价性质ID & "")
    txtInfo(e_麻醉方法).Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
    txtInfo(e_麻醉方法).Tag = txtInfo(e_麻醉方法).Text
    txtInfo(e_麻醉方法).SetFocus
    Call SeekNextCtl
    If Visible Then mblnChange = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItem主刀医生(ByVal intType As Integer)
'功能：获取主刀医生项目
'参数：0 文本框按回车，1 点按钮
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean, strTmp As String
    Dim blnDo As Boolean, str部门 As String
    Dim lng部门ID As Long, lng人员id As Long
    Dim i As Integer
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(e_主刀医生).Tag = txtInfo(e_主刀医生).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(e_主刀医生).Text = "" Then '相当于是清除该项目
            txtInfo(e_主刀医生).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
            
    strInput = Trim(UCase(txtInfo(e_主刀医生).Text))   '传入的值存在前缀空格
    
    strSQL = "Select A.ID,A.编号,A.姓名,A.简码,A.手术等级" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生'" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        IIF(intType = 0, " And (A.编号 Like [1] Or A.姓名 Like [2] Or A.简码 Like [2])", "") & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.编号"
    vRect = zlControl.GetControlRect(txtInfo(e_主刀医生).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医生", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(e_主刀医生).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
        
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            If mbln外院医生建档 Then
                Call MsgBox("没有找到匹配的医生!", vbInformation, gstrSysName)
                blnDo = False
            Else
                If MsgBox("没有找到匹配的医生，你确定要输入没有建立人员档案的医生吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    blnDo = True
                    strTmp = strInput
                Else
                    blnDo = False
                End If
            End If
        Else
            Call MsgBox("没有找到匹配的医生!", vbInformation, gstrSysName)
            blnDo = False
        End If
    Else
        blnDo = True
        txtInfo(e_主刀医生).Text = rsTmp!姓名 & ""
        txtInfo(e_主刀医生).Tag = rsTmp!姓名 & ""
        lng人员id = rsTmp!ID
        If gbln手术分级管理 Then mstr主刀等级 = rsTmp!手术等级 & ""
        txtInfo(e_主刀医生).SetFocus
    End If
    
    If blnDo Then
        strSQL = "Select b.名称,a.缺省 From 部门人员 A, 部门表 B, 部门性质说明 C" & _
            " Where a.部门id = b.Id And b.Id = c.部门id And c.工作性质 = '临床' And a.人员id = [1]" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng人员id)
        If Not rsTmp.EOF Then
            txtInfo(e_主刀医生科室).Text = rsTmp!名称 & ""
            rsTmp.Filter = "缺省=1"
            If Not rsTmp.EOF Then txtInfo(e_主刀医生科室).Text = rsTmp!名称 & ""
            
            txtInfo(e_主刀医生科室).Tag = txtInfo(e_主刀医生科室).Text
            txtInfo(e_主刀医生科室).SetFocus
        End If
        Call SeekNextCtl
    Else
        txtInfo(e_主刀医生).SetFocus
        Call zlControl.TxtSelAll(txtInfo(e_主刀医生))
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItem主刀医生科室(ByVal intType As Integer)
'功能：获取主刀医生科室
'参数：0 文本框按回车，1 点按钮
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    Dim i As Integer, strDoctor As String
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(e_主刀医生科室).Tag = txtInfo(e_主刀医生科室).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(e_主刀医生科室).Text = "" Then  '相当于是清除该项目
            txtInfo(e_主刀医生科室).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
       
    strDoctor = txtInfo(e_主刀医生).Text
    strInput = Trim(UCase(txtInfo(e_主刀医生科室).Text))    '传入的值存在前缀空格
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称 as 科室,A.简码 From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)  And a.Id = b.部门id" & _
        IIF(intType = 0, " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])", "") & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) And B.工作性质='临床'" & _
        IIF(strDoctor <> "", " And a.id in (select x.部门id from 部门人员 X, 人员表 Y where x.人员id=y.id and y.姓名=[3])", "") & _
        " Order by A.编码"
        
    vRect = zlControl.GetControlRect(txtInfo(e_主刀医生科室).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "科室", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(e_主刀医生科室).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", strDoctor)
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("没有找到匹配的科室!", vbInformation, gstrSysName)
            txtInfo(e_主刀医生科室).SetFocus
            zlControl.TxtSelAll txtInfo(e_主刀医生科室)
            Exit Sub
        End If
    Else
        txtInfo(e_主刀医生科室).Text = rsTmp!科室 & ""
        txtInfo(e_主刀医生科室).Tag = rsTmp!科室 & ""
        txtInfo(e_主刀医生科室).SetFocus
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItemDoctor(ByVal intType As Integer, ByVal intIndex As Integer)
'功能：提取医生
'参数：intType 0 输入匹配，1 点击按钮，intIndex 控件索引
    Dim strInput As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    
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

Private Sub Set手术执行科室(ByVal lng执行科室 As Long, Optional ByVal lng执行科室ID As Long)
'功能：设置执行科室
'参数：lng执行科室-执行性质，lng执行科室ID=如果传入，则表示设置此执行科室为当前执行科室
    Dim lngTmp As Long
    
    With cboInfo(e_执行科室)
        .Enabled = True
        If lng执行科室 = 5 Then
            .Clear:
            .AddItem "-"
            .ListIndex = 0
        Else
            If .ListIndex >= 0 And lng执行科室ID = 0 Then
                lngTmp = .ItemData(.ListIndex)
            ElseIf lng执行科室ID <> 0 Then
                lngTmp = lng执行科室ID
            End If
            
            If lngTmp = 0 Then lngTmp = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "F", mlng手术项目ID, 0, lng执行科室, mlng病人科室id, mlng开单科室ID, 1, IIF(mlng病人性质 = 1, 1, 2))
            
            Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cboInfo(e_执行科室), "F", mlng手术项目ID, 0, _
                lng执行科室, mlng病人科室id, mlng开单科室ID, lngTmp, 1, IIF(mlng病人性质 = 1, 1, 2))

            If lng执行科室ID = 0 Then
                If .ListIndex = -1 And .ListCount = 1 Then
                    .ListIndex = 0
                Else
                     '如果有多项，则取默认的执行科室
                    lng执行科室ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "F", mlng手术项目ID, 0, _
                            lng执行科室, mlng病人科室id, mlng开单科室ID, 1, IIF(mlng病人性质 = 1, 1, 2))
                End If
            End If

            If lng执行科室ID <> 0 Then Call Cbo.Locate(cboInfo(e_执行科室), lng执行科室ID, True)
         
        End If
        
        If .ListCount = 1 Then .Enabled = False
        
        If .ListIndex >= 0 Then .Tag = .ItemData(.ListIndex)
     
    End With
End Sub

Private Sub Set麻醉执行科室(ByVal lng执行科室 As Long, Optional ByVal lng执行科室ID As Long)
'功能：设置麻醉执行科室
'参数：lng执行科室-执行性质，lng执行科室ID=如果传入，则表示设置此执行科室为当前执行科室
    Dim lngTmp As Long
    
    With cboInfo(e_麻醉执行科室)
        .Enabled = True
        If lng执行科室 = 5 Then
            .Clear: .AddItem "-"
            .ListIndex = 0
        Else
            If .ListIndex >= 0 And lng执行科室ID = 0 Then
                lngTmp = .ItemData(.ListIndex)
            ElseIf lng执行科室ID <> 0 Then
                lngTmp = lng执行科室ID
            End If
            
            If lngTmp = 0 Then lngTmp = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "G", mlng手术项目ID, 0, lng执行科室, mlng病人科室id, mlng开单科室ID, 1, IIF(mlng病人性质 = 1, 1, 2))
            
            Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cboInfo(e_麻醉执行科室), "G", mlng麻醉项目ID, 0, _
                lng执行科室, mlng病人科室id, mlng开单科室ID, lngTmp, 1, IIF(mlng病人性质 = 1, 1, 2))
                
            If lng执行科室ID = 0 Then
                If .ListIndex = -1 And .ListCount = 1 Then
                    .ListIndex = 0
                Else
                     '如果有多项，则取默认的执行科室
                    lng执行科室ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "G", mlng麻醉项目ID, 0, _
                            lng执行科室, mlng病人科室id, mlng开单科室ID, 1, IIF(mlng病人性质 = 1, 1, 2))
                End If
            End If
            
            If lng执行科室ID <> 0 Then Call Cbo.Locate(cboInfo(e_麻醉执行科室), lng执行科室ID, True)
          
        End If
        
        If .ListCount = 1 Then cboInfo(e_麻醉执行科室).Enabled = False
        
        If .ListIndex >= 0 Then .Tag = .ItemData(.ListIndex)
 
    End With
End Sub

Private Function GetItemAppend(ByVal lng要素ID As Long, ByVal str中文名 As String, ByVal str项目 As String) As String
'功能：获取指定的申请附项值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strText As String
    Dim arrItem As Variant, i As Long
    
    On Error GoTo errH
    
    If mint调用类型 = 1 Then
        '3.未取到或未对应要素的，从病人之前已保存的医嘱中提取,以最后填写的为准
        strSQL = " Select 内容 From (" & _
            " Select B.内容 From 病人医嘱记录 A,病人医嘱附件 B" & _
            " Where A.ID=B.医嘱ID And A.病人ID=[1] And Nvl(A.婴儿,0)=[4]" & _
            IIF(mint场合 = 1, " And A.挂号单=[2]", " And A.主页ID=[3]") & _
            " And B.项目=[5] And B.内容 is Not Null and nvl(a.医嘱状态,0)<>4 Order by A.开嘱时间 Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, mlng主页ID, 0, str项目)
        If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
    End If
    
    '1.如果有对应要素，从要素提取函数读取
    If lng要素ID <> 0 And strText = "" Then
        '先老版，再新版
        If mint场合 = 1 Then
            '老版电子病历
            strSQL = "Select Zl_Replace_Element_Value(B.中文名,[1],A.ID,1) as 内容" & _
                " From 病人挂号记录 A,诊治所见项目 B Where A.NO=[2] And B.ID=[3] And a.记录性质=1 And a.记录状态=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, lng要素ID)
            If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
        Else
            '老版电子病历
            strSQL = "Select Zl_Replace_Element_Value(中文名,[1],[2],2) as 内容" & _
                " From 诊治所见项目 Where ID=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, lng要素ID)
            If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
        End If
        If strText = "" Then
            strText = GetItemAppendByEmr(str中文名)
        End If
    End If
    
    '2.如果诊断，从未保存的已录入诊断中提取
    If str项目 Like "*诊断" And strText = "" And txtInfo(e_术前诊断).Text <> "" Then
        strText = txtInfo(e_术前诊断).Text
    End If
    
    If mint调用类型 = 2 And strText = "" Then
        '3.未取到或未对应要素的，从病人之前已保存的医嘱中提取,以最后填写的为准
        strSQL = " Select 内容 From (" & _
            " Select B.内容 From 病人医嘱记录 A,病人医嘱附件 B" & _
            " Where A.ID=B.医嘱ID And A.病人ID=[1] And Nvl(A.婴儿,0)=[4]" & _
            IIF(mint场合 = 1, " And A.挂号单=[2]", " And A.主页ID=[3]") & _
            " And B.项目=[5] And B.内容 is Not Null and nvl(a.医嘱状态,0)<>4 Order by A.开嘱时间 Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, mlng主页ID, 0, str项目)
        If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
    End If
    
    GetItemAppend = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetItemAppendByEmr(ByVal str中文名 As String) As String
'功能：读取指定病人的指定提纲在病历填写的信息，例如：主诉，诊断等。从病历中获取附项值
    Dim strText As String
    Dim intType As Integer
    Dim lng就诊ID As Long
    
    On Error Resume Next
    
    If mobjEmrInterface Is Nothing Then Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
 
    If Not mobjEmrInterface Is Nothing Then
        If mint场合 = 0 Then
            intType = 2
            lng就诊ID = mlng主页ID
        Else
            intType = 1
            lng就诊ID = mlng挂号ID
        End If
        
        strText = mobjEmrInterface.GetOrderInspectInfoEx(intType, mlng病人ID, lng就诊ID, str中文名)
        If err.Number <> 0 Then
            strText = mobjEmrInterface.GetOrderInspectInfo(mlng病人ID, str中文名)
        End If
        
    End If
    
    err.Clear
    GetItemAppendByEmr = strText
End Function

Private Sub LoadPatiInfo()
'功能：提取病人基本信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
     
    '读取病人相关信息
    If mint场合 = 0 Then
        If 0 = mbytBaby Then
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
            txtInfo(e_住院号).Text = rsTmp!住院号 & ""
            txtInfo(e_姓名).Text = rsTmp!姓名 & ""
            txtInfo(e_性别).Text = rsTmp!性别 & ""
            txtInfo(e_科室).Text = rsTmp!科室 & ""
            txtInfo(e_床号).Text = rsTmp!当前床号 & ""
            txtInfo(e_年龄).Text = rsTmp!年龄 & ""
            txtInfo(e_病情).Text = rsTmp!病情 & ""
            mstr入院时间 = Format(rsTmp!入院日期 & "", "YYYY-MM-DD HH:mm")
            mint险类 = Val(rsTmp!险类 & "")
            mstr费别 = rsTmp!费别 & ""
        End If
        
        mstr上次转科时间 = Get上次转科日期
        
 
        strSQL = "select 信息值 as 主治医师 from 病案主页从表 where 病人id=[1] and 主页id=[2] and 信息名=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, "主治医师")
        If Not rsTmp.EOF Then txtInfo(e_主治医师).Text = rsTmp!主治医师 & ""
        
    Else
        strSQL = "Select a.id, A.姓名,A.性别,A.年龄,a.no,a.门诊号,a.险类,b.名称 as 科室,a.执行人 as 主治医师,decode(a.急诊,1,'是','否') as 病情,c.费别" & _
                " From 病人挂号记录 A,部门表 b,病人信息 c Where a.病人id=c.病人id and A.NO=[1] And a.记录性质=1 And a.记录状态=1 And A.病人ID+0=[2]" & _
                " and a.执行部门id=b.id"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单, mlng病人ID)
        
        If rsTmp.RecordCount > 0 Then
            lblInfo(e_住院号).Caption = "挂 号 单"
            txtInfo(e_住院号).Text = rsTmp!NO & ""
            
            txtInfo(e_姓名).Text = rsTmp!姓名 & ""
            txtInfo(e_性别).Text = rsTmp!性别 & ""
  
            txtInfo(e_科室).Text = rsTmp!科室 & ""
            lblInfo(e_床号).Caption = "门 诊 号"
            txtInfo(e_床号).Text = rsTmp!门诊号 & ""
            txtInfo(e_年龄).Text = rsTmp!年龄 & ""
            
            lblInfo(e_病情).Caption = "急    诊"
            txtInfo(e_病情).Text = rsTmp!病情 & ""
            
            mlng挂号ID = Val(rsTmp!ID & "")
            mint险类 = Val(rsTmp!险类 & "")
            mstr费别 = rsTmp!费别 & ""
            txtInfo(e_主治医师).Text = rsTmp!主治医师 & ""
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitDefault()
'功能：初始化一些表列，缺省值设定
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsTmpOther As ADODB.Recordset
    Dim datCur  As Date, str项目IDs As String, strTmp As String
    Dim blnCanSave As Boolean, str生效时间 As String, str手术时间 As String
    Dim lng执行科室ID As Long, lng附加执行ID As Long
    Dim lng手术项目ID As Long
    Dim int手术情况 As Integer
    Dim arrTmp As Variant
    Dim i As Long
    Dim bln再次 As Boolean
    
    On Error GoTo errH
    datCur = zlDatabase.Currentdate
 
    '公共的缺省值
    txtInfo(e_手术时间).Text = Format(datCur, "YYYY-MM-DD HH:mm")
    txtInfo(e_生效时间).Text = txtInfo(e_手术时间).Text
    txtInfo(e_生效时间).Tag = txtInfo(e_生效时间).Text
    
    If mint场合 = 0 Then
        If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
            If mdatTurn <> CDate(0) Then datCur = mdatTurn - 1 / 24 / 60
            mbln补录 = True
        End If
    End If
    If mlng开单科室ID <> 0 Then
        mstr开单科室 = Sys.RowValue("部门表", mlng开单科室ID, "名称")
        txtInfo(e_申请科室).Text = mstr开单科室
    End If
    txtInfo(e_申请医师).Text = UserInfo.姓名
    
    ' 再次手术类型  可见性
    If mintType = 2 Then
        '查看手术项目时单独判断
        strSQL = "select 1 from 病人医嘱附件 a where a.医嘱id=[1] and a.项目='再次手术类型'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
        bln再次 = Not rsTmp.EOF
    Else
        bln再次 = IsOperateAgain()
    End If
    
    lblInfo(e_再次手术类型).Visible = bln再次
    picInfo(e_再次手术类型).Visible = bln再次
    Line1(e_再次手术类型).Visible = bln再次
        
    If mintType = 0 Then '新增手术项目
        lblInfo(e_其它内容).Visible = False
        vsOther.Visible = False
        picNo.Visible = False
        If mlng手术项目ID <> 0 Then
            Call Load缺省手术
            If lng执行科室ID <> 0 Then
                Call Set手术执行科室(mlng手术执行科室性质, lng执行科室ID)
            Else
                Call Set手术执行科室(mlng手术执行科室性质)
            End If
        End If
        SetItemEditable 1, , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, , 1, 1
    ElseIf mintType = 2 Then '查看手术项目
        picNo.Visible = True
        Call InitFormContent
        SetItemEditable -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1
        mblnChange = False
        Exit Sub
    ElseIf mintType = 1 Then '修改
        picNo.Visible = False
        If mint调用场合 = 0 Then
            Call InitFormContent
        Else
            Call LoadDataFromCache
        End If
        SetItemEditable 1, , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, , 1, 1
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitFormContent()
'功能：根据已有数据加载界面控件
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Integer, lng开单科室ID As Long

    On Error GoTo errH
    '读取诊断
    mstr诊断IDs = GetAdviceDiag(mlngUpdateAdvice, strTmp)
    txtInfo(e_术前诊断).Text = strTmp
    
    '从附项中获取诊断如果附项中有以附项为准
    strSQL = "select 内容 from 病人医嘱附件 where 医嘱ID=[1] and 项目='申请单诊断'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
    If Not rsTmp.EOF Then
        txtInfo(e_术前诊断).Text = rsTmp!内容 & ""
    End If
    
    '手术医嘱信息 1、主手术，2、附加手术，3、麻醉
    strSQL = "select a.id,a.相关id,b.id as 项目id,b.类别,a.申请序号,a.标本部位 as 手术时间,b.编码,b.名称,a.执行科室id as 执行科室id,c.编码 as 科室编码,a.计价特性," & vbNewLine & _
        "c.名称 as 科室名称,b.操作类型 as 规模,a.开嘱科室id, nvl(a.手术情况,0) as 手术情况,a.开始执行时间 as 生效时间,a.执行性质,a.开嘱医生,a.执行标记" & vbNewLine & _
        "from 病人医嘱记录 a,诊疗项目目录 b,部门表 c" & vbNewLine & _
        "where a.诊疗项目id =b.id and a.执行科室id=c.id(+) and (a.相关id=[1] or a.id=[1]) order by a.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
    
    '1、主手术
    rsTmp.Filter = "id=" & mlngUpdateAdvice
    If Not rsTmp.EOF Then
        mlng开单科室ID = Val(rsTmp!开嘱科室id & "")
        mlng手术项目ID = Val(rsTmp!项目ID & "")
        mlng手术执行科室性质 = Val(rsTmp!执行性质 & "")
        mlng手术执行科室ID = Val(rsTmp!执行科室ID & "")
        
        txtInfo(e_No).Text = rsTmp!申请序号 & ""
        txtInfo(e_手术时间).Text = rsTmp!手术时间
        
        txtInfo(e_手术名称).Text = rsTmp!名称
        txtInfo(e_手术名称).Tag = rsTmp!名称
        
        txtInfo(e_手术级别).Text = GetMax手术等级(rsTmp!项目ID & "")
        txtInfo(e_生效时间).Text = Format(rsTmp!生效时间, "YYYY-MM-DD HH:MM")
		lblInfo(e_手术名称).Tag = Val(rsTmp!计价特性 & "")
        With cboInfo(e_执行科室)
            .Clear
            .AddItem IIF(rsTmp!科室编码 & "" <> "", rsTmp!科室编码 & "-" & rsTmp!科室名称, "-")
            .ItemData(.NewIndex) = mlng手术执行科室ID
            .AddItem "[其它...]"
            .ItemData(.NewIndex) = -1
            .ListIndex = 0
            .Tag = mlng手术执行科室ID
        End With
        cboInfo(e_手术情况).ListIndex = Val(rsTmp!手术情况 & "")
        
        txtInfo(e_申请科室).Text = Sys.RowValue("部门表", mlng开单科室ID, "名称")
        txtInfo(e_申请医师).Text = rsTmp!开嘱医生 & ""
    End If
    '2、附加手术
    rsTmp.Filter = "id<>" & mlngUpdateAdvice & " and 类别='F'"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With vsOper
                .Rows = rsTmp.RecordCount + 1
                .RowData(i) = Val(rsTmp!项目ID & "")
                .TextMatrix(i, COL_名称) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                .Cell(flexcpData, i, COL_名称) = .TextMatrix(i, COL_名称)
                .TextMatrix(i, COL_计价性质) = 0
                .TextMatrix(i, COL_执行性质) = Val(rsTmp!执行性质 & "")
                .TextMatrix(i, COL_手术级别) = GetMax手术等级(rsTmp!项目ID & "")
                .TextMatrix(i, COL_必要时) = IIF(Val(rsTmp!执行标记 & "") = 0, "", "-1")
                If mintType = 1 Then .AddItem ""
            End With
            rsTmp.MoveNext
        Next
    Else
        If mintType = 1 Then vsOper.Rows = 2
    End If
    '3、麻醉
    rsTmp.Filter = "类别='G'"
    If Not rsTmp.EOF Then
        mlng麻醉项目ID = Val(rsTmp!项目ID & "")
        mlng麻醉执行科室性质 = Val(rsTmp!执行性质 & "")
        mlng麻醉执行科室id = Val(rsTmp!执行科室ID & "")
        
        txtInfo(e_麻醉方法).Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
        txtInfo(e_麻醉方法).Tag = txtInfo(e_麻醉方法).Text
        
        With cboInfo(e_麻醉执行科室)
            .Clear
            .AddItem IIF(rsTmp!科室编码 & "" <> "", rsTmp!科室编码 & "-" & rsTmp!科室名称, "-")
            .ItemData(.NewIndex) = mlng麻醉执行科室id
            .AddItem "[其它...]"
            .ItemData(.NewIndex) = -1
            .ListIndex = 0
            .Tag = mlng麻醉执行科室id
        End With
    End If

    Call Init申请附项(False)
    
    '加载申请附项值
    strSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱id =[1] Order By 排列"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)

    If rsTmp.EOF Then Exit Sub
    txtInfo(e_主刀医生).Text = rsTmp!内容 & ""
    txtInfo(e_主刀医生).Tag = txtInfo(e_主刀医生).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    txtInfo(e_主刀医生科室).Text = rsTmp!内容 & ""
    txtInfo(e_主刀医生科室).Tag = txtInfo(e_主刀医生科室).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    txtInfo(e_第一助手).Text = rsTmp!内容 & ""
    txtInfo(e_第一助手).Tag = txtInfo(e_第一助手).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    txtInfo(e_第二助手).Text = rsTmp!内容 & ""
    txtInfo(e_第二助手).Tag = txtInfo(e_第二助手).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    txtInfo(e_第三助手).Text = rsTmp!内容 & ""
    txtInfo(e_第三助手).Tag = txtInfo(e_第三助手).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    If rsTmp!项目 & "" = "再次手术类型" Then
        If rsTmp!内容 & "" = "非计划" Then
            cboInfo(e_再次手术类型).ListIndex = 1
        End If
        rsTmp.MoveNext
    End If
    
    With vsOther
        For i = .FixedRows To .Rows - 1
            If rsTmp.EOF Then Exit Sub
            .TextMatrix(i, col_内容) = rsTmp!内容 & ""
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDataFromCache()
'功能：根据已有数据加载界面控件
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Integer, lng开单科室ID As Long
    Dim varA1 As Variant, varA2 As Variant, arrTmp As Variant
    Dim str项目IDs As String, int手术情况 As Integer
    Dim str生效时间 As String, str手术时间 As String
    Dim lng执行科室ID As Long, lng附加执行ID As Long
    Dim str附项 As String
    Dim blnDo As Boolean
    Dim str附项必要时 As String
    On Error GoTo errH
    
    '加载缓存数据
    If Not mrsCard Is Nothing Then
        If mrsCard.RecordCount > 0 Then
            mrsCard.MoveFirst
            blnDo = True
        End If
    End If
    
    If Not blnDo Then Exit Sub
     
    str项目IDs = mrsCard!附手术项目IDs & ""
    str附项必要时 = mrsCard!附手术必要时 & ""
    mlng手术项目ID = Val(mrsCard!主手术项目ID & "")
    If mrsCard!生效时间 & "" <> "" Then
        txtInfo(e_生效时间).Text = mrsCard!生效时间
        txtInfo(e_生效时间).Tag = txtInfo(e_生效时间).Text
    End If
    If mrsCard!手术时间 & "" <> "" Then
        txtInfo(e_手术时间).Text = mrsCard!手术时间 & ""
    End If
     
    lng执行科室ID = Val(mrsCard!手术执行科室ID & "")
    lng附加执行ID = Val(mrsCard!麻醉执行科室ID & "")
    str附项 = mrsCard!申请附项 & ""
    
    '读取诊断
    mstr诊断IDs = mrsCard!临床诊断IDs & ""
    txtInfo(e_术前诊断).Text = mrsCard!临床诊断描述 & ""
    
    mlng开单科室ID = Val(mrsCard!申请科室id & "")
    mstr开单科室 = Sys.RowValue("部门表", mlng开单科室ID, "名称")
    txtInfo(e_申请科室).Text = mstr开单科室
    mlng手术项目ID = Val(mrsCard!主手术项目ID & "")
    mlng手术执行科室ID = Val(mrsCard!手术执行科室ID & "")
    Set rsTmp = Get诊疗项目记录(mlng手术项目ID)
    mlng手术执行科室性质 = Val(rsTmp!执行科室 & "")
    Call Set手术执行科室(mlng手术执行科室性质, mlng手术执行科室ID)
    txtInfo(e_手术名称).Text = rsTmp!名称
    txtInfo(e_手术名称).Tag = rsTmp!名称
    
    txtInfo(e_手术级别).Text = GetMax手术等级(mlng手术项目ID)
    cboInfo(e_手术情况).ListIndex = Val(mrsCard!手术情况 & "")
 
    '2、附加手术
    vsOper.Rows = 2
    If str项目IDs <> "" Then
        Set rsTmp = Get诊疗项目记录ID(0, str项目IDs)
        With vsOper
            .Rows = rsTmp.RecordCount + 2
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!ID & "")
                .TextMatrix(i, COL_名称) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                .Cell(flexcpData, i, COL_名称) = .TextMatrix(i, COL_名称)
                .TextMatrix(i, COL_计价性质) = 0
                .TextMatrix(i, COL_执行性质) = Val(rsTmp!执行科室 & "")
                .TextMatrix(i, COL_手术级别) = rsTmp!手术级别 & ""
                .TextMatrix(i, COL_必要时) = IIF(InStr(str附项必要时, .RowData(i) & ":1") > 0, -1, 0)
                rsTmp.MoveNext
            Next
        End With
    End If
    '麻醉
    If Val(mrsCard!麻醉项目ID & "") <> 0 Then
        Set rsTmp = Get诊疗项目记录(Val(mrsCard!麻醉项目ID & ""))
        mlng麻醉项目ID = Val(rsTmp!ID & "")
        mlng麻醉执行科室性质 = Val(rsTmp!执行科室 & "")
        mlng麻醉执行科室id = lng附加执行ID
        txtInfo(e_麻醉方法).Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
        txtInfo(e_麻醉方法).Tag = txtInfo(e_麻醉方法).Text
        mlng麻醉执行科室id = Val(mrsCard!麻醉执行科室ID & "")
        Call Set麻醉执行科室(mlng麻醉执行科室性质, mlng麻醉执行科室id)
    End If
   
    Call Init申请附项(False)
    
    '加载申请附项值
    strSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱id =[1] Order By 排列"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 0)
    If str附项 <> "" Then
        Set rsTmp = zlDatabase.CopyNewRec(rsTmp, True)
        varA1 = Split(str附项, "<Split1>")
        For i = 0 To UBound(varA1)
            strTmp = varA1(i)
            If InStr(strTmp, "<Split2>") > 0 Then
                varA2 = Split(strTmp, "<Split2>")
                rsTmp.AddNew Array("项目", "内容"), Array(varA2(0), varA2(3))
            End If
        Next
    End If

'    rsTmp.MoveFirst
    rsTmp.Filter = "项目='主刀医生'"
    If Not rsTmp.EOF Then
        txtInfo(e_主刀医生).Text = rsTmp!内容 & ""
        txtInfo(e_主刀医生).Tag = txtInfo(e_主刀医生).Text
    End If
    
    rsTmp.Filter = "项目='主刀医生科室'"
    If Not rsTmp.EOF Then
        txtInfo(e_主刀医生科室).Text = rsTmp!内容 & ""
        txtInfo(e_主刀医生科室).Tag = txtInfo(e_主刀医生科室).Text
    End If
    
    rsTmp.Filter = "项目='第一助手'"
    If Not rsTmp.EOF Then
        txtInfo(e_第一助手).Text = rsTmp!内容 & ""
        txtInfo(e_第一助手).Tag = txtInfo(e_第一助手).Text
    End If
    
    rsTmp.Filter = "项目='第二助手'"
    If Not rsTmp.EOF Then
        txtInfo(e_第二助手).Text = rsTmp!内容 & ""
        txtInfo(e_第二助手).Tag = txtInfo(e_第二助手).Text
    End If
    
    rsTmp.Filter = "项目='第三助手'"
    If Not rsTmp.EOF Then
        txtInfo(e_第三助手).Text = rsTmp!内容 & ""
        txtInfo(e_第三助手).Tag = txtInfo(e_第三助手).Text
    End If
    
    With vsOther
        For i = .FixedRows To .Rows - 1
            rsTmp.Filter = "项目='" & .TextMatrix(i, COL_项目名称) & "'"
            If Not rsTmp.EOF Then
                .TextMatrix(i, col_内容) = rsTmp!内容 & ""
            End If
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check附项() As Boolean
'功能：查检申请附项中的必填项目
    Dim i As Long
    
    On Error GoTo errH
    
    mrsAppend.Filter = 0
    
    If Not mrsAppend.EOF Then
        Check附项 = False
        
        mrsAppend.Filter = "中文名='主刀医生'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!必填 & "") = 1 And txtInfo(e_主刀医生).Text = "" Then
                MsgBox """主刀医生""项目为必填项目。", vbInformation, Me.Caption
                txtInfo(e_主刀医生).SetFocus
                Exit Function
            End If
        End If
        
        mrsAppend.Filter = "中文名='主刀医生科室'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!必填 & "") = 1 And txtInfo(e_主刀医生科室).Text = "" Then
                MsgBox """主刀医生科室""项目为必填项目。", vbInformation, Me.Caption
                txtInfo(e_主刀医生科室).SetFocus
                Exit Function
            End If
        End If
        
        mrsAppend.Filter = "中文名='第一助手'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!必填 & "") = 1 And txtInfo(e_第一助手).Text = "" Then
                MsgBox """第一助手""项目为必填项目。", vbInformation, Me.Caption
                txtInfo(e_第一助手).SetFocus
                Exit Function
            End If
        End If
        
        mrsAppend.Filter = "中文名='第二助手'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!必填 & "") = 1 And txtInfo(e_第二助手).Text = "" Then
                MsgBox """第二助手""项目为必填项目。", vbInformation, Me.Caption
                txtInfo(e_第二助手).SetFocus
                Exit Function
            End If
        End If
        
        mrsAppend.Filter = "中文名='第三助手'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!必填 & "") = 1 And txtInfo(e_第三助手).Text = "" Then
                MsgBox """第三助手""项目为必填项目。", vbInformation, Me.Caption
                txtInfo(e_第三助手).SetFocus
                Exit Function
            End If
        End If
        
        If vsOther.Visible Then
            With vsOther
                For i = .FixedRows To .Rows - 1
                    If Val(.RowData(i)) <> 0 Then
                        mrsAppend.Filter = "序号=" & Val(.RowData(i))
                        If Not mrsAppend.EOF Then
                            If Val(mrsAppend!必填 & "") = 1 And .TextMatrix(i, col_内容) = "" Then
                                MsgBox """" & mrsAppend!项目 & """项目为必填项目。", vbInformation, Me.Caption
                                .Row = i
                                .Col = col_内容
                                .SetFocus
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End With
        End If
    End If
    
    Check附项 = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Init申请附项(Optional ByVal bln取值 As Boolean = True)
'功能：加载手术附项
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strFilter As String, i As Integer, intIndex As Integer
    Dim bln扩展 As Boolean, strText As String
    
    vsOther.Visible = False
    lblInfo(e_其它内容).Visible = False
    
    strSQL = "Select C.排列 as 序号,C.项目,C.内容,C.要素ID,C.必填,D.中文名" & _
        " From 病历单据应用 A,病历文件列表 B,病历单据附项 C,诊治所见项目 D" & _
        " Where A.诊疗项目ID=[1] And A.应用场合=[2]" & _
        " And A.病历文件ID=B.ID And B.种类=7 And B.ID=C.文件ID And c.要素id=d.id(+)" & _
        " Order by C.排列"
    
    On Error GoTo errH
    
    Set mrsAppend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng手术项目ID, mint服务对象)
    
    If Not mrsAppend.EOF Then
        lblInfo(e_主刀医生科室).Tag = ""
        lblInfo(e_主刀医生).Tag = ""
        lblInfo(e_第一助手).Tag = ""
        lblInfo(e_第二助手).Tag = ""
        lblInfo(e_第三助手).Tag = ""
        
        mrsAppend.Filter = "中文名='主刀医生'"
        
        mrsAppend.Sort = "序号"
        If Not mrsAppend.EOF Then
            lblInfo(e_主刀医生).Tag = Val(mrsAppend!序号 & "")
            lblInfo(e_主刀医生).ToolTipText = mrsAppend!项目 & ""
            strFilter = strFilter & " And 序号<>" & Val(mrsAppend!序号 & "")
            If bln取值 Then
                txtInfo(e_主刀医生).Text = GetItemAppend(Val(mrsAppend!要素ID & ""), mrsAppend!中文名 & "", mrsAppend!项目 & "")
                txtInfo(e_主刀医生).Tag = txtInfo(e_主刀医生).Text
            End If
        End If
        
        mrsAppend.Filter = "中文名='主刀医生科室'"
        mrsAppend.Sort = "序号"
        
        If Not mrsAppend.EOF Then
            lblInfo(e_主刀医生科室).Tag = Val(mrsAppend!序号 & "")
            lblInfo(e_主刀医生科室).ToolTipText = mrsAppend!项目 & ""
            strFilter = strFilter & " And 序号<>" & Val(mrsAppend!序号 & "")
            
            If bln取值 Then
                txtInfo(e_主刀医生科室).Text = GetItemAppend(Val(mrsAppend!要素ID & ""), mrsAppend!中文名 & "", mrsAppend!项目 & "")
                txtInfo(e_主刀医生科室).Tag = txtInfo(e_主刀医生科室).Text
            End If
        End If
        
        mrsAppend.Filter = "中文名='助手医生'"
        mrsAppend.Sort = "序号"
        If Not mrsAppend.EOF Then
            For i = 1 To mrsAppend.RecordCount
                If i = 1 Then
                    intIndex = e_第一助手
                ElseIf i = 2 Then
                    intIndex = e_第二助手
                ElseIf i = 3 Then
                    intIndex = e_第三助手
                Else
                    Exit For
                End If
                
                lblInfo(intIndex).Tag = Val(mrsAppend!序号 & "")
                lblInfo(intIndex).ToolTipText = mrsAppend!项目 & ""
                
                strFilter = strFilter & " And 序号<>" & Val(mrsAppend!序号 & "")
                
                If bln取值 Then
                    txtInfo(intIndex).Text = GetItemAppend(Val(mrsAppend!要素ID & ""), mrsAppend!中文名 & "", mrsAppend!项目 & "")
                    txtInfo(intIndex).Tag = txtInfo(intIndex).Text
                End If
                mrsAppend.MoveNext
            Next
        End If
        
        If strFilter <> "" Then
            strFilter = Mid(strFilter, 5)
            mrsAppend.Filter = strFilter
        Else
            mrsAppend.Filter = 0
        End If
        mrsAppend.Sort = "序号"
        
        If Not mrsAppend.EOF Then
            '加载扩展表格
            vsOther.Visible = True
            lblInfo(e_其它内容).Visible = True
            Me.Height = 10200
            i = mrsAppend.RecordCount
            With vsOther
                .Clear
                .Rows = IIF(i = 0, 2, i + 1)
                .Cols = 2
                .FixedRows = 1: .FixedCols = 1
                
                .ColAlignment(COL_项目名称) = 1
                .FixedAlignment(COL_项目名称) = 4
                .ColWidth(COL_项目名称) = 2000
                .TextMatrix(0, COL_项目名称) = "申请附项扩展"
                
                .ColAlignment(col_内容) = 1
                .FixedAlignment(col_内容) = 4
                 .ColWidth(col_内容) = 6200
                 .TextMatrix(0, col_内容) = "内容"
                 
                .Editable = flexEDKbdMouse
                
                For i = 1 To mrsAppend.RecordCount
                    .RowData(i) = Val(mrsAppend!序号 & "")
                    .TextMatrix(i, COL_项目名称) = mrsAppend!项目 & ""
                    
                    If bln取值 Then .TextMatrix(i, col_内容) = GetItemAppend(Val(mrsAppend!要素ID & ""), mrsAppend!中文名 & "", mrsAppend!项目 & "")
                    
                    mrsAppend.MoveNext
                Next
                .Row = 1: .Col = 0
            End With
        Else
            Me.Height = 8500
        End If
    End If
    '没有绑定则取出几个固定附项
    '如果没有绑定附项，找出要素ID，存入 Tag 中。助手医生是同一个要素，只需要判断  “第一助手” 即可
    If lblInfo(e_主刀医生科室).Tag = "" Or lblInfo(e_主刀医生).Tag = "" Or lblInfo(e_第一助手).Tag = "" Then
        strSQL = "Select i.Id As 要素id, i.中文名 From 诊治所见项目 I, 诊治所见分类 K" & _
            " Where i.分类id = k.Id And k.性质 = 1 And k.编码 = '06' And i.中文名 In ('助手医生', '主刀医生', '主刀医生科室')"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        For i = 1 To rsTmp.RecordCount
            If rsTmp!中文名 & "" = "主刀医生" And lblInfo(e_主刀医生).Tag = "" Then
                lblInfo(e_主刀医生).Tag = Val(rsTmp!要素ID & "")
                If txtInfo(e_主刀医生).Text = "" Then
                    txtInfo(e_主刀医生).Text = GetItemAppend(Val(rsTmp!要素ID & ""), "主刀医生", "主刀医生")
                    txtInfo(e_主刀医生).Tag = txtInfo(e_主刀医生).Text
                End If
            ElseIf rsTmp!中文名 & "" = "主刀医生科室" And lblInfo(e_主刀医生科室).Tag = "" Then
                lblInfo(e_主刀医生科室).Tag = Val(rsTmp!要素ID & "")
                If txtInfo(e_主刀医生科室).Text = "" Then
                    txtInfo(e_主刀医生科室).Text = GetItemAppend(Val(rsTmp!要素ID & ""), "主刀医生科室", "主刀医生科室")
                    txtInfo(e_主刀医生科室).Tag = txtInfo(e_主刀医生科室).Text
                End If
            ElseIf rsTmp!中文名 & "" = "助手医生" And lblInfo(e_第一助手).Tag = "" Then
                lblInfo(e_第一助手).Tag = Val(rsTmp!要素ID & "")
            End If
            rsTmp.MoveNext
        Next
    End If
    Call Form_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Select Case Index
        Case e_手术时间
            Call cmdDate_Click(e_手术时间)
        Case e_生效时间
            Call cmdDate_Click(e_生效时间)
        End Select
    End If
End Sub

Private Sub vsOper_DblClick()
    Call vsOper_KeyPress(32)
End Sub

Private Sub vsOper_GotFocus()
    Call vsOper_AfterRowColChange(vsOper.Row, vsOper.Col, vsOper.Row, vsOper.Col)
End Sub

Private Sub vsOper_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：删除数据行
    Dim i As Long
    If mintType = 2 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        With vsOper
            If .RowData(.Row) <> 0 Then
                If MsgBox("要删除当前行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                .RowData(.Row) = 0
                For i = 0 To .Cols - 1
                    .TextMatrix(.Row, i) = ""
                    .Cell(flexcpData, .Row, i) = ""
                Next
                If Not (.Rows = .FixedRows + 1 And .Row = .FixedRows) Then .RemoveItem .Row
            End If
        End With
    End If
End Sub

Private Sub vsOper_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：非回车确认完后编辑的处理(这里Text:=EditText,但ValidateEdit事件中还没有)
    If Not mblnReturn Then
        If Col = 0 Then
            vsOper.TextMatrix(Row, Col) = CStr(vsOper.Cell(flexcpData, Row, Col))
            Call vsOper_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
        End If
    End If
    mblnChange = True
End Sub

Private Sub vsOper_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'''''''''''''
    If mintType = 2 Then Exit Sub
    
    If NewCol = COL_名称 Then
        vsOper.Editable = flexEDKbdMouse
        vsOper.ComboList = "..."
        vsOper.FocusRect = flexFocusLight
    ElseIf NewCol = COL_必要时 And vsOper.TextMatrix(NewRow, COL_名称) <> "" Then
        vsOper.ComboList = ""
        vsOper.Editable = flexEDKbdMouse
    Else
        vsOper.ComboList = ""
        vsOper.Editable = flexEDNone
    End If
End Sub

Private Sub vsOper_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'功能：直接打开附加手术项目选择器
    Dim strSQLItem As String, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim int性别 As Integer
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    int性别 = GetType性别
 
    strSQLItem = " From 诊疗项目目录 A , 疾病诊断对照 B, 疾病编码目录 C Where A.类别='F' And A.ID<>-1*" & mlng手术项目ID & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) And a.Id = b.手术id(+) And b.疾病id = c.Id(+)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[4])" & _
            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
            " And A.服务对象 IN([1],3) And Nvl(A.执行频率,0) IN(0,[2]) And Nvl(A.适用性别,0) IN(0,[3])"
    
    strSQL = "Select 0 as 末级,Max(Level) as 级ID,ID,上级ID,编码,名称,NULL as 单位,NULL as 手术级别,NULL as 执行科室性质ID,null as 计价性质ID" & _
        " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Start With ID In (Select a.分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID" & _
        " Group by ID,上级ID,编码,名称  Union ALL" & _
        " Select 1 as 末级,1 as 级ID,A.ID,a.分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,c.手术类型 As 手术级别,A.执行科室 as 执行科室性质ID,a.计价性质 as 计价性质ID" & _
        strSQLItem & " Order By 末级,级ID Desc,编码"
        
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "手术", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
        mint服务对象, 1, int性别, mlng病人科室id)
        
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "未找到可用的手术项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
        End If
        vsOper.TextMatrix(Row, Col) = CStr(vsOper.Cell(flexcpData, Row, Col))
        Call vsOper_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
        vsOper.SetFocus
        Exit Sub
    End If
    
    Call Set附加手术(vsOper.Row, rsTmp)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsOper_KeyPress(KeyAscii As Integer)
    If mintType = 2 Then Exit Sub
    With vsOper
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call SeekNextCtl
        Else
            If .Col = 0 Then
                .Editable = flexEDKbdMouse
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsOper_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            ElseIf .Col = COL_必要时 And .TextMatrix(.Row, COL_名称) <> "" Then
                .ComboList = ""
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End If
    End With
End Sub

Private Sub vsOper_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If mintType = 2 Then Exit Sub
    vsOper.EditSelStart = 0
    vsOper.EditSelLength = zlCommFun.ActualLen(vsOper.EditText)
End Sub

Private Sub vsOper_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：开始编辑时认为没有按下回车
    If mintType = 2 Then
        Cancel = True
        Exit Sub
    End If
    If Col = COL_必要时 And vsOper.TextMatrix(Row, COL_名称) <> "" Then
        vsOper.ComboList = ""
        vsOper.Editable = flexEDKbdMouse
    End If
    mblnReturn = False
End Sub
    

Private Sub vsOper_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'功能：附加手术输入数据确认
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, int性别 As Integer
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI, strLike As String
    
    On Error GoTo errH
    If mintType = 2 Then Exit Sub
    If KeyAscii = 13 Then
    
        mblnReturn = True '标记是按回车确认编辑
        
        KeyAscii = 0
        
        int性别 = GetType性别
        
        '优化
        strLike = gstrLike
        If Len(vsOper.EditText) < 2 Then strLike = ""
        
        strSQL = " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,d.手术类型 As 手术级别,A.执行科室 as 执行科室性质ID,a.计价性质 as 计价性质ID" & _
            " From 诊疗项目目录 A,诊疗项目别名 B, 疾病诊断对照 C, 疾病编码目录 D" & _
            " Where A.ID=B.诊疗项目ID And A.类别='F' And A.ID<>-1*[3]" & IIF(strLike = "", "", " And Rownum<=100") & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) And a.Id = c.手术id(+) And c.疾病id = d.Id(+)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]) And B.码类=[4]" & _
            " And A.服务对象 IN([5],3) And Nvl(A.执行频率,0) IN(0,[6]) And Nvl(A.适用性别,0) IN(0,[7])" & _
            " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[8])" & _
            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
            " Order by A.编码"
        vPoint = zlControl.GetCoordPos(vsOper.hwnd, vsOper.CellLeft, vsOper.CellTop)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "手术", False, "", "", False, False, True, vPoint.X, vPoint.Y, vsOper.CellHeight, blnCancel, False, True, _
            UCase(vsOper.EditText) & "%", strLike & UCase(vsOper.EditText) & "%", mlng手术项目ID, gbytCode + 1, mint服务对象, 1, int性别, mlng病人科室id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到匹配项目！", vbInformation, gstrSysName
            End If
            vsOper.TextMatrix(Row, Col) = CStr(vsOper.Cell(flexcpData, Row, Col))
            Call vsOper_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
            vsOper.SetFocus
            Exit Sub
        End If
        Call Set附加手术(Row, rsTmp)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsOther_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If mintType = 2 Then Exit Sub
    If KeyAscii = 13 Then
        With vsOther
            If Row = .Rows - 1 Then
                Call SeekNextCtl
            Else
                .Row = Row + 1
                .Col = col_内容
                .Refresh
            End If
        End With
    End If
End Sub

Private Sub vsOther_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mintType = 2 Then Cancel = True: Exit Sub
    mblnChange = True
End Sub

Private Function GetType性别() As Integer
'功能：根据当前病人性别获取 适用性别
    If txtInfo(e_性别).Text Like "*男*" Then
        GetType性别 = 1
    ElseIf txtInfo(e_性别).Text Like "*女*" Then
        GetType性别 = 2
    End If
End Function

Private Sub Set附加手术(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    Dim i As Integer
    Dim strMsg As String
    
    '附加手术
    With vsOper
        '检查重复输入
     
        If Val(rsInput!ID & "") = mlng手术项目ID Then
            strMsg = "该附加手术与主手项目相同。"
        Else
            For i = .FixedRows To .Rows - 1
                If .RowData(i) = Val(rsInput!ID) Then
                    strMsg = "该附加手术已经在其它行录入。"
                    Exit For
                End If
            Next
        End If
        
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            .TextMatrix(lngRow, COL_名称) = CStr(.Cell(flexcpData, lngRow, COL_名称))
            Call vsOper_AfterRowColChange(lngRow, COL_名称, lngRow, COL_名称) '重新使按钮可见
            .SetFocus
            Exit Sub
        End If
    
        .EditText = "[" & rsInput!编码 & "]" & rsInput!名称 '对于输入直接匹配时有必要
        
        .RowData(lngRow) = Val(rsInput!ID)
        
        .TextMatrix(lngRow, COL_名称) = "[" & rsInput!编码 & "]" & rsInput!名称
        .Cell(flexcpData, lngRow, COL_名称) = .TextMatrix(lngRow, COL_名称)
        
        .TextMatrix(lngRow, COL_手术级别) = NVL(rsInput!手术级别)
        
        .TextMatrix(lngRow, COL_执行性质) = Val(rsInput!执行科室性质ID & "")
        .TextMatrix(lngRow, COL_计价性质) = Val(rsInput!计价性质ID & "")
        
        '下一输入行
        If .RowData(.Rows - 1) <> 0 Then .AddItem ""
        .Row = .Rows - 1: .Col = COL_名称
        Call .ShowCell(.Row, .Col)
    End With
    
    mblnChange = True
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
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " 保存(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, " 保存退出(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " 退出(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_18", Me)
        Case conMenu_File_Preview: Call PrintApply(1)
        Case conMenu_File_Print: Call PrintApply(2)
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit '保存
            If CheckData = False Then Exit Sub
            If mint调用场合 = 0 Then
                mblnOK = SaveData
            Else
                mblnOK = SaveCacheData
            End If
            If Control.ID = conMenu_Edit_SaveExit Then
                Unload Me
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnVisible As Boolean
    
    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit
            Control.Enabled = mblnChange
        Case conMenu_File_PrintSet, conMenu_File_Preview, conMenu_File_Print
            blnVisible = mint调用场合 = 0
    End Select
    Control.Visible = blnVisible
End Sub


Private Sub Form_Load()
    
    mblnOK = False
    
    mbln提醒对码 = True
    If mint服务对象 = 0 Then mint服务对象 = 2 '缺省为住院
 
    '外院医生必须建档
    mbln外院医生建档 = Val(zlDatabase.GetPara(253, glngSys)) <> 0
        
    Call LoadPatiInfo
    
    Me.Height = 8500
    
    Call InitCommandBar
    
    With cboInfo(e_手术情况)
        .Clear
        .AddItem "择期"
        .AddItem "急诊"
        .AddItem "限期"
        .ListIndex = 0
    End With
    
    With cboInfo(e_再次手术类型)
        .Clear
        .AddItem "计划"
        .AddItem "非计划"
        .ListIndex = 0
    End With
    
    With vsOper
        .Clear
        .Rows = 2
        .Cols = 6
        .FixedRows = 1: .FixedCols = 0
        
        .ColAlignment(COL_名称) = 1
        .FixedAlignment(COL_名称) = 4
        .ColWidth(COL_名称) = 6000
        .TextMatrix(0, COL_名称) = "附加手术"
        
        .ColAlignment(COL_手术级别) = 1
        .FixedAlignment(COL_手术级别) = 4
         .ColWidth(COL_手术级别) = 1500
         .TextMatrix(0, COL_手术级别) = "级别"
         
         
        .ColAlignment(COL_必要时) = 1
        .FixedAlignment(COL_必要时) = 4
        .ColWidth(COL_必要时) = 1000
        .TextMatrix(0, COL_必要时) = "必要时"
        .ColDataType(COL_必要时) = flexDTBoolean
         
        .ColHidden(COL_计价性质) = True
        .ColHidden(COL_执行性质) = True
        .ColHidden(COL_摘要) = True
        
        .Editable = flexEDKbdMouse
        .Row = 1: .Col = 0
    End With
    Call InitDefault
    mblnChange = False
End Sub

Private Function SeekNextCtl() As Boolean
'功能：定位到下一个焦点的控件上
    Call zlCommFun.PressKey(vbKeyTab)
    SeekNextCtl = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("当前申请单已经进行了调整尚未保存，是否要继续退出？", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
    If Not mclsDiagEdit Is Nothing Then Set mclsDiagEdit = Nothing
    mlng麻醉项目ID = 0
    mbln补录 = False
    mstr入院时间 = ""
    mstr上次转科时间 = ""
    mint险类 = 0
    Set mclsMipModule = Nothing
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub Form_Resize()
    Dim lngL As Long
    Dim lngT As Long
    Dim lngDY As Long
    Dim lngTmp As Long

    On Error Resume Next
    
    lngL = 480
    lngDY = 200
    
    lblHead.Top = 750
    linHead.Y1 = lblHead.Height + lblHead.Top + 20
    linHead.Y2 = linHead.Y1
    
    picNo.Top = linHead.Y1 + 30
    picNo.Left = Me.ScaleWidth - picNo.Width - 400
    Call SetFaceItemPos(e_No)
    
    
    lblInfo(e_姓名).Left = lngL
    lblInfo(e_姓名).Top = picNo.Top + picNo.Height + 150
    
        lblInfo(e_性别).Left = lngL + 2695
        lblInfo(e_性别).Top = lblInfo(e_姓名).Top
        
        lblInfo(e_年龄).Left = lngL + 5340
        lblInfo(e_年龄).Top = lblInfo(e_姓名).Top
        
        lblInfo(e_病情).Left = lngL + 7785
        lblInfo(e_病情).Top = lblInfo(e_姓名).Top
        
        txtInfo(e_病情).Width = 800
        
        Call SetFaceItemPos(e_姓名)
        Call SetFaceItemPos(e_性别)
        Call SetFaceItemPos(e_年龄)
        Call SetFaceItemPos(e_病情)
    
    
    lblInfo(e_床号).Left = lngL
    lblInfo(e_床号).Top = Line1(e_姓名).Y1 + lngDY
    
        lblInfo(e_住院号).Left = lblInfo(e_性别).Left
        lblInfo(e_住院号).Top = lblInfo(e_床号).Top
        
        lblInfo(e_科室).Left = lblInfo(e_年龄).Left
        lblInfo(e_科室).Top = lblInfo(e_床号).Top
        
        txtInfo(e_住院号).Width = txtInfo(e_床号).Width
        txtInfo(e_科室).Width = txtInfo(e_床号).Width
        
        Call SetFaceItemPos(e_床号)
        Call SetFaceItemPos(e_住院号)
        Call SetFaceItemPos(e_科室)
        
        
    lblInfo(e_手术时间).Left = lngL - 30
    lblInfo(e_手术时间).Top = Line1(e_床号).Y1 + lngDY
        Call SetFaceItemPos(e_手术时间)
        cmdDate(e_手术时间).Top = txtInfo(e_手术时间).Top - 30
        cmdDate(e_手术时间).Left = txtInfo(e_手术时间).Left + txtInfo(e_手术时间).Width
        
        lblInfo(e_再次手术类型).Top = lblInfo(e_手术时间).Top
        lblInfo(e_再次手术类型).Left = lblInfo(e_手术情况).Left + lblInfo(e_手术情况).Width - lblInfo(e_再次手术类型).Width
        picInfo(e_再次手术类型).Top = lblInfo(e_再次手术类型).Top
        picInfo(e_再次手术类型).Left = lblInfo(e_再次手术类型).Left + lblInfo(e_再次手术类型).Width + 160
        Line1(e_再次手术类型).Y1 = picInfo(e_再次手术类型).Top + picInfo(e_再次手术类型).Height
        Line1(e_再次手术类型).Y2 = Line1(e_再次手术类型).Y1
        Line1(e_再次手术类型).X1 = picInfo(e_再次手术类型).Left - 100
        Line1(e_再次手术类型).X2 = Line1(e_再次手术类型).X1 + picInfo(e_再次手术类型).Width
        
    lblInfo(e_术前诊断).Left = lngL
    lblInfo(e_术前诊断).Top = Line1(e_手术时间).Y1 + lngDY
        txtInfo(e_术前诊断).Width = Me.ScaleWidth - txtInfo(e_术前诊断).Left - 700
        Call SetFaceItemPos(e_术前诊断)
        cmdInfo(e_术前诊断).Top = txtInfo(e_术前诊断).Top - 30
        cmdInfo(e_术前诊断).Left = txtInfo(e_术前诊断).Left + txtInfo(e_术前诊断).Width
        
        
    lblInfo(e_手术名称).Left = lngL - 30
    lblInfo(e_手术名称).Top = Line1(e_术前诊断).Y1 + lngDY
        Call SetFaceItemPos(e_手术名称)
        cmdInfo(e_手术名称).Top = txtInfo(e_手术名称).Top - 30
        cmdInfo(e_手术名称).Left = txtInfo(e_手术名称).Left + txtInfo(e_手术名称).Width
        
        lblInfo(e_手术级别).Left = lblInfo(e_病情).Left
        lblInfo(e_手术级别).Top = lblInfo(e_手术名称).Top
        txtInfo(e_手术级别).Width = 800
        Call SetFaceItemPos(e_手术级别)
        
        
    lblInfo(e_执行科室).Left = lngL
    lblInfo(e_执行科室).Top = Line1(e_手术名称).Y1 + lngDY
        picInfo(e_执行科室).Left = lblInfo(e_执行科室).Left + lblInfo(e_执行科室).Width + 160
        picInfo(e_执行科室).Top = lblInfo(e_执行科室).Top
        picInfo(e_执行科室).Width = 2400
        picInfo(e_执行科室).Height = 250
        cboInfo(e_执行科室).Width = 2450
        
        Line1(e_执行科室).Y1 = picInfo(e_执行科室).Top + picInfo(e_执行科室).Height
        Line1(e_执行科室).Y2 = Line1(e_执行科室).Y1
        Line1(e_执行科室).X1 = picInfo(e_执行科室).Left - 100
        Line1(e_执行科室).X2 = Line1(e_执行科室).X1 + picInfo(e_执行科室).Width + 110
        
        lblInfo(e_手术情况).Left = lblInfo(e_科室).Left
        lblInfo(e_手术情况).Top = lblInfo(e_执行科室).Top
        picInfo(e_手术情况).Left = lblInfo(e_手术情况).Left + lblInfo(e_手术情况).Width + 160
        picInfo(e_手术情况).Top = lblInfo(e_手术情况).Top
        picInfo(e_手术情况).Width = 1020
        picInfo(e_手术情况).Height = 250
        cboInfo(e_手术情况).Width = 1070
        Line1(e_手术情况).Y1 = picInfo(e_手术情况).Top + picInfo(e_手术情况).Height
        Line1(e_手术情况).Y2 = Line1(e_手术情况).Y1
        Line1(e_手术情况).X1 = picInfo(e_手术情况).Left - 100
        Line1(e_手术情况).X2 = Line1(e_手术情况).X1 + picInfo(e_手术情况).Width + 100
        
                
    lblInfo(e_附加手术).Left = lngL
    lblInfo(e_附加手术).Top = Line1(e_执行科室).Y1 + lngDY
        vsOper.Top = lblInfo(e_附加手术).Top
        vsOper.Left = picInfo(e_执行科室).Left
    
    
    lblInfo(e_麻醉方法).Left = lngL
    lblInfo(e_麻醉方法).Top = vsOper.Top + vsOper.Height + lngDY
        txtInfo(e_麻醉方法).Width = 3800
        
        Call SetFaceItemPos(e_麻醉方法)
        cmdInfo(e_麻醉方法).Top = txtInfo(e_麻醉方法).Top - 30
        cmdInfo(e_麻醉方法).Left = txtInfo(e_麻醉方法).Left + txtInfo(e_麻醉方法).Width
        
        lblInfo(e_麻醉执行科室).Left = lblInfo(e_科室).Left
        lblInfo(e_麻醉执行科室).Top = lblInfo(e_麻醉方法).Top
        picInfo(e_麻醉执行科室).Left = lblInfo(e_麻醉执行科室).Left + lblInfo(e_麻醉执行科室).Width + 160
        picInfo(e_麻醉执行科室).Top = lblInfo(e_麻醉执行科室).Top
        picInfo(e_麻醉执行科室).Width = 2400
        picInfo(e_麻醉执行科室).Height = 250
        cboInfo(e_麻醉执行科室).Width = 2450
        
        Line1(e_麻醉执行科室).Y1 = picInfo(e_麻醉执行科室).Top + picInfo(e_麻醉执行科室).Height + 10
        Line1(e_麻醉执行科室).Y2 = Line1(e_麻醉执行科室).Y1
        Line1(e_麻醉执行科室).X1 = picInfo(e_麻醉执行科室).Left - 100
        Line1(e_麻醉执行科室).X2 = Line1(e_麻醉执行科室).X1 + picInfo(e_麻醉执行科室).Width + 100
        
    
    lblInfo(e_主刀医生).Left = lngL
    lblInfo(e_主刀医生).Top = Line1(e_麻醉方法).Y1 + lngDY
        Call SetFaceItemPos(e_主刀医生)
        cmdInfo(e_主刀医生).Top = txtInfo(e_主刀医生).Top - 30
        cmdInfo(e_主刀医生).Left = txtInfo(e_主刀医生).Left + txtInfo(e_主刀医生).Width
            
        lblInfo(e_主刀医生科室).Left = lblInfo(e_住院号).Left
        lblInfo(e_主刀医生科室).Top = lblInfo(e_主刀医生).Top
        Call SetFaceItemPos(e_主刀医生科室)
        cmdInfo(e_主刀医生科室).Top = txtInfo(e_主刀医生科室).Top - 30
        cmdInfo(e_主刀医生科室).Left = txtInfo(e_主刀医生科室).Left + txtInfo(e_主刀医生科室).Width
        
    lblInfo(e_第一助手).Left = lngL
    lblInfo(e_第一助手).Top = Line1(e_主刀医生).Y1 + lngDY
        Call SetFaceItemPos(e_第一助手)
        cmdInfo(e_第一助手).Top = txtInfo(e_第一助手).Top - 30
        cmdInfo(e_第一助手).Left = txtInfo(e_第一助手).Left + txtInfo(e_第一助手).Width
        
        lblInfo(e_第二助手).Left = lblInfo(e_住院号).Left
        lblInfo(e_第二助手).Top = lblInfo(e_第一助手).Top
        txtInfo(e_第二助手).Width = txtInfo(e_第一助手).Width
        Call SetFaceItemPos(e_第二助手)
        cmdInfo(e_第二助手).Top = txtInfo(e_第二助手).Top - 30
        cmdInfo(e_第二助手).Left = txtInfo(e_第二助手).Left + txtInfo(e_第二助手).Width
        
        lblInfo(e_第三助手).Left = lblInfo(e_科室).Left
        lblInfo(e_第三助手).Top = lblInfo(e_第一助手).Top
        Call SetFaceItemPos(e_第三助手)
        cmdInfo(e_第三助手).Top = txtInfo(e_第三助手).Top - 30
        cmdInfo(e_第三助手).Left = txtInfo(e_第三助手).Left + txtInfo(e_第三助手).Width
        
    
    lblInfo(e_其它内容).Left = lngL
    lblInfo(e_其它内容).Top = Line1(e_第一助手).Y1 + lngDY
        vsOther.Top = lblInfo(e_其它内容).Top
        vsOther.Left = picInfo(e_执行科室).Left
    
    If lblInfo(e_其它内容).Visible Then
        lngTmp = vsOther.Top + vsOther.Height + lngDY
    Else
        lngTmp = Line1(e_第一助手).Y1 + lngDY
    End If
    
    lblInfo(e_申请科室).Left = lngL
    lblInfo(e_申请科室).Top = lngTmp 'vsOther.Top + vsOther.Height + lngDY
        Call SetFaceItemPos(e_申请科室)
        
        lblInfo(e_申请医师).Top = lblInfo(e_申请科室).Top
        lblInfo(e_申请医师).Left = 4200
        Call SetFaceItemPos(e_申请医师)
        
        lblInfo(e_主治医师).Top = lblInfo(e_申请科室).Top
        lblInfo(e_主治医师).Left = 7800
        Call SetFaceItemPos(e_主治医师)
        
        
    lblInfo(e_生效时间).Top = Line1(e_申请科室).Y1 + lngDY
        lblInfo(e_生效时间).Left = 7050
        Call SetFaceItemPos(e_生效时间)
        cmdDate(e_生效时间).Top = txtInfo(e_生效时间).Top - 30
        cmdDate(e_生效时间).Left = txtInfo(e_生效时间).Left + txtInfo(e_生效时间).Width
        
End Sub

Private Sub SetFaceItemPos(ByVal lngIndex As Long)
'功能：设置界面 标签，文本框，下划线的位置
'参数：lngIndex 控件下标索引
    Dim lngL As Long
    Dim lngT As Long
    
    lngL = lblInfo(lngIndex).Left
    lngT = lblInfo(lngIndex).Top
    
    txtInfo(lngIndex).Top = lngT
    txtInfo(lngIndex).Left = lngL + lblInfo(lngIndex).Width + 160
    txtInfo(lngIndex).Height = 200
    Line1(lngIndex).Y1 = lngT + txtInfo(lngIndex).Height + 10
    Line1(lngIndex).Y2 = Line1(lngIndex).Y1
    
    Line1(lngIndex).X1 = txtInfo(lngIndex).Left - 100
    Line1(lngIndex).X2 = 110 + txtInfo(lngIndex).Width + Line1(lngIndex).X1

End Sub

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
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_18", Me, "医嘱ID=" & mlngUpdateAdvice, "手术ID=" & mlng手术项目ID, intType)
End Sub

Private Function CheckData() As Boolean
'功能：检查数据正确性
    Dim strIDs As String, str医嘱内容 As String, strMsg As String
    Dim lngTmp As Long, i As Integer
    Dim vMsg As VbMsgBoxResult
    Dim strExtra As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng执行性质 As Long
    Dim lng执行科室ID As Long
    Dim lng就诊ID As Long
    Dim strTmp As String
    Dim j As Long
    Dim strTabAdvice As String
    Dim rsPrice As ADODB.Recordset
    
    Call Me.ValidateControls
    '必须录入主手术项目
    If mlng手术项目ID = 0 Then
        MsgBox "没有确定主手术项目。", vbInformation, Me.Caption
        If txtInfo(e_手术名称).Enabled Then txtInfo(e_手术名称).SetFocus
        Exit Function
    End If
    
    '检查执行科室
    If cboInfo(e_执行科室).Text = "" Then
        MsgBox "没有确定执行科室。", vbInformation, Me.Caption
        If cboInfo(e_执行科室).Enabled Then cboInfo(e_执行科室).SetFocus
        Exit Function
    End If
    mlng手术执行科室ID = cboInfo(e_执行科室).ItemData(cboInfo(e_执行科室).ListIndex)
    '麻醉
    If mlng麻醉项目ID <> 0 Then
        If cboInfo(e_麻醉执行科室).Text = "" Then
            MsgBox "没有确定麻醉执行科室。", vbInformation, Me.Caption
            If cboInfo(e_麻醉执行科室).Enabled Then cboInfo(e_麻醉执行科室).SetFocus
            Exit Function
        End If
        mlng麻醉执行科室id = cboInfo(e_麻醉执行科室).ItemData(cboInfo(e_麻醉执行科室).ListIndex)
    End If

    '检查时间合法性
    If Not Check开始时间(txtInfo(e_生效时间).Text) Then
        If txtInfo(e_生效时间).Enabled Then txtInfo(e_生效时间).SetFocus
        Exit Function
    End If
    
    If Not Check安排时间(txtInfo(e_手术时间).Text, txtInfo(e_生效时间).Text) Then
        If txtInfo(e_手术时间).Enabled Then txtInfo(e_手术时间).SetFocus
        Exit Function
    End If
    If Not Check附项 Then
        Exit Function
    End If
    
    '如果启用了手术授权管理，则检查主刀医师执行权
    If gbln手术授权管理 And mint调用类型 = 2 Then
        If CheckDocEmpowerEx() = False Then
            If Not gbln手术等级管理 Then
                MsgBox "主刀医生不具备此手术的执行权，不允许下达。", vbInformation, "手术授权管理"
                Exit Function
            Else
                MsgBox "主刀医生不具备此手术的执行权。", vbInformation, "手术授权管理"
            End If
        End If
    End If
    
    If mint调用场合 = 0 Then
        strTmp = mlng手术项目ID & "||" & mint调用类型
        mstr摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", strTmp)
    
        strExtra = "||0||" & mlng麻醉项目ID & "|| ||0"
        lng就诊ID = IIF(mint场合 = 1, mlng挂号ID, mlng主页ID)
        
        lng执行性质 = IIF(cboInfo(e_执行科室).ItemData(cboInfo(e_执行科室).ListIndex) <= 0, 5, mlng手术执行科室性质)
        
        lng执行科室ID = IIF(cboInfo(e_执行科室).ItemData(cboInfo(e_执行科室).ListIndex) <= 0, 0, cboInfo(e_执行科室).ItemData(cboInfo(e_执行科室).ListIndex))
    
        If Not AdviceCheck(lng就诊ID, "F", mlng手术项目ID, lng执行科室ID, lng执行性质, mstr摘要 & strExtra) Then
            Exit Function
        End If
        j = 1
        strTabAdvice = "select " & j & " as ID," & j & " as 序号,-null as 相关ID,'F' as 诊疗类别," & mlng手术项目ID & " as 管码项目ID," & _
            mlng手术项目ID & " as 诊疗项目ID,1 As 总量, 0 As 单量,null as 标本部位,null As 检查方法," & _
            "0 as 执行标记,0 as 计价特性, null As 附加手术," & lng执行性质 & " As 执行性质," & lng执行科室ID & " as 执行科室id from dual"
    
        '对码检查
        str医嘱内容 = FormatAdviceContext
        
        strIDs = mlng手术项目ID & ":" & lng执行科室ID
        
        With vsOper
            For i = .FixedRows To .Rows - 1
                If .RowData(i) <> 0 Then
                    strIDs = strIDs & "," & .RowData(i) & ":" & lng执行科室ID
                    .TextMatrix(i, COL_摘要) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(.RowData(i)) & "||" & mint调用类型)
                    If Not AdviceCheck(lng就诊ID, "F", Val(.RowData(i)), lng执行科室ID, Val(.TextMatrix(i, COL_执行性质)), .TextMatrix(i, COL_摘要) & strExtra) Then
                        Exit Function
                    End If
                    j = j + 1
                    strTabAdvice = strTabAdvice & " Union ALL " & _
                        "select " & j & " as ID," & j & " as 序号,1 as 相关ID,'F' as 诊疗类别," & Val(.RowData(i)) & " as 管码项目ID," & _
                             Val(.RowData(i)) & " as 诊疗项目ID,1 As 总量, 0 As 单量,null as 标本部位,null As 检查方法," & _
                             Abs(Val(.TextMatrix(i, COL_必要时))) & " as 执行标记,0 as 计价特性, 1 As 附加手术," & Val(.TextMatrix(i, COL_执行性质)) & " As 执行性质," & lng执行科室ID & " as 执行科室id from dual"
                End If
            Next
        End With
        
        If mlng麻醉项目ID <> 0 Then
        
            lng执行科室ID = IIF(cboInfo(e_麻醉执行科室).ItemData(cboInfo(e_麻醉执行科室).ListIndex) <= 0, 0, cboInfo(e_麻醉执行科室).ItemData(cboInfo(e_麻醉执行科室).ListIndex))
            
            strIDs = strIDs & "," & mlng麻醉项目ID & ":" & lng执行科室ID
  
            lblInfo(e_麻醉执行科室).Tag = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(mlng麻醉项目ID) & "||" & mint调用类型)
            
            If Not AdviceCheck(lng就诊ID, "G", mlng麻醉项目ID, lng执行科室ID, lng执行性质, lblInfo(e_麻醉执行科室).Tag & strExtra) Then
                Exit Function
            End If
            
            j = j + 1
            strTabAdvice = strTabAdvice & " Union ALL " & _
                "select " & j & " as ID," & j & " as 序号,1 as 相关ID,'F' as 诊疗类别," & mlng麻醉项目ID & " as 管码项目ID," & _
                    mlng麻醉项目ID & " as 诊疗项目ID,1 As 总量, 0 As 单量,null as 标本部位,null As 检查方法," & _
                    "0 as 执行标记,0 as 计价特性, null As 附加手术," & lng执行性质 & " As 执行性质," & lng执行科室ID & " as 执行科室id from dual"
        End If
        
        If gint医保对码 = 2 Then mbln提醒对码 = True
    
        strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, IIF(mlng病人性质 = 0, 2, 1), "", strIDs, str医嘱内容)
        
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
        
        '医保管控实时监测
        If mint险类 <> 0 Then
            If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) Then
                If MakePriceRecord申请单("4" & mint调用类型, mlng病人ID, lng就诊ID, strTabAdvice, strIDs, mstr费别, mlng开单科室ID, rsPrice) Then
                    If Not gclsInsure.CheckItem(mint险类, IIF(mint调用类型 = 1, 0, 1), 0, rsPrice) Then
                        MsgBox "医保监测检查未通(执行Insure.CheckItem接口)，本次下达的手术申请单不能保存。", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
    End If
    CheckData = True
End Function

Private Function SaveCacheData() As Boolean
'功能：缓存数据
    Dim strTmp As String
    Dim str必要时 As String
    
    Dim i As Long
    On Error GoTo errH
    If mrsCard Is Nothing Then
         Call InitCardRsOperate(mrsCard)
         mrsCard.AddNew
    End If
    
    mrsCard!临床诊断描述 = txtInfo(e_术前诊断).Text
    mrsCard!临床诊断IDs = mstr诊断IDs
    mrsCard!手术情况 = cboInfo(e_手术情况).ListIndex
    mrsCard!主手术项目ID = mlng手术项目ID
    mrsCard!手术执行科室ID = mlng手术执行科室ID
    '附加手术
    With vsOper
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                strTmp = strTmp & "," & Val(.RowData(i))
                str必要时 = str必要时 & "," & Val(.RowData(i)) & ":" & Abs(Val(.TextMatrix(i, COL_必要时)))
            End If
        Next
    End With
    mrsCard!附手术项目IDs = Mid(strTmp, 2)
    mrsCard!附手术必要时 = Mid(str必要时, 2)
    mrsCard!麻醉项目ID = mlng麻醉项目ID
    mrsCard!手术执行科室ID = mlng手术执行科室ID
    mrsCard!麻醉执行科室ID = mlng麻醉执行科室id
    mrsCard!生效时间 = txtInfo(e_生效时间).Text
    mrsCard!手术时间 = txtInfo(e_手术时间).Text
    mrsCard!申请科室id = mlng开单科室ID
    
    strTmp = ""
    mrsAppend.Filter = 0
    mrsAppend.Filter = "项目='主刀医生'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>主刀医生<Split2>" & Val(mrsAppend!必填 & "") & "<Split2>" & mrsAppend!要素ID & "<Split2>" & txtInfo(e_主刀医生).Text
    Else
        strTmp = strTmp & "<Split1>主刀医生<Split2>0<Split2>0<Split2>" & txtInfo(e_主刀医生).Text
    End If
     
    mrsAppend.Filter = "项目='主刀医生科室'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>主刀医生科室<Split2>" & Val(mrsAppend!必填 & "") & "<Split2>" & mrsAppend!要素ID & "<Split2>" & txtInfo(e_主刀医生科室).Text
    Else
        strTmp = strTmp & "<Split1>主刀医生科室<Split2>0<Split2>0<Split2>" & txtInfo(e_主刀医生科室).Text
    End If
    
    mrsAppend.Filter = "项目='第一助手'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>第一助手<Split2>" & Val(mrsAppend!必填 & "") & "<Split2>" & mrsAppend!要素ID & "<Split2>" & txtInfo(e_第一助手).Text
    Else
        strTmp = strTmp & "<Split1>第一助手<Split2>0<Split2>0<Split2>" & txtInfo(e_第一助手).Text
    End If
    
    mrsAppend.Filter = "项目='第二助手'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>第二助手<Split2>" & Val(mrsAppend!必填 & "") & "<Split2>" & mrsAppend!要素ID & "<Split2>" & txtInfo(e_第二助手).Text
    Else
        strTmp = strTmp & "<Split1>第二助手<Split2>0<Split2>0<Split2>" & txtInfo(e_第二助手).Text
    End If
    
    mrsAppend.Filter = "项目='第三助手'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>第三助手<Split2>" & Val(mrsAppend!必填 & "") & "<Split2>" & mrsAppend!要素ID & "<Split2>" & txtInfo(e_第三助手).Text
    Else
        strTmp = strTmp & "<Split1>第三助手<Split2>0<Split2>0<Split2>" & txtInfo(e_第三助手).Text
    End If
    
    If lblInfo(e_再次手术类型).Visible Then
        strTmp = strTmp & "<Split1>再次手术类型<Split2>0<Split2>0<Split2>" & cboInfo(e_再次手术类型).Text
    End If
    
    With vsOther
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col_内容) <> "" Then
                mrsAppend.Filter = "项目='" & .TextMatrix(i, COL_项目名称) & "'"
                If Not mrsAppend.EOF Then
                    strTmp = strTmp & "<Split1>" & .TextMatrix(i, COL_项目名称) & "<Split2>" & Val(mrsAppend!必填 & "") & "<Split2>" & mrsAppend!要素ID & "<Split2>" & .TextMatrix(i, col_内容)
                End If
            End If
        Next
    End With
    
    If mstr诊断IDs <> "" Then
        strTmp = strTmp & "<Split1>申请单诊断<Split2>0<Split2>0<Split2>" & txtInfo(e_术前诊断).Text
    End If
    
    mrsCard!申请附项 = Mid(strTmp, 9)
    
    mrsCard.Update
    mblnChange = False
    SaveCacheData = True
    Exit Function
errH:
    If 2 = 1 Then
        Resume
    End If
    err.Clear
End Function

Private Function SaveData() As Boolean
'功能：保存数据
    Dim lng医嘱ID As Long, lng相关ID As Long, lng医嘱序号 As Long, lng申请序号 As Long
    Dim str医嘱内容 As String, str开嘱时间 As String, strTmp As String
    Dim lng执行科室ID As Long, arrSQLTmp As Variant
    Dim strSQL As String, arrSQL As Variant, blnTrans As Boolean
    Dim datCur As Date, str主页ID As String, str挂号单 As String
    Dim rsAffer As ADODB.Recordset, lngTmp As Long
    Dim i As Long, int紧急 As Integer
    Dim strSource As String 'SQL模版
    Dim rsTmp As ADODB.Recordset
    Dim str审核状态 As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    arrSQL = Array()
    
    If mintType = 1 Then
        strSQL = "Zl_病人医嘱记录_Delete(" & mlngUpdateAdvice & ",1)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        If gbln启用影像信息系统预约 Then
            Set rsTmp = Nothing
            Set rsTmp = GetDataRIS预约(CStr(mlngUpdateAdvice))
            On Error Resume Next
            If Not rsTmp.EOF Then
                If 0 <> gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!预约id & "")) Then
                    MsgBox "当前启用了影像信息系统接口，本次操作删除或修改了已经预约医嘱，但由于影像信息系统接口(HISSchedulingEx)取消息预约未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
                End If
                rsTmp.MoveNext
            End If
            err.Clear: On Error GoTo errH
        End If
    End If
    
    lng医嘱序号 = GetMaxAdviceNO(mlng病人ID, mlng主页ID, mbytBaby, mstr挂号单)
    If lng申请序号 = 0 Then lng申请序号 = Get申请序号
    mlngUpdateAdvice = zlDatabase.GetNextID("病人医嘱记录")
    
    If mint场合 = 0 Then
        str主页ID = mlng主页ID
        str挂号单 = "NULL"
    Else
        str主页ID = "NULL"
        str挂号单 = "'" & mstr挂号单 & "'"
    End If
    
    If cboInfo(e_手术情况).ListIndex = 1 Then int紧急 = 1
    '门诊场合手术医嘱不用审核
    str审核状态 = IIF(gbln手术分级管理 And int紧急 = 0 And mint调用类型 = 2, "1", "NULL")
    If str审核状态 = "1" Then
        If gbln手术授权管理 Then
            If gbln手术等级管理 Then
                If CheckDocEmpowerEx() Then
                    str审核状态 = "Null"
                End If
            End If
        Else
            If gbln手术等级管理 Then
                If CheckDoc手术等级 Then
                    str审核状态 = "Null"
                End If
            End If
        End If
    End If
    datCur = zlDatabase.Currentdate
    str开嘱时间 = IIF(datCur > CDate(txtInfo(e_生效时间).Text), txtInfo(e_生效时间).Text, datCur)
    '共9个项目依次为：[ID],[相关ID],[序号],[诊疗类别],[诊疗项目ID],[医嘱内容],[计价特性],[执行科室ID],[执行性质],[医生嘱托],[执行标记]
    strSource = "ZL_病人医嘱记录_Insert([0],[1],[2]," & mint调用类型 & "," & mlng病人ID & "," & str主页ID & "," & mbytBaby & ",1,1,'[3]',[4],NULL,NULL,NULL,1,'[5]',[9]," & _
        "'" & txtInfo(e_手术时间).Text & "','一次性',NULL,NULL,NULL,NULL,[6],[7],[8]," & IIF(mbln补录, 2, int紧急) & "," & _
        "To_Date('" & Format(txtInfo(e_生效时间).Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
        mlng病人科室id & "," & mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & _
        "To_Date('" & Format(str开嘱时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
        str挂号单 & "," & ZVal(mlng前提ID) & ",NULL,[10],NULL,[11],'" & UserInfo.姓名 & "',Null,NULL,NULL," & str审核状态 & _
        "," & lng申请序号 & ",NULL,NULL,NULL," & ZVal(cboInfo(e_手术情况).ListIndex) & ")"
    
    
    '主手术项目，手术医嘱的医生嘱托只存了主医嘱行中
    lng医嘱序号 = lng医嘱序号 + 1    '病人医嘱记录.序号，递增
    str医嘱内容 = FormatAdviceContext
    With cboInfo(e_执行科室)
        lng执行科室ID = IIF(.ItemData(.ListIndex) <= 0, 0, .ItemData(.ListIndex))
    End With
    strTmp = "Null"
    strSQL = GetStrExcSQL(strSource, mlngUpdateAdvice, "NULL", lng医嘱序号, "F", mlng手术项目ID, str医嘱内容, Val(lblInfo(e_手术名称).Tag), ZVal(lng执行科室ID), IIF(lng执行科室ID <= 0, "5", mlng手术执行科室性质), "Null", 0, IIF(mstr摘要 = "", "null", "'" & mstr摘要 & "'"))
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    
    lng相关ID = mlngUpdateAdvice
    
    '附加手术
    With vsOper
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                str医嘱内容 = .TextMatrix(i, COL_名称)
                str医嘱内容 = Mid(str医嘱内容, InStr(str医嘱内容, "]") + 1)
                strTmp = IIF(.TextMatrix(i, COL_摘要) = "", "null", "'" & .TextMatrix(i, COL_摘要) & "'")
                lng医嘱ID = zlDatabase.GetNextID("病人医嘱记录")
                lng医嘱序号 = lng医嘱序号 + 1
                strSQL = GetStrExcSQL(strSource, lng医嘱ID, lng相关ID, lng医嘱序号, "F", Val(.RowData(i)), str医嘱内容, Val(.TextMatrix(i, COL_计价性质)), ZVal(lng执行科室ID), Val(.TextMatrix(i, COL_执行性质)), "NULL", Abs(Val(.TextMatrix(i, COL_必要时))), strTmp)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        Next
    End With
    
    '麻醉项目
    If mlng麻醉项目ID <> 0 Then
        lng执行科室ID = 0
        With cboInfo(e_麻醉执行科室)
            lng执行科室ID = IIF(.ItemData(.ListIndex) <= 0, 0, .ItemData(.ListIndex))
        End With
        
        str医嘱内容 = txtInfo(e_麻醉方法).Text
        str医嘱内容 = Mid(str医嘱内容, InStr(str医嘱内容, "]") + 1)
        strTmp = IIF(lblInfo(e_麻醉执行科室).Tag = "", "null", "'" & lblInfo(e_麻醉执行科室).Tag & "'")
        lng医嘱ID = zlDatabase.GetNextID("病人医嘱记录")
        
        lng医嘱序号 = lng医嘱序号 + 1
        strSQL = GetStrExcSQL(strSource, lng医嘱ID, lng相关ID, lng医嘱序号, "G", mlng麻醉项目ID, str医嘱内容, Val(lblInfo(e_麻醉方法).Tag), ZVal(lng执行科室ID), IIF(lng执行科室ID <= 0, "5", mlng麻醉执行科室性质), "NULL", 0, strTmp)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
    
    '申请附项
    lngTmp = GetStrExcSQL申请附项(lng相关ID, arrSQLTmp)
    For i = 0 To UBound(arrSQLTmp)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = CStr(arrSQLTmp(i))
    Next
    
    '诊断关联信息
    If mstr诊断IDs <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(" & lng相关ID & ",'" & mstr诊断IDs & "')"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng相关ID & ",'申请单诊断',null," & lngTmp + 1 & ",null,'" & txtInfo(e_术前诊断).Text & "',0)"
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If mint场合 = 0 Then
        If gbln手术分级管理 And int紧急 = 0 Then
            Call ZLHIS_CIS_Audit("ZLHIS_CIS_028", mclsMipModule, mlng病人ID, txtInfo(e_姓名).Text, txtInfo(e_住院号).Text, , IIF(mlng病人性质 = 1, 1, 2), _
                mlng主页ID, mlng病区ID, , mlng病人科室id, "", , txtInfo(e_床号).Text, _
                lng医嘱ID, UserInfo.姓名, Format(str开嘱时间, "yyyy-MM-dd HH:mm:ss"), mlng开单科室ID, "", , "")
        Else
            Call ZLHIS_CIS_001(mclsMipModule, mlng病人ID, txtInfo(e_姓名).Text, txtInfo(e_住院号).Text, , IIF(mlng病人性质 = 1, 1, 2), _
                mlng主页ID, mlng病区ID, , mlng病人科室id, "", , txtInfo(e_床号).Text, _
                lng医嘱ID, IIF(int紧急 = 1, 1, 0), 1, "F", "", UserInfo.姓名, Format(str开嘱时间, "yyyy-MM-dd HH:mm:ss"), mlng开单科室ID, "", , , "")
        End If
    End If
    
    If mstr开单科室 = "" Then
        strSQL = "select 名称 from 部门表 where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng开单科室ID)
        mstr开单科室 = rsTmp!名称 & ""
        txtInfo(e_申请科室).Text = mstr开单科室
        txtInfo(e_申请医师).Text = UserInfo.姓名
    End If
    
    If gbln启用影像信息系统预约 Then
        On Error Resume Next
        Call gobjRis.HISScheduling(IIF(1 = mint场合, 1, 2), lng相关ID, mlng手术项目ID)
        err.Clear
    End If
    
    mintType = 1
    SaveData = True
    mblnChange = False
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

Private Sub ReSetAdviceNo(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
'功能：重新整理医嘱序号
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Count(*) as Num From (Select 序号,Count(ID) From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] Having Count(ID)>1 Group by 序号)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    If rsTmp.EOF Then Exit Sub
    
    If NVL(rsTmp!Num, 0) = 0 Then Exit Sub
    
    strSQL = "ZL_病人医嘱记录_更新序号(NULL,NULL," & mlng病人ID & "," & mlng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
 
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetAdviceNO(ByRef lngMinNo As Long, ByRef lngMaxNo As Long, ByRef lngNo As Long)
'功能：获取当前手医嘱的最大和最小序号,lngNo-申请序号
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "select min(序号) as 最小,max(序号) as 最大,max(申请序号) as 申请序号 from 病人医嘱记录 where id=[1] or 相关id=[1]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
    
    lngMinNo = rsTmp!最小
    lngMaxNo = rsTmp!最大
    lngNo = Val(rsTmp!申请序号 & "")
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetAfferAdvice() As ADODB.Recordset
'功能：获取当前手术医嘱之后的医嘱
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "select id as 医嘱id from 病人医嘱记录 where 病人id=[1] and nvl(婴儿,0)=[2]" & _
        IIF(mint场合 = 0, " and 主页id=[3]", " and 挂号单=[4]") & " order by 序号"
    
    On Error GoTo errH
    
    Set GetAfferAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, 0, mlng主页ID, mstr挂号单)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function


Private Function GetStrExcSQL申请附项(ByVal lng医嘱ID As Long, ByRef arrSQL As Variant) As Long
'功能：返加申请附项的可执行SQL，存于varSQL中，如果没有则函数反回 false， strItems 用医嘱编辑调用时
'返回：true 有，false 无
    Dim intCount As Integer, i As Integer
    Dim lng助手要素ID As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    arrSQL = Array()
    
    strTmp = txtInfo(e_主刀医生).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    mrsAppend.Filter = "序号=" & Val(lblInfo(e_主刀医生).Tag)
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'" & mrsAppend!项目 & "'," & mrsAppend!必填 & "," & intCount & "," & mrsAppend!要素ID & "," & strTmp & ",1)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'主刀医生',0," & intCount & "," & Val(lblInfo(e_主刀医生).Tag) & "," & strTmp & ",1)"
    End If
    
    strTmp = txtInfo(e_主刀医生科室).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    mrsAppend.Filter = "序号=" & Val(lblInfo(e_主刀医生科室).Tag)
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'" & mrsAppend!项目 & "'," & mrsAppend!必填 & "," & intCount & "," & mrsAppend!要素ID & "," & strTmp & ",0)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'主刀医生科室',0," & intCount & "," & Val(lblInfo(e_主刀医生科室).Tag) & "," & strTmp & ",0)"
    End If
 
    
    '先取出助手要素ID-------------------
    mrsAppend.Filter = "序号=" & Val(lblInfo(e_第一助手).Tag)
    If Not mrsAppend.EOF Then
        lng助手要素ID = Val(mrsAppend!要素ID & "")
    Else
        lng助手要素ID = Val(lblInfo(e_第一助手).Tag)
    End If
    
    strTmp = txtInfo(e_第一助手).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'" & mrsAppend!项目 & "'," & mrsAppend!必填 & "," & intCount & "," & lng助手要素ID & "," & strTmp & ",0)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'第一助手',0," & intCount & "," & lng助手要素ID & "," & strTmp & ",0)"
    End If
 
    '------------------------------
    strTmp = txtInfo(e_第二助手).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    mrsAppend.Filter = "序号=" & Val(lblInfo(e_第二助手).Tag)
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'" & mrsAppend!项目 & "'," & mrsAppend!必填 & "," & intCount & "," & lng助手要素ID & "," & strTmp & ",0)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'第二助手',0," & intCount & "," & lng助手要素ID & "," & strTmp & ",0)"
    End If
    
    strTmp = txtInfo(e_第三助手).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    mrsAppend.Filter = "序号=" & Val(lblInfo(e_第三助手).Tag)
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'" & mrsAppend!项目 & "'," & mrsAppend!必填 & "," & intCount & "," & lng助手要素ID & "," & strTmp & ",0)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'第三助手',0," & intCount & "," & lng助手要素ID & "," & strTmp & ",0)"
    End If
    
    If lblInfo(e_再次手术类型).Visible Then
        intCount = intCount + 1
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'再次手术类型',0," & intCount & ",null,'" & cboInfo(e_再次手术类型).Text & "',0)"
    End If
    
    '表格中的附项
    With vsOther
        If .Visible Then
            For i = .FixedRows To .Rows - 1
                strTmp = .TextMatrix(i, col_内容)
                strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
                mrsAppend.Filter = "序号=" & Val(.RowData(i))
                intCount = intCount + 1
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng医嘱ID & ",'" & mrsAppend!项目 & "'," & mrsAppend!必填 & "," & intCount & "," & ZVal(mrsAppend!要素ID & "") & "," & strTmp & ",0)"
            Next
        End If
    End With
    
    GetStrExcSQL申请附项 = intCount
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetStrExcSQL(ByVal strSource As String, ParamArray arrInput() As Variant) As String
'功能：生成 ZL_病人医嘱记录_Insert过程语句，arrInput参数 [ID],[相关ID],[序号],[诊疗类别],[诊疗项目ID],[医嘱内容],[计价特性],[执行科室ID],[执行性质]
    Dim i As Integer
    Dim strTmp As String
    Dim strResult As String
    
    strResult = strSource
    For i = 0 To UBound(arrInput)
        strTmp = arrInput(i)
        strResult = Replace(strResult, "[" & i & "]", strTmp)
    Next
    GetStrExcSQL = strResult
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

Private Function GetMax手术等级(ByVal str手术项目 As String) As String
'功能：取得当前医嘱最大手术类型
'参数：str手术项目：手术项目ID用，分隔，lng手术等级返回最高的手术等级
    Dim strSQL As String, rsTmp As Recordset
    Dim str手术等级 As String, i As Integer
    
    On Error GoTo errH
    strSQL = "Select a.手术类型 From 疾病编码目录 A,疾病诊断对照 B Where a.ID=b.疾病ID And a.类别='S' And instr([1], b.手术id)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str手术项目)
    For i = 1 To rsTmp.RecordCount
        If Decode(rsTmp!手术类型 & "", "一级", 1, "二级", 2, "三级", 3, "四级", 4, 0) > Decode(str手术等级, "一级", 1, "二级", 2, "三级", 3, "四级", 4, 0) Then
            str手术等级 = rsTmp!手术类型 & ""
        End If
        rsTmp.MoveNext
    Next
    GetMax手术等级 = str手术等级
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get诊疗项目记录ID(ByVal lngID As Long, Optional ByVal strIDs As String) As ADODB.Recordset
'功能：读取指定诊疗项目记录的ID
'参数：
    Dim strSQL As String
    
    strSQL = "Select " & vbNewLine & _
        " a.计算规则, a.站点, a.类别, a.分类id, a.Id, a.编码, a.名称, a.标本部位, a.计算单位, a.计算方式, a.执行频率, a.适用性别, a.单独应用, a.组合项目, c.手术类型 as 手术级别, a.执行安排," & vbNewLine & _
        " a.执行科室, a.服务对象, a.计价性质, a.参考目录id, a.人员id, a.建档时间, a.撤档时间, a.录入限量, a.试管编码, a.执行分类, a.执行标记" & vbNewLine & _
        "From 诊疗项目目录 A, 疾病诊断对照 B, 疾病编码目录 C" & vbNewLine & _
        "Where a.Id = b.手术id(+) And b.疾病id = c.Id(+) And a.ID"
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = strSQL & " IN (Select /*+cardinality(E,10)*/ Column_Value From Table(f_Num2list([1])) E)"
        Set Get诊疗项目记录ID = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strIDs)
    Else
        strSQL = strSQL & " = [1]"
        Set Get诊疗项目记录ID = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FormatAdviceContext() As String
'功能：根据系统基本参数，格式化医嘱内容
'参数：strBloodWay=输血途径,strAdvicePro=输血内容
    Dim strReturn As String, strText As String, strField As String
    Dim str麻醉 As String, str附术 As String
    Dim i As Integer, strTmp As String
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
 
    strReturn = mstrDefine
    
    strTmp = txtInfo(e_麻醉方法).Text
    strTmp = Mid(strTmp, InStr(strTmp, "]") + 1)
    str麻醉 = strTmp
    
    With vsOper
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
            
                strTmp = .TextMatrix(i, COL_名称)
                strTmp = Mid(strTmp, InStr(strTmp, "]") + 1)
                strTmp = strTmp & IIF(Abs(Val(.TextMatrix(i, COL_必要时))) = 1, "(必要时)", "")
                str附术 = str附术 & "," & strTmp
            End If
        Next
        str附术 = Mid(str附术, 2)
    End With
    
    If strReturn = "" Then
        strText = Format(txtInfo(e_手术时间).Text, "MM月dd日HH:mm")
        If str麻醉 <> "" Then
            strText = strText & IIF(str麻醉 <> "", " 在 " & str麻醉 & " 下行 ", " 行 ")
        End If
        strText = strText & txtInfo(e_手术名称).Text
        If str附术 <> "" Then
            strText = strText & " 及 " & str附术
        End If
        strReturn = strText
    Else
        strText = strReturn
        If InStr(strText, "[手术时间]") > 0 Then
            strField = txtInfo(e_手术时间).Text
            strText = Replace(strText, "[手术时间]", """" & strField & """")
        End If
        
        If InStr(strText, "[主要手术]") > 0 Then
            strField = txtInfo(e_手术名称).Text
            strText = Replace(strText, "[主要手术]", """" & strField & """")
        End If
        
        If InStr(strText, "[附加手术]") > 0 Then
            strField = str附术
            strText = Replace(strText, "[附加手术]", """" & strField & """")
        End If
        
        If InStr(strText, "[麻醉方法]") > 0 Then
            strField = str麻醉
            strText = Replace(strText, "[麻醉方法]", """" & strField & """")
        End If
        
        strReturn = mobjVBA.Eval(strText)
    End If

    FormatAdviceContext = strReturn
End Function

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
        MsgBox "输入的医嘱开始执行时间无效。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mint场合 = 0 Then
        strInDate = mstr入院时间
        If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
            strMsg = "医嘱的生效时间不能小于病人的入院时间 " & strInDate & " 。"
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
                    strMsg = "医嘱的生效时间应小于病人" & IIF(mintPState = ps最近转出, "转出", IIF(mintPState = ps预出, "预出院", "出院")) & "的时间 " & strInDate & " 。"
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
                    strMsg = "医嘱的生效时间不能小于病人最近的转科时间 " & strInDate & " 。"
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    Check开始时间 = True
End Function

Private Function Check安排时间(ByVal strDate As String, ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的手术安排时间是否合法  即界面上－－手术时间
'说明：
'1.输血时间不能小于医嘱的开始时间
    Dim strInDate As String, strDateType As String
    
    If Not IsDate(strDate) Then
        strMsg = "输入的手术时间无效。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") < Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = "手术时间不能小于医嘱生效时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check安排时间 = True
End Function

Private Sub txtInfo_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtInfo(Index)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
'按键事件，糊模查找
    If KeyAscii = 13 Then
        Select Case Index
        Case e_手术名称
            Call GetItemOper(0)
        Case e_麻醉方法
            Call GetItem麻醉(0)
        Case e_主刀医生
            Call GetItem主刀医生(0)
        Case e_主刀医生科室
            Call GetItem主刀医生科室(0)
        Case e_第一助手, e_第二助手, e_第三助手
            Call GetItemDoctor(0, Index)
        End Select
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Select Case Index
        Case e_手术名称
            Call GetItemOper(1)
        Case e_麻醉方法
            Call GetItem麻醉(1)
        Case e_主刀医生
            Call GetItem主刀医生(1)
        Case e_主刀医生科室
            Call GetItem主刀医生科室(1)
        Case e_第一助手, e_第二助手, e_第三助手
            Call GetItemDoctor(1, Index)
        End Select

    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
'合法性检查和值的恢复
    If mintType = 2 Then Exit Sub
    
    Select Case Index
    Case e_手术时间
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(e_生效时间).Text) Then
                    '恢复人为的清除缺省为开始时间
                    txtInfo(Index).Text = txtInfo(e_生效时间).Text
                End If
            End If
        Else
            '检查时间合法性
            If Not Check安排时间(txtInfo(Index).Text, txtInfo(e_生效时间).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
        End If
    Case e_生效时间
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
            If mint场合 = 0 Then
                '检查时间合法性
                If Not Check开始时间(txtInfo(Index).Text) Then
                    Cancel = True
                    Call txtInfo_GotFocus(Index)
                    Exit Sub
                End If
                '判断是否是补录医嘱
                If DateDiff("n", CDate(txtInfo(Index).Text), CDate(zlDatabase.Currentdate)) > gint补录间隔 _
                    Or mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
                    mbln补录 = True
                Else
                    mbln补录 = False
                End If
            Else
                mbln补录 = False
            End If
        End If
    
    Case e_手术名称, e_麻醉方法, e_主刀医生, e_主刀医生科室, e_第一助手, e_第二助手, e_第三助手
        If txtInfo(Index).Tag <> txtInfo(Index).Text And txtInfo(Index).Tag <> "" Then txtInfo(Index).Text = txtInfo(Index).Tag
         
    End Select
End Sub

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

Private Sub SetItemEditable(Optional int手术时间 As Integer, Optional int术前诊断 As Integer, _
    Optional int手术名称 As Integer, Optional int执行科室 As Integer, Optional int手术情况 As Integer, _
    Optional int麻醉方法 As Integer, Optional int麻醉执行科室 As Integer, _
    Optional int主刀医生 As Integer, Optional int主刀医生科室 As Integer, _
    Optional int第一助手 As Integer, Optional int第二助手 As Integer, Optional int第三助手 As Integer, _
    Optional int生效时间 As Integer, Optional int其它内容 As Integer, Optional int附加手术 As Integer, Optional int再次手术类型 As Integer)
'功能：设置指定编辑项的可用状态
'参数：0-保持不变,-1-锁定,1-允许
        
    If int手术时间 = 1 Then
        txtInfo(e_手术时间).Locked = False
        txtInfo(e_手术时间).TabStop = True
        txtInfo(e_手术时间).BackColor = vbWindowBackground
        cmdDate(e_手术时间).Enabled = True
    ElseIf int手术时间 = -1 Then
        txtInfo(e_手术时间).Locked = True
        txtInfo(e_手术时间).TabStop = False
        txtInfo(e_手术时间).BackColor = vbButtonFace
        cmdDate(e_手术时间).Enabled = False
    End If
    
    If int术前诊断 = 1 Then
        txtInfo(e_术前诊断).Locked = False
        txtInfo(e_术前诊断).TabStop = True
        txtInfo(e_术前诊断).BackColor = vbWindowBackground
        cmdInfo(e_术前诊断).Enabled = True
    ElseIf int术前诊断 = -1 Then
        txtInfo(e_术前诊断).Locked = True
        txtInfo(e_术前诊断).TabStop = False
        txtInfo(e_术前诊断).BackColor = vbButtonFace
        cmdInfo(e_术前诊断).Enabled = False
    End If
    
    If int手术名称 = 1 Then
        txtInfo(e_手术名称).Locked = False
        txtInfo(e_手术名称).TabStop = True
        txtInfo(e_手术名称).BackColor = vbWindowBackground
        cmdInfo(e_手术名称).Enabled = True
    ElseIf int手术名称 = -1 Then
        txtInfo(e_手术名称).Locked = True
        txtInfo(e_手术名称).TabStop = False
        txtInfo(e_手术名称).BackColor = vbButtonFace
        cmdInfo(e_手术名称).Enabled = False
    End If
    
    If int执行科室 = 1 Then
        cboInfo(e_执行科室).Locked = False
        cboInfo(e_执行科室).TabStop = True
        cboInfo(e_执行科室).BackColor = vbWindowBackground
    ElseIf int执行科室 = -1 Then
        cboInfo(e_执行科室).Locked = True
        cboInfo(e_执行科室).TabStop = False
        cboInfo(e_执行科室).BackColor = vbButtonFace
    End If
    
    If int手术情况 = 1 Then
        cboInfo(e_手术情况).Locked = False
        cboInfo(e_手术情况).TabStop = True
        cboInfo(e_手术情况).BackColor = vbWindowBackground
    ElseIf int手术情况 = -1 Then
        cboInfo(e_手术情况).Locked = True
        cboInfo(e_手术情况).TabStop = False
        cboInfo(e_手术情况).BackColor = vbButtonFace
    End If
    
    If int麻醉方法 = 1 Then
        txtInfo(e_麻醉方法).Locked = False
        txtInfo(e_麻醉方法).TabStop = True
        txtInfo(e_麻醉方法).BackColor = vbWindowBackground
        cmdInfo(e_麻醉方法).Enabled = True
    ElseIf int麻醉方法 = -1 Then
        txtInfo(e_麻醉方法).Locked = True
        txtInfo(e_麻醉方法).TabStop = False
        txtInfo(e_麻醉方法).BackColor = vbButtonFace
        cmdInfo(e_麻醉方法).Enabled = False
    End If
    
    If int麻醉执行科室 = 1 Then
        cboInfo(e_麻醉执行科室).Locked = False
        cboInfo(e_麻醉执行科室).TabStop = True
        cboInfo(e_麻醉执行科室).BackColor = vbWindowBackground
    ElseIf int麻醉执行科室 = -1 Then
        cboInfo(e_麻醉执行科室).Locked = True
        cboInfo(e_麻醉执行科室).TabStop = False
        cboInfo(e_麻醉执行科室).BackColor = vbButtonFace
    End If
    
    If int主刀医生 = 1 Then
        txtInfo(e_主刀医生).Locked = False
        txtInfo(e_主刀医生).TabStop = True
        txtInfo(e_主刀医生).BackColor = vbWindowBackground
        cmdInfo(e_主刀医生).Enabled = True
    ElseIf int主刀医生 = -1 Then
        txtInfo(e_主刀医生).Locked = True
        txtInfo(e_主刀医生).TabStop = False
        txtInfo(e_主刀医生).BackColor = vbButtonFace
        cmdInfo(e_主刀医生).Enabled = False
    End If
    
    If int主刀医生科室 = 1 Then
        txtInfo(e_主刀医生科室).Locked = False
        txtInfo(e_主刀医生科室).TabStop = True
        txtInfo(e_主刀医生科室).BackColor = vbWindowBackground
        cmdInfo(e_主刀医生科室).Enabled = True
    ElseIf int主刀医生科室 = -1 Then
        txtInfo(e_主刀医生科室).Locked = True
        txtInfo(e_主刀医生科室).TabStop = False
        txtInfo(e_主刀医生科室).BackColor = vbButtonFace
        cmdInfo(e_主刀医生科室).Enabled = False
    End If
    
    If int第一助手 = 1 Then
        txtInfo(e_第一助手).Locked = False
        txtInfo(e_第一助手).TabStop = True
        txtInfo(e_第一助手).BackColor = vbWindowBackground
        cmdInfo(e_第一助手).Enabled = True
    ElseIf int第一助手 = -1 Then
        txtInfo(e_第一助手).Locked = True
        txtInfo(e_第一助手).TabStop = False
        txtInfo(e_第一助手).BackColor = vbButtonFace
        cmdInfo(e_第一助手).Enabled = False
    End If
    
    If int第二助手 = 1 Then
        txtInfo(e_第二助手).Locked = False
        txtInfo(e_第二助手).TabStop = True
        txtInfo(e_第二助手).BackColor = vbWindowBackground
        cmdInfo(e_第二助手).Enabled = True
    ElseIf int第二助手 = -1 Then
        txtInfo(e_第二助手).Locked = True
        txtInfo(e_第二助手).TabStop = False
        txtInfo(e_第二助手).BackColor = vbButtonFace
        cmdInfo(e_第二助手).Enabled = False
    End If
    
    If int第三助手 = 1 Then
        txtInfo(e_第三助手).Locked = False
        txtInfo(e_第三助手).TabStop = True
        txtInfo(e_第三助手).BackColor = vbWindowBackground
        cmdInfo(e_第三助手).Enabled = True
    ElseIf int第三助手 = -1 Then
        txtInfo(e_第三助手).Locked = True
        txtInfo(e_第三助手).TabStop = False
        txtInfo(e_第三助手).BackColor = vbButtonFace
        cmdInfo(e_第三助手).Enabled = False
    End If
    
    If int生效时间 = 1 Then
        txtInfo(e_生效时间).Locked = False
        txtInfo(e_生效时间).TabStop = True
        txtInfo(e_生效时间).BackColor = vbWindowBackground
        cmdDate(e_生效时间).Enabled = True
    ElseIf int生效时间 = -1 Then
        txtInfo(e_生效时间).Locked = True
        txtInfo(e_生效时间).TabStop = False
        txtInfo(e_生效时间).BackColor = vbButtonFace
        cmdDate(e_生效时间).Enabled = False
    End If
    
 
    If int其它内容 = 1 Then
        vsOther.TabStop = True
    ElseIf int其它内容 = -1 Then
        vsOther.TabStop = False
        vsOper.Editable = flexEDNone
    End If
    
    If int附加手术 = 1 Then
        vsOper.TabStop = True
    ElseIf int附加手术 = -1 Then
        vsOper.TabStop = False
        vsOper.Editable = flexEDNone
    End If
    
    
    If int再次手术类型 = 1 Then
        cboInfo(e_再次手术类型).Locked = False
        cboInfo(e_再次手术类型).TabStop = True
        cboInfo(e_再次手术类型).BackColor = vbWindowBackground
    ElseIf int再次手术类型 = -1 Then
        cboInfo(e_再次手术类型).Locked = True
        cboInfo(e_再次手术类型).TabStop = False
        cboInfo(e_再次手术类型).BackColor = vbButtonFace
    End If
End Sub

Private Function CheckDoc手术等级() As Boolean
'功能：检查操作员或主刀医生是否达到手术项目的等级
'说明：人员等级和手术等级 同时有时值才进行判断，否则都需要审核
    Dim strSQL As String, rsTmp As Recordset
    Dim strDoc As String
    Dim str手术等级 As String
    Dim str人员等级 As String
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    
    If txtInfo(e_主刀医生).Text = "" Then
        str人员等级 = UserInfo.手术等级
    Else
        strDoc = txtInfo(e_主刀医生).Text
        strSQL = "Select b.手术等级 From 人员表 B Where B.姓名=[1] and b.手术等级 is not null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDoc手术等级", strDoc)
        If Not rsTmp.EOF Then str人员等级 = rsTmp!手术等级 & ""
    End If
    
    If str人员等级 <> "" Then
        strSQL = "select i.手术类型 as 手术等级 from 疾病编码目录 I, 疾病诊断对照 R where r.手术id=[1] and i.id=r.疾病id and i.手术类型 is not null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDoc手术等级", mlng手术项目ID)
        If Not rsTmp.EOF Then str手术等级 = rsTmp!手术等级 & ""
        
        If str人员等级 <> "" And str手术等级 <> "" Then
            If Decode(str人员等级, "丁", 1, "丙", 2, "乙", 3, "甲", 4, "一级", 1, "二级", 2, "三级", 3, "四级", 4, 0) >= _
                Decode(str手术等级, "丁", 1, "丙", 2, "乙", 3, "甲", 4, "一级", 1, "二级", 2, "三级", 3, "四级", 4, 0) Then
                blnTmp = True
            End If
        End If
    End If
    
    CheckDoc手术等级 = blnTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckDocEmpowerEx() As Boolean
'功能：检查操作员是否具有手术项目的执行权
'参数：strAppend=当前申请附项的填写情况串,格式为"项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
    Dim strSQL As String, rsTmp As Recordset
    Dim strDoc As String
    
    On Error GoTo errH
 
    If txtInfo(e_主刀医生).Text = "" Then
        strDoc = UserInfo.姓名
    Else
        strDoc = txtInfo(e_主刀医生).Text
    End If
    strSQL = "Select Count(*) as 权限 From 人员手术权限 A,人员表 B Where A.人员id = B.ID And B.姓名=[1] And A.诊疗项目id = [2] And A.记录性质 = 2"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", strDoc, mlng手术项目ID)
    CheckDocEmpowerEx = Val(rsTmp!权限 & "") > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceCheck(ByVal lng就诊ID As String, ByVal str诊疗类别 As String, ByVal lng项目id As Long, ByVal lng执行科室ID As Long, ByVal lng执行性质 As Long, ByVal strExtra As String) As Boolean
'功能：调用数据库方法zl_AdviceCheck对医嘱项目进行检查
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String
    
    strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as 结果 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", mint调用类型, mlng病人ID, lng就诊ID, mint险类, 1, _
         str诊疗类别, lng项目id, mlng开单科室ID, UserInfo.姓名, lng执行科室ID, lng执行性质, 0, 0, strExtra)
    
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
    AdviceCheck = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsOperateAgain() As Boolean
'功能：是不是再次手术，如果存在有效的手术医嘱则视为再次手术
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    If mstr挂号单 = "" Then
        strSQL = "Select 1 From 病人医嘱记录 Where 病人ID=[1]   And 主页id=[2] And Nvl(婴儿,0)=[3] and 医嘱状态=8 and 诊疗类别='F' and rownum<2"
    Else
        strSQL = "Select 1 From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[4] And Nvl(婴儿,0)=[3] and 医嘱状态=8 and 诊疗类别='F' and rownum<2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mbytBaby, mstr挂号单)
    IsOperateAgain = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
