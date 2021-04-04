VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "ZLIDKIND.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReplenishTheBalanceBill 
   Caption         =   "医保补充结算"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReplenishTheBalanceBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDiagnose 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   1140
      ScaleHeight     =   660
      ScaleWidth      =   6780
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1740
      Width           =   6780
      Begin VSFlex8Ctl.VSFlexGrid vsDiagnose 
         Height          =   600
         Left            =   30
         TabIndex        =   5
         Top             =   75
         Width           =   6555
         _cx             =   11562
         _cy             =   1058
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   350
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSCommLib.MSComm msCommSpeak 
      Left            =   14355
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   1650
      Left            =   -180
      ScaleHeight     =   1650
      ScaleWidth      =   14625
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5430
      Width           =   14625
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "全选(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   15
         TabIndex        =   28
         ToolTipText     =   "热键：Ctrl+A"
         Top             =   1230
         Width           =   1440
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1530
         TabIndex        =   27
         ToolTipText     =   "热键：Ctrl+R"
         Top             =   1230
         Width           =   1440
      End
      Begin VB.TextBox txtYB 
         Height          =   300
         Left            =   795
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txt摘要 
         Height          =   360
         Left            =   990
         MaxLength       =   100
         TabIndex        =   8
         Top             =   90
         Width           =   6960
      End
      Begin VB.TextBox txt退款合计 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   7755
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0.00"
         ToolTipText     =   "连续收费时未缴款单据的实收金额合计"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Frame fraDownSplit 
         Height          =   135
         Left            =   -525
         TabIndex        =   18
         Top             =   945
         Width           =   15075
      End
      Begin VB.CommandButton cmd预结算 
         Caption         =   "预结算(&V)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9840
         TabIndex        =   10
         ToolTipText     =   "热键：F5"
         Top             =   1230
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   13005
         TabIndex        =   12
         ToolTipText     =   "热键:Esc"
         Top             =   1230
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   11430
         TabIndex        =   11
         Top             =   1230
         Width           =   1440
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   375
         Left            =   -15
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   510
         Width           =   11265
         _cx             =   19870
         _cy             =   661
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   12632256
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReplenishTheBalanceBill.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   3
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl摘要 
         AutoSize        =   -1  'True
         Caption         =   "摘要"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   7
         Top             =   150
         Width           =   480
      End
      Begin VB.Label lbl实收 
         AutoSize        =   -1  'True
         Caption         =   "实收:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   13305
         TabIndex        =   25
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbl应收 
         AutoSize        =   -1  'True
         Caption         =   "应收:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11310
         TabIndex        =   24
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbl退款合计 
         AutoSize        =   -1  'True
         Caption         =   "当前退款"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   6450
         TabIndex        =   23
         Top             =   1290
         Width           =   1200
      End
   End
   Begin VB.PictureBox picFeeList 
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   2100
      ScaleHeight     =   1920
      ScaleWidth      =   5055
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3210
      Width           =   5055
      Begin VSFlex8Ctl.VSFlexGrid vsFeeList 
         Height          =   1515
         Left            =   -15
         TabIndex        =   6
         Top             =   -15
         Width           =   5325
         _cx             =   9393
         _cy             =   2672
         Appearance      =   1
         BorderStyle     =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReplenishTheBalanceBill.frx":691D
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   4
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   -2940
      ScaleHeight     =   1035
      ScaleWidth      =   14085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   270
      Width           =   14085
      Begin VB.TextBox txtMCInvoice 
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9570
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   100
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9550
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   100
         Width           =   1545
      End
      Begin VB.ComboBox cboNO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   12040
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   100
         Width           =   1350
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13500
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "热键:F8"
         Top             =   90
         Width           =   400
      End
      Begin VB.ComboBox cboPayMode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11820
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   630
         Width           =   2010
      End
      Begin VB.Frame fraInfo 
         Height          =   135
         Left            =   -150
         TabIndex        =   14
         Top             =   405
         Width           =   13980
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   375
         Left            =   735
         TabIndex        =   1
         Top             =   630
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmReplenishTheBalanceBill.frx":6A0E
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   2
         InputAppearance =   2
         ShowSortName    =   -1  'True
         DefaultCardType =   "0"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "单据号"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   11245
         TabIndex        =   34
         Top             =   160
         Width           =   720
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票号"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   9085
         TabIndex        =   33
         Top             =   160
         Width           =   480
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   9345
         TabIndex        =   21
         Top             =   60
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblPatient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   180
         TabIndex        =   0
         Top             =   690
         Width           =   480
      End
      Begin VB.Label lblPatiInfor 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4290
         TabIndex        =   2
         Top             =   705
         Width           =   480
      End
      Begin VB.Label lblPayMode 
         AutoSize        =   -1  'True
         Caption         =   "原支付方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10350
         TabIndex        =   3
         Top             =   705
         Width           =   1200
      End
      Begin VB.Label lbl险类 
         AutoSize        =   -1  'True
         Caption         =   "险类"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   90
         TabIndex        =   20
         Top             =   120
         Width           =   420
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   8070
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceBill.frx":6AC5
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13123
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   88
            Object.Tag             =   "用于记帐或收费个人帐户显示"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   370
            MinWidth        =   88
            Object.Tag             =   "用于收费预交显示"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "医保大类"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceBill.frx":7359
            Key             =   "Calc"
            Object.ToolTipText     =   "计算器:ALT+?"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmReplenishTheBalanceBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EM_Balance_Type
    EM_Balance_Register = 0 '挂号结算
    EM_Balance_Charge = 1 '收费结算
    EM_Balance_Err_Cancel = 2 'EM_异常作废
    EM_Balance_Err_ReCharge = 3 'EM_异常重新收费
End Enum
'-----------------------------------------------------------
'接口相关变量
Private mlngModule As Long, mstrPrivs As String
Private mEditType As EM_Balance_Type
Private mstrNo As String '当前操作的结算单号
Private mstr结算ID As String '当前操作的结算ID
Private mstr结算序号 As String '当前操作的结算序号
Private mblnFirst As Boolean
Private mblnUnLoad As Boolean
Private mblnElsePersonErrBill As Boolean '是否是他人的异常单据
'-----------------------------------------------------------
'本地相关变量
Private mobjPayCards As Cards
Private mblnNotClearLedDisplay As Boolean   '不清除显示
Private msngMinWidth As Single, msngMinHeight As Single
Private mstrTittle As String
Private mrsList As ADODB.Recordset
Private mblnNotClick As Boolean
Private mstrPreBalance As String '上次选择的支付方式
Private mstrPreDiagnose As String '上次选择的诊断
Private mintInsure As Integer
Private mstrYBPati  As String '医保病人
Private mobjPatiInfor As PatiInfor
Private mrs结算方式 As ADODB.Recordset
Private mstr应付款结算方式 As String
Private mcllDiagnose As Collection  '当前诊断请况

Private mstr个人帐户 As String '是否将个人帐户设置到收费可用
Private Enum Pan
    C2提示信息 = 2
    C3个人帐户 = 3
    C4预交信息 = 4
    C5医保大类 = 5
End Enum
Private mintSucces As Integer '调用成功次数
Private mstrPrePati  As String, mlngPrePati   As Long '上次病人信息
Private mobjInvoice As clsInvoice
Private mobjFactProperty As clsFactProperty
Private mlng领用ID As Long
Private mFrmBalanceWin As frmReplenishTheBalanceWin

Private Type Ty_Module_Para
     int提醒剩余票据张数 As Integer
     bln模糊查找病人 As Boolean
     int模糊天数 As Integer
     bln药房单位 As Boolean
     int清单打印方式 As Integer
     int补结算有效天数 As Integer
     str补结算允许收费方式 As String
End Type
Private mtyMoudlePara As Ty_Module_Para
Private Enum mEmPancelIDX
    EM_Pan_Pati = 1
    EM_Pan_Diagnose = 2
    EM_Pan_FeeList = 3
    EM_Pan_Down = 4
End Enum

Private Enum mEM_Diagnose_SelStatu
    EM_dgGrayToSeled = -1 '将选择的灰色置为选中
    EM_dgClearAllSeled = 0 '清除所有选中的诊断
    EM_dgSelAll = 1 '选择所有的诊断
    EM_dgSeledToGray = 5 '全部置选中的设置为灰色
End Enum
'-----------------------------------------------------------
'医保相关设置
Private Type TY_Insure
    dbl个帐透支 As Double
    dbl帐户余额 As Double
End Type
Private mTy_Insure As TY_Insure
 '当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    医保接口打印票据 As Boolean
    门诊预结算 As Boolean
    分币处理 As Boolean
    实时监控 As Boolean
    先自付 As Boolean
    全自付 As Boolean
    医保不走票号  As Boolean        '预结算时有效
    挂号使用个人帐户 As Boolean
    不收病历费 As Boolean   'support挂号不收取病历费
End Type
Private MCPAR As TYPE_MedicarePAR
Private mcolBalance As Collection '医保结算信息
Private mblnEdit As Boolean  '是否编辑过
Private mblnPrintBill As Boolean '票据是否打印
Private mlng病历费细目ID As Long '病历费对应收费细目ID
Private mcur病历费 As Currency
'-------------------------------------------------------------------------------------
'API声明:
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mrsBalanceNO As ADODB.Recordset '结算序号(No,结算序号,结帐ID)
Private mcllForceDelToCash As Collection '强制退现信息：Array(操作员,卡类别名称,结算方式)
Private mstr排除结算方式 As String '不能使用的结算方式,多个用逗号分隔

Public Function zlEditCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As EM_Balance_Type, Optional ByRef str结算ID As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的父窗口
    '     lngModule-模块号
    '     strPrivs-权限串
    '     EditCard-当前编辑类型
    '     str结算ID-结算ID(异常重收及异常作废时传入)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-16 11:32:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mintSucces = 0: mlngModule = lngModule: mstrPrivs = strPrivs
    mEditType = EditType: mblnFirst = True: mblnUnLoad = False
    mlngModule = 1124
    If CheckDepend = False Then Exit Function
    mstr结算ID = str结算ID
    Set mobjInvoice = New zlPublicExpense.clsInvoice
    If mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then Exit Function
    If CheckDepend = False Then Unload Me: Exit Function
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    
    zlEditCard = mintSucces > 0
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2014-09-16 16:28:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varTemp As Variant
    With mtyMoudlePara
        .bln药房单位 = zlDatabase.GetPara("药品单位显示", glngSys, mlngModule) = "1"
        .int清单打印方式 = Val(zlDatabase.GetPara("收费清单打印方式", glngSys, mlngModule))
        strTemp = zlDatabase.GetPara("姓名模糊查找方式", glngSys, mlngModule)
        varTemp = Split(strTemp & "|", "|")
        .bln模糊查找病人 = Val(varTemp(0)) = "1"
        .int模糊天数 = Val(varTemp(1))
        strTemp = Trim(zlDatabase.GetPara("票据剩余X张时开始提醒收费员", glngSys, mlngModule, "0|10"))
        varTemp = Split(strTemp & "|", "|")
        If Val(varTemp(0)) = 0 Then
            .int提醒剩余票据张数 = -1
        Else
            .int提醒剩余票据张数 = Val(varTemp(1))
        End If
        '84929
        .int补结算有效天数 = Val(zlDatabase.GetPara("补结算有效天数", glngSys, mlngModule, "3"))
        .str补结算允许收费方式 = zlDatabase.GetPara("允许补结算的收费结算方式", glngSys, mlngModule)
    End With
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面相关信息及相关变量
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-10 11:25:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strIDKindStr As String
    
    mstrTittle = "费用补充记录"
    Select Case mEditType
    Case EM_Balance_Charge
        mstrTittle = mstrTittle & "(收费补充结算)"
    Case EM_Balance_Err_Cancel
        mstrTittle = mstrTittle & "(异常结算作废)"
        cmdOK.Caption = "作废(&O)"
    Case EM_Balance_Err_ReCharge
        mstrTittle = mstrTittle & "(异常结算重收)"
        cmdOK.Caption = "重收(&O)"
    Case EM_Balance_Register
        mstrTittle = mstrTittle & "(挂号补充结算)"
    Case Else
        mstrTittle = mstrTittle & "(收费补充结算)"
    End Select
    Me.Caption = mstrTittle
    
    '是否可进行结算退费
    cmdDelete.Visible = (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) _
        And zlStr.IsHavePrivs(mstrPrivs, "结算退费")
    
    If gblnLED Then
        zl9LedVoice.Reset msCommSpeak
        zl9LedVoice.Init UserInfo.编号 & " 收费员为您服务", mlngModule, gcnOracle
    End If
    
    Call InitModulePara
    
    '获取病历费的收费细目ID,84965
    Dim rsRecord As ADODB.Recordset
    Set rsRecord = zlGetSpecialItemFee("病历费")
    If Not rsRecord Is Nothing Then
        If Not rsRecord.EOF Then mlng病历费细目ID = Val(Nvl(rsRecord!收费细目ID))
    End If
    
    Set mobjFactProperty = New clsFactProperty
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_收费收据, 0, 0, 0, mobjFactProperty)
    
    strIDKindStr = "姓|姓名或就诊卡;医|医保号;身|身份证号;IC|IC卡号|1;门|门诊号;单|收费单据号"
    msngMinWidth = (800 * Screen.TwipsPerPixelX) * 0.5
    msngMinHeight = (600 * Screen.TwipsPerPixelY) * 0.5
    mstrTittle = "费用补充记录"
    lbl险类.Caption = ""
    
    Call SetFeeListHead(True)   '初始化费用列头
    With vsDiagnose
        .Clear 1
        .Rows = 1: .COLS = 1
    End With
    Dim blnVisible As Boolean
    blnVisible = mEditType = EM_Balance_Register Or mEditType = EM_Balance_Charge
    cboPayMode.Visible = blnVisible: lblPayMode.Visible = blnVisible
    txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnStrictCtrl '89302
     
    Call InitPancel
    Call ClearData
    '初始化身份认别控件
    Call PatiIdentify.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, _
        gobjSquare.objSquareCard, strIDKindStr, gstrSysName)
End Sub

Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除数据
    '编制:刘兴洪
    '日期:2014-09-22 17:24:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PatiIdentify.Text = ""
    lblPatiInfor.Caption = ""
    Set mobjPatiInfor = Nothing
    cboPayMode.Clear
    
    vsBalance.Clear 1
    vsBalance.Rows = 1
    vsBalance.COLS = 1
    Set mcolBalance = New Collection
    
    vsFeeList.Clear 1
    vsFeeList.Rows = 2
    vsDiagnose.Clear 1
    vsDiagnose.Rows = 1
    vsDiagnose.COLS = 1
    lbl实收.Caption = "实收:0.00"
    lbl应收.Caption = "应收:0.00"
    staThis.Panels(Pan.C3个人帐户).Text = ""
    staThis.Panels(Pan.C3个人帐户).Visible = False
    txt摘要 = "": txt退款合计.Text = Format(0, "0.00")
    Call ClearDisplaySHow
    
    mblnEdit = False
    Call SetButtons '设置按钮
    
    mcur病历费 = 0
    mstr排除结算方式 = ""
End Sub

Private Sub cboPayMode_Click()
    If mblnNotClick Then Exit Sub
    If mstrPreBalance = Trim(cboPayMode.Text) Then Exit Sub
    mstrPreBalance = Trim(cboPayMode.Text)
    
    If mrsList Is Nothing Then Exit Sub
    Call LoadFeeData(mrsList)
    Call SetButtons
End Sub

Private Sub cboPayMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then
        Unload Me: Exit Sub
    End If
    If PatiIdentify.Locked Then
       SetPatientEnableModi True
       Call ClearData
       If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
       Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Call FromNosSel("", False, False, True)
    Call SetButtons
    vsDiagnose.Cell(flexcpChecked, 0, 0, vsDiagnose.Rows - 1, vsDiagnose.COLS - 1) = 2
End Sub

Private Sub cmdDelete_Click()
    '弹出退费窗体
    Call frmReplenishTheBalanceDel.zlShowMe(Me, mlngModule, mstrPrivs, EM_RBDTY_退费, "")
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim strNos As String, str结帐IDs As String, str冲销IDs As String
    Dim dtDate As Date, strNo As String, strReclaimInvoice As String
    Dim cur全自付 As Currency, cur先自付 As Currency, cur进入统筹 As Currency
    
    mblnNotClearLedDisplay = True
    strNos = GetSelFeeNos '获取本次结算单据号
    If mEditType = EM_Balance_Err_Cancel Then
        '异常作废
        If CancelBalance = False Then Call SetButtons: Exit Sub
        Unload Me: mintSucces = mintSucces + 1
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    
    If mEditType = EM_Balance_Err_ReCharge Then
        '并发检查
        If zlIsCheckExistErrBill(Val(mstr结算序号), True) = False Then
            MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(Val(mstr结算序号)) Then
            MsgBox "当前单据正在其它补结算窗口中进行处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        '三方卡结算方式有效性检查
        If ThreeBalanceCheck(mobjPayCards, IsRegister(), strNos, mcllForceDelToCash, mstr排除结算方式) = False Then Exit Sub
        
        If CheckFactValied(True, mblnPrintBill) = False Then
            Call SetButtons: mblnNotClearLedDisplay = False
            Exit Sub
        End If
        '异常重收
        dtDate = zlDatabase.Currentdate
        '显示和返回需要回收的发票，若选择取消则结束结算
        If ShowReclaimInvoice(strNos, strReclaimInvoice) = False Then Exit Sub
        Call GetAsyncKeyState(VK_RETURN)
        
        If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
        Set mFrmBalanceWin = New frmReplenishTheBalanceWin
        If mFrmBalanceWin.zlChargeWin(Me, EM_Balance_Err_ReCharge, mlngModule, mstrPrivs, mobjPatiInfor, mobjPayCards, mstrNo, dtDate, mstr结算ID, _
            mstr结算序号, MCPAR.分币处理, strNos, strReclaimInvoice, mcllForceDelToCash, mstr排除结算方式, mblnElsePersonErrBill, _
            IsRegister()) = False Then
            If Not gfrmMain Is Nothing Then
                Call zlExeBalanceWinRefrshData(mstrNo, False, dtDate)
            End If
            Call SetButtons
            mblnNotClearLedDisplay = False
            Exit Sub
        End If
        Call SetButtons
        If Not gfrmMain Is Nothing Then
            Call zlExeBalanceWinRefrshData(mstrNo, True, dtDate)
        End If
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    
    If isValied(strNos, str结帐IDs, str冲销IDs) = False Then
        If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
        Call SetButtons
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    
    '处理医保统筹金额
    If SaveItemYbMoney(mobjPatiInfor.病人ID, strNos, IIf(mEditType = EM_Balance_Register, 4, 1), _
        cur全自付, cur先自付, cur进入统筹) = False Then
        If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
        Call SetButtons
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    '不支持预结算时计算个帐支付金额
    If Not MCPAR.门诊预结算 And mEditType = EM_Balance_Charge Then
        If UpdateBalance(CCur(Val(lbl实收.Caption)), cur进入统筹, cur全自付, cur先自付) = False Then
            If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
            Call SetButtons
            mblnNotClearLedDisplay = False
            Exit Sub
        End If
    End If
    If CheckFactValied(False, mblnPrintBill) = False Then
        mblnNotClearLedDisplay = False
        Call SetButtons: Exit Sub
    End If
    
    dtDate = zlDatabase.Currentdate
    If SaveData(strNos, str结帐IDs, str冲销IDs, dtDate, mblnPrintBill, strNo, cur全自付, cur先自付, cur进入统筹) = False Then
        If Not gfrmMain Is Nothing Then
            Call zlExeBalanceWinRefrshData(strNo, False, dtDate)
        End If
        If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
        Call SetButtons
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    If Not gfrmMain Is Nothing Then
      Call zlExeBalanceWinRefrshData(strNo, True, dtDate)
    End If
    mblnNotClearLedDisplay = False
End Sub

Private Sub PrintBill(ByVal strNo As String, ByVal dtDate As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据打印
    '入参:blnPrintBill-发票是否允许打印
    '编制:刘兴洪
    '日期:2014-09-24 17:33:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotValiedNos As String
    Dim strReclaimInvoice As String '回收的发票号
    Dim blnPrintBillEmpty As Boolean
    Dim blnVirtualPrint As Boolean
    Dim intPrint As Integer
    strNo = IIf(InStr(1, strNo, "'") = 0, "'" & strNo & "'", strNo)
    blnVirtualPrint = MCPAR.医保接口打印票据
    If mblnPrintBill And Not (blnVirtualPrint And mstrYBPati <> "") Then
RePrint:
        strReclaimInvoice = ""
        Call frmReplenishTheBalancePrint.ReportPrint(1, strNo, mintInsure, mobjFactProperty, strReclaimInvoice, mlng领用ID, txtInvoice.Text, dtDate, _
                blnVirtualPrint, , blnPrintBillEmpty)
        If Not (blnVirtualPrint And mstrYBPati <> "") Then
            If mobjFactProperty.严格控制 And blnPrintBillEmpty = False Then
                If zlIsNotSucceedPrintBill(1, strNo, strNotValiedNos) = True Then
                       If MsgBox("单据[" & strNotValiedNos & "]票据打印未成功,是否重新进行票据打印!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                End If
            End If
        End If
    End If
    
    '打印费用清单:固定不分别打印
    If zlStr.IsHavePrivs(mstrPrivs, "门诊结算清单") Then
        intPrint = Val(zlDatabase.GetPara("结算清单打印方式", glngSys, mlngModule, "0"))
        If intPrint = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO=" & strNo, "药品单位=" & IIf(mtyMoudlePara.bln药房单位, 1, 0), 2)
        ElseIf intPrint = 2 Then
            If MsgBox("要打印结算的收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO=" & strNo, "药品单位=" & IIf(mtyMoudlePara.bln药房单位, 1, 0), 2)
            End If
        End If
    End If
End Sub

Private Function SaveData(ByVal strNos As String, ByVal str结帐IDs As String, ByVal str冲销IDs As String, _
     ByVal dtDate As Date, ByVal blnPrintBill As Boolean, ByRef strNo As String, _
     Optional ByRef cur全自付 As Currency, Optional ByRef cur先自付 As Currency, _
     Optional ByRef cur进入统筹 As Currency) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存单据
    '入参:str结帐IDs-返回本次二次结算的费用结帐IDs,多个用逗号分离
    '     str冲销IDs-返回本次二次结算的费用部分的冲销IDs,多个用逗号分离
    '     blnPrintBill-是否打印票据
    '     cur全自付 -全自费金额
    '     cur先自付-先自付金额
    '     cur进入统筹-统筹金额
    '出参:strNO-返回绵结算单号
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-17 11:42:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strAdvance As String, strDate As String, strFactNO As String
    Dim str结算序号 As String, str结帐ID As String, str虚拟结算 As String, str保险结算 As String
    Dim strSQL As String, strReclaimInvoice As String
    Dim cllIDs As Collection, cllPro As Collection
    Dim blnTrans  As Boolean, i As Long
    Dim varData As Variant
    Dim cur个帐 As Currency
    
    On Error GoTo errHandle
    
    If ShowReclaimInvoice(strNos, strReclaimInvoice) = False Then Exit Function
    str结帐ID = zlDatabase.GetNextId("病人结帐记录")
    strFactNO = Trim(txtInvoice.Text)
    strNo = zlDatabase.GetNextNo(13)    '收费单
    str结算序号 = "-" & str结帐ID
    str保险结算 = GetMedicareBalanceStr(cur个帐)
    str虚拟结算 = str保险结算
    strTemp = str结帐IDs & IIf(str冲销IDs <> "", "," & str冲销IDs, "")
    
    Set cllPro = New Collection
    Set cllIDs = New Collection
    If zlCommFun.ActualLen(strTemp) <= 4000 Then
        cllIDs.Add strTemp
    Else
        varData = Split(strTemp, ",")
        strTemp = ""
        For i = 1 To UBound(varData)
            If zlCommFun.ActualLen(strTemp & "," & varData(i)) >= 4000 Then
                strTemp = Mid(strTemp & "," & varData(i), 2)
                cllIDs.Add strTemp
                strTemp = ""
            End If
            strTemp = strTemp & "," & varData(i)
        Next
        If strTemp <> "" Then
            strTemp = Mid(strTemp, 2)
            cllIDs.Add strTemp
        End If
    End If
    strDate = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    For i = 1 To cllIDs.Count
        'Zl_费用补充记录_补结算
        strSQL = "Zl_费用补充记录_补结算("
        '  No_In          In 费用补充记录.No%Type,
        strSQL = strSQL & "'" & strNo & "',"
        '  实际票号_In    In 费用补充记录.实际票号%Type,
        strSQL = strSQL & IIf(blnPrintBill, "'" & strFactNO & "'", "null") & ","
        '  结算id_In      In 费用补充记录.结算id%Type,
        strSQL = strSQL & "" & str结帐ID & ","
        '  结算序号_In    In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "" & str结算序号 & ","
        '  收费结帐ids_In Varchar2,
        strSQL = strSQL & "'" & cllIDs(i) & "',"
        '  医保结算_In    Varchar2,:允许传入多个,格式为:结算方式,结算金额|.."
        strSQL = strSQL & "" & IIf(str保险结算 = "", "NULL", "'" & str保险结算 & "'") & ","
        '  操作员编号_In  In 费用补充记录.操作员编号%Type,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In  In 费用补充记录.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  登记时间_In    In 费用补充记录.登记时间%Type := Null,
        strSQL = strSQL & "to_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  备注_In    In 费用补充记录.备注%Type := Null,
        strSQL = strSQL & "'" & txt摘要.Text & "',"
        '  附加标志_In    In 费用补充记录.备注%Type := Null,
        strSQL = strSQL & "" & IIf(mEditType = EM_Balance_Register, 1, 0) & ","
        '  费用状态_In    In 费用补充记录.费用状态%Type := 0
        strSQL = strSQL & "1)"
        zlAddArray cllPro, strSQL
        str保险结算 = ""
    Next
    '80944,冉俊明,2014-12-18,将票据回收操作放到结算完成后,原因是若结算出现异常,则先不回收票据,到结算成功后再进行回收
'    If strReclaimInvoice <> "无可退票据" Then
'        varData = Split(strNos, ",")
'        For i = 0 To UBound(varData)
'            'Zl_门诊收费记录_Reprint
'            strSQL = "zl_门诊收费记录_RePrint("
'            '  No_In         门诊费用记录.No%Type,
'            strSQL = strSQL & "'" & varData(i) & "',"
'            '  票据号_In     票据使用明细.号码%Type,
'            strSQL = strSQL & "Null,"
'            '  领用id_In     票据使用明细.领用id%Type,
'            strSQL = strSQL & "0,"
'            '  使用人_In     票据使用明细.使用人%Type,
'            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
'            '  使用时间_In   票据使用明细.使用时间%Type,
'            strSQL = strSQL & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),"
'            '  退费_In       Number := 0,
'            strSQL = strSQL & "0,"
'            '  票据张数_In   Number := 0,
'            strSQL = strSQL & "0,"
'            '  收回票据号_In Varchar2 := Null,
'            strSQL = strSQL & "'" & strReclaimInvoice & "',"
'            '  票种_In Number:=1
'            strSQL = strSQL & "" & IIf(mEditType = EM_Balance_Register, 4, 1) & ")"
'            zlAddArray cllPro, strSQL
'        Next
'    End If
    If MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
        '38821
        '票据数据生成(因为不调HIS的打印，医保接口打印，所以先填票据数据)
        strSQL = "Zl_补充结算票据_Insert('" & strNo & "','" & strFactNO & "'," & ZVal(mlng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                  "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),0,1)"
        zlAddArray cllPro, strSQL
    End If
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If mEditType = EM_Balance_Register Then
        'strAdvance:结算模式|挂号费收取方式|挂号单号|补结算标志(1-补结算;0-普通挂号结算)
        strAdvance = "0|0|" & strNos & "|1"
        If Not gclsInsure.RegistSwap(Val(str结帐ID), cur个帐, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Function
        End If
        gcnOracle.CommitTrans
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, True, mintInsure)
    Else
        '调用结算接口
        If zlInsureClinicSwap(strFactNO, str结帐ID, str结算序号, str虚拟结算, cur全自付, cur先自付, _
            cur进入统筹) = False Then Exit Function
    End If
    
    blnTrans = False
    '显示退费结算窗口
    Call GetAsyncKeyState(VK_RETURN)
    If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
    Set mFrmBalanceWin = New frmReplenishTheBalanceWin
    If Not mFrmBalanceWin.zlChargeWin(Me, mEditType, mlngModule, mstrPrivs, mobjPatiInfor, mobjPayCards, strNo, dtDate, str结帐ID, _
        str结算序号, MCPAR.分币处理, strNos, strReclaimInvoice, mcllForceDelToCash, mstr排除结算方式, , mEditType = EM_Balance_Register) Then Exit Function
    If Not gfrmMain Is Nothing Then SaveData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlExeBalanceWinRefrshData(ByVal strNo As String, ByVal blnSaveOK As Boolean, ByVal dtDate As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行结算操作后的刷新操作
    '入参:blnSaveOK-是否保存成功
    '编制:刘兴洪
    '日期:2014-09-26 10:42:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln连续 As Boolean, i As Long, p As Long
    Dim blnGetFact As Boolean
    
    On Error GoTo errHandle
    
    If mEditType = EM_Balance_Err_Cancel Then
        If blnSaveOK Then mintSucces = mintInsure + 1
        Unload Me: Exit Sub
    End If
   
    If mEditType = EM_Balance_Err_ReCharge Then
        If blnSaveOK = False Then Exit Sub
        mintSucces = mintInsure + 1
        '打印单据
        Call PrintBill(strNo, dtDate)
        Unload Me: Exit Sub
    End If
    If blnSaveOK Then
        '加入单据历史记录(所有类型单据)
        cboNO.AddItem strNo, 0
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i '只显示10个
        Next
        mintSucces = mintInsure + 1
        '打印单据
        Call PrintBill(strNo, dtDate)
        Call ReInitPatiInvoice
    End If
    SetPatientEnableModi True
    Call ClearData: Call SetButtons
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlInsureClinicSwap(ByVal strFactNO As String, _
    ByVal str结帐ID As String, ByVal str结算序号 As String, ByVal str预结算信息 As String, _
    Optional ByRef cur全自付 As Currency, Optional ByRef cur先自付 As Currency, Optional ByRef cur进入统筹 As Currency) As Boolean
    '---------------------------------------------------------- -----------------------------------------------------------------------------------
    '功能:医保调用
    '入参:strFactNo-当前发票号
    '     str结帐ID-当前的结帐ID
    '     str结算序号-当前的结算序号
    '     str预结算信息-结算方式|结算金额||....
    '     cur全自付 -全自费金额
    '     cur先自付-先自付金额
    '     cur进入统筹-统筹金额
    '返回:医保调用成功 返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTransMedicare As Boolean
    Dim strAdvance As String, strSQL As String
    
    On Error GoTo errHandle
    If mintInsure = 0 Then Exit Function
    
    If MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
        '不严格控制票据时保存当前票号
        If Not mobjFactProperty.严格控制 = False Then
            zlDatabase.SetPara "当前收费票据号", strFactNO, glngSys, mlngModule
        End If
    End If
    strAdvance = str结算序号
    If Not gclsInsure.ClinicSwap(Val(str结帐ID), _
        GetMedicareBalanceSum(mstr个人帐户), GetMedicareBalanceSum("医保基金"), _
        cur全自付, cur先自付, mintInsure, strAdvance) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    blnTransMedicare = True
    If strAdvance = str结算序号 Then strAdvance = ""
     
    If strAdvance = "" Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
       zlInsureClinicSwap = True: Exit Function
    End If
    
    str预结算信息 = Replace(Replace(str预结算信息, "|", "||"), ",", "|") '转换为分隔符相同的字符串
    If Not zlInsureCheck(str预结算信息, strAdvance) Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
       zlInsureClinicSwap = True: Exit Function
    End If
    
    '需要更正
    'Zl_费用补充结算_Modify
    strSQL = "Zl_费用补充结算_Modify("
    '  操作类型_In   Number,
    '  --   0-普通结算方式:
    '  --     结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    strSQL = strSQL & "" & 2 & ","
    '  结算id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "" & str结帐ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & strAdvance & "')"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    '  卡号_In       病人预交记录.卡号%Type := Null,
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成结算_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mintInsure)
    zlInsureClinicSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, mintInsure)
End Function

Private Function isValied(ByVal strNos As String, ByRef str结帐IDs As String, ByRef str冲销IDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据有的有效性
    '出参:strNOs-本次结算的单据号
    '     str结帐IDs-返回本次二次结算的费用结帐IDs,多个用逗号分离
    '     str冲销IDs-返回本次二次结算的费用部分的冲销IDs,多个用逗号分离
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-17 10:46:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strTittle As String
    Dim int记录性质 As Integer
    
    On Error GoTo errHandle
    If Not CheckTextLength("摘要", txt摘要) Then Exit Function
    strTittle = IIf(mEditType = EM_Balance_Register, "挂号", "收费")
    int记录性质 = IIf(mEditType = EM_Balance_Register, 4, 1)
    If strNos = "" Then
        ShowMsgbox "当前病人没有需要补充结算的" & strTittle & "费用，请选择需要补充结算的" & strTittle & "费用！"
        Exit Function
    End If
    
    If mEditType = EM_Balance_Register Then
        If MCPAR.挂号使用个人帐户 Then
            If mstr个人帐户 = "" Then
                ShowMsgbox "挂号场合未设置个人帐户结算，病人帐户不能支付！"
                Exit Function
            End If
        End If
    End If
    '检查选择单据中是否存在已二次结算了的
    strSQL = _
    " Select 1" & _
    " From 费用补充记录 A," & _
    "      (Select /*+Cardinality(b,10)*/ Distinct 结帐id" & _
    "       From 门诊费用记录 A, Table(f_Str2list([1])) B" & _
    "       Where a.No = b.Column_Value And Mod(记录性质, 10)=[2]) B" & _
    " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And Nvl(费用状态,0) <> 2 And a.附加标志=[3] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos, int记录性质, _
        IIf(mEditType = EM_Balance_Register, 1, 0))
    If Not rsTemp.EOF Then
        ShowMsgbox "被选择单据中存在已补充结算了的数据或补充结算异常数据，不允许再进行补充结算！"
        Exit Function
    End If
    
    strSQL = _
    " Select /*+Cardinality(b,10)*/ a.记录性质, a.结帐ID," & _
    "       Max(Decode(a.记录状态,2,a.结帐ID,0)) As 冲销ID " & _
    " From 门诊费用记录 A, Table(f_Str2list([1])) B" & _
    " Where a.No = b.Column_Value And Mod(a.记录性质,10)=[2]" & _
    " Group By a.记录性质, a.结帐ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos, int记录性质)
    If rsTemp.EOF Then
        ShowMsgbox "当前病人没有需要补充结算的费用，请选择需要补充结算的费用！"
        Exit Function
    End If
    With rsTemp
        str冲销IDs = "": str结帐IDs = ""
        Do While Not .EOF
            If Val(Nvl(rsTemp!结帐ID)) = Val(Nvl(rsTemp!冲销ID)) Then
                str冲销IDs = str冲销IDs & "," & Val(Nvl(rsTemp!冲销ID))
            Else
                str结帐IDs = str结帐IDs & "," & Val(Nvl(rsTemp!结帐ID))
            End If
            .MoveNext
        Loop
        If str结帐IDs <> "" Then str结帐IDs = Mid(str结帐IDs, 2)
        If str冲销IDs <> "" Then str冲销IDs = Mid(str冲销IDs, 2)
    End With
    If str结帐IDs = "" Then
        ShowMsgbox strTittle & "单为:" & strNos & "中未找到原始的" & strTittle & "记录，不允许进行医保补充结算！"
        Exit Function
    End If

    strSQL = _
    " Select 1" & vbNewLine & _
    " From 病人预交记录 A, 结算方式 B" & vbNewLine & _
    " Where a.结算方式 = b.名称(+) And a.结帐id In (Select Column_Value From Table(f_Num2list([1])))" & vbNewLine & _
    "       And Decode(Mod(a.记录性质,10),1,0,Decode(b.性质,3,1,4,1,0)) = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结帐IDs)
    If rsTemp.EOF = False Then
        ShowMsgbox strTittle & "单为:" & strNos & "的" & strTittle & "单据中存在医保结算的数据，不允许进行医保补充结算！"
        Exit Function
    End If
    
    strSQL = _
    " Select 1" & vbNewLine & _
    " From 病人预交记录 A" & vbNewLine & _
    " Where a.结帐id In (Select Column_Value From Table(f_Num2list([1])))" & vbNewLine & _
    "       And Not Exists(Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 = 9)" & vbNewLine & _
    " Having Count(Distinct Decode(Mod(a.记录性质,10),1,'冲预存款',a.结算方式)) > 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结帐IDs)
    If rsTemp.EOF = False Then
        ShowMsgbox strTittle & "单为:" & strNos & "的" & strTittle & "单据中存在两种以上的结算方式，不允许进行医保补充结算！"
        Exit Function
    End If
    
    '三方卡结算方式有效性检查
    If ThreeBalanceCheck(mobjPayCards, mEditType = EM_Balance_Register, strNos, mcllForceDelToCash, mstr排除结算方式) = False Then Exit Function

    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ThreeBalanceCheck(objCards As Cards, ByVal blnIsRegister As Boolean, ByVal strNos As String, _
    ByRef cllForceDelToCash As Collection, ByRef str排除结算方式 As String) As Boolean
    '三方卡结算方式有效性检查
    '入参：
    '   objCards 补结算所有有效的支付方式
    '   blnIsRegister 是否挂号单
    '   strNos 本次选择补充结算的单据号
    '出参：
    '   cllForceDelToCash 强制退现信息：Array(操作员,卡类别名称,结算方式)
    '   str排除结算方式 排除结算方式,多个用逗号分隔
    '返回：检查通过，返回True；否则，返回False
    '105432
    Dim objCard As Card
    Dim cllFeeBalance As New Collection, i As Integer
    Dim blnFind As Boolean, blnQuestion As Boolean
    Dim str操作员 As String, strKey As String
    Dim dblMoney  As Double
    Dim j As Integer, lngCount As Long
    Dim varData As Variant
    Dim rsBalance As ADODB.Recordset
    
    On Error GoTo errHandler
    Set cllForceDelToCash = New Collection
    str排除结算方式 = ""
    Set rsBalance = zlFromIDGetChargeBalance(2, strNos, , , , IIf(blnIsRegister, 4, 1))
    If rsBalance Is Nothing Then ThreeBalanceCheck = True: Exit Function
    
    rsBalance.Filter = "类型=3"
    '去重
    With rsBalance
        Do While Not .EOF
            strKey = "_" & Val(Nvl(!卡类别ID))
            If CollectionExitsValue(cllFeeBalance, strKey) Then
                dblMoney = cllFeeBalance(strKey)(4) + Val(Nvl(!冲预交))
                cllFeeBalance.Remove strKey
            Else
                dblMoney = Val(Nvl(!冲预交))
            End If
            If RoundEx(dblMoney, 6) > 0 Then '全部退完的就不再加入
                'Array(结算方式,卡类别ID,是否退现,卡类别名称,冲预交,是否全退,是否转帐及代扣)
                cllFeeBalance.Add Array(Nvl(!结算方式), Val(Nvl(!卡类别ID)), Val(Nvl(!是否退现)), _
                    Nvl(!卡类别名称), dblMoney, Val(Nvl(!是否全退)), Nvl(!是否转帐及代扣)), strKey
            End If
            .MoveNext
        Loop
    End With
    If cllFeeBalance.Count = 0 Then ThreeBalanceCheck = True: Exit Function
    
    For i = 1 To cllFeeBalance.Count
        blnQuestion = False
        '医疗卡检查
        If objCards Is Nothing Then
            If MsgBox("『" & cllFeeBalance(i)(3) & "』未启用，该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnQuestion = True
        Else
            blnFind = False
            For Each objCard In objCards
                If objCard.接口序号 = cllFeeBalance(i)(1) Then blnFind = True: Exit For
            Next
            If blnFind = False Then
                If MsgBox("『" & cllFeeBalance(i)(3) & "』未启用，该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnQuestion = True
            End If
        End If
        
        If blnQuestion Then
            If cllFeeBalance(i)(2) = 0 Then '强制退现
                If str操作员 = "" Then '多种卡类别时只验证一次
                    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
                        str操作员 = UserInfo.姓名
                    Else
                        str操作员 = zlDatabase.UserIdentifyByUser(Me, "医疗卡『" & cllFeeBalance(i)(3) & "』强制退现，权限验证：", _
                            glngSys, mlngModule, "三方退款强制退现", , True)
                        If str操作员 = "" Then Exit Function
                    End If
                End If
                'Array(操作员,卡类别名称,结算方式)
                cllForceDelToCash.Add Array(str操作员, cllFeeBalance(i)(3), cllFeeBalance(i)(0))
            End If
        ElseIf cllFeeBalance(i)(5) = 1 Then '必须全退
            If cllFeeBalance(i)(2) = 1 Then '允许退现，必须全退
                If cllFeeBalance(i)(6) = 0 Then '不支持转帐及代扣
                    If MsgBox("『" & cllFeeBalance(i)(3) & "』必须全退，因此不能退回原卡。" & _
                        "如果继续操作，那么该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    str排除结算方式 = str排除结算方式 & "," & cllFeeBalance(i)(0)
                End If
            ElseIf cllFeeBalance(i)(6) = 0 Then '不允许退现，必须全退，且不支持转帐及代扣
                If MsgBox("『" & cllFeeBalance(i)(3) & "』必须全退且不能退现，同时也不支持转帐及代扣，因此无法退回原卡。" & _
                    "如果继续操作，那么该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                If str操作员 = "" Then '多种卡类别时只验证一次
                    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
                        str操作员 = UserInfo.姓名
                    Else
                        str操作员 = zlDatabase.UserIdentifyByUser(Me, "『" & cllFeeBalance(i)(3) & "』强制退现，权限验证：", _
                            glngSys, mlngModule, "三方退款强制退现", , True)
                        If str操作员 = "" Then Exit Function
                    End If
                End If
                'Array(操作员,卡类别名称,结算方式)
                cllForceDelToCash.Add Array(str操作员, cllFeeBalance(i)(3), cllFeeBalance(i)(0))
                str排除结算方式 = str排除结算方式 & "," & cllFeeBalance(i)(0)
            End If
        End If
    Next
    If str排除结算方式 <> "" Then str排除结算方式 = Mid(str排除结算方式, 2)
    

    If str排除结算方式 = "" Then ThreeBalanceCheck = True: Exit Function
    '判断是否还有有效的结算方式
    varData = Split(str排除结算方式, ",")
    lngCount = mobjPayCards.Count
    For i = 1 To mobjPayCards.Count
        If mobjPayCards(i).接口序号 <= 0 Or mobjPayCards(i).接口序号 > 0 And mobjPayCards(i).消费卡 Then
            Exit For
        End If
        
        blnFind = False
        For j = 0 To UBound(varData)
            If mobjPayCards(i).结算方式 = varData(j) Then
                lngCount = lngCount - 1: blnFind = True
            End If
        Next
        If blnFind = False Then Exit For
    Next
    If lngCount <= 0 Then
        MsgBox "排除强制退现的结算方式后，已没有可用的结算方式，不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    ThreeBalanceCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdSelAll_Click()
    Call FromNosSel("", True, False, True)
    Call SetButtons
    vsDiagnose.Cell(flexcpChecked, 0, 0, vsDiagnose.Rows - 1, vsDiagnose.COLS - 1) = 1
End Sub

Private Sub cmd预结算_Click()
    Dim strNos As String, strNone As String
    Dim strAdvance As String
    
    If mintInsure = 0 Then
        MsgBox "未进行医保身份验证,不允许预结算!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    strNos = GetSelFeeNos   '当前选中的单据
    If strNos = "" Then
        MsgBox "未选中需要预结的费用单据,不允许预结算", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If MCPAR.实时监控 Then
        '1.导入单据，2.修改单据，3.输入中药配方，4.修改中药付数后，其它行的付数同时变化，5.输入主项，自动产生从项，以及从项汇总计算折扣
        '6.修改单价，7.调整执行科室，药品价格重算，8.调整费别，实收金额重算,9.先输费用再验证医保身份,其它等等
        If gclsInsure.CheckItem(mintInsure, 0, 9, MakeDetailRecord(strNos), strAdvance) = False Then Exit Sub
    End If
    
    cmd预结算.Enabled = False
    '预结算
    If Not 门诊预结算(strNos, strNone) Then
        If strNone <> "" Then
            MsgBox "当前保险结算使用的结算方式" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                "在门诊未设置，请先到结算方式管理中设置这些结算方式！", vbInformation, gstrSysName
        End If
        cmd预结算.TabStop = True: cmdOK.Enabled = False: cmd预结算.Enabled = True
        If cmd预结算.Enabled And cmd预结算.Visible Then cmd预结算.SetFocus
        mblnEdit = True
        Exit Sub
    End If
    mblnEdit = False
    Call SetButtons
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
End Sub

Private Function 门诊预结算(ByVal strNos As String, ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊预结算
    '入参:strNos-本次预结算的单据号
    '出参:strNone-返回不存在的结算方式
    '返回:预结算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-16 17:30:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double, dblMoney As Double
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim strDate As String, str结算方式 As String
    Dim dbl合计 As Double
    
    strNone = ""
    
    Screen.MousePointer = 11
    On Error GoTo errH
    '初始化结算结果表格
    Call InitBalanceGrid
    
    '获取结算时间
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If zlInsureClinicPreSwap(strNos, strDate, strNone) = False Then Exit Function
    '要先设置以便其它地方识别
    If cmd预结算.Visible Then
        cmd预结算.TabStop = False
        cmdOK.Enabled = True
    End If
    
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            dblMoney = dblMoney + Val(.TextMatrix(0, i + 1))
        Next
        txt退款合计.Text = Format(dblMoney, "0.00")
    End With
    
    Call zl9InsureLedSpeak
    strNone = Mid(strNone, 2)
    If strNone = "" Then 门诊预结算 = True
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSelFeeNos() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选择的费用单据号
    '返回:多个用逗号分离
    '编制:刘兴洪
    '日期:2014-09-16 17:24:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, i As Long
    
    With vsFeeList
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then
                If Trim(.TextMatrix(i, .ColIndex("NO"))) <> "" _
                    And Abs(Val(.Cell(flexcpChecked, i, .ColIndex("选择")))) = 1 Then
                    strNos = strNos & "," & .TextMatrix(i, .ColIndex("NO"))
                End If
            End If
        Next
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetSelFeeNos = strNos
End Function

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.staThis.Visible Then Bottom = Me.staThis.Height
    staThis.Top = Me.ScaleHeight - Me.staThis.Height
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnNotClearLedDisplay = False
    
    If mblnUnLoad Then Unload Me: Exit Sub
    mblnFirst = False
    Call reSizeWinControl
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
    If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then cmdOK.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '参数：Shift=-1：表示是程序强行在调用
    Select Case KeyCode
        Case vbKeyF1  '帮助
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is PatiIdentify Then
                If mobjPatiInfor Is Nothing Then
                    If MCPatientProcess(mobjPatiInfor) = False Then Exit Sub
                End If
            End If
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF5
            If cmd预结算.Visible And cmd预结算.Enabled Then cmd预结算.SetFocus: cmd预结算_Click
        Case vbKeyF6 '定位到病人输入框
            If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Case vbKeyF8 '退费窗口
            If cmdDelete.Visible Then Call cmdDelete_Click
        Case vbKeyF9 '定位到单据号输入框
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Case vbKeyEscape
            cmdCancel.SetFocus: Call cmdCancel_Click
        Case 191 '"?"计算器
            If Shift = vbAltMask Then
                Call staThis_PanelClick(staThis.Panels("Calc"))
            End If
    End Select
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdSelAll.Visible Then Call cmdSelAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdClear.Visible Then Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Set mcolBalance = New Collection
    Call InitFace
    
    If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then
        mblnUnLoad = Not LoadErrBillData(mobjPatiInfor)
    End If
    Call SetControlEnabled
    
    RestoreWinState Me, App.ProductName, mstrTittle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTittle
    If gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset msCommSpeak
    End If
    PatiIdentify.AllowAutoCommCard = False
    PatiIdentify.AllowAutoICCard = False
    PatiIdentify.AllowAutoIDCard = False
    If Not mcllForceDelToCash Is Nothing Then Set mcllForceDelToCash = Nothing
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim bln收费单 As Boolean, strNo As String
    Dim lng病人ID As Long, str险类名称 As String
    
    strNo = Trim(PatiIdentify.Text)
    If Left(PatiIdentify.Text, 1) = "." Then bln收费单 = True: strNo = Mid(strNo, 2)
    Set mobjPatiInfor = New zlIDKind.PatiInfor
    
    If strNo = "" Then
        If CheckPatiInfor(objCardData) = False Then blnCancel = True: Exit Sub
        Exit Sub
    End If
    
    If objCard.名称 Like "*姓*名*" And Not blnCard And InStr("-*+/.", Left(Trim(PatiIdentify.Text), 1)) = 0 Then
        Dim strPati As String, vRect As RECT, rsTmp As ADODB.Recordset
        If Not gblnSeekName Then
            blnCancel = True: Exit Sub
        Else
             '问题号:50485
            strPati = _
                " Select /*+Rule */distinct 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位,decode(b.卡号,Null,Null,'√') As 是否有医疗卡" & _
                " From 病人信息 A, 病人医疗卡信息 B " & _
                " Where Rownum <101 And a.病人ID=b.病人ID(+) And b.状态(+)=0 And B.卡类别ID(+)=[3]  And A.停用时间 is NULL And A.姓名 Like [1]" & _
                IIf(gintNameDays = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])")
                
            strPati = strPati & " Order by 排序ID,姓名"
                
            vRect = zlControl.GetControlRect(PatiIdentify.hWnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, PatiIdentify.Height, blnCancel, False, True, strNo & "%", gintNameDays, Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, 0)), "bytSize=1")
            If Not rsTmp Is Nothing Then
                If Nvl(rsTmp!ID) = 0 Then '当作新病人
                    blnCancel = True: Exit Sub
                Else '以病人ID读取
                    lng病人ID = Nvl(rsTmp!ID)
                End If
            Else '取消选择
                blnCancel = True: Exit Sub
            End If
        End If
    Else
        If Not bln收费单 Then
            If objCard.接口序号 > 0 Then Exit Sub
            If objCard.名称 <> "收费单据号" Then Exit Sub
        End If
        strNo = zlCommFun.GetFullNO(strNo)
        If GetBillNoFromPati(strNo, lng病人ID) = False Then
            MsgBox "未找到对应的" & IIf(mEditType = EM_Balance_Register, "挂号", "收费") & "单据:" & strNo & ",请检查输入的单据是否正确!", vbInformation + vbOKOnly, gstrSysName
            blnCancel = True: Exit Sub
        End If
        If lng病人ID = 0 Then
            MsgBox "对应的" & IIf(mEditType = EM_Balance_Register, "挂号", "收费") & "单据:" & strNo & "不是建档病人的" & IIf(mEditType = EM_Balance_Register, "挂号", "收费") & "单,不能进行医保补结算!", vbInformation + vbOKOnly, gstrSysName
            blnCancel = True: Exit Sub
        End If
    End If
    Set mobjPatiInfor = Nothing
    If zlGetPati(lng病人ID, objCardData, str险类名称) = False Then blnCancel = True: Exit Sub
    strShowText = objCardData.姓名
    Set mobjPatiInfor = objCardData
    If CheckPatiInfor(objCardData) = False Then blnCancel = True: Exit Sub
    blnFindPatied = True
    If vsFeeList.Enabled And vsFeeList.Visible Then
        vsFeeList.SetFocus
    ElseIf cmd预结算.Enabled And cmd预结算.Visible Then
        cmd预结算.SetFocus
    ElseIf cmdOK.Enabled And cmdOK.Visible Then
        cmd预结算.SetFocus
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    '找到病人时,以病人ID作为判断依据
    If objHisPati Is Nothing Then blnCancel = True: Exit Sub
    
    If objHisPati.病人ID = 0 Then blnCancel = True: Exit Sub
    Set mobjPatiInfor = objHisPati
    PatiIdentify.Text = mobjPatiInfor.姓名
    
    If CheckPatiInfor(objHisPati) = False Then blnCancel = True: Exit Sub
    Set objCardData = mobjPatiInfor
    Call SetButtons
    If vsFeeList.Enabled And vsFeeList.Visible Then
        vsFeeList.SetFocus
    ElseIf cmd预结算.Enabled And cmd预结算.Visible Then
        cmd预结算.SetFocus
    ElseIf cmdOK.Enabled And cmdOK.Visible Then
        cmd预结算.SetFocus
    End If
End Sub

Private Function CheckPatiInfor(ByRef objPatiInfor As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证病人信息
    '入参:objPatiInfor-当前病人信息
    '出参:
    '返回:验证成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-16 15:10:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '先进行医保身份验证
    mblnNotClearLedDisplay = False
    If MCPatientProcess(objPatiInfor) = False Then GoTo GoClear
    '加载费用信息
    If ReadBills(objPatiInfor) = False Then GoTo GoClear
    Call SetButtons '设置按钮
    CheckPatiInfor = True
    Exit Function
    
GoClear:
    SetPatientEnableModi True
    ClearData
End Function

Private Function LoadErrBillData(ByRef objPati As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载单据数
    '入参:strNO-单据号
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-10 12:48:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strWhere1 As String, strWithTable As String
    Dim strSQL As String, j As Long
    Dim strTable As String, strFields As String, str险类名称 As String
    Dim rsTemp As ADODB.Recordset
    Dim dblMoney As Double
    Dim dbl结算金额 As Double
    Dim str结算方式 As String
    Dim strTemp As String
    
    On Error GoTo errHandle
    If Not (mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge) Then Exit Function
    
    strSQL = " Select NO,结算序号,备注 From 费用补充记录  A   Where a.结算ID =[1] and rownum <2"
    If mstr结算ID = "" Then mstr结算ID = "0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr结算ID)
    If rsTemp.EOF Then
        MsgBox "未找到需要" & IIf(mEditType = EM_Balance_Err_ReCharge, "重新结算", "作废") & "的异常作废记录！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    mstrNo = Nvl(rsTemp!NO): mstr结算序号 = "" & Val(Nvl(rsTemp!结算序号))
    cboNO.Text = mstrNo: txt摘要.Text = Nvl(rsTemp!备注)
    cboNO.Locked = True: txt摘要.Enabled = False
    
    strWhere = "": strFields = ",'' as 诊断": strTable = ""
    strTable = "," & _
    "         ( Select distinct 1 as 记录性质, A1.NO, f_List2str(Cast(COLLECT(distinct Q.诊断描述 ) as t_Strlist))  as 诊断" & _
    "           From  (Select distinct NO,医嘱序号 From 门诊费用记录 A,收费单据 N1　where mod(a.记录性质,10)=1 And a.记录状态 in (1,3) ANd A.结帐ID=N1.收费结帐ID) A1, " & _
    "               病人医嘱记录 H,病人诊断医嘱 J,病人诊断记录 Q  " & _
    "           Where   A1.医嘱序号=H.ID and Nvl(H.相关ID,H.ID)=J.医嘱ID and J.诊断ID=Q.ID " & _
    "           Group by  A1.NO ) C " & vbNewLine
    
    strFields = ",Max(C.诊断) as 诊断"
    strWithTable = "" & _
    "   With 收费单据 as ( " & _
    "       Select  Distinct A.收费结帐ID  From 费用补充记录  A   Where a.结算序号 =[1] )"
    strSQL = "" & strWithTable & vbCrLf & _
    "    Select A.记录性质,A.NO,A.记录状态,Nvl(A.价格父号,A.序号) as 序号,A.从属父号,A.开单部门ID,A.执行部门ID,A.收费类别,A.费别,A.收费细目ID," & _
    "          A.费用类型 ,max(A.开单人) as 开单人,A.计算单位,max(A.医嘱序号) as 医嘱序号," & _
    "          Avg(Nvl(A.付数,1)) as 付数,Avg(A.数次) as 数次," & _
    "          Sum(A.标准单价) as 单价, Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "          Max(Decode(a.记录状态, 2, '', a.操作员姓名)) As 操作员姓名, Max(Decode(a.记录状态, 2, To_Date('1900-01-01', 'YYYY-MM-DD'),a.登记时间)) As 登记时间," & _
    "          Max(A.摘要)  as 摘要,A.结帐ID,max(A.病人ID) as 病人ID" & strFields & _
    "   From 门诊费用记录 A,收费单据 b " & strTable & _
    "   Where  A.结帐ID=B.收费结帐ID " & _
    "          And a.记录性质 = c.记录性质(+) And a.No = c.No(+)" & _
    "   Group by  a.No, a.记录性质, a.结帐id, a.记录状态, Nvl(a.价格父号, a.序号), a.从属父号, a.开单部门id, a.执行部门id, a.收费类别, a.费别, a.收费细目id,  a.费用类型, a.计算单位, a.结帐id"
              
    strSQL = _
    " Select Decode(A.记录性质,1,'收费',4,'挂号','收费') as 单据, A.NO,A.序号,A.从属父号,A.费别,a.开单人,A.收费细目ID,C.编码 as 收费类别, " & _
    "       -1 as 选择,C.名称 as 类别,B.编码, " & _
    "       Nvl(M1.名称,B.名称) as 名称,E1.名称 as 商品名 ,B.规格,Max(Nvl(A.费用类型,B.费用类型)) 费用类型," & _
            IIf(mtyMoudlePara.bln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as  单位," & _
    "       Max(A.医嘱序号) as 医嘱序号,sum(A.付数) as 付数," & _
    "       sum(A.数次" & IIf(mtyMoudlePara.bln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") as 数次," & _
    "       Max(A.单价" & IIf(mtyMoudlePara.bln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
    "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "       D.名称 as 执行科室,E.名称 as 开单科室,Max(a.操作员姓名) As 操作员姓名, Max(a.登记时间) As 登记时间, " & _
    "       Max(A.摘要) as 摘要,'' as 结算方式,max(A.诊断) as 诊断,A.记录性质,max(A.病人ID) as 病人ID" & _
    " From (" & strSQL & ") A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E,药品规格 X," & _
    "       收费项目别名 M1,收费项目别名 E1" & _
    " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.收费细目ID=X.药品ID(+)" & _
    "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+) " & _
    "       And A.收费细目ID=M1.收费细目ID(+) And M1.码类(+)=1 And M1.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
    " Group by A.记录性质,A.NO,A.序号,A.从属父号,A.费别,a.开单人,A.收费细目ID,C.编码,C.名称,B.编码,Nvl(M1.名称,B.名称)," & _
    "       E1.名称,B.规格,A.计算单位,D.名称,E.名称,X.药品ID,X." & gstr药房单位 & _
    " Having Sum(A.数次)<>0 " & _
    " Order by 登记时间,NO,序号"
    
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr结算序号)
    If mrsList.RecordCount = 0 Then
        MsgBox "未找到需要" & IIf(mEditType = EM_Balance_Err_ReCharge, "重新结算", "作废") & "的异常作废记录！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If zlGetPati(Val(Nvl(mrsList!病人ID)), mobjPatiInfor, str险类名称) = False Then Exit Function
    
    mintInsure = GetBalanceInsure(mstr结算ID, str险类名称)
    txtYB.Text = mintInsure
    mrsList.Filter = ""
    Call LoadFeeData(mrsList)
    
    lbl险类.Caption = str险类名称
    PatiIdentify.Text = mobjPatiInfor.姓名
    PatiIdentify.PasswordChar = ""
    PatiIdentify.IMEMode = 0
    lblPatiInfor.Caption = "性别:" & mobjPatiInfor.性别
    lblPatiInfor.Caption = lblPatiInfor.Caption & Space(4) & "年龄:" & mobjPatiInfor.年龄
    lblPatiInfor.Caption = lblPatiInfor.Caption & Space(4) & "付款方式:" & mobjPatiInfor.医疗付款方式
    initInsurePara mobjPatiInfor.病人ID
    
    strSQL = _
    " Select Decode(Mod(a.记录性质,10),1,'冲预存款',a.结算方式) As 结算方式, Sum(a.冲预交) As 冲预交" & vbNewLine & _
    " From 病人预交记录 A" & vbNewLine & _
    " Where a.结算序号 = [1]" & vbNewLine & _
    " Group By Decode(Mod(a.记录性质,10),1, '冲预存款',a.结算方式)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr结算序号)

     '根据预结算结果设置结算集
    With vsBalance
        .Clear 1
        .Rows = 1
        .COLS = 1
        .TextMatrix(0, 0) = "医保结算"
        
        Do While Not rsTemp.EOF
            '报销方式;金额;是否允许修改
            str结算方式 = Nvl(rsTemp!结算方式, "未结金额")
            dbl结算金额 = Val(Nvl(rsTemp!冲预交))
            dblMoney = dblMoney + dbl结算金额
            .COLS = .COLS + 2
            .TextMatrix(0, .COLS - 2) = str结算方式
            .TextMatrix(0, .COLS - 1) = Format(dbl结算金额, "0.00")
            .Cell(flexcpData, 0, .COLS - 1) = dbl结算金额
            .ColData(.COLS - 1) = 0 '是否允许修改
            .ColData(.COLS - 2) = 0
            
            '结算方式;原始(最大)金额;可否修改;改后金额
            strTemp = str结算方式
            strTemp = strTemp & ";" & dbl结算金额
            strTemp = strTemp & ";" & 0
            strTemp = strTemp & ";" & dbl结算金额
            mcolBalance.Add strTemp
            rsTemp.MoveNext
        Loop
        .TabStop = False
    End With
    txt退款合计.Text = Format(dblMoney, "0.00")
    Call ReInitPatiInvoice(True, mintInsure, mobjPatiInfor.病人ID)
    LoadErrBillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReadBills(ByRef objPati As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载单据数
    '入参:strNO-单据号
    '     blnFilter-是否按条件筛选
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-10 12:48:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strWhere1 As String, strWithTable As String
    Dim strSQL As String, j As Long
    Dim strTable As String, strFields As String
    Dim strTemp As String
    Dim strBalance As String, blnFind As Boolean
    Dim rsTemp As ADODB.Recordset, varData As Variant, i As Long
    Dim strAllNOs As String
    
    On Error GoTo errHandle
    If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then
       '异常单据的处理
       ReadBills = LoadErrBillData(objPati)
       Exit Function
    End If
    If objPati.病人ID = 0 Then Exit Function
    
    If mEditType = EM_Balance_Register Then
        strWhere = " And a.病人ID=[1] And a.记录性质 =4 And a.记录状态 In(1, 3)"
    Else
        strWhere = " And a.病人ID=[1] And a.记录性质 =1 And a.记录状态 In(1, 3)"
    End If
    
    ' --问题号:79396,挂号单据退了其中一部分（号别或病历费）则不允许补结算
    ' --问题号:112811,"建档病人挂号存为划价单"时不能对该挂号单据进行医保补充结算
    If mEditType = EM_Balance_Register Then
        strWhere = strWhere & _
            " And Not Exists (Select 1 From 门诊费用记录 Where No = a.NO And 记录性质 = 4 And 记录状态 = 2)" & _
            " And Nvl(a.摘要,'-') Not Like '划价:%'"
    End If
    ' --只能对设置的“允许结算方式”进行补结算
    If mtyMoudlePara.str补结算允许收费方式 <> "" Then
        strWhere = strWhere & " And Instr('|'||[3]||'|', '|'||Decode(Mod(b.记录性质,10),1,'冲预存款',b.结算方式)||'|') > 0"
    End If
    ' --不包含已经二次结算的,但包含二次结算作废了的
    ' --不能包含医保结算了的
    ' --不包含消费卡结算的
    ' --排除误差费后只能有一种结算方式
    ' --无剩余项目的单据不进行补结算
    strWhere = strWhere & _
    " And Not Exists(Select 1 From 费用补充记录 Where 收费结帐id = a.结帐id And Nvl(费用状态, 0) <> 2)" & vbNewLine & _
    " And Not Exists(Select 1 From 保险结算记录 Where 性质 = 1 And 记录id = a.结帐id)" & vbNewLine & _
    " And (Mod(b.记录性质, 10) = 1 Or b.结算卡序号 Is Null)" & vbNewLine & _
    " And Exists(Select 1" & vbNewLine & _
    "            From 病人预交记录 F" & vbNewLine & _
    "            Where f.结帐id = b.结帐id" & vbNewLine & _
    "                  And Not Exists(Select 1 From 结算方式 Where 名称 = f.结算方式 And 性质 = 9)" & vbNewLine & _
    "            Having Count(Distinct Decode(Mod(记录性质,10),1,'冲预存款',结算方式)) = 1)" & vbNewLine & _
    " And Exists(Select 1" & vbNewLine & _
    "            From 门诊费用记录" & vbNewLine & _
    "            Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号 And 价格父号 Is Null" & vbNewLine & _
    "            Having Sum(Nvl(付数,1)*数次) <> 0)"

    strWithTable = _
    "With 收费单据 As(" & _
    "    Select a.记录性质, a.No, Max(Decode(Mod(b.记录性质,10),1,'冲预存款',b.结算方式)) As 结算方式" & vbNewLine & _
    "    From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
    "    Where a.结帐id = b.结帐id" & vbNewLine & _
    "          And Decode(a.记录性质,1,a.登记时间,a.发生时间) Between Trunc(Sysdate)-[2] And Trunc(Sysdate)+1-1/24/60/60" & vbNewLine & _
    "          And a.病人id = [1]" & vbNewLine & _
               strWhere & vbNewLine & _
    "    Group By a.记录性质, a.No)"
    
    If mEditType = EM_Balance_Charge Then
       strWithTable = strWithTable & "," & vbNewLine & _
        "医嘱诊断 As (" & _
        "    Select Distinct 1 As 记录性质, A2.No, f_List2str(Cast(Collect(Distinct c.诊断描述) As t_Strlist)) As 诊断" & _
        "    From(Select Distinct 1 As 记录性质, B1.No, B1.医嘱序号" & _
        "         From 门诊费用记录 B1, 收费单据 A1" & _
        "         Where B1.No = A1.No And A1.记录性质 = 1 And B1.记录性质 = 1 And B1.记录状态 In (1, 3)" & _
        "        ) A2, 病人医嘱记录 A, 病人诊断医嘱 B, 病人诊断记录 C" & _
        "    Where A2.记录性质 = 1 And A2.医嘱序号 = a.Id And Nvl(a.相关Id,a.Id) = b.医嘱id And b.诊断id = c.Id" & _
        "    Group By A2.No)"
        
        strTable = ",医嘱诊断 C"
        strFields = ",Max(C.诊断) as 诊断 "
        strWhere1 = " And A.NO=C.NO(+)"
    Else
        strTable = ""
        strFields = ",'' as 诊断"
        strWhere1 = ""
    End If
    
    strSQL = strWithTable & vbNewLine & _
    " Select A.记录性质,A.NO,A.记录状态,Nvl(A.价格父号,A.序号) as 序号,A.从属父号,A.开单部门ID,A.执行部门ID," & _
    "       A.收费类别,A.费别,A.收费细目ID, A.费用类型 ,max(A.开单人) as 开单人,A.计算单位," & _
    "       Max(A.医嘱序号) as 医嘱序号,Avg(Nvl(A.付数,1)) as 付数,Avg(A.数次) as 数次," & _
    "       Sum(A.标准单价) as 单价, Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "       Max(Decode(a.记录状态, 2, '', a.操作员姓名)) As 操作员姓名, " & _
    "       Max(Decode(a.记录状态, 2, To_Date('1900-01-01', 'YYYY-MM-DD'), a.登记时间)) As 登记时间," & _
    "       Max(A.摘要)  as 摘要,A.结帐ID,max(A.病人ID) as 病人ID,max(B.结算方式) as 结算方式" & strFields & _
    " From 门诊费用记录 A,收费单据 B " & strTable & _
    " Where A.记录性质 IN(1,4) And A.记录性质=b.记录性质 And a.No=b.No " & strWhere1 & _
    " Group By a.记录性质, a.记录状态, A.NO,Nvl(A.价格父号,A.序号),A.从属父号,A.开单部门ID,A.执行部门ID," & _
    "       A.收费类别,A.费别,A.收费细目ID,A.费用类型,A.计算单位,A.结帐ID"
    
    strSQL = _
    " Select Decode(A.记录性质,1,'收费',4,'挂号','收费') as 单据, A.NO,A.序号,A.从属父号,A.费别,a.开单人,A.收费细目ID,C.编码 as 收费类别, " & _
    "       " & IIf(mEditType = EM_Balance_Charge, -1, 2) & " as 选择,C.名称 as 类别,B.编码, " & _
    "       Nvl(M1.名称,B.名称) as 名称,E1.名称 as 商品名 ,B.规格,Max(Nvl(A.费用类型,B.费用类型)) 费用类型," & _
            IIf(mtyMoudlePara.bln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as  单位," & _
    "       Max(A.医嘱序号) as 医嘱序号,Avg(Decode(a.记录状态, 1, a.付数, 1)) As 付数, " & _
    "       Sum(Decode(a.记录状态, 1, 1, a.付数) * a.数次" & IIf(mtyMoudlePara.bln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") As 数次," & _
    "       Max(A.单价" & IIf(mtyMoudlePara.bln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
    "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "       D.名称 as 执行科室,E.名称 as 开单科室,Max(a.操作员姓名) As 操作员姓名, Max(a.登记时间) As 登记时间, " & _
    "       Max(A.摘要) as 摘要,max(A.结算方式) as 结算方式,max(A.诊断) as 诊断,A.记录性质,max(A.病人ID) as 病人ID" & _
    " From (" & strSQL & ") A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E,药品规格 X," & _
    "       收费项目别名 M1,收费项目别名 E1" & _
    " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.收费细目ID=X.药品ID(+)" & _
    "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+) " & _
    "       And A.收费细目ID=M1.收费细目ID(+) And M1.码类(+)=1 And M1.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
    " Group by A.记录性质,A.NO,A.序号,A.从属父号,A.费别,a.开单人,A.收费细目ID,C.编码,C.名称,B.编码,Nvl(M1.名称,B.名称)," & _
    "       E1.名称,B.规格,A.计算单位,D.名称,E.名称,X.药品ID,X." & gstr药房单位 & _
    " Having Sum(A.数次)<>0 " & _
    " Order by 登记时间,NO,序号"
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objPati.病人ID, mtyMoudlePara.int补结算有效天数 - 1, _
        mtyMoudlePara.str补结算允许收费方式)
        
    If mrsList.RecordCount = 0 Then
        MsgBox "当前病人未找到需要补结算的费用数据!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set mcllDiagnose = New Collection
    strAllNOs = ""
    With mrsList
        Do While Not .EOF
            If InStr("," & strBalance & ",", "," & Nvl(!结算方式) & ",") = 0 Then
                strBalance = strBalance & "," & Nvl(!结算方式)
            End If
            If InStr("|" & strAllNOs & "|", "|" & Nvl(!记录性质) & "," & Nvl(!NO) & "|") = 0 Then
                strAllNOs = strAllNOs & "|" & Nvl(!记录性质) & "," & Nvl(!NO)
            End If
            
            If mEditType = EM_Balance_Charge Then
                strTemp = ""
                For i = 1 To mcllDiagnose.Count
                    If mcllDiagnose(i)(0) = Nvl(!诊断) Then
                        strTemp = mcllDiagnose(i)(1)
                        mcllDiagnose.Remove i: Exit For
                    End If
                Next
                If InStr("," & strTemp & ",", "," & Nvl(!NO) & ",") = 0 Then strTemp = strTemp & "," & Nvl(!NO)
                If Left(strTemp, 1) = "," Then strTemp = Mid(strTemp, 2)
                mcllDiagnose.Add Array(Nvl(!诊断), strTemp)
            End If
            .MoveNext
        Loop
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    If strAllNOs <> "" Then strAllNOs = Mid(strAllNOs, 2)
    
    '加载原收款方式
    mblnNotClick = True
    cboPayMode.Clear
    strSQL = "Select 编码,名称 From 结算方式 Where Instr(','||[1]||',',','||名称||',')>0 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strBalance)
    With rsTemp
        Do While Not .EOF
            cboPayMode.AddItem Nvl(!名称)
            rsTemp.MoveNext
        Loop
        varData = Split(strBalance, ",")
        For i = 0 To UBound(varData)
            blnFind = False
            For j = 0 To cboPayMode.ListCount - 1
                If varData(i) = cboPayMode.List(j) Then blnFind = True: Exit For
            Next
            If blnFind = False Then
                cboPayMode.AddItem varData(i)
            End If
        Next
    End With
    If cboPayMode.ListCount > 0 Then cboPayMode.ListIndex = 0
    mstrPreBalance = cboPayMode.Text
    mblnNotClick = False
    
    '加载诊断
    With vsDiagnose
        .Clear
        .Rows = 1: .COLS = mcllDiagnose.Count
        If mEditType = EM_Balance_Charge Then
            For i = 1 To mcllDiagnose.Count
               .TextMatrix(0, i - 1) = IIf(mcllDiagnose(i)(0) = "", "无诊断收费", mcllDiagnose(i)(0))
               .Cell(flexcpData, 0, i - 1) = mcllDiagnose(i)(1)
               .Cell(flexcpChecked, 0, i - 1) = 2
            Next
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .COLS - 1
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
    
    Call LoadFeeData(mrsList)
    Call LoadBalanceNO(strAllNOs)
    ReadBills = True
    Exit Function
errHandle:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadBalanceNO(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结算对照
    '入参:strNOs-记录性质,NO|....
    '编制:刘兴洪
    '日期:2014-09-26 16:42:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, strTemp As String
    
    On Error GoTo errHandle
    Set mrsBalanceNO = New ADODB.Recordset
    mrsBalanceNO.Fields.Append "记录性质", adBigInt, , adFldIsNullable
    mrsBalanceNO.Fields.Append "NO", adVarChar, 50, adFldIsNullable
    mrsBalanceNO.Fields.Append "结算序号", adBigInt, , adFldIsNullable
    mrsBalanceNO.Fields.Append "结帐ID", adBigInt, , adFldIsNullable
    mrsBalanceNO.CursorLocation = adUseClient
    mrsBalanceNO.LockType = adLockOptimistic
    mrsBalanceNO.CursorType = adOpenStatic
    mrsBalanceNO.Open
    
    If strNos = "" Then Exit Function
    If zlCommFun.ActualLen(strNos) < 4000 Then
        If ReadBalanceData(mrsBalanceNO, strNos) = False Then Exit Function
        LoadBalanceNO = True
        Exit Function
    End If
    
    varData = Split(strNos, "|")
    strTemp = ""
    For i = 0 To UBound(varData)
        If zlCommFun.ActualLen(strTemp & "|" & varData(i)) >= 4000 Then
            If ReadBalanceData(mrsBalanceNO, Mid(strTemp, 2)) = False Then Exit Function
            strTemp = ""
        End If
        strTemp = strTemp & "|" & varData(i)
    Next
    If strTemp <> "" Then
        strTemp = Mid(strTemp, 2)
        If ReadBalanceData(mrsBalanceNO, strTemp) = False Then Exit Function
    End If
    LoadBalanceNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBalanceData(ByRef rsBalanceNO As ADODB.Recordset, ByVal strNos As String) As Boolean
    '向结算序号(No,结算序号,结帐ID)记录集中加入数据
    '入参：
    '   strNos - 记录性质,NO|....
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strFilter As String
    
    On Error GoTo errHandle
    strSQL = _
    " Select Distinct Mod(A.记录性质,10) As 记录性质, A.NO, A.结帐ID, Nvl(B.结算序号,0) As 结算序号" & _
    " From 门诊费用记录 A,病人预交记录 B  " & _
    " Where a.结帐ID=b.结帐ID And ( A.记录性质,A.NO) IN (Select C1,C2 From Table(f_Str2list2([1], '|', ',')))"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    With rsTemp
        Do While Not .EOF
            strFilter = "NO='" & Nvl(!NO) & "'"
            strFilter = strFilter & " And 记录性质= " & Val(Nvl(!记录性质))
            strFilter = strFilter & " And 结帐ID= " & Val(Nvl(!结帐ID))
            strFilter = strFilter & " And 结算序号= " & Val(Nvl(!结算序号))
            rsBalanceNO.Filter = strFilter
            If rsBalanceNO.EOF Then
                rsBalanceNO.Filter = 0
                rsBalanceNO.AddNew
                rsBalanceNO!记录性质 = Val(Nvl(!记录性质))
                rsBalanceNO!NO = CStr(Nvl(!NO))
                rsBalanceNO!结帐ID = Val(Nvl(!结帐ID))
                rsBalanceNO!结算序号 = Val(Nvl(!结算序号))
                rsBalanceNO.Update
            End If
            .MoveNext
        Loop
    End With
    ReadBalanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadFeeData(ByVal rsList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费用数据到网格列表中
    '入参:rsList-费用列表
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-11 11:16:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, strDiagnose As String
    Dim strNo As String
    Dim strFilter As String
    Dim i As Long, j As Long
    Dim str诊断 As String, strTemp As String
    
    On Error GoTo errHandle
    If rsList Is Nothing Then Exit Function
    If rsList.State <> 1 Then Exit Function
    
    strBalance = Trim(cboPayMode.Text)
    If mEditType = EM_Balance_Charge Then '收费结算
        For i = 0 To vsDiagnose.Rows - 1
            For j = 0 To vsDiagnose.COLS - 1
                If Abs(Val(vsDiagnose.Cell(flexcpChecked, i, j))) = 1 Then
                    strDiagnose = "'" & vsDiagnose.TextMatrix(i, j) & "'"
                    If strDiagnose = "'无诊断收费'" Then strDiagnose = "null"
                    strFilter = strFilter & "or (诊断=" & strDiagnose & "" & " And 结算方式='" & strBalance & "') "
                End If
            Next
        Next
        
        If strBalance = "" And strFilter = "" Then
            rsList.Filter = ""
        ElseIf strFilter = "" Then
            rsList.Filter = "结算方式='" & strBalance & "'"
        Else
            rsList.Filter = Mid(strFilter, 3)
        End If
    Else '挂号结算
        If strBalance = "" Then
            rsList.Filter = ""
        Else
            rsList.Filter = "结算方式='" & strBalance & "'"
        End If
    End If
    Set vsFeeList.DataSource = rsList
    Call SetFeeListHead
    
    If Not (mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge) Then
         strNo = ""
        If vsFeeList.Rows < 2 Then
            strNo = ""
        ElseIf vsFeeList.ColIndex("类别") >= 0 Then
            strNo = vsFeeList.TextMatrix(1, vsFeeList.ColIndex("类别"))
        End If
        vsFeeList.Editable = IIf(strNo <> "", flexEDKbdMouse, flexEDNone)
    End If
    
    If mrsList.RecordCount <> 0 Then
        mrsList.MoveFirst: strTemp = ""
        Do While Not mrsList.EOF
            str诊断 = IIf(Nvl(mrsList!诊断) = "", "无诊断收费", Nvl(mrsList!诊断))
            If InStr(strTemp & ",", "," & str诊断 & ",") = 0 Then
                For i = 0 To vsDiagnose.COLS - 1
                    If vsDiagnose.TextMatrix(0, i) = str诊断 And Abs(vsDiagnose.Cell(flexcpChecked, 0, i)) <> 1 Then
                        vsDiagnose.Cell(flexcpChecked, 0, i) = 1
                    End If
                Next
                strTemp = strTemp & "," & str诊断
            End If
            mrsList.MoveNext
        Loop
    End If
    
    Call CalcTotalMoney
    Call CalcRegisterYBMoney
    mblnEdit = True
    LoadFeeData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetFeeListHead(Optional blnInitHead As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置费用信息列头
    '入参:blnInitHead-是否初始调用
    '编制:刘兴洪
    '日期:2014-09-10 17:01:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strHead As String, i As Long, varData As Variant
    
    On Error GoTo errHandle
    With vsFeeList
        .Redraw = flexRDNone
        
        If blnInitHead Then
            strHead = "单据|NO|序号|从属父号|费别|开单人|收费细目ID|收费类别|选择|类别|编码|名称|商品名|规格|" & _
                      "费用类型|单位|医嘱序号|付数|数次|单价|应收金额|实收金额|执行科室|开单科室|操作员姓名|" & _
                      "登记时间|摘要|结算方式|诊断|记录性质|病人ID"
            
            .Clear
            .Rows = 2
            varData = Split(strHead, "|")
            .COLS = UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, i) = varData(i)
            Next
        ElseIf .Rows <= 1 Then
            .Clear 1
            .Rows = 2
        End If
        
        For i = 0 To .COLS - 1
            .ColKey(i) = UCase(Trim(.TextMatrix(0, i)))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            If .ColKey(i) Like "*ID" Or InStr(",序号,从属父号,医嘱序号,收费类别,记录性质,", "," & .ColKey(i) & ",") > 0 Then
                .ColHidden(i) = True
            ElseIf .ColKey(i) Like "*数*" Or .ColKey(i) Like "*额" Or .ColKey(i) Like "*价" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf InStr(",选择,登记时间,", "," & .ColKey(i) & ",") > 0 Then
                .ColAlignment(i) = flexAlignCenterCenter
            End If
        Next
        
        Select Case gTy_System_Para.byt药品名称显示
        Case 0
            .ColHidden(.ColIndex("名称")) = False
            .ColHidden(.ColIndex("商品名")) = True
        Case 1
            .ColHidden(.ColIndex("名称")) = True
            .ColHidden(.ColIndex("商品名")) = False
        Case 2
            .ColHidden(.ColIndex("名称")) = False
            .ColHidden(.ColIndex("商品名")) = False
        End Select
        
        .HighLight = flexHighlightWithFocus
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .COLS - 1
        zl_vsGrid_Para_Restore mlngModule, vsFeeList, mstrTittle, "费用信息列表", True, False
        
        .RowHeight(0) = 350
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        If .TextMatrix(1, .ColIndex("NO")) <> "" Then Call SplitGroupToFeeList
        
        For i = 0 To .COLS - 1
            If i >= .ColIndex("选择") Then Exit For
            .ColHidden(i) = True
        Next
        .ColHidden(.ColIndex("开单科室")) = True
        If .ColIndex("结算方式") >= 0 Then .ColHidden(.ColIndex("结算方式")) = True
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    vsFeeList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SplitGroupToFeeList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:费用明细数据分组显示
    '编制:刘兴洪
    '日期:2014-09-10 17:12:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String
    Dim bytCheck As Byte
    
    On Error GoTo errHandle
    bytCheck = 1
    
    With vsFeeList
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("NO"), .ColIndex("实收金额"), gstrDec, &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("NO"), .ColIndex("应收金额"), gstrDec, &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("类别")
        .OutlineCol = .ColIndex("类别")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                .RowHeight(i) = 350
                .TextMatrix(i, .ColIndex("NO")) = Trim(.Cell(flexcpTextDisplay, i + 1, .ColIndex("NO")))
                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("NO")) & "(" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("单据")) & ")"
                 strTemp = strTemp & Space(2) & "费别:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("费别"))
                 strTemp = strTemp & Space(2) & "开单部门:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单科室"))
                 strTemp = strTemp & Space(2) & "开单人:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单人"))
                 
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("类别"), i, .ColIndex("类别")) = 1
                 For j = 0 To .COLS - 1
                    If j < .ColIndex("应收金额") Then
                        If j >= .ColIndex("类别") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = True
                        ElseIf j = .ColIndex("选择") Then
                            .Cell(flexcpChecked, i, j) = bytCheck
                            .Cell(flexcpAlignment, i, j, i, j) = 4
                            If mEditType = EM_Balance_Register Then bytCheck = 2
                        End If
                    ElseIf .ColIndex("实收金额") = j Then
                        .Cell(flexcpData, i, j) = Val(.TextMatrix(i, j))
                        .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    ElseIf .ColIndex("应收金额") = j Then
                        .Cell(flexcpData, i, j) = Val(.TextMatrix(i, j))
                        .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("选择")) = ""
                .TextMatrix(i, .ColIndex("单价")) = Format(Val(.TextMatrix(i, .ColIndex("单价"))), gstrFeePrecisionFmt)
                .TextMatrix(i, .ColIndex("数次")) = FormatEx(Val(.TextMatrix(i, .ColIndex("数次"))), 5)
                .Cell(flexcpData, i, .ColIndex("应收金额")) = Val(.TextMatrix(i, .ColIndex("应收金额")))
                .Cell(flexcpData, i, .ColIndex("实收金额")) = Val(.TextMatrix(i, .ColIndex("实收金额")))
                
                .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(.TextMatrix(i, .ColIndex("应收金额"))), gstrDec)
                .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(.TextMatrix(i, .ColIndex("实收金额"))), gstrDec)
                
            End If
        Next
        
        Call .AutoSize(.ColIndex("类别"))
        Call .AutoSize(.ColIndex("单价"))
        
        For j = 0 To .COLS - 1
            If j < .ColIndex("应收金额") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetBillNoFromPati(ByVal strNo As String, ByRef lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号，获取对应的病人ID
    '入参:strNo-单据号
    '出参:lng病人ID-返回病从ID
    '返回:找到对应的单据，返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-10 12:38:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    lng病人ID = 0
    strSQL = "Select 病人ID From 门诊费用记录 Where 记录性质=1 and NO=[1] and 记录状态 in (1,3) and rownum< 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTemp.EOF Then Exit Function
    lng病人ID = Val(Nvl(rsTemp!病人ID))
    GetBillNoFromPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPati(ByVal lng病人ID As String, ByRef objPati As PatiInfor, ByRef str险类名称 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID,重新获取数据
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String
    
    Set objPati = New PatiInfor
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select a.病人id, a. 门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式,p.编码 as 医疗付款方式编码, a. 姓名, a.性别, a. 年龄, a.出生日期, a.出生地点, a.身份证号, a.其他证件, a.身份, " & _
    "        a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.监护人, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, " & _
    "        a.合同单位id, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人, a.担保额, a.担保性质, a.就诊时间, a.就诊状态, a.就诊诊室, a.在院, a.Ic卡号, " & _
    "        a.健康号, a.医保号, a.登记时间, a.停用时间, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, '' as 卡号, 0As 卡状态,'' as 密码, '' as 挂失方式, " & _
    "       sysdate as 挂失时间, 0  as 挂失有效天数,sysdate as 当前时间,C.名称 as 险类名称" & _
    "   From 病人信息 A,医疗付款方式 P,保险类别 C " & _
    "   Where A.险类 = C.序号(+) And a.医疗付款方式=P.名称(+) And 病人ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病人ID)
    If rsTemp.EOF Then Exit Function
    objPati.病人ID = rsTemp!病人ID
    objPati.门诊号 = IIf(Val(Nvl(rsTemp!门诊号)) = 0, "", Nvl(rsTemp!门诊号))
    objPati.姓名 = Nvl(rsTemp!姓名)
    objPati.性别 = Nvl(rsTemp!性别)
    objPati.年龄 = Nvl(rsTemp!年龄)
    objPati.出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
    objPati.出生地址 = Nvl(rsTemp!出生地点)
    objPati.身份证号 = Nvl(rsTemp!身份证号)
    objPati.其他证件 = Nvl(rsTemp!其他证件)
    objPati.职业 = Nvl(rsTemp!职业)
    objPati.民族 = Nvl(rsTemp!民族)
    objPati.国籍 = Nvl(rsTemp!国籍)
    objPati.学历 = Nvl(rsTemp!学历)
    objPati.婚姻状况 = Nvl(rsTemp!婚姻状况)
    objPati.区域 = Nvl(rsTemp!婚姻状况)
    objPati.家庭地址 = Nvl(rsTemp!家庭地址)
    objPati.家庭电话 = Nvl(rsTemp!家庭电话)
    objPati.家庭邮编 = Nvl(rsTemp!家庭地址邮编)
    objPati.监护人 = Nvl(rsTemp!监护人)
    objPati.联系人 = Nvl(rsTemp!联系人姓名)
    objPati.联系人关系 = Nvl(rsTemp!联系人关系)
    objPati.联系人地址 = Nvl(rsTemp!联系人地址)
    objPati.联系人电话 = Nvl(rsTemp!联系人电话)
    objPati.工作单位 = Nvl(rsTemp!工作单位)
    objPati.工作单位电话 = Nvl(rsTemp!单位电话)
    objPati.工作单位邮编 = Nvl(rsTemp!单位邮编)
    objPati.工作单位开户行 = Nvl(rsTemp!单位开户行)
    objPati.工作单位开户行帐户 = Nvl(rsTemp!单位帐号)
    objPati.户口地址 = Nvl(rsTemp!户口地址)
    objPati.户口地址邮编 = Nvl(rsTemp!户口地址邮编)
    objPati.籍贯 = Nvl(rsTemp!籍贯)
    objPati.密码 = Nvl(rsTemp!密码)
    objPati.医疗付款方式编码 = Nvl(rsTemp!医疗付款方式编码)
    objPati.医疗付款方式 = Nvl(rsTemp!医疗付款方式)
    str险类名称 = Nvl(rsTemp!险类名称)
    zlGetPati = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Function PatiErrBillPay(ByVal lng病人ID As Long) As Boolean
    '功能:根据病人,对结算异常单据进行重结
    '入参:lng病人ID-指定的病人ID
    '返回:存在异常单据,并进行重新结算,返回true,否则返回False
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim str操作员姓名 As String, blnDoElsePersonErr As Boolean
    Dim lng结算ID As Long
    Dim blnRegister As Boolean
    Dim editTypeTemp As EM_Balance_Type
    
    mblnElsePersonErrBill = False
    blnRegister = mEditType = EM_Balance_Register
   
    On Error GoTo errHandle
    strSQL = "Select 结算ID, 操作员姓名" & vbNewLine & _
            " From 费用补充记录" & vbNewLine & _
            " Where Nvl(费用状态,0) = 1 And 记录性质 = 1 And 记录状态 = 1" & vbNewLine & _
            "       And Nvl(附加标志,0) = [2] And 病人id =[1] And Rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, IIf(blnRegister, 1, 0))
    If rsTemp.EOF Then Exit Function
    
    lng结算ID = Val(Nvl(rsTemp!结算ID))
    str操作员姓名 = Nvl(rsTemp!操作员姓名)
    
    If str操作员姓名 <> UserInfo.姓名 Then
        '判断是否能够对他人的收费异常单据进行重收
        strSQL = "Select 结算序号" & vbNewLine & _
                " From 病人预交记录 A, 结算方式 B" & vbNewLine & _
                " Where Nvl(a.结算方式, '-') = b.名称 And b.性质 Not In ('3', '4') And a.结帐id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结算ID)
        If rsTemp.EOF Then
            '107905，具有“重结他人异常单据”权限时，可以对只进行了医保结算的他人的异常结算单据进行重结
            blnDoElsePersonErr = zlStr.IsHavePrivs(mstrPrivs, "重结他人异常单据")
        Else
            '存在其他非医保结算方式，其它操作员就不能处理了
            blnDoElsePersonErr = False
        End If
        
        If blnDoElsePersonErr = False Then
            If MsgBox("注意:" & vbCrLf & _
                "       该病人存在异常的" & IIf(blnRegister, "挂号", "费用") & _
                "补充结算单据，操作员[" & str操作员姓名 & "]收取了一部分，" & _
                "注意到操作员[" & str操作员姓名 & "]处对异常单据进行重结！" & vbCrLf & vbCrLf & _
                "       是否继续对该病人进行" & IIf(blnRegister, "挂号", "费用") & "补充结算？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                PatiErrBillPay = True
            End If
            Exit Function
        End If
    End If
    
    If MsgBox("注意:" & vbCrLf & _
            "       该病人存在异常的" & IIf(blnRegister, "挂号", "费用") & "补充结算单据" & _
            IIf(str操作员姓名 <> UserInfo.姓名, "，该单据是操作员[" & str操作员姓名 & "]收取的", "") & _
            "，是否重新对该单据进行重新结算？" & vbCrLf & vbCrLf & _
            "『是』代表重新对异常单据进行重新结算" & vbCrLf & _
            "『否』代表不对异常单据进行处理，继续进行" & IIf(blnRegister, "挂号", "费用") & "补充结算操作", _
            vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Function
    End If
    
    '重新对异常单据进行重结
    mblnElsePersonErrBill = blnDoElsePersonErr
    mEditType = EM_Balance_Err_ReCharge
    mstr结算ID = lng结算ID
    If LoadErrBillData(mobjPatiInfor) = False Then
        PatiErrBillPay = True
        Exit Function
    End If
    Call cmdOK_Click
    
    PatiErrBillPay = True
    mstr结算ID = ""
    mEditType = IIf(blnRegister, EM_Balance_Register, EM_Balance_Charge)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub PatiIdentify_GotFocus()
    zlControl.TxtSelAll PatiIdentify.objTxtInput
    If gblnLED Then zl9LedVoice.Speak "#51" '请问你的姓名
End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(PatiIdentify.Text) = "" Then
        KeyAscii = 0
        Call CheckPatiInfor(mobjPatiInfor)
    End If
End Sub

Private Sub picDiagnose_Resize()
    Err = 0: On Error Resume Next
    With picDiagnose
        vsDiagnose.Left = .ScaleLeft
        vsDiagnose.Top = .ScaleTop
        vsDiagnose.Height = .ScaleHeight
        vsDiagnose.Width = .ScaleWidth
    End With
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        lbl实收.Left = .ScaleWidth - lbl实收.Width - 300
        lbl应收.Left = lbl实收.Left - lbl应收.Width - 400
        txt摘要.Width = lbl应收.Left - txt摘要.Left - 400
        vsBalance.Width = .ScaleWidth - vsBalance.Left - 50
        fraDownSplit.Width = .ScaleWidth + 100
        fraDownSplit.Left = .ScaleLeft
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        cmd预结算.Left = cmdOK.Left - cmd预结算.Width - 50
        
        txt退款合计.Left = IIf(cmd预结算.Visible = False, cmdOK.Left, cmd预结算.Left) - txt退款合计.Width - 200
        lbl退款合计.Left = txt退款合计.Left - lbl退款合计.Width - 20
    End With
End Sub

Private Sub picFeeList_Resize()
    Err = 0: On Error Resume Next
    With picFeeList
        vsFeeList.Left = .ScaleLeft
        vsFeeList.Top = .ScaleTop
        vsFeeList.Height = .ScaleHeight
        vsFeeList.Width = .ScaleWidth
    End With
End Sub
 
Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    With picTop
        fraInfo.Left = .ScaleLeft
        fraInfo.Width = .ScaleWidth
        If cmdDelete.Visible Then
            cmdDelete.Left = .ScaleWidth - .ScaleLeft - cmdDelete.Width - 100
            cboNO.Left = .ScaleWidth - .ScaleLeft - cboNO.Width - 550
        Else
            cboNO.Left = .ScaleWidth - .ScaleLeft - cboNO.Width - 50
        End If
        lblNO.Left = cboNO.Left - lblNO.Width - 20
        txtInvoice.Left = lblNO.Left - txtInvoice.Width * 1.3
        txtMCInvoice.Left = txtInvoice.Left
        lblFact.Left = txtInvoice.Left - lblFact.Width - 20
        If txtInvoice.Visible Then
            lblFormat.Left = lblFact.Left - lblFormat.Width - 50
            lblFormat.Top = lblFact.Top
        Else
            lblFormat.Left = lblPayMode.Left - lblPayMode.Width - 50
            lblFormat.Top = lblPayMode.Top
        End If
        cboPayMode.Left = .ScaleWidth - cboPayMode.Width - 50
        lblPayMode.Left = cboPayMode.Left - lblPayMode.Width - 20
        lbl险类.Left = .ScaleLeft + 100
    End With
End Sub

Private Function MCPatientProcess(ByRef objPatiInfor As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保身份验证
    '编制:刘兴洪
    '日期:2014-09-16 09:59:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnTran As Boolean
    Dim lng病人ID As Long, lng病人IDOut As Long
    Dim lngTemp As Long, str险类名称 As String
    Dim strAdvance As String
    
    On Error GoTo errH
'    PatiIdentify.AllowAutoCommCard = False
'    PatiIdentify.AllowAutoICCard = False
'    PatiIdentify.AllowAutoIDCard = False
    If Not objPatiInfor Is Nothing Then
        lng病人ID = objPatiInfor.病人ID
    Else
        lng病人ID = 0
    End If
    
    If gblnLED Then zl9LedVoice.Speak "#50"
    lng病人IDOut = lng病人ID '避免Identify接口中修改该变量后返回新值
    
    '返回：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,24就诊类型(1=急诊门诊),25开单科室名称
    '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    strAdvance = "2"
    mstrYBPati = gclsInsure.Identify(IIf(mEditType = EM_Balance_Register, 3, 0), lng病人IDOut, mintInsure, strAdvance)
    If mstrYBPati = "" Then GoTo CheckValied:
    
    '获取病人信息
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
            lngTemp = Val(CLng(Split(mstrYBPati, ";")(8)))
            If lng病人ID <> lngTemp And lng病人ID <> 0 Then
                MsgBox "医保验证的病人与之前提取的病人不是同一个病人!", vbInformation, gstrSysName
                staThis.Panels(Pan.C2提示信息) = "医保验证的病人与之前提取的病人不是同一个病人!！"
                Call YBIdentifyCancel
                GoTo CheckValied:
                Exit Function
            End If
        End If
        lng病人ID = lng病人IDOut
    End If
            
    '问题:29283
    '  -- 参数:调用场合-1-挂号;2-收费
    '  --        病人id_In-病人ID(未建档的,传入零)
    '  --        卡号_In: 刷卡卡号;未刷卡时,为空
    '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
    If zlPatiCardCheck(IIf(mEditType = EM_Balance_Register, 1, 2), lng病人ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
        Call YBIdentifyCancel
        GoTo CheckValied: Exit Function
    End If
    
    Call initInsurePara(lng病人ID)    '初始化医保参数
    
    If zlGetPati(lng病人ID, objPatiInfor, str险类名称) = False Then
        Call YBIdentifyCancel
        GoTo CheckValied: Exit Function
    End If
    objPatiInfor.险类 = mintInsure
    txtYB.Text = mintInsure
    PatiIdentify.ForeColor = vbRed
    If objPatiInfor.病人类型 <> "" Then
        Call SetPatiColor(PatiIdentify.objTxtInput, objPatiInfor.病人类型, vbRed)
    End If
    
    PatiIdentify.Text = Split(mstrYBPati, ";")(3)
    PatiIdentify.PasswordChar = ""
    PatiIdentify.IMEMode = 0
    lblPatiInfor.Caption = "性别:" & objPatiInfor.性别
    lblPatiInfor.Caption = lblPatiInfor.Caption & Space(4) & "年龄:" & objPatiInfor.年龄
    lblPatiInfor.Caption = lblPatiInfor.Caption & Space(4) & "付款方式:" & objPatiInfor.医疗付款方式
    lbl险类.Caption = str险类名称
    
    '个人帐户
    Dim cur透支额 As Currency
    cur透支额 = RoundEx(mTy_Insure.dbl个帐透支, 2)
    
    mTy_Insure.dbl帐户余额 = gclsInsure.SelfBalance(lng病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, cur透支额, mintInsure)
    staThis.Panels(Pan.C3个人帐户).Text = "个人帐户余额:" & Format(mTy_Insure.dbl帐户余额, "0.00")
    staThis.Panels(Pan.C3个人帐户).Visible = True
    mTy_Insure.dbl个帐透支 = cur透支额
    
    Call SetButtons '设置按钮
    
   If MCPAR.门诊预结算 And mstr个人帐户 <> "" Then  '只有使用个人帐户才用
        vsBalance.COLS = 3
        vsBalance.TextMatrix(0, 0) = "医保结算"
        vsBalance.TextMatrix(0, 1) = mstr个人帐户
        vsBalance.TextMatrix(0, 2) = "0.00"
        vsBalance.ColData(1) = 0
        vsBalance.ColData(2) = 0
    End If
    
    staThis.Panels(Pan.C2提示信息) = ""
    SetPatientEnableModi (False)
    Call ShowWelcomeByLed
    Call ReInitPatiInvoice(True, mintInsure, lng病人ID)
    
    '根据病人,对异常单据进行重结
    If PatiErrBillPay(lng病人ID) Then
        Call YBIdentifyCancel
        GoTo CheckValied: Exit Function
    End If
    
    MCPatientProcess = True
    
    Exit Function
CheckValied:    '检查失败
    mintInsure = 0: mTy_Insure.dbl帐户余额 = 0: mTy_Insure.dbl个帐透支 = 0
    Set objPatiInfor = Nothing
    staThis.Panels(Pan.C3个人帐户).Text = ""
    staThis.Panels(Pan.C3个人帐户).Visible = False
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
    Call PatiIdentify_GotFocus
'    PatiIdentify.AllowAutoCommCard = True
'    PatiIdentify.AllowAutoICCard = True
'    PatiIdentify.AllowAutoIDCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function YBIdentifyCancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消医保病人身份验证
    '返回:返回假时不退出界面或清除操作
    '编制:刘兴洪
    '日期:2014-09-16 16:07:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    YBIdentifyCancel = True
    If mstrYBPati = "" Or PatiIdentify.Text = "" Then Exit Function
    If UBound(Split(mstrYBPati, ";")) < 8 Then Exit Function
    If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
        lng病人ID = Val(CLng(Split(mstrYBPati, ";")(8)))
    End If
    If lng病人ID = 0 Then Exit Function
    YBIdentifyCancel = gclsInsure.IdentifyCancel(IIf(mEditType = EM_Balance_Register, 3, 0), lng病人ID, mintInsure)
End Function

Public Function zlPatiCardCheck(ByVal byt调用场合 As Byte, lng病人ID As Long, str卡号 As String, byt刷卡方式 As Byte) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查病人刷卡方式
    '入参：byt调用场合: 1-挂号;2-收费
    '         lng病人ID:病人ID(未建档的,传入零)
    '         str卡号;未刷卡时,为空
    '         byt刷卡方式: 1-普能刷卡;2-医保刷卡
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-04-27 16:09:08
    '说明：一汽集团的离休病人，使用的医保卡同时也是就诊卡；医院要求必须以医保方式进行
    '          身份验证挂号、收费，而不能以自费方式直接刷卡进行；因此要求在挂号、收费时，离休病人刷卡后如果不是以医保身份验证方式刷的卡，
    '          而是直接刷的卡，就提示并不允许继续。
    '问题:29283
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = " Select Zl_Paticardcheck([1],[2],[3],[4]) as 提示信息 From Dual "
    ' Zl_Paticardcheck
    '  调用场合_IN NUMBER ,
    '  病人id_In Number,
    '  卡号_In   Varchar2,
    '  刷卡方式_In Number:=1
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病人刷卡方式是否合法", byt调用场合, lng病人ID, str卡号, byt刷卡方式)
    strSQL = Nvl(rsTemp!提示信息)
    If strSQL <> "" Then
        MsgBox strSQL, vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    zlPatiCardCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initInsurePara(ByVal lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2011-08-27 12:25:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure)
    MCPAR.门诊预结算 = gclsInsure.GetCapability(support门诊预算, lng病人ID, mintInsure)
    MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, mintInsure)
    MCPAR.先自付 = gclsInsure.GetCapability(support收费帐户首先自付, lng病人ID, mintInsure)
    MCPAR.全自付 = gclsInsure.GetCapability(support收费帐户全自费, lng病人ID, mintInsure)
    MCPAR.实时监控 = gclsInsure.GetCapability(support实时监控, lng病人ID, mintInsure)
    MCPAR.医保不走票号 = False
    MCPAR.挂号使用个人帐户 = gclsInsure.GetCapability(support挂号使用个人帐户, lng病人ID, mintInsure)
    MCPAR.不收病历费 = gclsInsure.GetCapability(support挂号不收取病历费, lng病人ID, mintInsure)
End Sub

Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在关联数据
    '返回:如果不存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 16:49:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '是否存在误差费的处理
    If IsCheck误差费 = False Then Exit Function
    
    '结算方式检查
    Set mrs结算方式 = Get结算方式("收费")
    If mrs结算方式.RecordCount = 0 Then
        MsgBox "收费场合没有可用的结算方式，请先到结算方式管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    If mstr个人帐户 = "" Then
        mrs结算方式.Filter = "性质=3"
        If Not mrs结算方式.EOF Then mstr个人帐户 = mrs结算方式!名称
    End If
    If mstr应付款结算方式 = "" Then
        mrs结算方式.Filter = "应付款=1"
        If Not mrs结算方式.EOF Then mstr应付款结算方式 = Nvl(mrs结算方式!名称)
    End If
    mrs结算方式.Filter = 0
    
    Set mobjPayCards = GetPayCardsObject
    If mobjPayCards Is Nothing Then Exit Function
    If mobjPayCards.Count = 0 Then Exit Function
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPayCardsObject() As Cards
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取补结算支持的结算类别对象
    '返回:返回Cards对象
    '编制:刘兴洪
    '日期:2015-03-18 09:56:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objCards As Cards, objPayCards As Cards
    Dim rsTemp As ADODB.Recordset
    Dim lngKey As Long, i As Long, blnFind As Boolean
    
    On Error GoTo errHandle
    
    Set objCards = New Cards: Set objPayCards = New Cards
    Set rsTemp = Get结算方式("补结算")
    '83533:李南春,2015/3/25,没有有效的补结算
    If rsTemp.RecordCount = 0 Then
        MsgBox "补结算没有可用的结算方式，请先到『结算方式管理』中设置补结算的应用场合。", vbInformation, gstrSysName
        Exit Function
    End If
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
        '入参:bytType-0-所有医疗卡;
        '             1-启用的医疗卡,
        '             2-所有存在三方账户的三方卡
        '             3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.Count
                If objCards(i).结算方式 = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                '83266:李南春,2015/3/18,医疗卡还需判断是否启用
                If InStr(",1,2,", "," & Val(Nvl(rsTemp!性质)) & ",") > 0 _
                    And Val(Nvl(rsTemp!应付款)) <> 1 Then
                    '不加入医保的结算方式或退支票的
                     Set objCard = New Card
                     objCard.短名 = Mid(Nvl(!名称), 1, 1)
                     objCard.接口编码 = Nvl(!编码)
                     objCard.接口程序名 = ""
                     objCard.接口序号 = -1 * lngKey
                     objCard.结算方式 = Nvl(!名称)
                     objCard.名称 = Nvl(!名称)
                     objCard.启用 = True
                     objCard.缺省标志 = Val(Nvl(rsTemp!缺省)) = 1
                     objCard.支付启用 = True
                     objCard.结算性质 = Val(!性质)
                    objPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
                End If
            End If
            .MoveNext
        Loop
    End With
    '加三方卡
    For Each objCard In objCards
        If objCard.消费卡 = False Then
            rsTemp.Filter = "名称='" & objCard.结算方式 & "'"
            If Not rsTemp.EOF Then
                objPayCards.Add objCard, "K" & lngKey
                lngKey = lngKey + 1
            End If
        End If
    Next
    If objPayCards.Count = 0 Then
        MsgBox "结算卡设置有误,原因可能如下:" & vbCrLf & _
        "未正常启用结算卡,请到『医疗卡类别』和『设备配置』中启用", vbInformation, gstrSysName
    End If
    Set GetPayCardsObject = objPayCards
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsCheck误差费() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查误差费是否正常设置
    '返回:正常设置,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 15:17:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gstr误差费名称 <> "" Then IsCheck误差费 = True: Exit Function
    If Not (mEditType = EM_Balance_Register Or mEditType = EM_Balance_Charge) Then IsCheck误差费 = True: Exit Function
    MsgBox "系统中尚未设置有效的误差处理,请在[结算方式管理]中设置。", vbInformation, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    '选择退费原因
    If KeyCode <> vbKeyReturn Then Exit Sub

    If Trim(txt摘要.Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt摘要.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt摘要, Trim(txt摘要.Text), "常用退费原因", "常用退费原因选择", True, True) = False Then
        If zlCommFun.IsCharChinese(Trim(txt摘要.Text)) = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt摘要_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt摘要
End Sub

Private Sub txt摘要_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt摘要, KeyAscii, m文本式
End Sub

Private Sub txt摘要_LostFocus()
    zlCommFun.OpenIme False
    If zlCommFun.ActualLen(txt摘要.Text) > 100 Then
        MsgBox "摘要最多允许输入100个字符或50个汉字！", vbInformation, gstrSysName
        If txt摘要.Visible And txt摘要.Enabled Then txt摘要.SetFocus
    End If
End Sub

Private Sub txt摘要_Change()
    txt摘要.Tag = ""
End Sub

Private Sub vsBalance_DblClick()
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    With vsBalance
      '不允许修改的医保项目
      If Val(.ColData(.Col)) = 0 Then Exit Sub
      .EditCell
      .EditSelStart = 0
      .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsBalance_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Cancel = True: Exit Sub
    With vsBalance
        If Val(.ColData(Col)) = 0 Then Cancel = True: Exit Sub
        '设置单元格的编辑长度
        .EditMaxLength = 16
    End With
End Sub

Private Sub vsBalance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '医保接口检查
     Dim curOrig As Currency, curTotal As Currency
     Dim i As Integer, strKey As String, str结算方式 As String, varData As Variant
    '数据验证
    With vsBalance
        If Val(.ColData(Col)) = 0 Then Exit Sub
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        .EditText = Format(Val(strKey), "0.00")
        If strKey = "" Then Exit Sub
        
        If Not IsNumeric(strKey) Then
            MsgBox .TextMatrix(.Row, 0) & "输入了非法字符,只能输入数字型！", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If Val(strKey) = 0 Then Exit Sub
        
        str结算方式 = Trim(.TextMatrix(0, .Col - 1))
        If str结算方式 = "" Then Exit Sub
        '结算金额不允许超过返回的原始金额(个人帐户允许透支时再判断)
        curOrig = GetMedicareBalanceSum(str结算方式, True)      '该结算方式所有原始返回金额
        If (str结算方式 <> mstr个人帐户 Or mTy_Insure.dbl个帐透支 = 0) _
            And Val(strKey) > curOrig And Val(strKey) <> 0 And curOrig <> 0 Then
            Cancel = True
            MsgBox "输入的""" & str结算方式 & """结算金额不能超过 " & Format(curOrig, "0.00") & " ！", vbInformation, gstrSysName
            Exit Sub
        End If
            
        '个人帐户检查
        If str结算方式 = mstr个人帐户 Then
            '不允许超过允许透支金额
            If mTy_Insure.dbl帐户余额 - Val(strKey) < -1 * mTy_Insure.dbl个帐透支 Then
                Cancel = True
                MsgBox "帐户余额:" & Format(mTy_Insure.dbl帐户余额, "0.00") & _
                    IIf(mTy_Insure.dbl个帐透支 = 0, "", "(" & "允许透支:" & Format(mTy_Insure.dbl个帐透支, "0.00") & ")") & _
                    "不足要结算的金额。", vbInformation, gstrSysName
                Exit Sub
             End If
        End If
            
        '不允许超出单据剩余可结算金额
        curTotal = RoundEx(Val(lbl实收.Tag), "0.00")
        For i = 1 To mcolBalance.Count
           '结算方式;原始(最大)金额;可否修改;改后金额
            varData = Split(mcolBalance(i) & ";;;;", ";")
            If varData(0) <> str结算方式 Then
                curTotal = curTotal - CCur(varData(3))
            End If
        Next
        If Val(strKey) > curTotal Then
            Cancel = True
            MsgBox "结算金额过大，超过单据允许结算金额:" & Format(curTotal, "0.00") & "。", vbInformation, gstrSysName
            Exit Sub
        End If
        .EditText = FormatEx(Val(strKey), 6)
        Call SetBalanceVal(str结算方式, CCur(Val(strKey)))
    End With
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    With vsBalance
        '不允许修改的医保项目
        If Val(.ColData(Col)) = 0 Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsBalance_GotFocus()
    vsBalance_EnterCell
End Sub

Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    Call VsFlxGridCheckKeyPress(vsBalance, vsBalance.Row, vsBalance.Col, KeyAscii, m金额式)
End Sub

Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Val(vsBalance.ColData(Col)) = 0 Then Exit Sub
    Call VsFlxGridCheckKeyPress(vsBalance, Row, Col, KeyAscii, m金额式)
End Sub

Private Sub vsBalance_EnterCell()
    With vsBalance
        If .Col <= 0 Then Exit Sub
    End With
    
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    With vsBalance
        If .ColData(.Col) = 0 Then
             .FocusRect = flexFocusLight
        Else
             .FocusRect = flexFocusHeavy
        End If
    End With
End Sub

Private Sub SetPatientEnableModi(blnModi As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置病人编辑信息
    '编制:刘兴洪
    '日期:2014-09-16 11:42:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PatiIdentify.Locked = Not blnModi
    If blnModi Then
        PatiIdentify.BackColor = &HFFFFFF
    Else
        PatiIdentify.BackColor = &HE0E0E0
    End If
End Sub

Private Sub ShowWelcomeByLed()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示欢迎信息和病人信息
    '编制:刘兴洪
    '日期:2014-06-06 17:56:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String, lngPatient As Long
    If gblnLED = False Then Exit Sub
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    If gblnLedWelcome Then
        zl9LedVoice.Reset msCommSpeak
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.编号 & " 收费员为您服务", mlngModule, gcnOracle
    End If
    strInfo = Trim(PatiIdentify.Text)
    If Not mobjPatiInfor Is Nothing Then strInfo = strInfo & " " & mobjPatiInfor.性别 & " " & mobjPatiInfor.年龄: lngPatient = mobjPatiInfor.病人ID
    zl9LedVoice.DisplayPatient strInfo, lngPatient
End Sub

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0, Optional ByVal lng病人ID_In As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '入参:blnFact-是否重新取发票号
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng病人ID As Long
    Dim intInsure As Integer, lngCur病人ID As Long, lng主页ID As Long
    
    If Not mobjPatiInfor Is Nothing Then lngCur病人ID = mobjPatiInfor.病人ID
    
    lng病人ID = IIf(lng病人ID_In <> 0, lng病人ID_In, lngCur病人ID)
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    
    If lng病人ID = 0 Then
        '上次病人ID
        If PatiIdentify.Text = mstrPrePati And mlngPrePati <> 0 Then
            lng病人ID = mlngPrePati
        End If
    End If
    If lng病人ID = 0 Then lng病人ID = lngCur病人ID
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_收费收据, lng病人ID, lng主页ID, mintInsure, mobjFactProperty)
    Call ZlShowBillFormat(mlngModule, lblFormat, mobjFactProperty.打印格式)
    If blnFact Then Call RefreshFact
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mobjInvoice.zlGetInvoiceGroupID(mlngModule, UserInfo.姓名, EM_收费收据, mobjFactProperty.使用类别, lng领用ID, mobjFactProperty.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng领用ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng领用ID
        Case 0 '操作失败
        Case -1
            If Trim(mobjFactProperty.使用类别) = "" Then
                MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "你没有自用和共用的『" & mobjFactProperty.使用类别 & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFactProperty.使用类别) = "" Then
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "本地的共用票据的『" & mobjFactProperty.使用类别 & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新收费票据号
    '编制:刘兴洪
    '日期:2014-06-06 14:21:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
    If mobjFactProperty Is Nothing Then Exit Sub
    If mobjFactProperty.打印方式 = 0 And Not MCPAR.医保接口打印票据 Then Exit Sub
    
    If mobjFactProperty.严格控制 Then
        'lblFact.tag主要是检查发票号是否手工输入的.手工输入的,发票号为空,否则是自动产生的发票号
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            
            If zlGetInvoiceGroupUseID(mlng领用ID) = False Then
                txtInvoice.Text = "": txtInvoice.Tag = "": Exit Sub
            End If
            
            '严格：取下一个号码
            If mobjInvoice.zlGetNextBill(mlngModule, mlng领用ID, strFactNO) = False Then strFactNO = ""
            txtInvoice.Text = strFactNO
            
            'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
            '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
            '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
            txtInvoice.Tag = txtInvoice.Text
            lblFact.Tag = txtInvoice.Tag
            If mobjFactProperty.启用使用类别 Then Call zlCheckFactIsEnough
        End If
    Else
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            '松散：取下一个号码
            txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("当前收费票据号", glngSys, mlngModule)))
        End If
        txtInvoice.Tag = txtInvoice.Text
        lblFact.Tag = txtInvoice.Tag
    End If
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Sub zlCheckFactIsEnough(Optional ByVal intInvoicePages As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前票据是否允足
    ' 入参:intInvoicePages-需要的发票张数,如果为0,按系统参数提醒
    '编制:刘兴洪
    '日期:2011-05-10 17:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng剩余数量 As Long, lngNums As Long
    
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    
    '刘兴洪 问题:26948 日期:2009-12-28 17:43:00
    '需要检查剩余数量是否充足:
 
    If intInvoicePages <> 0 Then
        If mobjInvoice.zlCheckInvoiceOverplusEnough(1, intInvoicePages, lng剩余数量, mlng领用ID, mobjFactProperty.使用类别) = False Then
            MsgBox "注意:" & vbCrLf & _
                   "    当前剩余票据不足(" & lng剩余数量 & ") ,当前需要" & intInvoicePages & "张票据,请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    Else
        If mobjInvoice.zlCheckInvoiceOverplusEnough(1, mtyMoudlePara.int提醒剩余票据张数, lng剩余数量, mlng领用ID, mobjFactProperty.使用类别) = False Then
            MsgBox "注意:" & vbCrLf & _
                   "    当前剩余票据(" & lng剩余数量 & ") 小于了报警的张数(" & mtyMoudlePara.int提醒剩余票据张数 & "),请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    End If
End Sub

Public Function MakeDetailRecord(ByVal strNos As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据本次结算内容,构建明细数据集
    '入参:strNOs-单据号,多个用逗号分隔
    '出参:
    '返回:返回数据集,格式:病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室
    '编制:刘兴洪
    '日期:2014-09-16 17:20:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim i As Integer, j As Integer, p As Integer, strSQL As String
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl单价 As Double, cur实收 As Currency
    Dim rsTmp As New ADODB.Recordset, rsPrice As ADODB.Recordset
    
    On Error GoTo errHandle
    
    rsTmp.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    '79420,李南春,2014/11/10:调整记录集字段大小
    rsTmp.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    strSQL = "" & _
    "   Select NO,记录状态,结帐ID,nvl(价格父号,序号) as 序号,病人ID,NULL as 主页ID,收费类别,收费细目ID, " & _
    "           Avg(数次*Nvl(付数,0)) 数量,Sum(标准单价) as 单价,Sum(实收金额) 实收金额,max(a.开单人) as 开单人,max(C.名称) as 开单科室" & _
    "   From 门诊费用记录 A,部门表 C" & _
    "   Where A.NO in (Select Column_Value From Table(f_str2List([1])) ) and A.开单部门ID=C.ID(+) " & _
    "           And mod(A.记录性质,10)=1 And A.记录状态 IN (1,2,3)  " & _
    "   Group By NO,记录状态,nvl(价格父号,序号),收费细目ID,病人ID,收费类别,结帐ID "
    
    Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, "读取重新收费单据", strNos)
    With rsPrice
        For i = 1 To .RecordCount
            rsTmp.Filter = "收费细目ID=" & !收费细目ID
            If rsTmp.RecordCount = 0 Then
                rsTmp.AddNew
                
                rsTmp!病人ID = Nvl(!病人ID, mobjPatiInfor.病人ID)
                rsTmp!主页ID = Nvl(!主页ID, 0)
                rsTmp!收费类别 = !收费类别
                rsTmp!收费细目ID = !收费细目ID
                rsTmp!数量 = !数量
                rsTmp!单价 = !单价
                rsTmp!实收金额 = !实收金额
                rsTmp!开单人 = !开单人
                rsTmp!开单科室 = !开单科室
            Else
                rsTmp!数量 = rsTmp!数量 + !数量
                rsTmp!单价 = (rsTmp!单价 + !单价) / 2
                rsTmp!实收金额 = rsTmp!实收金额 + !实收金额
            End If
            rsTmp.Update
            .MoveNext
        Next
    End With
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitBalanceGrid(Optional blnOnlyClearBalace As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化保险结算表格
    '入参:blnOnlyBalace-仅清除结算算信息
    '编制:刘兴洪
    '日期:2011-11-02 13:53:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    vsBalance.Clear 1
    vsBalance.Rows = 1
    vsBalance.COLS = 1
    
    vsBalance.ColAlignment(0) = 1
'    vsBalance.ColAlignment(1) = 7
    vsBalance.Row = 0
    vsBalance.Col = 0
    
    vsBalance.TabStop = False
    With vsBalance
        .Cell(flexcpFontBold, 0, 0, 0, .COLS - 1) = False
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .COLS - 1) = Me.ForeColor
    End With
    If mEditType = EM_Balance_Charge And mEditType = EM_Balance_Register Then vsBalance.Editable = flexEDKbdMouse
    For i = 0 To vsBalance.COLS - 1
        vsBalance.ColData(i) = 0
    Next
    If blnOnlyClearBalace Then Exit Sub
    '清除结算集内容
    Set mcolBalance = New Collection
End Sub

Private Function zlInsureClinicPreSwap(ByVal strNos As String, ByVal strDate As String, _
    ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用门诊预结算接口
    '入参:strNos-当前选中的单据
    '     strDate-结算时间
    '出参:strNone-不支持的结算方式
    '返回:接口调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-16 17:34:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim varBalance As Variant, varTemp As Variant
    Dim j As Long, strTemp As String
    
    Dim arrPage As Variant, arrBalance() As String, strInvoice As String
    Dim str结算方式 As String, dbl结算金额 As Double, dbl可分配额 As Double
    Dim rsTemp  As ADODB.Recordset, i As Long, k As Long, p As Long
    
    On Error GoTo errHandle
    strInvoice = Trim(txtInvoice.Text)
    
    Set rsTemp = MakePreRecord(strNos, strDate, strInvoice)
    
    strAdvance = "2"
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, mintInsure, strAdvance) Then
        staThis.Panels(Pan.C2提示信息).Text = "单据预结算失败。"
        If mstr个人帐户 <> "" And Not MCPAR.门诊预结算 Then  '只有使用个人帐户才用
            vsBalance.COLS = 3
            vsBalance.TextMatrix(0, 1) = mstr个人帐户
            vsBalance.TextMatrix(0, 2) = "0"
            vsBalance.ColData(1) = 0
            vsBalance.ColData(2) = 0
        End If
        Screen.MousePointer = 0
        Exit Function
    End If
    
    If strAdvance <> "" And strAdvance <> "2" Then '医保票据号
        txtMCInvoice.Text = strAdvance
        txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
        txtMCInvoice.Visible = True
    End If
    
    MCPAR.医保不走票号 = False
    If InStr(1, strAdvance, ";") > 0 Then
          '38821:strAdvance:发票号;是否不走票据号
          MCPAR.医保不走票号 = Val(Split(strAdvance & ";", ";")(1)) = 1
    End If
        

     '根据预结算结果设置结算集
    Set mcolBalance = New Collection
    
    With vsBalance
        .Clear 1
        .Rows = 1
        .COLS = 1
        .TextMatrix(0, 0) = "医保结算"
        
        varBalance = Split(strBalance, "|")
        For i = 0 To UBound(varBalance)
            '报销方式;金额;是否允许修改
            varTemp = Split(varBalance(i) & ";;;;", ";")
            str结算方式 = varTemp(0)
            dbl结算金额 = Val(varTemp(1))
            
            mrs结算方式.Filter = "名称='" & str结算方式 & "' And  性质>=3 and 性质<= 4"
            If mrs结算方式.EOF Then
                '记录医保有但本地没有的结算方式
                If InStr(strNone & ",", "," & str结算方式 & ",") = 0 Then
                    strNone = strNone & "," & str结算方式
                End If
            End If
            If Not mrs结算方式.EOF And dbl结算金额 <> 0 Then
                .COLS = .COLS + 2
                .TextMatrix(0, .COLS - 2) = str结算方式
                .TextMatrix(0, .COLS - 1) = FormatEx(dbl结算金额, 6)
                .Cell(flexcpData, 0, .COLS - 1) = dbl结算金额
                .ColData(.COLS - 1) = Val(varTemp(2)) '是否允许修改
                .ColData(.COLS - 2) = 0
                
                '结算方式;原始(最大)金额;可否修改;改后金额
                strTemp = str结算方式
                strTemp = strTemp & ";" & dbl结算金额
                strTemp = strTemp & ";" & Val(varTemp(2))
                strTemp = strTemp & ";" & GetYBActualMoeny(str结算方式, dbl结算金额)
                mcolBalance.Add strTemp
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        For i = 0 To .COLS - 1
            If .ColData(i) <> 0 Then
                .Row = 0:  .Col = i: .TabStop = True
            End If
            If i > 0 And i Mod 2 = 0 Then vsBalance.ColWidth(i) = 1000
        Next
    End With
    
    zlInsureClinicPreSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function MakePreRecord(ByVal strNos As String, ByVal str结算时间 As String, ByVal strInvoice As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据对象内容创建一个记录信息(以售价单位)
    '入参:strNos-当前的单据信息
    '     str结算时间=结算时间(yyyy-mm-dd HH:MM:SS)
    '     strInvoice=票据号
    '出参:
    '返回:医保相关数据的数据集(单据序号(1--n),病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保)
    '编制:刘兴洪
    '日期:2011-08-15 16:40:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, intStartPage As Integer, intPages As Integer
    Dim p As Integer, strSQL As String
    Dim dbl单价 As Double, cur实收 As Currency, cur统筹 As Currency
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    Dim strAllNOs As String
    
    Err = 0: On Error GoTo Errhand:
    rsTmp.Fields.Append "单据序号", adBigInt, 50, adFldIsNullable
    rsTmp.Fields.Append "费别", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsTmp.Fields.Append "序号", adBigInt, , adFldIsNullable '问题:42961
    rsTmp.Fields.Append "实际票号", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "结算时间", adDBTimeStamp, , adFldIsNullable
    rsTmp.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "收据费目", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "计算单位", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "统筹金额", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "保险支付大类ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "是否医保", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "保险编码", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "摘要", adVarChar, 2000, adFldIsNullable
    rsTmp.Fields.Append "是否急诊", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "开单部门ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    strSQL = _
    "Select '" & strInvoice & "' as 实际票号,NO,记录状态,Nvl( 价格父号, 序号) as 序号,To_Date('" & str结算时间 & "','YYYY-MM-DD HH24:MI:SS') as 结算时间," & _
    "       max(A.病人ID) as 病人ID ,max(A.费别) As 费别,收费类别,收据费目,计算单位,开单人," & _
    "       收费细目ID,保险大类ID As 保险支付大类ID,Nvl(保险项目否,0) As 是否医保,保险编码," & _
    "       Avg(Nvl(付数,0)*数次) As 数量,Avg(标准单价) As 单价," & _
    "       Sum(实收金额) As 实收金额,Sum(统筹金额) As 统筹金额,摘要," & _
    "       max(加班标志) as 是否急诊,开单部门ID,执行部门ID,结帐ID " & _
    "       From 门诊费用记录 a" & _
    "   Where 记录性质=1 And A.NO in (Select Column_Value From Table(f_str2List([1])) ) " & _
    " Group By NO,记录状态,Nvl(价格父号,序号),收费类别,收据费目,计算单位,开单人," & _
    "       收费细目ID,保险大类ID,Nvl(保险项目否,0),保险编码,摘要,开单部门ID,执行部门ID,结帐ID" & _
    " Order by  NO,序号,记录状态 "
    Set rsNo = zlDatabase.OpenSQLRecord(strSQL, "获取划价单数据-医保", strNos)
    If rsNo.RecordCount <> 0 Then rsNo.MoveFirst
    p = 0
    Do While Not rsNo.EOF
        rsTmp.AddNew
        If InStr(strAllNOs & ",", "," & Nvl(rsNo!NO) & ",") = 0 Then p = p + 1
        
        rsTmp!单据序号 = p
        rsTmp!费别 = Nvl(rsNo!费别)
        rsTmp!NO = Nvl(rsNo!NO)   '仅提取划价单时才有值
        rsTmp!序号 = Val(Nvl(rsNo!序号))   '仅提取划价单时才有值
        rsTmp!实际票号 = strInvoice
        rsTmp!结算时间 = CDate(str结算时间)
        rsTmp!病人ID = Nvl(rsNo!病人ID)
        rsTmp!收费类别 = Nvl(rsNo!收费类别)
        rsTmp!收据费目 = Nvl(rsNo!收据费目)
        rsTmp!开单人 = Nvl(rsNo!开单人)
        rsTmp!收费细目ID = Val(Nvl(rsNo!收费细目ID))
        rsTmp!计算单位 = Nvl(rsNo!计算单位)
        rsTmp!数量 = Val(Nvl(rsNo!数量))
        rsTmp!单价 = Val(Nvl(rsNo!单价))
        rsTmp!实收金额 = Val(Nvl(rsNo!实收金额))
        rsTmp!统筹金额 = Val(Nvl(rsNo!统筹金额))
        rsTmp!保险支付大类ID = IIf(Val(Nvl(rsNo!保险支付大类ID)) = 0, Null, Val(Nvl(rsNo!保险支付大类ID)))
        rsTmp!是否医保 = Val(Nvl(rsNo!是否医保))
        rsTmp!保险编码 = Nvl(rsNo!保险编码)
        rsTmp!摘要 = Nvl(rsNo!摘要)
        rsTmp!是否急诊 = Val(Nvl(rsNo!是否急诊))
        rsTmp!开单部门ID = Val(Nvl(rsNo!开单部门ID))
        rsTmp!执行部门ID = Val(Nvl(rsNo!执行部门ID))
        rsTmp.Update
        If InStr(1, strAllNOs & ",", "," & Nvl(rsNo!NO) & ",") = 0 Then
            strAllNOs = strAllNOs & "," & Nvl(rsNo!NO)
        End If
        rsNo.MoveNext
    Loop
                 
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakePreRecord = rsTmp
    Exit Function
Errhand::
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetYBActualMoeny(ByVal str结算方式 As String, ByVal dbl结算金额 As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保结算的实际使用金额
    '入参:str结算方式-医保的结算方式
    '     dbl结算金额-医保的结算金额
    '返回:实际金额,否则返回传入的dbl结算金额
    '编制:刘兴洪
    '日期:2014-09-16 18:05:13
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim dbl个帐 As Double, dbl个帐合计 As Double
    
    On Error GoTo errHandle
    
    If dbl结算金额 = 0 Then Exit Function
    If str结算方式 <> mstr个人帐户 Then GetYBActualMoeny = dbl结算金额: Exit Function
    
    '咸阳医保无法返回余额
     If (mTy_Insure.dbl帐户余额 > -1 * mTy_Insure.dbl个帐透支 Or mintInsure = 61) _
        And CCur(lbl实收.Tag) > 0 Then
        dbl个帐 = dbl结算金额
        If mintInsure <> 61 Then
            '计算个人帐户支付金额
            If RoundEx(mTy_Insure.dbl帐户余额 - dbl个帐合计 - dbl个帐, 6) = -1 * mTy_Insure.dbl个帐透支 Then
                dbl个帐 = dbl个帐 '在允许透支范围内足够(允许透支0为特例)
            Else
                If mTy_Insure.dbl个帐透支 = 0 And RoundEx(mTy_Insure.dbl帐户余额 - dbl个帐合计, 6) > 0 Then
                    dbl个帐 = mTy_Insure.dbl帐户余额 - dbl个帐合计 '不允许透支且有余额
                Else
                    '超过允许透支范围或不允许透支时无余额
                    If mTy_Insure.dbl个帐透支 <> 0 Then
                        dbl个帐 = mTy_Insure.dbl帐户余额 - dbl个帐合计 + mTy_Insure.dbl个帐透支 '在允许透支范围内支付
                    Else
                        dbl个帐 = 0
                    End If
                End If
            End If
        End If
        dbl个帐合计 = dbl个帐合计 + dbl个帐
        dbl个帐 = Format(dbl个帐, "0.00")
        GetYBActualMoeny = dbl个帐
    Else
        GetYBActualMoeny = dbl结算金额
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    GetYBActualMoeny = dbl结算金额
End Function

Private Sub zl9InsureLedSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保预结Led报价
    '编制:刘兴洪
    '日期:2014-09-18 13:43:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double
    If Not gblnLED Then Exit Sub
    dbl个帐合计 = GetMedicareBalanceSum(mstr个人帐户)
    zl9LedVoice.DisplayBank "医保结算:", "帐户余额" & Format(mTy_Insure.dbl帐户余额, "0.00"), "帐户支付" & Format(dbl个帐合计, "0.00"), "统筹支付" & Format(GetMedicareBalanceSum - dbl个帐合计, "0.00")
    zl9LedVoice.Speak "#21 " & Format(-1 * GetMedicareBalanceSum, "0.00")
End Sub

Public Function GetMedicareBalanceSum(Optional strItem As String, Optional blnOrig As Boolean) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保险结算的金额
    '   strItem=是否指定结算方式,否则为所有结算方式
    '   blnOrig=是否取原始(最大)结算金额,否则取现在(修改后)有效金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 13:44:21
    '说明：该函数以mcolBalance为准计算,对于医保划价收费也是
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, curMoney As Currency
    Dim i As Integer
    For i = 1 To mcolBalance.Count
        '结算方式;原始(最大)金额;可否修改;有效金额
        varData = Split(mcolBalance(i), ";")
        If strItem = "" Or (strItem <> "" And varData(0) = strItem) Then
            If blnOrig Then
                curMoney = curMoney + CCur(varData(1))
            Else
                curMoney = curMoney + CCur(varData(3))
            End If
        End If
    Next
    GetMedicareBalanceSum = Format(curMoney, "0.00")
End Function

Private Function GetMedicareBalanceStr(ByRef cur个帐 As Currency) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保险结算数据
    '出参:cur个帐-返回挂号时的个人帐户支付
    '返回:返回保险结算方式串,"结算方式,金额|...."
    '编制:刘兴洪
    '日期:2014-09-17 16:01:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strTemp As String
    Dim varData As Variant
    strTemp = ""
    cur个帐 = 0
    If mEditType = EM_Balance_Register Then
        With vsBalance
            For i = 1 To .COLS - 1 Step 2
                If .TextMatrix(0, i) = mstr个人帐户 Then
                    cur个帐 = Val(.TextMatrix(0, i + 1))
                    strTemp = strTemp & "|" & .TextMatrix(0, i) & "," & Format(Val(.TextMatrix(0, i + 1)), "0.00")
                End If
            Next
        End With
    Else
        For i = 1 To mcolBalance.Count
            '结算方式;原始(最大)金额;可否修改;有效金额
            varData = Split(mcolBalance(i), ";")
            strTemp = strTemp & "|" & varData(0) & "," & Format(varData(3), "0.00")
        Next
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    GetMedicareBalanceStr = strTemp
End Function

Private Function IsRegister(Optional ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断当前费用是挂号费用
    '出参:strNO-挂号单号
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-28 12:37:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnRegister As Boolean
    
    On Error GoTo errHandle
    strNo = ""
    If mrsList Is Nothing Then Exit Function
    If mrsList.State <> 1 Then Exit Function
    If mrsList.RecordCount = 0 Then Exit Function
    mrsList.Filter = "记录性质=1"
    blnRegister = mrsList.RecordCount = 0
    mrsList.Filter = 0
    If blnRegister Then
        If Not mrsList.EOF Then strNo = Nvl(mrsList!NO)
    End If
    IsRegister = blnRegister
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CancelBalance() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:异常作废
    '编制:刘兴洪
    '日期:2014-06-19 14:42:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str冲销ID As String, dtDelDate As Date
    Dim blnTrans As Boolean, strSQL As String
    Dim cllPro As Collection, strRegNO As String '挂号单号
    Dim blnReg As Boolean
    
    '并发检查
    If zlIsCheckExistErrBill(Val(mstr结算序号), True) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(Val(mstr结算序号)) Then
        MsgBox "当前单据正在其它补结算窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    
    blnReg = IsRegister(strRegNO)
    dtDelDate = zlDatabase.Currentdate
    str冲销ID = zlDatabase.GetNextId("病人结帐记录")
    
    
    Set cllPro = New Collection
    'Zl_费用补充记录_补结算作废
    strSQL = "Zl_费用补充记录_补结算作废("
    '  No_In         In 费用补充记录.No%Type,
    strSQL = strSQL & "'" & mstrNo & "',"
    '  冲销id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "" & str冲销ID & ","
    '  结算序号_In   In 费用补充记录.结算序号%Type,
    strSQL = strSQL & "" & "-" & str冲销ID & ","
    '  操作员编号_In In 费用补充记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In In 费用补充记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  登记时间_In   In 费用补充记录.登记时间%Type := Null,
    strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    zlAddArray cllPro, strSQL
    Err = 0: On Error GoTo Errhand:
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If blnReg Then
        '调用挂号作废
        If ExcuteInsureRegistDel(mintInsure, mobjPatiInfor.病人ID, mstr结算ID, str冲销ID, strRegNO) = False Then Exit Function
    Else
        If ExcuteInsureDel(mintInsure, mobjPatiInfor.病人ID, mstr结算ID, str冲销ID) = False Then Exit Function
    End If
    blnTrans = False: CancelBalance = True
    Exit Function
Errhand:
   If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ExcuteInsureRegistDel(ByVal intInsure As Integer, ByVal lng病人ID As Long, _
    ByVal str原结帐ID As String, ByVal str冲销ID As Long, ByVal strRegNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用医保退号接口
    '返回:调用成功,返回true,否则返回False
    '编制:冉俊明
    '日期:2014-10-27
    '说明:需要在外层启用事务;
    '     如果失败,则事务将回退(主要是避免弹出界面造成死锁)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, blnTransMedicare As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    strAdvance = "0|" & strRegNO & "|1"
    If Not gclsInsure.RegistDelSwap(Val(str原结帐ID), intInsure, strAdvance) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    blnTransMedicare = True
    'Zl_费用补充结算_Modify
    strSQL = "Zl_费用补充结算_Modify("
    '  操作类型_In   Number,
    strSQL = strSQL & "" & "2" & ","
    '  结算id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "" & str冲销ID & ","
    '  结算方式_In   Varchar2,:结算方式|结算金额||.."
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "NULL,"
    '  完成结算_In Number:=0
    strSQL = strSQL & "2)" '1-完成补充结算;0-未完成补充结算;2-完成了异常作废
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, True, intInsure)
    
    ExcuteInsureRegistDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, False, intInsure)
    Call ErrCenter
End Function

Private Function ExcuteInsureDel(ByVal intInsure As Integer, ByVal lng病人ID As Long, _
    ByVal str原结帐ID As String, ByVal str冲销ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用医保退费接口
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 18:20:38
    '说明:需要在外层启用事务,正常退费后,该过程不提交,需要调用者提交;
    '     如果失败,则事务将回退(主要是避免弹出界面造成死锁)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, blnTransMedicare As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strAdvance = str冲销ID & "|1"
    
    blnTransMedicare = False
    If Not gclsInsure.ClinicDelSwap(Val(str原结帐ID), , intInsure, strAdvance) Then
         gcnOracle.RollbackTrans: Exit Function
    End If
    blnTransMedicare = True
    If Val(strAdvance) = str冲销ID Or strAdvance = "" Then
        'Zl_费用补充结算_Modify
        strSQL = "Zl_费用补充结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & "2" & ","
        '  结算id_In     In 费用补充记录.结算id%Type,
        strSQL = strSQL & "" & str冲销ID & ","
        '  结算方式_In   Varchar2,:结算方式|结算金额||.."
        strSQL = strSQL & "NULL,"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "NULL,"
        '  完成结算_In Number:=0
        strSQL = strSQL & "2)" '1-完成补充结算;0-未完成补充结算;2-完成了异常作废
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
        gcnOracle.CommitTrans: ExcuteInsureDel = True
        Exit Function
    End If
    '根据返回的结算信息，修正预交记录，strAdvance返回格式:结算方式1|金额||结算方式2|金额...
    If InStr(strAdvance, "|") > 0 Then
        '更新标志:
        'Zl_费用补充结算_Modify
        strSQL = "Zl_费用补充结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & "2" & ","
        '  结算id_In     In 费用补充记录.结算id%Type,
        strSQL = strSQL & "" & str冲销ID & ","
        '  结算方式_In   Varchar2,:结算方式|结算金额||.."
        strSQL = strSQL & "'" & strAdvance & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "NULL,"
        '  完成结算_In Number:=0
        strSQL = strSQL & "2)" '1-完成补充结算;0-未完成补充结算;2-完成了异常作废
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
    gcnOracle.CommitTrans
    ExcuteInsureDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intInsure)
    Call ErrCenter
End Function

Public Function GetBalanceInsure(ByVal str结算ID As String, _
    ByRef str险类名称 As String, Optional ByRef lng病人ID As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保险序事情
    '出参:str险类名称-险类名称
    '     lng病人ID-病人ID
    '返回:返回险类
    '编制:刘兴洪
    '日期:2014-09-22 13:57:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errH
    strSQL = "" & _
    "    Select /*+ rule */  B.记录ID,B.险类,B.病人ID,C.名称" & _
    "    From 　保险结算记录 B,保险类别 C" & _
    "    Where B.记录ID=[1] and B.险类=C.序号(+) And B.性质=1  "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结算ID)
    If Not rsTmp.EOF Then
        lng病人ID = Nvl(rsTmp!病人ID, 0)
        str险类名称 = Nvl(rsTmp!名称)
        GetBalanceInsure = Nvl(rsTmp!险类, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsDiagnose_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '修改后,需要重新过滤
    If Not mEditType = EM_Balance_Charge Then Exit Sub
    '选择
    Call FromDiagnoseSelFee
    
    mblnEdit = True
    Call SetButtons
End Sub

Private Sub FromDiagnoseSelFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据诊断选择指定的费用
    '编制:刘兴洪
    '日期:2014-09-26 16:23:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, varData As Variant, strTemp As String
    Dim i As Long, j As Long, q As Long, blnHaversBalanceNo As Boolean
    Dim lng结算序号 As Long
    
    blnHaversBalanceNo = False
    If Not mrsBalanceNO Is Nothing Then
        blnHaversBalanceNo = mrsBalanceNO.State = 1
    End If
    '将全部选中的(灰色部分),置为选中
    Call SetDiagnoseSelStatu(EM_dgGrayToSeled)
    
    '获取所有选中的诊断所对应的Nos
    strNos = GetDiagnoseNos

    '清除所有选择
    Call FromNosSel("", False, False, True)
    Call FromNosSel(strNos, True, True)
End Sub

Private Function GetDiagnoseNos() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选择的诊断的单据号
    '返回:成功,返回诊断所对应的NOs
    '编制:刘兴洪
    '日期:2014-09-28 10:39:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, varData As Variant, strTemp As String
    Dim i As Long, j As Long, q As Long, blnHaversBalanceNo As Boolean
    
    blnHaversBalanceNo = False
    If Not mrsBalanceNO Is Nothing Then
        blnHaversBalanceNo = mrsBalanceNO.State = 1
    End If
    On Error GoTo errHandle
    
    '获取选中的诊断的单据号(多个用逗号分离)
    strNos = ""
    With vsDiagnose
        For i = 0 To .Rows - 1
            For j = 0 To .COLS - 1
                If Abs(Val(.Cell(flexcpChecked, i, j))) = 1 And vsDiagnose.TextMatrix(i, j) <> "" Then
                    strTemp = vsDiagnose.Cell(flexcpData, i, j)
                    If strTemp <> "" Then
                         varData = Split(strTemp, ",")
                         For q = 0 To UBound(varData)
                             Call GetRelatedNos(varData(q), strNos)
                         Next
                    End If
                End If
            Next
        Next
    End With
    GetDiagnoseNos = strNos
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetRelatedNos(ByVal strNo As String, ByRef strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定单据，获取关联单据号
    '入参:strNO-当前单据号
    '出参:strNos-返回关联的单据号(含本身单号)
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-28 10:20:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnHaversBalanceNo As Boolean, varData As Variant
    Dim str结算序号 As String, str结算序号集 As String
    On Error GoTo errHandle
    
    blnHaversBalanceNo = False
    If Not mrsBalanceNO Is Nothing Then
        blnHaversBalanceNo = mrsBalanceNO.State = 1
    End If
    If Not blnHaversBalanceNo Then GoTo CurNOs:
    
    mrsBalanceNO.Filter = "NO='" & strNo & "'"
    str结算序号集 = ""
    Do While Not mrsBalanceNO.EOF
        str结算序号 = ""
        If Not mrsBalanceNO.EOF Then str结算序号 = Val(Nvl(mrsBalanceNO!结算序号))
        If str结算序号 <> 0 And InStr(str结算序号集 & ",", "," & str结算序号 & ",") = 0 Then
            str结算序号集 = str结算序号集 & "," & str结算序号
        End If
        mrsBalanceNO.MoveNext
    Loop
    mrsBalanceNO.Filter = 0
    If str结算序号集 = "" Then GoTo CurNOs:
    
    str结算序号集 = Mid(str结算序号集, 2)
    varData = Split(str结算序号集, ",")
    
    For i = 0 To UBound(varData)
        mrsBalanceNO.Filter = "结算序号=" & IIf(varData(i) = "", "0", varData(i))
        Do While Not mrsBalanceNO.EOF
             If InStr(1, "," & strNos & ",", "," & mrsBalanceNO!NO & ",") = 0 Then
                strNos = strNos & "," & mrsBalanceNO!NO
             End If
            mrsBalanceNO.MoveNext
        Loop
    Next
    
    If InStr(1, "," & strNos & ",", "," & strNo & ",") = 0 Then
       strNos = strNos & "," & strNo
    End If
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    
    GetRelatedNos = True
    Exit Function
CurNOs:
    strNos = strNos & "," & strNo
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    GetRelatedNos = True: Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetDiagnoseSelStatu(ByVal intStatu As mEM_Diagnose_SelStatu)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置诊断的选择状态
    '入参:intStatu=-1，将选择的灰色置为选中
    '     intStatu=0  清除所有选中的诊断
    '     intStatu=1  选择所有的诊断
    '     intStatu=5  全部置选中的设置为灰色
    '编制:刘兴洪
    '日期:2014-09-28 10:00:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, intCurStatu As Integer
    With vsDiagnose
        For i = 0 To .Rows - 1
            For j = 0 To .COLS - 1
                intCurStatu = Abs(Val(.Cell(flexcpChecked, i, j)))
                If intCurStatu = 0 Then intCurStatu = 2
                '将灰色的调整为
                Select Case intStatu
                Case EM_dgGrayToSeled '将选择的灰色置为选中
                     If Abs(intCurStatu) = 5 Then intCurStatu = -1
                Case EM_dgClearAllSeled   '清除所有选中的诊断
                    intCurStatu = 2
                Case EM_dgClearAllSeled  '选择所有的诊断
                    intCurStatu = -1
                Case Else '5-全部置选中的设置为灰色
                    If Abs(intCurStatu) = 1 Then intCurStatu = 5
                End Select
                .Cell(flexcpChecked, i, j) = intCurStatu
            Next
        Next
    End With
End Sub

Private Sub FromNosSel(ByVal strNos As String, ByVal blnSel As Boolean, _
    ByVal blnBeforClearSel As Boolean, Optional blnAllNo As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号选择
    '入参:blnSel-选择
    '     blnBeforClearSel-先清除所有选择
    '     blnAllNo-不区分单据进行处理
    '编制:刘兴洪
    '日期:2014-09-26 17:18:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, i As Long
    
    '选中所有单据
    With vsFeeList
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then
                strNo = Trim(.TextMatrix(i, .ColIndex("NO")))
                If strNo <> "" Then
                    If blnAllNo Then
                        .Cell(flexcpChecked, i, .ColIndex("选择")) = IIf(blnSel, -1, 2)
                    Else
                        If InStr(1, "," & strNos & ",", "," & strNo & ",") > 0 Then
                            .Cell(flexcpChecked, i, .ColIndex("选择")) = IIf(blnSel, -1, 2)
                        ElseIf blnBeforClearSel Then
                            .Cell(flexcpChecked, i, .ColIndex("选择")) = 2
                        End If
                    End If
                End If
            End If
        Next
    End With
    Call CalcTotalMoney
    Call CalcRegisterYBMoney
    mblnEdit = True
End Sub

Private Sub vsDiagnose_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub vsFeeList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, strNo As String, strNos As String
    
    Dim blnSel As Boolean
    With vsFeeList
        If .Col <> .ColIndex("选择") Then Exit Sub
        strNo = Trim(.TextMatrix(Row, .ColIndex("NO")))
        If strNo = "" Then Exit Sub
        
        If mEditType = EM_Balance_Register Then
            '只能选择一个挂号单
            If Abs(Val(.Cell(flexcpChecked, Row, Col))) <> 1 Then
                mblnEdit = True
                Call SetButtons
                Call CalcTotalMoney
                Call CalcRegisterYBMoney
                Exit Sub
            End If
            
            For i = 1 To .Rows - 1
                If .IsSubtotal(i) And i <> Row Then
                    .Cell(flexcpChecked, i, Col) = 2
                End If
            Next
        End If
        blnSel = Abs(Val(.Cell(flexcpChecked, Row, Col))) = 1
        '关联选择
        Call GetRelatedNos(strNo, strNos)
        
        Call FromNosSel(strNos, blnSel, False)
        '将选中的置为灰色状态
        Call SetDiagnoseSelStatu(EM_dgSeledToGray)
        
        mblnEdit = True
        Call SetButtons
        Call CalcTotalMoney
        Call CalcRegisterYBMoney
    End With
End Sub

Private Sub vsFeeList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, mstrTittle, "费用信息列表", True, False
End Sub

Private Sub vsFeeList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, mstrTittle, "费用信息列表", True, False
End Sub

Private Sub vsFeeList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    With vsFeeList
        If .Col <> .ColIndex("选择") Then Cancel = True: Exit Sub
        If .IsSubtotal(Row) = False Then Cancel = True: Exit Sub
        If .ColIndex("类别") < 0 Then Cancel = True: Exit Sub
        If Trim(.TextMatrix(Row, .ColIndex("类别"))) = "" Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsFeeList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    With vsFeeList
        If Col <= .ColIndex("选择") Then
             Position = Col
        End If
    End With
End Sub

Private Sub vsFeeList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsFeeList
        If Col <= .ColIndex("选择") Then Cancel = True
    End With
End Sub

Private Sub vsFeeList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cmd预结算.Enabled And cmd预结算.Visible Then
        cmd预结算.SetFocus
    ElseIf cmdOK.Visible And cmdOK.Enabled Then
        cmdOK.SetFocus
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Dim lngR As Long
    If Panel.Key = "Calc" Then
        lngR = FindWindow("SciCalc", "计算器")
        If lngR <> 0 Then
            BringWindowToTop lngR
        Else
            On Error Resume Next
            Shell "calc.exe", vbNormalFocus
        End If
    End If
End Sub

Private Sub SetButtons()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置功能按钮
    '编制:刘兴洪
    '日期:2014-09-23 11:51:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdCancel.Enabled = True: cmdCancel.Visible = True
    
    If mEditType = EM_Balance_Register Or mEditType = EM_Balance_Err_ReCharge Or mEditType = EM_Balance_Err_Cancel Then
        cmd预结算.Visible = False: cmd预结算.Enabled = False
        cmdOK.Enabled = True: cmdOK.Visible = True
        cmdSelAll.Visible = False: cmdClear.Visible = False
        Call picDown_Resize
        Exit Sub
    End If
    If mobjPatiInfor Is Nothing Then
        cmdOK.Enabled = False: cmd预结算.Visible = False: cmd预结算.Enabled = False
        cmdSelAll.Visible = False: cmdClear.Visible = False
        Exit Sub
    End If
    cmdSelAll.Visible = True: cmdClear.Visible = True
    
    '支持预结算时就不固定显示个人帐户,否则显示
    If MCPAR.门诊预结算 Then
        '显示预结算按钮
        cmd预结算.Enabled = mblnEdit  '是否编辑过且未重新预结算的
        cmd预结算.Visible = True
        cmdOK.Enabled = Not mblnEdit: cmdOK.Visible = True
        Call picDown_Resize
        Exit Sub
    End If
    If mstr个人帐户 <> "" Then '只有使用个人帐户才用
        cmd预结算.Visible = False: cmd预结算.Enabled = False
        cmdOK.Enabled = True: cmdOK.Visible = True
        Call picDown_Resize
    End If
End Sub

Private Sub CalcTotalMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算相关的合计金额
    '编制:刘兴洪
    '日期:2014-09-23 12:09:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnSel As Boolean
    Dim dblMoney(0 To 1) As Double
    Dim blnTotalSelect As Boolean
    
    lbl应收.Tag = "": lbl实收.Tag = ""
    With vsFeeList
        dblMoney(0) = 0: dblMoney(1) = 0
        For i = 1 To .Rows - 1
            If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then
                blnSel = True
            Else
                blnSel = Abs(Val(.Cell(flexcpChecked, i, .ColIndex("选择")))) = 1
            End If
            
            If .IsSubtotal(i) Then
                If blnSel Then
                    dblMoney(0) = dblMoney(0) + Val(.Cell(flexcpData, i, .ColIndex("应收金额")))
                    dblMoney(1) = dblMoney(1) + Val(.Cell(flexcpData, i, .ColIndex("实收金额")))
                End If
                blnTotalSelect = blnSel '记录汇总行选择状态
            Else
                '汇总行被选择，下面的子项也就被选择
                If blnTotalSelect Then
                    '统计病历费，84965
                    If mEditType = EM_Balance_Register _
                        And Val(.TextMatrix(i, .ColIndex("收费细目ID"))) = mlng病历费细目ID Then
                        mcur病历费 = mcur病历费 + Val(.Cell(flexcpData, i, .ColIndex("实收金额")))
                    End If
                End If
            End If
        Next
    End With
    lbl应收.Caption = "应收:" & Format(dblMoney(0), "0.00")
    lbl应收.Tag = dblMoney(0)
    lbl实收.Caption = "实收:" & Format(dblMoney(1), "0.00")
    lbl实收.Tag = dblMoney(1)
    
    '清除预结算信息
    vsBalance.Clear 1: vsBalance.COLS = 1: txt退款合计.Text = "0.00"
End Sub

Private Sub CalcRegisterYBMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算并显示挂号当前医保病人个人帐户可以支持的金额
    '编制:刘兴洪
    '日期:2014-09-23 12:07:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur合计 As Currency
    Dim strInfo As String, i As Long, j As Long, lng病人ID As Long
    Dim dbl个人帐户 As Double
    Dim blnFind As Boolean
    
    '80238,冉俊明,2014-11-27
    If mEditType <> EM_Balance_Register Then Exit Sub
    If mstrYBPati <> "" Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
    
    cur合计 = Val(lbl实收.Tag)
    If MCPAR.不收病历费 Then '不收病历费，84965
        cur合计 = cur合计 - mcur病历费
    End If
    '计算并显示个人帐户支付金额
    '要求医保支持个人帐户支付及ZLHIS允许使用个人帐户
    dbl个人帐户 = 0
    If mintInsure <> 0 And mstr个人帐户 <> "" Then
        If gclsInsure.GetCapability(support挂号使用个人帐户, lng病人ID, mintInsure) Then
            If mTy_Insure.dbl帐户余额 - cur合计 >= -1 * mTy_Insure.dbl个帐透支 Then
               dbl个人帐户 = Format(cur合计, "0.00")  '在允许透支范围内足够(允许透支0为特例)
            Else
                If mTy_Insure.dbl个帐透支 = 0 And mTy_Insure.dbl帐户余额 > 0 Then
                    dbl个人帐户 = mTy_Insure.dbl帐户余额  '不允许透支且有余额
                Else
                    dbl个人帐户 = 0 '超过允许透支范围或不允许透支时无余额
                End If
            End If
        End If
    End If
    blnFind = False
    With vsBalance
        .Clear 1
        .Rows = 1
        .COLS = 1
        If blnFind = False Then
            j = -1
            For i = 1 To .COLS - 1 Step 2
                If .TextMatrix(0, i) = "" Then j = i: Exit For
            Next
            If j < 0 Then .COLS = .COLS + 2: j = .COLS - 2
            .TextMatrix(0, i) = mstr个人帐户
            .TextMatrix(0, i + 1) = Format(dbl个人帐户, "0.00")
        End If
        txt退款合计 = Format(dbl个人帐户, "0.00")
    End With
End Sub

Private Function SaveItemYbMoney(ByVal lng病人ID As Long, ByVal strNos As String, _
    ByVal int记录性质 As Integer, Optional ByRef cur全自付 As Currency, _
    Optional ByRef cur先自付 As Currency, Optional ByRef cur进入统筹 As Currency) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改挂号或收费项目的统筹金额
    '入参:strNo-单号,多个用逗号分离
    '     int记录性质-1-收费;4-挂号
    '出参:cur全自付-全自费金额
    '     cur先自付-先自付金额
    '     cur进入统筹-统筹金额
    '返回:修改成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-23 14:41:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim cur实收 As Currency, cllPro As Collection
    Dim varData As Variant, strInfo As String
    Dim rsItem As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select ID,NO,记录状态,收费细目id, 收入项目ID, 实收金额 As 实收 " & _
    "   From 门诊费用记录  " & _
    "   Where NO in (Select Column_Value From Table(f_str2List([1]))) " & _
    "       And 记录性质=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos, int记录性质)
    
    If rsTemp.EOF Then Exit Function
    Set rsItem = New ADODB.Recordset
    
    rsItem.Fields.Append "NO", adVarChar, 100, adFldIsNullable
    rsItem.Fields.Append "收费细目id", adBigInt, , adFldIsNullable
    rsItem.Fields.Append "收入项目ID", adBigInt, , adFldIsNullable
    rsItem.Fields.Append "实收金额", adDouble, , adFldIsNullable
    rsItem.Fields.Append "保险项目否", adBigInt, , adFldIsNullable
    rsItem.Fields.Append "保险大类id", adBigInt, , adFldIsNullable
    rsItem.Fields.Append "保险编码", adVarChar, 100, adFldIsNullable
    rsItem.Fields.Append "费用类型", adVarChar, 100, adFldIsNullable
    rsItem.Fields.Append "摘要", adVarChar, 2000, adFldIsNullable
    rsItem.Fields.Append "统筹金额", adDouble, , adFldIsNullable
    rsItem.CursorLocation = adUseClient
    rsItem.LockType = adLockOptimistic
    rsItem.CursorType = adOpenStatic
    rsItem.Open
        
    Do While Not rsTemp.EOF
        rsItem.Filter = "NO='" & Nvl(rsTemp!NO) & "' And 收费细目id=" & Val(Nvl(rsTemp!收费细目ID)) & " And 收入项目ID=" & Val(Nvl(rsTemp!收入项目ID))
        If rsItem.EOF Then
            rsItem.AddNew
            rsItem!NO = CStr(Nvl(rsTemp!NO))
            rsItem!收费细目ID = Val(Nvl(rsTemp!收费细目ID))
            rsItem!收入项目ID = Val(Nvl(rsTemp!收入项目ID))
        End If
        rsItem!实收金额 = Val(Nvl(rsItem!实收金额)) + Val(Nvl(rsTemp!实收))
        rsItem.Update
        rsTemp.MoveNext
    Loop
    rsItem.Filter = 0
    Set cllPro = New Collection
    If rsItem.RecordCount <> 0 Then rsItem.MoveFirst
    Do While Not rsItem.EOF
        cur实收 = Val(Nvl(rsItem!实收金额))
        strInfo = gclsInsure.GetItemInsure(lng病人ID, Val(Nvl(rsItem!收费细目ID)), cur实收, True, mintInsure)
        If strInfo <> "" Then
            '保险项目否(0/1);保险大类ID;进入统筹金额;保险项目编码;摘要;费用类型
            varData = Split(strInfo & ";;;;;", ";")
            rsItem!保险项目否 = Val(varData(0))
            rsItem!保险大类ID = Val(varData(1))
            rsItem!统筹金额 = Val(varData(2))
            rsItem!保险编码 = Trim(varData(3))
            rsItem!摘要 = Trim(varData(4))
            rsItem!费用类型 = Trim(varData(5))
            If Val(varData(2)) = 0 Or Val(varData(0)) = 0 Then
                '以原始金额为准,不管分币处理
                cur全自付 = cur全自付 + cur实收
            Else
                cur进入统筹 = cur进入统筹 + Val(varData(2))
                '以原始金额为准,不管分币处理
                cur先自付 = cur先自付 + (cur实收 - cur进入统筹)
            End If
            rsItem.Update
        Else
            cur全自付 = cur全自付 + cur实收
        End If
        rsItem.MoveNext
    Loop
    
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        rsItem.Filter = "NO='" & Nvl(rsTemp!NO) & "' And 收费细目id=" & Val(Nvl(rsTemp!收费细目ID)) & " And 收入项目ID=" & Val(Nvl(rsTemp!收入项目ID))
        If Not rsItem.EOF And InStr(1, ",1,3,", "," & Val(Nvl(rsTemp!记录状态))) > 0 Then
            'Zl_门诊收费记录_Update
            strSQL = "Zl_门诊收费记录_Update("
            '  Id_In         In 门诊费用记录.Id%Type,
            strSQL = strSQL & "" & rsTemp!ID & ","
            '  保险大类id_In In 门诊费用记录.保险大类id%Type,
            strSQL = strSQL & "" & ZVal(Val(Nvl(rsItem!保险大类ID))) & ","
            '  保险项目否_In In 门诊费用记录.保险项目否%Type,
            strSQL = strSQL & "" & ZVal(Val(Nvl(rsItem!保险项目否))) & ","
            '  保险编码_In   In 门诊费用记录.保险编码%Type,
            strSQL = strSQL & "'" & Nvl(rsItem!保险编码) & "',"
            '  费用类型_In   In 门诊费用记录.费用类型%Type,
            strSQL = strSQL & "'" & Nvl(rsItem!费用类型) & "',"
            '  统筹金额_In   In 门诊费用记录.统筹金额%Type,
            strSQL = strSQL & "" & Val(Nvl(rsItem!统筹金额)) & ","
            '  摘要_In       In 门诊费用记录.摘要%Type
            strSQL = strSQL & "'" & Nvl(rsItem!摘要) & "')"
            zlAddArray cllPro, strSQL
        End If
        rsTemp.MoveNext
    Loop
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveItemYbMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub SetBalanceVal(strItem As String, curVal As Currency)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置指定保险结算方式的有效值
    '编制:刘兴洪
    '日期:2014-09-24 14:39:57
    '说明：该函数以mcolBalance为准计算,对于医保划价收费也是
    '说明：用于正常医保收费修改保险结算金额；及划价单医保收费设置个人帐户等结算金额
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, strTemp As String
    Dim cllNewBalance As Collection, i As Long
    
    Set cllNewBalance = New Collection
    If mcolBalance.Count <> 0 Then
        For i = 1 To mcolBalance.Count
            '结算方式;原始(最大)金额;可否修改;有效金额
            varTemp = Split(mcolBalance(i), ";")
            If varTemp(0) = strItem And varTemp(3) <> curVal Then
                strTemp = varTemp(0) & ";" & varTemp(1) & ";" & varTemp(2) & ";" & Format(curVal, "0.00")
            Else
                strTemp = varTemp(0) & ";" & varTemp(1) & ";" & varTemp(2) & ";" & varTemp(3)
            End If
            cllNewBalance.Add strTemp
        Next
    Else
        '无内容时强行增加:不支持预结算或医保划价收费时用
        strTemp = strItem & ";" & Format(curVal, "0.00") & ";0;" & Format(curVal, "0.00")
        cllNewBalance.Add strTemp
    End If
    Set mcolBalance = cllNewBalance
End Sub

Private Function CheckFactValied(Optional blnReCharge As Boolean = False, _
    Optional ByRef blnPrintBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据号码的合法性检查
    '入参:blnReCharge-是否重新收费的检查
    '出参:mblnPrintBill-是否打印票据
    '返回:数据合法,返回tru,否则返回false
    '编制:刘兴洪
    '日期:2014-09-24 17:30:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    '票据号码检查,工本费打印检查
    blnPrintBill = True
    '检查是否打印票据
    If mobjFactProperty.打印方式 = 0 Then
        blnPrintBill = False
    Else
        If mobjFactProperty.打印方式 = 2 Then
            If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                blnPrintBill = False
            End If
        End If
    End If
    
    '不打印直接退出
    If Not blnPrintBill Then CheckFactValied = True: Exit Function

    If Not mobjFactProperty.严格控制 Then
        If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
            MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus: Exit Function
        End If
        CheckFactValied = True: Exit Function
    End If
    
    If Trim(txtInvoice.Text) = "" Then
        MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
        If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus: Exit Function
    End If

InvoiceHandle:
    If zlGetInvoiceGroupUseID(mlng领用ID, 1, txtInvoice.Text) = False Then Exit Function

    '并发操作检查,票号是否已用
    If CheckBillRepeat(mlng领用ID, 1, txtInvoice.Text) Then
        If txtInvoice.Locked = False And txtInvoice.Tag <> Trim(txtInvoice.Text) Then
            MsgBox "票据号""" & txtInvoice.Text & """已经被使用，请重新输入。", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus: Exit Function
        End If
        
        Call RefreshFact
        If txtInvoice.Text = "" Then
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus: Exit Function
        End If
        MsgBox "当前票据号已经被使用，已重新获取票据号:" & txtInvoice.Text, vbInformation, gstrSysName
        GoTo InvoiceHandle:
    End If
    CheckFactValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case EM_Pan_Pati
        Item.Handle = picTop.hWnd
    Case EM_Pan_Diagnose
        Item.Handle = picDiagnose.hWnd
    Case EM_Pan_FeeList
        Item.Handle = picFeeList.hWnd
    Case EM_Pan_Down
        Item.Handle = picDown.hWnd
    End Select
End Sub

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区哉设置
    '编制:刘兴洪
    '日期:2009-09-14 18:06:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim sngHight As Single
    Dim panLeft As Pane
    
    '病人信息及单据部分
    
    sngHight = picTop.Height \ Screen.TwipsPerPixelY
    Set panThis = dkpMan.CreatePane(EM_Pan_Pati, 200, sngHight, DockLeftOf, Nothing)
    panThis.MaxTrackSize.Height = sngHight
    panThis.MinTrackSize.Height = sngHight
    panThis.Title = "": panThis.Tag = EM_Pan_Diagnose
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picTop.hWnd
    
    
    If mEditType = EM_Balance_Charge Then
        '诊断列表
        sngHight = picDiagnose.Height \ Screen.TwipsPerPixelY
        Set panThis = dkpMan.CreatePane(EM_Pan_Diagnose, 200, sngHight, DockBottomOf, panThis)
        panThis.MaxTrackSize.Height = sngHight
        panThis.MinTrackSize.Height = sngHight
        panThis.Title = "诊断选择": panThis.Tag = EM_Pan_Diagnose
        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        panThis.Handle = picDiagnose.hWnd
    Else
        picDiagnose.Visible = False
    End If
    
    
    '费用信息列表
    Set panThis = dkpMan.CreatePane(EM_Pan_FeeList, 250, 580, DockBottomOf, panThis)
    panThis.Title = "当前已结费用信息"
    panThis.Tag = EM_Pan_FeeList
    panThis.Handle = picFeeList.hWnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    '表下项
    sngHight = picDown.Height \ Screen.TwipsPerPixelY
    Set panThis = dkpMan.CreatePane(EM_Pan_Down, 200, 580, DockBottomOf, panLeft)
    panThis.Title = "": panThis.Tag = EM_Pan_Down
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picDown.hWnd
    panThis.MaxTrackSize.Height = sngHight
    panThis.MinTrackSize.Height = sngHight
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub

Private Sub SetControlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的属性
    '编制:刘兴洪
    '日期:2014-09-26 14:53:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    blnEdit = mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register
    txt摘要.Enabled = blnEdit
    txt摘要.BackColor = IIf(blnEdit, &H80000005, Me.BackColor)
    cboNO.Enabled = blnEdit
    cboNO.BackColor = IIf(blnEdit, &H80000005, Me.BackColor)
    
    PatiIdentify.Enabled = blnEdit
    PatiIdentify.AllowAutoCommCard = blnEdit
    PatiIdentify.AllowAutoICCard = blnEdit
    PatiIdentify.AllowAutoIDCard = blnEdit
    
    blnEdit = Not mEditType = EM_Balance_Err_Cancel
    txtInvoice.Enabled = blnEdit
    txtInvoice.BackColor = IIf(blnEdit, &H80000005, Me.BackColor)
End Sub

Private Sub reSizeWinControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新调整窗体控件位置
    '编制:刘兴洪
    '日期:2014-09-26 14:54:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dkpMan.RecalcLayout
    Call picTop_Resize
    Call picDiagnose_Resize
    Call picDown_Resize
    Call picFeeList_Resize
End Sub

Private Function ShowReclaimInvoice(ByVal strNos As String, ByRef strReclaimInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示和返回需要回收的发票
    '入参:strNos-当前的单据号,多个用逗号分离(如果是被充结算,则为补充结算单号)
    '出参:strReclaimInvoice-返回回收的发票号(多个用逗号分隔),格式:AAAA,BBB,....)
    '返回:显示或获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-10 17:53:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmReInvoiceTemp As frmReInvoice
    Dim strSQL As String, rsTemp As ADODB.Recordset, blnFee As Boolean '当前结算是否为收费结算
    
    On Error GoTo errHandle
    '确定当前结算是否为收费结算
    If mEditType = EM_Balance_Err_ReCharge Then
        strSQL = "Select 1 From 费用补充记录 Where Nvl(附加标志, 0) = 0 And 结算序号 = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "确定当前异常结算是否为收费异常结算", Val(mstr结算序号))
        If Not rsTemp.EOF Then blnFee = True
    End If
    blnFee = blnFee Or mEditType = EM_Balance_Charge
    
    Set frmReInvoiceTemp = New frmReInvoice
    If frmReInvoiceTemp.ShowMe(Me, strNos, 0, 0, strReclaimInvoice, True, IIf(blnFee, 1, 4)) = False Then Exit Function
    If Not frmReInvoiceTemp Is Nothing Then Unload frmReInvoiceTemp
    Set frmReInvoiceTemp = Nothing
    ShowReclaimInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ClearDisplaySHow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除双屏显示
    '编制:刘兴洪
    '日期:2014-10-13 15:07:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    If Not gblnLED Then Exit Sub
    If mblnNotClearLedDisplay Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub

Private Function UpdateBalance(ByVal cur实收合计 As Currency, ByVal cur进入统筹 As Currency, _
        ByVal cur全自付 As Currency, ByVal cur先自付 As Currency) As Boolean
    '更新当前单据个人帐户支付金额:不支持预结算时
    '医保病人且满足相应条件才处理,合计为负不能退到个人帐户
    Dim cur个帐 As Currency, cur可用个帐 As Currency
    Dim i As Integer, j As Integer, blnFind As Boolean
    
    On Error GoTo Errhand
    If mstrYBPati <> "" And mstr个人帐户 <> "" And mTy_Insure.dbl帐户余额 > -1 * mTy_Insure.dbl个帐透支 Then
        If cur实收合计 >= 0 Then
            cur个帐 = cur进入统筹 + IIf(MCPAR.先自付, cur先自付, 0) + IIf(MCPAR.全自付, cur全自付, 0)
            cur可用个帐 = mTy_Insure.dbl帐户余额
            '计算个人帐户支付金额
            If cur可用个帐 - cur个帐 >= -1 * mTy_Insure.dbl个帐透支 Then
                Call SetBalanceVal(mstr个人帐户, Format(cur个帐, "0.00"))   '在允许透支范围内足够(允许透支0为特例)
            Else
                If mTy_Insure.dbl个帐透支 = 0 And cur可用个帐 > 0 Then
                    Call SetBalanceVal(mstr个人帐户, Format(cur可用个帐, "0.00"))  '不允许透支且有余额
                Else
                    '超过允许透支范围或不允许透支时无余额
                    If mTy_Insure.dbl个帐透支 <> 0 Then
                        Call SetBalanceVal(mstr个人帐户, cur可用个帐 + mTy_Insure.dbl个帐透支) '在允许透支范围内支付
                    Else
                        Call SetBalanceVal(mstr个人帐户, 0)
                    End If
                End If
            End If
        Else
            Call SetBalanceVal(mstr个人帐户, 0)
        End If
        '刷新显示个人帐户支付情况
        '-------------------------------------------------------------------------
        With vsBalance
            For i = 1 To .COLS - 1 Step 2
                If .TextMatrix(0, i) = mstr个人帐户 And mstr个人帐户 <> "" Then
                    .TextMatrix(0, i + 1) = Format(GetMedicareBalanceSum(mstr个人帐户), "0.00")
                    blnFind = True: Exit For
                End If
            Next
            If blnFind = False Then
                j = -1
                For i = 1 To .COLS - 1 Step 2
                    If .TextMatrix(0, i) = "" Then j = i: Exit For
                Next
                If j < 0 Then .COLS = .COLS + 2: j = .COLS - 2
                .TextMatrix(0, i) = mstr个人帐户
                .TextMatrix(0, i + 1) = Format(GetMedicareBalanceSum(mstr个人帐户), "0.00")
            End If
        End With
    End If
    UpdateBalance = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
