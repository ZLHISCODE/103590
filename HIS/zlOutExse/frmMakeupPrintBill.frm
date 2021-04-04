VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMakeupPrintBill 
   Caption         =   "门诊收费补打发票"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmMakeupPrintBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11790
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBalance 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   6525
      ScaleHeight     =   4695
      ScaleWidth      =   4380
      TabIndex        =   17
      Top             =   1185
      Width           =   4380
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   2685
         Left            =   525
         TabIndex        =   18
         Top             =   300
         Width           =   8505
         _cx             =   15002
         _cy             =   4736
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMakeupPrintBill.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
         ExplorerBar     =   2
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
      Begin VB.Label lbl合计 
         AutoSize        =   -1  'True
         Caption         =   "补打合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   465
         TabIndex        =   19
         Top             =   3165
         Width           =   1155
      End
   End
   Begin VB.PictureBox PicDetail 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   75
      ScaleHeight     =   2775
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   4290
      Width           =   5535
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         Height          =   2685
         Left            =   180
         TabIndex        =   16
         Top             =   330
         Width           =   8505
         _cx             =   15002
         _cy             =   4736
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMakeupPrintBill.frx":03CF
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
         ExplorerBar     =   2
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
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   90
      ScaleHeight     =   2295
      ScaleWidth      =   5400
      TabIndex        =   13
      Top             =   1290
      Width           =   5400
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   105
         TabIndex        =   14
         Top             =   75
         Width           =   4650
         _cx             =   8202
         _cy             =   4736
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMakeupPrintBill.frx":03E5
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
         ExplorerBar     =   2
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
   Begin VB.PictureBox picCon 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   180
      ScaleHeight     =   675
      ScaleWidth      =   14475
      TabIndex        =   5
      Top             =   135
      Width           =   14475
      Begin VB.CheckBox chkRegistFee 
         Caption         =   "含挂号费"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8850
         TabIndex        =   25
         Top             =   210
         Width           =   1320
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   375
         Left            =   555
         TabIndex        =   24
         Top             =   150
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         Appearance      =   2
         IDKindStr       =   "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         MustSelectItems =   "姓名"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   9
         Top             =   150
         Width           =   2040
      End
      Begin VB.CommandButton cmdBrush 
         Caption         =   "刷新(&N)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10245
         TabIndex        =   6
         Top             =   150
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   5175
         TabIndex        =   7
         Top             =   180
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   7125
         TabIndex        =   10
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "忽略发生时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3420
         TabIndex        =   8
         Top             =   210
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "病人"
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
         Index           =   7
         Left            =   90
         TabIndex        =   12
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
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
         Left            =   6840
         TabIndex        =   11
         Top             =   225
         Width           =   240
      End
   End
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   45
      ScaleHeight     =   660
      ScaleWidth      =   11700
      TabIndex        =   0
      Top             =   7170
      Width           =   11700
      Begin VB.TextBox txtInvoice 
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
         IMEMode         =   3  'DISABLE
         Left            =   3780
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "补打(&O)"
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
         Left            =   8565
         TabIndex        =   4
         Top             =   210
         Width           =   1440
      End
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
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "热键：Ctrl+A"
         Top             =   225
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
         TabIndex        =   2
         ToolTipText     =   "热键：Ctrl+R"
         Top             =   225
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   10095
         TabIndex        =   1
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   6105
         TabIndex        =   23
         Top             =   300
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3285
         TabIndex        =   22
         Top             =   330
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   8055
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMakeupPrintBill.frx":03FB
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15716
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   1500
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeupPrintBill.frx":0C8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeupPrintBill.frx":0FE3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   30
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMakeupPrintBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPanel
    Pane_Search = 1
    Pane_List = 2
    Pane_Detail = 3
    Pane_Balance = 4
End Enum
'-----------------------------------------------------------------------------------
'结算卡相关
Private mSquareCard As SquareCard '结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mrsInfo As ADODB.Recordset
Private mrsList As ADODB.Recordset  '单据列表
Private mrsDetail As ADODB.Recordset
Private mrsBalance As ADODB.Recordset
Private mlngModule As Long
Attribute mlngModule.VB_VarHelpID = -1
Private mblnValid As Boolean
Private mblnSel As Boolean
Private mstrPrivs As String
Private mintSucces As Integer  '成功打印次数
Private mlng病人ID As Long
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mlngShareUseID As Long '共享领用批次ID
Private mstrUseType As String '使用类别
Private mintInvoiceFormat As Integer  '打印的发票格式,发票格式序号
Private mblnStartFactUseType As Boolean   '是否启用了使用类别
Private mintInvoicePrint As Integer  '0-不打印;1-自动打印;2-提示打印
Private mlng领用ID As Long
Private mintInsure As Integer

'相关参数
Private mbln不分结算次数  As Boolean
Private mintPatiInvoiceFormat As Integer '不分结算次数打印的发票格式

Public Function zlRePrintBill(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal strPrivs As String, Optional lng病人ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打票据入口
    '返回:打印成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-09-04 22:39:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mintSucces = 0
    mlng病人ID = lng病人ID
    
    mbln不分结算次数 = Val(zlDatabase.GetPara("按病人补打发票不区分结算次数", glngSys, mlngModule, "")) = 1
  
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlRePrintBill = mintSucces > 0
End Function

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区哉
    '编制:刘兴洪
    '日期:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, strKey As String
    Dim objTemp As Object
    With dkpMan
        .ImageList = imlPaneIcons
        Set objPane = .CreatePane(mPanel.Pane_Search, 200, 100, DockLeftOf, Nothing)
        objPane.Tag = mPanel.Pane_Search
        objPane.Title = "条件设置": objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoCaption
        objPane.MaxTrackSize.Height = 675 \ Screen.TwipsPerPixelY
        objPane.MinTrackSize.Height = 675 \ Screen.TwipsPerPixelY
        objPane.Handle = picCon.hWnd
        Set objTemp = .CreatePane(mPanel.Pane_List, 300, 100, DockBottomOf, objPane)
        objTemp.Tag = mPanel.Pane_List
        objTemp.Title = "单据列表": objTemp.Options = PaneNoCloseable Or PaneNoHideable
        objTemp.Handle = picList.hWnd
        
        Set objPane = .CreatePane(mPanel.Pane_Balance, 100, 100, DockRightOf, objTemp)
        objPane.Tag = mPanel.Pane_Balance
        objPane.Title = "结算信息": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
       '
        Set objPane = .CreatePane(mPanel.Pane_Detail, 300, 100, DockBottomOf, objTemp)
        objPane.Tag = mPanel.Pane_Detail
        objPane.Title = "单据明细列表": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = PicDetail.hWnd
       '  .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    dkpMan.RecalcLayout: DoEvents
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
    'Call GetRegInFor(g私有模块, Me.Name, "隐藏", strKey)
    'If Val(strKey) = 1 Then mPanSearch.Hide
End Function

Private Sub chkDate_Click()
    dtpBegin.Enabled = chkDate.Value <> 1
    dtpEnd.Enabled = chkDate.Value <> 1
End Sub

Private Sub cmdBrush_Click()
    Call ReadListData
End Sub

Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdClear_Click()
    With vsList
        If .Rows <= .FixedRows Then Exit Sub
        If .ColIndex("选择") < 0 Then Exit Sub
        .Cell(flexcpChecked, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
    End With
    Call SetBlanceShow
    Call InitPatiInsure
End Sub

Private Function zlMakeupPrint(ByVal lng病人ID As Long, _
    ByVal strNos As String, _
    ByVal strUseType As String, _
    ByVal strBillNameDemo As String, _
    ByVal intInvoiceFormat As Integer, _
    ByVal blnVirtualPrint As Boolean, _
    ByVal intInusre As Integer, _
    ByRef lng领用ID As Long, _
    ByVal lngShareUseID As Long, _
    ByVal strFactNO As String, _
    Optional ByVal str结帐IDs As String = "", _
    Optional strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补打票据
    '入参:strNos-需要打印的NO
    '     strUseType-使用类别
    '     strBillNameDemo-票据格式说明
    '     intInvoiceFormat-发票打印格式
    '     blnVirtualPrint-是否医只接口打印票据
    '     intInusre-险类
    '     blnOnePrint-是否一次打印(true-是一次打印，不分结算次数,否则分结算次数打印)
    '     strFactNo-发票号
    '     str结帐IDs-本次打印涉及的结帐IDs,多个用逗号分隔
    '出参:lng领用ID-返回领用ID
    '返回:补打成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-04 22:58:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long
    Dim bln分别打印 As Boolean, lng打印ID As Long
    


    If strNos = "" Then Exit Function
    If lng病人ID = 0 Then Exit Function
    If strNos = "" Then
        MsgBox "当前没有单据可以重打票据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not CheckFP(strNos, strUseType, strBillNameDemo, strFactNO, lng领用ID, lngShareUseID) Then Exit Function
           
    '--------------------------------------------------------------------------------------
    '处理临时数据
    If mbln不分结算次数 Then
        If zlSaveTempPrintData(strNos, lng领用ID, strFactNO, lng打印ID) = False Then Exit Function
    End If
    '--------------------------------------------------------------------------------------
    
    '单据有剩余数量的才可以重打，北京医保，即使退完了也可以重新打印
    If Not blnVirtualPrint Then
        If Not BillExistMoney(strNos, 1, , lng打印ID) Then
            MsgBox "单据[" & strNos & "]中的项目已经全部退费，不能进行打印！", vbInformation, gstrSysName
            Call zlDeleteTempPrintData(lng打印ID)
            Exit Function
        End If
    End If
    
    '冉俊明,2014-12-17,补结算后的收费单据不允许重打票据
    If CheckBillExistReplenishData(2, , Replace(strNos, "'", ""), lng打印ID) = True Then
        MsgBox "单据[" & strNos & "]中存在已经进行了保险补充结算的项目，不能进行打印！", vbInformation, gstrSysName
        Call zlDeleteTempPrintData(lng打印ID)
        Exit Function
    End If
    
    Dim dtDate As Date
    dtDate = zlDatabase.Currentdate
    
    bln分别打印 = gTy_Module_Para.bln分别打印
    If mbln不分结算次数 Then bln分别打印 = False
    
    strNos = "'" & Replace(strNos, ",", "','") & "'"
    Call frmPrint.ReportPrint(1, strNos, "", "", lng领用ID, lngShareUseID, strFactNO, dtDate, "", "", _
        bln分别打印, intInvoiceFormat, blnVirtualPrint, , strUseType, , mbln不分结算次数, lng打印ID, strPriceGrade)
    
    mintSucces = mintSucces + 1
    zlMakeupPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckFP(ByVal strNos As String, ByVal strUserType As String, ByVal strBillNameDemo As String, ByRef strFactNO As String, ByRef lng领用ID As Long, ByRef lngShareID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否正确
    '入参:strNos-根据NO来获取发票号
    '     strUserType-票据使用类别
    '     lngShareID-当前共用批次
    '     strFactNo-发票号
    '出参:lng领用ID-返回领用ID
    '     lngShareID-返回共用ID
    '     strFactNo-发票号
    '返回: 发票合法 返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-07-12 11:30:22
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intNum As Integer, varData As Variant
    
    On Error GoTo errHandle
    varData = Split(strNos, ",")
    intNum = UBound(varData) + 1
    If strNos = "" Then
        MsgBox "不存在需要补打的票据", vbInformation, gstrSysName
        Exit Function
    End If
    If Not gblnStrictCtrl Then
       If Len(strFactNO) <> gbytFactLength And strFactNO <> "" Then
           MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
           If InputFactNo(strUserType, strBillNameDemo, lng领用ID, lngShareID, strFactNO) Then Exit Function

        End If
       CheckFP = True
       Exit Function
    End If
     
    If Trim(strFactNO) = "" Then
       MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
       Exit Function
    End If
    
    If Not gTy_Module_Para.bln分别打印 Or mbln不分结算次数 Then intNum = 1
 
InvoiceHandle:
    If zlCheckInvoiceValied(lng领用ID, intNum, strFactNO, lngShareID, strUserType) = False Then Exit Function

    '并发操作检查,票号是否已用
    If CheckBillRepeat(lng领用ID, 1, strFactNO) Then
        MsgBox "票据号""" & strFactNO & """已经被使用，请重新输入。", vbInformation, gstrSysName
        If mblnStartFactUseType = False Then
            txtInvoice.Text = GetNextFactNo(strUserType, lng领用ID, lngShareID)
        End If
        If InputFactNo(strUserType, strBillNameDemo, lng领用ID, lngShareID, strFactNO) Then Exit Function
    End If
   CheckFP = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckPrintValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查打印的合法性
    '入参:
    '返回:
    '编制:刘兴洪
    '日期:2016-04-29 11:58:23
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then
        MsgBox "未选择指定的病人,请选择需要打印发票的病人!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    If mrsInfo.State <> 1 Then
        MsgBox "未选择指定的病人,请选择需要打印发票的病人!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.EOF Then
        MsgBox "未选择指定的病人,请选择需要打印发票的病人!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    CheckPrintValied = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SplitGroupPrint(ByRef cllPrint As Collection, ByRef cllUseType As Collection, _
    ByRef cllRegistNos As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动进行分组,以便分组打印
    '出参:cllPrint-分组打印数据()
    '     格式:array(Key,结帐IDs,结算序号s,单据号,使用类别,票据格式,是否医保接口打印,险类),"K_" & 票据格式 & "_" & 险类 & "_" & 接口打印标志 & "_"  & 结算序号
    '     cllUseType-当前选中要打印的票据的使用类别,格式:array(使用类别,票据)，“K" & 使用类别
    '返回:分组成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-04-29 12:00:30
    '说明：95543
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, strKey As String, lngRow As Long
    Dim blnVirtualPrint As Boolean, blnYb As Boolean
    Dim lng结帐ID As Long, lng结算序号 As String, intInsure As Integer, lng病人ID As Long
    Dim str结帐IDs As String, str结算序号s As String, strNos As String
    Dim intPrintFormat As Integer, strUserType As String
    Dim cllUserTypes As New Collection, strInsureIDs As String
    Dim lngTemp As Long, varData As Variant, intGeneralFromat As Integer, strGeneralUserType As String
    Dim strUseType As String, strUseTypes As String, strBillNameDemo As String
    
    On Error GoTo errHandle
    '一、如果按病人补打发票，处理规则如下:
    '1.如果医保和普通病人使用的相同发票(不分使用类别)且同一种发票格式，同时医保接口不打印发票，则不分医保和普通病人，一起打印.
    '2.如果医保和普通病人使用不同发票(分使用类别)或不同发票格式，同时医保接口不打印，则需要分医保和普通病人，分别打印.
    '3.如果医保接口打印，则还是根据接口返回的单据来分组，确定打印次数(接口打印的放在一起，接口不打印的放在一起)
    '4.按病人补打发票时，分单据打印将失效!
    '二、不按病人打印发票，则分结算次数进行打印
    Set cllPrint = New Collection
    Set cllUseType = New Collection
    Set cllRegistNos = New Collection
    
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    
    '普通格式
    strGeneralUserType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
 
    intGeneralFromat = zl_GetInvoicePrintFormat(mlngModule, strGeneralUserType, , mbln不分结算次数)
    
    Set cllUserTypes = New Collection
    strInsureIDs = "": strUseTypes = ""
    With vsList
        For lngRow = 1 To .Rows - 1
            If .IsSubtotal(lngRow) Then
                '分组行不处理
            ElseIf GetVsGridBoolColVal(vsList, lngRow, .ColIndex("选择")) Then
                strNo = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
                If .TextMatrix(lngRow, .ColIndex("单据")) = "挂号单" Then
                    cllRegistNos.Add strNo
                Else
                    lng结算序号 = Val(.TextMatrix(lngRow, .ColIndex("结算序号ID")))
                    lng结帐ID = Val(.TextMatrix(lngRow, .ColIndex("结帐ID")))
                    blnYb = .TextMatrix(lngRow, .ColIndex("医保")) = "√"
                    intInsure = .TextMatrix(lngRow, .ColIndex("险类ID"))
                    
                    blnVirtualPrint = False
                    If intInsure <> 0 Then  'InStr(strInsureIDs & ",", "," & intInsure & ",") = 0 And
                        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
                    End If
                    
                    '判断使用类别
                    If InStr(strInsureIDs & ",", "," & intInsure & ",") = 0 Then
                        strInsureIDs = strInsureIDs & "," & intInsure
                        strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
                        intPrintFormat = zl_GetInvoicePrintFormat(mlngModule, strUseType, , mbln不分结算次数)
                        If mblnStartFactUseType = False Then strUseType = ""
                        cllUserTypes.Add Array(strUseType, intPrintFormat), "K" & intInsure
                    Else
                        strUseType = cllUserTypes("K" & intInsure)(0)
                        intPrintFormat = cllUserTypes("K" & intInsure)(1)
                    End If
                    
                    '获取使用类别
                    If InStr(1, strUseTypes & ",", "," & IIf(strUseType = "", "-", strUseType) & ",") = 0 Then
                        strBillNameDemo = ZlGetBillFormat(mlngModule, intPrintFormat)
                        cllUseType.Add Array(strUseType, strBillNameDemo), "K" & strUseType
                        strUseTypes = strUseTypes & "," & IIf(strUseType = "", "-", strUseType)
                    End If
                    
                    lngTemp = IIf(mbln不分结算次数, 0, lng结算序号)
                
                    '104391
                    '1.如果医保和普通病人使用的相同发票(不分使用类别)且同一种发票格式，同时医保接口不打印发票，则不分医保和普通病人，一起打印.
                    If Not blnVirtualPrint And Not mblnStartFactUseType And intPrintFormat = intGeneralFromat And intInsure <> 0 Then
                        '一起打印:1.不是医保接口打印
                        '         2.不分使用类别，且与普通票据是一种格式
                        intInsure = 0
                    End If
                    
                    'Key:"K_" & 票据格式 & "_" & 险类 & "_" & 接口打印标志 & "_"  & 结算序号
                    strKey = "K_" & intPrintFormat & "_" & intInsure & "_" & IIf(blnVirtualPrint, 1, 0) & "_" & lngTemp
                    'array(Key,结帐IDs,结算序号s,单据号,使用类别,票据格式,是否医保接口打印,险类)
                    If FindCllKeyIsExsits(cllPrint, strKey) Then
                        varData = cllPrint(strKey)
                        str结帐IDs = varData(1) & "," & lng结帐ID
                        str结算序号s = varData(1) & "," & lng结算序号
                        strNos = varData(3) & "," & strNo
                        cllPrint.Remove strKey
                    Else
                        str结帐IDs = lng结帐ID
                        str结算序号s = lng结算序号
                        strNos = strNo
                    End If
                    cllPrint.Add Array(strKey, str结帐IDs, str结算序号s, strNos, strUseType, intPrintFormat, IIf(blnVirtualPrint, 1, 0), intInsure), strKey
                End If
            End If
        Next
    End With

    SplitGroupPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function FindCllKeyIsExsits(ByVal cllData As Collection, ByVal strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找集合中的Key值是否存在，存在返回true,否则返回False
    '入参:cllData-集合数据
    '     strKey-查找的Key值
    '返回:如果Key存在，返回True,否则返回False
    '编制:刘兴洪
    '日期:2016-05-03 10:57:45
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrData As Variant
    Err = 0: On Error Resume Next
    arrData = cllData(strKey)
    If Err <> 0 Then Err = 0: Exit Function
    FindCllKeyIsExsits = True
    Exit Function
End Function

Private Sub cmdOK_Click()
    Dim cllPrint As Collection, arrPrint As Variant
    Dim strNos As String, str结帐IDs As String, lng病人ID As Long
    Dim i As Long, j As Long, strPrintUserType As String
    Dim cllUseType As Collection, strFactNO As String
    Dim strUseType As String, strBillNameDemo As String, lngShareUseID As Long, lng领用ID As Long
    Dim strPriceGrade As String
    Dim cllRegistNos As Collection
    Dim blnPrintSccess As Boolean
    
    On Error GoTo errHandler
    '1.票据打印的合法性检查
    If Not CheckPrintValied Then Exit Sub
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
      
    '2.分解票据打印内容
    '   格式:array(Key,结帐IDs,结算序号s,单据号,使用类别,票据格式,是否医保接口打印,险类),"K_" & 票据格式 & "_" & 险类 & "_" & 接口打印标志 & "_"  & 结算序号
    If SplitGroupPrint(cllPrint, cllUseType, cllRegistNos) = False Then Exit Sub

    If cllPrint.Count = 0 And cllRegistNos.Count = 0 Then
        MsgBox "未选择需要打印的票据,请选择后再补打票据", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    '获取价格等级
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lng病人ID, 0, "", , , strPriceGrade)
    Else
        strPriceGrade = gstr普通价格等级
    End If
    
    '3.进行相关的收费单票据打印
    If cllPrint.Count > 0 Then
        For j = 1 To cllUseType.Count
            strUseType = cllUseType(j)(0)
            strBillNameDemo = cllUseType(j)(1)
            '确定共用批次
            lngShareUseID = zl_GetInvoiceShareID(mlngModule, strUseType)
            If InputFactNo(strUseType, strBillNameDemo, lng领用ID, lngShareUseID, strFactNO) = False Then GoTo PrintEnd:
            
            For i = 1 To cllPrint.Count
                'array(Key,结帐IDs,结算序号s,单据号,使用类别,票据格式,是否医保接口打印,险类)
                arrPrint = cllPrint(i)
                If arrPrint(4) = strUseType Then
                    strNos = strNos & "," & arrPrint(3)
                    str结帐IDs = str结帐IDs & "," & arrPrint(1)
                    '获取票据
                    If Not zlMakeupPrint(lng病人ID, arrPrint(3), strUseType, strBillNameDemo, Val(arrPrint(5)), _
                        IIf(Val(arrPrint(6)) = 1, True, False), Val(arrPrint(7)), lng领用ID, lngShareUseID, _
                        strFactNO, str结帐IDs, strPriceGrade) Then GoTo PrintEnd:
                    blnPrintSccess = True
                    strFactNO = GetNextFactNo(arrPrint(4), lng领用ID, lngShareUseID)
                    txtInvoice.Text = strFactNO
                End If
            Next
        Next
        If str结帐IDs <> "" Then str结帐IDs = Mid(str结帐IDs, 2)
        If strNos <> "" Then strNos = Mid(strNos, 2)
          
        '银医一卡通写卡，85950
        Call WriteInforToCard(Me, mlngModule, mstrPrivs, mSquareCard.objSquareCard, 0, strNos)
        
        '81688:李南春,2015/5/18,评价器
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.OutPatiInvoicePrintAfter(lng病人ID, str结帐IDs)
            Err.Clear
        End If
    End If
    
    If cllRegistNos.Count > 0 Then
        If PrintRegistBill(cllRegistNos, lng病人ID, blnPrintSccess) = False Then GoTo PrintEnd:
    End If
    GoTo PrintEnd:
    
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
PrintEnd:
    If blnPrintSccess Then Call ReadListData '刷新数据
End Sub

Private Function PrintRegistBill(ByVal cllRegistNos As Collection, ByVal lng病人ID As Long, _
    ByRef blnPrinted As Boolean) As Boolean
    '补打挂号单据
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
    Dim i As Long, blnFirstNO As Boolean
    
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If gobjRegist Is Nothing Then Exit Function
    End If
    
    Err.Clear: On Error GoTo 0
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    
    For i = 1 To cllRegistNos.Count
        'Public Function PrintRegistBill(frmMain As Object, cnMain As ADODB.Connection, _
         ByVal lngSys As Long, ByVal strDbUser As String, _
         ByVal strNO As String, ByVal lng病人ID As Long, _
         Optional ByVal blnFirstNO As Boolean) As Boolean
        blnFirstNO = (i = 1)
        If gobjRegist.PrintRegistBill(Me, gcnOracle, glngSys, gstrDBUser, cllRegistNos(i), lng病人ID, blnFirstNO) = False Then
             Call GlobalDeleteAtom(intAtom)
             Exit Function
        End If
        blnPrinted = True
    Next
    
    Call GlobalDeleteAtom(intAtom)
    PrintRegistBill = True
End Function

Private Function InputFactNo(ByVal strUseType As String, ByVal strBillNameDemo As String, ByRef lng领用ID As Long, ByRef lngShareUseID As Long, ByRef strFactNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入有效的发票号
    '入参:strUseType-使用类别
    '     strBillNameDemo-票据名称说是
    '     lng领用ID-当前的领用ID
    '     lngShareUseID-共用批次ID
    '出参:返回的发票号
    '返回:输入成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-06-08 11:00:28
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean
    On Error GoTo errHandle
    
    If Not mblnStartFactUseType Then
        '不启用使用类型时，直接从主界面中录入的发票号中取数
        strFactNO = Trim(txtInvoice.Text)
        If strFactNO = "" Then GoTo ReInput:
        If gblnStrictCtrl Then
            If Not zlCheckInvoiceValied(lng领用ID, 1, strFactNO, lngShareUseID, strUseType) Then Exit Function
        End If
        InputFactNo = True
        Exit Function
    End If
    
ReInput:
    Do
        '根据票据领用读取
        blnValid = False
        '确定共用批次
        strFactNO = GetNextFactNo(strUseType, lng领用ID, lngShareUseID)
        
        If frmInputBox.InputBox(Me, "发票号输入:" & IIf(strUseType = "", "", "『" & strUseType & "』，格式:" & strBillNameDemo), "请确认补打使用的开始票据号码：", 30, 1, False, False, strFactNO, _
        False, Me.Left + 1500, Me.Top + 1500) = False Then Exit Function
        '用户取消输入,不打印
        If strFactNO = "" Then Exit Function
        If gblnStrictCtrl Then
            If zlCheckInvoiceValied(lng领用ID, 1, strFactNO, lngShareUseID, strUseType) Then blnValid = True
        Else
            blnValid = True
        End If
    Loop While Not blnValid
    
    InputFactNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdSelAll_Click()
    With vsList
        If .Rows <= .FixedRows Then Exit Sub
        If .ColIndex("选择") < 0 Then Exit Sub
        .Cell(flexcpChecked, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = -1
    End With
    Call SetBlanceShow
    Call InitPatiInsure
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True: Exit Sub
    If Action = PaneActionFloating Then Cancel = True: Exit Sub
    If Action = PaneActionPinning Then Cancel = True: Exit Sub
    If Action = PaneActionCollapsing Then Cancel = True: Exit Sub
    If Action = PaneActionAttaching Then Cancel = True: Exit Sub
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pane_Search    '1
        Item.Handle = picCon.hWnd
    Case Pane_List      ' 2
        Item.Handle = picList.hWnd
    Case Pane_Detail    '3
        Item.Handle = PicDetail.hWnd
    Case Pane_Balance  ' 4
        Item.Handle = picBalance.hWnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
     Bottom = stbThis.Height + picDown.Height
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    Call zlClearPatiInfor
    Call ReadListData: Call ShowDetail '处理一下界面
    mblnFirst = False
    If mlng病人ID <> 0 Then
        If GetPatient(IDKind.GetCurCard, "-" & mlng病人ID, False) = False Then
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            zlControl.TxtSelAll txtPatient
            Exit Sub
        End If
         vsList.SetFocus: Exit Sub
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdSelAll.Visible And cmdSelAll.Enabled Then Call cmdSelAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdClear.Visible And cmdClear.Enabled Then Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Dim intPrintDays As Integer
    
    mblnFirst = True
    mblnStartFactUseType = zlStartFactUseType(1)
    mlng领用ID = 0
    lblFormat.Alignment = 0

    dtpBegin.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    dtpEnd.MaxDate = dtpBegin.MaxDate
    intPrintDays = Val(zlDatabase.GetPara("缺省发票打印天数", glngSys, mlngModule, "0"))
    If intPrintDays <= 0 Then
        chkDate.Value = vbChecked
        intPrintDays = 7
    Else
        chkDate.Value = vbUnchecked
    End If
    dtpBegin.Enabled = (chkDate.Value = vbUnchecked): dtpEnd.Enabled = dtpBegin.Enabled
    dtpBegin.Value = Format(DateAdd("d", -1 * (intPrintDays - 1), dtpBegin.MaxDate), "yyyy-mm-dd")
    dtpEnd.Value = Format(dtpEnd.MaxDate, "yyyy-mm-dd")
    chkRegistFee.Value = Val(zlDatabase.GetPara("按病人补打票据含挂号费", glngSys, mlngModule, "0"))
    
    '未启用使用类别时，才能在主界面中显示
    txtInvoice.Visible = Not mblnStartFactUseType
    lblFact.Visible = Not mblnStartFactUseType
    lblFormat.Visible = Not mblnStartFactUseType
 
    Call InitPanel
    Call zlCardSquareObject
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picDown
        .Top = ScaleHeight - stbThis.Height - .Height
        .Width = ScaleWidth
        .Left = ScaleLeft
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlDatabase.SetPara "按病人补打票据含挂号费", chkRegistFee.Value, glngSys, mlngModule
    Call zlCardSquareObject(True)
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "结算列表", False
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "单据列表", False
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If txtPatient.Locked Then Exit Sub
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If mSquareCard.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub picDown_Resize()
    With picDown
        cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.Width - 50
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        vsList.Left = .ScaleLeft
        vsList.Width = .ScaleWidth
        vsList.Height = .ScaleHeight
        vsList.Top = .ScaleTop
    End With
End Sub

Private Sub PicDetail_Resize()
    Err = 0: On Error Resume Next
    With PicDetail
        vsDetail.Left = .ScaleLeft
        vsDetail.Width = .ScaleWidth
        vsDetail.Height = .ScaleHeight
        vsDetail.Top = .ScaleTop
    End With
End Sub

Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Width = .ScaleWidth
        vsBalance.Height = .ScaleHeight - lbl合计.Height - 50
        vsBalance.Top = .ScaleTop
        lbl合计.Top = .ScaleHeight - lbl合计.Height - 10
        lbl合计.Left = .ScaleLeft
    End With
End Sub

Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    If mSquareCard Is Nothing Then
         Set mSquareCard = New SquareCard
    End If
    '只有:执行或退费时,才可能管结算卡的
    If blnClosed Then
       If Not mSquareCard.objSquareCard Is Nothing Then
            Call mSquareCard.objSquareCard.CloseWindows
            Set mSquareCard.objSquareCard = Nothing
        End If
        Set mSquareCard = Nothing
        Exit Sub
    End If
    
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    Set mSquareCard.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    If mSquareCard.objSquareCard Is Nothing Then Exit Sub
    Dim objCard As Card
    
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKind.Cards.按缺省卡查找
       
   '安装了结算卡的部件
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   '功能:zlInitComponents (初始化接口部件)
   '    ByVal frmMain As Object, _
   '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
   '        ByVal cnOracle As ADODB.Connection, _
   '        Optional blnDeviceSet As Boolean = False, _
   '        Optional strExpand As String
   '出参:
   '返回:   True:调用成功,False:调用失败
   '编制:刘兴洪
   '日期:2009-12-15 15:16:22
   'HIS调用说明.
   '   1.进入门诊收费时调用本接口
   '   2.进入住院结帐时调用本接口
   '   3.进入预交款时
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   If mSquareCard.objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
        '初始部件不成功,则作为不存在处理
        Exit Sub
   End If
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '出参: blnOutMsg-已经提示,不用再外部再提示
    '返回:
    '编制:刘兴洪
    '日期:2011-01-25 16:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = _
        "Select A.病人ID,Nvl(B.主页ID,0) as 主页ID,A.门诊号 as 门诊号,A.当前床号,B.出院病床,A.姓名,A.性别,Nvl(B.年龄,A.年龄) as 年龄,A.IC卡号,A.就诊卡号,A.卡验证码," & _
        "       Nvl(B.费别,A.费别) as 费别,C.名称 as 当前科室,A.当前科室ID,D.名称 as 出院科室,B.出院科室ID, A.险类 as 险类,E.卡号,E.医保号,E.密码," & _
        "       A.登记时间,Nvl(B.状态,0) as 状态,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Nvl(B.审核标志,0) as 审核标志,B.入院日期,B.出院日期,B.病人性质,NVL(A.病人类型,B.病人类型) as 病人类型" & _
        " From 病人信息 A,病案主页 B,部门表 C,部门表 D,医保病人档案 E,医保病人关联表 F" & _
        " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+) And A.主页ID=B.主页ID(+) " & _
        "           And A.病人ID=F.病人ID(+) And F.标志(+)=1 And F.医保号=E.医保号(+) And F.险类=E.险类(+) And F.中心 = E.中心(+)" & _
        "           And A.当前科室ID=C.ID(+) And B.出院科室ID=D.ID(+)" & _
        "           And A.停用时间 is NULL "
    
    If blnCard = True And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If mSquareCard.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strSQL = strSQL & " And A.住院号=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If mrsInfo.State = 1 Then
                    If Not mrsInfo.EOF Then
                        If mrsInfo!姓名 = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                    End If
                End If
              strSQL = strSQL & " And A.姓名=[2]"
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[2]"
                '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.住院号=[2]"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If mSquareCard.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If mSquareCard.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
            End Select
    End If
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If Not mrsInfo.EOF Then
        '75259：李南春,2014-7-10，病人姓名的显示颜色处理
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), Me.ForeColor, vbRed))
        txtPatient.Text = Nvl(mrsInfo!姓名)
        'txtOld.Text = Nvl(mrsInfo!年龄): txtSex.Text = Nvl(mrsInfo!性别)
        ' txt住院号.Text = Nvl(mrsInfo!门诊号)
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        GetPatient = True
        Exit Function
    Else
        Call zlClearPatiInfor
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
NotFoundPati:
    Call zlClearPatiInfor
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Function

Private Sub zlClearPatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检除病人信息
    '编制:刘兴洪
    '日期:2011-09-04 18:09:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtPatient.Text = ""

    Set mrsInfo = New ADODB.Recordset
    vsList.Clear 1: vsList.Rows = 1: vsDetail.Clear 1: vsDetail.Rows = 1
    vsBalance.Clear 1: vsBalance.Rows = 1
    lbl合计.Caption = "补打合计:" & Format(0, "0.00")
    
    Set mrsList = Nothing: Set mrsDetail = Nothing: Set mrsBalance = Nothing
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    '问题:51488
    If (IDKind.Cards.读卡快键 = "空格键" Or IDKind.Cards.读卡快键 = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
        
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
    Else
        If IDKind.GetCurCard.名称 Like "姓名*" Then
            '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
            If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
                blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
            End If
        ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
         Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-09-03 09:32:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim strSQL As String, curTotal As Currency, blnIDCard As Boolean
    Dim blnICCard As Boolean, blnMsg As Boolean
    
    If objCard.名称 Like "IC卡*" And objCard.系统 And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        If blnCard Then
            If Not blnMsg Then MsgBox "不能确定病人信息，请检查是否正确刷卡！", vbInformation, gstrSysName
            Call zlClearPatiInfor
            Exit Sub
        End If
        If Not blnMsg Then MsgBox "不能读取病人信息！", vbInformation, gstrSysName
        Call zlClearPatiInfor
        Exit Sub
    End If
    '读取成功
    '就诊卡密码检查
    If (objCard.名称 Like "IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 = True And blnCard Then blnCard = False
    If Mid(gstrCardPass, 3, 1) = "1" And (blnCard Or blnICCard Or blnIDCard Or IDKind.GetCurCard.接口序号 <> 0) And mstrPassWord <> "" Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            Call zlClearPatiInfor
             txtPatient.SetFocus: Exit Sub
        End If
    End If
    Call ReadListData
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Function zlGetFpToBIllNOs(ByVal strFpNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的发票号,找出对应的单据号
    '返回:返回对应的单据号,用逗号分隔
    '编制:刘兴洪
    '日期:2011-02-25 10:50:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strNos As String
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct NO From 票据打印内容 A,票据使用明细 B " & _
    "   Where A.数据性质=1 and A.ID=B.打印ID and B.票种=1 And B.号码=[1]  " & _
    "   Order by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFpNo)
    strNos = ""
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    zlGetFpToBIllNOs = strNos
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub ShowDetail(Optional ByVal byt记录性质 As Byte = 1, Optional ByVal strNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细数据
    '参数:
    '编制:刘兴洪
    '日期:2011-09-04 20:18:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, strSQL As String
    Dim lngCol As Long
    
    On Error GoTo errHandler
    strSQL = _
    " Select C.名称 as 类别,Nvl(E.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
            IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
    "       Trim(To_Char(Avg(Nvl(A.付数,1)*A.数次)" & _
            IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000')) as 数量, " & _
    "       Trim(To_Char(Sum(A.标准单价)" & _
            IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "')) as 单价, " & _
    "       Trim(To_Char(Sum(A.应收金额),'9999999" & gstrDec & "')) as 应收金额, " & _
    "       Trim(To_Char(Sum(A.实收金额),'9999999" & gstrDec & "')) as 实收金额, " & _
    "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型," & _
    "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行',9,'异常收费单','第'||ABS(A.执行状态)||'次退费') as 说明" & _
    " From  门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,药品规格 X" & _
              IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
    " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
    "       And A.记录性质=[1] And A.NO=[2] And A.记录状态 IN(1,3)" & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
    " Group by Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称)," & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 ,", "") & " B.规格,A.计算单位,A.费别,D.名称," & _
    "       Nvl(A.费用类型,B.费用类型),A.执行状态,A.记录状态,X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1)" & _
    " Order by Nvl(A.价格父号,A.序号)"
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, byt记录性质, strNo)
    With vsDetail
        .Clear 1
        .Redraw = flexRDNone
        Set .DataSource = mrsDetail
        For lngCol = 0 To .COLS - 1
            .ColAlignment(lngCol) = flexAlignLeftCenter
            .FixedAlignment(lngCol) = flexAlignCenterCenter
            .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
            If .ColKey(lngCol) Like "*ID" Or InStr(1, ",符号", "," & .ColKey(lngCol) & ",") > 0 Then
                .ColHidden(lngCol) = True
            ElseIf .ColKey(lngCol) Like "*数*" Or .ColKey(lngCol) Like "*价*" Or .ColKey(lngCol) Like "*额" Then
                .ColAlignment(lngCol) = flexAlignRightCenter
            End If
        Next
        .ColHidden(.ColIndex("单位")) = byt记录性质 = 4 '挂号单隐藏“单位”列
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "单据明细列表", False
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitBlanceData(ByVal str费用结算 As String, ByVal str挂号结算 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算数据
    '入参:str费用结算-费用单据，指定的结算序号，格式：结算序号,结算序号,...
    '     str挂号结算-挂号单据，指定的结帐ID，格式：结帐ID,结帐ID,...
    '返回:
    '编制:刘兴洪
    '日期:2011-09-04 21:32:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strSubSql As String
    
    Err = 0: On Error GoTo errHandle
    If str费用结算 = "" And str挂号结算 = "" Then
        Set mrsBalance = Nothing
        InitBlanceData = True: Exit Function
    End If
    
    If str费用结算 <> "" Then
        If zlStr.ActualLen(str费用结算) <= 4000 Then
            strSubSql = "Select Column_Value As 结算序号 From Table(f_Str2list([1]))"
        Else
            strSubSql = FromStrListBulidSQL(str费用结算, "结算序号")
        End If
        strSQL = _
            "Select /*+cardinality(c,10)*/'收费单' As 单据,Min(a.No)||Decode(Min(a.No),Max(a.No),'','～'||Max(a.No)) As NO, b.结算序号 As 结算序号" & vbNewLine & _
            "From 门诊费用记录 A, 病人预交记录 B,(" & strSubSql & ") C" & vbNewLine & _
            "Where a.结帐id = b.结帐id And Mod(a.记录性质,10)=1 And (b.结算序号 = c.结算序号 Or b.结帐id = c.结算序号)" & vbNewLine & _
            "Group By b.结算序号" & vbNewLine
    End If
    If str挂号结算 <> "" Then
        If strSQL <> "" Then strSQL = strSQL & "Union All"
        If zlStr.ActualLen(str挂号结算) <= 4000 Then
            strSubSql = "Select Column_Value As 结帐id From Table(f_Str2list([2]))"
        Else
            strSubSql = FromStrListBulidSQL(str挂号结算, "结帐id")
        End If
        strSQL = strSQL & vbNewLine & _
            "Select /*+cardinality(c,10)*/'挂号单' As 单据,Min(a.No)||Decode(Min(a.No),Max(a.No),'','～'||Max(a.No)) As NO, b.结帐id As 结算序号" & vbNewLine & _
            "From 门诊费用记录 A, 病人预交记录 B,(" & strSubSql & ") C" & vbNewLine & _
            "Where a.结帐id = b.结帐id And a.记录性质=4 And b.结帐id = c.结帐id" & vbNewLine & _
            "Group By b.结帐id"
    End If

    '结算信息
    strSQL = _
        " Select Max(单据) as 单据, Max(t.No) As NO," & vbNewLine & _
        "       Decode(Mod(s.记录性质, 10), 1, '冲预交', s.结算方式) As 结算方式, Sum(s.冲预交) As 金额, t.结算序号" & vbNewLine & _
        " From 病人预交记录 S, (" & strSQL & ") T" & vbNewLine & _
        " Where s.结算序号 = t.结算序号 Or s.结帐id = t.结算序号" & vbNewLine & _
        " Group By t.单据, t.结算序号, Decode(Mod(s.记录性质, 10), 1, '冲预交', s.结算方式)"
            
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str费用结算, str挂号结算)
    InitBlanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FromStrListBulidSQL(ByVal strData As String, _
    Optional ByVal strColumnName As String, _
    Optional ByVal strSplit As String = ",") As String
    '功能：获取长字符串列表的SQL,字符串长度超过4000时
    Dim strSQL As String
    Dim varData As Variant, i As Long, strTemp As String
    
    On Error GoTo errHandler
    varData = Split(strData, strSplit)
    For i = 0 To UBound(varData)
        If zlStr.ActualLen(strTemp) > 4000 Then
            strSQL = strSQL & _
                " Union All" & _
                " Select Column_Value" & IIf(strColumnName <> "", " As " & strColumnName, "") & _
                " From Table(f_Str2list('" & Mid(strTemp, 2) & "'))"
            strTemp = ""
        End If
        strTemp = strTemp & "," & varData(i)
    Next
    If strTemp <> "" Then
        strSQL = strSQL & _
            " Union All" & _
            " Select Column_Value" & IIf(strColumnName <> "", " As " & strColumnName, "") & _
            " From Table(f_Str2list('" & Mid(strTemp, 2) & "'))"
    End If
    If strSQL <> "" Then strSQL = Mid(strSQL, 11)
    FromStrListBulidSQL = strSQL
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetBlanceShow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示结算方式
    '入参:blnAllSel-选择所有的单据
    '编制:刘兴洪
    '日期:2011-02-23 14:54:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, lngRow As Long, i As Long, str结算 As String
    Dim bln全选 As Boolean, bln未选 As Boolean
    Dim strFilter As String
    Dim strSelNos As String, dblMoney As Double
    Dim str单据类型 As String, lng结算序号 As Long
    
    lbl合计.Caption = "补打合计:0.00"
    vsBalance.Clear 1: vsBalance.Rows = 1
    If mrsBalance Is Nothing Then Exit Sub
    
    With vsList
        bln全选 = True: bln未选 = True
        For lngRow = 1 To .Rows - 1
            If .IsSubtotal(lngRow) = False Then
                str单据类型 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
                If str单据类型 = "挂号单" Then
                    lng结算序号 = .TextMatrix(lngRow, .ColIndex("结帐ID"))
                Else
                    lng结算序号 = .TextMatrix(lngRow, .ColIndex("结算序号ID"))
                End If
                If GetVsGridBoolColVal(vsList, lngRow, .ColIndex("选择")) Then
                    If InStr(1, strSelNos & ",", "," & str单据类型 & ":" & lng结算序号 & ",") = 0 Then
                        strSelNos = strSelNos & "," & str单据类型 & ":" & lng结算序号
                        bln未选 = False
                        
                        If strFilter <> "" Then strFilter = strFilter & " Or "
                        strFilter = strFilter & "(单据='" & str单据类型 & "' And 结算序号=" & lng结算序号 & ")"
                    End If
                End If
                If InStr(1, strSelNos & ",", "," & str单据类型 & ":" & lng结算序号 & ",") = 0 Then bln全选 = False
            End If
        Next
    End With
    If strSelNos <> "" Then strSelNos = Mid(strSelNos, 2)
    
    '显示所有选择的单据的结算方式之和
    If bln全选 Or bln未选 Then
        mrsBalance.Filter = 0
    Else
        mrsBalance.Filter = strFilter
    End If
    mrsBalance.Sort = "单据,NO Desc,结算方式"
    
    With vsBalance
        .Redraw = flexRDNone
        Set .DataSource = mrsBalance
        
        If chkRegistFee.Value = vbChecked Then
            '分组显示
            .OutlineBar = flexOutlineBarComplete
            .Subtotal flexSTClear
            .Subtotal flexSTNone, .ColIndex("单据"), , , &H8000000F
            .OutlineCol = .ColIndex("单据号")
        End If
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then
                .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(i, 0)
            Else
                dblMoney = dblMoney + Val(.TextMatrix(i, .ColIndex("结算金额")))
                .TextMatrix(i, .ColIndex("结算金额")) = FormatEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6, , , 2)
            End If
        Next
        
        '首列合并
        .MergeCol(.ColIndex("单据号")) = True
        .MergeCells = flexMergeRestrictColumns
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Caption, "结算列表", False
        .Redraw = flexRDBuffered
        lbl合计.Caption = "补打合计:" & Format(dblMoney, "0.00")
    End With
End Sub

Private Function ReadListData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取需要明细数据
    '返回:读取成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, strTable As String
    Dim lngCol As Long, strSQL As String
    Dim strWhere As String, dtStartDate As Date, dtEndDate As Date
    Dim i As Long, lng结算序号 As Long, byt单据类型 As Byte
    Dim blnRemove As Boolean, blnVirtualPrint As Boolean
    Dim intInsure As Integer
    Dim lng结帐ID As Long, dbl剩余数量 As Double
    Dim strPreNo As String, strPreNoType As String
    Dim str费用结算 As String '结算序号，格式：结算序号,结算序号,...
    Dim str挂号结算 As String '结帐ID，格式：结帐ID,结帐ID,...
    Dim j As Long
    
    dtStartDate = CDate("1901-01-01")
    dtEndDate = dtStartDate
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    If chkDate.Value = 0 Then
        strWhere = strWhere & " And A.发生时间 betWeen [2] and [3]"
        dtStartDate = CDate(Format(dtpBegin.Value, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpEnd.Value, "yyyy-mm-dd") & " 23:59:59")
    End If
    '排除进行补充结算的单据(收费结帐ID和非最后一次退费的结帐ID都在费用补充记录的收费结帐id中，但最后一次退费的结帐ID不在)
    If chkRegistFee.Value = vbChecked Then
        strWhere = strWhere & vbNewLine & _
            " And Mod(a.记录性质, 10) In (1, 4) " & vbNewLine & _
            " And Not Exists(Select 1 From 费用补充记录 Where 记录性质 = 1 " & _
                           "And (Mod(a.记录性质,10)=1 And Nvl(附加标志,0)=0 Or a.记录性质=4 And Nvl(附加标志,0)=1) And 收费结帐id = a.结帐id)"
    Else
        strWhere = strWhere & _
            " And Mod(a.记录性质, 10) = 1 " & vbNewLine & _
            " And Not Exists(Select 1 From 费用补充记录 Where 记录性质 = 1 And Nvl(附加标志, 0) = 0 And 收费结帐id = a.结帐id)"
    End If
    strWhere = strWhere & vbNewLine & _
        " And Not Exists(Select 1 From 费用补充记录 M, 病人预交记录 N Where m.结算序号 = n.结算序号 And n.结帐id = a.结帐id)"
    
    mblnSel = False
    On Error GoTo errHandle
    zlCommFun.ShowFlash "正在读取单据数据,请稍候 ..."
    Screen.MousePointer = 11
    DoEvents
    strTable = "" & _
            " Select Mod(a.记录性质,10) As 记录性质,a.No, Max(a.实际票号) As 实际票号, Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄," & vbNewLine & _
            "        Max(Decode(a.门诊标志, 2, '', a.标识号)) As 门诊号, Max(Decode(a.门诊标志, 2, a.标识号, '')) As 住院号," & vbNewLine & _
            "        Max(a.费别) As 费别, Max(a.开单人) As 开单人, Max(a.开单部门id) As 开单部门id, Max(a.付款方式) As 付款方式," & vbNewLine & _
            "        Max(a.划价人) As 划价人," & vbNewLine & _
            "        Max(Decode(Decode(a.记录性质,4,1,a.记录性质), 1, Decode(a.记录状态, 1, a.操作员姓名, 3, a.操作员姓名, ''), '')) As 操作员姓名," & vbNewLine & _
            "        Max(Decode(Decode(a.记录性质,4,1,a.记录性质), 1, Decode(a.记录状态, 1, a.登记时间, 3, a.登记时间, Null), Null)) As 登记时间," & vbNewLine & _
            "        Sum(Decode(Decode(a.记录性质,4,1,a.记录性质), 1, Decode(a.记录状态, 1, a.应收金额, 3, a.应收金额, 0), 0)) As 应收金额," & vbNewLine & _
            "        Sum(Decode(Decode(a.记录性质,4,1,a.记录性质), 1, Decode(a.记录状态, 1, a.实收金额, 3, a.实收金额, 0), 0)) As 实收金额," & vbNewLine & _
            "        Max(Decode(Decode(a.记录性质,4,1,a.记录性质), 1, Decode(a.记录状态, 1, a.结帐id, 3, a.结帐id, 0), 0)) As 结帐id," & vbNewLine & _
            "        Sum(Nvl(a.付数, 1) * a.数次) As 剩余数量" & vbNewLine & _
            " From 门诊费用记录 A" & vbNewLine & _
            " Where a.记录状态 In (1, 2, 3) And a.病人ID=[1]" & strWhere & vbNewLine & _
            "       And Nvl(a.附加标志, 0) <> 9 And Nvl(a.费用状态, 0) <> 1" & vbNewLine & _
            " Group By Mod(a.记录性质,10),a.No"
        
    strSQL = "Select /*+ RULE */" & _
            "  Decode(a.记录性质,4,'挂号单','收费单') As 单据, -1 As 选择, Decode(Nvl(Max(t.险类), 0), 0, Null, '√') As 医保, " & _
            "  a.No As 单据号, Max(b.名称) As 开单科室, Max(a.开单人) As 开单人, Max(a.门诊号) As 门诊号," & _
            "  Max(a.住院号) As 住院号, Max(c.名称) As 医疗付款方式, Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄, Min(a.费别) As 费别," & _
            "  To_Char(Max(a.应收金额), '99999990.00') As 应收金额, To_Char(Max(a.实收金额), '99999990.00') As 实收金额, Max(a.划价人) As 划价人," & _
            "  Max(a.操作员姓名) As 操作员, To_Char(Max(a.登记时间), 'YYYY-MM-DD HH24:MI:SS') As 登记时间, a.结帐id, " & _
            "  Max(Decode(a.记录性质,4,a.结帐ID,Nvl(m.结算序号, a.结帐id))) As 结算序号id," & _
            "  Nvl(Max(t.险类), 0) As 险类id, Max(a.剩余数量) As 剩余数量" & _
            " From (" & strTable & ") A, 病人预交记录 M, 部门表 B, 医疗付款方式 C, 保险结算记录 T" & _
            " Where a.开单部门id = b.Id And a.付款方式 = c.编码(+) And a.结帐id = t.记录id(+) And t.性质(+) = 1" & _
            "       And a.结帐id = m.结帐id(+) And (b.站点 = '" & gstrNodeNo & "' Or b.站点 Is Null)" & _
            "       And a.实际票号 Is Null " & _
            "       And (a.记录性质=1 And (Nvl(t.险类,0)<>0 Or Nvl(t.险类,0)=0 And a.剩余数量<>0) " & _
            "            Or a.记录性质=4 And a.剩余数量<>0)" & _
            " Group By a.记录性质,a.No, a.结帐id" & _
            " Order By 单据,结帐id Desc,单据号"
            
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, dtStartDate, dtEndDate)
    '102113,普通病人以及非医保接口打印的单据全部退费的不显示
    With vsList
        .Redraw = flexRDNone
        Set .DataSource = mrsList
        
        For lngCol = 0 To .COLS - 1
            .ColAlignment(lngCol) = flexAlignLeftCenter
            .FixedAlignment(lngCol) = flexAlignCenterCenter
            .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
            If .ColKey(lngCol) Like "*ID" Then
                .ColHidden(lngCol) = True
            ElseIf .ColKey(lngCol) = "剩余数量" Then
                .ColHidden(lngCol) = True
            ElseIf .ColKey(lngCol) Like "*数*" Or .ColKey(lngCol) Like "*额" Then
                .ColAlignment(lngCol) = flexAlignRightCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .ColDataType(.ColIndex("选择")) = flexDTBoolean
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "单据列表", False

        For i = 1 To .Rows - 1
            If i > .Rows - 1 Then Exit For
            lng结帐ID = Val(Trim(.TextMatrix(i, .ColIndex("结帐ID"))))

            If strPreNoType <> .TextMatrix(i, .ColIndex("单据")) _
                Or strPreNo <> Trim(.TextMatrix(i, .ColIndex("单据号"))) Then
                blnVirtualPrint = False: blnRemove = False
                
                strPreNoType = .TextMatrix(i, .ColIndex("单据"))
                strPreNo = Trim(.TextMatrix(i, .ColIndex("单据号")))
                intInsure = Val(Trim(.TextMatrix(i, .ColIndex("险类Id"))))
                dbl剩余数量 = Val(Trim(.TextMatrix(i, .ColIndex("剩余数量"))))

                If intInsure <> 0 Then
                    blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
                End If

                If blnVirtualPrint = False And RoundEx(dbl剩余数量, 6) = 0 Then
                    blnRemove = True
                    .RemoveItem i
                    i = i - 1
                End If
            ElseIf blnRemove Then
                .RemoveItem i
                i = i - 1
            End If

            If blnRemove = False Then
                lng结算序号 = Val(Trim(.TextMatrix(i, .ColIndex("结算序号ID"))))
                byt单据类型 = IIf(.TextMatrix(i, .ColIndex("单据")) = "挂号单", 4, 1)
                If Not (byt单据类型 = 1 And InStr(1, str费用结算 & ",", "," & lng结算序号 & ",") > 0 _
                    Or byt单据类型 = 4 And InStr(1, str挂号结算 & ",", "," & lng结帐ID & ",") > 0) Then

                    If byt单据类型 = 1 Then
                        str费用结算 = str费用结算 & "," & lng结算序号
                    Else
                        str挂号结算 = str挂号结算 & "," & lng结帐ID
                    End If

                    '画出分隔线
                    If i > .FixedRows Then
                        .CellBorderRange i, .FixedCols, i, .COLS - 1, vbBlue, 0, 1, 0, 0, 0, 0
                    End If
                End If
            End If
        Next
        
        If chkRegistFee.Value = vbChecked Then
            '分组显示
            .OutlineBar = flexOutlineBarComplete
            .Subtotal flexSTClear
            .Subtotal flexSTNone, .ColIndex("单据"), , , &H8000000F
            .Outline .ColIndex("选择")
            .OutlineCol = .ColIndex("选择")
            For i = 1 To .Rows - 1
                .MergeRow(i) = False
                If .IsSubtotal(i) Then
                    .TextMatrix(i, .ColIndex("选择")) = "-1"
                    For j = 0 To .COLS - 1
                        If j > .ColIndex("选择") Then
                            .Cell(flexcpText, i, j) = .TextMatrix(i, 0)
                        End If
                    Next
                    .MergeRow(i) = True
                End If
            Next
            .MergeCells = flexMergeRestrictRows
        End If
        .ColHidden(.ColIndex("单据")) = True
        
        .Editable = flexEDKbdMouse
        .Redraw = flexRDBuffered
        vsList_AfterRowColChange 0, 0, .Row, .Col
    End With
    If str费用结算 <> "" Then str费用结算 = Mid(str费用结算, 2)
    If str挂号结算 <> "" Then str挂号结算 = Mid(str挂号结算, 2)
    
    If str费用结算 = "" And str挂号结算 = "" Then
        vsDetail.Clear 1: vsDetail.Rows = 1
        vsBalance.Clear 1: vsBalance.Rows = 1
    End If
    
    '加载结算方式
    Call InitBlanceData(str费用结算, str挂号结算)
    Call SetBlanceShow
    Call InitPatiInsure
    
    zlCommFun.StopFlash
    
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsList.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
   zlCommFun.StopFlash
End Function

Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "结算列表", False
End Sub

Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "结算列表", False
End Sub

Private Sub vsDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
End Sub

Private Sub vsDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
End Sub

Private Sub vsList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call SetSelect(Row)
    Call SetBlanceShow
    '根据选择单据确病人险类
    Call InitPatiInsure
End Sub

Private Sub InitPatiInsure()
    '根据选择单据确定病人险类
    Dim strNo As String, lngRow As Long
    
    mintInsure = 0
    With vsList
        For lngRow = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsList, lngRow, .ColIndex("选择")) Then
                strNo = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
                If Val(.TextMatrix(lngRow, .ColIndex("险类ID"))) <> 0 Then
                    mintInsure = Val(.TextMatrix(lngRow, .ColIndex("险类ID")))
                    Exit For
                End If
            End If
        Next
    End With
    '重新初始化病人发票信息
    Call ReInitPatiInvoice
End Sub

Private Sub SetSelect(ByVal Row As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置选择标志
    '编制:刘兴洪
    '日期:2011-09-04 22:14:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    With vsList
        '73270,冉俊明,2014-5-23,鼠标点击选择列下的复选框，报错“运行时错误13，类型不匹配”
        If Row < 0 Or .ColIndex("结算序号ID") < 0 Or .ColIndex("选择") < 0 Then Exit Sub
        
        If .IsSubtotal(Row) Then
            For i = Row + 1 To .Rows - 1
                If .IsSubtotal(i) Then Exit For
                .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
            Next
        Else
            For i = Row - 1 To 1 Step -1
                If .IsSubtotal(i) Then Exit For
                If Not (.TextMatrix(i, .ColIndex("单据")) = .TextMatrix(Row, .ColIndex("单据")) _
                    And Val(.TextMatrix(i, .ColIndex("结算序号ID"))) = Val(.TextMatrix(Row, .ColIndex("结算序号ID")))) Then Exit For
                .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
            Next
            For i = Row + 1 To .Rows - 1
                If .IsSubtotal(i) Then Exit For
                If Not (.TextMatrix(i, .ColIndex("单据")) = .TextMatrix(Row, .ColIndex("单据")) _
                    And Val(.TextMatrix(i, .ColIndex("结算序号ID"))) = Val(.TextMatrix(Row, .ColIndex("结算序号ID")))) Then Exit For
                .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
            Next
        End If
    End With
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "单据列表", False
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNo As String, byt记录性质 As Byte
    
    If OldRow = NewRow Then Exit Sub
    With vsList
        If NewRow < .FixedRows Then Exit Sub
        byt记录性质 = IIf(.TextMatrix(NewRow, .ColIndex("单据")) = "挂号单", 4, 1)
        strNo = Trim(.TextMatrix(NewRow, .ColIndex("单据号")))
    End With
    ShowDetail byt记录性质, strNo
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "单据列表", False
End Sub
 
Private Sub vsList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsList
        Select Case Col
        Case .ColIndex("选择")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '入参:blnFact-是否重新取发票号
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng病人ID As Long
    Dim intInsure As Integer
  
    If Not mrsInfo Is Nothing Then
      If mrsInfo.State = 1 Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModule, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType)
    mintPatiInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, , True)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModule, mstrUseType)
    Call ShowBillFormat
    If blnFact Then Call RefreshFact
End Sub

Private Function ShowBillFormat()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前登录的收费操作员显示它所使用收费票据格式
    '编制:刘兴洪
    '日期:2016-06-08 10:06:20
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, intFormat As Integer
    
    On Error GoTo errHandle
    If mblnStartFactUseType Then Exit Function
    
    If mbln不分结算次数 Then
        intFormat = mintPatiInvoiceFormat
    Else
        intFormat = mintInvoiceFormat
    End If
    Call ZlShowBillFormat(mlngModule, lblFormat, intFormat)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetNextFactNo(ByVal strUseType As String, ByRef lng领用ID As Long, ByRef lngShareUseID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取下一张发票号
    '入参:strUserType-使用类别
    '     lng领用ID-领用ID
    '     lngShareUseID-共用ID
    '出参:lng领用ID-领用ID
    '返回:下一张发票号
    '编制:刘兴洪
    '日期:2016-06-08 10:27:46
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gblnStrictCtrl Then
        If zlCheckInvoiceValied(lng领用ID, 1, , lngShareUseID, strUseType) = False Then Exit Function
        '严格：取下一个号码
        GetNextFactNo = GetNextBill(lng领用ID)
        Exit Function
    End If
    GetNextFactNo = zlStr.Increase(UCase(zlDatabase.GetPara("当前收费票据号", glngSys, mlngModule)))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshFact()
    '功能：刷新收费票据号
    If mblnStartFactUseType Then Exit Sub
    
    If gblnStrictCtrl Then
        'lblFact.tag主要是检查发票号是否手工输入的.手工输入的,发票号为空,否则是自动产生的发票号
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            If zlCheckInvoiceValied(mlng领用ID, 1, , mlngShareUseID, mstrUseType) = False Then
                txtInvoice.Text = "": txtInvoice.Tag = "": Exit Sub
            End If
            '严格：取下一个号码
            txtInvoice.Text = GetNextBill(mlng领用ID)
            'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
            '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
            '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
            txtInvoice.Tag = txtInvoice.Text
            lblFact.Tag = txtInvoice.Tag
            Call zlCheckFactIsEnough
        End If
    Else
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            '松散：取下一个号码
            txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("当前收费票据号", glngSys, mlngModule)))
            'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
            '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
            '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
        End If
        txtInvoice.Tag = txtInvoice.Text
        lblFact.Tag = txtInvoice.Tag
    End If
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Sub zlCheckFactIsEnough()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前票据是否允足
    '编制:刘兴洪
    '日期:2011-05-10 17:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng剩余数量 As Long
    '刘兴洪 问题:26948 日期:2009-12-28 17:43:00
    '需要检查剩余数量是否充足:
    If zlCheckInvoiceOverplusEnough(1, gTy_Module_Para.int提醒剩余票据张数, lng剩余数量, mlng领用ID, mstrUseType) = False Then
        MsgBox "注意:" & vbCrLf & _
               "    当前剩余票据(" & lng剩余数量 & ") 小于了报警的张数(" & gTy_Module_Para.int提醒剩余票据张数 & "),请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    End If
End Sub
