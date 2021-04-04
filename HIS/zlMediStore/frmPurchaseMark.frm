VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPurchaseMark 
   Caption         =   "核对发票"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   Icon            =   "frmPurchaseMark.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   9135
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   2895
      TabIndex        =   31
      Top             =   480
      Width           =   2895
      Begin VB.PictureBox picCol3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   0
         Width           =   260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "未修改"
         Height          =   180
         Left            =   2280
         TabIndex        =   37
         Top             =   37
         Width           =   540
      End
      Begin VB.Label lblNotExecute 
         AutoSize        =   -1  'True
         Caption         =   "已付款"
         Height          =   180
         Left            =   360
         TabIndex        =   36
         Top             =   37
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "已修改"
         Height          =   180
         Left            =   1320
         TabIndex        =   35
         Top             =   37
         Width           =   540
      End
   End
   Begin VB.PictureBox picDetails 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3840
      ScaleHeight     =   1335
      ScaleWidth      =   3495
      TabIndex        =   15
      Top             =   1080
      Width           =   3495
      Begin XtremeSuiteControls.TabControl tbcDetails 
         Height          =   975
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1720
         _StockProps     =   64
      End
   End
   Begin VB.Frame fra条件 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3015
      Begin VB.CommandButton cmd药品 
         Caption         =   "…"
         Height          =   300
         Left            =   2640
         TabIndex        =   30
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txt药品名称 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   28
         Top             =   1182
         Width           =   1725
      End
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   1725
      End
      Begin VB.CommandButton Cmd供应商 
         Caption         =   "…"
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.ComboBox cbo审核日期 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1623
         Width           =   1725
      End
      Begin VB.TextBox txt结束发票号 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   5
         Top             =   3420
         Width           =   1725
      End
      Begin VB.TextBox txt开始发票号 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2976
         Width           =   1725
      End
      Begin VB.TextBox txt供应商 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   741
         Width           =   1485
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   2064
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   57671683
         CurrentDate     =   40848
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   2520
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   57671683
         CurrentDate     =   40848
      End
      Begin VB.Label lbl药品 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "药品名称"
         Height          =   180
         Left            =   180
         TabIndex        =   29
         Top             =   1242
         Width           =   720
      End
      Begin VB.Label lbl库房 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "外购库房"
         Height          =   180
         Left            =   180
         TabIndex        =   23
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl开始日期 
         Caption         =   "开始日期"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblEnd发票 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "结束发票号"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label lblStart发票 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "开始发票号"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   3030
         Width           =   900
      End
      Begin VB.Label lbl结束日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "结束日期"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   2580
         Width           =   720
      End
      Begin VB.Label lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "审核日期"
         Height          =   180
         Left            =   180
         TabIndex        =   3
         Top             =   1710
         Width           =   720
      End
      Begin VB.Label lbl供应商 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "供 应 商"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   6780
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchaseMark.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9763
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseMark.frx":70E6
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseMark.frx":75E8
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   3720
      ScaleHeight     =   4095
      ScaleWidth      =   5175
      TabIndex        =   17
      Top             =   2640
      Width           =   5175
      Begin VSFlex8Ctl.VSFlexGrid vsf列头 
         Height          =   3855
         Left            =   2640
         TabIndex        =   27
         Top             =   0
         Width           =   2415
         _cx             =   4260
         _cy             =   6800
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.CheckBox chk全选 
         Caption         =   "全选"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox picSplit 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   255
         ScaleWidth      =   4935
         TabIndex        =   21
         Top             =   2760
         Width           =   4935
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf库存 
         Height          =   855
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   4455
         _cx             =   7858
         _cy             =   1508
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VSFlex8Ctl.VSFlexGrid vsf未标记 
         Height          =   975
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   5175
         _cx             =   9128
         _cy             =   1720
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         Begin VB.PictureBox picSetCols 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   220
            Index           =   0
            Left            =   0
            Picture         =   "frmPurchaseMark.frx":7AEA
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf已标记 
         Height          =   1215
         Left            =   0
         TabIndex        =   20
         Top             =   1440
         Width           =   5295
         _cx             =   9340
         _cy             =   2143
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         Begin VB.PictureBox picSetCols 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   220
            Index           =   1
            Left            =   0
            Picture         =   "frmPurchaseMark.frx":801C
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpPanel 
      Bindings        =   "frmPurchaseMark.frx":854E
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPicture 
      Left            =   1320
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPurchaseMark.frx":8562
   End
End
Attribute VB_Name = "frmPurchaseMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const menuToolSave As Integer = 101
Private Const menuToolGetData As Integer = 102
Private Const menuToolExit As Integer = 103
Private Const menuTool简洁 As Integer = 104
Private Const menuTool详细 As Integer = 105
Private Const menuTool全选 As Integer = 106
Private Const menuToolSave2 As Integer = 107
Private mobjMnu As ICommandBarControl

Private Const CSTCOLOR_UNMODIFY = &HC0C0FF       '粉红 选项页颜色
Private Const CSTCOLOR_NORECORDS = &HFFFFFF   '

Private Const mColumn As String = "付款|NO|药品名称|批号|规格|单位|数量|采购价|采购金额|发票号|发票金额|发票日期|审核人|审核日期"
Private mintShowCostDigit As Integer            '成本价小数位数
Private mintShowPriceDigit As Integer           '售价小数位数
Private mintShowNumberDigit As Integer          '数量小数位数
Private mintShowMoneyDigit As Integer           '金额小数位数
Private mintUnit As Integer '单位系数
Private mblnDo As Boolean
Private mstrSelColumn As String                 '记录选中需要显示的列
Private mstrColumn As String
Private mvMsg As VbMsgBoxResult                 '提示信息
Private mstrLike As String                      '是按那种方式进行输入匹配
Private mstrPrivs As String                     '模块权限

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4
Private mint单位系数 As Integer

Private mStr库房 As String
Private mstr当前库房 As String
Private Const MStrCaption As String = "药品外购入库管理"

Private Enum mPage
    未标记单据 = 0
    已标记单据 = 1
End Enum

Private Enum mColumnMark
    序号 = 0
    id = 1
    项目Id
    付款标志
    NO
    药品名称
    批号
    规格
    计算单位
    数量
    采购价
    采购金额
'    当前库存
'    全院库存
    付款序号
    发票号
    发票金额
    发票日期
    审核人
    审核日期
    剂量系数
    门诊包装
    住院包装
    药库包装
    count = 22
End Enum

Public Sub showMe(ByVal str库房 As String, ByVal str当前库房 As String, ByVal objFrm As frmMainList, ByVal strPrivs As String)
    '公共过程，用来供其他窗体调用本窗体，并传入相应过程
    
'    mlng库房 = lng库房
'    mStr库房 = str库房
    mStr库房 = str库房
    mstr当前库房 = str当前库房
    mstrPrivs = strPrivs
    Me.Show vbModal, objFrm
End Sub

Private Sub initComandbar()
    '初始化工具栏
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPicture.Icons
    
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    
    With cbrToolBar.Controls    'menuToolSave2
        Set cbrControlMain = .Add(xtpControlButton, menuToolSave, "标记")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, menuToolSave2, "取消标记")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, menuToolGetData, "提取数据")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, menuToolExit, "退出")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        
'        Set cbrControlMain = .Add(xtpControlButton, menuTool简洁, "简洁")
'            cbrControlMain.flags = xtpFlagRightAlign
'            cbrControlMain.Style = xtpButtonIconAndCaption '同时显示图标和文字
'        Set cbrControlMain = .Add(xtpControlButton, menuTool详细, "详细")
'            cbrControlMain.flags = xtpFlagRightAlign
'            cbrControlMain.Style = xtpButtonIconAndCaption '同时显示图标和文字
'            cbrControlMain.Checked = True
    End With
    
    cbsMain.Item(1).Delete
End Sub

Private Sub InitTabControl()
    '初始化Tabcontrol控件
    With Me.tbcDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mPage.未标记单据, "未标记单据", picList.hWnd, 0).Tag = "未标记单据_"
        .InsertItem(mPage.已标记单据, "已标记单据", picList.hWnd, 0).Tag = "已标记单据_"
        .Item(mPage.已标记单据).Selected = True
        .Item(mPage.未标记单据).Selected = True
    End With
End Sub

Private Sub cbo库房_Click()
    Call SetSelectorRS(1, "药品外购入库管理", cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex))
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    vsf列头.Visible = False
    Select Case Control.id
        Case menuTool简洁   '简洁列表
'            Call Simple(Control)
        Case menuTool详细   '详细列表
'            Call Full(Control)
        Case menuToolGetData   '提取数据
            Call checkUpdate
        Case menuToolSave, menuToolSave2  '保存
            Call Save
        Case menuToolExit
            Call ExitForm
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    picDetails.Move fra条件.Width, lngTop, Me.Width - fra条件.Width - fra条件.Left, lngBottom - staThis.Height - lngTop
    tbcDetails.Move 0, 0, picDetails.Width, picDetails.Height
    cbo库房.Move lbl库房.Left + lbl库房.Width + 280, lbl库房.Top, fra条件.Width - cbo库房.Left - 100
    txt供应商.Move cbo库房.Left, txt供应商.Top, fra条件.Width - cbo库房.Left - 100
    Cmd供应商.Left = txt供应商.Left + txt供应商.Width - Cmd供应商.Width
    txt药品名称.Move cbo库房.Left, txt药品名称.Top, fra条件.Width - cbo库房.Left - 100
    cmd药品.Left = txt供应商.Left + txt供应商.Width - Cmd供应商.Width
    cbo审核日期.Move cbo库房.Left, cbo审核日期.Top, txt药品名称.Width
    dtp开始时间.Move cbo库房.Left, dtp开始时间.Top, txt药品名称.Width
    dtp结束时间.Move cbo库房.Left, dtp结束时间.Top, txt药品名称.Width
    txt开始发票号.Move cbo库房.Left, txt开始发票号.Top, txt药品名称.Width
    txt结束发票号.Move cbo库房.Left, txt结束发票号.Top, txt药品名称.Width
    
    Call initOtherControl
End Sub

'Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    If Control.Id = menuTool简洁 Then
'        If Control.Checked Then
'            Control.IconId = 105
'        Else
'            Control.IconId = 104
'        End If
'    ElseIf Control.Id = menuTool详细 Then
'        If Control.Checked Then
'            Control.IconId = 105
'        Else
'            Control.IconId = 104
'        End If
'    End If
'End Sub

Private Sub cbo审核日期_Click()
    If cbo审核日期.Text = "自定义日期" Then
        lbl开始日期.Visible = True
        dtp开始时间.Visible = True
        lbl结束日期.Visible = True
        dtp结束时间.Visible = True
        
        lblStart发票.Top = dtp结束时间.Top + dtp结束时间.Height + 130
        txt开始发票号.Top = dtp结束时间.Top + dtp结束时间.Height + 80
        lblEnd发票.Top = txt开始发票号.Top + txt开始发票号.Height + 130
        txt结束发票号.Top = txt开始发票号.Top + txt开始发票号.Height + 80
    Else
        lbl开始日期.Visible = False
        dtp开始时间.Visible = False
        lbl结束日期.Visible = False
        dtp结束时间.Visible = False
        
        lblStart发票.Top = cbo审核日期.Top + cbo审核日期.Height + 130
        txt开始发票号.Top = cbo审核日期.Top + cbo审核日期.Height + 80
        lblEnd发票.Top = txt开始发票号.Top + txt开始发票号.Height + 130
        txt结束发票号.Top = txt开始发票号.Top + txt开始发票号.Height + 80
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If tbcDetails.Item(0).Selected = True Then
'        cbsMain(1).Controls(1).Caption = "标记"
        cbsMain.FindControl(xtpControlButton, menuToolSave).Caption = "标记"
    Else
        cbsMain.FindControl(xtpControlButton, menuToolSave).Caption = "取消标记"
    End If
End Sub

Private Sub chk全选_Click()
    Dim i As Integer
    If tbcDetails.Item(mPage.未标记单据).Selected = True Then
        With vsf未标记
            For i = 1 To .rows - 1
                If chk全选.Value = 1 Then
                    .TextMatrix(i, mColumnMark.付款标志) = "√"
                    .Cell(flexcpFontBold, i, mColumnMark.付款标志) = True
                    .Cell(flexcpFontSize, i, mColumnMark.付款标志) = 10
                    .Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlue
                Else
                    .TextMatrix(i, mColumnMark.付款标志) = ""
                End If
            Next
        End With
    ElseIf tbcDetails.Item(mPage.已标记单据).Selected = True Then
        With vsf已标记
        For i = 1 To .rows - 1
            If Trim(.TextMatrix(i, mColumnMark.付款序号)) = "未付款" Then '已经付款的单据不能修改付款标志
                If chk全选.Value = 1 Then
                    .TextMatrix(i, mColumnMark.付款标志) = "√"
                    .Cell(flexcpFontBold, i, mColumnMark.付款标志) = True
                    .Cell(flexcpFontSize, i, mColumnMark.付款标志) = 10
                    .Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlue
                Else
                    .TextMatrix(i, mColumnMark.付款标志) = ""
                End If
            End If
        Next
    End With
    End If
End Sub

Private Sub Cmd供应商_Click()
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsRecord As ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt供应商.hWnd) '获取位置
    
    gstrSQL = "select id,编码,名称,简码 from 供应商 Where 末级 = 1 and (站点 = '-' Or 站点 Is Null) And" & _
                " (Substr(类型, 1, 1) = 1 Or Nvl(末级, 0) = 0)"
                
    Set rsRecord = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "供应商", False, "", "", False, False, _
    True, vRect.Left, vRect.Top, txt供应商.Height, blnCancel, False, True)

    If rsRecord Is Nothing Then
        Exit Sub
    Else
        If txt供应商.Tag <> rsRecord!id Then
            txt药品名称.Tag = ""
            txt药品名称.Text = ""
            txt开始发票号.Text = ""
            txt结束发票号.Text = ""
        End If
        txt供应商.Text = rsRecord!名称
        txt供应商.Tag = rsRecord!id
    End If
    zlControl.TxtSelAll txt供应商
    OS.OpenIme True
End Sub

Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品外购入库管理", cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex))
    End If
    
    Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex), , , , , , False, mstrPrivs)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        txt药品名称.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        txt药品名称.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    txt药品名称.Tag = RecReturn!药品id
End Sub

Private Sub dkpPanel_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.id = 1 Then
         Item.Handle = fra条件.hWnd
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 12000
    mblnDo = Val(zlDataBase.GetPara("使用个性化风格")) <> 0
    mstrLike = IIf(Val(zlDataBase.GetPara("输入匹配")) = 0, "%", "")
    staThis.Panels(2).Picture = picColor
    
    Call initComandbar  '初始化工具栏
    Call initPanel  '初始化面板
    Call InitTabControl
    Call initOtherControl   '控制其他控件位置
    Call initColumn '初始化列
    Call initComboBox
        
    If mblnDo Then
        RestoreWinState Me, App.ProductName, MStrCaption
    End If
    
    mstrColumn = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "付款标记", "列显示隐藏", "")
    If mstrColumn <> "" And mblnDo = True Then
        Call SetColumnVisible
    End If
    
    lbl开始日期.Visible = False
    dtp开始时间.Visible = False
    lbl结束日期.Visible = False
    dtp结束时间.Visible = False
    
    Me.Caption = "核对发票"
    dtp开始时间.Value = DateAdd("d", -7, Sys.Currentdate)
    dtp结束时间.Value = Sys.Currentdate
    
    mintUnit = zlDataBase.GetPara("药品单位", glngSys, 1300, 0, 0, True)
    Select Case mintUnit
        Case 4 '售价单位
            mint单位系数 = 4
        Case 2 '门诊单位
            mint单位系数 = 2
        Case 3 '住院单位
            mint单位系数 = 3
        Case 1 '药库单位
            mint单位系数 = 1
        Case 0
            mint单位系数 = 1
    End Select
    Call GetDrugDigit(cbo库房.ItemData(cbo库房.ListIndex), MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
End Sub

Private Sub SetColumnVisible()
    Dim i As Integer
    Dim strTemp As String
    Dim arrColumn As Variant
    Dim j As Integer
    
    ReDim arrColumn(UBound(Split(mstrColumn, "|"))) As String
    For i = 0 To UBound(arrColumn) - 1
        arrColumn(i) = Split(mstrColumn, "|")(i)
    Next
    With vsf列头
        For i = 1 To .rows - 1
            For j = 0 To UBound(arrColumn) - 1
                If InStr(1, arrColumn(j), .TextMatrix(i, 2)) > 0 Then
                    .TextMatrix(i, 1) = IIf(Split(arrColumn(j), ",")(0) = "0", "", Split(arrColumn(j), ",")(0))
                End If
            Next
        Next
    End With
    
    For i = 1 To vsf列头.rows - 1
        If vsf列头.TextMatrix(i, 1) = "" Then
            vsf未标记.colHidden(vsf未标记.ColIndex(vsf列头.TextMatrix(i, 2))) = True
            vsf已标记.colHidden(vsf未标记.ColIndex(vsf列头.TextMatrix(i, 2))) = True
        End If
    Next
End Sub

Private Sub initPanel()
    '初始化分栏控件
    'DockingPane
    '-----------------------------------------------------
    Dim objPaneCon As Pane
    Dim objPaneDetail As Pane
    
    Me.dkpPanel.SetCommandBars Me.cbsMain
    Me.dkpPanel.Options.UseSplitterTracker = False '实时拖动
    Me.dkpPanel.Options.ThemedFloatingFrames = True
    Me.dkpPanel.Options.AlphaDockingContext = True
    
    Set objPaneCon = Me.dkpPanel.CreatePane(1, 200, 0, DockLeftOf, Nothing)
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
    objPaneCon.Title = "提取条件"
End Sub

Private Sub initOtherControl()
    '初始化其他控件 各个列表
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    picDetails.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - staThis.Height - lngTop
    tbcDetails.Move 0, 0, picDetails.Width, picDetails.Height
    
    If tbcDetails.Item(0).Selected = True Then
        vsf已标记.Visible = False
        vsf未标记.Visible = True
        
        chk全选.Move 0, 0, 1215, 255
        vsf未标记.Move 0, chk全选.Height, picList.Width, (picList.Height / 6) * 5
        picSplit.Move 0, vsf未标记.Top + vsf未标记.Height, picList.ScaleWidth, 50
        vsf库存.Move 0, picSplit.Top + picSplit.Height, picList.ScaleWidth, picList.Height - picSplit.Top - picSplit.Height
        
    Else
        vsf已标记.Visible = True
        vsf未标记.Visible = False
        
        chk全选.Move 0, 0, 1215, 255
        vsf已标记.Move 0, chk全选.Height, picList.ScaleWidth, (picList.ScaleHeight / 6) * 5
        picSplit.Move 0, vsf已标记.Top + vsf已标记.Height, picList.ScaleWidth, 50
        vsf库存.Move 0, picSplit.Top + picSplit.Height, picSplit.ScaleWidth, picList.Height - picSplit.Top - picSplit.Height
    End If
End Sub

Private Sub initColumn()
    Dim i As Integer
    '初始化表格中列头 未标记
    With vsf未标记
        .rows = 1
        .Cols = mColumnMark.count
        .Editable = flexEDNone
        .MergeCells = flexMergeRestrictColumns
        
        .AllowSelection = False '不能多选
        .SelectionMode = flexSelectionByRow '整行选择
        
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '不能多选单元格
        .RowHeight(0) = 310
    End With
    With vsf已标记
        .rows = 1
        .Cols = mColumnMark.count
        
        .Editable = flexEDNone
        .MergeCells = flexMergeRestrictColumns
        
        .AllowSelection = False '不能多选
        .SelectionMode = flexSelectionByRow '整行选择
        
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '不能多选单元格
        .RowHeight(0) = 310
    End With
    With vsf库存
        .rows = 0
        .Editable = flexEDNone
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '不能多选单元格
    End With
    With vsf列头
        .Cols = 3
        .ColDataType(1) = flexDTBoolean
        .Editable = flexEDNone
        .MergeCells = flexMergeRestrictAll
        .MergeRow(0) = True
        .TextMatrix(0, 0) = "未选中的列将隐藏"
        .TextMatrix(0, 1) = "未选中的列将隐藏"
        .TextMatrix(0, 2) = "未选中的列将隐藏"
        .colHidden(0) = True
        .rows = UBound(Split(mColumn, "|")) + 1
        For i = 1 To .rows - 1
            .TextMatrix(i, 1) = "1"
            .TextMatrix(i, 2) = Split(mColumn, "|")(i)
            .RowHeight(i) = 300
        Next
        .Visible = False
    End With
    
    VsfGridColFormat vsf未标记, mColumnMark.序号, "序号", 640, flexAlignCenterCenter, "序号"
    VsfGridColFormat vsf未标记, mColumnMark.id, "id", 600, flexAlignCenterCenter, "id"
    VsfGridColFormat vsf未标记, mColumnMark.项目Id, "项目id", 600, flexAlignCenterCenter, "项目id"
    VsfGridColFormat vsf未标记, mColumnMark.付款标志, "付款", 1000, flexAlignCenterCenter, "付款"
    VsfGridColFormat vsf未标记, mColumnMark.NO, "NO", 1500, flexAlignLeftCenter, "NO"
    VsfGridColFormat vsf未标记, mColumnMark.药品名称, "药品名称", 1500, flexAlignLeftCenter, "药品名称"
    
    VsfGridColFormat vsf未标记, mColumnMark.批号, "批号", 600, flexAlignLeftCenter, "批号"
    VsfGridColFormat vsf未标记, mColumnMark.规格, "规格", 1500, flexAlignLeftCenter, "规格"
    
    VsfGridColFormat vsf未标记, mColumnMark.计算单位, "单位", 600, flexAlignLeftCenter, "单位"
    VsfGridColFormat vsf未标记, mColumnMark.数量, "数量", 1000, flexAlignRightCenter, "数量"
    VsfGridColFormat vsf未标记, mColumnMark.采购价, "采购价", 1000, flexAlignRightCenter, "采购价"
    VsfGridColFormat vsf未标记, mColumnMark.采购金额, "采购金额", 1000, flexAlignRightCenter, "采购金额"
    
'    VsfGridColFormat vsf未标记, mColumnMark.当前库存, "当前库存", 1000, flexAlignRightCenter, "当前库存"
'    VsfGridColFormat vsf未标记, mColumnMark.全院库存, "全院库存", 1500, flexAlignRightCenter, "全院库存"
    VsfGridColFormat vsf未标记, mColumnMark.付款序号, "付款序号", 1000, flexAlignLeftCenter, "付款序号"
    VsfGridColFormat vsf未标记, mColumnMark.审核人, "审核人", 1000, flexAlignLeftCenter, "审核人"
    VsfGridColFormat vsf未标记, mColumnMark.审核日期, "审核日期", 1000, flexAlignLeftCenter, "审核日期"
    VsfGridColFormat vsf未标记, mColumnMark.发票号, "发票号", 1000, flexAlignLeftCenter, "发票号"
    VsfGridColFormat vsf未标记, mColumnMark.发票金额, "发票金额", 1000, flexAlignRightCenter, "发票金额"
    VsfGridColFormat vsf未标记, mColumnMark.发票日期, "发票日期", 1000, flexAlignLeftCenter, "发票日期"
    VsfGridColFormat vsf未标记, mColumnMark.剂量系数, "剂量系数", 1000, flexAlignLeftCenter, "剂量系数"
    VsfGridColFormat vsf未标记, mColumnMark.门诊包装, "门诊包装", 1000, flexAlignLeftCenter, "门诊包装"
    VsfGridColFormat vsf未标记, mColumnMark.住院包装, "住院包装", 1000, flexAlignRightCenter, "住院包装"
    VsfGridColFormat vsf未标记, mColumnMark.药库包装, "药库包装", 1000, flexAlignLeftCenter, "药库包装"
    vsf未标记.Cell(flexcpPicture, vsf未标记.Row, 0, vsf未标记.Row, 0) = picSetCols(0)
    
    '已标记
    VsfGridColFormat vsf已标记, mColumnMark.序号, "序号", 640, flexAlignCenterCenter, "序号"
    VsfGridColFormat vsf已标记, mColumnMark.id, "id", 600, flexAlignCenterCenter, "id"
    VsfGridColFormat vsf已标记, mColumnMark.项目Id, "项目id", 600, flexAlignCenterCenter, "项目id"
    VsfGridColFormat vsf已标记, mColumnMark.付款标志, "取消付款", 1000, flexAlignCenterCenter, "取消付款"
    VsfGridColFormat vsf已标记, mColumnMark.NO, "NO", 1500, flexAlignLeftCenter, "NO"
    VsfGridColFormat vsf已标记, mColumnMark.药品名称, "药品名称", 1500, flexAlignLeftCenter, "药品名称"
    VsfGridColFormat vsf已标记, mColumnMark.批号, "批号", 600, flexAlignLeftCenter, "批号"
    VsfGridColFormat vsf已标记, mColumnMark.规格, "规格", 1500, flexAlignLeftCenter, "规格"
    VsfGridColFormat vsf已标记, mColumnMark.计算单位, "单位", 1000, flexAlignLeftCenter, "单位"
    VsfGridColFormat vsf已标记, mColumnMark.数量, "数量", 1000, flexAlignRightCenter, "数量"
    VsfGridColFormat vsf已标记, mColumnMark.采购价, "采购价", 1000, flexAlignRightCenter, "采购价"
    VsfGridColFormat vsf已标记, mColumnMark.采购金额, "采购金额", 1000, flexAlignRightCenter, "采购金额"
'    VsfGridColFormat vsf已标记, mColumnMark.当前库存, "当前库存", 1000, flexAlignRightCenter, "当前库存"
'    VsfGridColFormat vsf已标记, mColumnMark.全院库存, "全院库存", 1500, flexAlignRightCenter, "全院库存"
    VsfGridColFormat vsf已标记, mColumnMark.付款序号, "付款序号", 1000, flexAlignLeftCenter, "付款序号"
    VsfGridColFormat vsf已标记, mColumnMark.审核人, "审核人", 1000, flexAlignLeftCenter, "审核人"
    VsfGridColFormat vsf已标记, mColumnMark.审核日期, "审核日期", 1000, flexAlignLeftCenter, "审核日期"
    VsfGridColFormat vsf已标记, mColumnMark.发票号, "发票号", 1000, flexAlignLeftCenter, "发票号"
    VsfGridColFormat vsf已标记, mColumnMark.发票金额, "发票金额", 1000, flexAlignRightCenter, "发票金额"
    VsfGridColFormat vsf已标记, mColumnMark.发票日期, "发票日期", 1000, flexAlignLeftCenter, "发票日期"
    VsfGridColFormat vsf已标记, mColumnMark.剂量系数, "剂量系数", 1000, flexAlignLeftCenter, "剂量系数"
    VsfGridColFormat vsf已标记, mColumnMark.门诊包装, "门诊包装", 1000, flexAlignLeftCenter, "门诊包装"
    VsfGridColFormat vsf已标记, mColumnMark.住院包装, "住院包装", 1000, flexAlignRightCenter, "住院包装"
    VsfGridColFormat vsf已标记, mColumnMark.药库包装, "药库包装", 1000, flexAlignLeftCenter, "药库包装"
    vsf已标记.Cell(flexcpPicture, vsf已标记.Row, 0, vsf已标记.Row, 0) = picSetCols(1)
    
    vsf未标记.colHidden(mColumnMark.id) = True
    vsf已标记.colHidden(mColumnMark.id) = True
    vsf未标记.colHidden(mColumnMark.项目Id) = True
    vsf已标记.colHidden(mColumnMark.项目Id) = True
    vsf未标记.colHidden(mColumnMark.付款序号) = True
    vsf已标记.colHidden(mColumnMark.付款序号) = True
    
    vsf未标记.colHidden(mColumnMark.剂量系数) = True
    vsf未标记.colHidden(mColumnMark.门诊包装) = True
    vsf未标记.colHidden(mColumnMark.住院包装) = True
    vsf未标记.colHidden(mColumnMark.药库包装) = True
    
    vsf已标记.colHidden(mColumnMark.剂量系数) = True
    vsf已标记.colHidden(mColumnMark.门诊包装) = True
    vsf已标记.colHidden(mColumnMark.住院包装) = True
    vsf已标记.colHidden(mColumnMark.药库包装) = True
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐）
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Private Sub initComboBox()
    Dim i As Integer
    Dim strTemp As String
    Dim strIndex As String
    Dim arrtemp As Variant
    
    With cbo审核日期
        .Clear
        .AddItem "今日", "0"
        .AddItem "一星期内", "1"
        .AddItem "一个月内", "2"
        .AddItem "三个月内", "3"
        .AddItem "自定义日期", "4"
        .ListIndex = 0
    End With
    
    ReDim arrtemp(UBound(Split(mStr库房, "|"))) As String
    
    With cbo库房
        .Clear
        For i = 0 To UBound(arrtemp) - 1
            strIndex = ""
            strTemp = ""
            arrtemp(i) = Split(mStr库房, "|")(i)
            strIndex = Mid(arrtemp(i), 1, InStr(1, arrtemp(i), ",") - 1)
            strTemp = Mid(arrtemp(i), InStr(1, arrtemp(i), ",") + 1)
            .AddItem strTemp
            .ItemData(.NewIndex) = strIndex
        Next
        .ListIndex = mstr当前库房
    End With
End Sub

'Private Sub Simple(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    '简单模式
'    If Control.Checked = False Then
'        Control.Checked = True
'        cbsMain.Item(1).Controls.Item(5).Checked = False
'    End If
'
'    If cbsMain.Item(1).Controls.Item(4).Checked = True Then '简单模式被选中的话
'        With vsf未标记
'            .ColHidden(mColumnMark.发票号) = True
'            .ColHidden(mColumnMark.发票金额) = True
'            .ColHidden(mColumnMark.发票日期) = True
'        End With
'        With vsf已标记
'            .ColHidden(mColumnMark.发票号) = True
'            .ColHidden(mColumnMark.发票金额) = True
'            .ColHidden(mColumnMark.发票日期) = True
'        End With
'    End If
'End Sub

'Private Sub Full(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    '完整模式
'    If Control.Checked = False Then
'        Control.Checked = True
'        cbsMain.Item(1).Controls.Item(4).Checked = False
'    End If
'
'    If cbsMain.Item(1).Controls.Item(5).Checked = True Then '简单完整模式被选中的话
'        With vsf未标记
'            .ColHidden(mColumnMark.发票号) = False
'            .ColHidden(mColumnMark.发票金额) = False
'            .ColHidden(mColumnMark.发票日期) = False
'        End With
'
'        With vsf已标记
'            .ColHidden(mColumnMark.发票号) = False
'            .ColHidden(mColumnMark.发票金额) = False
'            .ColHidden(mColumnMark.发票日期) = False
'        End With
'    End If
'End Sub

Private Sub checkUpdate()
    '检查是否修改了记录
    Dim i As Integer
    Dim blnChange As Boolean
    Dim lngResult As Long
    
    blnChange = False
    For i = 1 To vsf未标记.rows - 1
        If vsf未标记.Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlue Then
            blnChange = True
        End If
    Next
    
    For i = 1 To vsf已标记.rows - 1
        If vsf已标记.Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlue Then
            blnChange = True
        End If
    Next
    
    If blnChange = True Then
        lngResult = MsgBox("刚有内容被修改了，是否继续？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
    End If
    
    If lngResult = vbYes Or blnChange = False Then
        Call GetData
    End If
End Sub

Private Sub GetData()
    Dim rsRecord As ADODB.Recordset
    Dim i As Integer
    Dim blnChange As Boolean
    Dim lngResult As Long
    Dim dbDate As Date
    
    On Error GoTo errHandle
'    blnChange = False
'    For i = 1 To vsf未标记.rows - 1
'        If vsf未标记.Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlue Then
'            blnChange = True
'        End If
'    Next
'
'    For i = 1 To vsf已标记.rows - 1
'        If vsf已标记.Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlue Then
'            blnChange = True
'        End If
'    Next
'
'    If blnChange = True Then
'        lngResult = MsgBox("刚有内容被修改了，是否保存？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
'    End If
'
'    If lngResult = vbYes Then
'        Call Save
'    End If
        
    '提取数据的方法
    If Trim(txt供应商.Text) = "" And Trim(txt开始发票号.Text) = "" And Trim(txt结束发票号.Text) = "" And Trim(cbo审核日期.Text) = "" Then
        Exit Sub
    End If
    If txt供应商.Text = "" Then
        txt供应商.Tag = ""
    End If
    gstrSQL = "Select distinct a.Id, a.入库单据号 no, a.项目id, Decode(a.付款标志, Null, 0, 0, 0, 1) 付款标志, a.批号, a.品名 As 药品名称, a.规格, a.数量, a.计量单位, a.采购价, a.采购金额," & _
              "     A.付款序号 , A.审核人, A.审核日期, A.发票号, A.发票金额, A.发票日期, b.剂量系数, b.门诊包装, b.住院包装, b.药库包装,b.门诊单位,b.住院单位,b.药库单位 " & _
              "  From 应付记录 A, 药品规格 B where a.项目id = b.药品id And a.审核人 is not null and a.库房id=[1] and a.发票号 is not null and  a.发票日期 is not null and a.记录状态=1 and a.记录性质=0 "
    
    
    If txt供应商.Tag <> "" Then
        gstrSQL = gstrSQL & " and a.单位id=[2]"
    End If
        
    If Me.txt开始发票号 <> "" And Me.txt结束发票号 <> "" Then gstrSQL = gstrSQL & " And a.发票号 >= [3] And a.发票号 <=[4] "
    If Me.txt开始发票号 <> "" And Me.txt结束发票号 = "" Then gstrSQL = gstrSQL & " And a.发票号 >= [3] "
    If Me.txt开始发票号 = "" And Me.txt结束发票号 <> "" Then gstrSQL = gstrSQL & " And a.发票号 <= [3] "
    
    If cbo审核日期.Text = "今日" Then
        dbDate = CDate(Format(Date, "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.审核日期 between [5] and sysdate"
    End If
    If cbo审核日期.Text = "一星期内" Then
        dbDate = CDate(Format(DateAdd("d", -7, Date), "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.审核日期 between [5] and sysdate"
    End If
    If cbo审核日期.Text = "一个月内" Then
        dbDate = CDate(Format(DateAdd("d", -30, Date), "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.审核日期 between [5] and sysdate"
    End If
    If cbo审核日期.Text = "三个月内" Then
        dbDate = CDate(Format(DateAdd("d", -90, Date), "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.审核日期 between [5] and sysdate"
    End If
    If cbo审核日期.Text = "自定义日期" Then
        dbDate = CDate(Format(dtp开始时间, "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.审核日期 between [5] and [6]"
    End If
    
    If txt药品名称.Tag <> "" Then
        gstrSQL = gstrSQL & " and a.项目id=[7]"
    End If
    
    Set rsRecord = zlDataBase.OpenSQLRecord(gstrSQL, "提取数据", cbo库房.ItemData(cbo库房.ListIndex), txt供应商.Tag, UCase(txt开始发票号.Text), UCase(txt结束发票号.Text), dbDate, CDate(Format(dtp结束时间, "yyyy-mm-dd") & " 23:59:59"), txt药品名称.Tag)
    If rsRecord Is Nothing Then
        vsf未标记.rows = 1
        vsf已标记.rows = 1
        Exit Sub
    End If
    
    If vsf未标记.rows > 1 Then
        With vsf未标记
            .Cell(flexcpFontBold, 1, 0, .rows - 1, .Cols - 1) = False
            .Cell(flexcpForeColor, 1, 0, .rows - 1, .Cols - 1) = vbBlack
        End With
    End If
    If vsf已标记.rows > 1 Then
        With vsf已标记
            .Cell(flexcpFontBold, 1, 0, .rows - 1, .Cols - 1) = False
            .Cell(flexcpForeColor, 1, 0, .rows - 1, .Cols - 1) = vbBlack
        End With
    End If
    
    Call SetColumn(rsRecord)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get当前库存(ByVal lng药品ID As Long) As String
    '功能：某个药品在某个科室的当前库存
    '返回值：返回查询到的库存数
    '参数：药品id
    Dim rsRecord As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Sum(a.实际数量) 当前库存,b.计算单位  From 药品库存 a,收费项目目录 b where a.药品id=b.id and a.药品id = [1] And a.库房id = [2] group by b.计算单位"
    Set rsRecord = zlDataBase.OpenSQLRecord(gstrSQL, "获取当前库存", lng药品ID, cbo库房.ItemData(cbo库房.ListIndex))
    If rsRecord Is Nothing Or rsRecord.RecordCount = 0 Then
        Get当前库存 = "0"
        Exit Function
    Else
        Get当前库存 = IIf(IsNull(rsRecord!当前库存), "0", rsRecord!当前库存) & IIf(IsNull(rsRecord!计算单位), "", rsRecord!计算单位)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get全院库存(ByVal lng药品ID As Long) As String
    '功能：查询某个药品在全院的库存
    '返回值：返回查询到的库存数
    '参数：药品id
    Dim rsRecord As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Sum(a.实际数量) 全院库存,b.计算单位  From 药品库存 a,收费项目目录 b where a.药品id=b.id  and a.药品id=[1] group by b.计算单位"
    Set rsRecord = zlDataBase.OpenSQLRecord(gstrSQL, "获取当前库存", lng药品ID)
    If rsRecord Is Nothing Or rsRecord.RecordCount = 0 Then
        Get全院库存 = "0"
        Exit Function
    Else
        Get全院库存 = IIf(IsNull(rsRecord!全院库存), "0", rsRecord!全院库存) & IIf(IsNull(rsRecord!计算单位), "", rsRecord!计算单位)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetColumn(ByVal rsRecord As ADODB.Recordset)
    '给表格添加值
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    Set rsTemp = rsRecord
    
    rsRecord.Filter = "付款标志=0"
    With vsf未标记
        .rows = rsRecord.RecordCount + 1
        For i = 1 To rsRecord.RecordCount
            .TextMatrix(i, mColumnMark.序号) = i
            .TextMatrix(i, mColumnMark.批号) = IIf(IsNull(rsRecord!批号), "", rsRecord!批号)
            .TextMatrix(i, mColumnMark.id) = rsRecord!id
            .TextMatrix(i, mColumnMark.项目Id) = rsRecord!项目Id
            
            .TextMatrix(i, mColumnMark.NO) = rsRecord!NO
            If rsRecord!付款标志 = 0 Then
                .TextMatrix(i, mColumnMark.付款标志) = ""
            End If
            
            .TextMatrix(i, mColumnMark.药品名称) = rsRecord!药品名称
            .TextMatrix(i, mColumnMark.规格) = rsRecord!规格
            
            Select Case mint单位系数
                Case 4  '售价单位
                    .TextMatrix(i, mColumnMark.计算单位) = IIf(IsNull(rsRecord!计量单位), "", rsRecord!计量单位)
                Case 1  '药库单位
                    .TextMatrix(i, mColumnMark.计算单位) = IIf(IsNull(rsRecord!药库单位), "", rsRecord!药库单位)
                Case 2  '门诊单位
                    .TextMatrix(i, mColumnMark.计算单位) = IIf(IsNull(rsRecord!门诊单位), "", rsRecord!门诊单位)
                Case 3  '住院单位
                    .TextMatrix(i, mColumnMark.计算单位) = IIf(IsNull(rsRecord!住院单位), "", rsRecord!住院单位)
            End Select
            Select Case mint单位系数
                Case 4  '售价单位
                    .TextMatrix(i, mColumnMark.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.采购价) = zlStr.FormatEx(rsRecord!采购价, mintShowCostDigit, , True)
                Case 1  '药库单位
                    .TextMatrix(i, mColumnMark.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量 / IIf(IsNull(rsRecord!药库包装), 1, rsRecord!药库包装)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.采购价) = zlStr.FormatEx(rsRecord!采购价 * IIf(IsNull(rsRecord!药库包装), 1, rsRecord!药库包装), mintShowCostDigit, , True)
                Case 2  '门诊单位
                    .TextMatrix(i, mColumnMark.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量 / IIf(IsNull(rsRecord!门诊包装), 1, rsRecord!门诊包装)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.采购价) = zlStr.FormatEx(rsRecord!采购价 * IIf(IsNull(rsRecord!门诊包装), 1, rsRecord!门诊包装), mintShowCostDigit, , True)
                Case 3  '住院单位
                    .TextMatrix(i, mColumnMark.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量 / IIf(IsNull(rsRecord!住院包装), 1, rsRecord!住院包装)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.采购价) = zlStr.FormatEx(rsRecord!采购价 * IIf(IsNull(rsRecord!住院包装), 1, rsRecord!住院包装), mintShowCostDigit, , True)
            End Select
            
            .TextMatrix(i, mColumnMark.采购金额) = zlStr.FormatEx(rsRecord!采购金额, mintShowMoneyDigit, , True)
'            .TextMatrix(i, mColumnMark.当前库存) = Get当前库存(rsRecord!项目Id)
'            .TextMatrix(i, mColumnMark.全院库存) = Get全院库存(rsRecord!项目Id)
            
            If IsNull(rsRecord!付款序号) Or rsRecord!付款序号 = 0 Then
                .TextMatrix(i, mColumnMark.付款序号) = "未付款"
            Else
                .TextMatrix(i, mColumnMark.付款序号) = "已付款"
            End If
            .TextMatrix(i, mColumnMark.审核人) = rsRecord!审核人
            .TextMatrix(i, mColumnMark.审核日期) = Format(rsRecord!审核日期, "yyyy-mm-dd")
            .TextMatrix(i, mColumnMark.发票号) = IIf(IsNull(rsRecord!发票号), "", rsRecord!发票号)
            .TextMatrix(i, mColumnMark.发票金额) = zlStr.FormatEx(IIf(IsNull(rsRecord!发票金额), "", rsRecord!发票金额), mintShowMoneyDigit, , True)
            .TextMatrix(i, mColumnMark.发票日期) = IIf(IsNull(rsRecord!发票日期), "", Format(rsRecord!发票日期, "yyyy-mm-dd"))
            .TextMatrix(i, mColumnMark.剂量系数) = rsRecord!剂量系数
            .TextMatrix(i, mColumnMark.门诊包装) = rsRecord!门诊包装
            .TextMatrix(i, mColumnMark.住院包装) = rsRecord!住院包装
            .TextMatrix(i, mColumnMark.药库包装) = rsRecord!药库包装
            .RowHeight(i) = 310
            If .TextMatrix(i, mColumnMark.付款序号) = "已付款" Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
            rsRecord.MoveNext
        Next
    End With
    
    rsTemp.Filter = "付款标志=1"
    With vsf已标记
        .rows = rsRecord.RecordCount + 1
        For i = 1 To rsRecord.RecordCount
            .TextMatrix(i, mColumnMark.序号) = i
            .TextMatrix(i, mColumnMark.批号) = IIf(IsNull(rsRecord!批号), "", rsRecord!批号)
            .TextMatrix(i, mColumnMark.id) = rsRecord!id
            .TextMatrix(i, mColumnMark.项目Id) = rsRecord!项目Id
            
            .TextMatrix(i, mColumnMark.NO) = rsRecord!NO
            If rsRecord!付款标志 = 1 Then
                .TextMatrix(i, mColumnMark.付款标志) = ""
            End If
            
            .TextMatrix(i, mColumnMark.药品名称) = rsRecord!药品名称
            .TextMatrix(i, mColumnMark.规格) = rsRecord!规格
            
            Select Case mint单位系数
                Case 4  '售价单位
                    .TextMatrix(i, mColumnMark.计算单位) = IIf(IsNull(rsRecord!计量单位), "", rsRecord!计量单位)
                Case 1  '药库单位
                    .TextMatrix(i, mColumnMark.计算单位) = IIf(IsNull(rsRecord!药库单位), "", rsRecord!药库单位)
                Case 2  '门诊单位
                    .TextMatrix(i, mColumnMark.计算单位) = IIf(IsNull(rsRecord!门诊单位), "", rsRecord!门诊单位)
                Case 3  '住院单位
                    .TextMatrix(i, mColumnMark.计算单位) = IIf(IsNull(rsRecord!住院单位), "", rsRecord!住院单位)
            End Select
            Select Case mint单位系数
                Case 4  '售价单位
                    .TextMatrix(i, mColumnMark.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.采购价) = zlStr.FormatEx(rsRecord!采购价, mintShowCostDigit, , True)
                Case 1  '药库单位
                    .TextMatrix(i, mColumnMark.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量 / IIf(IsNull(rsRecord!药库包装), 1, rsRecord!药库包装)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.采购价) = zlStr.FormatEx(rsRecord!采购价 * IIf(IsNull(rsRecord!药库包装), 1, rsRecord!药库包装), mintShowCostDigit, , True)
                Case 2  '门诊单位
                    .TextMatrix(i, mColumnMark.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量 / IIf(IsNull(rsRecord!门诊包装), 1, rsRecord!门诊包装)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.采购价) = zlStr.FormatEx(rsRecord!采购价 * IIf(IsNull(rsRecord!门诊包装), 1, rsRecord!门诊包装), mintShowCostDigit, , True)
                Case 3  '住院单位
                    .TextMatrix(i, mColumnMark.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量 / IIf(IsNull(rsRecord!住院包装), 1, rsRecord!住院包装)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.采购价) = zlStr.FormatEx(rsRecord!采购价 * IIf(IsNull(rsRecord!住院包装), 1, rsRecord!住院包装), mintShowCostDigit, , True)
            End Select
            
            .TextMatrix(i, mColumnMark.采购金额) = zlStr.FormatEx(rsRecord!采购金额, mintShowMoneyDigit, , True)
'            .TextMatrix(i, mColumnMark.当前库存) = Get当前库存(rsRecord!项目Id)
'            .TextMatrix(i, mColumnMark.全院库存) = Get全院库存(rsRecord!项目Id)
            
            If IsNull(rsRecord!付款序号) Or rsRecord!付款序号 = 0 Then
                .TextMatrix(i, mColumnMark.付款序号) = "未付款"
            Else
                .TextMatrix(i, mColumnMark.付款序号) = "已付款"
            End If
            .TextMatrix(i, mColumnMark.审核人) = rsRecord!审核人
            .TextMatrix(i, mColumnMark.审核日期) = Format(rsRecord!审核日期, "yyyy-mm-dd")
            .TextMatrix(i, mColumnMark.发票号) = IIf(IsNull(rsRecord!发票号), "", rsRecord!发票号)
            .TextMatrix(i, mColumnMark.发票金额) = zlStr.FormatEx(IIf(IsNull(rsRecord!发票金额), "", rsRecord!发票金额), mintShowMoneyDigit, , True)
            .TextMatrix(i, mColumnMark.发票日期) = IIf(IsNull(rsRecord!发票日期), "", Format(rsRecord!发票日期, "yyyy-mm-dd"))
            .TextMatrix(i, mColumnMark.剂量系数) = rsRecord!剂量系数
            .TextMatrix(i, mColumnMark.门诊包装) = rsRecord!门诊包装
            .TextMatrix(i, mColumnMark.住院包装) = rsRecord!住院包装
            .TextMatrix(i, mColumnMark.药库包装) = rsRecord!药库包装
            .RowHeight(i) = 310
            If .TextMatrix(i, mColumnMark.付款序号) = "已付款" Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
            rsRecord.MoveNext
        Next
    End With
    If vsf未标记.rows > 1 Then
        vsf未标记.Select 1, 1
    End If
End Sub

Private Sub Form_Resize()
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - staThis.Panels(5).Width - staThis.Panels(6).Width - .Width - 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim strTemp As String
    
    SaveWinState Me, App.ProductName, MStrCaption
    mblnDo = False
    mvMsg = vbYes
    With vsf列头
        For i = 1 To .rows - 1
            strTemp = strTemp & IIf(.TextMatrix(i, 1) = "", 0, .TextMatrix(i, 1)) & "," & .TextMatrix(i, 2) & "|"
        Next
    End With
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\付款标记", "列显示隐藏", strTemp)
    Call ReleaseSelectorRS
End Sub

Private Sub picSetCols_Click(Index As Integer)
    With vsf列头
        .Top = vsf未标记.Top + .CellHeight
        .Left = vsf未标记.Left + 10
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    If tbcDetails.Item(0).Selected = True Then
        If Button = 1 And vsf未标记.Height + y > 200 And picSplit.Top + y < staThis.Top - 1500 Then
            Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
            vsf未标记.Move vsf未标记.Left, chk全选.Height, lngRight - lngLeft, vsf未标记.Height + y
            picSplit.Move vsf未标记.Left, picSplit.Top + y, lngRight - lngLeft
            vsf库存.Move vsf未标记.Left, picSplit.Top + picSplit.Height, lngRight - lngLeft, picList.Height - y
        End If
    Else
        If Button = 1 And vsf已标记.Height + y > 200 And picSplit.Top + y < staThis.Top - 1500 Then
            Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
            vsf已标记.Move vsf已标记.Left, chk全选.Height, lngRight - lngLeft, vsf已标记.Height + y
            picSplit.Move vsf已标记.Left, picSplit.Top + y, lngRight - lngLeft
            vsf库存.Move vsf已标记.Left, picSplit.Top + picSplit.Height, lngRight - lngLeft, picList.Height - y
        End If
    End If
End Sub

Private Sub tbcDetails_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim cbrToolBar As CommandBar
    
    If tbcDetails.Item(0).Selected = True Then
        cbsMain.FindControl(xtpControlButton, menuToolSave).Visible = True
        cbsMain.FindControl(xtpControlButton, menuToolSave2).Visible = False
        vsf已标记.Visible = False
        vsf未标记.Visible = True
        chk全选.Enabled = True
        vsf未标记.Move 0, chk全选.Height, picList.Width, (picList.Height / 6) * 5
        picSplit.Move 0, vsf未标记.Top + vsf未标记.Height, picList.ScaleWidth, 50
        vsf库存.Move 0, picSplit.Top + picSplit.Height, picList.ScaleWidth, picList.Height - picSplit.Top - picSplit.Height
    Else
        cbsMain.FindControl(xtpControlButton, menuToolSave).Visible = False
        cbsMain.FindControl(xtpControlButton, menuToolSave2).Visible = True
        vsf已标记.Visible = True
        vsf未标记.Visible = False
        chk全选.Enabled = False
        vsf已标记.Move 0, chk全选.Height, picList.ScaleWidth, (picList.ScaleHeight / 6) * 5
        picSplit.Move 0, vsf已标记.Top + vsf已标记.Height, picList.ScaleWidth, 50
        vsf库存.Move 0, picSplit.Top + picSplit.Height, picSplit.ScaleWidth, picList.Height - picSplit.Top - picSplit.Height
    End If
    Call setTabControlColor(tbcDetails)
End Sub

Private Sub txt供应商_GotFocus()
    zlControl.TxtSelAll txt供应商
    OS.OpenIme True
End Sub

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsRecord As ADODB.Recordset
    
    If KeyCode = 13 Then
        vRect = zlControl.GetControlRect(txt供应商.hWnd) '获取位置
        
        gstrSQL = "select id,编码,名称,简码 from 供应商 where 末级 = 1 and (编码 like [1] OR 名称 like [1] OR 简码 like [1])"
        Set rsRecord = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "供应商", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txt供应商.Height, blnCancel, False, True, UCase(txt供应商.Text) & "%")
        
        If blnCancel Then txt供应商.SetFocus: Exit Sub
        
        If rsRecord Is Nothing Then
            MsgBox "没有您输入的供应商，请重输！", vbOKOnly + vbInformation, gstrSysName
            txt供应商.SelStart = 0
            txt供应商.SelLength = Len(txt供应商)
            Exit Sub
        Else
            If txt供应商.Tag <> rsRecord!id Then
                txt药品名称.Tag = ""
                txt药品名称.Text = ""
                txt开始发票号.Text = ""
                txt结束发票号.Text = ""
            End If
            txt供应商.Text = rsRecord!名称
            txt供应商.Tag = rsRecord!id
        End If
    End If
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then
        txt供应商.Text = ""
        txt供应商.Tag = ""
    End If
End Sub

Private Sub txt药品名称_Change()
    If txt药品名称.Text = "" Then
         txt药品名称.Tag = ""
    End If
End Sub

Private Sub txt药品名称_GotFocus()
    zlControl.TxtSelAll txt药品名称
End Sub

Private Sub txt药品名称_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecReturn As Recordset
    Dim strkey As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt药品名称.Text) = "" Then Exit Sub
    vRect = zlControl.GetControlRect(txt药品名称.hWnd) '获取位置
    dblLeft = vRect.Left
    dblTop = vRect.Top + txt药品名称.Height
    
    strkey = Trim(txt药品名称.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品外购入库管理", cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex))
    End If
    
    Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, dblLeft, dblTop, cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex), , , , , , False, mstrPrivs)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        txt药品名称.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        txt药品名称.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    txt药品名称.Tag = RecReturn!药品id
End Sub

Private Sub txt药品名称_Validate(Cancel As Boolean)
    If Trim(txt药品名称.Text) <> "" Then
        Call txt药品名称_KeyDown(vbKeyReturn, 1)
    End If
End Sub

Private Sub vsf列头_LostFocus()
    vsf列头.Visible = False
End Sub

Private Sub vsf列头_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsf列头
        If .Row = 0 Then Exit Sub
        If .Col = 1 And Button = 1 Then
            If .TextMatrix(.Row, .Col) = "1" Then
                .TextMatrix(.Row, .Col) = ""
'                If tbcDetails.Item(0).Selected = True Then
                    vsf未标记.colHidden(vsf未标记.ColIndex(.TextMatrix(.Row, 2))) = True
'                Else
                    vsf已标记.colHidden(vsf已标记.ColIndex(.TextMatrix(.Row, 2))) = True
'                End If
            ElseIf .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "1"
'                If tbcDetails.Item(0).Selected = True Then
                    vsf未标记.colHidden(vsf未标记.ColIndex(.TextMatrix(.Row, 2))) = False
'                Else
                    vsf已标记.colHidden(vsf已标记.ColIndex(.TextMatrix(.Row, 2))) = False
'                End If
            End If
        End If
    End With
End Sub

Private Sub vsf未标记_DblClick()
    With vsf未标记
        vsf列头.Visible = False
        If .Row = 0 Then
            Exit Sub
        End If
        If .TextMatrix(.Row, mColumnMark.付款标志) = "√" Then
            .TextMatrix(.Row, mColumnMark.付款标志) = ""
            .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = False
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlack
        Else
            .TextMatrix(.Row, mColumnMark.付款标志) = "√"
            .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlue
        End If
    End With
End Sub

Private Sub vsf未标记_EnterCell()
    Dim rsTemp As ADODB.Recordset
    Dim rs单位 As ADODB.Recordset
    Dim i As Integer
    Dim int住院系数 As Integer
    Dim int门诊系数 As Integer
    Dim int药库系数 As Integer
    
    On Error GoTo errHandle
    With vsf未标记
        If .rows = 1 Then
            vsf库存.rows = 0
            Exit Sub
        End If
        If .TextMatrix(.Row, mColumnMark.项目Id) <> "" And .Row <> 0 Then
            gstrSQL = "Select 名称, Sum(实际数量) 数量" & _
                      "  From (Select 实际数量, b.名称 " & _
                              " From 药品库存 A, 部门表 B, (Select Distinct 执行科室id From 收费执行科室 Where 收费细目id = [1]) D" & _
                              " Where a.库房id = b.Id And a.库房id = d.执行科室id And a.药品id = [1]" & _
                              " Union All " & _
                              " Select 0 数量, a.名称" & _
                              " From 部门表 A, (Select Distinct 执行科室id From 收费执行科室 Where 收费细目id = [1]) B" & _
                              " Where a.Id = b.执行科室id)" & _
                      "  Group By 名称 "
           Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, .TextMatrix(.Row, mColumnMark.项目Id))
           If Not rsTemp Is Nothing Then
                With vsf库存
                    .Cols = rsTemp.RecordCount + 1
                    .rows = 2
                    For i = 1 To rsTemp.RecordCount
                        .TextMatrix(0, 0) = "库房"
                        .TextMatrix(1, 0) = "数量"
                        VsfGridColFormat vsf库存, i, rsTemp!名称, 1500, flexAlignCenterCenter, rsTemp!名称
                        
                        int住院系数 = Val(vsf未标记.TextMatrix(vsf未标记.Row, mColumnMark.住院包装))
                        int门诊系数 = Val(vsf未标记.TextMatrix(vsf未标记.Row, mColumnMark.门诊包装))
                        int药库系数 = Val(vsf未标记.TextMatrix(vsf未标记.Row, mColumnMark.药库包装))
                        Select Case mint单位系数
                            Case 4  '售价单位
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!数量), 0, rsTemp!数量), mintShowNumberDigit, , True)
                            Case 1  '药库单位
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!数量), 0, rsTemp!数量 / int药库系数), mintShowNumberDigit, , True)
                            Case 2  '门诊单位
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!数量), 0, rsTemp!数量 / int门诊系数), mintShowNumberDigit, , True)
                            Case 3  '住院单位
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!数量), 0, rsTemp!数量 / int住院系数), mintShowNumberDigit, , True)
                        End Select
                        rsTemp.MoveNext
                    Next
                    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
                    .Cell(flexcpAlignment, 1, 1, 1, .Cols - 1) = flexAlignRightCenter
                    .ColAlignment(0) = flexAlignCenterCenter
                    .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
                    .RowHeight(0) = 300
                    .RowHeight(1) = 300
                End With
           End If
        End If
        
        If .Row = 0 Then
            vsf库存.Clear
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf已标记_DblClick()
    With vsf已标记
        vsf列头.Visible = False
        If .Row = 0 Then
            Exit Sub
        End If
        If Trim(.TextMatrix(.Row, mColumnMark.付款序号)) = "未付款" Then '已经付款的单据不能修改付款标志
            If .TextMatrix(.Row, mColumnMark.付款标志) = "√" Then
                .TextMatrix(.Row, mColumnMark.付款标志) = ""
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = False
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlack
            Else
                .TextMatrix(.Row, mColumnMark.付款标志) = "√"
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlue
            End If
        End If
    End With
End Sub

'Private Sub vsf未标记_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    With vsf未标记
'        vsf列头.Visible = False
'        If .Row = 0 Then
'            Exit Sub
'        End If
'        If y < .CellHeight * .rows Then
'            If Button = 1 Then
'                If .Col = mColumnMark.付款标志 Then
'                    If .TextMatrix(.Row, .Col) = "√" Then
'                        .TextMatrix(.Row, .Col) = ""
'                    Else
'                        .TextMatrix(.Row, .Col) = "√"
'                        .Cell(flexcpFontBold, .Row, .Col) = True
'                        .Cell(flexcpFontSize, .Row, .Col) = 10
'                        .Cell(flexcpForeColor, .Row, .Col) = vbBlue
'                    End If
'                End If
'            End If
'        End If
'    End With
'End Sub


Private Sub vsf已标记_EnterCell()
    Dim rsTemp As ADODB.Recordset
    Dim rs单位 As ADODB.Recordset
    Dim i As Integer
    Dim int住院系数 As Integer
    Dim int门诊系数 As Integer
    Dim int药库系数 As Integer
    
    On Error GoTo errHandle
    With vsf已标记
        If .rows = 1 Then
            vsf库存.rows = 0
            Exit Sub
        End If
        If .TextMatrix(.Row, mColumnMark.项目Id) <> "" And .Row <> 0 Then
            gstrSQL = "Select 名称, Sum(实际数量) 数量" & _
                      "  From (Select 实际数量, b.名称 " & _
                              " From 药品库存 A, 部门表 B, (Select Distinct 执行科室id From 收费执行科室 Where 收费细目id = [1]) D" & _
                              " Where a.库房id = b.Id And a.库房id = d.执行科室id And a.药品id = [1]" & _
                              " Union All " & _
                              " Select 0 数量, a.名称" & _
                              " From 部门表 A, (Select Distinct 执行科室id From 收费执行科室 Where 收费细目id = [1]) B" & _
                              " Where a.Id = b.执行科室id)" & _
                      "  Group By 名称 "
           Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, .TextMatrix(.Row, mColumnMark.项目Id))
           If Not rsTemp Is Nothing Then
                With vsf库存
                    .Cols = rsTemp.RecordCount + 1
                    .rows = 2
                    For i = 1 To rsTemp.RecordCount
                        .TextMatrix(0, 0) = "库房"
                        .TextMatrix(1, 0) = "数量"
                        VsfGridColFormat vsf库存, i, rsTemp!名称, 1500, flexAlignCenterCenter, rsTemp!名称
                        
                        int住院系数 = Val(vsf已标记.TextMatrix(vsf已标记.Row, mColumnMark.住院包装))
                        int门诊系数 = Val(vsf已标记.TextMatrix(vsf已标记.Row, mColumnMark.门诊包装))
                        int药库系数 = Val(vsf已标记.TextMatrix(vsf已标记.Row, mColumnMark.药库包装))
                        Select Case mint单位系数
                            Case 4  '售价单位
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!数量), 0, rsTemp!数量), mintShowNumberDigit, , True)
                            Case 1  '药库单位
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!数量), 0, rsTemp!数量 / int药库系数), mintShowNumberDigit, , True)
                            Case 2  '门诊单位
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!数量), 0, rsTemp!数量 / int门诊系数), mintShowNumberDigit, , True)
                            Case 3  '住院单位
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!数量), 0, rsTemp!数量 / int住院系数), mintShowNumberDigit, , True)
                        End Select
                        .ColAlignment(i) = flexAlignRightCenter
                        .ColWidth(i) = 1500
                        rsTemp.MoveNext
                    Next
                    .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
                End With
           End If
        End If
        
        If .Row = 0 Then
            vsf库存.Clear
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'Private Sub vsf已标记_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim blnMsg As Boolean
'
'    With vsf已标记
'        vsf列头.Visible = False
'        If .Row = 0 Then
'            Exit Sub
'        End If
'        If y < .CellHeight * .rows Then
'            If Button = 1 Then
'                If .Col = mColumnMark.付款标志 Then
'                    If Trim(.TextMatrix(.Row, mColumnMark.付款序号)) = "未付款" Then '已经付款的单据不能修改付款标志
'                        If .TextMatrix(.Row, .Col) = "√" Then
''                            If mvMsg <> vbCancel And mvMsg <> vbIgnore Then
''                                mvMsg = frmMsgBox.ShowMsgBox("该单据已经标记付款，确定要取消标记？", Me)
''                                blnMsg = True
''                            End If
''                            If (mvMsg = vbYes Or mvMsg = vbIgnore) And blnMsg = True Then
'                                .TextMatrix(.Row, .Col) = ""
''                            End If
''                            If (mvMsg = vbCancel Or mvMsg = vbIgnore) And blnMsg = False Then
''                                .TextMatrix(.Row, .Col) = ""
''                            End If
'                        Else
'                            .TextMatrix(.Row, .Col) = "√"
'                            .Cell(flexcpFontBold, .Row, .Col) = True
'                            .Cell(flexcpFontSize, .Row, .Col) = 10
'                            .Cell(flexcpForeColor, .Row, .Col) = vbBlue
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    End With
'End Sub



Private Sub ExitForm()
    '退出窗体的方法
'    Dim i As Integer
'    Dim blnChange As Boolean
'    Dim lngResult As Long
'
'    blnChange = False
'    For i = 1 To vsf未标记.rows - 1
'        If vsf未标记.Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlue Then
'            blnChange = True
'        End If
'    Next
'
'    For i = 1 To vsf已标记.rows - 1
'        If vsf已标记.Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlue Then
'            blnChange = True
'        End If
'    Next
'
'    If blnChange = True Then
'        lngResult = MsgBox("刚有内容被修改了，是否保存？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
'
'        If lngResult = vbNo Then    '退出窗体 不保存
'            Unload Me
'        Else    '退出窗体，保存
'            Call Save
'            Unload Me
'        End If
'    Else
'        Unload Me
'    End If
    Unload Me
End Sub

Private Sub Save()
    '保存方法
    Dim i As Integer
    Dim intTemp As Integer
    Dim strTemp As String
    Dim blnContinue As Boolean
    
    blnContinue = False
    If tbcDetails.Item(mPage.未标记单据).Selected = True Then
        If MsgBox("将标记单据，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'            strTemp = "标记成功,"
            With vsf未标记
                For i = 1 To .rows - 1
                    If .TextMatrix(i, mColumnMark.付款标志) = "√" Then
                        intTemp = 1
                    Else
                        intTemp = 0
                    End If
                    gstrSQL = "zl_应付记录_付款标志(" & .TextMatrix(i, mColumnMark.id) & ","
                    gstrSQL = gstrSQL & intTemp & ")"
                    
                    zlDataBase.ExecuteProcedure gstrSQL, MStrCaption
                Next
                    
                For i = 1 To .rows - 1
                    .Cell(flexcpFontBold, i, mColumnMark.付款标志) = False
                    .Cell(flexcpFontSize, i, mColumnMark.付款标志) = 9
                    .Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlack
                Next
            End With
            blnContinue = True
        End If
    End If
    
    If tbcDetails.Item(mPage.已标记单据).Selected = True Then
        If MsgBox("将取消单据标记，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'            strTemp = "取消标记成功,"
            With vsf已标记
                For i = 1 To .rows - 1
                    If .TextMatrix(i, mColumnMark.付款标志) = "√" Then
                        intTemp = 0
                    Else
                        intTemp = 1
                    End If
                    gstrSQL = "zl_应付记录_付款标志(" & .TextMatrix(i, mColumnMark.id) & ","
                    gstrSQL = gstrSQL & intTemp & ")"
                    
                    zlDataBase.ExecuteProcedure gstrSQL, MStrCaption
                Next
                
                For i = 1 To .rows - 1
                    .Cell(flexcpFontBold, i, mColumnMark.付款标志) = False
                    .Cell(flexcpFontSize, i, mColumnMark.付款标志) = 9
                    .Cell(flexcpForeColor, i, mColumnMark.付款标志) = vbBlack
                Next
            End With
            blnContinue = True
        End If
    End If
'    If MsgBox(strTemp & "是否清空条件！", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'        txt供应商.Text = ""
'        txt药品名称.Tag = ""
'        txt药品名称.Text = ""
'        cbo审核日期.ListIndex = 0
'        txt开始发票号.Text = ""
'        txt结束发票号.Text = ""
'    End If
    If blnContinue = True Then
        Call GetData
    End If
End Sub

Private Sub setTabControlColor(ByVal objtbc As TabControl)
    '对Tabcontrol控件进行颜色判断
    Dim i As Integer
    
    With objtbc
        For i = 0 To .ItemCount - 1
            If .Item(i).Selected = True Then
                .Item(i).Color = CSTCOLOR_UNMODIFY
            Else
                .Item(i).Color = CSTCOLOR_NORECORDS
            End If
        Next
    End With
End Sub





