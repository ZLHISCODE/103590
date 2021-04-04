VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLR_S 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   4140
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5010
      ScaleWidth      =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1140
      Width           =   45
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6150
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReport.frx":014A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11298
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmReport.frx":09DE
      Height          =   5175
      LargeChange     =   20
      Left            =   9225
      Max             =   100
      SmallChange     =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   250
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmReport.frx":0CE8
      Height          =   250
      LargeChange     =   20
      Left            =   4185
      Max             =   100
      SmallChange     =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5895
      Width           =   4995
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   5145
      Left            =   4230
      ScaleHeight     =   5085
      ScaleWidth      =   4950
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   735
      Width           =   5010
      Begin VB.PictureBox picRotate 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4440
         ScaleHeight     =   285
         ScaleWidth      =   330
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   330
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfRelations 
         Height          =   1215
         Left            =   720
         TabIndex        =   35
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
         _cx             =   1964641833
         _cy             =   1964640351
         Appearance      =   1
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
         BackColor       =   14737632
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14737632
         GridColor       =   16761024
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   30
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         FillStyle       =   1
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
      Begin VSFlex8Ctl.VSFlexGrid msh 
         Height          =   1575
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   3135
         _cx             =   1964643738
         _cy             =   1964640986
         Appearance      =   0
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
         MouseIcon       =   "frmReport.frx":0FF2
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16777215
         ForeColorFixed  =   0
         BackColorSel    =   10251637
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   3960
         ScaleHeight     =   765
         ScaleWidth      =   330
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1815
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Index           =   0
         Left            =   255
         ScaleHeight     =   3390
         ScaleWidth      =   3315
         TabIndex        =   6
         Top             =   165
         Width           =   3315
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   -8888
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   30
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   330
         ScaleHeight     =   3390
         ScaleWidth      =   3315
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   3315
      End
   End
   Begin VB.PictureBox picGroup 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   0
      ScaleHeight     =   5010
      ScaleWidth      =   4140
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "Save"
      Top             =   1140
      Width           =   4140
      Begin VB.PictureBox picPar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F4F4F4&
         Height          =   3090
         Left            =   45
         ScaleHeight     =   3030
         ScaleWidth      =   4050
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "Save"
         Top             =   2325
         Width           =   4110
         Begin VB.CommandButton cmdSelAll 
            Caption         =   "全选"
            Height          =   350
            Left            =   120
            TabIndex        =   34
            Top             =   930
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmdSelNone 
            Cancel          =   -1  'True
            Caption         =   "全清"
            Height          =   350
            Left            =   765
            TabIndex        =   33
            Top             =   930
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmdLoad 
            BackColor       =   &H00F4F4F4&
            Caption         =   "确定(&O)"
            Height          =   350
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   930
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdDefault 
            BackColor       =   &H00F4F4F4&
            Caption         =   "条件(&D)"
            Height          =   350
            Left            =   2850
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   930
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.Frame fraGroup 
            BackColor       =   &H00F4F4F4&
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   16
            Top             =   -60
            Visible         =   0   'False
            Width           =   3825
         End
         Begin VB.Frame fra 
            BackColor       =   &H00F4F4F4&
            ForeColor       =   &H00800000&
            Height          =   645
            Index           =   0
            Left            =   210
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   3825
            Begin VB.OptionButton opt 
               BackColor       =   &H00F4F4F4&
               Caption         =   "#"
               Height          =   180
               Index           =   0
               Left            =   105
               MaskColor       =   &H8000000F&
               TabIndex        =   18
               Top             =   270
               Visible         =   0   'False
               Width           =   1150
            End
         End
         Begin VB.CheckBox chk 
            BackColor       =   &H00F4F4F4&
            Caption         =   "#"
            Height          =   195
            Index           =   0
            Left            =   1455
            TabIndex        =   23
            Top             =   255
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cbo 
            BackColor       =   &H00F4F4F4&
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   21
            Top             =   195
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00F4F4F4&
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   20
            Top             =   195
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00F4F4F4&
            Caption         =   "…"
            Height          =   240
            Index           =   0
            Left            =   4425
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "按 F2 打开选择器"
            Top             =   225
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   22
            Top             =   195
            Visible         =   0   'False
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16053492
            CalendarTitleBackColor=   12946264
            CalendarTitleForeColor=   16053492
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   43778051
            CurrentDate     =   36731
         End
         Begin VB.Frame fraSplit 
            BackColor       =   &H00F4F4F4&
            Height          =   75
            Left            =   -180
            TabIndex        =   25
            Top             =   750
            Visible         =   0   'False
            Width           =   10000
         End
         Begin VB.Label lblName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "参数名称"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   675
            TabIndex        =   24
            Top             =   255
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1845
         Left            =   45
         TabIndex        =   11
         Tag             =   "Save"
         Top             =   225
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3254
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "编号"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "说明"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblPar_S 
         BackColor       =   &H009B6737&
         Caption         =   " 报表条件"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         MousePointer    =   7  'Size N S
         TabIndex        =   13
         Top             =   2100
         Width           =   4080
      End
      Begin VB.Label lblGroup_S 
         BackColor       =   &H009B6737&
         Caption         =   " 报表组"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   12
         Top             =   15
         Width           =   4095
      End
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   2011
      _CBWidth        =   9480
      _CBHeight       =   1140
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   4500
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Caption2        =   "格式"
      Child2          =   "cboFormat"
      MinWidth2       =   2505
      MinHeight2      =   330
      Width2          =   4005
      NewRow2         =   0   'False
      Caption3        =   "查找"
      Child3          =   "txtFind"
      MinWidth3       =   1005
      MinHeight3      =   330
      Width3          =   1935
      NewRow3         =   0   'False
      Begin VB.TextBox txtFind 
         Height          =   330
         Left            =   585
         TabIndex        =   32
         Top             =   780
         Width           =   8805
      End
      Begin MSComctlLib.ImageCombo cboFormat 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   225
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   16053492
         Locked          =   -1  'True
         ImageList       =   "img16"
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "打印预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "图形"
               Key             =   "Graph"
               Description     =   "图形"
               Object.ToolTipText     =   "对当前表格进行图形分析"
               Object.Tag             =   "图形"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重置"
               Key             =   "Par"
               Description     =   "重置"
               Object.ToolTipText     =   "重设条件"
               Object.Tag             =   "重置"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Par_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "列宽"
               Key             =   "ColWidth"
               Description     =   "列宽"
               Object.ToolTipText     =   "列宽"
               Object.Tag             =   "列宽"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Object.Tag             =   "合适匹配"
                     Text            =   "合适匹配"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Fill"
                     Object.Tag             =   "补齐表宽"
                     Text            =   "补齐表宽"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Def"
                     Object.Tag             =   "缺省定义"
                     Text            =   "缺省定义"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "选择"
               Key             =   "SelMode"
               Description     =   "选择"
               Object.ToolTipText     =   "表格行列选择模式"
               Object.Tag             =   "选择"
               ImageKey        =   "SelMode"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RowMode"
                     Object.Tag             =   "整行选择"
                     Text            =   "整行选择"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ColMode"
                     Object.Tag             =   "整列选择"
                     Text            =   "整列选择"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "列表"
               Key             =   "Style"
               Object.ToolTipText     =   "报表组列表显示方式"
               Object.Tag             =   "列表"
               ImageKey        =   "Style"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Large"
                     Object.Tag             =   "大图标"
                     Text            =   "大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "小图标"
                     Text            =   "小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "列表"
                     Text            =   "列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "详细资料"
                     Text            =   "详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Style_"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "前张"
               Key             =   "Pre"
               Description     =   "前张"
               Object.ToolTipText     =   "切换到前一张报表(Page Up)"
               Object.Tag             =   "前张"
               ImageKey        =   "Pre"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "后张"
               Key             =   "Next"
               Description     =   "后张"
               Object.ToolTipText     =   "切换到后一张报表(Page Down)"
               Object.Tag             =   "后张"
               ImageKey        =   "Next"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Page_"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   705
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":18CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2134
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":234E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2568
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2782
            Key             =   "Style"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":299C
            Key             =   "Pre"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2BB6
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2DD0
            Key             =   "SelMode"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   75
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3204
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":341E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3638
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3852
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3A6C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3EA0
            Key             =   "Style"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":40BA
            Key             =   "Pre"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":42D4
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":44EE
            Key             =   "SelMode"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timHead 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSScriptControlCtl.ScriptControl Srt 
      Left            =   6855
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   2745
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":4708
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2100
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":4A22
            Key             =   "Format"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":4B7C
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin C1Chart2D8.Chart2D Chart 
      Height          =   1230
      Index           =   0
      Left            =   4275
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   1650
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   2910
      _ExtentY        =   2170
      _StockProps     =   0
      ControlProperties=   "frmReport.frx":4CD6
   End
   Begin VB.Image imgCode 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   4230
      Stretch         =   -1  'True
      Top             =   2415
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   4230
      Stretch         =   -1  'True
      Top             =   2415
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Line lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   4380
      X2              =   5655
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Shape Shp 
      FillColor       =   &H80000005&
      Height          =   315
      Index           =   0
      Left            =   4365
      Top             =   1995
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4380
      MouseIcon       =   "frmReport.frx":5335
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_Setup 
         Caption         =   "打印设置(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_Preview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "打印报表(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "Excel数据输出(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile_Graph 
         Caption         =   "Excel图形分析(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Par 
         Caption         =   "条件重置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuEdit_Par_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SetCol 
         Caption         =   "调整列宽(&C)"
         Begin VB.Menu mnuEdit_SetCol_Auto 
            Caption         =   "合适匹配(&A)"
         End
         Begin VB.Menu mnuEdit_SetCol_Fill 
            Caption         =   "补齐表宽(&I)"
         End
         Begin VB.Menu mnuEdit_SetCol_Def 
            Caption         =   "缺省定义(&D)"
         End
      End
      Begin VB.Menu mnuEdit_SelMode 
         Caption         =   "选择模式(&S)"
         Begin VB.Menu mnuEdit_SelMode_Row 
            Caption         =   "整行选择(&R)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuEdit_SelMode_Col 
            Caption         =   "整列选择(&C)"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "视图(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolFormat 
            Caption         =   "报表格式(&F)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolGroup 
            Caption         =   "报表组(&G)"
            Checked         =   -1  'True
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "列表(&L)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuViewStyle_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_Pre 
         Caption         =   "前一张(&P)"
      End
      Begin VB.Menu mnuView_Next 
         Caption         =   "后一张(&N)"
      End
      Begin VB.Menu mnuView_Page_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_reFlash 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "WEB上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPop_Cond 
         Caption         =   "条件1"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop_Split1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop_Save 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu mnuPop_SaveAs 
         Caption         =   "另存为(&A)"
      End
      Begin VB.Menu mnuPop_Del 
         Caption         =   "删除(&C)"
      End
      Begin VB.Menu mnuPop_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop_Default 
         Caption         =   "缺省(&D)"
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mobjCurDLL As clsReport '打开当前报表时所初始化的一个公共报表类(DLL:clsReport,主要用作处理事件)
Public mbytStyle As Byte    '决定报表的显示处理方式

Public mblnDisabledPrint As Boolean     '只预览不允许打印
Public mblnPrintEmpty As Boolean '是否打印无表格数据的格式
Public bytFormat As Byte '当前报表中所打开的格式号
Public marrPars As Variant '公共参数数组,用于本张报表操作

Public frmParent As Object '用于刷新
Public mobjReport As Report '当前报表对象(包含报表组中的当前)

Private arrReport() As Report '报表数组,存放报表组中的多张报表
Private arrLibDatas() As LibDatas '报表组中各个报表的数据源数组(注意：占资源,因为切换格式时不重新读取,所以要每个保留)
Private arrDefPars() As RPTPars '报表组中各个报表的缺省参数内容

Public intReport As Integer '>=0,报表组时当前报表序号,对应picPaper的索引

'出口参数：(用于打印、预览、Excel、图形分析)-------------------------
Public mLibDatas As LibDatas '出：报表中的多个数据源数据,打开报表时产生
Public marrPage As Variant   '出：PageCells集合的数组,打印单元内容描述,用于预览或打印
Public marrPageCard As Variant   '出;卡片的页数
Public mcolRowIDs As New Collection '出：用于记录任意表格的行ID(ID来源于数据源字段,不一定绑定在表格上)

'模块变量------------------------------------------------------------
Private mstrExcelFile As String
Private mblnAllFormat As Boolean
Private lngPreX As Long, lngPreY As Long
Private intGridCount As Integer '当前报表所具有的独立表格数(一个表格整体而言)
Private intGridID As Integer '如果只有一个独立表格,则为其控件ID
Private objCurGrid As Object
Private mobjPars As RPTPars '报表组中处理参数所临时使用
Private mobjDefPars As RPTPars '存放当前报表原始的参数内容,用于恢复缺省值
Private objScript As clsScript
Private blnMatch As Boolean, blnExcel As Boolean
Private blnRefresh As Boolean
Private lngCurInx As Long
Private lngTmpColor As Long
Private mstrPDFFile As String
Private mlngReportID As Long
Private mlngRPTID As Long               '组的子报表ID或独立报表ID
Private mblnLeftClick As Boolean
Private mlngRelationReport As Long
Private mintGridIndex As Integer
Private mintLblIndex As Integer
Private mlngRelationMouseRow As Long
Private mlngRelationMouseCol As Long
Private mlngBackX As Long
Private mlngBackY As Long
Private mbytType As Byte
Private mintCurMenuIndex As Integer
Private mintCurCondID As Integer
Private mobjfrmShow As frmPreview
Private mobjfrmShowDock As frmPreviewDock
Private mlngSys As Long

Private Const CON_SETFOCES As Long = &H9C6D75

Public Sub ShowMe(objParent As Object, objCurDLL As clsReport, arrPars As Variant, ByVal bytStyle As Byte)
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    mlngSys = glngSys
    
    On Error Resume Next
    
    If mbytStyle <> 0 Then
        Load Me
        If Err.Number = 0 Then
            If mbytStyle = 1 Then       '自动预览
                mnuFile_Preview_Click
            ElseIf mbytStyle = 2 Then   '自动打印
                mnuFile_Print_Click
            ElseIf mbytStyle = 3 Then   '输出到Excel
                mnuFile_Excel_Click
            ElseIf mbytStyle = 4 Then   '固定输出到PDF
                mnuFile_Print_Click
            End If
        ElseIf Err.Number <> 0 Then
            '364:对象已卸载(在Form_Load内部Unload,如取消条件窗体)
            Err.Clear
        End If
        Unload Me
    Else
        '先尝试以非模态显示报表
        If frmParent Is Nothing Then
            Me.Show
        ElseIf frmParent.name = "frmDesign" Then
            Me.Show 1, frmParent
        Else
            Me.Show , frmParent
        End If
        
        '两种情况不能以非模态显示
        If Err.Number = 373 Or Err.Number = 401 Then
            '373:不支持编译和设计环境部件的内部操作(源程序调用zlReport.dll,不支持加父窗体)
            '401:当打开有模式窗体时不能显示无模式窗体
            '已自动Load，再显示时不会再激活Form_Load事件
            Err.Clear: Me.Show 1
        ElseIf Err.Number = 364 Then
            '364:对象已卸载(在Form_Load内部Unload,如取消条件窗体)
            Err.Clear
        ElseIf Err.Number <> 0 Then
            Err.Clear: Unload Me '已自动Load，未知错误时卸载窗体
        End If
    End If
End Sub

Private Sub CopyLibDatas(objS As LibDatas, objO As LibDatas)
'功能：拷贝不同报表之间的多个数据源
    Dim tmpData As LibData
    
    Set objO = New LibDatas
    
    For Each tmpData In objS
        objO.Add tmpData.DataName, tmpData.DataSet.Clone, "_" & tmpData.DataName
    Next
End Sub

Private Sub CboFormat_Click()
    Dim strErr As String
    Dim strStartTime As String
    
    If CByte(Mid(cboFormat.SelectedItem.Key, 2)) = bytFormat Then Exit Sub
    bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
    mobjReport.bytFormat = bytFormat
    
    mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).图样 <> 0)
    tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).图样 <> 0)
    
    If mobjReport.blnLoad Then
        If gblnReportRunLog Then
            strStartTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
        End If
        '读取当前格式需要的数据(在已经读取过其它格式数据源的情况下)
        strErr = OpenReportData(False)
        If strErr <> "" Then
            MsgBox "在读取报表数据""" & strErr & """时遇到意外错误,报表不能产生！", vbInformation, App.Title
            Exit Sub
        End If

        '激活条件提交事件
        If Not mobjCurDLL Is Nothing Then
            mobjCurDLL.Act_CommitCondition mobjReport.编号, GetParsStr(MakeNamePars(mobjReport, True)), Me
        End If
        
        Call ShowItems
        If Val(cboFormat.Tag) = 0 Then
            If mlngReportID > 0 Then
                Call RecordsExecute(mlngReportID, strStartTime, 2)
            ElseIf lvw.Visible Then
                Call RecordsExecute(Val(Mid(lvw.SelectedItem.Key, 2)), strStartTime, 2)
            End If
        End If
    End If
End Sub

Private Sub CboFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set cboFormat.SelectedItem = cboFormat.ComboItems("_" & bytFormat)
        KeyAscii = 0
    End If
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Chart_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngX As Single, sngY As Single
    Dim lngSeries As Long, lngPoint As Long, lngDS As Long
    Dim strSeries As String, vArea As RegionConstants
    Dim dblX As Double, strX As String, strY As String
    Dim strLabelX As String, strLabelY As String
    
    With Chart(Index).ChartGroups(1)
        sngX = X / Screen.TwipsPerPixelX
        sngY = Y / Screen.TwipsPerPixelY
        vArea = .CoordToDataIndex(sngX, sngY, oc2dFocusXY, lngSeries, lngPoint, lngDS)
        If vArea = oc2dRegionInChartArea Then
            If lngDS <= 3 Then
                strSeries = ""
                If lngSeries <= .SeriesLabels.count Then
                    strSeries = .SeriesLabels(lngSeries).Text & ":"
                End If
                                
                If .Data.Layout = oc2dDataGeneral Then
                    dblX = .Data.X(lngSeries, lngPoint)
                Else
                    dblX = .Data.X(1, lngPoint)
                End If
                
                If Chart(Index).ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels Then '从1970-01-01 08:00:00开始的秒数
                    strX = Format(DateAdd("s", dblX, CDate("1970-01-01 08:00:00")), "yyyy-MM-dd HH:mm:ss")
                    strX = Replace(strX, " 00:00:00", "")
                    strX = Replace(strX, ":00:00", "")
                    strX = Replace(strX, ":00", "")
                Else
                    strX = dblX
                End If
                strY = .Data.Y(lngSeries, lngPoint)
                
                If Chart(Index).ChartArea.Axes("X").Title.Text <> "" Then
                    strLabelX = Chart(Index).ChartArea.Axes("X").Title.Text & "="
                End If
                If Chart(Index).ChartArea.Axes("Y").Title.Text <> "" Then
                    strLabelY = Chart(Index).ChartArea.Axes("Y").Title.Text & "="
                End If
                
                sta.Panels(3).Text = strSeries & strLabelX & strX & "," & strLabelY & strY
            Else
                sta.Panels(3).Text = ""
            End If
        Else
            sta.Panels(3).Text = ""
        End If
    End With
End Sub

Private Sub cmdDefault_Click()
    Dim sngTop As Single
    
    sngTop = cmdDefault.Top + cmdDefault.Height + picPar.Top + IIF(cbr.Visible, cbr.Height, 0) + 15
    Call Me.PopupMenu(mnuPop, , cmdDefault.Left + 30, sngTop)
End Sub

Private Sub cmdLoad_Click()
    mnuView_reFlash_Click
End Sub

Private Sub cmdSelAll_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        chkTmp.Value = 1
    Next
End Sub

Private Sub cmdSelNone_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        chkTmp.Value = 0
    Next
End Sub

Private Sub Form_Activate()
    Dim tmpMsh As Object
    Static blnAct As Boolean
    
    If blnExcel Then blnExcel = False: Exit Sub
    
    cbr.Bands(2).Width = cbr.Bands(2).Width + 15
    cbr.Bands(2).Width = cbr.Bands(2).Width - 15
    
    '激活事件
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_ReportActive(mobjReport.编号, Me)
    End If
    
    If cbr.Bands(1).Visible Then cbr.Bands(1).MinHeight = tbr.ButtonHeight

    '定位在第一个表格上
    If Not blnAct Then
        blnAct = True
        For Each tmpMsh In msh
            If tmpMsh.Index <> 0 And tmpMsh.Container Is picPaper(intReport) And Not tmpMsh.Tag Like "H_*" Then
                Call msh_EnterCell(tmpMsh.Index)
                On Error Resume Next
                tmpMsh.SetFocus: Exit For
            End If
        Next
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (scrVsc.Visible And scrHsc.Visible) And KeyCode <> vbKeyF3 Then Exit Sub
    Select Case KeyCode
        Case vbKeyUp
            If scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then
                If Shift = 2 Then
                    scrVsc.Value = IIF(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIF(scrVsc.Value - scrVsc.SmallChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.SmallChange)
                End If
            End If
        Case vbKeyDown
            If scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then
                If Shift = 2 Then
                    scrVsc.Value = IIF(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIF(scrVsc.Value + scrVsc.SmallChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.SmallChange)
                End If
            End If
        Case vbKeyLeft
            If scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then
                If Shift = 2 Then
                    scrHsc.Value = IIF(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIF(scrHsc.Value - scrHsc.SmallChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.SmallChange)
                End If
            End If
        Case vbKeyRight
            If scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then
                If Shift = 2 Then
                    scrHsc.Value = IIF(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIF(scrHsc.Value + scrHsc.SmallChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.SmallChange)
                End If
            End If
        Case vbKeyF3
            Call FindItem(txtFind.Text, True)
    End Select
End Sub

Private Sub Form_Load()
    Dim strErr As String, i As Integer, j As Integer
    Dim objItem As Object, rsTmp As ADODB.Recordset
    Dim strPrivs As String, lng程序ID As Long, lng系统ID As Long
    Dim blnPriv As Boolean, bytMode As Byte
    Dim strSQL As String, lngReport As Long
    Dim rsReport As New ADODB.Recordset
    Dim frmNewParInput As New frmParInput
    Dim strBasePrivs As String
    Dim strTmp As String
    Dim strStartTime As String
    Dim rsData As ADODB.Recordset
    Dim lngPersonID As Long
    
    Set objScript = New clsScript
    Srt.AddObject "clsScript", objScript, True
    
    garrBill = Empty
    mblnPrintEmpty = False
    bytFormat = 0
    blnExcel = False

    '获取报表数据
    If gobjReport Is Nothing Then
        '打开报表组
        '显示报表组信息
        Set rsTmp = GetGroupInfo(glngGroup)
        If rsTmp Is Nothing Then Unload Me: Exit Sub '错误退出
        Caption = rsTmp!名称
        lblGroup_S.Caption = lblGroup_S.Caption & ":" & rsTmp!名称
        Me.Tag = rsTmp!编号 '存入报表组编号
        
        lng系统ID = IIF(IsNull(rsTmp!系统), 0, rsTmp!系统)
        lng程序ID = IIF(IsNull(rsTmp!程序id), 0, rsTmp!程序id)
        
        '行列选择模式
        bytMode = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & Me.Tag, "选择模式", 0)
        If bytMode = 1 Then
            Call mnuEdit_SelMode_Col_Click
        Else
            Call mnuEdit_SelMode_Row_Click
        End If
        
        '装入及显示报表
        Set rsTmp = GetSubReport(glngGroup)
        If rsTmp Is Nothing Then
            MsgBox "该报表组中没有任何报表可以执行！", vbInformation, App.Title
            Unload Me: Exit Sub '错误退出
        End If
        Screen.MousePointer = 11
        strPrivs = GetPrivFunc(lng系统ID, lng程序ID)
        i = 0
        Do While Not rsTmp.EOF
            '未授权的子报表，并且非管理工具调用，就不列出该子报表
            If InStr(";" & strPrivs & ";", ";" & Nvl(rsTmp!功能, "NONE") & ";") <= 0 _
                And Not mobjCurDLL Is Nothing Then
                GoTo makContinue
            End If
            
            blnPriv = True
            '合法性判断
            blnPriv = CheckPass(rsTmp!报表ID)
            '权限判断
            If lng程序ID > 0 And Not IsNull(rsTmp!功能) And blnPriv Then
                blnPriv = (InStr(";" & strPrivs & ";", ";" & rsTmp!功能 & ";") > 0)
            End If
            If blnPriv Then
                If i = 0 Then
                    ReDim arrReport(0)
                    ReDim arrLibDatas(0) '数据暂时未装入
                    ReDim arrDefPars(0)
                Else
                    Load picPaper(i): picPaper(i).Visible = False
                    ReDim Preserve arrReport(i)
                    ReDim Preserve arrLibDatas(i) '数据暂时未装入
                    ReDim Preserve arrDefPars(i)
                End If
                
                '报表内容
                Set arrReport(i) = New Report
                Set arrReport(i) = ReadReport(rsTmp!报表ID)
                Call ReplaceSysNo(arrReport(i)) '处理参数定义中的系统变量
                Call GetUserName(arrReport(i).系统, gstrUserName, gstrUserNO)
                Call SetReportIndex(i, arrReport(i))
                
                '缺省参数内容
                Set arrDefPars(i) = New RPTPars
                Set arrDefPars(i) = MakeNamePars(arrReport(i))
                
                Set objItem = lvw.ListItems.Add(, "_" & rsTmp!报表ID, arrReport(i).名称, "Report", "Report")
                objItem.SubItems(1) = arrReport(i).编号
                objItem.SubItems(2) = arrReport(i).说明
                
                 '所有的报表格式中不用的数据源删除
                If arrReport(i).Datas.count > 0 Then Call DelUnUseData(arrReport(i))
                '替换用户通过函数传入的参数
                If ParCount(arrReport(i)) > 0 Then Call ReplaceUserPars(arrReport(i))
                
                '从注册表读上次格式,缺省1
                arrReport(i).bytFormat = CByte(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & arrReport(i).编号, "格式", 1))
                
                i = i + 1
            End If
            
makContinue:
            rsTmp.MoveNext
        Loop
        
        Screen.MousePointer = 0
        If rsTmp.RecordCount > 0 And lvw.ListItems.count = 0 Then
            MsgBox "你没有权限执行该报表组中的所有报表，请检查报表的合法性以及是否正确授权！", vbInformation, App.Title
            Unload Me: Exit Sub '错误退出
        ElseIf lvw.ListItems.count = 0 Then
            Unload Me: Exit Sub '错误退出
        End If
        
        mnuEdit_Par.Visible = False
        mnuEdit_Par_.Visible = False
        tbr.Buttons("Par").Caption = "重装"
        tbr.Buttons("Par").Tag = "重装"
        
        lvw.ColumnHeaders(2).Position = 1
        RestoreWinState Me, App.ProductName, Me.Tag

        SetView lvw.View
        
        '设置报表列表高度，以尽量增加条件区域高度
        lvw.Height = lvw.ListItems.count * 350
        If lvw.Height < 1000 Then lvw.Height = 1000
        If lvw.Height > picGroup.Height / 2 Then
            lvw.Height = picGroup.Height / 2
        End If
        
        picLR_S.Visible = mnuViewToolGroup.Checked
        picGroup.Visible = mnuViewToolGroup.Checked
        
        If Not lvw.SelectedItem Is Nothing Then Call lvw_ItemClick(lvw.SelectedItem)
    Else
        '打开单独报表
        picBack.BorderStyle = 0
        picLR_S.Visible = False
        picGroup.Visible = False
        For i = 0 To mnuViewStyle.UBound
            mnuViewStyle(i).Visible = False
        Next
        mnuViewStyle_.Visible = False
        mnuView_Pre.Visible = False
        mnuView_Next.Visible = False
        mnuView_Page_.Visible = False
        mnuViewToolGroup.Visible = False

        tbr.Buttons("Style").Visible = False
        tbr.Buttons("Style_").Visible = False
        tbr.Buttons("Pre").Visible = False
        tbr.Buttons("Next").Visible = False
        tbr.Buttons("Page_").Visible = False
        
        intReport = 0
        Call CopyReport(gobjReport, mobjReport)
        Call ReplaceSysNo(mobjReport) '处理参数定义中的系统变量
        Call GetUserName(mobjReport.系统, gstrUserName, gstrUserNO)
        Call SetReportIndex(intReport, mobjReport)
        Caption = mobjReport.名称
        
        If mbytStyle = 0 Then '不显示窗体时就不处理以加快速度
            RestoreWinState Me, App.ProductName, mobjReport.编号
        End If
        
        '查询是否允许在当前时间执行此报表
        If Format(mobjReport.禁止结束时间, "HH:mm:ss") <> "00:00:00" Or Format(mobjReport.禁止开始时间, "HH:mm:ss") <> "00:00:00" Then
            If CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss")) > CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")) Then
                If Between(CDate(Format(Currentdate, "HH:mm:ss")), CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")), CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss"))) Then
                    MsgBox "当前报表在" & CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")) & "-" & CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss")) & "禁止执行，如有疑问请联系信息科。", vbInformation, App.Title
                    Unload Me: Exit Sub
                End If
            Else
                If CDate(Format(Currentdate, "HH:mm:ss")) < CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")) Or CDate(Format(Currentdate, "HH:mm:ss")) > CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss")) Then
                    MsgBox "当前报表在" & CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")) & "-第二天" & CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss")) & "禁止执行，如有疑问请联系信息科。", vbInformation, App.Title
                    Unload Me: Exit Sub
                End If
            End If
        End If
        
        '行列选择模式
        bytMode = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.编号, "选择模式", 0)
        If bytMode = 1 Then
            Call mnuEdit_SelMode_Col_Click
        Else
            Call mnuEdit_SelMode_Row_Click
        End If
        
        '记录报表执行开始时间
        If gblnReportRunLog Then
            strStartTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
        End If
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_BeforeReportLoad(mobjReport.编号, Me)
        End If
    
         '所有的报表格式中不用的数据源删除
        If mobjReport.Datas.count > 0 Then Call DelUnUseData(mobjReport)
    
        '缺省显示第一种格式
        bytFormat = 1
        
        '根据本地打印设置读取要打印的格式
        '如果是程序内部调用打印用要打印所有格式，则当前默认为第一种格式
        strTmp = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\LocalSet\" & mobjReport.编号, "AllFormat", "")
        If strTmp = "" Then strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & mobjReport.编号, "AllFormat", 0)
        mblnAllFormat = Val(strTmp) = 1
        If Not (mbytStyle = 2 And mblnAllFormat) Then
            '根据注册表取用户格式
            bytFormat = CByte(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.编号, "格式", 1))
            '读取打印设置指定的格式 'If mobjReport.票据 Then
            strTmp = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\LocalSet\" & mobjReport.编号, "Format", "")
            If strTmp = "" Then
                i = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & mobjReport.编号, "Format", -1))
            Else
                i = Val(strTmp)
            End If
            If i <> -1 Then bytFormat = i
        End If
        
        '取该报表的ID
        lngReport = 0
        If ReportReaded(, mobjReport.编号, mobjReport.系统) Then
            lngReport = grsReport!id '利用缓存
        Else
            strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间 From zlReports Where 编号=[1] And Nvl(系统,0)=[2]"
            Set rsReport = OpenSQLRecord(strSQL, Me.Caption, mobjReport.编号, mobjReport.系统)
            If Not rsReport.EOF Then '缓存处理
                Set grsReport = New ADODB.Recordset
                Set grsReport = rsReport
                gdatModiTime = grsReport!修改时间
                
                lngReport = rsReport!id
            End If
        End If
        mlngReportID = lngReport
        mlngRPTID = lngReport
        
        '根据用户传入参数取一些控制参数
        If IsArray(marrPars) Then
            If UBound(marrPars) <> -1 Then
                For i = 0 To UBound(marrPars)
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        'ReportFormat
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("ReportFormat") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                bytFormat = CByte(Trim(Mid(CStr(marrPars(i)), j + 1)))
                                mblnAllFormat = False '程序指定了格式则打印设置无效
                            End If
                        'DisabledPrint
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("DisabledPrint") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                mblnDisabledPrint = CByte(Trim(Mid(CStr(marrPars(i)), j + 1))) = 1
                            End If
                        'PrintEmpty
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("PrintEmpty") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                mblnPrintEmpty = CByte(Trim(Mid(CStr(marrPars(i)), j + 1))) = 1
                            End If
                        'ExcelFile
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("ExcelFile") Then
                            mstrExcelFile = Trim(Mid(CStr(marrPars(i)), j + 1))
                        'PDF
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("PDF") Then
                            mstrPDFFile = Trim(Mid(CStr(marrPars(i)), j + 1))
                        End If
                    End If
                Next
            End If
        End If
        
        '加入现有格式
        For i = 1 To mobjReport.Fmts.count
            Set objItem = cboFormat.ComboItems.Add(, "_" & mobjReport.Fmts(i).序号, mobjReport.Fmts(i).说明, "Format")
            If mobjReport.Fmts(i).序号 = bytFormat Then objItem.Selected = True
        Next
        If cboFormat.SelectedItem Is Nothing And cboFormat.ComboItems.count > 0 Then
            cboFormat.ComboItems(1).Selected = True
            bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
        End If
        mobjReport.bytFormat = bytFormat
        mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).图样 <> 0)
        tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).图样 <> 0)
                
'        If cboFormat.ComboItems.Count = 1 Then
'            mnuViewToolFormat.Checked = False
'            cbr.Bands(2).Visible = False
'        End If
        cboFormat.Locked = cboFormat.ComboItems.count > 1
                
        '条件输入
        If ParCount(mobjReport) > 0 Then
            If Not ReplaceUserPars(mobjReport) Then
                '未全部正确地传入参数,则要求输入参数
                
                Set mobjPars = MakeNamePars(mobjReport)
                Call CopyPars(mobjPars, mobjDefPars)
                frmNewParInput.mlngReport = lngReport
                Set frmNewParInput.mobjPars = mobjPars
                Set frmNewParInput.mobjDefPars = mobjDefPars
                Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
                
                frmNewParInput.mstrTitle = mobjReport.名称
                frmNewParInput.mblnReset = False
                frmNewParInput.Show 1, Me
                
                If frmNewParInput.mblnOK Then
                    '激活条件提交事件
                    If Not mobjCurDLL Is Nothing Then
                        mobjCurDLL.Act_CommitCondition mobjReport.编号, GetParsStr(frmNewParInput.mobjPars), Me
                    End If
                    
                    ReplaceInputPars frmNewParInput.mobjPars
                    Unload frmNewParInput
                Else
                    Unload Me: Exit Sub '第一次取消则退出
                End If
            Else
                '全部正确传入参数,则也不能重置报表条件
                tbr.Buttons("Par").Visible = False
                mnuEdit_Par.Visible = False
                tbr.Buttons("Par_").Visible = False
                mnuEdit_Par_.Visible = False
                
                Set mobjDefPars = MakeNamePars(mobjReport)
                
                '激活条件提交事件
                If Not mobjCurDLL Is Nothing Then
                    mobjCurDLL.Act_CommitCondition mobjReport.编号, GetParsStr(MakeNamePars(mobjReport, True)), Me
                End If
            End If
        Else
            '激活条件提交事件
            If Not mobjCurDLL Is Nothing Then
                mobjCurDLL.Act_CommitCondition mobjReport.编号, "ReportFormat=" & bytFormat, Me
            End If
            
            '如果没有定义参数,则也不能重置报表条件
            tbr.Buttons("Par").Visible = False
            mnuEdit_Par.Visible = False
            tbr.Buttons("Par_").Visible = False
            mnuEdit_Par_.Visible = False
        End If
        
        '使用者传入参数或输入参数都保存在报表参数的缺省值中(mobjReport)
        '产生数据
        If Not frmParent Is Nothing Then frmParent.Refresh
        Me.Refresh
        strErr = OpenReportData(False)
        If strErr <> "" Then
            If gblnSilentMode = False Then
                MsgBox "在读取报表数据""" & strErr & """时遇到意外错误,报表不能产生！", vbInformation, App.Title
            End If
            Unload Me: Exit Sub
        End If
        
        '显示报表
        Call ShowItems
       
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_AfterReportLoad(mobjReport.编号, Me)
        End If
        Call RecordsExecute(lngReport, strStartTime, 3)
        
    End If
    
    'Excel输出、打印权限判断
    strBasePrivs = GetPrivFunc(0, 16)
    If InStr(";" & strBasePrivs & ";", ";Excel输出;") = 0 Then
        mnuFile_Excel.Visible = False
        mnuFile_Graph.Visible = False
        mnuFile_1.Visible = False
        tbr.Buttons("Graph").Visible = False
        tbr.Buttons(5).Visible = False
    End If
    If InStr(";" & strBasePrivs & ";", ";打印;") = 0 Or mblnDisabledPrint Then
        mnuFile_Print.Visible = False
        tbr.Buttons("Print").Visible = False
    End If
    
End Sub

Private Sub RecordsExecute(ByVal lngReportID As Long, ByVal strStartTime As String, _
    Optional ByVal intType As Integer = 0)
'功能：记录报表执行
'参数：
'  lngReportID：报表ID
'  strStartTime：
'  intType：要记录的类型；2-报表运行日志；3-报表使用状态和报表运行日志

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnRunLog As Boolean
    Dim strEndTime As Date
    
    If intType <= 0 Then Exit Sub
    If Not (gblnReportUse Or gblnReportRunLog) Then Exit Sub
    
    On Error GoTo ErrHand
    strEndTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
    Select Case intType
    Case 2
        GoSub makTwo
    Case 3
        GoSub makOne
        GoSub makTwo
    End Select
    Exit Sub
    
makOne:
    If gblnReportUse Then
        strSQL = "Zl_Rptrun_Update(" & lngReportID & "," & _
                    "'" & gstrUserName & "')"
        Call ExecuteProcedure(strSQL, "报表运行记录")
    End If
    Return
    
makTwo:
    If gblnReportRunLog Then
        strSQL = "Zl_Rptrunhistory_Update(" & _
                    lngReportID & "," & _
                    "'" & gstrUserName & "'," & _
                    "to_date('" & strStartTime & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "to_date('" & strEndTime & "','YYYY-MM-DD HH24:MI:SS'))"
        Call ExecuteProcedure(strSQL, "报表运行日志")
    End If
    Return
    
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lbl_Click(Index As Integer)
    Dim objRelations As RPTRelations
    Dim i As Long
    Dim lngRec As Long, strDataName As String
    Dim strLisName As String
    Dim lngRelationReport As Long
    Dim tmpData As RPTData, tmpPar As RPTPar
    
    If mblnLeftClick = False And mlngRelationReport = 0 Then Exit Sub
    
    If lbl(Index).Tag <> "" Then
        Set objRelations = mobjReport.Items("_" & Index).Relations
        lngRec = Val(lbl(Index).Tag)
        For i = 1 To objRelations.count
            If objRelations.Item(i).默认 = 1 Then
                lngRelationReport = objRelations.Item(i).关联报表ID
                Exit For
            End If
        Next
        If lngRelationReport = 0 Then lngRelationReport = objRelations.Item(1).关联报表ID
        If mlngRelationReport <> 0 Then lngRelationReport = mlngRelationReport
        If Not CheckReportPriv(lngRelationReport) Then
            MsgBox "你没有权限查询该报表某些数据源中的对象！", vbInformation, App.Title: Exit Sub
        End If
        '执行报表
        If CheckPass(lngRelationReport) = False Then
            MsgBox "报表数据错误，不能执行该报表！", vbInformation, App.Title: Exit Sub
        End If
        
        Set gobjReport = ReadReport(lngRelationReport)
        '初始化参数
        garrPars = Array()
        '定位记录集
        On Error Resume Next
        For i = 1 To objRelations.count
            If objRelations.Item(i).关联报表ID = lngRelationReport Then
                If InStr(objRelations.Item(i).参数值来源, ".") > 0 Then
                    strDataName = Mid(objRelations.Item(i).参数值来源, 1, InStr(objRelations.Item(i).参数值来源, ".") - 1)
                End If
            End If
            If strDataName <> "" Then Exit For
        Next

        '点击的具体数据行
        If strDataName <> "" Then mLibDatas("_" & strDataName).DataSet.AbsolutePosition = lngRec
        
        For i = 1 To objRelations.count
            With objRelations.Item(i)
                strLisName = ""
                If objRelations.Item(i).关联报表ID = lngRelationReport Then
                    If InStr(.参数值来源, ".") > 0 Then
                        If mLibDatas("_" & strDataName).DataSet.RecordCount > 0 Then
                            strLisName = mLibDatas("_" & strDataName).DataSet.Fields(Mid(.参数值来源, InStr(.参数值来源, ".") + 1)).Value
                        End If
                    ElseIf InStr(.参数值来源, "=") = 1 Then
                        For Each tmpData In mobjReport.Datas
                            For Each tmpPar In tmpData.Pars
                                If tmpPar.名称 = Mid(.参数值来源, 2) Then
                                    strLisName = tmpPar.缺省值
                                    Exit For
                                End If
                            Next
                            If strLisName <> "" Then Exit For
                        Next
                    End If
                    ReDim Preserve garrPars(UBound(garrPars) + 1)
                    garrPars(UBound(garrPars)) = .参数名 & "=" & strLisName
                End If
            End With
        Next
        
        
        If Not ShowReport(Me) Then MsgBox "报表打开失败！", vbInformation, App.Title
    End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl(Index).Tag <> "" Then lbl(Index).MousePointer = 99
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim dbRowHeight As Double

    If Button = 2 Then
        mintLblIndex = Index
        mblnLeftClick = False
        If lbl(Index).FontUnderline = True Then
            Call LoadRelation(1, Index)
            vsfRelations.Visible = True
            vsfRelations.SetFocus
            For i = 0 To vsfRelations.Rows - 1
                dbRowHeight = dbRowHeight + vsfRelations.RowHeight(i)
            Next
            vsfRelations.Height = dbRowHeight
            vsfRelations.Left = lbl(Index).Left + X + 150
            vsfRelations.Top = lbl(Index).Top + 90
        Else
            vsfRelations.Visible = False
        End If
    Else
        mblnLeftClick = True
    End If
End Sub

Private Sub lblPar_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lngPreY = Y
End Sub

Private Sub lblPar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If lblPar_S.Top + Y - lngPreY < 1000 Or picPar.Height - (Y - lngPreY) < 1000 Then Exit Sub
        lblPar_S.Top = lblPar_S.Top + Y - lngPreY
        lvw.Height = lvw.Height + Y - lngPreY
        picPar.Top = picPar.Top + Y - lngPreY
        picPar.Height = picPar.Height - (Y - lngPreY)
        Me.Refresh
    End If
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, objItem As Object
    
    Set objCurGrid = Nothing
    
    LockWindowUpdate Me.hwnd
    
    '保存前一张报表的数据状态
    If Not mobjReport Is Nothing Then
        Call CopyReport(mobjReport, arrReport(intReport)) '报表
        Call CopyPars(mobjDefPars, arrDefPars(intReport)) '缺省参数内容
        If mLibDatas Is Nothing Then '数据源数据
            Set arrLibDatas(intReport) = Nothing
        Else
            Call CopyLibDatas(mLibDatas, arrLibDatas(intReport))
        End If
    End If
    
    '获取当前报表数据状态
    intReport = Item.Index - 1
    Call CopyReport(arrReport(intReport), mobjReport) '报表
    Call CopyPars(arrDefPars(intReport), mobjDefPars) '缺省参数内容
    If arrLibDatas(intReport) Is Nothing Then '数据源数据
        Set mLibDatas = Nothing
    Else
        Call CopyLibDatas(arrLibDatas(intReport), mLibDatas)
    End If
    
    bytFormat = mobjReport.bytFormat
    intGridCount = mobjReport.intGridCount
    intGridID = mobjReport.intGridID
        
    '加入格式
    cboFormat.ComboItems.Clear
    For i = 1 To mobjReport.Fmts.count
        Set objItem = cboFormat.ComboItems.Add(, "_" & mobjReport.Fmts(i).序号, mobjReport.Fmts(i).说明, "Format")
        If mobjReport.Fmts(i).序号 = bytFormat Then objItem.Selected = True
    Next
    If cboFormat.SelectedItem Is Nothing And cboFormat.ComboItems.count > 0 Then
        cboFormat.ComboItems(1).Selected = True
        bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
        mobjReport.bytFormat = bytFormat
    End If
    cboFormat.Refresh
    cboFormat.Locked = cboFormat.ComboItems.count > 1
   
    mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).图样 <> 0)
    tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).图样 <> 0)
    
    '显示当前纸张
    picBack.Visible = False
    For i = 0 To picPaper.UBound
        picPaper(i).Visible = (i = intReport)
    Next
    picPaper(intReport).ZOrder
    
    scrVsc.Visible = Not (intGridCount = 1 And Not mobjReport.票据)
    scrHsc.Visible = Not (intGridCount = 1 And Not mobjReport.票据)
    picShadow.Visible = Not (intGridCount = 1 And Not mobjReport.票据)
    If Not (intGridCount = 1 And Not mobjReport.票据) Then
        scrVsc.Value = scrVsc.Min
        scrHsc.Value = scrHsc.Min
        Call scrhsc_Change
        Call scrVsc_Change
    End If
    
    '调整页面
    Call Form_Resize
    picBack.Visible = True

    '显示参数
    picPar.Visible = False

    Call CopyPars(mobjDefPars, mobjPars)
    mlngRPTID = Val(Mid(lvw.SelectedItem.Key, 2))
    Call InitReportPars
    picPar.Visible = True
    
    LockWindowUpdate 0

    '定位在第一个表格上
    For Each objItem In msh
        If objItem.Index <> 0 And objItem.Container Is picPaper(intReport) And Not objItem.Tag Like "H_*" Then
            Call msh_EnterCell(objItem.Index)
            Exit For
        End If
    Next
End Sub

Private Sub mnuEdit_SelMode_Col_Click()
    Dim tmpMsh As Object
    
    '(报表组中)所有报表
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_LeaveCell(tmpMsh.Index)
        End If
    Next
    
    mnuEdit_SelMode_Col.Checked = True
    mnuEdit_SelMode_Row.Checked = False
    
    '(报表组中)所有报表
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
        End If
        '列选择
        msh(tmpMsh.Index).SelectionMode = flexSelectionByColumn
    Next
End Sub

Private Sub mnuEdit_SelMode_Row_Click()
    Dim tmpMsh As Object
    
    '(报表组中)所有报表
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_LeaveCell(tmpMsh.Index)
        End If
    Next
    
    mnuEdit_SelMode_Row.Checked = True
    mnuEdit_SelMode_Col.Checked = False
    
    '(报表组中)所有报表
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
        End If
        '行选择
        msh(tmpMsh.Index).SelectionMode = flexSelectionByRow
    Next
End Sub

Private Sub mnuEdit_SetCol_Auto_Click()
'功能：调整表格列宽为最小适应文字宽度,以最后一行固定行为准(如果有)
    Dim tmpMsh As Object
    Dim i As Integer

    If Not mobjReport.blnLoad Then Exit Sub
    
    On Error Resume Next
    
    Screen.MousePointer = 11
    
    LockWindowUpdate Me.hwnd
    
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 Then
            If tmpMsh.Container Is picPaper(intReport) Then
                Call SetColWidth(tmpMsh)
            ElseIf UCase(tmpMsh.Container.name) = "PIC" Then
                If tmpMsh.Container.Container Is picPaper(intReport) Then
                    Call SetColWidth(tmpMsh)
                End If
            End If
        End If
    Next
    '表头与表体相同列的宽度取较大者
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") _
            And tmpMsh.FixedRows = 0 Then
            For i = 0 To tmpMsh.Cols - 1
                If tmpMsh.ColWidth(i) > msh(tmpMsh.Tag).ColWidth(i) Then
                    msh(tmpMsh.Tag).ColWidth(i) = tmpMsh.ColWidth(i)
                Else
                    tmpMsh.ColWidth(i) = msh(tmpMsh.Tag).ColWidth(i)
                End If
            Next
            tmpMsh.LeftCol = 0: msh(tmpMsh.Tag).LeftCol = 0
            
            '重整行高
            Call AdjustRowHight(tmpMsh.Index)
        End If
    Next
    
    Screen.MousePointer = 0
    LockWindowUpdate 0
End Sub

Private Sub mnuEdit_SetCol_Def_Click()
'功能：还原表格列宽为缺省设计列宽
    Dim objItem As RPTItem, objCurItem As RPTItem
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim i As Integer, j As Integer, strWidth As String
    Dim lngColB As Long, lngColE As Long
    
    If Not mobjReport.blnLoad Then Exit Sub
        
    On Error Resume Next
    
    LockWindowUpdate Me.hwnd
    
    For Each objItem In mobjReport.Items
        If objItem.格式号 = bytFormat Then
            If objItem.类型 = 4 Then
                With objItem
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.id)
                        msh(.id).ColWidth(tmpItem.序号) = tmpItem.W
                        msh(.SubIDs(1).id).ColWidth(tmpItem.序号) = tmpItem.W
                        msh(.id).LeftCol = 0: msh(.SubIDs(1).id).LeftCol = 0
                    Next
                    '重整行高
                    Call AdjustRowHight(objItem.id)
                End With
            ElseIf objItem.类型 = 5 And objItem.性质 = 0 Then
                For i = 0 To UBound(Split(objItem.表头, "|"))
                    Set objCurItem = mobjReport.Items("_" & Split(Split(objItem.表头, "|")(i), ",")(0))
                    With objCurItem
                        strWidth = ""
                        For Each tmpID In .SubIDs
                            Set tmpItem = mobjReport.Items("_" & tmpID.id)
                            Select Case tmpItem.类型
                                Case 7
                                    If i = 0 Then msh(objItem.id).ColWidth(tmpItem.序号) = tmpItem.W
                                Case 9
                                    strWidth = strWidth & "," & tmpItem.W
                            End Select
                        Next
                        strWidth = Mid(strWidth, 2)
                        
                        If i = 0 Then
                            lngColB = msh(objItem.id).FixedCols
                        Else
                            lngColB = lngColE + 1
                        End If
                        lngColE = CLng(Split(Split(objItem.表头, "|")(i), ",")(1)) - 1
                        
                        For j = lngColB To lngColE
                            msh(objItem.id).ColWidth(j) = _
                                CLng(Split(strWidth, ",")((j - lngColB) Mod (UBound(Split(strWidth, ",")) + 1)))
                        Next
                    End With
                Next
            End If
        End If
    Next
    
    '针对附加表体特殊处理
    Call SetGridAlign
    
    LockWindowUpdate 0
End Sub

Private Sub mnuEdit_SetCol_Fill_Click()
'功能：自动调整列宽(按当前各列比例补齐表格宽度,且附加体中各表右边列对齐)
    Dim tmpMsh As VSFlexGrid
    Dim i As Integer
    Dim lngCurW As Long
    Dim sngScale As Single
    
    If Not mobjReport.blnLoad Then Exit Sub
    
    On Error Resume Next
    
    LockWindowUpdate Me.hwnd
    
    Call SetGridAlign(Val("1-补齐"))
    
    LockWindowUpdate 0
    
'    '补宽
'    For Each tmpMsh In msh
'        If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") Then
'            tmpMsh.Redraw = False
'
'            lngCurW = GetGridColWidth(tmpMsh)
'            If lngCurW < tmpMsh.Width - 300 Then
'                '计算页面表格宽度与设计界面表格的比例
'                sngScale = (tmpMsh.Width - 300) / lngCurW
'                For i = 0 To tmpMsh.Cols - 1
'                    tmpMsh.ColWidth(i) = tmpMsh.ColWidth(i) * sngScale
'                Next
'            End If
'
'            '重整行高
'            Call AdjustRowHight(tmpMsh.Index)
'
'            tmpMsh.Redraw = True
'        End If
'    Next
'
'    LockWindowUpdate 0
End Sub

Private Sub mnuFile_Graph_Click()
    Dim objHead As Object
    Dim objItem As RPTItem
    Dim bytKind As Byte
    Dim tmpMsh As Object
    
    If Not mobjReport.blnLoad Then Exit Sub
    
    If zlRegInfo("授权性质") <> "1" Then
        MsgBox "试用或测试版本不能使用该功能。", vbInformation, App.Title
        Exit Sub
    End If
    
    If intGridCount = 0 Then
        MsgBox "当前报表中没有数据表可供图形分析！", vbInformation, App.Title
        Exit Sub
    End If
    If objCurGrid Is Nothing Then
        If msh.count > 1 Then
            For Each tmpMsh In msh
                If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") And Not tmpMsh.Tag Like "H_*" Then
                    Set objCurGrid = tmpMsh
                    Exit For
                End If
            Next
        End If
        If objCurGrid Is Nothing Then
            MsgBox "请先选择一个要进行图形分析的数据表！", vbInformation, App.Title
            Exit Sub
        End If
    End If
    If objCurGrid.Tag Like "H_*" Then
        MsgBox "数据表头不能用作图形分析！", vbInformation, App.Title
        Exit Sub
    End If
    
    Set objItem = mobjReport.Items("_" & objCurGrid.Index)
    If objItem.类型 = 4 Then
        bytKind = GetGridStyle(mobjReport, objItem.id)
        If bytKind = 0 Then Set objHead = msh(CInt(objCurGrid.Tag))
    End If
    blnExcel = True
    Call ExcelChart(Me, objCurGrid, objHead, IIF(mobjReport.Items("_" & objCurGrid.Index).类型 = 5, 1, 2), mobjReport.名称, mobjReport.Fmts("_" & bytFormat).图样)
End Sub

Private Sub mnuFile_Preview_Click()
    Dim frmShow As New frmPreview

    If Not mobjReport.blnLoad Then Exit Sub
    
    If mobjReport.Items.count = 0 Then Exit Sub
    
    If Not InitPrinter(Me) Then
        gblnError = True
        MsgBox "设备初始化失败.可能是系统没有安装打印机或与当前设置不兼容！", vbInformation, App.Title: Exit Sub
    End If
    
    If Not CalcCellPage Then
        gblnError = True
        MsgBox "无法处理的表格格式,操作不能继续！", vbInformation, App.Title: Exit Sub
    End If
    If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
        lbl(lngCurInx).BackColor = lngTmpColor
        lngCurInx = 0: lngTmpColor = 0
    End If
    
    SetRedraw False
    
    Set frmShow.frmParent = Me
    
    If mbytStyle = Val("1-自动预览") Then
        If Not frmParent Is Nothing Then
            On Error Resume Next
            frmShow.Show 1, frmParent
            If Err.Number <> 0 Then
                On Error GoTo 0
                frmShow.Show 1
            End If
            On Error GoTo 0
        Else
            frmShow.Show 1
        End If
    Else
        frmShow.Show 1, Me
    End If
    
    SetRedraw True
End Sub

Private Sub mnuFile_Print_Click()
    Dim objItem As RPTItem, strSource As String
    Dim lngPrintH As Long, blnReset As Boolean
    Dim blnExit As Boolean, intCopy As Integer
    Dim blnDo As Boolean, blnCancel As Boolean
    Dim k As Integer, i As Integer, j As Integer
    Dim arrBill As Variant, strItem As String
    Dim objFmt As RPTFmt, blnGoOn As Boolean
    Dim blnPrint As Boolean, blnALLEmpty As Boolean
    Dim strTmp As String
    Dim strDefault As String
    Dim lngEndPage As Long
    
    If Not mobjReport.blnLoad Then Exit Sub
    If mobjReport.Items.count = 0 Then Exit Sub
    
    If Not mobjCurDLL Is Nothing Then
        mobjCurDLL.DataIsEmpty = False
    End If
    blnALLEmpty = True
    
    strDefault = mobjReport.Fmts(mobjReport.bytFormat).说明
    strTmp = GetRegPrinterInfo("PaperCopy", mobjReport.编号, strDefault)
    intCopy = Val(strTmp)
    If intCopy < 1 Then intCopy = 1
    If gblnSingleTask Then intCopy = 1 '多报表单任务打印时不支持打印份数
    If mobjReport.票据 Then intCopy = 1 '如果是票据，则只能打印1份
    
    cboFormat.Tag = "1"
    blnGoOn = True
    Do While blnGoOn
        Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)

        '直接打印时,当前格式的表格都为空时,则不打印
        blnExit = False
        'mblnPrintEmpty=Fale表示程序不强制打印空表； 打印方式=1表示不“空表打印”
        If mblnPrintEmpty = False And mobjReport.打印方式 = Val("1-空表不打印") And InStr(";0;2;4;", ";" & mbytStyle & ";") > 0 Then
            strSource = ""
            For Each objItem In mobjReport.Items
                If objItem.格式号 = bytFormat Then
                    If objItem.类型 = 4 Then        '任意表格
                        strItem = GetGridSource(objItem, True) '"病人信息,药品信息,..."
                        If strItem <> "" Then strSource = strSource & "," & strItem
                    ElseIf objItem.类型 = 5 Then    '汇总表格
                        strSource = strSource & "," & objItem.内容
                    End If
                End If
            Next
            '使用了数据源(或表格)才判断,否则可能只打印一些标签之类的
            If strSource <> "" Then
                blnExit = True
                strSource = Mid(strSource, 2)
                For i = 0 To UBound(Split(strSource, ","))
                    On Error Resume Next
                    blnExit = blnExit And mLibDatas("_" & Split(strSource, ",")(i)).DataSet.RecordCount = 0
                    Err.Clear: On Error GoTo 0
                Next
            End If
            If blnExit Then GoTo NextFormat
        End If
        blnALLEmpty = False
        
        On Error GoTo errH
        
        If mbytStyle = Val("4-PDF") Then
            If PDFInitialize(objFmt) Then
                If PDFFile(mstrPDFFile, , , True) = False Then
                    MsgBox "未指定PDF输出路径和文件名，请检查！", vbInformation, App.Title
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        Else
            '初始化打印机信息
            If Not InitPrinter(Me, intCopy) Then
                MsgBox "设备初始化失败.可能是系统没有安装打印机或与当前设置不兼容！", vbInformation, App.Title
                gblnError = True: GoTo ExitHandle
            End If
        End If
        
        k = intCopy '缺省为强制循环打印k份
        If Printer.Copies = intCopy Then k = 1 '支持时使用打印机功能
        
        '计算打印内容
        If Not CalcCellPage Then
            gblnError = True
            MsgBox "无法处理的表格格式,操作不能继续！", vbInformation, App.Title: GoTo ExitHandle
        End If
        If mbytStyle <> 2 And mbytStyle <> 4 Then
            If MsgBox("报表内容即将输出,打印设备准备就绪吗？", vbQuestion + vbYesNo, App.Title) = vbNo Then GoTo ExitHandle
        End If
        
        If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
            lbl(lngCurInx).BackColor = lngTmpColor
            lngCurInx = 0: lngTmpColor = 0
        End If
        
        '输出到打印之前，激活打印事件
        If Not mobjCurDLL Is Nothing Then
            arrBill = Empty: blnCancel = False: i = 1
            If IsArray(marrPage) Then i = UBound(marrPage) + 1
            Call mobjCurDLL.Act_BeforePrint(mobjReport.编号, i * intCopy, blnCancel, arrBill)
            If blnCancel Then GoTo ExitHandle
            
            '实际要打印的票据数据
            If IsArray(arrBill) Then garrBill = arrBill
        End If
        
        SetRedraw False
        
        '直接打印报表
        If mbytStyle <> 2 Then Screen.MousePointer = 11
        
        j = 0
        blnReset = False
        Do
            k = k - 1
            j = j + 1
            If Not IsArray(marrPage) Then
                If IsArray(marrPageCard) Then
                    '卡片
                    GoTo makPage
                End If
                
                If mbytStyle <> Val("2-打印") Then
                    If Printer.Copies <> intCopy And intCopy <> 1 Then
                        ShowFlash "输出" & mobjReport.名称 & ",共 1 页 " & intCopy & " 份,当前第 " & j & " 份", j / intCopy, Me
                    Else
                        ShowFlash "输出" & mobjReport.名称 & "…", 1, Me
                    End If
                End If
                
                '动态计算及设置纸张高度
                If objFmt.动态纸张 And objFmt.纸向 = 1 Then
                    Call PrintPage(0, Me, Me, 1, False, True, lngPrintH)
                    blnDo = lngPrintH > 0 And lngPrintH < objFmt.H
                    If blnDo Then '空白部份高于30mm且高于原纸张的1/8
                        blnDo = objFmt.H - lngPrintH > 30 * Twip_mm And objFmt.H - lngPrintH > objFmt.H / 8
                    End If
                    If blnDo Then
                        lngPrintH = lngPrintH + 567 '比实际打印多留10mm高度
                        If Not SetPrinterPaper(Me.hwnd, mobjReport, lngPrintH, intCopy) Then
                            '设置失败时恢复成原始纸张
                            Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                        End If
                    End If
                End If
                
                blnPrint = True
                Call PrintPage(0, Printer, Me)
            Else
makPage:
                If IsArray(marrPage) Then
                    lngEndPage = UBound(marrPage)
                ElseIf IsArray(marrPageCard) Then
                    lngEndPage = UBound(marrPageCard)
                Else
                    lngEndPage = -1
                End If
                
                For i = 0 To lngEndPage
                    If mbytStyle <> 2 Then
                        If Printer.Copies <> intCopy And intCopy <> 1 Then
                            ShowFlash "输出" & mobjReport.名称 & ",共 " & lngEndPage + 1 & " 页 " & intCopy & " 份,当前第 " & j & " 份", ((i + 1) + ((j - 1) * (lngEndPage + 1))) / ((lngEndPage + 1) * intCopy), Me
                        Else
                            ShowFlash "输出" & mobjReport.名称 & ",共 " & lngEndPage + 1 & " 页,当前第 " & i + 1 & " 页…", (i + 1) / (lngEndPage + 1), Me
                        End If
                    End If
                    
                    '动态计算及设置纸张高度
                    If objFmt.动态纸张 And objFmt.纸向 = 1 Then
                        Call PrintPage(i, Me, Me, 1, False, True, lngPrintH)
                        blnDo = lngPrintH > 0 And lngPrintH < objFmt.H
                        If blnDo Then '空白部份高于30mm且高于原纸张的1/8
                            blnDo = objFmt.H - lngPrintH > 30 * Twip_mm And objFmt.H - lngPrintH > objFmt.H / 8
                        End If
                        If blnDo Then
                            lngPrintH = lngPrintH + 567 '比实际打印多留10mm高度
                            If Not SetPrinterPaper(Me.hwnd, mobjReport, lngPrintH, intCopy) Then
                                '设置失败时恢复成原始纸张
                                Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                                blnReset = False
                            Else
                                blnReset = True '本页已设置过动态纸张,下页计算出不必设时要恢复成原始的
                            End If
                        ElseIf blnReset Then
                            Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                            blnReset = False
                        End If
                    End If
                    
                    blnPrint = True
                    If Not PrintPage(i, Printer, Me) Then Exit For
                    If i <> lngEndPage Then Printer.NewPage: blnPrint = True '多页
                Next
            End If
            If k > 0 Then Printer.NewPage: blnPrint = True '多份
        Loop Until k = 0
        
NextFormat:
        '检查是否继续打印下一格式
        blnGoOn = False
        If InStr(";0;2;4;", ";" & mbytStyle & ";") > 0 And mblnAllFormat And cboFormat.ComboItems.count > 1 And cboFormat.SelectedItem.Index < cboFormat.ComboItems.count Then
            cboFormat.ComboItems(cboFormat.SelectedItem.Index + 1).Selected = True
            Call CboFormat_Click: blnGoOn = True
            If Not (mblnPrintEmpty = False And mobjReport.打印方式 = 1) Or blnExit = False Then
                Printer.NewPage
            End If
            blnPrint = True '多格式：如果新页无实际输出,则不会产生新打印页
        End If
    Loop
    cboFormat.Tag = ""

    If Not mobjCurDLL Is Nothing Then
        mobjCurDLL.DataIsEmpty = blnALLEmpty
    End If

ExitHandle:
    cboFormat.Tag = ""
    If blnPrint Then
        If gblnSingleTask Then
            Printer.NewPage '如果新页无实际输出,则不会产生新打印页
        Else
            Printer.EndDoc
        End If
        
        '输出PDF结束
        If mbytStyle = Val("4-PDF") Then
            Call PDFFileSuccess
        End If
        
        '输出到打印结束，激活打印事件
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_AfterPrint(mobjReport.编号)
        End If
    End If

    If mbytStyle <> 2 Then ShowFlash
    SetRedraw True
    Screen.MousePointer = 0
    Exit Sub
    
errH:
    cboFormat.Tag = ""
    Screen.MousePointer = 0
    If mbytStyle <> 2 Then Call ShowFlash
    Printer.KillDoc
    SetRedraw True
    MsgBox Err.Number & ":" & Err.Description & vbCrLf & "打印过程被强行中断！", vbExclamation, App.Title
    Err.Clear
    gblnError = True
End Sub

Private Sub mnuHelpTitle_Click()
    If Me.Tag = "" Then
        Call ShowHelpRpt(Me.hwnd, mobjReport.编号, Int((mobjReport.系统) / 100))
    Else
        Call ShowHelpRpt(Me.hwnd, Me.Tag, Int((mobjReport.系统) / 100))
    End If
End Sub

Private Sub mnuPop_Cond_Click(Index As Integer)
    Set mobjPars = mdlPublic.RPTParsCondExec(mlngRPTID, Val(mnuPop_Cond(Index).Tag), mobjDefPars)
    If Not mobjPars Is Nothing Then
        mintCurMenuIndex = Index
        mintCurCondID = Val(mnuPop_Cond(Index).Tag)
        Call InitReportPars
        If cmdLoad.Enabled And cmdLoad.Visible Then cmdLoad.SetFocus
    End If
End Sub

Private Sub mnuPop_Default_Click()
    '基于缺省参数对象，更新当前参数对象
    Call CopyPars(mobjDefPars, mobjPars)
    If Not mobjPars Is Nothing Then
        mintCurMenuIndex = 0
        mintCurCondID = 0
        Call InitReportPars
        If cmdLoad.Enabled And cmdLoad.Visible Then cmdLoad.SetFocus
    End If
End Sub

Private Sub mnuPop_Del_Click()
    If mdlPublic.RPTParsCondDel(mlngRPTID, mintCurCondID) Then
        Call mnuPop_Default_Click
    End If
End Sub

Private Sub mnuPop_Save_Click()
    '保存条件
    If mdlPublic.RPTParsCondSave(mlngRPTID, mintCurCondID, mobjPars, mobjDefPars, Me) Then
        '更新参数控件
        If mintCurCondID = 0 Then
            '从缺省状态下保存，更新为新增的条件
            Call mnuPop_Cond_Click(mnuPop_Cond.count - 1)
        Else
            '从条件状态下保存
            Call mnuPop_Cond_Click(mintCurCondID)
        End If
    End If
End Sub

Private Sub mnuPop_SaveAs_Click()
    If mdlPublic.RPTParsCondSave(mlngRPTID, mintCurCondID, mobjPars, mobjDefPars, Me, True) Then
        '更新参数控件
        If mintCurCondID = 0 Then
            '从缺省状态下保存，更新为新增的条件
            Call mnuPop_Cond_Click(mnuPop_Cond.count - 1)
        Else
            '从条件状态下保存
            Call mnuPop_Cond_Click(mintCurCondID)
        End If
    End If
End Sub

Private Sub mnuView_Next_Click()
    Dim intIdx As Integer
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    intIdx = lvw.SelectedItem.Index
    If intIdx + 1 <= lvw.ListItems.count Then
        lvw.ListItems(intIdx + 1).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuView_Pre_Click()
    Dim intIdx As Integer
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    intIdx = lvw.SelectedItem.Index
    If intIdx - 1 >= 1 Then
        lvw.ListItems(intIdx - 1).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuView_reFlash_Click()
    Dim strErr As String, strCond As String
    Dim tmpMsh As Object
    Dim strStartTime As String, strInfo As String, strName As String
    Dim intState As Integer
    
    If gblnReportRunLog Then
        strStartTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
    End If
    '取当前参数
    If lblName.UBound > 0 Then
        If Not ReSetReportPars Then Exit Sub
    End If
    
    '激活条件提交事件
    If Not mobjCurDLL Is Nothing Then
        '检查报表或报表组的状态
        strName = mdlPublic.FormatString("【[1]】[2]", mobjReport.编号, mobjReport.名称)
        intState = mdlPublic.ReportStateSwitch(mlngSys, mobjReport.编号, False, strInfo)
        Select Case intState
        Case Val("0-报表不存在或未发布")
            If strInfo = "" Then
                MsgBox mdlPublic.FormatString("报表[1]不存在，请联系管理员！", strName), vbInformation, App.Title
                Exit Sub
'            Else
'                MsgBox mdlPublic.FormatString("“[1]”报表未发布，请联系管理员！", strName), vbInformation, App.Title
            End If
        Case Val("1-报表启用中")
            '正常
        Case Val("2-报表停用中")
            MsgBox mdlPublic.FormatString("“[1]”报表停用中，请联系管理员！", strName), vbInformation, App.Title
            Exit Sub
        Case Else
            Exit Sub
        End Select
        
        '抛出事件
        strCond = GetParsStr(MakeNamePars(mobjReport, True))
        mobjCurDLL.Act_CommitCondition mobjReport.编号, strCond, Me
    End If
    
     '查询是否允许在当前时间执行此报表
    If Format(mobjReport.禁止结束时间, "HH:mm:ss") <> "00:00:00" Or Format(mobjReport.禁止开始时间, "HH:mm:ss") <> "00:00:00" Then
        If CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss")) > CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")) Then
            If Between(CDate(Format(Currentdate, "HH:mm:ss")), CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")), CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss"))) Then
                MsgBox "当前报表在" & CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")) & "-" & CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss")) & "禁止执行，如有疑问请联系信息科。", vbInformation, App.Title
                Exit Sub
            End If
        Else
            If CDate(Format(Currentdate, "HH:mm:ss")) < CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")) Or CDate(Format(Currentdate, "HH:mm:ss")) > CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss")) Then
                MsgBox "当前报表在" & CDate(Format(mobjReport.禁止开始时间, "HH:mm:ss")) & "-第二天" & CDate(Format(mobjReport.禁止结束时间, "HH:mm:ss")) & "禁止执行，如有疑问请联系信息科。", vbInformation, App.Title
                Exit Sub
            End If
        End If
    End If
    
    '重新数据
    strErr = OpenReportData(True)
    If strErr <> "" Then
        MsgBox "在读取报表数据""" & strErr & """时遇到意外错误,报表不能产生！", vbInformation, App.Title
        Exit Sub
    End If
    '重显内容
    Call ShowItems
    
    If lblName.UBound > 0 Then
        '重新显示参数(注意：需要重新显示)
        picPar.Visible = False
        Set mobjPars = New RPTPars
        Set mobjPars = MakeNamePars(mobjReport)
        Call InitReportPars
        picPar.Visible = True
        
        '根据当前参数,替换其它报表中相同参数的值
        Call KeepParsSame
    End If
    
    '定位在第一个表格上
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And tmpMsh.Container Is picPaper(intReport) And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
            On Error Resume Next
            tmpMsh.SetFocus
            On Error GoTo 0
            Exit For
        End If
    Next
    If lvw.ListItems.count > 0 Then
        Call RecordsExecute(Val(Mid(lvw.SelectedItem.Key, 2)), strStartTime, 2)
    Else
        Call RecordsExecute(mlngReportID, strStartTime, 2)
    End If
End Sub

Private Sub KeepParsSame()
'功能：报表组中，根据当前报表有效的参数内容,保持其它报表相同参数的值相同
'传入：mobjPars=当前报表所使用到的参数
'说明：1.无类型的参数不处理,因为可能参数值书写方式不一样
    Dim objPar As RPTPar, tmpPar As RPTPar
    Dim objData As RPTData, i As Integer
    For i = 0 To UBound(arrReport)
        If i <> intReport Then
            For Each objData In arrReport(i).Datas
                For Each objPar In objData.Pars
                    For Each tmpPar In mobjPars
                        If tmpPar.名称 = objPar.名称 _
                            And tmpPar.类型 = objPar.类型 _
                            And objPar.类型 <> 3 Then
                            objPar.缺省值 = tmpPar.缺省值
                            objPar.Reserve = tmpPar.Reserve
                        End If
                    Next
                Next
            Next
        End If
    Next
End Sub

Private Sub mnuViewToolFormat_Click()
    mnuViewToolFormat.Checked = Not mnuViewToolFormat.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolGroup_Click()
    mnuViewToolGroup.Checked = Not mnuViewToolGroup.Checked
    picLR_S.Visible = Not picLR_S.Visible
    picGroup.Visible = Not picGroup.Visible
    Call Form_Resize
End Sub

Private Sub msh_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim objRelations As RPTRelations
    Dim i As Long
    Dim intIdx As Integer
    Dim colRelations As Collection
    
    If mobjReport.Items("_" & Index).类型 = 6 Then
        intIdx = Val(Mid(msh(Index).Tag, 3))
    
        '由于Sort被赋值，首行的链接报表信息会发生变化。因此，需要先缓存再重新设置
        For i = 0 To msh(intIdx).Cols - 1
            If TypeName(msh(intIdx).Cell(flexcpData, 0, i)) = "RPTRelations" Then
                '对象缓存
                If colRelations Is Nothing Then
                    Set colRelations = New Collection
                End If
                colRelations.Add msh(intIdx).Cell(flexcpData, 0, i), "_" & i
                '清除原记录Data信息
                msh(intIdx).Cell(flexcpData, 0, i) = Empty
            End If
        Next
    
        msh(intIdx).Col = Col
        msh(intIdx).Sort = Order
        
        '对象恢复
        If Not colRelations Is Nothing Then
            For i = 0 To msh(intIdx).Cols - 1
                Set objRelations = Nothing
                On Error Resume Next
                Set objRelations = colRelations("_" & i)
                On Error GoTo 0
                If Not objRelations Is Nothing Then
                    Set msh(intIdx).Cell(flexcpData, 0, i) = objRelations
                End If
            Next
        End If
    End If
End Sub

Private Sub msh_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, intBegin As Integer, intEnd As Integer
    Dim lngID As Long
    Dim objRelaID As RelatID, objItem As RPTItem, objBody As RPTItem
    Dim sngNew As Single, sngOld As Single
    
    If mobjReport.Items("_" & Index).类型 = Val("6-任意表（表头）") Then
        If msh(Index).Tag Like "H_*" Then
            '调用列对象(RPTItem)宽度
            lngID = Val(Mid(msh(Index).Tag, 3))
            Set objBody = mobjReport.Items("_" & lngID)
            If Not objBody Is Nothing Then
                intBegin = -1
                intEnd = -1
                For Each objRelaID In objBody.SubIDs
                    Set objItem = mobjReport.Items("_" & objRelaID.id)
                    If objItem.序号 = Col Then
                        sngOld = objItem.W
                        objItem.W = msh(Index).ColWidth(Col)
                        sngNew = msh(Index).ColWidth(Col)
                    End If
                    If objItem.自适应行高 Then
                        If intBegin < 0 Then intBegin = objItem.序号
                        intEnd = objItem.序号
                    End If
                Next
            End If
            '调整表体宽度
            msh(lngID).ColWidth(Col) = msh(Index).ColWidth(Col)
            '调整行高
            If intBegin >= 0 Then
                msh(lngID).AutoSize intBegin, intEnd
            End If
        End If
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_ColResize(mobjReport.编号, CInt(Col), sngNew, sngOld)
        End If
    End If
End Sub

Private Sub msh_Click(Index As Integer)
    Dim objRelations As RPTRelations
    Dim i As Long
    Dim lngRec As Long, strDataName As String
    Dim strLisName As String
    Dim tmpData As RPTData, tmpPar As RPTPar
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim strFilter As String, lngItemID As Long
    Dim lngMouseRow As Long, lngMouseCol As Long
    Dim lngRelationReport As Long
    
    If mblnLeftClick = False And mlngRelationReport = 0 Then Exit Sub
    If mblnLeftClick = False And mlngRelationReport <> 0 Then
        lngMouseRow = mlngRelationMouseRow
        lngMouseCol = mlngRelationMouseCol
    Else
        lngMouseRow = msh(Index).MouseRow: lngMouseCol = msh(Index).MouseCol
    End If

    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Sub
    If grsObject.State = adStateClosed Then
        Set grsObject = Nothing
        Set grsObject = UserObject
        If grsObject Is Nothing Then Exit Sub
    End If
    
    If lngMouseRow > -1 And lngMouseCol > -1 Then
        If msh(Index).Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
            If mobjReport.Items("_" & Index).类型 = 4 Then
'                Set objRelations = msh(Index).Cell(flexcpData, lngMouseRow, lngMouseCol)(2)
'                '任意表可定位到具体的行上
'                lngRec = msh(Index).Cell(flexcpData, lngMouseRow, lngMouseCol)(1)
                
                '优化
                Set objRelations = msh(Index).Cell(flexcpData, 0, lngMouseCol)  '从第一行的单元格中获取Relations对象
                lngRec = msh(Index).RowData(lngMouseRow)                        '从行的RowData中获取记录集行号
            Else
                '优化；只在特定行绑定链报表对象
                lngRec = msh(Index).FixedRows
                If TypeName(msh(Index).Cell(flexcpData, lngRec, lngMouseCol)) = "Empty" Then Exit Sub
                If msh(Index).Cell(flexcpData, lngRec, lngMouseCol).Relations.count <= 0 Then Exit Sub
                
                Set objRelations = msh(Index).Cell(flexcpData, lngRec, lngMouseCol).Relations
                lngItemID = msh(Index).Cell(flexcpData, lngRec, lngMouseCol).id
            End If
            
            For i = 1 To objRelations.count
                If objRelations.Item(i).默认 = 1 Then
                    lngRelationReport = objRelations.Item(i).关联报表ID
                    Exit For
                End If
            Next
            If lngRelationReport = 0 Then lngRelationReport = objRelations.Item(1).关联报表ID
            If mlngRelationReport <> 0 Then lngRelationReport = mlngRelationReport
            If Not CheckReportPriv(lngRelationReport) Then
                MsgBox "你没有权限查询该报表某些数据源中的对象！", vbInformation, App.Title: Exit Sub
            End If
            '执行报表
            If CheckPass(lngRelationReport) = False Then
                MsgBox "报表数据错误，不能执行该报表！", vbInformation, App.Title: Exit Sub
            End If
            
            Set gobjReport = ReadReport(lngRelationReport)
            '初始化参数
            garrPars = Array()
            '定位记录集
            On Error Resume Next
            For i = 1 To objRelations.count
                If objRelations.Item(i).关联报表ID = lngRelationReport Then
                    If InStr(objRelations.Item(i).参数值来源, ".") > 0 Then
                        strDataName = Mid(objRelations.Item(i).参数值来源, 1, InStr(objRelations.Item(i).参数值来源, ".") - 1)
                    End If
                End If
                If strDataName <> "" Then Exit For
            Next
            If mobjReport.Items("_" & Index).类型 = 4 Then
                '任意表能够确定点击的具体数据行
                If strDataName <> "" Then mLibDatas("_" & strDataName).DataSet.AbsolutePosition = lngRec
            Else
                '汇总表只能根据纵向和横向分类定位数据行
                If strDataName <> "" Then
                    For Each tmpID In mobjReport.Items("_" & Index).SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.id)
                        Select Case mobjReport.Items("_" & lngItemID).类型
                            Case 7 '纵向分类
                                If tmpItem.类型 = 7 Then
                                    If Decode(Trim(msh(Index).TextMatrix(lngMouseRow, tmpItem.序号)), "合计", 1, "平均值", 2, "最大值", 3, "最小值", 4, "记录数", 5, 0) > 0 Then
                                        '如果是合计行则取上面一行
                                        lngMouseRow = lngMouseRow - 1
                                    End If
                                    strFilter = strFilter & " And " & tmpItem.内容 & "='" & msh(Index).TextMatrix(lngMouseRow, tmpItem.序号) & "'"
                                End If
                            Case 8 '横向分类
                                If tmpItem.类型 = 8 Then
                                    strFilter = strFilter & " And " & tmpItem.内容 & "='" & msh(Index).TextMatrix(tmpItem.序号, lngMouseCol) & "'"
                                End If
                            Case 9 '统计项
                                '统计项根据横向和纵向分类进行确定
                                If tmpItem.类型 = 7 Then
                                    If Decode(Trim(msh(Index).TextMatrix(lngMouseRow, tmpItem.序号)), "合计", 1, "平均值", 2, "最大值", 3, "最小值", 4, "记录数", 5, 0) > 0 Then
                                        '如果是合计行则取上面一行
                                        lngMouseRow = lngMouseRow - 1
                                    End If
                                    strFilter = strFilter & " And " & tmpItem.内容 & "='" & msh(Index).TextMatrix(lngMouseRow, tmpItem.序号) & "'"
                                ElseIf tmpItem.类型 = 8 Then
                                    strFilter = strFilter & " And " & tmpItem.内容 & "='" & msh(Index).TextMatrix(tmpItem.序号, lngMouseCol) & "'"
                                End If
                        End Select
                    Next
                    mLibDatas("_" & strDataName).DataSet.Filter = Mid(strFilter, 6)
                End If
            End If
            For i = 1 To objRelations.count
                With objRelations.Item(i)
                    strLisName = ""
                    If objRelations.Item(i).关联报表ID = lngRelationReport Then
                        If InStr(.参数值来源, ".") > 0 Then
                            If mLibDatas("_" & strDataName).DataSet.RecordCount > 0 Then
                                strLisName = mLibDatas("_" & strDataName).DataSet.Fields(Mid(.参数值来源, InStr(.参数值来源, ".") + 1)).Value
                            End If
                        ElseIf InStr(.参数值来源, "=") = 1 Then
                            For Each tmpData In mobjReport.Datas
                                For Each tmpPar In tmpData.Pars
                                    If tmpPar.名称 = Mid(.参数值来源, 2) Then
                                        strLisName = tmpPar.缺省值
                                        Exit For
                                    End If
                                Next
                                If strLisName <> "" Then Exit For
                            Next
                        End If
                    
                        ReDim Preserve garrPars(UBound(garrPars) + 1)
                        garrPars(UBound(garrPars)) = .参数名 & "=" & strLisName
                    End If
                End With
            Next
            
            If Not ShowReport(Me) Then MsgBox "报表打开失败！", vbInformation, App.Title
        End If
    End If
End Sub

Private Sub msh_DblClick(Index As Integer)
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetDblClick(mobjReport.编号, msh(Index), Me)
        msh(Index).SetFocus
    End If
End Sub

Private Sub msh_EnterCell(Index As Integer)
    Dim i As Long, lngRow As Long, lngCol As Long
    Dim strRowText As String, strText As String
    Dim intA As Integer, intB As Integer
    Static strRow As String
    Static strCol As String
    
    If blnRefresh = False Then Exit Sub
    Set objCurGrid = msh(Index)
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        lngRow = msh(Index).Row
        lngCol = msh(Index).Col
        intA = IIF(lngRow > 30000, 30000, lngRow)
        intB = IIF(lngCol > 30000, 30000, lngCol)
        strText = msh(Index).Text
        Call mobjCurDLL.Act_EnterCell(mobjReport.编号, intA, intB, strText)
        '改变后的值
        If lngRow >= 0 And lngRow <= msh(Index).Rows - 1 And lngCol >= 0 And lngCol <= msh(Index).Cols - 1 Then
            msh(Index).Row = lngRow
            msh(Index).Col = lngCol
            msh(Index).Text = strText
        End If
        
        If strRow <> Index & "," & msh(Index).Row Then
            For i = 0 To msh(Index).Cols - 1
                strRowText = strRowText & "|" & msh(Index).TextMatrix(msh(Index).Row, i)
            Next
            intA = IIF(msh(Index).Row > 30000, 30000, msh(Index).Row)
            Call mobjCurDLL.Act_EnterRow(mobjReport.编号, intA, Mid(strRowText, 2), msh(Index))
            strRow = Index & "," & msh(Index).Row
        End If
        
        If strCol <> Index & "," & msh(Index).Col Then
            Call mobjCurDLL.Act_EnterCol(mobjReport.编号, msh(Index).Col, msh(Index))
            strCol = Index & "," & msh(Index).Col
        End If
    End If
End Sub

Private Sub msh_GotFocus(Index As Integer)
    On Error Resume Next
    If msh(Index).Tag Like "H_*" Then
        msh(CInt(Mid(msh(Index).Tag, 3))).SetFocus
    Else
        Call msh_EnterCell(Index)
    End If
End Sub

Private Sub msh_LeaveCell(Index As Integer)
    Dim intRow As Integer
    
    If blnRefresh = False Then Exit Sub
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        intRow = IIF(msh(Index).Row > 30000, 30000, msh(Index).Row)
        Call mobjCurDLL.Act_LevelCell(mobjReport.编号, intRow, msh(Index).Col, msh(Index).Text)
    End If
End Sub

Private Sub msh_LostFocus(Index As Integer)
    If Not msh(Index).Tag Like "H_*" Then Call msh_LeaveCell(Index)
End Sub

Private Sub msh_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseDown(mobjReport.编号, Button, Shift, X, Y, msh(Index), Me)
    End If
End Sub

Private Sub msh_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    msh(Index).ToolTipText = msh(Index).TextMatrix(msh(Index).MouseRow, msh(Index).MouseCol)
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseMove(mobjReport.编号, Button, Shift, X, Y, msh(Index), Me)
    End If
    If msh(Index).MouseRow > -1 And msh(Index).MouseCol > -1 Then
        If msh(Index).Cell(flexcpFontUnderline, msh(Index).MouseRow, msh(Index).MouseCol) = True Then
            msh(Index).MousePointer = 99
        Else
            msh(Index).MousePointer = 0
        End If
    Else
        msh(Index).MousePointer = 0
    End If
End Sub

Private Sub msh_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim dbRowHeight As Double
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseUp(mobjReport.编号, Button, Shift, X, Y, msh(Index), Me)
    End If
    
    If Button = vbRightButton Then
        If msh(Index).MouseRow < 0 Or msh(Index).MouseCol < 0 Then
            vsfRelations.Visible = False
            Exit Sub
        End If
        mintGridIndex = Index
        mblnLeftClick = False
        mlngRelationMouseRow = msh(Index).MouseRow
        mlngRelationMouseCol = msh(Index).MouseCol
        If msh(Index).Cell(flexcpFontUnderline, msh(Index).MouseRow, msh(Index).MouseCol) Then
            Call LoadRelation(0, Index, msh(Index).MouseRow, msh(Index).MouseCol)
            If TypeName(msh(Index).Cell(flexcpData, msh(Index).MouseRow, msh(Index).MouseCol)) = "RPTRelations" Then
                vsfRelations.Visible = True
                vsfRelations.SetFocus
            End If
            
            For i = 0 To vsfRelations.Rows - 1
                dbRowHeight = dbRowHeight + vsfRelations.RowHeight(i)
            Next
            vsfRelations.Height = dbRowHeight
            vsfRelations.Left = msh(Index).Left + X + 150
            vsfRelations.Top = msh(Index).Top + Y + 90
        Else
            vsfRelations.Visible = False
        End If
    Else
        mblnLeftClick = True
    End If
End Sub

Private Sub LoadRelation(ByVal bytType As Byte, ByVal cIndex As Integer, Optional lngMouseRow As Long, Optional lngMouseCol As Long)
    Dim i As Long
    Dim objRelations As RPTRelations
    Dim strFlag As String
    
    mbytType = bytType
    If mbytType = 0 Then
        If TypeName(msh(cIndex).Cell(flexcpData, lngMouseRow, lngMouseCol)) <> "RPTRelations" Then
            Exit Sub
        End If
        If mobjReport.Items("_" & cIndex).类型 = 4 Then
            Set objRelations = msh(cIndex).Cell(flexcpData, lngMouseRow, lngMouseCol)(2)
        Else
            Set objRelations = msh(cIndex).Cell(flexcpData, lngMouseRow, lngMouseCol).Relations
        End If
    ElseIf mbytType = 1 Then
        Set objRelations = mobjReport.Items("_" & cIndex).Relations
    End If
    If objRelations.count = 0 Then Exit Sub

    With vsfRelations
        .Rows = 0
        If .Cols = 0 Then
            .Cols = 2
            .ColKey(0) = "ID"
            .ColDataType(0) = flexDTString
            .ColWidth(0) = 0
            .ColKey(1) = "名称"
            .ColDataType(1) = flexDTString
            .ColWidth(1) = vsfRelations.Width
        End If
        For i = 1 To objRelations.count
            If InStr(strFlag, "," & objRelations.Item(i).关联报表ID) = 0 Then
                strFlag = strFlag & "," & objRelations.Item(i).关联报表ID
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = objRelations.Item(i).关联报表ID
                .TextMatrix(.Rows - 1, .ColIndex("名称")) = " " & Split(objRelations.Item(i).关联报表名称, "(")(0)
            End If
        Next
    End With
End Sub

Private Sub msh_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim intPre As Integer
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetScroll(mobjReport.编号, msh(Index))
    End If
    
    If IsNumeric(msh(Index).Tag) Then
        intPre = msh(msh(Index).Tag).LeftCol
        msh(msh(Index).Tag).LeftCol = msh(Index).LeftCol
        If msh(msh(Index).Tag).LeftCol = intPre Then msh(Index).LeftCol = intPre
    ElseIf Left(msh(Index).Tag, 2) = "H_" Then
        intPre = msh(Mid(msh(Index).Tag, 3)).LeftCol
        msh(Mid(msh(Index).Tag, 3)).LeftCol = msh(Index).LeftCol
        If msh(Mid(msh(Index).Tag, 3)).LeftCol = intPre Then msh(Index).LeftCol = intPre
    End If
End Sub

Private Sub opt_GotFocus(Index As Integer)
    If opt(Index).Value Then
        '这样做的目的是避免按TAB键时自动切换到下一个选项
        opt(Index).Value = False
        opt(Index).Value = True
    End If
End Sub

Private Sub picLR_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objTmp As Object
    
    On Error Resume Next
    
    If Button = 1 Then
        If picGroup.Width + X < 1000 Or picBack.Width - X < 3000 Then Exit Sub
        picLR_S.Left = picLR_S.Left + X

        picGroup.Width = picGroup.Width + X
        picBack.Left = picBack.Left + X
        picBack.Width = picBack.Width - X
        scrHsc.Left = scrHsc.Left + X
        scrHsc.Width = scrHsc.Width - X
        
        lblGroup_S.Width = lblGroup_S.Width + X
        lvw.Width = lvw.Width + X
        lblPar_S.Width = lblPar_S.Width + X
        picPar.Width = picPar.Width + X
        
        lvw.ColumnHeaders(1).Width = lvw.Width - 500    '动态列宽
        
        For Each objTmp In fraGroup
            objTmp.Width = picGroup.ScaleWidth - objTmp.Left * 2
        Next
        For Each objTmp In fra
            objTmp.Width = picGroup.ScaleWidth - objTmp.Left * 2
        Next
        
        picPaper(intReport).Cls
        Call SetPaper
        Call SetPlace
        Me.Refresh
    End If
End Sub

Private Sub picPane_Click()

End Sub

Private Sub picPaper_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnPop As Boolean
    
    lngPreX = X: lngPreY = Y
    
    If Not mobjCurDLL Is Nothing Then
        blnPop = True
        Call mobjCurDLL.Act_PaperMouseDown(mobjReport.编号, Button, Shift, X, Y, blnPop)
        If blnPop Then
            If Button = 2 Then PopupMenu mnuEdit, 2
        End If
    Else
        If Button = 2 Then PopupMenu mnuEdit, 2
    End If
End Sub

Private Sub picPaper_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_PaperMouseMove(mobjReport.编号, Button, Shift, X, Y)
    End If
    If Button = 1 Then
        If scrVsc.Enabled And scrVsc.Visible Then
            If (Y - lngPreY) / 15 > 0 Then
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - lngPreY) / 15)
            Else
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - lngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled And scrHsc.Visible Then
            If (X - lngPreX) / 15 > 0 Then
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - lngPreX) / 15)
            Else
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - lngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picPaper_GotFocus(Index As Integer)
    Oldwinproc = GetWindowLong(picPaper(Index).hwnd, GWL_WNDPROC)
    SetWindowLong picPaper(Index).hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub picPaper_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        '下
        If scrVsc.Value + scrVsc.Max / 10 > scrVsc.Max Then
            scrVsc.Value = scrVsc.Max
        Else
            scrVsc.Value = scrVsc.Value + scrVsc.Max / 10
        End If
    ElseIf KeyCode = vbKeyPageUp Then
        '上
        If scrVsc.Value - scrVsc.Max / 10 < 0 Then
            scrVsc.Value = 0
        Else
            scrVsc.Value = scrVsc.Value - scrVsc.Max / 10
        End If
    End If
End Sub

Private Sub picPaper_LostFocus(Index As Integer)
    SetWindowLong picPaper(Index).hwnd, GWL_WNDPROC, Oldwinproc
End Sub

Private Sub mnuEdit_Par_Click()
'功能：重置报表条件
    Dim strErr As String, objPars As RPTPars
    Dim strCond As String, blnInhere As Boolean
    Dim lngReport As Long, strSQL As String
    Dim rsReport As New ADODB.Recordset
    Dim frmNewParInput As New frmParInput
    Dim strStartTime As String
    
    '取该报表的ID
    lngReport = 0
    strSQL = "Select ID from zlReports Where 编号=[1]"
    Set rsReport = OpenSQLRecord(strSQL, Me.Caption, mobjReport.编号)
    If Not rsReport.EOF Then lngReport = rsReport!id
    
    If gblnReportRunLog Then
        strStartTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
    End If
    
    If Not mobjCurDLL Is Nothing Then
        blnInhere = True
        Set objPars = MakeNamePars(mobjReport, True)
        strCond = GetParsStr(objPars)
        
        '激活重置条件事件
        mobjCurDLL.Act_ResetCondition mobjReport.编号, strCond, blnInhere, Me
        
        If Not blnInhere Then
             '调用程序取消条件重置,或条件格式设置错误
            If strCond = "" Or Not strCond Like "*=*" Then Exit Sub
            
            Set objPars = SetStrPars(strCond, objPars)
            
            '激活条件提交事件
            strCond = GetParsStr(objPars)
            mobjCurDLL.Act_CommitCondition mobjReport.编号, strCond, Me
            
            ReplaceInputPars objPars
            
            Me.Refresh
            strErr = OpenReportData(True)
            If strErr <> "" Then MsgBox "在读取报表数据""" & strErr & """时遇到意外错误,报表不能产生！", vbInformation, App.Title: Exit Sub
            Call ShowItems
        Else
            Set objPars = MakeNamePars(mobjReport) '需要重新以这种方式取
            
            frmNewParInput.mlngReport = lngReport
            Set frmNewParInput.mobjPars = objPars
            Set frmNewParInput.mobjDefPars = mobjDefPars
            Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
            frmNewParInput.mstrTitle = mobjReport.名称
            frmNewParInput.mblnReset = True
            frmNewParInput.Show 1, Me
            If frmNewParInput.mblnOK Then
                '激活条件提交事件
                strCond = GetParsStr(frmNewParInput.mobjPars)
                mobjCurDLL.Act_CommitCondition mobjReport.编号, strCond, Me
                
                ReplaceInputPars frmNewParInput.mobjPars
                Unload frmNewParInput
                
                '产生数据
                Me.Refresh
                strErr = OpenReportData(True)
                If strErr <> "" Then MsgBox "在读取报表数据""" & strErr & """时遇到意外错误,报表不能产生！", vbInformation, App.Title: Exit Sub
                Call ShowItems
            End If
        End If
    Else
        frmNewParInput.mlngReport = lngReport
        Set frmNewParInput.mobjPars = MakeNamePars(mobjReport)
        Set frmNewParInput.mobjDefPars = mobjDefPars
        Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
        frmNewParInput.mstrTitle = mobjReport.名称
        frmNewParInput.mblnReset = True
        frmNewParInput.Show 1, Me
        If frmNewParInput.mblnOK Then
            ReplaceInputPars frmNewParInput.mobjPars
            Unload frmNewParInput
           
            '产生数据
            Me.Refresh
            strErr = OpenReportData(True)
            If strErr <> "" Then MsgBox "在读取报表数据""" & strErr & """时遇到意外错误,报表不能产生！", vbInformation, App.Title: Exit Sub
            Call ShowItems
        End If
    End If
    Call RecordsExecute(lngReport, strStartTime, 2)
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    Dim lngTmp As Long
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    'lblPar_S、lblGropu_S、picLR_S 的命名方式是为了窗体记忆功能的处理
    
    '靠齐控件宽度和高度
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(sta.Visible, sta.Height, 0)
    
    lblGroup_S.Width = picGroup.ScaleWidth - lblGroup_S.Left * 2
    
    lblPar_S.Width = lblGroup_S.Width
    
    lvw.Top = lblGroup_S.Top + lblGroup_S.Height + 15
    lvw.Width = picGroup.ScaleWidth
    lvw.Height = lblPar_S.Top - lblGroup_S.Top - lblGroup_S.Height - 15 * 2
    
    picPar.Top = lblPar_S.Top + lblPar_S.Height + 15
    picPar.Left = 0
    picPar.Width = lvw.Width
    picPar.Height = ScaleHeight - staH - cbrH - (lblGroup_S.Height + 30) - (lblPar_S.Height + 30) - lvw.Height
    
    picBack.Top = ScaleTop + cbrH
    picBack.Left = ScaleLeft + IIF(picGroup.Visible, picGroup.Width + picLR_S.Width, 0)
    picBack.Width = ScaleWidth - IIF(scrVsc.Visible, scrVsc.Width, 0) - IIF(picGroup.Visible, picGroup.Width + picLR_S.Width, 0)
    picBack.Height = ScaleHeight - staH - cbrH - IIF(scrHsc.Visible, scrHsc.Height, 0)
    
    If scrVsc.Visible Then
        scrVsc.Top = picBack.Top
        scrVsc.Left = ScaleWidth - scrVsc.Width
        scrVsc.Height = picBack.Height
        
        scrHsc.Left = picBack.Left
        scrHsc.Top = picBack.Top + picBack.Height
        scrHsc.Width = picBack.Width
    End If
    
    On Error GoTo 0
    
    If Not mobjReport Is Nothing And Visible Then
        picPaper(intReport).Cls
        Call SetPaper
        Call SetPlace
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, bytMode As Byte
    
    bytMode = IIF(mnuEdit_SelMode_Row.Checked, 0, 1)
    
    If lvw.ListItems.count > 0 Then
        SaveWinState Me, App.ProductName, Me.Tag
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & Me.Tag, "选择模式", bytMode
        For i = 0 To UBound(arrReport)
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & arrReport(i).编号, "格式", arrReport(i).bytFormat
        Next
    ElseIf Not mobjReport Is Nothing Then
        If mbytStyle = 0 Then '不显示窗体时就不处理以加快速度
            SaveWinState Me, App.ProductName, mobjReport.编号
        End If
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.编号, "选择模式", bytMode
    End If
    
    If Not mobjReport Is Nothing Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.编号, "格式", bytFormat
    End If
    
    '激活窗体卸载事件
    If Not mobjCurDLL Is Nothing And Not mobjReport Is Nothing Then
        Call mobjCurDLL.Act_ReportUnload(mobjReport.编号, Me)
    End If
    
    '释放模块变量
    '---------------------------------------------------
    mbytStyle = 0
    mstrExcelFile = ""
    mstrPDFFile = ""
    
    Unload frmFlash
    
    Set frmParent = Nothing
    Set mobjCurDLL = Nothing
    Set mobjReport = Nothing
    Set mLibDatas = Nothing
    Set objCurGrid = Nothing
    Set mobjPars = Nothing
    Set mobjDefPars = Nothing
    Set objScript = Nothing
    
    Erase arrReport, arrLibDatas, arrDefPars

    If IsArray(marrPars) Then Erase marrPars
    If IsArray(marrPage) Then Erase marrPage
    marrPars = Empty
    marrPage = Empty

    Err.Clear
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuFile_Setup_Click()
    Dim objFmt As RPTFmt
    Dim strTmp As String
    Dim strDefault As String
    
    Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)
    strTmp = GetRegPrinterInfo("Printer", mobjReport.编号, objFmt.说明, mobjReport)
    If Not ReportLocalSet(mobjReport.系统, mobjReport.编号, False, mobjReport.bytFormat, Me) Then Exit Sub
    sta.Panels(2) = "打印机:" & strTmp & _
        "   纸张:" & GetPaperName(objFmt.纸张, objFmt.W, objFmt.H) & " " & _
        IIF(objFmt.纸张 = 256, CInt(objFmt.W / Twip_mm) & "mm × " & CInt(objFmt.H / Twip_mm) & "mm", "") & _
        IIF(objFmt.纸向 = 1, "   纵向", "   横向")
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta.Visible = Not sta.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.count
        tbr.Buttons(i).Caption = IIF(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub picPaper_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_PaperMouseUp(mobjReport.编号, Button, Shift, X, Y)
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Par"
            If lvw.ListItems.count = 0 Then
                mnuEdit_Par_Click '单个报表时重置条件
            Else
                mnuView_reFlash_Click
            End If
        Case "Preview"
            mnuFile_Preview_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Graph"
            mnuFile_Graph_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Style"
            Call SetView((lvw.View + 1) Mod 4)
        Case "Pre"
            mnuView_Pre_Click
        Case "Next"
            mnuView_Next_Click
        Case "ColWidth"
            mnuEdit_SetCol_Auto_Click
        Case "SelMode"
            If mnuEdit_SelMode_Row.Checked Then
                mnuEdit_SelMode_Col_Click
            Else
                mnuEdit_SelMode_Row_Click
            End If
    End Select
End Sub

Private Sub SetPaper()
'功能：设置报表纸张尺寸,位置,滚动
'说明：不依赖于打印设备
    Dim strPrinter As String
    Dim strDefault As String
    
    strDefault = mobjReport.Fmts(mobjReport.bytFormat).说明
    strPrinter = GetRegPrinterInfo("Printer", mobjReport.编号, strDefault, mobjReport)
    With mobjReport.Fmts("_" & mobjReport.bytFormat)
        sta.Panels(2).Text = "打印机:" & strPrinter & "   纸张:" & GetPaperName(.纸张, .W, .H) & " " & _
            IIF(.纸张 = 256, CInt(.W / Twip_mm) & "mm × " & CInt(.H / Twip_mm) & "mm", "") & _
            IIF(.纸向 = 1, "   纵向", "   横向")
    End With
    On Error GoTo errH
    
    If intGridCount = 1 And Not mobjReport.票据 Then
        picPaper(intReport).Top = 45
        picPaper(intReport).Left = 45
        picPaper(intReport).Width = picBack.ScaleWidth - picPaper(intReport).Left * 2
        picPaper(intReport).Height = picBack.ScaleHeight - picPaper(intReport).Top * 2
    Else
        With mobjReport.Fmts("_" & mobjReport.bytFormat)
            If .纸向 = 1 Then
                picPaper(intReport).Width = .W
                picPaper(intReport).Height = .H
            Else
                picPaper(intReport).Width = .H
                picPaper(intReport).Height = .W
            End If
        End With
        picShadow.Width = picPaper(intReport).Width
        picShadow.Height = picPaper(intReport).Height
        
        If picBack.ScaleWidth >= picPaper(intReport).Width + 180 Then
            picPaper(intReport).Left = (picBack.ScaleWidth - (picPaper(intReport).Width + 180)) / 2 + 60
            scrHsc.Enabled = False
        Else
            picPaper(intReport).Left = 60
            scrHsc.Max = (picPaper(intReport).Width + 180 - picBack.ScaleWidth) / 15
            If scrHsc.Max / 3 < scrHsc.SmallChange Then
                scrHsc.LargeChange = scrHsc.SmallChange
            Else
                scrHsc.LargeChange = scrHsc.Max / 3
            End If
            scrHsc.Enabled = True
        End If
        
        If picBack.ScaleHeight >= picPaper(intReport).Height + 180 Then
            picPaper(intReport).Top = (picBack.ScaleHeight - (picPaper(intReport).Height + 180)) / 2 + 60
            scrVsc.Enabled = False
        Else
            picPaper(intReport).Top = 60
            scrVsc.Max = (picPaper(intReport).Height + 180 - picBack.ScaleHeight) / 15
            If scrVsc.Max / 3 < scrVsc.SmallChange Then
                scrVsc.LargeChange = scrVsc.SmallChange
            Else
                scrVsc.LargeChange = scrVsc.Max / 3
            End If
            scrVsc.Enabled = True
        End If
        
        picShadow.Top = picPaper(intReport).Top + 60
        picShadow.Left = picPaper(intReport).Left + 60
    End If
    Exit Sub
errH:
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub scrhsc_Change()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrHsc.Value / (scrHsc.Max - scrHsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.编号, 0, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrHsc.Value = (scrHsc.Max - scrHsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Left = -scrHsc.Value * 15# + 60
    picShadow.Left = picPaper(intReport).Left + 60
    Me.Refresh
End Sub

Private Sub scrhsc_Scroll()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrHsc.Value / (scrHsc.Max - scrHsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.编号, 0, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrHsc.Value = (scrHsc.Max - scrHsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Left = -scrHsc.Value * 15# + 60
    picShadow.Left = picPaper(intReport).Left + 60
    Me.Refresh
End Sub

Private Sub scrVsc_Change()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrVsc.Value / (scrVsc.Max - scrVsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.编号, 1, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrVsc.Value = (scrVsc.Max - scrVsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Top = -scrVsc.Value * 15# + 60
    picShadow.Top = picPaper(intReport).Top + 60
    Me.Refresh
End Sub

Private Sub scrVsc_Scroll()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrVsc.Value / (scrVsc.Max - scrVsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.编号, 1, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrVsc.Value = (scrVsc.Max - scrVsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Top = -scrVsc.Value * 15# + 60
    picShadow.Top = picPaper(intReport).Top + 60
    Me.Refresh
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Auto"
            mnuEdit_SetCol_Auto_Click
        Case "Def"
            mnuEdit_SetCol_Def_Click
        Case "Fill"
            mnuEdit_SetCol_Fill_Click
        Case "Large"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
        Case "RowMode"
            mnuEdit_SelMode_Row_Click
        Case "ColMode"
            mnuEdit_SelMode_Col_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Function GetGridSource(objItem As RPTItem, Optional ByVal blnHead As Boolean) As String
'功能：从任意表格中搜索出所用倒的数据源名
'参数：objItem=任意表格主项
'      blnHead=是否从表头标签中检查
'返回："病人信息,药品信息,...",""
    Dim tmpID As RelatID
    Dim strSource As String, strFormula As String
    
    For Each tmpID In objItem.SubIDs
        strFormula = mobjReport.Items("_" & tmpID.id).内容
        Do While InStr(strFormula, "[") > 0
            strSource = Trim(Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1))
            strFormula = Mid(strFormula, InStr(strFormula, "]") + 1)
            If InStr(strSource, ".") > 0 Then
                If InStr(GetGridSource & ",", "," & Left(strSource, InStr(strSource, ".") - 1) & ",") = 0 Then
                    GetGridSource = GetGridSource & "," & Left(strSource, InStr(strSource, ".") - 1)
                End If
            End If
        Loop
        
        If blnHead Then
            strFormula = mobjReport.Items("_" & tmpID.id).表头
            Do While InStr(strFormula, "[") > 0
                strSource = Trim(Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1))
                strFormula = Mid(strFormula, InStr(strFormula, "]") + 1)
                If InStr(strSource, ".") > 0 Then
                    If InStr(GetGridSource & ",", "," & Left(strSource, InStr(strSource, ".") - 1) & ",") = 0 Then
                        GetGridSource = GetGridSource & "," & Left(strSource, InStr(strSource, ".") - 1)
                    End If
                End If
            Loop
        End If
    Next
    If GetGridSource <> "" Then GetGridSource = Mid(GetGridSource, 2)
End Function

Private Function ReplaceUserPars(objReport As Report) As Boolean
'功能：根据使用者传入参数设置报表参数值
'返回：使用者是否传入(全部)(正确)参数
'说明：为了避免"="号与无类型参数冲突,分析使用者参数格式时不用Split函数,而用Instr函数
    Dim tmpData As RPTData, tmpPar As RPTPar
    Dim i As Integer, j As Integer, k As Integer
    Dim blnCur As Boolean, blnALL As Boolean
    Dim strTmp As String
    
    If Not IsArray(marrPars) Then Exit Function
    If UBound(marrPars) <> -1 Then
        '先判断是否全部传完
        blnALL = True
        For Each tmpData In objReport.Datas
            For Each tmpPar In tmpData.Pars
                blnCur = False: k = k + 1
                For i = 0 To UBound(marrPars)
                    '参数名称相同且格式合法才替换
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase(tmpPar.名称) Then
                            strTmp = Trim(Mid(CStr(marrPars(i)), j + 1))
                            If strTmp <> "" Then
                                If InStr(strTmp, "|") > 0 And (tmpPar.缺省值 = "选择器定义…" Or tmpPar.缺省值 = "固定值列表…") Then
                                    blnCur = True: Exit For
                                Else
                                    Select Case tmpPar.类型
                                        Case 0, 3
                                            blnCur = True: Exit For
                                        Case 1
                                            If IsNumeric(strTmp) Then blnCur = True: Exit For
                                        Case 2
                                            If IsDate(strTmp) Then blnCur = True: Exit For
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next
                blnALL = blnALL And blnCur
            Next
        Next
        
        '再处理
        For Each tmpData In objReport.Datas
            For Each tmpPar In tmpData.Pars
                k = k + 1
                For i = 0 To UBound(marrPars)
                    '参数名称相同且格式合法才替换
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase(tmpPar.名称) Then
                            strTmp = Trim(Mid(CStr(marrPars(i)), j + 1))
                            If strTmp <> "" Then
                                '当由程序只传入了绑定值时,模拟为传入了显示值
                                If InStr(strTmp, "|") = 0 And (tmpPar.缺省值 = "选择器定义…" Or tmpPar.缺省值 = "固定值列表…") Then
                                    If tmpPar.缺省值 = "固定值列表…" Then
                                        For j = 0 To UBound(Split(tmpPar.值列表, "|"))
                                            If Split(Split(tmpPar.值列表, "|")(j), ",")(1) = strTmp Then
                                                strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0) & "|" & strTmp
                                                If Left(strTmp, 1) = "√" Then strTmp = Mid(strTmp, 2)
                                                Exit For
                                            End If
                                        Next
                                    Else
                                        '如果不弹出参数窗体则无用,否则在参数窗体中再作处理
                                        strTmp = "程序传入|" & strTmp
                                    End If
                                End If
                                If InStr(strTmp, "|") > 0 And (tmpPar.缺省值 = "选择器定义…" Or tmpPar.缺省值 = "固定值列表…") Then
                                    '带显示值,绑定值定义的参数。
                                    If Not blnALL Then
                                        '在未传完参数的情况下,原理是将参数对象模拟为重置时的值
                                        tmpPar.Reserve = strTmp
                                    Else
                                        '在传完的情况下,原理是将参数对象模拟为将要执行时的值
                                        tmpPar.Reserve = tmpPar.缺省值 & "|" & Split(strTmp, "|")(0)
                                        tmpPar.缺省值 = Split(strTmp, "|")(1)
                                    End If
                                    Exit For '当前参数已替换,不用再找
                                Else
                                    '一般传入参数,一般要传完,这种情况不能选择
                                    '不管是否传完,都直接处理缺省值
                                    If tmpPar.Reserve = "" And Left(tmpPar.缺省值, 1) = "&" Then
                                        tmpPar.Reserve = tmpPar.缺省值
                                    End If
                                    Select Case tmpPar.类型
                                        Case 0, 3
                                            tmpPar.缺省值 = strTmp: Exit For
                                        Case 1
                                            If IsNumeric(strTmp) Then tmpPar.缺省值 = strTmp: Exit For
                                        Case 2
                                            If IsDate(strTmp) Then tmpPar.缺省值 = strTmp: Exit For
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        Next
    End If
    ReplaceUserPars = blnALL
End Function

Private Function ParCount(objReport As Report) As Integer
'功能：从报表对象中返回不重复名称参数个数
    Dim tmpPar As RPTPar, tmpData As RPTData, StrPar As String
    
    If objReport.Datas.count = 0 Then ParCount = 0: Exit Function
    For Each tmpData In objReport.Datas
        For Each tmpPar In tmpData.Pars
            If InStr(StrPar & ",", "," & tmpPar.名称 & ",") = 0 Then
                StrPar = StrPar & "," & tmpPar.名称
                ParCount = ParCount + 1
            End If
        Next
    Next
End Function

Private Sub ReplaceInputPars(objPars As RPTPars)
'功能：根据参数输入窗体输入的参数值(名称唯一)替换报表数据源的参数集
    Dim tmpData As RPTData, tmpPar As RPTPar, objPar As RPTPar
    
    For Each tmpData In mobjReport.Datas
        For Each tmpPar In tmpData.Pars
            '对当前参数进行替换
            For Each objPar In objPars
                If objPar.名称 = tmpPar.名称 Then
                    tmpPar.缺省值 = objPar.缺省值
                    tmpPar.Reserve = objPar.Reserve
                    Exit For '下一个报表参数
                End If
            Next
        Next
    Next
End Sub

Private Function OpenReportData(Optional ByVal blnAllReLoad As Boolean = True) As String
'功能：根据报表对象(mobjReport)当前格式的数据源内容,产生具体的报表数据集
'功能：blnAllReLoad=是否是全部重新读取数据源(比如刷新,重置条件时,而切换格式时只重读需要的)
'返回：成功="",失败="数据源名"
    Dim tmpData As RPTData, strName As String
    Dim rsTmp As ADODB.Recordset
    Dim blnDo As Boolean, i As Integer
    
    '没有定义数据源
    mobjReport.blnLoad = True '表示报表本次读取是否正确
    If mobjReport.Datas.count = 0 Then Exit Function
    
    If blnAllReLoad Then
        Set mLibDatas = Nothing
        Set mLibDatas = New LibDatas
    ElseIf mLibDatas Is Nothing Then
        Set mLibDatas = New LibDatas
    End If
    
    On Error GoTo hErr
            
    For Each tmpData In mobjReport.Datas
        '判断该数据源是否已读取
        blnDo = True
        For i = 1 To mLibDatas.count
            If mLibDatas(i).Key = tmpData.名称 Then
                blnDo = False: Exit For
            End If
        Next
        '读取当前格式用到的数据源
        If blnDo And DataUsed(mobjReport, tmpData.名称, True) Then
            strName = tmpData.名称
            Set rsTmp = Nothing
            Set rsTmp = OpenReportSQL(tmpData)
            If rsTmp Is Nothing Then
                OpenReportData = tmpData.名称
                mobjReport.blnLoad = False
                Call ShowFlash: Exit Function
            End If
            mLibDatas.Add strName, rsTmp, "_" & strName
        End If
    Next
    
    Call ShowFlash
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Private Function OpenReportSQL(objData As RPTData) As ADODB.Recordset
'功能：根据数据源对象内容打开记录集
'说明：当静态ADO.Command变量返回的记录集被Clone出去，并且该Clone处于打开状态时，重复执行Command会出现对象已打开错误。
'1.执行报表组时,存在这种情况,Clone的记录集又不能关闭,因此不使用静态变量。
'2.单个报表不存在这种情况,可用Static变量.但该函数应放到公共模块中,不然窗体关闭了Static也没有效果。
'  因为单个报表重复执行频率不会很高,连接赋值的效率影响可以忽略。
'3.允许特殊写法变量绑定。如：select '[0]' 名称 from ...

    Dim rsTmp As New ADODB.Recordset
    Dim cmdData As New ADODB.Command
    Dim strLeft As String, strRight As String
    Dim StrPar As String, strParOld As String, bytPar As Byte
    Dim strSQL As String, strLog As String
    Dim intMax As Integer
    Dim strSQLtmp As String, i As Long, arrStr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim intDateType  As Integer  '0=无加减运算，1=加减常数，2=加减其他，如字段,3=其他运算，保持以前的规则，传入字符绑定变量
    Dim j As Long, k As Long, datValue As Date
    Dim l As Long

    If mbytStyle = 0 Or mbytStyle = 1 Then
        ShowFlash "正在读取数据""" & objData.名称 & """，请稍候．．．", , Me
    End If

    On Error GoTo errHandle

    '解析原始SQL
    'strSql = SQLOwner(TrimChar(objData.SQL), objData.对象)
    strSQL = SQLOwner(RemoveNote(objData.SQL), objData.对象)
    
    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLtmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLtmp, 7)), 1, 2) <> "/*" And Mid(strSQLtmp, 1, 6) = "SELECT" Then
        If Not Replace(strSQLtmp, " ", "") Like "*/[*]+CARDINALITY*[*]/*" Then      '/**/里可能出现多个CARDINALITY
            arrStr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
            For i = 0 To UBound(arrStr)
                strSQLtmp1 = strSQLtmp
                Do While InStr(strSQLtmp1, arrStr(i)) > 0
                    '判断前面是否用了IN 用了则不加Rule
                    '先找到最近一个SELECT
                    strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrStr(i)) - 1)
                    strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                    If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)               '取后面3个字符
                    
                    If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                       strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrStr(i)) + Len(arrStr(i)))
                    Else
                        Exit For
                    End If
                Loop
            Next
            If i <= UBound(arrStr) Then
                strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
            End If
        End If
    End If
    
    strLog = strSQL
        
    i = 1
    Do While i <= Len(strLog)
        If InStr(i, strLog, "[") <= 0 Then
            i = i + 1
            GoTo makContinue1
        End If
        strLeft = Left(strLog, InStr(i, strLog, "[") - 1)
        strTmp = Mid(strLog, InStr(i, strLog, "["))
        If mdlPublic.AtString(strLeft) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '单引号内的字符串，并且格式非[0-99]
            i = i + 1
            GoTo makContinue1
        End If
        
        If InStr(i, strLog, "]") <= 0 Then
            i = i + 1
            GoTo makContinue1
        End If
        strRight = Mid(strLog, InStr(i, strLog, "]") + 1)
        If strRight <> "" And mdlPublic.AtString(strRight) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '单引号内的字符串，并且格式非[0-99]
            i = i + 1
            GoTo makContinue1
        End If
        
        '单引号外的参数串
        i = InStr(i, strLog, "[")
        strRight = Mid(strLog, InStr(i, strLog, "]") + 1)
        
        StrPar = Mid(strLog, InStr(i, strLog, "[") + 1, InStr(i, strLog, "]") - InStr(i, strLog, "[") - 1)
        strParOld = StrPar
        bytPar = Val(StrPar)
        Select Case objData.Pars("_" & CInt(bytPar)).类型
            Case 0 '字符
                StrPar = "'" & Replace(objData.Pars("_" & CInt(bytPar)).缺省值, "'", "''") & "'"
            Case 1 '数字
                StrPar = objData.Pars("_" & CInt(bytPar)).缺省值
            Case 2 '日期
                If Left(objData.Pars("_" & CInt(bytPar)).缺省值, 1) = "&" Then
                    StrPar = GetParSQLMacro(objData.Pars("_" & CInt(bytPar)).缺省值)
                Else
                    If Format(objData.Pars("_" & CInt(bytPar)).缺省值, "HH:mm:ss") = "00:00:00" Then
                        '短时间格式
                        StrPar = "To_Date('" & Format(objData.Pars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                    Else
                        '长时间格式
                        StrPar = "To_Date('" & Format(objData.Pars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                End If
            Case 3 '无类型
                StrPar = objData.Pars("_" & CInt(bytPar)).缺省值
        End Select
        strLog = strLeft & StrPar & strRight
        
        i = Len(strLeft & strParOld)
        
makContinue1:
    Loop
        
    If InStr(UCase(objData.SQL), "--UNBOUND") > 0 Then GoTo LineOld

    '创建绑定参数SQL
    cmdData.CommandText = ""                        '不为空有时清除参数出错
    cmdData.CommandType = adCmdText                 '设置为adCmdText性能更优
    
    '清除原有参数:不然不能重复执行
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    l = 1
    Do While l <= Len(strSQL)
        If InStr(l, strSQL, "[") <= 0 Then
            l = l + 1
            GoTo makContinue2
        End If
        strLeft = Left(strSQL, InStr(l, strSQL, "[") - 1)
        strTmp = Mid(strSQL, InStr(l, strSQL, "["))
        If mdlPublic.AtString(strLeft) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '单引号内的字符串，并且格式非[0-99]
            l = l + 1
            GoTo makContinue2
        End If
        
        If InStr(l, strSQL, "]") <= 0 Then
            l = l + 1
            GoTo makContinue2
        End If
        strRight = Mid(strSQL, InStr(l, strSQL, "]") + 1)
        If strRight <> "" And mdlPublic.AtString(strRight) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '单引号内的字符串，并且格式非[0-99]
            l = l + 1
            GoTo makContinue2
        End If
        
        '单引号外的参数串
        l = InStr(l, strSQL, "[")
        strRight = Mid(strSQL, InStr(l, strSQL, "]") + 1)
        
        StrPar = Mid(strSQL, InStr(l, strSQL, "[") + 1, InStr(l, strSQL, "]") - InStr(l, strSQL, "[") - 1)
        strParOld = StrPar
        bytPar = Val(StrPar)
        intDateType = 0
        datValue = CDate(0)
        strTmp = ""
        
        Select Case objData.Pars("_" & CInt(bytPar)).类型
            Case 0 '字符
                StrPar = objData.Pars("_" & CInt(bytPar)).缺省值
                intMax = LenB(StrConv(StrPar, vbFromUnicode))
                
                If intMax <= 2000 Then
                    intMax = IIF(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarChar, adParamInput, intMax, StrPar)
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adLongVarChar, adParamInput, intMax, StrPar)
                End If
                
                strSQL = strLeft & "?" & strRight
            Case 1 '数字
                StrPar = objData.Pars("_" & CInt(bytPar)).缺省值
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarNumeric, adParamInput, 30, Val(StrPar))

                strSQL = strLeft & "?" & strRight
            Case 2 '日期
                If Left(objData.Pars("_" & CInt(bytPar)).缺省值, 1) = "&" Then
                    StrPar = GetParVBMacro(objData.Pars("_" & CInt(bytPar)).缺省值)
                Else
                    If Format(objData.Pars("_" & CInt(bytPar)).缺省值, "HH:mm:ss") = "00:00:00" Then
                        '短时间格式
                        StrPar = Format(objData.Pars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd")
                    Else
                        '长时间格式
                        StrPar = Format(objData.Pars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd HH:mm:ss")
                    End If
                End If
'                1）如果不存在日期（加减）运算，则直接使用日期型绑定变量；
'                2）如果存在日期（加减）运算，如果后面是常数，则先执行一条SQL或者在程序中计算，得到计算后的值（例如：XX+1/24），再根据得到的值使用日期型绑定变量。
'                      如果后面不是常数（例如：对与某个字段或sysdate等运算），则不使用绑定变量，直接代入参数值(sql拼接)。
'                      这种不使用绑定变量，虽然每次执行需要硬解析，但相对于执行计划可能出错来比较，代价更小一些。
 '                     只识别常用的加减算法（1、+1-1/24/60/60  2、 -1/24/60/60+1 ，3、-1/24/60/60  4、加减一个数字，没有连加的情况
 '                     如果不满足常用算法，如连加，+1/24 这种保持以前的规则，传入字符绑定变量
 '测试SQL：
'                select * from 部门表 where 撤档时间 >1+ [0]- 1  and  ID>0
'                Union All
'                select * from 部门表 where  [0]- 1=撤档时间  and  ID>0
'                Union All
'                select * from 部门表 where  撤档时间>[0]+1 - 1/24 /60 /60  and  ID>0
'                Union All
'                select * from 部门表 where  撤档时间>[0] - 1/24 /60 /60+1  and  ID>0
'                Union All
'                select * from 部门表 where  撤档时间>[0] - 1 /24 /60 /60  and  ID>0
'                Union All
'                select * from 部门表 where  撤档时间>[0] - 1 /24 /60   and  ID>0
'                Union All
'                select * from 部门表 where  撤档时间>1+[0] - 1 /24 /60/60   and  ID>0

                '先查看后边是否有加减运算
                datValue = CDate(StrPar)
                
                For i = 1 To Len(strRight)
                    If Mid(strRight, i, 1) <> " " Then
                        If InStr("+-", Mid(strRight, i, 1)) > 0 Then
                            For j = i + 1 To Len(strRight)
                                If Mid(strRight, j, 1) <> " " Then
                                    '找到计算的值
                                    For k = j + 1 To Len(strRight)
                                        If Mid(strRight, k, 1) = " " Or (IsNumeric(Mid(strRight, j, 1)) And Not IsNumeric(Mid(strRight, j, k - j + 1))) Then
                                            If Not Mid(strRight, k, 1) = " " And Not IsNumeric(Mid(strRight, k - 1, 1)) Then
                                                k = k - 1
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    If IsNumeric(Mid(strRight, j, k - j)) Then
                                        intDateType = 1
                                        
                                        '计算具体值
                                        '常见的计算方式优先判断长的
                                        If InStr(Replace(strRight, " ", ""), "+1-1/24/60/60") = 1 Then
                                            datValue = datValue + 1 - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "60") + 2), "60") + 2 + InStr(strRight, "60") + 1)
                                        ElseIf InStr(Replace(strRight, " ", ""), "-1/24/60/60+1") = 1 Then
                                            datValue = datValue + 1 - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "+") + 1), "1") + InStr(strRight, "+") + 1)
                                        ElseIf InStr(Replace(strRight, " ", ""), "-1/24/60/60") = 1 Then
                                            datValue = datValue - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "60") + 2), "60") + 2 + InStr(strRight, "60") + 1)
                                        Else
                                            If Mid(strRight, i, 1) = "+" Then
                                                datValue = datValue + Val(Mid(strRight, j, k - j))
                                            Else
                                                datValue = datValue - Val(Mid(strRight, j, k - j))
                                            End If
                                            strTmp = Mid(strRight, k)
                                        End If
                                        If InStr("+-*/", Mid(Replace(strTmp, " ", ""), 1, 1)) > 0 And Replace(strRight, " ", "") <> "" Then
                                            '如果后面没有+-*/则表示是单独的加减法,否则保持以前的规则
                                            intDateType = 3
                                        End If
                                    Else
                                        intDateType = 2
                                    End If
                                    Exit For
                                End If
                            Next
                        Else
                            Exit For
                        End If
                        Exit For
                    End If
                Next
                '前面加减不做处理，保持传入字符绑定变量的规则
                If intDateType <> 2 Then
                    For i = Len(strLeft) To 1 Step -1
                        If Mid(strLeft, i, 1) <> " " Then
                            If InStr("+-", Mid(strLeft, i, 1)) > 0 Then
                               intDateType = 3
                            End If
                            Exit For
                        End If
                    Next
                End If
                If intDateType = 2 Then
                    '不使用绑定变量
                    strSQL = strLeft & "To_Date('" & Format(datValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & strRight
                ElseIf intDateType = 3 Then
                    '将日期转换为字符类型到SQL中绑定,因为日期类型的绑定变量与数字运算要出错
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarChar, adParamInput, Len(StrPar), StrPar)
                    If StrPar Like "*:*:*" Then
                        strSQL = strLeft & "To_Date(?,'YYYY-MM-DD HH24:MI:SS')" & strRight
                    Else
                        strSQL = strLeft & "To_Date(?,'YYYY-MM-DD')" & strRight
                    End If
                Else
                    '在这里赋值主要是处理 后面满足条件的，但前面又有运算的
                    If intDateType = 1 Then strRight = strTmp
                    '如果是常数加减，或没有加减则直接使用日期绑定变量
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adDBTimeStamp, adParamInput, , datValue)
                    strSQL = strLeft & "?" & strRight
                End If
            Case 3 '无类型
                StrPar = objData.Pars("_" & CInt(bytPar)).缺省值

                strSQL = strLeft & StrPar & strRight
        End Select
        
        l = Len(strLeft & strParOld)
        
makContinue2:
    Loop
    
    '存在ROLLUP的或者WITH打头的都加上SELECT* 嵌套
    If InStr(strSQL, "ROLLUP") > 0 Or Mid(strSQL, 1, 4) = "WITH" Then
        strSQL = "SELECT * FROM(" & strSQL & ")"
    End If
    '执行返回记录集
'    If cmdData.ActiveConnection Is Nothing Then
'        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
'    End If
    Set cmdData.ActiveConnection = mdlPublic.GetDBConnection(objData.数据连接编号)
    cmdData.CommandText = strSQL

LineBand:
    Call SQLTest(App.ProductName, "OpenReportSQL", strLog)
    
    If SQLExistLOB(objData) Then
        If rsTmp Is Nothing Then Set rsTmp = New ADODB.Recordset
        rsTmp.Open cmdData, , adOpenStatic, adLockOptimistic
    Else
        Set rsTmp = cmdData.Execute
    End If
    
    Set rsTmp.ActiveConnection = Nothing        '与zl9ComLib的Recordset对象处理一致
    Call SQLTest
    Set OpenReportSQL = rsTmp
    Exit Function
    
LineOld:
    Call OpenRecord(rsTmp, strLog, "OpenReportSQL", objData.数据连接编号)
    Set rsTmp.ActiveConnection = Nothing        '与zl9ComLib的Recordset对象处理一致
    Set OpenReportSQL = rsTmp
    Exit Function
    
LineBlob:
    '由于ODBC不支持含有LOB列的SQL，所以改为使用OLEDB查询
    If objData.数据连接编号 <= 0 Then
        If gcolOLEDBConnect Is Nothing Then
            Set gcolOLEDBConnect = New Collection
        End If
        '获取缓存连接对象
        Set gcnOLEDB = mdlPublic.GetOLEDBConnect(gcnOracle, gcolOLEDBConnect, gobjRegister)
        If gcnOLEDB Is Nothing Then
            Set gcnOLEDB = gobjRegister.ReGetConnection(Val("1-OracleOLEDB"), "", gcnOracle)
            
            If gcnOLEDB.State = adStateClosed Then
                strTmp = "创建数据连接失败，请检查：" & Chr(10) & _
                         "1.业务程序使用zlRegister部件方法的参数是否正确；" & Chr(10) & _
                         "2.非导航台程序调用报表部件传入的连接对象（微软驱动），必须通过" & Chr(10) & _
                         "  zlRegister部件创建连接对象。"
                If gblnSilentMode Then
                    gstrErrorContent = strTmp
                Else
                    MsgBox strTmp, vbInformation, App.Title
                End If
                Exit Function
            End If
            
            '缓存
            Call gcolOLEDBConnect.Add(gcnOLEDB)
        End If
        Set cmdData.ActiveConnection = gcnOLEDB
    Else
        Set cmdData.ActiveConnection = mdlPublic.GetDBConnectionEx(Val("1-OracleOLEDB"), objData.数据连接编号)
    End If
    If Not cmdData.ActiveConnection Is Nothing Then
        'Set rsTmp = cmdData.Execute
        'CLOB、BLOB字段类型如果使用Command对象，记录集对象默认的锁adOpenUnspecified执行会很慢
        '因此，改用记录集对象的Open方法
        If rsTmp Is Nothing Then Set rsTmp = New ADODB.Recordset
        rsTmp.Open cmdData, , adOpenStatic, adLockOptimistic
        Set rsTmp.ActiveConnection = Nothing        '与zl9ComLib的Recordset对象处理一致
        Set OpenReportSQL = rsTmp
    End If
    Exit Function
    
errHandle:
    'ORA-00979:不是 GROUP BY 表达式
    'SQL中的"?"在提交给Oracle时被ADO顺序换为":P1,:P2"方式,Group by认为这是不同的分组字段,即使绑定值相同
    '老连接方式下,ADO的SQL中不能使用":P"这种参数(始终说变量未关联)
    '新连接方式下,ADO的SQL中可以使用":P"这种参数,但和Parameters对象中创建的只是顺序对应,名称不对应.
    '    即可通过名称使用Group中涉及参数的相同字段一致,但不能因为名称相同而少创建一些参数
    If Err.Description Like "*ORA-00979*" Then Err.Clear: GoTo LineOld
    
    'ORA-00932: 数据类型不一致: 应为 NUMBER, 但却获得 -
    '当Group By Rollup和Decode混用时，可能会出现该错误，暂未知明确原因,无明确解决办法
    '实验情况看当出现该错误时，SQL尚未真正执行，速度上应没有影响
    If Err.Description Like "*ORA-00932*" Then Err.Clear: GoTo LineOld
    
     'MS的ODBC连接，查询BLOB字段时会报错(执行一个提供程序命令时数据提供程序失败。)，这里临时改为OraOLEDB连接对象来访问
    If Err.Number = -2147467259 Then Err.Clear: GoTo LineBlob
    
    Call ShowFlash
    If Err.Description Like "*ORA-00920*" Then
        MsgBox "参数输入错误，导致不能正确读取数据""" & objData.名称 & """！", vbExclamation, App.Title
    ElseIf ErrCenter() = 1 Then
        If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "正在读取数据""" & objData.名称 & """，请稍候．．．", , Me
        Resume
    End If
    Call SaveErrLog
End Function

Private Function EvalFormula(ByVal strFormula As String, idx As Integer, Row As Long) As String
'功能：计算表达式的值
'参数：strFormula=表达公式,idx:已处理一部份数据的表格索引,Row=当前行
'返回：计算后的值,计算错误返回空
'参考：mLibDatas
    Dim strLeft As String, strRight As String, strVar As String
    
    On Error Resume Next
    
    strFormula = Trim(strFormula)
    
    If strFormula = "" Then '空列
        Exit Function
    ElseIf InStr(strFormula, "[") = 0 Then '纯计算列
        EvalFormula = Srt.Eval(strFormula)
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         '只有字段引用的列
         EvalFormula = GetFieldValue(Me, Mid(strFormula, 2, Len(strFormula) - 2))
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         '只有列引用的列
         EvalFormula = msh(idx).TextMatrix(Row, CInt(Mid(strFormula, 2, Len(strFormula) - 2)))
    Else '复合计算
        Do While InStr(strFormula, "[") > 0
            strLeft = Left(strFormula, InStr(strFormula, "[") - 1)
            strRight = Mid(strFormula, InStr(strFormula, "]") + 1)
            strVar = Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1)
            
            If IsNumeric(Mid(strVar, 2)) And Left(strVar, 1) = "@" Then
                If Row = msh(idx).FixedRows Then
                    strVar = "" '第一数据行无法取
                Else
                    If InStr(strFormula, """[" & strVar & "]""") > 0 And InStr(msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))), """") > 0 Then
                        '字符串运算及单元值中包含字符串
                        strVar = Replace(msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))), """", """""")
                    Else
                        strVar = msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))) '直接取对应列值
                    End If
                End If
                If strVar = "" Then strVar = 0
            ElseIf IsNumeric(strVar) Then
                If InStr(strFormula, """[" & strVar & "]""") > 0 And InStr(msh(idx).TextMatrix(Row, CInt(strVar)), """") > 0 Then
                    '字符串运算及单元值中包含字符串
                    strVar = Replace(msh(idx).TextMatrix(Row, CInt(strVar)), """", """""")
                Else
                    strVar = msh(idx).TextMatrix(Row, CInt(strVar)) '直接取对应列值
                End If
                If strVar = "" Then strVar = 0
            ElseIf InStr(strVar, ".") > 0 Then
                '如果为空,返回"Null",表达式要作处理判断
                strVar = GetFieldValue(Me, strVar, True) '可能有日期或字符参与运算时,自动转换格式
            End If
            
            '替换避免死循环
            If InStr(strVar, "[") > 0 Or InStr(strVar, "]") > 0 Then
                strVar = Replace(strVar, "[", Chr(1) & "SKIPCYCLEFT" & Chr(1))
                strVar = Replace(strVar, "]", Chr(1) & "SKIPCYCRIGHT" & Chr(1))
            End If
            strFormula = strLeft & strVar & strRight
        Loop
        strFormula = Replace(strFormula, Chr(1) & "SKIPCYCLEFT" & Chr(1), "[")
        strFormula = Replace(strFormula, Chr(1) & "SKIPCYCRIGHT" & Chr(1), "]")
        EvalFormula = Srt.Eval(strFormula)
    End If
End Function

Private Function SortFormula(objItem As RPTItem) As Variant
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrFormula() As String, strTmp As String
    Dim strReferCols As String, intReferCols As Integer
    Dim intCol As Integer, intCur As Integer
    Dim i As Integer, j As Integer
    Dim strDie As String, strOrder As String
    
    
    ReDim arrFormula(objItem.SubIDs.count - 1) As String
    
    '先顺序放在数组中
    For Each tmpID In objItem.SubIDs
        Set tmpItem = mobjReport.Items("_" & tmpID.id)
        arrFormula(tmpItem.序号) = tmpItem.内容 & "|" & tmpItem.格式 & "|" & tmpItem.序号 & "|" & tmpItem.汇总
    Next
    
    '根据"内容"的包含关系排序
    i = 0
    strOrder = GetOrder(arrFormula)
    Do While i <= UBound(arrFormula)
        '引用了哪些列
        strReferCols = GetReferCols(CStr(Split(arrFormula(i), "|")(0)))
        intReferCols = UBound(Split(strReferCols, ","))
        
        intCur = i '该项当前位置
        For j = 0 To intReferCols
            '引用项当前位置
            intCol = GetReferLoc(arrFormula, CInt(Split(strReferCols, ",")(j)))
            If intCol > intCur Then
                strTmp = arrFormula(intCur)
                arrFormula(intCur) = arrFormula(intCol)
                arrFormula(intCol) = strTmp
                intCur = intCol
            End If
        Next
        '如果一次也没有换,表示没有引用,则继续下一个排序
        '否则仍然从老位置开始分析排序
        strDie = GetOrder(arrFormula)
        If intCur = i Or (intCur <> i And strOrder = strDie) Then
            i = i + 1
            strOrder = strDie
        End If
    Loop
    
    SortFormula = arrFormula
End Function

Private Function GetOrder(arrFormula() As String) As String
'功能：返回当前公式数组中各列号排列的顺序,用于检查死循环
    Dim i As Integer
    For i = 0 To UBound(arrFormula)
        GetOrder = GetOrder & "," & CInt(Split(arrFormula(i), "|")(2))
    Next
    GetOrder = Mid(GetOrder, 2)
End Function

Private Function GetReferLoc(arrFormula() As String, intCol As Integer) As Integer
'功能：返回序号为intCol的项目在数组中的位置
    Dim i As Integer
    For i = 0 To UBound(arrFormula)
        If CInt(Split(arrFormula(i), "|")(2)) = intCol Then
            GetReferLoc = i: Exit Function
        End If
    Next
End Function

Private Function GetReferCols(ByVal strFormula As String) As String
'功能：返回公式中引用的列号,如"3,5,6"
    Dim strRight As String, strCol As String, strCols As String
    
    strFormula = Trim(strFormula)
    
    Do While InStr(strFormula, "[") > 0
        strRight = Mid(strFormula, InStr(strFormula, "]") + 1)
        strCol = Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1)
        If IsNumeric(strCol) Then strCols = strCols & "," & strCol
        strFormula = strRight
    Loop
    GetReferCols = Mid(strCols, 2)
End Function

Private Sub SetRedraw(blnDraw As Boolean)
    Dim obj As Object
    For Each obj In msh
        If obj.Index <> 0 And (obj.Container Is picPaper(intReport) Or UCase(obj.Container.name) = "PIC") Then obj.Redraw = blnDraw
    Next
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Function GetLR(msh As Object, Col As Integer) As Byte
    Select Case msh.ColAlignment(Col)
        Case 0, 1, 2 '左对齐
            GetLR = 2 '右加空格
        Case 3, 4, 5 '中对齐
            GetLR = 1 '双加空格
        Case 6, 7, 8 '右对齐
            GetLR = 0 '左加空格
    End Select
End Function

Private Function GetRowText(msh As Object, Row As Long, Col As Long) As String
    Dim i As Integer
    Dim strTmp As String
    
    For i = 0 To Col
        strTmp = strTmp & Trim(msh.TextMatrix(Row, i))
    Next
    GetRowText = strTmp
End Function

Private Function GetColText(msh As Object, Row As Long, Col As Long) As String
    Dim i As Integer
    Dim strTmp As String
    
    For i = 0 To Row
        strTmp = strTmp & Trim(msh.TextMatrix(i, Col))
    Next
    GetColText = strTmp
End Function

Private Function GetColType(ByVal strFormula As String) As Byte
'功能：判断任意表格某一列的数据类型
'参数：strFormula=列计算公式
'返回：0=不确定,1-字符(其它),2=数字,3=日期
'参考：mLibDatas
    Dim varR As Variant, strData As String, strField As String
    
    On Error Resume Next
    
    strFormula = Trim(strFormula)
    
    If strFormula = "" Then '空列
        GetColType = 1
    ElseIf InStr(strFormula, "[") = 0 Then '纯计算列
        varR = Srt.Eval(strFormula)
        If IsNumeric(varR) Then
            GetColType = 2
        ElseIf IsDate(varR) Then
            GetColType = 3
        Else
            GetColType = 1
        End If
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         '只有字段引用的列
        strFormula = Mid(strFormula, 2, Len(strFormula) - 2)
        strData = Left(strFormula, InStr(strFormula, ".") - 1)
        strField = Mid(strFormula, InStr(strFormula, ".") + 1)
        
        Select Case mLibDatas("_" & strData).DataSet.Fields(strField).type
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                GetColType = 1
            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                GetColType = 2
            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                GetColType = 3
        End Select
    End If
End Function

Private Function GetParsStr(ByVal objPars As RPTPars) As String
'功能：将不重复名称的参数集中的内容转换为字符串
'返回："参数名=参数值|参数名=参数值..."
'说明：如果报表具有多个格式,同时包含"ReportFormat=x"
    Dim tmpPar As RPTPar
    Dim strPars As String
    
    If mobjReport.Fmts.count > 1 Then
        strPars = strPars & "|ReportFormat=" & bytFormat
    End If
    
    For Each tmpPar In objPars
        If tmpPar.缺省值 Like "&*" And tmpPar.类型 = 2 Then
            strPars = strPars & "|" & tmpPar.名称 & "=" & GetParVBMacro(tmpPar.缺省值)
        Else
            strPars = strPars & "|" & tmpPar.名称 & "=" & tmpPar.缺省值
        End If
    Next
    GetParsStr = Mid(strPars, 2)
End Function

Private Function SetStrPars(ByVal strPars As String, ByVal objPars As RPTPars) As RPTPars
'功能：将字符串中的参数描述内容填写到参数对象中
'参数：strPars="参数名=参数值|参数名=参数值..."
'返回：被设置的参数对象集
'说明：如果当前报表有多种格式且参数描述中有格式指定,则替换
    Dim tmpPar As RPTPar, tmpPars As RPTPars
    Dim i As Integer, j As Integer
    Dim bytTmp As Byte, strTmp As String
    
    If strPars = "" Or Not strPars Like "*=*" Then Set SetStrPars = objPars: Exit Function
    
    Set tmpPars = objPars
    
    For i = 0 To UBound(Split(strPars, "|"))
        For Each tmpPar In tmpPars
            strTmp = Split(strPars, "|")(i)
            If UCase(Split(strTmp, "=")(0)) = UCase("ReportFormat") And mobjReport.Fmts.count > 1 Then
                If IsNumeric(Split(strTmp, "=")(1)) Then
                    bytTmp = CByte(Split(strTmp, "=")(1))
                    For j = 1 To cboFormat.ComboItems.count
                        If CByte(Mid(cboFormat.ComboItems(j).Key, 2)) = bytTmp Then
                            cboFormat.ComboItems(j).Selected = True
                            bytFormat = bytTmp: mobjReport.bytFormat = bytFormat: Exit For
                        End If
                    Next
                End If
            ElseIf UCase(tmpPar.名称) = UCase(Split(strTmp, "=")(0)) Then
                Select Case tmpPar.类型
                    Case 1 '数字型
                        If IsNumeric(Split(strTmp, "=")(1)) Then tmpPar.缺省值 = Split(strTmp, "=")(1)
                    Case 2 '日期型
                        If IsDate(Split(strTmp, "=")(1)) Then tmpPar.缺省值 = Split(strTmp, "=")(1)
                    Case Else
                        tmpPar.缺省值 = Split(strTmp, "=")(1)
                End Select
            End If
        Next
    Next
    Set SetStrPars = tmpPars
End Function

Private Sub mnuFile_Excel_Click()
    Dim lngRow As Long, lngCol As Long
    Dim bytKind As Byte, tmpMsh As Object
    Dim i As Long, j As Long
    
    '必要条件检查
    If Not mobjReport.blnLoad Then Exit Sub
    
    If zlRegInfo("授权性质") <> "1" Then
        MsgBox "试用或测试版本不能使用该功能。", vbInformation, App.Title
        Exit Sub
    End If
    
    If isExporting Then
        gblnError = True
        MsgBox "另外一张报表正在输出到 Excel,请稍候再执行该操作！", vbInformation, App.Title
        Exit Sub
    End If
    
    If intGridCount = 0 Then
        MsgBox "报表中没有数据表表格可以输出到 Excel！", vbInformation, App.Title
        Exit Sub
    End If
    If objCurGrid Is Nothing Then
        If msh.count > 1 Then
            For Each tmpMsh In msh
                If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") And Not tmpMsh.Tag Like "H_*" Then
                    Set objCurGrid = tmpMsh
                    Exit For
                End If
            Next
        End If
        If objCurGrid Is Nothing Then
            MsgBox "请先选择一个要输出到 Excel的数据表！", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    If Not HaveExcel Then
        gblnError = True
        MsgBox "系统检测到本机没有安装 Microsoft Excel 程序,操作不能继续！", vbInformation, App.Title
        Exit Sub
    End If
    
    '确定[表头]、表体
    Set gobjHead = Nothing
    Set gobjBody = Nothing
    
    Set gobjBody = objCurGrid
    If Val(objCurGrid.Tag) > 0 Then
        bytKind = GetGridStyle(mobjReport, objCurGrid.Index)
        If bytKind <> 2 Then Set gobjHead = msh(CInt(objCurGrid.Tag))
    End If
    
    '产生附加标签项目
    Call MakeAppend(Me, picPaper(intReport))
    
    '输出到Excel
    lngRow = gobjBody.Row
    lngCol = gobjBody.Col
    If Not gobjHead Is Nothing Then gobjHead.Redraw = False
    gobjBody.Redraw = False
    
    '替换回车换行符
    If Not gobjHead Is Nothing Then
        For i = 0 To gobjHead.Rows - 1
            For j = 0 To gobjHead.Cols - 1
                gobjHead.TextMatrix(i, j) = Replace(Replace(Replace(gobjHead.TextMatrix(i, j), vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
            Next
        Next
    End If
    For i = 0 To gobjBody.Rows - 1
        For j = 0 To gobjBody.Cols - 1
            gobjBody.TextMatrix(i, j) = Replace(Replace(Replace(gobjBody.TextMatrix(i, j), vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
        Next
    Next
    
    blnExcel = True
    Call ExportExcel(Me, IIF(mbytStyle = 3, mstrExcelFile, ""))
    
    gobjBody.Row = lngRow
    gobjBody.Col = lngCol
    Call msh_EnterCell(gobjBody.Index)
    If Not gobjHead Is Nothing Then
        gobjHead.Redraw = True
    
        '恢复回车换行符
        For i = 0 To gobjHead.Rows - 1
            For j = 0 To gobjHead.Cols - 1
                gobjHead.TextMatrix(i, j) = Replace(gobjHead.TextMatrix(i, j), "<换行分隔符>", vbCrLf)
            Next
        Next
    End If
    For i = 0 To gobjBody.Rows - 1
        For j = 0 To gobjBody.Cols - 1
            gobjBody.TextMatrix(i, j) = Replace(gobjBody.TextMatrix(i, j), "<换行分隔符>", vbCrLf)
        Next
    Next
    gobjBody.Redraw = True
    
End Sub

Public Function DelUnUseData(objReport As Report) As Boolean
'功能：从对象mobjReport中删除未使用的数据源对象
'返回：是否存在未使用的数据源
'说明：1.该函数只在打开报表前调用一次
'      2.是删除所有报表格式中未使用的。
    Dim tmpData As RPTData
    
    If objReport Is Nothing Then Exit Function
    
    For Each tmpData In objReport.Datas
        If Not DataUsed(objReport, tmpData.名称) Then objReport.Datas.Remove "_" & tmpData.Key
    Next
End Function

Private Function GetStatText(strStat As String) As String
    Select Case strStat
        Case "SUM"
            GetStatText = "合计"
        Case "AVG"
            GetStatText = "平均值"
        Case "MAX"
            GetStatText = "最大值"
        Case "MIN"
            GetStatText = "最小值"
        Case "COUNT"
            GetStatText = "记录数"
    End Select
End Function

Public Sub AddCol(msh As Object, Optional ByVal intCol As Integer = -1, Optional ByVal intCols As Integer = 1)
'功能：在指定表格msh中插入intCols个列,插入的第一列的列号为intCol,如果没有intCol参数,则追加列
'说明：插入列后,只对数据处理,对格式不作处理(包括列宽)
    Dim i As Integer, j As Integer, k As Integer
    
    If intCol >= msh.Cols Then intCol = -1
    msh.Cols = msh.Cols + intCols
    If intCol = -1 Then Exit Sub
    '移动数据
    For j = msh.Cols - 1 To intCol + intCols Step -intCols
        For i = 0 To msh.FixedRows - 1
            For k = 0 To intCols - 1
                msh.TextMatrix(i, j - k) = msh.TextMatrix(i, j - k - intCols)
                msh.ColData(j - k) = msh.ColData(j - k - intCols)
            Next
        Next
    Next
    '新列数据清除
    For j = intCol To intCol + intCols - 1
        For i = 0 To msh.FixedRows - 1
            msh.TextMatrix(i, j) = ""
        Next
    Next
End Sub

Private Sub ShowFreeGrid(objItem As RPTItem)
'功能：在查询界面上组织显示一个任意表格
    Dim strData As String, strTmp As String, bytKind As Byte
    Dim lngCol As Long, strState As String, arrState() As Variant
    Dim mshBody As Object, mshHead As Object
    Dim tmpItem As RPTItem, tmpID As RelatID, objBody As RPTItem
    Dim strValue As String, lngHead As Long, arrHead() As String
    Dim strSource As String, arrSource() As String, arrFormula() As String, arrField() As String
    Dim strFormula As String, strFormat As String, arrType() As Long
    Dim i As Long, j As Long, k As Long, l As Long, blnDo As Boolean
    Dim arrRowIDs() As Variant, strIDSource As String, strFirstSource As String
    Dim objPic As StdPicture
    Dim objColProtertys As RPTColProtertys
    Dim varIFValue As Variant
    Dim blntmp As Boolean, blnRPTLink As Boolean
    
    arrRowIDs = Array()
    On Error GoTo hErr
    
    With objItem
        bytKind = GetGridStyle(mobjReport, .id)
        Load msh(.id) '表体部份
        Load msh(.SubIDs(1).id) '表头部份
        Set msh(.id).Container = picPaper(intReport)
        Set msh(.SubIDs(1).id).Container = picPaper(intReport)
        If .父ID <> 0 Then
            Set msh(.id).Container = pic(.父ID)
            Set msh(.SubIDs(1).id).Container = pic(.父ID)
        End If
        Set mshBody = msh(.id)
        Set mshHead = msh(.SubIDs(1).id) '利用第一列的ID作为控件索引
        
        mshBody.Redraw = False
        mshHead.Redraw = False
                            
        mshHead.Tag = "H_" & mshBody.Index '标志该表格为固定表头
        mshBody.Tag = mshHead.Index
        
        '表格外形
        '表头
        mshHead.Left = .X: mshHead.Top = .Y
        mshHead.Width = .W: mshHead.Height = .H '为了使表头可滚动
        
        mshHead.Cols = .SubIDs.count
        mshHead.FixedCols = 0
        mshHead.Rows = UBound(Split(mobjReport.Items("_" & .SubIDs(1).id).表头, "|")) + 2
        mshHead.RowHeight(mshHead.Rows - 1) = 0
        mshHead.FixedRows = mshHead.Rows - 1
        
        mshHead.ForeColor = .前景
        mshHead.ForeColorFixed = .前景
        mshHead.BackColor = .背景
        mshHead.BackColorFixed = .背景
        mshHead.GridColor = .网格
        mshHead.GridColorFixed = IIF(.格式 = "", .网格, Val(.格式))
        mshHead.Font.name = .字体
        mshHead.Font.Size = .字号
        mshHead.Font.Bold = .粗体
        mshHead.Font.Italic = .斜体
        mshHead.Font.Underline = .下线
        mshHead.GridLineWidth = IIF(.表格线加粗, 2, 1)
        'Set mshHead.FontFixed = mshHead.Font
        '支持排序
        mshHead.ExplorerBar = flexExSortShow

        '处理表头内容(单元对齐、行高、内容,列宽)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.id)
            If tmpItem.Relations.count > 0 Then
                If blnRPTLink = False Then blnRPTLink = True
            End If
            arrHead = Split(tmpItem.表头, "|")
            lngHead = 0 '纯表头部份高度
            For i = 0 To UBound(arrHead) '对齐^高度^内容
                mshHead.Col = tmpItem.序号: mshHead.Row = i
                mshHead.CellAlignment = CInt(Split(arrHead(i), "^")(0))
                
                mshHead.RowHeight(i) = CLng(Split(arrHead(i), "^")(1))
                lngHead = lngHead + mshHead.RowHeight(i)
                
                If CStr(Split(arrHead(i), "^")(2)) = "#" Then '为空
                    mshHead.TextMatrix(i, tmpItem.序号) = ""
                ElseIf CStr(Split(arrHead(i), "^")(2)) = "←" Then '与左边单元格相同
                    mshHead.TextMatrix(i, tmpItem.序号) = mshHead.TextMatrix(i, tmpItem.序号 - 1)
                ElseIf CStr(Split(arrHead(i), "^")(2)) = "↑" Then '与上边单元格相同
                    mshHead.TextMatrix(i, tmpItem.序号) = mshHead.TextMatrix(i - 1, tmpItem.序号)
                Else
                    strValue = CStr(Split(arrHead(i), "^")(2))
                    
                    '数据指针复位(可能用到多个数据源、多个字段)
                    '先处理数据字段(查询时只取第一个值)
                    strData = GetLabelDataName(strValue)
                    If strData <> "" Then
                        For j = 0 To UBound(Split(strData, "|"))
                            strTmp = Split(Split(strData, "|")(j), ".")(0)
                            If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                mLibDatas("_" & strTmp).DataSet.MoveFirst
                            End If
                            strTmp = GetFieldValue(Me, CStr(Split(strData, "|")(j)))
                            strValue = Replace(strValue, "[" & Split(strData, "|")(j) & "]", strTmp)
                        Next
                    End If
                    
                    '再处理报表变量:[=参数名]、[n>=0]、[日期格式串][单位名称]
                    strValue = GetLabelMacro(Me, strValue)
                    
                    mshHead.TextMatrix(i, tmpItem.序号) = strValue
                    
                End If
                If UBound(Split(arrHead(i), "^")) > 3 Then
                    '处理表头颜色、加粗
                    If Split(arrHead(i), "^")(3) = 1 Then
                        mshHead.Cell(flexcpFontBold, i, tmpItem.序号) = True
                    End If
                    mshHead.Cell(flexcpForeColor, i, tmpItem.序号) = Val(Split(arrHead(i), "^")(4))
                End If
            Next
            mshHead.ColWidth(tmpItem.序号) = tmpItem.W
        Next
        
        '表头任意合并
        For i = 0 To mshHead.FixedRows - 1
            mshHead.MergeRow(i) = True
        Next
        For i = 0 To mshHead.Cols - 1
            mshHead.MergeCol(i) = True
        Next
        
        '表体格式
        If bytKind = 2 Then '仅有表体
            mshBody.Top = .Y: mshBody.Left = .X
            mshBody.Height = .H: mshBody.Width = .W
        Else
            mshBody.Top = .Y + lngHead: mshBody.Left = .X
            If .H - lngHead + 15 < 0 Then
                mshBody.Height = 0
            Else
                mshBody.Height = .H - lngHead + 15
            End If
            mshBody.Width = .W
        End If
        mshBody.Cols = .SubIDs.count: mshBody.FixedCols = 0
        mshBody.Rows = 1: mshBody.FixedRows = 0 '空数据时只有一空行
        mshBody.RowHeight(0) = .行高
        mshBody.RowHeightMin = .行高
        
        mshBody.ForeColor = .前景
        mshBody.ForeColorFixed = .前景
        mshBody.BackColor = .背景
        mshBody.BackColorFixed = .背景
        mshBody.GridColor = .网格
        mshBody.GridColorFixed = .网格
        mshBody.Font.name = .字体
        mshBody.Font.Size = .字号
        mshBody.Font.Bold = .粗体
        mshBody.Font.Italic = .斜体
        mshBody.Font.Underline = .下线
        mshBody.GridLineWidth = IIF(.表格线加粗, 2, 1)

        'Set mshBody.FontFixed = mshBody.Font
        
        '表体内容(列宽，列对齐)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.id)
            With mshBody
                .ColData(tmpItem.序号) = tmpItem
                .ColWidth(tmpItem.序号) = tmpItem.W
                .ColAlignment(tmpItem.序号) = Switch(tmpItem.对齐 = Val("0-左"), flexAlignLeftCenter _
                                                   , tmpItem.对齐 = Val("1-中"), flexAlignCenterCenter _
                                                   , tmpItem.对齐 = Val("2-右"), flexAlignRightCenter)
                If .FixedRows - 1 >= 0 And .Rows - 1 >= 0 Then
                    .Cell(flexcpAlignment, .FixedRows - 1, tmpItem.序号, .Rows - 1, tmpItem.序号) = .ColAlignment(tmpItem.序号)
                End If
                .MergeCol(tmpItem.序号) = tmpItem.自调
            End With
        Next
        
        '--------------------------------------------------------------------------------------
        '处理表体数据
        '--------------------------------------------------------------------------------------
        '1.搜索该表格用到的数据源
        strSource = GetGridSource(objItem) '"病人信息,药品信息,..."
        
        '2.对表列公式包含关系排序
        arrFormula = SortFormula(objItem) '(数组元素="公式|格式|列号|汇总")
        
        '3.初始统计数据
        ReDim arrState(.SubIDs.count - 1)
        ReDim arrType(.SubIDs.count - 1) '各列数据类型(0=不确定,1-字符(其它),2-数字,3-日期)
        If strSource <> "" Then
            arrSource = Split(strSource, ",")
            strFirstSource = arrSource(0)
            ''第一次时判断是否有数据:只要一个有数据,则不结束
            blnDo = False
            For i = 0 To UBound(arrSource)
                If mLibDatas("_" & arrSource(i)).DataSet.RecordCount > 0 Then
                    mLibDatas("_" & arrSource(i)).DataSet.MoveFirst '指针复位
                End If
                blnDo = blnDo Or Not mLibDatas("_" & arrSource(i)).DataSet.EOF
                
                '确定有ID字段的数据源:以第一个为准
                If strIDSource = "" Then
                    For j = 0 To mLibDatas("_" & arrSource(i)).DataSet.Fields.count - 1
                        If UCase(mLibDatas("_" & arrSource(i)).DataSet.Fields(j).name) = "ID" Then
                            If IsType(mLibDatas("_" & arrSource(i)).DataSet.Fields(j).type, adNumeric) Then
                                strIDSource = arrSource(i): Exit For
                            End If
                        End If
                    Next
                End If
            Next

        Else
            blnDo = True
        End If
        
        mshHead.WordWrap = .自调
        mshBody.WordWrap = .自调
        
        '自适应行高
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.id)
            If Not tmpItem Is Nothing Then
                If tmpItem.自适应行高 Then
                    '列一旦设置自适应行高，整个表格对象就要设置该属性
                    mshBody.AutoSizeMode = flexAutoSizeRowHeight
                End If
            End If
        Next
        
        '4.组织数据
        j = 0
        Do While blnDo
            If j > 0 Then
                mshBody.Rows = mshBody.Rows + 1 '缺省有一行
                mshBody.RowHeight(mshBody.Rows - 1) = .行高
            End If
            
            '对表格对应的ID行数组进行赋值
            ReDim Preserve arrRowIDs(UBound(arrRowIDs) + 1)
            arrRowIDs(UBound(arrRowIDs)) = 0
            If strIDSource <> "" Then
                If Not mLibDatas("_" & strIDSource).DataSet.EOF Then
                    arrRowIDs(UBound(arrRowIDs)) = Val(Nvl(mLibDatas("_" & strIDSource).DataSet.Fields("ID").Value, 0))
                End If
            End If
            
            For i = 0 To UBound(arrFormula)
                arrField = Split(arrFormula(i), "|")
                strFormula = arrField(0)    '公式
                strFormat = arrField(1)     '格式
                lngCol = Val(arrField(2))   '列号
                strState = arrField(3)      '汇总
                
                '！输出数据
                strValue = EvalFormula(strFormula, mshBody.Index, j)
                'If gobjFile.FileExists(strValue) Then   '该方法速度慢，特别是在记录数多的情况下特别明显
                If LCase$(Right$(strValue, 4)) = ".pic" Then
                    Set objPic = Nothing
                    On Error Resume Next
                    Set objPic = LoadPicture(strValue)
                    gobjFile.DeleteFile strValue, True
                    On Error GoTo 0
                    
                    If Not objPic Is Nothing Then
                        mshBody.Row = j: mshBody.Col = lngCol
                        
                        Me.picTemp.Cls '不清有问题
                        If objPic.Height / objPic.Width < mshBody.CellHeight / mshBody.CellWidth Then
                            Me.picTemp.Width = mshBody.CellWidth
                            Me.picTemp.Height = (objPic.Height / objPic.Width) * mshBody.CellWidth
                        Else
                            Me.picTemp.Height = mshBody.CellHeight
                            Me.picTemp.Width = (objPic.Width / objPic.Height) * mshBody.CellHeight
                        End If
                        Me.picTemp.PaintPicture objPic, 0, 0, Me.picTemp.Width, Me.picTemp.Height
                                            
                        Set mshBody.CellPicture = Me.picTemp.Image
                        mshBody.CellPictureAlignment = 4 '固定中对齐
                    End If
                Else
                    mshBody.TextMatrix(j, lngCol) = strValue
                    mshBody.Cell(flexcpFont, j, lngCol) = mshBody.Font          '强制更新字体对象，单元格自适应行高才能生效
                    If (strIDSource <> "" Or strFirstSource <> "") And blnRPTLink = True Then
'                        Set colRelation = New Collection
'                        colRelation.Add mLibDatas("_" & IIF(strIDSource = "", strFirstSource, strIDSource)).DataSet.AbsolutePosition
'                        mshBody.Cell(flexcpData, j, lngCol) = colRelation
                        
                        '优化
                        '固定将RowData存放记录集的行号，第一行的单元格存放报表链接关系
                        mshBody.RowData(j) = mLibDatas("_" & IIF(strIDSource = "", strFirstSource, strIDSource)).DataSet.AbsolutePosition
                    End If
                    '列属性设置
                    Set objColProtertys = mshBody.ColData(lngCol).ColProtertys
                    If objColProtertys.count > 0 Then
                        For l = 1 To objColProtertys.count
                            If InStr(objColProtertys.Item(l).条件值, strIDSource & ".") > 0 Then
                                varIFValue = EvalFormula("[" & objColProtertys.Item(l).条件值 & "]", mshBody.Index, j)
                            Else
                                varIFValue = objColProtertys.Item(l).条件值
                            End If
                            If CheckColProtertys(EvalFormula("[" & objColProtertys.Item(l).条件字段 & "]", mshBody.Index, j), objColProtertys.Item(l).条件关系, varIFValue) Then
                                If objColProtertys.Item(l).是否整行应用 Then
                                    mshBody.Cell(flexcpBackColor, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(l).背景颜色
                                    mshBody.Cell(flexcpForeColor, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(l).字体颜色
                                    mshBody.Cell(flexcpFontBold, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(l).是否加粗
                                Else
                                    mshBody.Cell(flexcpBackColor, j, lngCol) = objColProtertys.Item(l).背景颜色
                                    mshBody.Cell(flexcpForeColor, j, lngCol) = objColProtertys.Item(l).字体颜色
                                    mshBody.Cell(flexcpFontBold, j, lngCol) = objColProtertys.Item(l).是否加粗
                                End If
                                
                                '对齐方式
                                Select Case objColProtertys.Item(l).对齐
                                Case Val("1-居左")
                                    mshBody.Cell(flexcpAlignment, j, lngCol) = flexAlignLeftCenter
                                Case Val("2-居中")
                                    mshBody.Cell(flexcpAlignment, j, lngCol) = flexAlignCenterCenter
                                Case Val("3-居右")
                                    mshBody.Cell(flexcpAlignment, j, lngCol) = flexAlignRightCenter
                                Case Else
                                    '缺省，不处理
                                End Select
                            End If
                        Next
                    End If
                End If
                
                '求列数据类型
                If j = 0 And (strState = "MAX" Or strState = "MIN") Then
                    arrType(lngCol) = GetColType(strFormula)
                    arrState(lngCol) = "初始值"
                End If
                If strState = "MAX" Or strState = "MIN" Then
                    If arrType(lngCol) = 0 Then
                        If IsNumeric(mshBody.TextMatrix(j, lngCol)) Then
                            arrType(lngCol) = 2
                        ElseIf IsDate(mshBody.TextMatrix(j, lngCol)) Then
                            arrType(lngCol) = 3
                        Else
                            arrType(lngCol) = 1
                        End If
                    End If
                End If
                
                '汇总数据
                On Error Resume Next
                If mshBody.TextMatrix(j, lngCol) <> "" Then
                    Select Case strState
                        Case "SUM", "AVG" '平均值先加(再除)
                            If IsNumeric(mshBody.TextMatrix(j, lngCol)) Then
                                arrState(lngCol) = arrState(lngCol) + CDbl(mshBody.TextMatrix(j, lngCol))
                            ElseIf IsDate(mshBody.TextMatrix(j, lngCol)) Then
                                arrState(lngCol) = arrState(lngCol) + CDate(mshBody.TextMatrix(j, lngCol))
                            Else
                                arrState(lngCol) = arrState(lngCol) + mshBody.TextMatrix(j, lngCol)
                            End If
                        Case "MAX"
                            If arrState(lngCol) = "初始值" Then
                                If arrType(lngCol) = 2 Then
                                    arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                ElseIf arrType(lngCol) = 3 Then
                                    arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                Else
                                    arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                End If
                            Else
                                If arrType(lngCol) = 2 Then
                                    If CDbl(mshBody.TextMatrix(j, lngCol)) > arrState(lngCol) Then
                                        arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                    End If
                                ElseIf arrType(lngCol) = 3 Then
                                    If CDate(mshBody.TextMatrix(j, lngCol)) > arrState(lngCol) Then
                                        arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                    End If
                                Else
                                    If mshBody.TextMatrix(j, lngCol) > arrState(lngCol) Then
                                        arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                    End If
                                End If
                            End If
                        Case "MIN"
                            If arrState(lngCol) = "初始值" Then
                                If arrType(lngCol) = 2 Then
                                    arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                ElseIf arrType(lngCol) = 3 Then
                                    arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                Else
                                    arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                End If
                            Else
                                If arrType(lngCol) = 2 Then
                                    If CDbl(mshBody.TextMatrix(j, lngCol)) < arrState(lngCol) Then
                                        arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                    End If
                                ElseIf arrType(lngCol) = 3 Then
                                    If CDate(mshBody.TextMatrix(j, lngCol)) < arrState(lngCol) Then
                                        arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                    End If
                                Else
                                    If mshBody.TextMatrix(j, lngCol) < arrState(lngCol) Then
                                        arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                    End If
                                End If
                            End If
                        Case "COUNT"
                            arrState(lngCol) = arrState(lngCol) + 1
                    End Select
                End If
                
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                
                '先汇总再格式化,减少误差
                If strFormat <> "" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = Format(mshBody.TextMatrix(j, lngCol), strFormat)
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
            Next
            
            If strSource <> "" Then
                '只要一个有数据,则不结束
                blnDo = False
                For i = 0 To UBound(arrSource)
                    If Not mLibDatas("_" & arrSource(i)).DataSet.EOF Then
                        mLibDatas("_" & arrSource(i)).DataSet.MoveNext
                    End If
                    blnDo = blnDo Or Not mLibDatas("_" & arrSource(i)).DataSet.EOF
                Next
            Else
                '如果没有用到数据源,则只有一行数据
                blnDo = False
            End If
            
            j = j + 1
        Loop
        
        '5.处理汇总行,只要有一列有汇总,则处理
        blnDo = False
        For i = 0 To UBound(arrFormula)
            blnDo = blnDo Or (Split(arrFormula(i), "|")(3) <> "")
            '优化
            If blnDo Then Exit For
        Next
        If blnDo Then
            mshBody.Rows = mshBody.Rows + 1
            mshBody.RowHeight(mshBody.Rows - 1) = .行高
            For i = 0 To UBound(arrFormula)
                arrField = Split(arrFormula(i), "|")
                strState = arrField(3)      '汇总
                lngCol = Val(arrField(2))   '列号
                strFormat = arrField(1)     '格式
                strFormula = arrField(0)    '公式
                If strState = "AVG" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = arrState(lngCol) / j
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                ElseIf strState <> "" Then
                    If TypeName(arrState(lngCol)) = "String" Then
                        If arrState(lngCol) = "初始值" Then arrState(lngCol) = ""
                    End If
                    mshBody.TextMatrix(j, lngCol) = arrState(lngCol)
                ElseIf strFormula <> "" Then
                    '汇总行中，没有汇总的列如果有公式，则计算公式
                    strValue = EvalFormula(strFormula, mshBody.Index, j)
                    'If gobjFile.FileExists(strValue) Then   '该方法速度慢，特别是在记录数多的情况下特别明显
                    If LCase$(Right$(strValue, 4)) = ".pic" Then
                        '图片字段
                        On Error Resume Next
                        gobjFile.DeleteFile strValue, True
                        On Error GoTo 0
                    Else
                        mshBody.TextMatrix(j, lngCol) = strValue
                    End If
                End If
                '格式化单元值
                If strFormat <> "" And mshBody.TextMatrix(j, lngCol) <> "" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = Format(mshBody.TextMatrix(j, lngCol), strFormat)
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
            Next
            '显示汇总标志
            For k = 0 To mshBody.Cols - 1
                If mshBody.ColWidth(k) > 0 Then Exit For
            Next
            If mshBody.TextMatrix(j, k) = "" Then
                blnDo = True: l = 0
                For i = 0 To UBound(arrFormula)
                    arrField = Split(arrFormula(i), "|")
                    If arrField(3) <> "" Then
                        If l = 0 Then
                            strState = arrField(3)
                        Else
                            blnDo = blnDo And (Split(arrFormula(i), "|")(3) = strState)
                        End If
                        l = l + 1
                    End If
                Next
                If blnDo Then '一种汇总方式
                    mshBody.TextMatrix(j, k) = Switch(strState = "SUM", "合计", strState = "AVG", "平均值", strState = "MAX", "最大值", strState = "MIN", "最小值", strState = "COUNT", "记录数")
                Else '多种汇总方式
                    mshBody.TextMatrix(j, k) = "汇总"
                End If
                mshBody.Row = j: mshBody.Col = k: mshBody.CellAlignment = 4
            End If
        End If
        
        For i = 0 To mshBody.Rows - 1
            mshBody.RowHeight(i) = .行高
        Next

        '其它属性
        mshHead.ScrollBars = flexScrollBarHorizontal
        mshBody.MergeCells = flexMergeRestrictRows
        mshBody.ScrollBars = flexScrollBarBoth
        mshHead.Row = mshHead.FixedRows
        mshBody.Row = 0: mshBody.Col = 0
        
        '表体内容(列宽，列对齐)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.id)
            '设置关联查询超链接样式
            If tmpItem.Relations.count > 0 Then
                For i = 0 To mshBody.Rows - 1
                    '合计列不设置
                    If TypeName(mshBody.RowData(i)) <> "Empty" Then
                        If mshBody.Cell(flexcpForeColor, i, tmpItem.序号) = 0 Then
                            mshBody.Cell(flexcpForeColor, i, tmpItem.序号) = &HFF0001
                        End If
                        mshBody.Cell(flexcpFontUnderline, i, tmpItem.序号) = True
                    End If
                Next
                
                '优化。第一行的单元格存放报表链接关系对象
                For i = 0 To mshBody.Cols - 1
                    If tmpItem.序号 = i Then
                        '列相同时取Relations对象
                        mshBody.Cell(flexcpData, 0, i) = tmpItem.Relations
                        Exit For
                    End If
                Next
            End If
            '该列没有任何数据时隐藏
            blntmp = False
            If tmpItem.分栏 = 1 Then
                For i = mshBody.FixedRows To mshBody.Rows - 1
                    If mshBody.TextMatrix(i, tmpItem.序号) <> "" Then
                        blntmp = True: Exit For
                    End If
                Next
                If blntmp = False Then
                    mshBody.ColHidden(tmpItem.序号) = True
                    mshHead.ColHidden(tmpItem.序号) = True
                    mshBody.ColWidth(tmpItem.序号) = 0
                    mshHead.ColWidth(tmpItem.序号) = 0
                End If
            End If
        Next
        
        '自动重整表格的行高
        For Each objBody In mobjReport.Items
            '当前表格
            If objBody.类型 = Val("4-任意表") And mshBody.Index = objBody.id Then
                Call AdjustRowHight(mshBody.Index)
                Exit For
            End If
        Next
        
        If bytKind <> 2 Then '仅有表体
            mshHead.ZOrder
            mshHead.Visible = True
        End If
        If bytKind <> 1 Then '仅有表头
            mshBody.ZOrder
            mshBody.Visible = True
        End If
    End With
    
    mshBody.Redraw = True
    mshHead.Redraw = True
    
    '补齐其他行的数据
    If UBound(arrRowIDs) + 1 < mshBody.Rows Then
        ReDim Preserve arrRowIDs(UBound(arrRowIDs) + (mshBody.Rows - (UBound(arrRowIDs) + 1)))
    End If
    mcolRowIDs.Add arrRowIDs, "_" & mshBody.Index
    
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub AdjustRowHight(ByVal Index As Integer)
    Dim intBegin As Integer, intEnd As Integer, i As Integer
    Dim objBody As RPTItem, tmpItem As RPTItem
    
    intBegin = -1
    intEnd = -1
    Set objBody = mobjReport.Items("_" & Index)
    If objBody Is Nothing Then Exit Sub
    
    For i = 1 To objBody.SubIDs.count
        Set tmpItem = mobjReport.Items("_" & objBody.SubIDs(i).id)
        If Not tmpItem Is Nothing Then
            If tmpItem.自适应行高 Then
                If intBegin < 0 Then intBegin = tmpItem.序号
                intEnd = tmpItem.序号
            End If
        End If
    Next
    If intBegin >= 0 Then
        '重整行高
        msh(objBody.id).AutoSize intBegin, intEnd
    End If
End Sub

Private Function CheckColProtertys(ByVal var条件字段 As Variant, ByVal str条件关系 As String, ByVal var条件值 As Variant) As Boolean
'功能：根据传入的条件，进行判断是否满足,若为空，则整列执行
    
    Select Case str条件关系
        Case ""
            CheckColProtertys = True
        Case "等于"
            If IsNumeric(var条件字段) Then var条件值 = ValEx(var条件值): var条件字段 = ValEx(var条件字段)
            CheckColProtertys = (var条件字段 = var条件值)
        Case "大于"
            var条件值 = ValEx(var条件值)
            var条件字段 = ValEx(var条件字段)
            CheckColProtertys = (var条件字段 > var条件值)
        Case "小于"
            var条件值 = ValEx(var条件值)
            var条件字段 = ValEx(var条件字段)
            CheckColProtertys = (var条件字段 < var条件值)
        Case "不等于"
            var条件值 = ValEx(var条件值)
            var条件字段 = ValEx(var条件字段)
            CheckColProtertys = (var条件字段 <> var条件值)
        Case "大于等于"
            var条件值 = ValEx(var条件值)
            var条件字段 = ValEx(var条件字段)
            CheckColProtertys = (var条件字段 >= var条件值)
        Case "小于等于"
            var条件值 = ValEx(var条件值)
            var条件字段 = ValEx(var条件字段)
            CheckColProtertys = (var条件字段 <= var条件值)
        Case "左匹配"
            If var条件字段 <> "" And var条件值 <> "" Then
                CheckColProtertys = (var条件字段 Like var条件值 & "*")
            End If
        Case "双向匹配"
            If var条件字段 <> "" And var条件值 <> "" Then
                CheckColProtertys = (var条件字段 Like "*" & var条件值 & "*")
            End If
    End Select
End Function

Private Sub ShowItems()
    Dim i As Integer, lngW As Long, lngH As Long
    Dim objItem As RPTItem, objLoad As Object
    Dim strData As String, strFormat As String, strTmp As String
    Dim strValue As String, objPic As StdPicture
    Dim objFmt As RPTFmt, objFont As StdFont
    Dim lngSize As Long, sngWidth As Single
    Dim lngRec As Long
    Dim objRotate As clsRotateFont
    
    On Error GoTo errH
    blnRefresh = False
    If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "正在组织报表数据,请稍候．．．", , Me

    LockWindowUpdate Me.hwnd
    
    Set mcolRowIDs = New Collection
    For Each objLoad In msh
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In lbl
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In img
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In imgCode
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In lin
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In Shp
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In Chart
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In pic
        If objLoad.Index <> 0 And objLoad.Container Is picPaper(intReport) Then Unload objLoad
    Next
    
    picPaper(intReport).Cls
    intGridCount = 0
    intGridID = 0
    Set objCurGrid = Nothing
    
    Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)
    If objFmt.纸向 = 1 Then
        lngW = objFmt.W
        lngH = objFmt.H
    Else
        lngW = objFmt.H
        lngH = objFmt.W
    End If
    
    '先打载体
    For Each objItem In mobjReport.Items
        '性质为2的为左联接表格,在处理其参照表格的同时处理
        If objItem.格式号 = bytFormat Then
            With objItem
                If .类型 = Val("14-卡片元素") Then
                    Load pic(.id)
                    Set pic(.id).Container = picPaper(intReport)
                    Set objLoad = pic(.id)
                    .内容 = "111"
                    objLoad.Left = .X
                    objLoad.Top = .Y
                    
                    objLoad.Height = IIF(.H > lngH, lngH, .H)
                    objLoad.Width = IIF(.W > lngW, lngW, .W)
                    objLoad.BorderStyle = IIF(.边框, 1, 0)
                    
                    objLoad.ZOrder
                    objLoad.Visible = True
                End If
            End With
        End If
    Next
    
    For Each objItem In mobjReport.Items
        '性质为2的为左联接表格,在处理其参照表格的同时处理
        If objItem.格式号 = bytFormat Then
            With objItem
                Select Case .类型
                    Case 1 '线条
                        Load lin(.id)
                        Set lin(.id).Container = picPaper(intReport)
                        If .父ID <> 0 Then
                            Set lin(.id).Container = pic(.父ID)
                        End If
                        Set objLoad = lin(.id)
                        objLoad.X1 = .X
                        objLoad.X2 = IIF(.X + .W - IIF(.W > 0, Screen.TwipsPerPixelX, 0) > lngW, lngW, .X + .W - IIF(.W > 0, Screen.TwipsPerPixelX, 0))
                        objLoad.Y1 = .Y
                        objLoad.Y2 = IIF(.Y + .H - IIF(.H > 0, Screen.TwipsPerPixelY, 0) > lngH, lngH, .Y + .H - IIF(.H > 0, Screen.TwipsPerPixelY, 0))
                        objLoad.BorderColor = .前景
                        If .粗体 Then objLoad.BorderWidth = 2
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 10 '框线
                        Load Shp(.id)
                        Set Shp(.id).Container = picPaper(intReport)
                        If .父ID <> 0 Then
                            Set Shp(.id).Container = pic(.父ID)
                        End If
                        Set objLoad = Shp(.id)
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderColor = 0
                        If .粗体 Then objLoad.BorderWidth = 2
                        objLoad.Shape = IIF(.边框, ShapeConstants.vbShapeOval, ShapeConstants.vbShapeRectangle)
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 11 '图片
                        Load img(.id)
                        Set img(.id).Container = picPaper(intReport)
                        If .父ID <> 0 Then
                            Set img(.id).Container = pic(.父ID)
                        End If
                        Set objLoad = img(.id)
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        
                        Set objPic = LoadPictureFromPar(Me, .名称)
                        If objPic Is Nothing Then Set objPic = .图片
                        If .自调 And Not objPic Is Nothing Then
                            .W = objPic.Width * (15 / 26.46)
                            .H = objPic.Height * (15 / 26.46)
                        End If
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderStyle = IIF(.边框, 1, 0)
                        
                        '保持比例
                        If Not objPic Is Nothing Then
                            If .粗体 Then
                                Set objLoad.Picture = ScalePicture(picTemp, objPic, objLoad.Width, objLoad.Height)
                            Else
                                Set objLoad.Picture = objPic
                            End If
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 2, 3 '标签,标签绑定图片
                        strValue = .内容
                        
                        '数据指针复位(可能用到多个数据源、多个字段)
                        strData = GetLabelDataName(strValue)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = Split(Split(strData, "|")(i), ".")(0)
                                If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                    If Val(.源行号 & "") <> 0 Then
                                        If mLibDatas("_" & strTmp).DataSet.RecordCount >= Val(.源行号 & "") Then
                                            mLibDatas("_" & strTmp).DataSet.AbsolutePosition = Val(.源行号 & "")
                                        Else
                                            mLibDatas("_" & strTmp).DataSet.MoveFirst
                                        End If
                                    Else
                                        mLibDatas("_" & strTmp).DataSet.MoveFirst
                                    End If
                                End If
                                
                                '先处理数据字段(查询时只取第一个值)
                                strFormat = GetFieldValue(Me, CStr(Split(strData, "|")(i)))
                                If .格式 <> "" Then
                                    On Error Resume Next
                                    strFormat = Format(strFormat, .格式)
                                    If Err.Number <> 0 Then Err.Clear
                                    On Error GoTo errH
                                End If
                                strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strFormat)
                            Next
                            lngRec = mLibDatas("_" & strTmp).DataSet.AbsolutePosition
                        End If
                        
                        '再处理报表变量:[=参数名]、[n>=0]、[日期格式串]、[单位名称]
                        strValue = GetLabelMacro(Me, strValue)
                        
                        If gobjFile.FileExists(strValue) Then
                            '二进制字段当作图形
                            On Error Resume Next
                            Set .图片 = LoadPicture(strValue)
                            If .图片 Is Nothing Then Set .图片 = New StdPicture '以此区分是图片还是文字
                            Kill strValue
                            Err.Clear
                            On Error GoTo errH
                            
                            If .自调 Then
                                .W = .图片.Width * (15 / 26.46)
                                .H = .图片.Height * (15 / 26.46)
                            End If
                            
                            Load img(.id)
                            Set img(.id).Container = picPaper(intReport)
                            Set objLoad = img(.id)
                            objLoad.BorderStyle = IIF(.边框, 1, 0)
                            
                            '保持比例
                            If .粗体 Then
                                Set objLoad.Picture = ScalePicture(picTemp, .图片, objLoad.Width, objLoad.Height)
                            Else
                                Set objLoad.Picture = .图片
                            End If
                        Else
                            Set .图片 = Nothing '以此区分是图片还是文字
                            
                            If .自调 Then Call ItemAutoSize(objItem, strValue, picBack)
                            
                            Load lbl(.id)
                            Set lbl(.id).Container = picPaper(intReport)
                            If .父ID <> 0 Then
                                Set lbl(.id).Container = pic(.父ID)
                            End If
                            Set objLoad = lbl(.id)
                            
                            objLoad.FontName = .字体
                            objLoad.FontSize = .字号
                            objLoad.FontBold = .粗体
                            objLoad.FontItalic = .斜体
                            objLoad.FontUnderline = .下线
                            
                            objLoad.Alignment = IIF(.对齐 = 2, 1, IIF(.对齐 = 1, 2, 0))
                            objLoad.BorderStyle = IIF(.边框, 1, 0)
                            objLoad.ForeColor = .前景
                            objLoad.BackColor = .背景
                            objLoad.Caption = strValue
                            '设置超链接字样
                            If objItem.Relations.count > 0 Then
                                objLoad.ForeColor = &HFF0001
                                objLoad.FontUnderline = True
                                objLoad.Tag = lngRec
                            End If
                        End If
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        
                        If .水平反转 Then
                            Set objRotate = New clsRotateFont
                            
                            Load picRotate(.id)
                            With picRotate(.id)
                                Set .Font = objLoad.Font
                                If objItem.父ID = 0 Then
                                    Set .Container = picPaper(intReport)
                                Else
                                    Set .Container = objLoad.Container      '卡片内的标签
                                End If
                                .AutoRedraw = True
                                .Left = objLoad.Left
                                .Top = objLoad.Top
                                .Width = objLoad.Width
                                .Height = objLoad.Height
                                .BackColor = objLoad.BackColor
                                If objItem.边框 Then
                                    picRotate(objItem.id).Line (0, 0)-(objLoad.Width - 15, objLoad.Height - 15), , B
                                End If
                                .ForeColor = objLoad.ForeColor
                                .ZOrder
                                .Visible = True
                            End With
                            
                            Set objRotate.LogFont = picRotate(.id).Font
                            objRotate.OutputReverse picRotate(.id), lbl(.id).Caption, .对齐
                        Else
                            objLoad.ZOrder
                            objLoad.Visible = True
                        End If
                    Case 4 '任意表格(含类型为6的子项)
                        If objItem.性质 = 0 Then
                            If .父ID = 0 Then
                                '卡片中的表格不算独立表
                                intGridCount = intGridCount + 1
                                intGridID = objItem.id
                            End If
                        End If
                        Call ShowFreeGrid(objItem)
                    Case 5 '分类表格(含类型为7,8,9的子项)
                        If objItem.性质 = 0 Then
                            intGridCount = intGridCount + 1
                            intGridID = objItem.id
                            Call ShowStatGrid(objItem)
                        End If
                    Case 12 '图表@@@
                        Load Chart(.id)
                        Set Chart(.id).Container = picPaper(intReport)
                        If .父ID <> 0 Then
                            Set Chart(.id).Container = pic(.父ID)
                        End If
                        Set objLoad = Chart(.id)
                        
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                                                                
                        strTmp = GetChartFileFromPar(Me, .名称)
                        If strTmp <> "" Then
                            Call objLoad.Load(strTmp)
                            objLoad.Height = IIF(.H > lngH, lngH, .H)
                            objLoad.Width = IIF(.W > lngW, lngW, .W)
                        Else
                            If objItem.内容 <> "" Then
                                Call GetChartDataName(objItem.内容, , , , strTmp)
                            End If
                            If strTmp <> "" Then
                                Call SetChartStyleAndData(objLoad, objItem, mLibDatas("_" & strTmp).DataSet)
                            Else
                                Call SetChartStyleAndData(objLoad, objItem, , , , True)
                            End If
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 13 '条码
                        Load imgCode(.id)
                        Set imgCode(.id).Container = picPaper(intReport)
                        If .父ID <> 0 Then
                            Set imgCode(.id).Container = pic(.父ID)
                        End If
                        Set objLoad = imgCode(.id)
                        
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderStyle = 0
                        
                        '获取条码内容
                        strValue = .内容
                        
                        '数据指针复位(可能用到多个数据源、多个字段)
                        strData = GetLabelDataName(strValue)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = Split(Split(strData, "|")(i), ".")(0)
                                If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                    mLibDatas("_" & strTmp).DataSet.MoveFirst
                                End If
                            Next
                        End If
                        
                        '先处理数据字段(查询时只取第一个值)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = GetFieldValue(Me, CStr(Split(strData, "|")(i)))
                                If .格式 <> "" Then
                                    On Error Resume Next
                                    strTmp = Format(strTmp, .格式)
                                    If Err.Number <> 0 Then Err.Clear
                                    On Error GoTo errH
                                End If
                                strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                            Next
                        End If
                        
                        '再处理报表变量:[=参数名]、[n>=0]、[日期格式串]、[单位名称]
                        strValue = GetLabelMacro(Me, strValue)
                        '[页号]、[页数]预览时才有值
                        strValue = Replace(strValue, "[页号]", "")
                        strValue = Replace(strValue, "[页数]", "")
                        
                        Set objPic = Nothing
                        If strValue <> "" Then
                            Unload frmFlash '强制初始Picture，不然切换绘制有问题
                            If .序号 = 1 Then
                                Set objPic = DrawBarCode128(frmFlash.picTemp, 3, strValue, Mid(.表头, 1, 1) = "1")
                            ElseIf .序号 = 2 Then
                                Set objPic = DrawBarCode39(frmFlash.picTemp, 3, strValue, Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                            ElseIf .序号 = 3 Then
                                Set objPic = DrawBarCode128Auto(frmFlash.picTemp, strValue, sngWidth, .行高, Mid(.表头, 1, 1) = "1")
                            ElseIf .序号 = 10 Then
                                Set objPic = DrawBarCode2D(strValue, frmFlash.picTemp, lngSize)
                            End If
                            If Val(Mid(.表头, 3, 1)) <> 0 Then
                                Set objPic = PictureSpin(objPic, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                            End If
                        End If
                        Set objLoad.Picture = objPic
                        
                        If .序号 = 3 Then
                            '128码自动调整宽度
                            If Val(Mid(.表头, 3, 1)) = 0 Then
                                .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                                objLoad.Width = .W
                            Else
                                .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                                objLoad.Height = .H
                            End If
                        ElseIf .序号 = 10 And .自调 Then
                            '二维条码缺省自动调整大小
                            objLoad.Width = lngSize
                            objLoad.Height = lngSize
                            .W = lngSize: .H = lngSize
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                End Select
            End With
        End If
    Next
    
    '处理标签元素为指定表格单元格值的设置
    For Each objItem In mobjReport.Items
        If objItem.类型 = 4 Or objItem.类型 = 5 Then
            Call mdlPublic.SetCellValue(Val("0-初预览"), Me, objItem)
        End If
    Next
    
    scrVsc.Visible = Not (intGridCount = 1 And Not mobjReport.票据)
    scrHsc.Visible = Not (intGridCount = 1 And Not mobjReport.票据)
    picShadow.Visible = Not (intGridCount = 1 And Not mobjReport.票据)
    
    '设置附加体缺省对齐
    Call SetGridAlign
        
    mobjReport.intGridCount = intGridCount
    mobjReport.intGridID = intGridID
        
    Call Form_Resize
    
    ShowFlash
    blnRefresh = True
    LockWindowUpdate 0
    Exit Sub
errH:
    ShowFlash
    LockWindowUpdate 0
    If ErrCenter() = 1 Then
        If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "正在组织报表数据,请稍候．．．", , Me
        LockWindowUpdate Me.hwnd
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetFixHeight(objGrid As Object) As Long
'功能：获取指定表格的固定部份高度
    Dim i As Integer, lngH As Long
    
    For i = 0 To objGrid.FixedRows - 1
        lngH = lngH + objGrid.RowHeight(i)
    Next
    GetFixHeight = lngH
End Function

Private Sub GetGridCurSize(ByVal intID As Integer, ByRef X As Long, ByRef Y As Long, _
    ByRef W As Long, ByRef H As Long, Optional ByRef Bottom As Long)
'功能：获取指定表格的当前显示整体尺寸(指加上左联接或附加表格或分栏)
'返回：X,Y,W,H；Bottom(设计时)
    Dim objItem As RPTItem, tmpItem As RPTItem, lngCurH As Long, lngBottom As Long
    
    X = msh(intID).Left
    W = msh(intID).Width
    
    If Val(msh(intID).Tag) = 0 Then
        Y = msh(intID).Top
        lngCurH = msh(intID).Height
    Else
        Y = msh(CInt(msh(intID).Tag)).Top
        lngCurH = msh(CInt(msh(intID).Tag)).Height
    End If
        
    lngBottom = mobjReport.Items("_" & intID).Y + mobjReport.Items("_" & intID).H
    
    Set objItem = mobjReport.Items("_" & intID)
    W = W * objItem.分栏
    
    '汇总表也可能有附加表
    For Each tmpItem In mobjReport.Items '加上附加表高度
        If tmpItem.格式号 = bytFormat And tmpItem.类型 = 4 _
            And tmpItem.性质 = 1 And tmpItem.参照 = objItem.名称 Then
            lngCurH = lngCurH + msh(CInt(msh(tmpItem.id).Tag)).Height
            lngBottom = lngBottom + tmpItem.H
        End If
    Next
    H = lngCurH
    Bottom = lngBottom
End Sub

Private Function GetDependID(strName As String) As Integer
'功能：根据参照名称,获取其索引.
    Dim objItem As RPTItem
    
    For Each objItem In mobjReport.Items
        If objItem.格式号 = bytFormat And objItem.名称 = strName _
            And (objItem.类型 = 4 Or objItem.类型 = 5) And objItem.性质 = 0 Then
            GetDependID = objItem.id: Exit Function
        End If
    Next
End Function

Private Sub SetPlace()
'功能：根据报表内容，设置表格、标签的相对位置
    Dim objItem As RPTItem, tmpItem As RPTItem
    Dim lngDesignH As Long, lngShowH As Long, lngAppH As Long
    Dim lngCurH As Long, lngCurTop As Long, bytKind As Byte
    Dim strAppGrid As String, lngFixH As Long
    Dim strGridScale As String, i As Integer
    Dim intCurID As Integer, sngScale As Single
    Dim lngTX As Long, lngTY As Long, lngTW As Long, lngTH As Long
    Dim lngBottom As Long
    
    On Error GoTo errH
    
    If mobjReport Is Nothing Then Exit Sub
    If Not mobjReport.blnLoad Then Exit Sub
    
    '调整表格适合窗体大小
    If intGridCount = 1 And Not mobjReport.票据 Then
        '当只有一个表格体时(包含左联接、附加表格)：
        '1:表Top、Left使用绝对位置
        '2.表Width使用相对表Left的相对位置
        '3.表下标签总高度不能超过表高的的一半,表高取除开表下标签位置的相对高度
        Set objItem = mobjReport.Items("_" & intGridID)
        
        '计算表格设计和当前高度(注意任意表可能没有表头或表体)
        lngDesignH = objItem.H '设计高度
        If Val(msh(intGridID).Tag) > 0 Then
            '表头实际高度总是整个高度
            lngShowH = msh(CInt(msh(intGridID).Tag)).Height
        Else
            lngShowH = msh(intGridID).Height
        End If
        
        For Each tmpItem In mobjReport.Items '加上附加表高度
            If tmpItem.格式号 = bytFormat And tmpItem.类型 = 4 _
                And tmpItem.性质 = 1 And tmpItem.参照 = objItem.名称 Then
                lngDesignH = lngDesignH + tmpItem.H
                lngShowH = lngShowH + msh(CInt(msh(tmpItem.id).Tag)).Height
                strAppGrid = strAppGrid & "," & tmpItem.id '附加表格们
            End If
        Next
        strAppGrid = Mid(strAppGrid, 2)
        
        '计算表下标签占用总高度
        For Each tmpItem In mobjReport.Items
            If tmpItem.格式号 = bytFormat And tmpItem.类型 = 2 _
                And tmpItem.图片 Is Nothing And tmpItem.Y >= objItem.Y + lngDesignH Then
                If tmpItem.Y + lbl(tmpItem.id).Height > lngAppH Then
                    lngAppH = tmpItem.Y + lbl(tmpItem.id).Height '用真实高度比较
                End If
            End If
        Next
        
        If lngAppH > 0 Then lngAppH = lngAppH - (objItem.Y + lngDesignH)
        
        '下边高度不能超过表高的一半
        If lngAppH > lngShowH / 2 Then lngAppH = lngShowH / 2
        
        msh(intGridID).Width = (picPaper(intReport).ScaleWidth - msh(intGridID).Left * 2) / objItem.分栏

        '整个表体应该占的高度
        If Val(msh(intGridID).Tag) = 0 Then
            lngCurH = picPaper(intReport).ScaleHeight - msh(intGridID).Top - (lngAppH + 200)
        Else
            msh(CInt(msh(intGridID).Tag)).Width = msh(intGridID).Width
            lngCurH = picPaper(intReport).ScaleHeight - msh(CInt(msh(intGridID).Tag)).Top - (lngAppH + 200)
        End If
        
        If strAppGrid = "" Then '没有附加表格时
            If objItem.类型 = 5 Then
                msh(intGridID).Height = lngCurH
            Else
                bytKind = GetGridStyle(mobjReport, intGridID)
                If bytKind = 2 Then
                    msh(intGridID).Height = lngCurH
                Else
                    lngFixH = GetFixHeight(msh(CInt(msh(intGridID).Tag))) '纯表头高度
                    If lngCurH < lngFixH + 300 Then lngCurH = lngFixH + 300 '至少要保证可以显示表头(加滚动条)
                    msh(CInt(msh(intGridID).Tag)).Height = lngCurH
                    msh(intGridID).Height = lngCurH - lngFixH
                End If
            End If
        Else
            '有附加表格时
            '计算各个附加表格占总高度的比例
            strGridScale = "|" & objItem.id & "," & objItem.H / lngDesignH
            For i = 0 To UBound(Split(strAppGrid, ","))
                Set tmpItem = mobjReport.Items("_" & Split(strAppGrid, ",")(i))
                strGridScale = strGridScale & "|" & tmpItem.id & "," & tmpItem.H / lngDesignH
            Next
            strGridScale = Mid(strGridScale, 2) '"表ID,比例|表ID,比例..."
            lngCurTop = objItem.Y
            For i = 0 To UBound(Split(strGridScale, "|"))
                intCurID = CInt(Split(Split(strGridScale, "|")(i), ",")(0))
                sngScale = CSng(Split(Split(strGridScale, "|")(i), ",")(1))
                
                If i > 0 Then
                    msh(intCurID).Width = msh(intGridID).Width
                    msh(CInt(msh(intCurID).Tag)).Width = msh(intCurID).Width
                End If
                
                bytKind = GetGridStyle(mobjReport, intCurID)
                
                If Val(msh(intCurID).Tag) = 0 Then '为分类表,也可能带附加,但不可能为附加
                    msh(intCurID).Height = lngCurH * sngScale
                    lngCurTop = lngCurTop + msh(intCurID).Height
                Else
                    lngFixH = GetFixHeight(msh(CInt(msh(intCurID).Tag)))
                    If lngCurH * sngScale < lngFixH + 300 Then '至少要保证可以显示表头(加滚动条)
                        msh(CInt(msh(intCurID).Tag)).Height = lngFixH + 300
                    Else
                        msh(CInt(msh(intCurID).Tag)).Height = lngCurH * sngScale
                    End If
                    msh(CInt(msh(intCurID).Tag)).Top = lngCurTop
                    lngCurTop = lngCurTop + msh(CInt(msh(intCurID).Tag)).Height
                    
                    bytKind = GetGridStyle(mobjReport, intCurID)
                    If bytKind = 2 Then
                        msh(intCurID).Top = msh(CInt(msh(intCurID).Tag)).Top
                        msh(intCurID).Height = msh(CInt(msh(intCurID).Tag)).Height
                    Else
                        msh(intCurID).Top = msh(CInt(msh(intCurID).Tag)).Top + lngFixH
                        msh(intCurID).Height = msh(CInt(msh(intCurID).Tag)).Height - lngFixH
                    End If
                End If
            Next
        End If
    End If
    
    For Each tmpItem In mobjReport.Items
        If tmpItem.格式号 = bytFormat And tmpItem.类型 = Val("4-任意表") And tmpItem.分栏 > 1 _
            And tmpItem.性质 = 0 And tmpItem.参照 = "" Then
            '分栏标志
            For i = 2 To tmpItem.分栏
                With msh(tmpItem.id)
                    DrawCell picPaper(intReport), "数据分栏位置", tmpItem.X + ((i - 1) * .Width), tmpItem.Y, .Width, _
                        msh(CInt(.Tag)).Height - 15, , , .GridColor, .ForeColor, .BackColor, .Font, , 1, 1
                End With
            Next
        End If
    Next
    
    '调整标签适合其照参表格位置(多表格时)
    For Each tmpItem In mobjReport.Items
        If tmpItem.格式号 = bytFormat And tmpItem.类型 = 2 And tmpItem.图片 Is Nothing Then
            '定义的左右靠齐：不管多少表格都处理
            If tmpItem.性质 <> 0 And tmpItem.参照 <> "" Then
                GetGridCurSize GetDependID(tmpItem.参照), lngTX, lngTY, lngTW, lngTH
                Select Case tmpItem.性质
                    Case 11, 21 '靠左
                        lbl(tmpItem.id).Left = lngTX
                    Case 12, 22 '靠中
                        lbl(tmpItem.id).Left = lngTX + (lngTW - lbl(tmpItem.id).Width) / 2
                    Case 13, 23 '靠右
                        lbl(tmpItem.id).Left = lngTX + lngTW - lbl(tmpItem.id).Width
                End Select
            End If
            '表下的自动靠齐：只有一个表整体时才处理(所有标签,包含表附项)
            If intGridCount = 1 Then
                GetGridCurSize intGridID, lngTX, lngTY, lngTW, lngTH, lngBottom
                If tmpItem.Y >= lngBottom Then
                    lbl(tmpItem.id).Top = lngTY + lngTH + (tmpItem.Y - lngBottom)
                End If
            End If
        End If
    Next
    Exit Sub
errH:
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GridHaveApp(intID As Integer) As Boolean
'功能：判断一个表格是否具有附加表格
    Dim tmpItem As RPTItem, strName As String
    
    strName = mobjReport.Items("_" & intID).名称
    For Each tmpItem In mobjReport.Items
        If tmpItem.格式号 = bytFormat And tmpItem.类型 = 4 And tmpItem.性质 = 1 And tmpItem.参照 = strName Then
            GridHaveApp = True: Exit Function
        End If
    Next
End Function

Private Function GetGridDesignWidth(objItem As RPTItem) As Long
'功能：获取表格加上其左联接表格在设计时的总宽度
    Dim lngW As Long, tmpItem As RPTItem
    
    lngW = objItem.W
    For Each tmpItem In mobjReport.Items
        If tmpItem.格式号 = bytFormat And tmpItem.类型 = 5 _
            And tmpItem.性质 = 2 And tmpItem.参照 = objItem.名称 Then
            lngW = lngW + tmpItem.W
        End If
    Next
    GetGridDesignWidth = lngW
End Function

Private Function GetPreAppGrid(intID As Integer, arrGrids As Variant) As Long
'功能：获取当前附加表格的前一个表格(可能为附加表格或独立表格)
'参数：arrGrids=按XY先后顺序存放的表格索引数组
'说明：1.对象集中的数据必须保证附加表格按Y先后顺序存放
'      2.当当前输出表格为附加表格时,其参照表格一定已经输出
    Dim objItem As RPTItem, tmpItem As RPTItem, i As Integer
    
    Set objItem = mobjReport.Items("_" & intID)
    For i = 0 To UBound(arrGrids)
        Set tmpItem = mobjReport.Items("_" & arrGrids(i))
        If tmpItem.格式号 = bytFormat And _
            ((tmpItem.类型 = 4 And tmpItem.性质 = 1 And tmpItem.参照 = objItem.参照) Or _
            (tmpItem.性质 = 0 And objItem.参照 = tmpItem.名称)) Then
            If tmpItem.id <> intID Then
                GetPreAppGrid = tmpItem.id
            Else '当处理到当前附加表格时,上一表格即为所得
                Exit Function
            End If
        End If
    Next
End Function

Private Function GetGridDesignHeight(intID As Integer) As Long
'功能：获取表格的设计时高度(包含所有附加表格)
'参数：整个附加体中的任何一个表格索引
    Dim objItem As RPTItem, tmpItem As RPTItem
    Dim lngH As Long
    
    Set objItem = mobjReport.Items("_" & intID)
    If objItem.性质 = 1 And objItem.参照 <> "" Then
        Set objItem = mobjReport.Items("_" & GetDependID(objItem.参照))
    End If
    
    lngH = objItem.H
    For Each tmpItem In mobjReport.Items
        If tmpItem.格式号 = bytFormat And tmpItem.类型 = 4 _
            And tmpItem.性质 = 1 And tmpItem.参照 = objItem.名称 Then
            lngH = lngH + tmpItem.H
        End If
    Next
    GetGridDesignHeight = lngH
End Function

Private Function GetGridPageCol(objItem As RPTItem) As Integer
'功能：返回清册表中的用于自动换页的列号
'参数：objItem=表格对象
'返回：-1=没有
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    GetGridPageCol = -1
    If objItem.类型 <> 4 Then Exit Function
    
    For Each tmpID In objItem.SubIDs
        Set tmpItem = mobjReport.Items("_" & tmpID.id)
        If tmpItem.边框 Then
            GetGridPageCol = tmpItem.序号
            Exit For
        End If
    Next
End Function

Private Sub AddPrintPage(ByVal intPage As Integer, ByVal objBody As Object, ByVal colCard As Collection _
    , ByVal lngPageBeginRow As Long, ByVal lngPageEndRow As Long _
    , ByVal lngW As Long, ByVal lngL As Long)

    '动态定义对象数组
    If intPage > 0 Then
        ReDim Preserve marrPageCard(intPage) As PageCards
    Else
        ReDim marrPageCard(intPage) As PageCards
    End If
    Set marrPageCard(intPage) = New PageCards
    
    '加入新的打印页对象
    marrPageCard(intPage).Add objBody.Index, objBody.Left, objBody.Top, objBody.Width _
        , objBody.Height, lngPageBeginRow, lngPageEndRow, lngW, lngL, colCard, "_" & objBody.Index
End Sub

Private Function CalcCellPage() As Boolean
'功能：计算单元格与页的对应关系
'参数：mobjreport=报表对象
'      marrPage=打印页集
'返回：是否可以进行打印或预览(如固定行列尺寸比较表格尺寸还大)
'说明：如果运行该函数之后isArray(marrPage)=False,则表明没有表格输出
    Dim objBody As Control, objPageCell As PageCell, arrPage As Variant '当前处理表格对象
    Dim lngFixW As Long, lngFixH As Long '当前表格固定行列尺寸
    Dim lngRowB As Long, lngRowE As Long
    Dim lngColB As Long, lngColE As Long '起止行列
    Dim lngBodyW As Long, lngBodyH As Long '除开固定行列后可用宽高
    Dim lngCurW As Long, lngCurH As Long '当前页中表格计算累计到的宽高
    Dim lngOutX As Long, lngOutY As Long '当前布中表格输出的实际位置(主要用于附加表格)
    Dim bytKind As Byte, intPage As Integer  '当前处理到的页(0-N)
    Dim i As Long, j As Long, k As Long, strTmp As String
    Dim objItem As RPTItem, blnHaveApp As Boolean, blnHorPage As Boolean
    Dim blnApp As Boolean, lngMinH As Long, arrGrids As Variant
    Dim lngPreID As Long, intDepend As Integer, lngDesignH As Long
    Dim tmpPageCell As PageCell
    Dim lngL As Long, lngW As Long, lngC As Long, lngZ As Long
    Dim lngTop As Long, lngLeft As Long
    Dim lngCount As Long, lngRowsHeight As Long
    Dim blnData As Boolean, tmpSubID As RelatID
    Dim Y As Long, X As Long, Z As Long
    
    '根据列内容自动换页相关变量
    Dim strCurText As String, blnNewPage As Boolean, lngPageCol As Long
    Dim lngBaseRows As Long, lngVRowE As Long
    Dim colCardRow As New Collection  '记录卡片内表格最小显示行数
    Dim lngLastID As Long, lngRow As Long
    Dim lngRowCount As Long, colCard As New Collection
    Dim blnRePage As Boolean, blnPage As Boolean
    Dim arrPageTmp As Variant, arrTmp As Variant
    Dim intGridID As Integer
    Dim vsfTmp As VSFlexGrid
    
    '将表格按X,Y先后次序排序
    arrGrids = Array()
    For Each objBody In msh
        If objBody.Index <> 0 And (objBody.Container Is picPaper(intReport) Or objBody.Container.name = "pic") _
            And Left(objBody.Tag, 2) <> "H_" Then
            ReDim Preserve arrGrids(UBound(arrGrids) + 1)
            arrGrids(UBound(arrGrids)) = objBody.Left & "," & objBody.Top & "," & objBody.Index
        End If
    Next
    For i = 0 To UBound(arrGrids) - 1
        For j = i To UBound(arrGrids)
            If CLng(Split(arrGrids(j), ",")(0)) < CLng(Split(arrGrids(i), ",")(0)) Then
                strTmp = arrGrids(i): arrGrids(i) = arrGrids(j): arrGrids(j) = strTmp
            End If
        Next
    Next
    For i = 0 To UBound(arrGrids) - 1
        For j = i To UBound(arrGrids)
            If CLng(Split(arrGrids(j), ",")(1)) < CLng(Split(arrGrids(i), ",")(1)) Then
                strTmp = arrGrids(i): arrGrids(i) = arrGrids(j): arrGrids(j) = strTmp
            End If
        Next
    Next
    For i = 0 To UBound(arrGrids)
        arrGrids(i) = CInt(Split(arrGrids(i), ",")(2))
    Next
    
    arrPage = Empty
    marrPage = Empty
    marrPageCard = Empty
    
    For k = 0 To UBound(arrGrids)
        '逐个表格计算
        Set objBody = msh(arrGrids(k))
        '强制处理最小行高
        For i = 0 To objBody.Rows - 1
            If objBody.RowHeight(i) < objBody.RowHeightMin Then
                objBody.RowHeight(i) = objBody.RowHeightMin
            End If
        Next
        
        strTmp = ""
        lngLeft = 0: lngTop = 0
        If objBody.Container.name = "pic" Then
            If objBody.Container.Container Is picPaper(intReport) Then
                lngLeft = mobjReport.Items("_" & objBody.Container.Index).X
                lngTop = mobjReport.Items("_" & objBody.Container.Index).Y
            End If
        End If
        Set objItem = mobjReport.Items("_" & objBody.Index)
        blnApp = (objItem.类型 = Val("4-任意表") And objItem.性质 = Val("1-附加表格") And objItem.参照 <> "") '是否附加表格
        
        '求该表格固定行宽及固定行高
        lngFixW = 0: lngFixH = 0
        lngDesignH = GetGridDesignHeight(objItem.id) '包含附加表格的高度
        
        If objItem.类型 = Val("5-汇总表") Then
            For i = 0 To objBody.FixedCols - 1
                lngFixW = lngFixW + objBody.ColWidth(i)
            Next
            For i = 0 To objBody.FixedRows - 1
                lngFixH = lngFixH + objBody.RowHeight(i)
            Next
            '除去固定行列之后一页可用的宽度和高度(不算分栏)
            lngBodyW = GetGridDesignWidth(objItem) - lngFixW
            lngBodyH = lngDesignH - lngFixH
        Else
            bytKind = GetGridStyle(mobjReport, objBody.Index)
            For i = 0 To msh(CInt(objBody.Tag)).FixedRows - 1
                lngFixH = lngFixH + msh(objBody.Tag).RowHeight(i)
            Next
            Select Case bytKind
            Case Val("0-表头表体")
                lngBodyH = lngDesignH - lngFixH
            Case Val("1-表头")
                lngBodyH = 0
            Case Val("2-表体")
                lngBodyH = lngDesignH
                lngFixH = 0
            End Select
            lngBodyW = objItem.W
        End If
        
        '调整为整除像素的缇值
        lngBodyW = Round(lngBodyW / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
        
        If objItem.类型 = Val("4-任意表") Then blnHaveApp = GridHaveApp(objItem.id)
        
        lngPageCol = GetGridPageCol(objItem) '任意表才有,没有为-1
        lngRowB = objBody.FixedRows
        lngColB = objBody.FixedCols
        lngRowE = lngRowB - 1
        lngColE = lngColB - 1
        
        '当前表格输出起始页号
        If blnApp Then
            '参照的表格ID
            intDepend = GetDependID(objItem.参照)
            '上一个输出的附加表格ID
            '因为该表格为附加表格,其上一个表格一定已经输出
            lngPreID = GetPreAppGrid(objItem.id, arrGrids)
            intPage = -1
            If objItem.父ID <> 0 Then
                arrTmp = arrPageTmp
            Else
                arrTmp = arrPage
            End If
            For i = 0 To UBound(arrTmp)
                For Each objPageCell In arrTmp(i)
                    If objPageCell.id = lngPreID Then
                        '最后一行(及第一列)输出页为上个表格输出结束页
                        '(因为汇总表可能横向分页,但这些分页不输出附加表格)
                        If objPageCell.RowE >= msh(objPageCell.id).Rows - 1 _
                            And objPageCell.ColB = msh(objPageCell.id).FixedCols Then
                            '判断剩余高度是否够输出(最小高度为表头加一行)
                            Select Case bytKind
                                Case 0 '输出整体
                                    lngMinH = lngFixH + objItem.行高
                                Case 1 '仅输出表头
                                    lngMinH = lngFixH
                                Case 2 '仅输出表体
                                    lngMinH = objItem.行高
                            End Select
                            If lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y) >= lngMinH Then
                                lngOutX = objPageCell.X + lngLeft
                                lngOutY = objPageCell.Y + objPageCell.H + lngTop
                                Select Case bytKind
                                    Case 0 '输出整体
                                        lngBodyH = lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y) - lngFixH
                                    Case 1 '仅输出表头
                                        lngBodyH = 0
                                    Case 2 '仅输出表体
                                        lngBodyH = lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y)
                                End Select
                                intPage = i
                            Else
                                '在新页中开始输出,在整个框架内输出
                                lngOutX = mobjReport.Items("_" & intDepend).X + lngLeft
                                lngOutY = mobjReport.Items("_" & intDepend).Y + lngTop
                                Select Case bytKind
                                    Case 0 '输出整体
                                        lngBodyH = lngDesignH - lngFixH
                                    Case 1 '仅输出表头
                                        lngBodyH = 0
                                    Case 2 '仅输出表体
                                        lngBodyH = lngDesignH
                                End Select
                                intPage = i + 1
                                
                                '附加表格跳过其参照表格的横向分页
                                For j = intPage To UBound(arrTmp)
                                    For Each tmpPageCell In arrTmp(j)
                                        If tmpPageCell.id = intDepend Then
                                            If tmpPageCell.ColB <> msh(intDepend).FixedCols Then
                                                intPage = intPage + 1
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                            Exit For
                        End If
                    End If
                Next
                If intPage <> -1 Then Exit For
            Next
            If intPage = -1 Then intPage = 0
        Else
            lngOutX = objItem.X + lngLeft
            lngOutY = objItem.Y + lngTop
            intPage = 0
        End If
        
        '页间循环(每个表格在多页中计算)
        Do
            '页内循环(两个DO)
            
            '计算当前页行范围
            lngCurH = 0
            blnNewPage = False
            Do
                If lngPageCol <> -1 Then
                    If lngRowE + 1 = lngRowB Then
                        '每页第一次为lngRowE=lngRowB-1,不用比较,且该行在该页必定要打印
                        strCurText = objBody.TextMatrix(lngRowE + 1, lngPageCol)
                    ElseIf lngRowE + 1 > lngRowB Then
                        If strCurText <> objBody.TextMatrix(lngRowE + 1, lngPageCol) Then
                            blnNewPage = True
                        End If
                    End If
                End If
                If Not blnNewPage Then
                    lngCurH = lngCurH + objBody.RowHeight(lngRowE + 1)
                    If lngCurH <= lngBodyH Then
                        lngRowE = lngRowE + 1   '根据实现行高计算出每页可输出的行数，以及表格高度
                        If lngPageCol <> -1 Then
                            strCurText = objBody.TextMatrix(lngRowE, lngPageCol)
                        End If
                    End If
                End If
            Loop Until (lngCurH > lngBodyH) Or (lngRowE = objBody.Rows - 1) Or blnNewPage
            
            '取实际高度
            If lngCurH > lngBodyH Then lngCurH = lngCurH - objBody.RowHeight(lngRowE + 1)
            
            '取当前页能容纳的实际行数
            lngRowsHeight = 0
            lngBaseRows = 0
            For i = lngRowB To objBody.Rows - 1
                lngRowsHeight = lngRowsHeight + objBody.RowHeight(i)
                If lngBodyH < lngRowsHeight Then
                    Exit For
                Else
                    lngBaseRows = lngBaseRows + 1
                End If
            Next
            If lngBodyH > lngRowsHeight Then
                '设计表格高与实际行数高的差计算空行行数
                lngBaseRows = lngBaseRows + (lngBodyH - lngRowsHeight) \ objItem.行高
            End If
            
            '不足打印一行,则强行打印一行
            If lngRowE < lngRowB Then lngRowE = lngRowB
            
            '计算分栏和票据时的输出行数：表格所有行高相同
            lngVRowE = 0 '分栏或者票据输出时的虚拟结束行
            If objItem.分栏 > 1 Then
                '求出真实换页尾行(前面只是可能超出高度了)
                If lngPageCol <> -1 Then
                    strCurText = objBody.TextMatrix(lngRowE, lngPageCol)
                    For i = lngRowE + 1 To objBody.Rows - 1
                        If i - lngRowB + 1 > lngBaseRows * objItem.分栏 Then
                            lngRowE = i - 1: Exit For
                        ElseIf strCurText <> objBody.TextMatrix(i, lngPageCol) Then
                            lngRowE = i - 1: Exit For
                        Else
                            lngRowE = i
                        End If
                        strCurText = objBody.TextMatrix(i, lngPageCol)
                    Next
                Else
                    For i = lngRowE + 1 To objBody.Rows - 1
                        If i - lngRowB + 1 > lngBaseRows * objItem.分栏 Then
                            lngRowE = i - 1: Exit For
                        Else
                            lngRowE = i
                        End If
                    Next
                End If
                '求出虚拟换页尾行(因分栏补填的空白行)
                If mobjReport.票据 Then
                    '分栏且是票据时，按设计补空行
                    lngVRowE = lngRowE + (lngBaseRows * objItem.分栏 - (lngRowE - lngRowB + 1))
                Else
                    '分栏不是票据时，按单栏实际输出行数分栏补空行
                    If lngRowE - lngRowB + 1 <= lngBaseRows Then
                        lngVRowE = lngRowE + (lngRowE - lngRowB + 1) * (objItem.分栏 - 1)
                    Else
                        lngVRowE = lngRowE + (lngBaseRows * objItem.分栏 - (lngRowE - lngRowB + 1))
                    End If
                End If
            Else
                '没有分栏时，票据需要补空行
                '不管是不是数据变化强行换页
                If mobjReport.票据 Then
                    lngVRowE = lngRowE + (lngBaseRows - (lngRowE - lngRowB + 1))
                End If
            End If
            If lngVRowE = lngRowE Then lngVRowE = 0
            
            '计算列范围(横向分页是多页)
            Do
                '计算当前页列范围
                lngCurW = 0
                Do
                    lngCurW = lngCurW + objBody.ColWidth(lngColE + 1)
                    If lngCurW <= lngBodyW Then lngColE = lngColE + 1
                Loop Until lngCurW > lngBodyW Or lngColE = objBody.Cols - 1
                
                '取真实宽度
                If lngCurW > lngBodyW Then lngCurW = lngCurW - objBody.ColWidth(lngColE + 1)
                
                '不足打印一列,则强行打印一列
                If lngColE < lngColB Then lngColE = lngColB
                
                If objItem.父ID = 0 Then
                    '卡片外的表格要分页
                    blnPage = True
                Else
                    '卡片内部的表格卡片有数据源时不分页
                    If mobjReport.Items("_" & objItem.父ID).数据源 = "" Then
                        blnPage = True
                    Else
                        blnPage = False
                    End If
                End If
                
                '新的一页初始
                If blnPage Then
                    If Not IsArray(arrPage) Then
                        ReDim arrPage(intPage) As PageCells  '第一次初始页
                        Set arrPage(intPage) = New PageCells
                    ElseIf intPage > UBound(arrPage) Then
                        '如果该页已被其它表格占用,则不用再初始
                        ReDim Preserve arrPage(intPage) As PageCells
                        Set arrPage(intPage) = New PageCells
                    End If
                Else
                    If intPage = 0 Then
                        If Not IsArray(arrPageTmp) Then
                            ReDim arrPageTmp(intPage) As PageCells  '第一次初始页
                            Set arrPageTmp(intPage) = New PageCells
                        ElseIf intPage > UBound(arrPageTmp) Then
                            '如果该页已被其它表格占用,则不用再初始
                            ReDim Preserve arrPageTmp(intPage) As PageCells
                            Set arrPageTmp(intPage) = New PageCells
                        End If
                    End If
                End If
                blnData = False
                If objBody.Container.name = "pic" Then
                    '卡片
                    If objBody.Container.Container Is picPaper(intReport) Then
                        If mobjReport.Items("_" & objBody.Index).SubIDs.count > 0 And mobjReport.Items("_" & objBody.Container.Index).数据源 <> "" And lngLastID <> objBody.Index Then
                            For Each tmpSubID In mobjReport.Items("_" & objBody.Index).SubIDs
                                If mobjReport.Items("_" & tmpSubID.id).内容 <> "" Then
                                    With mobjReport.Items("_" & tmpSubID.id)
                                        X = InStr(1, .内容, "]")
                                        Y = InStr(1, .内容, ".")
                                        Z = InStr(1, .内容, "[")
                                        If X > Z And X > Y And X <> 0 And Z <> 0 Then
                                            If Mid(.内容, Z + 1, Y - Z - 1) = mobjReport.Items("_" & objBody.Container.Index).数据源 Then
                                                blnData = True
                                                Exit For
                                            End If
                                        End If
                                    End With
                                End If
                            Next
                            If blnData Then
                                On Error Resume Next
                                If lngCurH \ mobjReport.Items("_" & objBody.Index).行高 < colCardRow("_" & objBody.Container.Index) Then
                                    If Err.Number = 0 Then colCardRow.Remove "_" & objBody.Container.Index
                                    colCardRow.Add lngCurH \ mobjReport.Items("_" & objBody.Index).行高, "_" & objBody.Container.Index
                                End If
                                On Error GoTo 0
                            End If
                        End If
                    End If
                End If
                lngLastID = objBody.Index
                
                '加入新的打印页描述
                '只有卡片外的表格才分页
                If blnPage Then
                    arrPage(intPage).Add objBody.Index, lngOutX, lngOutY, lngCurW + lngFixW, lngCurH + lngFixH, _
                        lngDesignH, lngRowB, lngRowE, lngVRowE, lngColB, lngColE, _
                        lngFixW, lngFixH, objItem.分栏, "_" & objBody.Index
                Else
                    If intPage = 0 Then
                        arrPageTmp(intPage).Add objBody.Index, lngOutX, lngOutY, lngCurW + lngFixW, lngCurH + lngFixH, _
                            lngDesignH, lngRowB, lngRowE, lngVRowE, lngColB, lngColE, _
                            lngFixW, lngFixH, objItem.分栏, "_" & objBody.Index
                    End If
                End If
                lngColB = lngColE + 1
                lngColE = lngColB - 1
                
                intPage = intPage + 1
            
                If blnApp Then
                    '附加表格跳过其参照表格的横向分页
                    If objItem.父ID <> 0 Then
                        arrTmp = arrPageTmp
                    Else
                        arrTmp = arrPage
                    End If
                    For i = intPage To UBound(arrTmp)
                        For Each objPageCell In arrTmp(i)
                            If objPageCell.id = intDepend Then
                                If objPageCell.ColB <> msh(intDepend).FixedCols Then
                                    intPage = intPage + 1
                                End If
                            End If
                        Next
                    Next
                    '重算下页可用的位置尺寸
                    '在新页中开始输出,在整个框架内输出
                    lngOutX = mobjReport.Items("_" & intDepend).X
                    lngOutY = mobjReport.Items("_" & intDepend).Y
                    Select Case bytKind
                        Case 0 '输出整体
                            lngBodyH = lngDesignH - lngFixH
                        Case 1 '仅输出表头
                            lngBodyH = 0
                        Case 2 '仅输出表体
                            lngBodyH = lngDesignH
                    End Select
                End If
            
            '任意表有分栏时或属于附加体中时不横向分页
            Loop Until lngColB > objBody.Cols - 1 Or _
                (objItem.类型 = 4 And (objItem.分栏 > 1 Or objItem.性质 = Val("1-附加表") Or blnHaveApp))
            
            lngColB = objBody.FixedCols
            lngColE = lngColB - 1
                            
            lngRowB = lngRowE + 1
            lngRowE = lngRowB - 1
        '所有行处理完了,由该表格也完了；只显示表头时只有一页
        Loop Until lngRowB > objBody.Rows - 1 Or (objItem.类型 = 4 And bytKind = 1)
    Next
    
    '卡片动态打印
    For Each objBody In pic
        '如果是动态打印
        If objBody.Index <> 0 And Not mobjReport.Items("_" & objBody.Index) Is Nothing Then
            If mobjReport.Items("_" & objBody.Index).数据源 <> "" Then
                lngRowB = 0
                lngRowE = 0
                intPage = 0
                lngCount = 0
                If mobjReport.Items("_" & objBody.Index).横向分栏 = 0 Then
                    If mobjReport.Fmts.Item("_" & bytFormat).纸向 = 1 Then
                        lngL = (mobjReport.Fmts.Item("_" & bytFormat).W - objBody.Left + mobjReport.Items("_" & objBody.Index).左右间距) \ (objBody.Width + mobjReport.Items("_" & objBody.Index).左右间距)
                    Else
                        lngL = (mobjReport.Fmts.Item("_" & bytFormat).H - objBody.Left + mobjReport.Items("_" & objBody.Index).左右间距) \ (objBody.Width + mobjReport.Items("_" & objBody.Index).左右间距)
                    End If
                Else
                    lngL = mobjReport.Items("_" & objBody.Index).横向分栏
                End If
                If mobjReport.Items("_" & objBody.Index).纵向分栏 = 0 Then
                    If mobjReport.Fmts.Item("_" & bytFormat).纸向 = 1 Then
                        lngW = (mobjReport.Fmts.Item("_" & bytFormat).H - objBody.Top + mobjReport.Items("_" & objBody.Index).上下间距) \ (objBody.Height + mobjReport.Items("_" & objBody.Index).上下间距)
                    Else
                        lngW = (mobjReport.Fmts.Item("_" & bytFormat).W - objBody.Top + mobjReport.Items("_" & objBody.Index).上下间距) \ (objBody.Height + mobjReport.Items("_" & objBody.Index).上下间距)
                    End If
                Else
                    lngW = mobjReport.Items("_" & objBody.Index).纵向分栏
                End If
                
                '一页可以容纳多少卡片
                lngC = lngW * lngL
                With mLibDatas("_" & mobjReport.Items("_" & objBody.Index).数据源).DataSet
                    If .RecordCount > 0 Then .MoveFirst
                    
                    '检查分组标识列
                    For i = 0 To .Fields.count - 1
                        If .Fields(i).name = "分组标识" Then
                            Exit For
                        End If
                    Next
                    
                    If i >= 0 And i <= .Fields.count - 1 Then
                        '有“分组标识”列
                        
                        '获取网格控件
                        intGridID = -1
                        For Each vsfTmp In msh
                            If vsfTmp.Index > 0 And Not vsfTmp.Container Is Nothing Then
                                If objBody.Index = vsfTmp.Container.Index Then
                                    intGridID = vsfTmp.Index
                                    Exit For
                                End If
                            End If
                        Next
                        
                        If intGridID >= 0 Then
                            '获取表格高
                            lngDesignH = GetGridDesignHeight(intGridID)
                            lngTop = msh(CInt(msh(intGridID).Tag)).Top
                            lngFixH = 0
                            For i = 0 To msh(CInt(msh(intGridID).Tag)).FixedRows - 1
                                lngFixH = lngFixH + msh(CInt(msh(intGridID).Tag)).RowHeight(i)
                            Next
'                            Select Case bytKind
'                            Case Val("0-表头表体")
'                                lngBodyH = lngDesignH - lngFixH
'                            Case Val("1-表头")
'                                lngBodyH = 0
'                            Case Val("2-表体")
'                                lngBodyH = lngDesignH
'                                lngFixH = 0
'                            End Select
                            lngBodyH = lngDesignH - lngFixH
                            
                            lngRowB = 0                 '卡片的开始行
                            lngRowE = 0                 '卡片的结束行
                            lngCurH = 0                 '实际行高
                            i = 0                       '页卡片计数
                            If .RecordCount > 0 Then
                                strTmp = "" & Nvl(!分组标识)
                            Else
                                strTmp = ""
                            End If
                            
                            '计算表格行数
                            Do While .EOF = False
                                lngRow = .AbsolutePosition - 1

                                '页的所有卡片
                                '分卡片逻辑：1.累计行高 > 表体高； 2.分组标识
                                If lngCurH + msh(intGridID).RowHeight(lngRow) > lngBodyH _
                                    Or strTmp <> "" & !分组标识 Then
                                    '行高超出卡片高
                                    colCard.Add lngRowB + 1 & "-" & lngRowE + 1
                                    strTmp = "" & !分组标识
                                    If lngCurH = 0 Then
                                        '只一行，不回退行
                                    Else
                                        '回退一行
                                        .MovePrevious
                                    End If
                                    lngCurH = 0
                                    i = i + 1
                                    lngRowB = lngRowE + 1
                                    lngRowE = lngRowB
                                Else
                                    lngCurH = lngCurH + msh(intGridID).RowHeight(lngRow)
                                    lngRowE = lngRow
                                    strTmp = "" & !分组标识
                                End If
                                
                                .MoveNext

                                '超当页卡片数量就产生一页的marrPageCard对象
                                If i > lngC - 1 Or .EOF Then
                                    If .EOF Then
                                        lngRowE = .RecordCount - 1
                                        colCard.Add lngRowB + 1 & "-" & lngRowE + 1
                                    End If
                                    Call AddPrintPage(intPage, objBody, colCard, lngRowB, lngRowE, lngW, lngL)
                                    Set colCard = New Collection
                                    intPage = intPage + 1
                                    i = 0
                                End If
                            Loop
                        Else
                            '无网格控件
                            lngRowB = 0                 '卡片的开始行
                            lngRowE = 0                 '卡片的结束行
                            i = 0                       '页卡片计数
                            If .RecordCount > 0 Then
                                strTmp = "" & Nvl(!分组标识)
                            Else
                                strTmp = ""
                            End If
                            
                            '计算卡片行数
                            Do While .EOF = False
                                lngRow = .AbsolutePosition - 1
                                
                                '页的所有卡片
                                '分卡片逻辑：1.分组标识
                                If strTmp <> "" & !分组标识 Then
                                    colCard.Add lngRowB + 1 & "-" & lngRowE + 1
                                    strTmp = "" & !分组标识

                                    '回退一行
                                    .MovePrevious
                                    
                                    i = i + 1
                                    lngRowB = lngRowE + 1
                                    lngRowE = lngRowB
                                Else
                                    lngRowE = lngRow
                                    strTmp = "" & !分组标识
                                End If
                                
                                .MoveNext

                                '超当页卡片数量就产生一页的marrPageCard对象
                                If i > lngC - 1 Or .EOF Then
                                    If .EOF Then
                                        lngRowE = .RecordCount - 1
                                        colCard.Add lngRowB + 1 & "-" & lngRowE + 1
                                    End If
                                    Call AddPrintPage(intPage, objBody, colCard, lngRowB, lngRowE, lngW, lngL)
                                    Set colCard = New Collection
                                    intPage = intPage + 1
                                    i = 0
                                End If
                            Loop
                        End If

                    Else
                        '无“分组标识”列，一个卡片只有一个记录
                        If .RecordCount <= lngC Then
                            '只一页多卡片
                            lngRowB = 0
                            lngRowE = .RecordCount - 1
                            For i = lngRowB To lngRowE
                                colCard.Add i + 1 & "-" & i + 1
                            Next
                            Call AddPrintPage(intPage, objBody, colCard, lngRowB, lngRowE, lngW, lngL)
                        Else
                            '多页多卡片
                            lngRowB = 0
                            Do While .EOF = False
                                '所有卡片
                                lngRowE = .AbsolutePosition - 1
                                colCard.Add lngRowE + 1 & "-" & lngRowE + 1
                                
                                .MoveNext
                                
                                If colCard.count >= lngC Or .EOF Then
                                    Call AddPrintPage(intPage, objBody, colCard, lngRowB, lngRowE, lngW, lngL)
                                    Set colCard = New Collection
                                    intPage = intPage + 1
                                    lngRowB = lngRowE + 1
                                End If
                            Loop
                        End If
                    End If
                End With
            End If
        End If
    Next
    
    '动态打印的表格单独处理
    If IsArray(arrPageTmp) Then
        If arrPageTmp(0).count > 0 Then
            For Each objPageCell In arrPageTmp(0)
                If IsArray(marrPageCard) Then
                    For i = 0 To UBound(marrPageCard)
                        On Error Resume Next
                        j = marrPageCard(i).Item("_" & mobjReport.Items("_" & objPageCell.id).父ID).id
                        If Err.Number = 0 Then
                            On Error GoTo 0
                            With marrPageCard(i).Item("_" & mobjReport.Items("_" & objPageCell.id).父ID)
                                    For j = 1 To .Item.count
                                        If Not IsArray(arrPage) Then
                                            ReDim arrPage(i) As PageCells  '第一次初始页
                                            Set arrPage(i) = New PageCells
                                        ElseIf i > UBound(arrPage) Then
                                            '如果该页已被其它表格占用,则不用再初始
                                            ReDim Preserve arrPage(i) As PageCells
                                            Set arrPage(i) = New PageCells
                                        End If
                                        If mobjReport.票据 Then
                                            lngVRowE = (mobjReport.Items("_" & objPageCell.id).H - objPageCell.FixH) \ mobjReport.Items("_" & objPageCell.id).行高 _
                                                    - (Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) + 1) _
                                                    + Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - 1
                                        Else
                                            lngVRowE = objPageCell.VRowE
                                        End If
                                        arrPage(i).Add objPageCell.id _
                                            , objPageCell.X + ((j - 1) Mod .Col) * (mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.id).父ID).W + mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.id).父ID).左右间距) _
                                            , objPageCell.Y + ((j - 1) \ .Col) * (mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.id).父ID).H + mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.id).父ID).上下间距) _
                                            , objPageCell.W, objPageCell.FixH + (Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) + 1) * mobjReport.Items("_" & objPageCell.id).行高 _
                                            , objPageCell.MaxH, Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) - 1, Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - 1 _
                                            , lngVRowE, objPageCell.ColB, objPageCell.ColE, objPageCell.FixW, objPageCell.FixH _
                                            , objPageCell.Copys, "_" & objPageCell.id + (j - 1)
                                    Next
                                
                            End With
                        End If
                        On Error GoTo 0
                    Next
                End If
            Next
        End If
    End If
    
    marrPage = arrPage
    CalcCellPage = True
End Function

Private Sub SetReportIndex(intIndex As Integer, objReport As Report)
'功能：根据当前要显示的报表,对其元素的索引加上后缀
'说明：主要用于区别报表组中的多张报表的不同元素
'注意：将元素的真实关键字也一并作调整
    Dim tmpItem As RPTItem, objItems As RPTItems
    Dim tmpSubID As RelatID, objSubIDs As RelatIDs
    Dim tmpCopyID As RelatID, objCopyIDs As RelatIDs
    
    Set objItems = New RPTItems
    For Each tmpItem In objReport.Items
        With tmpItem
            Set objSubIDs = New RelatIDs
            For Each tmpSubID In .SubIDs
                objSubIDs.Add tmpSubID.id & intIndex, "_" & tmpSubID.id & intIndex
            Next
            Set objCopyIDs = New RelatIDs
            For Each tmpCopyID In .CopyIDs
                objCopyIDs.Add tmpCopyID.id & intIndex, "_" & tmpCopyID.id & intIndex
            Next
            objItems.Add .id & intIndex, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头, _
                .X, .Y, .W, .H, .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, _
                .边框, .分栏, .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, _
                IIF(.父ID = 0, 0, .父ID & intIndex), objSubIDs, objCopyIDs, "_" & .id & intIndex, _
                .数据源, .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏, .Relations, _
                .ColProtertys, .水平反转
        End With
    Next
    
    Set objReport.Items = New RPTItems
    For Each tmpItem In objItems
        With tmpItem
            objReport.Items.Add .id, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头, _
                .X, .Y, .W, .H, .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, _
                .边框, .分栏, .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, .父ID, .SubIDs, .CopyIDs, _
                "_" & .id, .数据源, .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏, .Relations, _
                .ColProtertys, .水平反转
        End With
    Next
End Sub

Private Sub mnuViewStyle_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub SetView(bytStyle As Byte)
'功能：调整床位列表显示方式
'参数：bytstyle=0-大图标,1-小图标,2-列表,3-详细资料
    mnuViewStyle(0).Checked = False
    mnuViewStyle(1).Checked = False
    mnuViewStyle(2).Checked = False
    mnuViewStyle(3).Checked = False
    mnuViewStyle(bytStyle).Checked = True
    lvw.View = bytStyle
End Sub

Private Function GetSubReport(lngGroup As Long) As ADODB.Recordset
'功能：根据报表组ID,获取其子报表的信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH

    strSQL = "Select 组ID,报表ID,序号,功能 From zlRPTSubs Where 组ID=[1] Order by 序号"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngGroup)
    If Not rsTmp.EOF Then Set GetSubReport = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetGroupInfo(lngGroup As Long) As ADODB.Recordset
'功能：根据报表组ID获取其信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ID,编号,名称,说明,系统,程序ID,发布时间 From zlRPTGroups Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngGroup)
    If Not rsTmp.EOF Then Set GetGroupInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitReportPars()
'功能：在报表组中,根据当前报表的数据内容显示应输入参数
    Dim i As Integer, j As Integer
    Dim tmpPar As RPTPar, strTmp As String
    Dim lngCurH As Long, objTmp As Object
    Dim intCurTab As Integer, objLoad As Object
    Dim strGroup As String, objGroup As Object
    Dim blnCmd As Boolean, blnExist As Boolean
    Dim strPre As String, strCur As String
    Dim blntmp As Boolean, lngTmp As Long
    
    For Each objLoad In lblName
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In txt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cmd
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cbo
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In dtp
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In opt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In chk
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fra
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fraGroup
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    
    blnMatch = False
    
    '产生参数输入框组
    i = 0: lngCurH = lblName(0).Top
    For Each tmpPar In mobjPars
        i = i + 1
        
        Load lblName(i)
        lblName(i).Caption = tmpPar.名称 & "(&" & i & ")"
        lblName(i).ToolTipText = tmpPar.名称
        lblName(i).Left = txt(0).Left - lblName(i).Width - 30
        lblName(i).Top = lngCurH
        lblName(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
        lblName(i).Visible = True
        
        If tmpPar.缺省值 = "固定值列表…" Then
            If tmpPar.格式 = 0 Then '下拉框
                Load cbo(i): Set objTmp = cbo(i)
                If tmpPar.是否锁定 Then objTmp.Enabled = False
                cbo(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                cbo(i).Left = cbo(0).Left: cbo(i).Top = lblName(i).Top - (cbo(i).Height - lblName(i).Height) / 2
                '不好的分隔符
                For j = 0 To UBound(Split(tmpPar.值列表, "|"))
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    
                    If Left(strTmp, 1) = "√" Then
                        cbo(i).AddItem Mid(strTmp, 2)
                        If cbo(i).ListIndex = -1 Then cbo(i).ListIndex = cbo(i).NewIndex
                    Else
                        cbo(i).AddItem strTmp
                    End If
                    '重置条件时Reserve存放了"显示值|绑定值"
                    '根据上次显示值来定位缺省项
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "√" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then cbo(i).ListIndex = cbo(i).NewIndex
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then cbo(i).ListIndex = cbo(i).NewIndex
                        End If
                        
                        '上次人为输入的值与某个绑定值相同,则定位
                        '因为多个选择值中绑定值可能重复,所以此段可不要
                        If Split(tmpPar.Reserve, "|")(0) = Split(Split(tmpPar.值列表, "|")(j), ",")(1) Then
                            cbo(i).ListIndex = cbo(i).NewIndex
                        End If
                    End If
                Next
                cbo(i).Visible = True
            ElseIf tmpPar.格式 = 1 Then '单选框
                Load fra(i): Set objTmp = fra(i)
                fra(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                fra(i).Left = fra(0).Left: fra(i).Top = lblName(i).Top - 50
                
                lblName(i).Visible = False
                fra(i).Caption = lblName(i).Caption
                                
                j = UBound(Split(tmpPar.值列表, "|")) + 1 '可选数
                j = CInt((j / 3) + 0.4) '行数
                
                fra(i).Height = fra(0).Height + (j - 1) * (opt(0).Height * 1.6) - opt(0).Height * 0.3
                
                blnExist = False '是否已经按上次条件值设置了当前值
                '不好的分隔符
                For j = 0 To UBound(Split(tmpPar.值列表, "|"))
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    
                    Load opt(opt.UBound + 1)
                    If tmpPar.是否锁定 Then opt(opt.UBound).Enabled = False
                    Set opt(opt.UBound).Container = fra(i)
                    opt(opt.UBound).TabIndex = intCurTab: intCurTab = intCurTab + 1
                    opt(opt.UBound).Tag = Split(Split(tmpPar.值列表, "|")(j), ",")(1) '存放绑定值
                    
                    If InStr(",0,1,3,", "," & UBound(Split(tmpPar.值列表, "|")) & ",") > 0 Then
                        '只有1,2,4个的情况特殊处理
                        If j = 0 Or j = 1 Then 'Top
                            opt(opt.UBound).Top = opt(0).Top
                        Else
                            opt(opt.UBound).Top = opt(0).Top + opt(0).Height * 1.6
                        End If
                        If j = 0 Or j = 2 Then 'Left
                            opt(opt.UBound).Left = opt(0).Left + 150
                        Else
                            opt(opt.UBound).Left = opt(0).Left + (opt(0).Width * 1.4 + 60) + 150
                        End If
                        
                        If Left(strTmp, 1) = "√" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    Else
                        opt(opt.UBound).Top = opt(0).Top + (CInt(((j + 1) / 3) + 0.4) - 1) * (opt(0).Height * 1.6)
                        opt(opt.UBound).Left = opt(0).Left + (IIF(((j + 1) Mod 3) = 0, 3, ((j + 1) Mod 3)) - 1) * (opt(0).Width + 60)
                        
                        If Left(strTmp, 1) = "√" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    End If

                    opt(opt.UBound).Width = TextWidth(opt(opt.UBound).Caption) + 300
                    
                    '重置条件时Reserve存放了"显示值|绑定值"
                    '根据上次选择值来定位缺省项
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "√" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                opt(opt.UBound).Value = True
                                blnExist = True
                            End If
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                opt(opt.UBound).Value = True
                                blnExist = True
                            End If
                        End If
                    End If
                    
                    opt(opt.UBound).Visible = True
                Next
                
                fra(i).ZOrder 1 '放在最下面
                fra(i).Visible = True
            ElseIf tmpPar.格式 = 2 Then '单个复选框
                lblName(i).Visible = False
                
                blntmp = True
                Load chk(i): Set objTmp = chk(i)
                If tmpPar.是否锁定 Then objTmp.Enabled = False
                chk(i).Caption = lblName(i).Caption
                chk(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                chk(i).Left = chk(0).Left: chk(i).Top = lblName(i).Top - (chk(i).Height - lblName(i).Height) / 2
                chk(i).Width = TextWidth(chk(i).Caption) + 230
                
                '不好的分隔符
                If Left(Split(Split(tmpPar.值列表, "|")(0), ",")(0), 1) = "√" Then chk(i).Value = 1
                For j = 0 To 1
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    '重置条件时Reserve存放上次了"显示值|绑定值"
                    '根据上次选择值来定位本次缺省项
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "√" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                If Left(strTmp, 1) = "√" Then
                                    chk(i).Value = 1
                                Else
                                    chk(i).Value = 0
                                End If
                            End If
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                If Left(strTmp, 1) = "√" Then
                                    chk(i).Value = 1
                                Else
                                    chk(i).Value = 0
                                End If
                            End If
                        End If
                    End If
                Next
                chk(i).Visible = True
            End If
        ElseIf tmpPar.缺省值 = "选择器定义…" Then
            Load txt(i): Set objTmp = txt(i)
            If tmpPar.是否锁定 Then objTmp.Enabled = False
            txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
            txt(i).Left = txt(0).Left: txt(i).Top = lblName(i).Top - (txt(i).Height - lblName(i).Height) / 2
            txt(i).ToolTipText = "按 F2 打开选择器"
            txt(i).Locked = True
                                                
            blnCmd = True
            If tmpPar.Reserve Like "*|*" Then
                If Split(tmpPar.Reserve, "|")(0) <> "" Then
                    '重置条件时Reserve存放了"显示值|绑定值"
                    txt(i).Text = Split(tmpPar.Reserve, "|")(0)
                    txt(i).Tag = Split(tmpPar.Reserve, "|")(1)
                    
                    '虽然有缺省,但如果没有其它可选则不可见
                    strTmp = ""
                    If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                Else
                    '使用缺省定义的缺省值
                    If tmpPar.值列表 Like "*|*" Then
                        txt(i).Text = Split(tmpPar.值列表, "|")(0)
                        txt(i).Tag = Split(tmpPar.值列表, "|")(1)
                    ElseIf tmpPar.明细SQL <> "" Then
                        '取明细SQL结果中第一行值,如果只有一行,则不用选
                        strTmp = ""
                        If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                        If strTmp <> "" Then
                            txt(i).Text = Split(strTmp, "|")(0)
                            txt(i).Tag = Split(strTmp, "|")(1)
                            If tmpPar.格式 = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                            blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                        Else
                            blnCmd = False
                        End If
                    End If
                End If
            Else
                If tmpPar.值列表 Like "*|*" Then
                    '使用缺省定义的缺省值
                    txt(i).Text = Split(tmpPar.值列表, "|")(0)
                    txt(i).Tag = Split(tmpPar.值列表, "|")(1)
                    
                    '虽然有缺省,但如果没有其它可选则不可见
                    strTmp = ""
                    If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                ElseIf tmpPar.明细SQL <> "" Then
                    '取明细SQL结果中第一行值,如果只有一行,则不用选
                    strTmp = ""
                    If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        txt(i).Text = Split(strTmp, "|")(0)
                        txt(i).Tag = Split(strTmp, "|")(1)
                        If tmpPar.格式 = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                    Else
                        blnCmd = False
                    End If
                End If
            End If
                        
            Load cmd(i)
            If tmpPar.是否锁定 Then cmd(i).Enabled = False
            cmd(i).Top = txt(i).Top + 30
            cmd(i).Left = txt(i).Left + txt(i).Width - cmd(i).Width - 30
            cmd(i).Height = txt(i).Height - 45
            cmd(i).TabStop = False
            cmd(i).ZOrder
            
            txt(i).Visible = True
            cmd(i).Visible = blnCmd
            
            '可否输入匹配
            txt(i).Locked = Not ((InStr(tmpPar.分类SQL, "[*]") > 0 Or InStr(tmpPar.明细SQL, "[*]") > 0) And blnCmd)
        Else
            If tmpPar.类型 = 2 Then
                Load dtp(i): Set objTmp = dtp(i)
                If tmpPar.是否锁定 Then objTmp.Enabled = False
                dtp(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                dtp(i).Left = dtp(0).Left: dtp(i).Top = lblName(i).Top - (dtp(i).Height - lblName(i).Height) / 2
                If InStr(tmpPar.缺省值, ":") > 0 Or InStr(tmpPar.缺省值, "时间") > 0 Then
                    dtp(i).CustomFormat = "yyyy年MM月dd日 HH:mm:ss"
                    dtp(i).Width = 2460
                Else
                    dtp(i).CustomFormat = "yyyy年MM月dd日"
                    dtp(i).Width = 1635
                End If
                If tmpPar.缺省值 <> "" Then
                    If Left(tmpPar.缺省值, 1) = "&" Then
                        dtp(i).Value = GetParVBMacro(tmpPar.缺省值)
                    Else
                        dtp(i).Value = Format(tmpPar.缺省值, dtp(i).CustomFormat)
                    End If
                Else
                    dtp(i).Value = Currentdate
                End If
                
'                '注册表保存值
'                If dtp(i).CustomFormat Like "*HH:mm:ss" And Left(tmpPar.缺省值, 1) <> "&" Then
'                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.名称, lblName(i).ToolTipText & "时间", Format(dtp(i).Value, "HH:mm:ss"))
'                    dtp(i).Value = CDate(Format(dtp(i).Value, Left(dtp(i).CustomFormat, InStr(dtp(i).CustomFormat, "HH:mm:ss") - 1)) & strTmp)
'                End If
                
                dtp(i).Visible = True
            Else
                Load txt(i): Set objTmp = txt(i)
                If tmpPar.是否锁定 Then objTmp.Enabled = False
                txt(i).Left = txt(0).Left: txt(i).Top = lblName(i).Top - (txt(i).Height - lblName(i).Height) / 2
                txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                txt(i).Text = tmpPar.缺省值
                txt(i).Visible = True
            End If
        End If
        If objTmp.name = "fra" Then
            lngCurH = lngCurH + objTmp.Height + 180
        Else
            lngCurH = lngCurH + txt(0).Height + 150
        End If
        
        lblName(i).Tag = tmpPar.组名 & "," & objTmp.name
        If tmpPar.缺省值 = "选择器定义…" Then lblName(i).Tag = lblName(i).Tag & ",cmd"
    Next
    
    fraSplit.Top = lngCurH
    
    '处理参数组
    For i = 1 To lblName.UBound
        strCur = ""
        If strGroup <> CStr(Split(lblName(i).Tag, ",")(0)) And CStr(Split(lblName(i).Tag, ",")(0)) <> "" Then
            Load fraGroup(fraGroup.UBound + 1)
            Set objGroup = fraGroup(fraGroup.UBound)
            objGroup.Caption = CStr(Split(lblName(i).Tag, ",")(0))
            objGroup.Top = lblName(i).Top - 150
            objGroup.ZOrder 1
            objGroup.Visible = True
            
            Select Case CStr(Split(lblName(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            lngCurH = 195 '当前Top位置
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            objTmp.Left = 1250
            
            Set lblName(i).Container = objGroup
            lblName(i).Top = objTmp.Top + (objTmp.Height - lblName(i).Height) / 2
            lblName(i).Left = objTmp.Left - lblName(i).Width - 30
            lblName(i).Caption = GetLenStr(lblName(i).ToolTipText, 900, Me) & Mid(lblName(i).Caption, InStr(lblName(i).Caption, "("))
            
            If UBound(Split(lblName(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If

            lngCurH = lngCurH + txt(0).Height + 50 '当前Top位置
        ElseIf strGroup = CStr(Split(lblName(i).Tag, ",")(0)) And CStr(Split(lblName(i).Tag, ",")(0)) <> "" Then
            strCur = "Add"
            Select Case CStr(Split(lblName(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            objTmp.Left = 1250
            
            Set lblName(i).Container = objGroup
            lblName(i).Top = objTmp.Top + (objTmp.Height - lblName(i).Height) / 2
            lblName(i).Left = objTmp.Left - lblName(i).Width - 30
            lblName(i).Caption = GetLenStr(lblName(i).ToolTipText, 900, Me) & Mid(lblName(i).Caption, InStr(lblName(i).Caption, "("))
            
            If UBound(Split(lblName(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If
                        
            lngCurH = lngCurH + txt(0).Height + 50 '当前Top位置
            
            objGroup.Height = objTmp.Top + objTmp.Height + 90  '框高度
            
            '该框以下的条件输入全部下移
            For j = i + 1 To lblName.UBound
                If Split(lblName(j).Tag, ",")(0) <> "fra" Then
                    lblName(j).Top = lblName(j).Top + 60
                    Select Case CStr(Split(lblName(j).Tag, ",")(1))
                        Case "txt"
                            txt(j).Top = txt(j).Top + 60
                        Case "cbo"
                            cbo(j).Top = cbo(j).Top + 60
                        Case "dtp"
                            dtp(j).Top = dtp(j).Top + 60
                        Case "chk"
                            chk(j).Top = chk(j).Top + 60
                    End Select
                    If UBound(Split(lblName(j).Tag, ",")) = 2 Then
                        cmd(j).Top = cmd(j).Top + 60
                    End If
                End If
            Next
        End If
        If strPre = "Add" And strCur = "" Then
            fraSplit.Top = fraSplit.Top + 60
        End If
        strPre = strCur
        strGroup = CStr(Split(lblName(i).Tag, ",")(0))
    Next
    
    '没有参数组但有多项单选框时,向该框对齐
    If fraGroup.UBound = 0 And fra.UBound > 0 Then
        For Each objTmp In fra
            objTmp.Left = txt(0).Left - 1000
        Next
    End If
    
    cmdLoad.Top = fraSplit.Top + 180
    cmdDefault.Top = fraSplit.Top + 180
    
    fraSplit.Visible = (lblName.UBound > 0)
    cmdLoad.Visible = (lblName.UBound > 0)
    cmdDefault.Visible = (lblName.UBound > 0)
    
    cmdLoad.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdDefault.TabIndex = intCurTab
    
    cmdSelAll.Top = cmdLoad.Top: cmdSelNone.Top = cmdSelAll.Top
    cmdSelAll.Visible = blntmp
    cmdSelNone.Visible = blntmp
    If Me.Visible Then
        On Error Resume Next
        If picPar.Height < cmdLoad.Top + cmdLoad.Height + 100 Then
            lngTmp = cmdLoad.Top + cmdLoad.Height + 100 - picPar.Height
            picPar.Height = picPar.Height + lngTmp
            picPar.Top = picPar.Top - lngTmp: lblPar_S.Top = lblPar_S.Top - lngTmp
            lvw.Height = lvw.Height - lngTmp
        End If
    End If
    
    '更新弹出菜单
    Call LoadCondsMenu
End Sub

Private Function ReSetReportPars() As Boolean
'功能：重新设置当前报表由用户输入的参数
    Dim i As Integer, j As Integer
    Dim strTmp As String, strDisp As String
    Dim strParName As String, curDate As Date
    
    '先检查合法性
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).缺省值 = "固定值列表…" Then
            Select Case mobjPars("_" & strParName).格式
                Case 0
                    If Trim(cbo(i).Text) = "" Then
                        MsgBox "请选择""" & strParName & """的条件值！", vbInformation, App.Title
                        If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                        Exit Function
                    End If
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '是否人为输入
                        '类型检查
                        Select Case mobjPars("_" & strParName).类型
                            Case 1
                                If Not IsNumeric(cbo(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                                    If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                                    Exit Function
                                End If
                            Case 2
                                If Not IsDate(cbo(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                    If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                                    Exit Function
                                End If
                        End Select
                    End If
            End Select
        ElseIf mobjPars("_" & strParName).缺省值 = "选择器定义…" Then
            If Trim(txt(i).Text) = "" Then
                MsgBox "请选择""" & strParName & """的条件值！", vbInformation, App.Title
                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                Exit Function
            End If
            If txt(i).Tag = "" Then '是否人为输入
                If mobjPars("_" & strParName).值列表 Like "*|*" Then
                    If Split(mobjPars("_" & strParName).值列表, "|")(0) <> txt(i).Text Then
                        '类型检查
                        Select Case mobjPars("_" & strParName).类型
                            Case 1
                                If Not IsNumeric(txt(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                                    If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                    Exit Function
                                End If
                            Case 2
                                If Not IsDate(txt(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                    If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                    Exit Function
                                End If
                        End Select
                    Else
                        '输入值与定义的缺省值相同,则还原为缺省值
                        txt(i).Tag = Split(mobjPars("_" & strParName).值列表, "|")(1)
                    End If
                Else
                    '类型检查
                    Select Case mobjPars("_" & strParName).类型
                        Case 1
                            If Not IsNumeric(txt(i).Text) Then
                                MsgBox "你输入的""" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                Exit Function
                            End If
                        Case 2
                            If Not IsDate(txt(i).Text) Then
                                MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                Exit Function
                            End If
                    End Select
                End If
            End If
        Else
            Select Case mobjPars("_" & strParName).类型
                Case 0, 3
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "请输入""" & strParName & """的条件值！", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If TLen(txt(i).Text) > 255 Then
                        MsgBox """" & strParName & """的条件值长度不能超过255个字符！", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                Case 1
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "请输入""" & strParName & """的条件值！", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If TLen(txt(i).Text) > 255 Then
                        MsgBox """" & strParName & """的条件值长度不能超过255个字符！", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If Not IsNumeric(txt(i).Text) Then
                        MsgBox """" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                Case 2 '日期时间最大值检查
                    curDate = Currentdate
                    If Not (mobjPars("_" & strParName).缺省值 Like "&下一*" Or mobjPars("_" & strParName).Reserve Like "&下一*" Or _
                        mobjPars("_" & strParName).缺省值 Like "&后一*" Or mobjPars("_" & strParName).Reserve Like "&后一*" Or _
                        mobjPars("_" & strParName).缺省值 Like "&*结束*" Or mobjPars("_" & strParName).Reserve Like "&*结束*" Or _
                        mobjPars("_" & strParName).缺省值 Like "&*月末*" Or mobjPars("_" & strParName).缺省值 Like "&*年末*" Or _
                        mobjPars("_" & strParName).Reserve Like "&*月末*" Or mobjPars("_" & strParName).Reserve Like "&*年末*") Then
                        
                        If mobjPars("_" & strParName).缺省值 Like "*时间*" Or mobjPars("_" & strParName).Reserve Like "*时间*" Then
                            If Format(dtp(i).Value, "yyyy-MM-dd HH:mm:ss") > Format(curDate, "yyyy-MM-dd HH:mm:ss") Then
                                MsgBox """" & strParName & """ 的条件值不能超过当前时间！", vbInformation, App.Title
                                If dtp(i).Enabled And dtp(i).Visible Then dtp(i).SetFocus
                                Exit Function
                            End If
                        Else
                            If Format(dtp(i).Value, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
                                MsgBox """" & strParName & """ 的条件值不能超过当前日期！", vbInformation, App.Title
                                If dtp(i).Enabled And dtp(i).Visible Then dtp(i).SetFocus
                                Exit Function
                            End If
                        End If
                    End If
            End Select
        End If
    Next
        
    '再取值
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).缺省值 = "固定值列表…" Then '不好的分隔符
            Select Case mobjPars("_" & strParName).格式
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '是否人为输入
                        'Reserve字段保存本次条件的"宏条件值|显示值"
                        mobjPars("_" & strParName).Reserve = "固定值列表…|" & cbo(i).Text
                        mobjPars("_" & strParName).缺省值 = cbo(i).Text
                    Else
                        '列表选择
                        'Reserve字段保存本次条件的"宏条件值|显示值"
                        mobjPars("_" & strParName).Reserve = "固定值列表…|" & cbo(i).Text
                        strTmp = mobjPars("_" & strParName).值列表
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "√" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                mobjPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                'Reserve字段保存本次条件的"宏条件值|显示值"
                                mobjPars("_" & strParName).Reserve = "固定值列表…|" & opt(j).ToolTipText
                                mobjPars("_" & strParName).缺省值 = opt(j).Tag
                            End If
                        End If
                    Next
                Case 2
                    'Reserve字段保存本次条件的"宏条件值|显示值"
                    strTmp = mobjPars("_" & strParName).值列表
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "√" Then
                                mobjPars("_" & strParName).Reserve = "固定值列表…|" & strDisp
                                mobjPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        Else
                            If Left(strDisp, 1) = "√" Then
                                mobjPars("_" & strParName).Reserve = "固定值列表…|" & Mid(strDisp, 2)
                                mobjPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).缺省值 = "选择器定义…" Then
            If txt(i).Tag = "" Then '是否人为输入
                'Reserve字段保存本次条件的"宏条件值|显示值"
                mobjPars("_" & strParName).Reserve = "选择器定义…|"
                mobjPars("_" & strParName).缺省值 = txt(i).Text
            Else
                '列表选择
                'Reserve字段保存本次条件的"宏条件值|显示值"
                mobjPars("_" & strParName).Reserve = "选择器定义…|" & txt(i).Text
                mobjPars("_" & strParName).缺省值 = txt(i).Tag
            End If
        Else
            Select Case mobjPars("_" & strParName).类型
                Case 0, 1, 3
                    mobjPars("_" & strParName).缺省值 = txt(i).Text
                Case 2
                    If mobjPars("_" & strParName).缺省值 Like "&*" Then
                        mobjPars("_" & strParName).Reserve = mobjPars("_" & strParName).缺省值
                    End If
                    mobjPars("_" & strParName).缺省值 = Format(dtp(i).Value, dtp(i).CustomFormat)
                    '保存到注册表
                    If dtp(i).CustomFormat Like "*HH:mm:ss" Then
                        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.名称, lblName(i).ToolTipText & "时间", Format(dtp(i).Value, "HH:mm:ss")
                    End If
            End Select
        End If
    Next
    
    Call ReplaceInputPars(mobjPars)
    
    ReSetReportPars = True
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim LngIdx As Long
    
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
    If InStr("~`!@#$^&"";|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If mobjPars("_" & lblName(Index).ToolTipText).类型 = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii <> 8 Then
        If SendMessage(cbo(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
        LngIdx = MatchIndex(cbo(Index), KeyAscii)
        If LngIdx <> -2 Then cbo(Index).ListIndex = LngIdx
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Function GetValues() As Collection
'功能：获取现有的界面上的参数值
    Dim i As Integer, j As Integer
    Dim strParName As String, strTmp As String
    Dim strDisp As String, colValue As New Collection
     
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).缺省值 = "固定值列表…" Then
            Select Case mobjPars("_" & strParName).格式
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '是否人为输入
                        'Reserve字段保存本次条件的"宏条件值|显示值"
                        colValue.Add cbo(i).Text, "_" & strParName
                    Else
                        '列表选择
                        'Reserve字段保存本次条件的"宏条件值|显示值"
                        '不好的分隔符
                        strTmp = mobjPars("_" & strParName).值列表
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "√" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                colValue.Add opt(j).Tag, "_" & strParName
                            End If
                        End If
                    Next
                Case 2
                    'Reserve字段保存本次条件的"宏条件值|显示值"
                    '不好的分隔符
                    strTmp = mobjPars("_" & strParName).值列表
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "√" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        Else
                            If Left(strDisp, 1) = "√" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).缺省值 = "选择器定义…" Then
            If txt(i).Tag = "" Then '是否人为输入
                'Reserve字段保存本次条件的"宏条件值|显示值"
                colValue.Add txt(i).Text, "_" & strParName
            Else
                '列表选择
                'Reserve字段保存本次条件的"宏条件值|显示值"
                colValue.Add txt(i).Tag, "_" & strParName
            End If
        Else
            Select Case mobjPars("_" & strParName).类型
                Case 0, 1, 3
                    colValue.Add txt(i).Text, "_" & strParName
                Case 2
                    colValue.Add Format(dtp(i).Value, dtp(i).CustomFormat), "_" & strParName
            End Select
        End If
    Next
    Set GetValues = colValue
End Function

Private Sub cmd_Click(Index As Integer)
    Dim tmpPar As RPTPar, str明细对象 As String, str分类对象 As String
    Dim frmNewSelect As New frmSelect
    Dim strSQL明细 As String, strSQL分类 As String
    Dim colValue As New Collection    '参数现有的值
    
    For Each tmpPar In mobjPars
        If tmpPar.名称 = lblName(Index).ToolTipText Then
            If blnMatch And txt(Index).Tag = "" Then frmNewSelect.strMatch = txt(Index).Text
            
            If InStr(tmpPar.对象, "|") > 0 Then
                str明细对象 = Split(tmpPar.对象, "|")(0)
                str分类对象 = Split(tmpPar.对象, "|")(1)
            End If
            strSQL明细 = tmpPar.明细SQL
            strSQL分类 = tmpPar.分类SQL
            Set colValue = GetValues
            Call CheckParsRela(strSQL明细, Nothing, tmpPar.名称, True, colValue, mobjPars)
            Call CheckParsRela(strSQL分类, Nothing, tmpPar.名称, True, colValue, mobjPars)
            frmNewSelect.strSQLList = SQLOwner(RemoveNote(strSQL明细), str明细对象)
            frmNewSelect.strSQLTree = SQLOwner(RemoveNote(strSQL分类), str分类对象)
            frmNewSelect.strFLDList = tmpPar.明细字段
            frmNewSelect.strFLDTree = tmpPar.分类字段
            frmNewSelect.strParName = tmpPar.名称
            frmNewSelect.bytType = tmpPar.类型
            frmNewSelect.mblnMulti = tmpPar.格式 = 1
            frmNewSelect.mintConnect = GetDBConnectNo(tmpPar, mobjReport.Datas)
            frmNewSelect.lngSeekHwnd = cmd(Index).hwnd
            
            On Error Resume Next
            Err.Clear
            
            frmNewSelect.Show 1, Me
            If frmNewSelect.mblnOK Then
                txt(Index).Text = frmNewSelect.strOutDisp
                txt(Index).Tag = frmNewSelect.strOutBand
                Unload frmNewSelect
                SendKeys "{Tab}"
            ElseIf blnMatch Then
                txt(Index).Text = ""
                txt(Index).Tag = ""
            End If
            
            blnMatch = False
            Exit For
        End If
    Next
    txt(Index).SetFocus
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub dtp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SelAll txt(Index)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And txt(Index).ToolTipText <> "" Then
        If cmd(Index).Enabled And cmd(Index).Visible Then Call cmd_Click(Index)
    End If
    If txt(Index).Locked Then Exit Sub
    
    '人为输入时(不选择)，清除绑定值作为人为输入的标志
    '144=Num;112-123=F1-F12;229=开始输入汉字
    If KeyCode >= 48 And KeyCode <> 144 And KeyCode <> 229 _
        And Not (KeyCode >= 112 And KeyCode <= 123) Then
        txt(Index).Tag = ""
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
            '想输入匹配
            KeyAscii = 0
            If txt(Index).Text <> "" Then
                If cmd(Index).Enabled And cmd(Index).Visible Then
                    blnMatch = True
                    Call cmd_Click(Index)
                End If
            End If
            Exit Sub
        Else
            '想移动焦点
            KeyAscii = 0: SendKeys "{Tab}": Exit Sub
        End If
    End If
    
    If txt(Index).Locked Then Exit Sub
    
    If InStr("~`!@#$^&"";|'" & Chr(3) & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If txt(Index).ToolTipText = "" And mobjPars("_" & lblName(Index).ToolTipText).类型 = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '人为输入时(不选择)，清除绑定值作为人为输入的标志
    '这里只处理汉字,其它在KeyDown中处理
    If KeyAscii < 0 Then txt(Index).Tag = ""
End Sub

Private Sub ShowStatGrid(objItem As RPTItem)
'功能：在查询界面上组织显示一个分类表格(包含其左联接表格)
    Dim mshBody As Object, tmpItem As RPTItem, tmpID As RelatID
    Dim rsGroup As ADODB.Recordset, rsVsc As ADODB.Recordset, rsHsc As ADODB.Recordset
    Dim arrStat() As Variant, strVscStat As String, strHscStat As String, strStat As String
    Dim strVsc As String, strHsc As String, strVscOrder As String, strHscOrder As String
    Dim strFilter As String, strAlign As String, strTmp As String
    Dim i As Long, j As Long, k As Long, l As Long, M As Long
    Dim X As Long, Y As Long, Z As Long '分类子项数
    Dim strFormat As String, strSort As String, blnHide As Boolean, blnDo As Boolean
    Dim arrLevel() As String, arrMerge() As String, arrCount() As Long
    Dim arrHead() As String
        
    '处理左联接表格所加变量
    Dim lngCurCols As Long, objCurItem As RPTItem
    Dim strLink As String, lngMaxY As Long
    Dim lngGrid As Long, strTopRow As String
    Dim lngDiff As Long
    Dim lngStatistics As Long
    
    '改进后的填数方法变量
    Dim colVsc As Collection, colHsc As Collection
    Dim strKey As String, lngRow As Long, lngCol As Long, StrFmt As String
    
    Dim varIFValue As Variant
    Dim objColProp As RPTColProterty
    Dim objStatusGridItem As RPTItem
    Dim colTmp As Collection
    
    On Error GoTo hErr
    
    With objItem
        Load msh(.id)
        Set msh(.id).Container = picPaper(intReport)
        Set mshBody = msh(.id)
        
        mshBody.Redraw = False
        
        mshBody.ForeColor = .前景
        mshBody.ForeColorFixed = .前景
        mshBody.BackColor = .背景
        mshBody.BackColorFixed = .背景
        mshBody.GridColor = .网格
        mshBody.GridColorFixed = .网格
        mshBody.Font.name = .字体
        mshBody.Font.Size = .字号
        mshBody.Font.Bold = .粗体
        mshBody.Font.Italic = .斜体
        mshBody.Font.Underline = .下线
        mshBody.GridLineWidth = IIF(.表格线加粗, 2, 1)
        'Set mshBody.FontFixed = mshBody.Font
        
        mshBody.Left = .X: mshBody.Top = .Y
        mshBody.Height = .H: mshBody.Width = 0
        mshBody.FixedRows = 0
    
        '获取左接接表格相关信息
        strLink = strLink & "|" & .id
        For Each tmpItem In mobjReport.Items
            If tmpItem.格式号 = bytFormat And tmpItem.类型 = 5 And tmpItem.性质 = 2 And tmpItem.参照 = .名称 Then
                strLink = strLink & "|" & tmpItem.id
            End If
        Next
        strLink = Mid(strLink, 2)
    End With
        
    objItem.表头 = ""
    strTopRow = ""
    
    blnHide = True
    lngCurCols = 0
    lngMaxY = 0
    For lngGrid = 0 To UBound(Split(strLink, "|"))
        Set objCurItem = mobjReport.Items("_" & Split(strLink, "|")(lngGrid))
        With objCurItem
            mshBody.Width = mshBody.Width + .W
            
            '统计子项数
            '左联接的表格不再处理纵向分类
            If lngGrid = 0 Then
                strVsc = "": strVscOrder = "": X = 0
            End If
            strHsc = "": strHscOrder = "" '存放横纵向项目名称及排序字段名称
            Y = 0: Z = 0
            For Each tmpID In .SubIDs
                Set tmpItem = mobjReport.Items("_" & tmpID.id)
                Select Case tmpItem.类型
                    Case 7 '纵向分类
                        If lngGrid = 0 Then
                            X = X + 1
                            If tmpItem.排序 <> "" Then strVscOrder = strVscOrder & "|" & tmpItem.排序
                            strVsc = strVsc & "|" & tmpItem.内容
                        End If
                    Case 8 '横向分类
                        Y = Y + 1
                        If tmpItem.排序 <> "" Then strHscOrder = strHscOrder & "|" & tmpItem.排序
                        strHsc = strHsc & "|" & tmpItem.内容
                    Case 9 '统计项
                        Z = Z + 1
                End Select
            Next
            If Y > lngMaxY Then lngMaxY = Y
            If lngGrid = 0 Then
                strVsc = Mid(strVsc, 2)
                strVscOrder = Mid(strVscOrder, 2)
            End If
            strHsc = Mid(strHsc, 2)
            strHscOrder = Mid(strHscOrder, 2)
            
            '构造空表格架
            If lngGrid = 0 Then
                mshBody.FixedRows = Y + 1
            Else
                If Y + 1 > mshBody.FixedRows Then
                    lngDiff = Y + 1 - mshBody.FixedRows '行数定位偏移(左联接时可能增加固定行数)
                    For i = 1 To Y + 1 - mshBody.FixedRows
                        mshBody.AddItem "", mshBody.FixedRows
                        mshBody.FixedRows = mshBody.FixedRows + 1
                        For j = 0 To mshBody.Cols - 1 '加入新横向分类行时补充前一个表格的行内容
                            mshBody.TextMatrix(mshBody.FixedRows - 1, j) = mshBody.TextMatrix(mshBody.FixedRows - 2, j)
                        Next
                    Next
                End If
            End If
            mshBody.Cols = lngCurCols + IIF(lngGrid = 0, X, 0) + Z
            If lngGrid = 0 Then
                mshBody.Rows = Y + 2
                mshBody.FixedCols = X
            End If
            lngStatistics = 0
            For Each tmpID In .SubIDs
                Set tmpItem = mobjReport.Items("_" & tmpID.id)
                Select Case tmpItem.类型
                    Case 7 '纵向分类
                        If lngGrid = 0 Then
                            For i = 0 To Y
                                mshBody.TextMatrix(i, tmpItem.序号) = tmpItem.内容
                            Next
                        End If
                    Case 8 '横向分类
                    Case 9 '统计项
                        lngStatistics = lngStatistics + 1
                        For i = mshBody.FixedRows - 1 To Y
                            mshBody.TextMatrix(i, lngCurCols + IIF(lngGrid = 0, X, 0) + tmpItem.序号) = tmpItem.内容
                        Next
                End Select
            Next
            
            '-------------------------------------------------------------------------------------
            '处理表格数据
            '-------------------------------------------------------------------------------------
            If mLibDatas("_" & .内容).DataSet.RecordCount > 0 Then
                Set rsGroup = Nothing
                Set rsGroup = mLibDatas("_" & .内容).DataSet.Clone
                
                '1.生成表格框架
                
                '1.1:构造纵向分类表头对象(隐含排序字段)
                If lngGrid = 0 Then
                    Set rsVsc = Nothing
                    Set rsVsc = New ADODB.Recordset
                    '1.1.1:先添加纵向排序字段
                    For i = 0 To UBound(Split(strVscOrder, "|"))
                        If Left(Split(strVscOrder, "|")(i), 1) = "," Then
                            With rsGroup.Fields(Mid(Split(strVscOrder, "|")(i), 2))
                                '！！！adNumeric类型有时不行,可替换为adBigInt或adSingle/adDouble
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        Else
                            With rsGroup.Fields(Split(strVscOrder, "|")(i))
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    '1.1.2:再添加纵向分类字段(不与排序字段重复)
                    For i = 0 To UBound(Split(strVsc, "|"))
                        If InStr("|" & Replace(strVscOrder, ",", "") & "|", "|" & Split(strVsc, "|")(i) & "|") = 0 Then
                            With rsGroup.Fields(Split(strVsc, "|")(i))
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    rsVsc.CursorLocation = adUseClient
                    rsVsc.LockType = adLockBatchOptimistic
                    rsVsc.CursorType = adOpenStatic
                    rsVsc.Open
                End If
                
                '1.2:构造横向分类表头对象(隐含排序字段)
                Set rsHsc = Nothing
                If strHsc <> "" Then
                    Set rsHsc = New ADODB.Recordset
                    '1.2.1:先添加横向排序字段
                    For i = 0 To UBound(Split(strHscOrder, "|"))
                        If Left(Split(strHscOrder, "|")(i), 1) = "," Then
                            With rsGroup.Fields(Mid(Split(strHscOrder, "|")(i), 2))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        Else
                            With rsGroup.Fields(Split(strHscOrder, "|")(i))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    '1.2.2:再添加横向分类字段(不与排序字段重复)
                    For i = 0 To UBound(Split(strHsc, "|"))
                        If InStr("|" & Replace(strHscOrder, ",", "") & "|", "|" & Split(strHsc, "|")(i) & "|") = 0 Then
                            With rsGroup.Fields(Split(strHsc, "|")(i))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    rsHsc.CursorLocation = adUseClient
                    rsHsc.LockType = adLockBatchOptimistic
                    rsHsc.CursorType = adOpenStatic
                    rsHsc.Open
                End If
                
                '1.3:添加表头数据集
                rsGroup.MoveFirst
                For i = 1 To rsGroup.RecordCount
                    '纵向表头
                    If Not rsVsc Is Nothing And lngGrid = 0 Then
                        strFilter = "" '是否已经加入该分类组值
                        For j = 0 To UBound(Split(strVsc, "|")) '此时不管排序字段
                            strFilter = strFilter & " And " & Split(strVsc, "|")(j) & "="
                            Select Case rsGroup.Fields(Split(strVsc, "|")(j)).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "'" & Replace(rsGroup.Fields(Split(strVsc, "|")(j)).Value, " ", "♂♂") & "'"
                                    Else
                                        strFilter = strFilter & "'#'"
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        strFilter = strFilter & rsGroup.Fields(Split(strVsc, "|")(j)).Value
                                    Else
                                        strFilter = strFilter & "123456707654321"
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        '必须格式化以正确识别格式,如#02-4-9#会认成"2009-02-04"
                                        strFilter = strFilter & "#" & Format(rsGroup.Fields(Split(strVsc, "|")(j)).Value, "yyyy-MM-dd HH:mm:ss") & "#"
                                    Else
                                        strFilter = strFilter & "#3000-05-05#"
                                    End If
                            End Select
                        Next
                        rsVsc.Filter = Replace(Mid(strFilter, 6), "♂♂", " ")
                        If rsVsc.EOF Then
                            rsVsc.AddNew
                            For j = 0 To rsVsc.Fields.count - 1 '加入新的分类组值
                                If Not IsNull(rsGroup.Fields(rsVsc.Fields(j).name).Value) Then
                                    rsVsc.Fields(j).Value = rsGroup.Fields(rsVsc.Fields(j).name).Value
                                Else
                                    Select Case rsGroup.Fields(rsVsc.Fields(j).name).type
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            rsVsc.Fields(j).Value = "#" '空标志
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            rsVsc.Fields(j).Value = 123456707654321# '空标志
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            rsVsc.Fields(j).Value = #5/5/3000#   '空标志
                                    End Select
                                End If
                            Next
                        End If
                    End If
                    '横向表头
                    If Not rsHsc Is Nothing Then
                        strFilter = "" '是否已经加入该分类组值
                        For j = 0 To UBound(Split(strHsc, "|")) '此时不管排序字段
                            strFilter = strFilter & " And " & Split(strHsc, "|")(j) & "="
                            Select Case rsGroup.Fields(Split(strHsc, "|")(j)).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "'" & rsGroup.Fields(Split(strHsc, "|")(j)).Value & "'"
                                    Else
                                        strFilter = strFilter & "'#'"
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & rsGroup.Fields(Split(strHsc, "|")(j)).Value
                                    Else
                                        strFilter = strFilter & "123456707654321"
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "#" & Format(rsGroup.Fields(Split(strHsc, "|")(j)).Value, "yyyy-MM-dd HH:mm:ss") & "#"
                                    Else
                                        strFilter = strFilter & "#3000-05-05#"
                                    End If
                            End Select
                        Next
                        rsHsc.Filter = Mid(strFilter, 6)
                        If rsHsc.EOF Then
                            rsHsc.AddNew
                            For j = 0 To rsHsc.Fields.count - 1 '加入新的分类组值
                                If Not IsNull(rsGroup.Fields(rsHsc.Fields(j).name).Value) Then
                                    rsHsc.Fields(j).Value = rsGroup.Fields(rsHsc.Fields(j).name).Value
                                Else
                                    Select Case rsGroup.Fields(rsHsc.Fields(j).name).type
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            rsHsc.Fields(j).Value = "#" '空标志
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            rsHsc.Fields(j).Value = 123456707654321# '空标志
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            rsHsc.Fields(j).Value = #5/5/3000#   '空标志
                                    End Select
                                End If
                            Next
                        End If
                    End If
                    rsGroup.MoveNext
                Next
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    rsVsc.UpdateBatch adAffectAllChapters
                    rsVsc.Filter = 0
                End If
                If Not rsHsc Is Nothing Then
                    rsHsc.UpdateBatch adAffectAllChapters
                    rsHsc.Filter = 0
                End If
                
                '1.4:表头内容按定义排序
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    strSort = ""
                    For i = 0 To UBound(Split(strVscOrder, "|"))
                        If Left(Split(strVscOrder, "|")(i), 1) = "," Then
                            strSort = strSort & "," & Mid(Split(strVscOrder, "|")(i), 2) & " Desc"
                        Else
                            strSort = strSort & "," & Split(strVscOrder, "|")(i)
                        End If
                    Next
                    If strSort <> "" Then rsVsc.Sort = Mid(strSort, 2)
                    rsVsc.MoveFirst
                End If
                If Not rsHsc Is Nothing Then
                    strSort = ""
                    For i = 0 To UBound(Split(strHscOrder, "|"))
                        If Left(Split(strHscOrder, "|")(i), 1) = "," Then
                            strSort = strSort & "," & Mid(Split(strHscOrder, "|")(i), 2) & " Desc"
                        Else
                            strSort = strSort & "," & Split(strHscOrder, "|")(i)
                        End If
                    Next
                    If strSort <> "" Then rsHsc.Sort = Mid(strSort, 2)
                    rsHsc.MoveFirst
                End If
                
                '1.5:填写横纵向表头单元数据
                '纵向表头
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    Set colVsc = New Collection
                    '汇总函数
                    strVscStat = ""
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.id)
                        If tmpItem.类型 = 7 Then strVscStat = strVscStat & "," & tmpItem.汇总
                    Next
                    strVscStat = Mid(strVscStat, 2)
                    
                    '产生表头
                    k = Y  '当前应该处理的行
                    ReDim arrLevel(X - 1) '用判断某级汇总是否应该加入汇总行了
                    ReDim arrMerge(X - 1) '用于处理同级汇总不同上级防止合并
                    For i = 1 To X - 1
                        arrMerge(i) = Space(i Mod 2)
                    Next
                    For i = 1 To rsVsc.RecordCount
                        k = k + 1
                        If mshBody.Rows - 1 < k Then mshBody.Rows = mshBody.Rows + 1
                        strKey = ""
                        For j = 0 To X - 1
                            strTmp = Trim(mshBody.TextMatrix(Y, j))
                            Select Case rsVsc.Fields(strTmp).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If rsVsc.Fields(strTmp).Value = "#" Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " " '用空格是为了强制合并
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "♂")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If rsVsc.Fields(strTmp).Value = 123456707654321# Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " "
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "♂")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If rsVsc.Fields(strTmp).Value = #5/5/3000# Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " "
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "♂")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                            End Select
                        Next
                        
                        '插入汇总行(在倒数第二行)
                        For j = X - 1 To 1 Step -1 '一定要反向
                            strTmp = GetRowText(mshBody, k, j - 1)
                            If strTmp <> arrLevel(j) And k > Y + 1 Then
                                If strVscStat <> "" Then
                                    If Split(strVscStat, ",")(j) <> "" Then
                                        mshBody.AddItem "", k
                                        mshBody.Row = k
                                        For l = 0 To j - 1
                                            mshBody.TextMatrix(k, l) = mshBody.TextMatrix(k - 1, l)
                                        Next
                                        For l = j To X - 1
                                            mshBody.Col = l
                                            mshBody.CellAlignment = 4
                                            'mshBody.TextMatrix(k, L) = Space(j Mod 2) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j Mod 2)
                                            mshBody.TextMatrix(k, l) = Space(j) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j)
                                        Next
                                        mshBody.RowData(k) = j + 1
                                        mshBody.MergeRow(k) = True
                                        
                                        k = k + 1
                                    End If
                                End If
                                arrMerge(j) = IIF(arrMerge(j) = "", " ", "")
                            End If
                        Next
                        
                        '注意：k要为插入汇总行之后的行(因为汇总行不是插入到最后行)
                        colVsc.Add k, "_" & Mid(strKey, 2) '纵向数据行定位集合
                        
                        '此时K为非汇总行(最后一行)
                        For j = 1 To X - 1
                            mshBody.TextMatrix(k, j) = mshBody.TextMatrix(k, j) & arrMerge(j)
                            arrLevel(j) = GetRowText(mshBody, k, j - 1)
                        Next
                        
                        rsVsc.MoveNext
                    Next
                    
                    '加入最后的汇总行
                    k = mshBody.Rows
                    If strVscStat <> "" And k > Y + 1 Then
                        For j = X - 1 To 0 Step -1
                            If Split(strVscStat, ",")(j) <> "" Then
                                mshBody.AddItem "", k
                                mshBody.Row = k
                                For l = 0 To j - 1
                                    mshBody.TextMatrix(k, l) = mshBody.TextMatrix(k - 1, l)
                                Next
                                For l = j To X - 1
                                    mshBody.Col = l
                                    mshBody.CellAlignment = 4
                                    '对于汇总行：0,2列合计,1列不合计,这时0,2的合计挨在一起了,造成横纵向同时合并。
                                    '汇总列因为显示方式有点不同，所以不存在这个问题。
                                    'mshBody.TextMatrix(k, L) = Space(j Mod 2) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j Mod 2)
                                    mshBody.TextMatrix(k, l) = Space(j) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j)
                                Next
                                mshBody.RowData(k) = j + 1
                                mshBody.MergeRow(k) = True
                                
                                k = k + 1
                            End If
                        Next
                    End If
                End If
                
                '横向表头
                If Y > 0 And Not rsHsc Is Nothing Then
                    Set colHsc = New Collection
                    '汇总函数
                    strHscStat = ""
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.id)
                        If tmpItem.类型 = 8 Then strHscStat = strHscStat & "," & tmpItem.汇总
                    Next
                    strHscStat = Mid(strHscStat, 2)
                    
                    '产生表头
                    ReDim arrLevel(Y - 1) '用判断某级汇总是否应该加入汇总列了
                    ReDim arrMerge(Y - 1) '用于处理同级汇总不同上级防止合并
                    For i = 1 To Y - 1
                        arrMerge(i) = Space(i Mod 2)
                    Next
                    l = lngCurCols + IIF(lngGrid = 0, X, 0) - Z    '当前应该处理的列
                    For i = 1 To rsHsc.RecordCount
                        l = l + Z
                        If mshBody.Cols - 1 < l Then mshBody.Cols = mshBody.Cols + Z
                        strKey = "" '之所以在这里处理(与纵向表头不同),是为了取记录原始值
                        For j = 0 To Y - 1
                            For k = 0 To Z - 1
                                Select Case rsHsc.Fields(CStr(Split(strHsc, "|")(j))).type
                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = "#" Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, l + k) = " " '用空格是为了强制合并
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "♂")
                                            mshBody.TextMatrix(j, l + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = 123456707654321# Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, l + k) = " "
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "♂")
                                            mshBody.TextMatrix(j, l + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = #5/5/3000# Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, l + k) = " "
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "♂")
                                            mshBody.TextMatrix(j, l + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                End Select
                            Next
                        Next
                        
                        '加入汇总列
                        For j = Y - 1 To 1 Step -1 '一定要反向
                            strTmp = GetColText(mshBody, j - 1, l)
                            If strTmp <> arrLevel(j) And l > lngCurCols + IIF(lngGrid = 0, X, 0) Then
                                If strHscStat <> "" Then
                                    If Split(strHscStat, ",")(j) <> "" Then
                                        AddCol mshBody, l, Z
                                        For k = 0 To Z - 1
                                            For M = 0 To j - 1
                                                mshBody.TextMatrix(M, l + k) = mshBody.TextMatrix(M, l + k - Z)
                                            Next
                                            mshBody.Col = l + k
                                            mshBody.Row = j
                                            mshBody.CellAlignment = 4
                                            mshBody.TextMatrix(j, l + k) = Space((j + 1) Mod 2) & GetStatText(CStr(Split(strHscStat, ",")(j))) & Space((j + 1) Mod 2)
                                            mshBody.ColData(l + k) = j + 1
                                            mshBody.MergeCol(l + k) = True
                                        Next
                                        l = l + Z
                                    End If
                                End If
                                arrMerge(j) = IIF(arrMerge(j) = "", " ", "")
                            End If
                        Next
                        
                        '注意：L要为插入汇总列之后的列(因为汇总列不是插入到最后列)
                        colHsc.Add l, "_" & Mid(strKey, 2) '横向数据行定位集合
                        
                        '此时L为非汇总列(最后一组列)
                        For j = 1 To Y - 1
                            For k = 0 To Z - 1
                                mshBody.TextMatrix(j, l + k) = mshBody.TextMatrix(j, l + k) & arrMerge(j)
                            Next
                            arrLevel(j) = GetColText(mshBody, j - 1, l)
                        Next
                        rsHsc.MoveNext
                    Next
                    '加入最后的汇总列
                    l = mshBody.Cols
                    If strHscStat <> "" And l > lngCurCols + IIF(lngGrid = 0, X, 0) Then
                        For j = Y - 1 To 0 Step -1
                            If Split(strHscStat, ",")(j) <> "" Then
                                AddCol mshBody, l, Z
                                For k = 0 To Z - 1
                                    For M = 0 To j - 1
                                        mshBody.TextMatrix(M, l + k) = mshBody.TextMatrix(M, l + k - Z)
                                    Next
                                    mshBody.Col = l + k
                                    mshBody.Row = j
                                    mshBody.CellAlignment = 4
                                    mshBody.TextMatrix(j, l + k) = Space((j + 1) Mod 2) & GetStatText(CStr(Split(strHscStat, ",")(j))) & Space((j + 1) Mod 2)
                                    mshBody.ColData(l + k) = j + 1
                                    mshBody.MergeCol(l + k) = True
                                Next
                                l = l + Z
                            End If
                        Next
                    End If
                End If
                
                '填写统计项表头
                strFormat = ""
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.id)
                    If tmpItem.类型 = 9 Then
                        strFormat = strFormat & "|~" & tmpItem.序号 & "~"
                        For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                            '如果当前表格横向层次比前一个表格少,则对补加的固定行补填行内容
                            For j = Y To mshBody.FixedRows - 1
                                mshBody.TextMatrix(j, tmpItem.序号 + i) = Space((tmpItem.序号 + i) Mod 2) & tmpItem.内容 & Space((tmpItem.序号 + i) Mod 2)
                            Next
                            '合计列内容
                            If mshBody.ColData(i) > 0 Then
                                For M = Y - 1 To mshBody.ColData(i) Step -1
                                    For k = 0 To Z - 1
                                        mshBody.TextMatrix(M, i + k) = mshBody.TextMatrix(M + 1, i + k)
                                    Next
                                Next
                            End If
                        Next
                    End If
                Next
                strFormat = Mid(strFormat, 2)
                
                '统计项字段格式
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.id)
                    '间隔字符不能为格式字符
                    If tmpItem.类型 = 9 Then strFormat = Replace(strFormat, "~" & tmpItem.序号 & "~", tmpItem.格式)
                Next
                
                '列宽(含汇总列)、对齐
                strAlign = ""
                Set colTmp = New Collection
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.id)
                    Select Case tmpItem.类型
                        Case 7 '纵向分类
                           If lngGrid = 0 Then mshBody.ColWidth(tmpItem.序号) = tmpItem.W
                        Case 9 '统计项
                            strAlign = strAlign & "," & tmpItem.对齐
                            For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                                mshBody.ColAlignment(i + tmpItem.序号) = Switch(tmpItem.对齐 = 0, 1, tmpItem.对齐 = 1, 4, tmpItem.对齐 = 2, 7)
                                If mshBody.FixedRows - 1 >= 0 And mshBody.Rows - 1 >= 0 Then mshBody.Cell(flexcpAlignment, mshBody.FixedRows - 1, i + tmpItem.序号, mshBody.Rows - 1, i + tmpItem.序号) = mshBody.ColAlignment(i + tmpItem.序号)
                                mshBody.ColWidth(i + tmpItem.序号) = tmpItem.W
                            Next
                            
                            '保存交叉统计的列对象
                            colTmp.Add tmpItem, "_" & tmpItem.序号
                    End Select
                Next
                strAlign = Mid(strAlign, 2)
                
                '处理表体数据
                rsGroup.MoveFirst
                For i = 1 To rsGroup.RecordCount
                    '取行
                    strKey = ""
                    For j = 0 To UBound(Split(strVsc, "|"))
                        strKey = strKey & "^" & IIF(IsNull(rsGroup.Fields(CStr(Split(strVsc, "|")(j))).Value), "", Replace(Nvl(rsGroup.Fields(CStr(Split(strVsc, "|")(j))).Value, ""), " ", "♂"))
                    Next
                    
                    '左联连表的时候,如果右边表的数据比左边多,因为以左为准,如果不能定位行,则不处理数据。
                    lngRow = 0
                    If lngGrid > 0 Then On Local Error Resume Next
                    lngRow = CLng(colVsc("_" & Mid(strKey, 2))) + lngDiff
                    On Error GoTo 0
                    If lngRow > 0 Then
                        '取列
                        lngCol = lngCurCols + IIF(lngGrid = 0, X, 0)
                        If strHsc <> "" Then
                            strKey = ""
                            For j = 0 To UBound(Split(strHsc, "|"))
                                strKey = strKey & "^" & IIF(IsNull(rsGroup.Fields(CStr(Split(strHsc, "|")(j))).Value), "", Replace(rsGroup.Fields(CStr(Split(strHsc, "|")(j))).Value & "", " ", "♂"))
                            Next
                            lngCol = CLng(colHsc("_" & Mid(strKey, 2)))
                        End If
                        
                        '填数(暂不处理汇总行列)
                        For j = 0 To Z - 1
                            strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + j))
                            If Not IsNull(rsGroup.Fields(strTmp).Value) Then
                                '填数时格式化
                                StrFmt = ""
                                If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(j))
                                If StrFmt <> "" Then
                                    On Local Error Resume Next
                                    Select Case rsGroup.Fields(strTmp).type
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Format(Val(Replace(mshBody.TextMatrix(lngRow, lngCol + j), ",", "")) + rsGroup.Fields(strTmp).Value, StrFmt)
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Format(Val(mshBody.TextMatrix(lngRow, lngCol + j)) + Val(rsGroup.Fields(strTmp).Value), StrFmt)
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            If mshBody.TextMatrix(lngRow, lngCol + j) = "" Then
                                                mshBody.TextMatrix(lngRow, lngCol + j) = Format(CDate(rsGroup.Fields(strTmp).Value), StrFmt)
                                            Else
                                                mshBody.TextMatrix(lngRow, lngCol + j) = Format(CDate(mshBody.TextMatrix(lngRow, lngCol + j)) + rsGroup.Fields(strTmp).Value, StrFmt)
                                            End If
                                    End Select
                                    On Local Error GoTo 0
                                Else
                                    Select Case rsGroup.Fields(strTmp).type
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Val(Replace(mshBody.TextMatrix(lngRow, lngCol + j), ",", "")) + rsGroup.Fields(strTmp).Value
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Val(mshBody.TextMatrix(lngRow, lngCol + j)) + Val(rsGroup.Fields(strTmp).Value)
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            If mshBody.TextMatrix(lngRow, lngCol + j) = "" Then
                                                mshBody.TextMatrix(lngRow, lngCol + j) = CDate(rsGroup.Fields(strTmp).Value)
                                            Else
                                                mshBody.TextMatrix(lngRow, lngCol + j) = CDate(mshBody.TextMatrix(lngRow, lngCol + j)) + rsGroup.Fields(strTmp).Value
                                            End If
                                    End Select
                                End If
                                
                                '根据对齐设置防止合并
                                Select Case CByte(Split(strAlign, ",")(j))
                                    Case 0 '左对齐
                                        mshBody.TextMatrix(lngRow, lngCol + j) = mshBody.TextMatrix(lngRow, lngCol + j) & Space((lngRow + lngCol + j) Mod 2)
                                    Case 1 '中对齐
                                        mshBody.TextMatrix(lngRow, lngCol + j) = Space((lngRow + lngCol + j) Mod 2) & mshBody.TextMatrix(lngRow, lngCol + j) & Space((lngRow + lngCol + j) Mod 2)
                                    Case 2 '右对齐
                                        mshBody.TextMatrix(lngRow, lngCol + j) = Space((lngRow + lngCol + j) Mod 2) & mshBody.TextMatrix(lngRow, lngCol + j)
                                End Select
                                
                                '数据体当前列对象
                                Set objStatusGridItem = Nothing
                                For Each objStatusGridItem In colTmp
                                    If objStatusGridItem.序号 = j And objStatusGridItem.类型 = Val("9-数据区对象") Then
                                        Exit For
                                    End If
                                Next
                                
                                If Not objStatusGridItem Is Nothing Then
                                    For k = 1 To objStatusGridItem.ColProtertys.count
                                        Set objColProp = objStatusGridItem.ColProtertys.Item(k)
                                        If InStr(objColProp.条件值, objCurItem.内容 & ".") > 0 Then
                                            varIFValue = GetStatGridData(mshBody.Index, objColProp.条件值, lngRow, lngCol + j)
                                        Else
                                            varIFValue = objColProp.条件值
                                        End If
                                        If lngCol + j = mshBody.FixedCols And objColProp.是否整行应用 Then
                                            If CheckColProtertys(Trim(mshBody.TextMatrix(lngRow, lngCol + j)), objColProp.条件关系, varIFValue) Then
                                                If objColProp.背景颜色 <> vbWhite Then
                                                    mshBody.Cell(flexcpBackColor, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.背景颜色
                                                End If
                                                If objColProp.字体颜色 <> vbBlack Then
                                                    mshBody.Cell(flexcpForeColor, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.字体颜色
                                                End If
                                                If objColProp.是否加粗 Then
                                                    mshBody.Cell(flexcpFontBold, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.是否加粗
                                                End If
                                            End If
                                        Else
                                            If CheckColProtertys(Trim(mshBody.TextMatrix(lngRow, lngCol + j)), objColProp.条件关系, varIFValue) Then
                                                If objColProp.背景颜色 <> vbWhite Then
                                                    mshBody.Cell(flexcpBackColor, lngRow, lngCol + j) = objColProp.背景颜色
                                                End If
                                                If objColProp.字体颜色 <> vbBlack Then
                                                    mshBody.Cell(flexcpForeColor, lngRow, lngCol + j) = objColProp.字体颜色
                                                End If
                                                If objColProp.是否加粗 Then
                                                    mshBody.Cell(flexcpFontBold, lngRow, lngCol + j) = objColProp.是否加粗
                                                End If
                                                
                                                '对齐方式
                                                Select Case objColProp.对齐
                                                Case Val("1-居左")
                                                    mshBody.Cell(flexcpAlignment, lngRow, lngCol + j) = flexAlignLeftCenter
                                                Case Val("2-居中")
                                                    mshBody.Cell(flexcpAlignment, lngRow, lngCol + j) = flexAlignCenterCenter
                                                Case Val("3-居右")
                                                    mshBody.Cell(flexcpAlignment, lngRow, lngCol + j) = flexAlignRightCenter
                                                Case Else
                                                    '缺省，不处理
                                                End Select
                                            End If
                                        End If
                                    Next
                                End If
                                
                            End If
                        Next
                    End If
                    rsGroup.MoveNext
                Next
                Set colTmp = Nothing
                
                '计算汇总行列数据(纵向优先)
                '横向汇总列
                If strHsc <> "" And strHscStat <> "" Then
                    For l = UBound(Split(strHsc, "|")) To 0 Step -1
                        strStat = CStr(Split(strHscStat, ",")(l))
                        If strStat <> "" Then
                            ReDim arrStat(mshBody.FixedRows To mshBody.Rows - 1, Z - 1)  '保存汇总数据
                            ReDim arrCount(mshBody.FixedRows To mshBody.Rows - 1, Z - 1) '保存非空记录个数
                            blnDo = False
                            For j = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                                For i = mshBody.FixedRows To mshBody.Rows - 1 '因为可能多个表格左联接,Y不准,用FixedRows
                                    '显示汇总行结果
                                    If mshBody.ColData(j) = l + 1 Then
                                        For k = 0 To Z - 1
                                            If strStat = "AVG" Then
                                                strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + k))
                                                Select Case rsGroup.Fields(strTmp).type
                                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                        arrStat(i, k) = Val(arrStat(i, k) / arrCount(i, k))
                                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                        arrStat(i, k) = Val(arrStat(i, k) / arrCount(i, k))
                                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                        arrStat(i, k) = CDate(arrStat(i, k) / arrCount(i, k))
                                                End Select
                                            End If
                                            StrFmt = ""
                                            If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(k))
                                            If StrFmt <> "" Then
                                                On Local Error Resume Next
                                                Select Case Split(strAlign, ",")(k)
                                                    Case 0 '左
                                                        mshBody.TextMatrix(i, j + k) = Format(arrStat(i, k), StrFmt) & Space((i + j + k) Mod 2)
                                                    Case 1 '中
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & Format(arrStat(i, k), StrFmt) & Space((i + j + k) Mod 2)
                                                    Case 2 '右
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & Format(arrStat(i, k), StrFmt)
                                                End Select
                                                On Local Error GoTo 0
                                            Else
                                                Select Case Split(strAlign, ",")(k)
                                                    Case 0 '左
                                                        mshBody.TextMatrix(i, j + k) = arrStat(i, k) & Space((i + j + k) Mod 2)
                                                    Case 1 '中
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & arrStat(i, k) & Space((i + j + k) Mod 2)
                                                    Case 2 '右
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & arrStat(i, k)
                                                End Select
                                            End If
                                        Next
                                    '计算汇总数据
                                    ElseIf mshBody.ColData(j) = 0 Then
                                        For k = 0 To Z - 1
                                            If Trim(mshBody.TextMatrix(i, j + k)) <> "" Then
                                                strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + k))
                                                arrCount(i, k) = arrCount(i, k) + 1
                                                Select Case strStat
                                                    Case "SUM", "AVG"
                                                        Select Case rsGroup.Fields(strTmp).type
                                                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                                arrStat(i, k) = arrStat(i, k) + Val(Replace(Trim(mshBody.TextMatrix(i, j + k)), ",", ""))
                                                            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                                arrStat(i, k) = arrStat(i, k) + Val(Trim(mshBody.TextMatrix(i, j + k)))
                                                            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                                arrStat(i, k) = arrStat(i, k) + CDate(Trim(mshBody.TextMatrix(i, j + k)))
                                                        End Select
                                                    Case "MIN"
                                                        If Not blnDo Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k)): blnDo = True
                                                        If Trim(mshBody.TextMatrix(i, j + k)) < arrStat(i, k) Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k))
                                                    Case "MAX"
                                                        If Not blnDo Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k)): blnDo = True
                                                        If Trim(mshBody.TextMatrix(i, j + k)) > arrStat(i, k) Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k))
                                                    Case "COUNT"
                                                        arrStat(i, k) = arrStat(i, k) + 1
                                                End Select
                                            End If
                                        Next
                                    End If
                                Next
                                If mshBody.ColData(j) = l + 1 Then
                                    ReDim arrStat(mshBody.FixedRows To mshBody.Rows - 1, Z - 1)  '保存汇总数据
                                    ReDim arrCount(mshBody.FixedRows To mshBody.Rows - 1, Z - 1) '保存非空记录个数
                                    blnDo = False
                                End If
                            Next
                        End If
                    Next
                End If

                '纵向汇总行
                If strVscStat <> "" Then
                    For l = UBound(Split(strVsc, "|")) To 0 Step -1
                        strStat = CStr(Split(strVscStat, ",")(l))
                        If strStat <> "" Then
                            ReDim arrStat(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1) '保存汇总数据
                            ReDim arrCount(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1) '保存非空记录个数
                            blnDo = False
                            For i = mshBody.FixedRows To mshBody.Rows - 1 '因为可能多个表格左联接,Y不准,用FixedRows
                                For j = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1
                                    '显示汇总行结果
                                    If mshBody.RowData(i) = l + 1 Then
                                        If strStat = "AVG" Then
                                            strTmp = Trim(mshBody.TextMatrix(Y, j))
                                            Select Case rsGroup.Fields(strTmp).type
                                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                    arrStat(j) = Val(arrStat(j) / arrCount(j))
                                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                    arrStat(j) = Val(arrStat(j) / arrCount(j))
                                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                    arrStat(j) = CDate(arrStat(j) / arrCount(j))
                                            End Select
                                        End If
                                        k = 0
                                        If Z > 1 Then k = ((j - (lngCurCols + IIF(lngGrid = 0, X, 0)) + 1) Mod Z) - 1
                                        If k = -1 Then k = Z - 1
                                        StrFmt = ""
                                        If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(k))
                                        If StrFmt <> "" Then
                                            On Local Error Resume Next
                                            Select Case Split(strAlign, ",")(k)
                                                Case 0 '左
                                                    mshBody.TextMatrix(i, j) = Format(arrStat(j), StrFmt) & Space((i + j) Mod 2)
                                                Case 1 '中
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & Format(arrStat(j), StrFmt) & Space((i + j) Mod 2)
                                                Case 2 '右
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & Format(arrStat(j), StrFmt)
                                            End Select
                                            On Local Error GoTo 0
                                        Else
                                            Select Case Split(strAlign, ",")(k)
                                                Case 0 '左
                                                    mshBody.TextMatrix(i, j) = arrStat(j) & Space((i + j) Mod 2)
                                                Case 1 '中
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & arrStat(j) & Space((i + j) Mod 2)
                                                Case 2 '右
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & arrStat(j)
                                            End Select
                                        End If
                                    '计算汇总数据
                                    ElseIf mshBody.RowData(i) = 0 And Trim(mshBody.TextMatrix(i, j)) <> "" Then
                                        strTmp = Trim(mshBody.TextMatrix(Y, j))
                                        arrCount(j) = arrCount(j) + 1
                                        Select Case strStat
                                            Case "SUM", "AVG"
                                                Select Case rsGroup.Fields(strTmp).type
                                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                        arrStat(j) = arrStat(j) + Val(Replace(Trim(mshBody.TextMatrix(i, j)), ",", ""))
                                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                        arrStat(j) = arrStat(j) + Val(Trim(mshBody.TextMatrix(i, j)))
                                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                        arrStat(j) = arrStat(j) + CDate(Trim(mshBody.TextMatrix(i, j)))
                                                End Select
                                            Case "MIN"
                                                If Not blnDo Then arrStat(j) = Trim(mshBody.TextMatrix(i, j)): blnDo = True
                                                If Trim(mshBody.TextMatrix(i, j)) < arrStat(j) Then arrStat(j) = Trim(mshBody.TextMatrix(i, j))
                                            Case "MAX"
                                                If Not blnDo Then arrStat(j) = Trim(mshBody.TextMatrix(i, j)): blnDo = True
                                                If Trim(mshBody.TextMatrix(i, j)) > arrStat(j) Then arrStat(j) = Trim(mshBody.TextMatrix(i, j))
                                            Case "COUNT"
                                                arrStat(j) = arrStat(j) + 1
                                        End Select
                                    End If
                                Next
                                If mshBody.RowData(i) = l + 1 Then
                                    ReDim arrStat(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1)
                                    ReDim arrCount(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1)
                                    blnDo = False
                                End If
                            Next
                        End If
                    Next
                End If
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.id)
                    Select Case tmpItem.类型
                        Case 7 '纵向分类
                            '报表关联查询超链接设置
                            If tmpItem.Relations.count > 0 Then
                                mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.序号, mshBody.Rows - 1, tmpItem.序号) = &HFF0001
                                mshBody.Cell(flexcpFontUnderline, mshBody.FixedRows, tmpItem.序号, mshBody.Rows - 1, tmpItem.序号) = True
                                mshBody.Cell(flexcpData, mshBody.FixedRows, tmpItem.序号, mshBody.Rows - 1, tmpItem.序号) = tmpItem
                            End If
                            '设置表格部分字体颜色和加粗（优先）
                            mshBody.Cell(flexcpFontBold, mshBody.FixedRows, tmpItem.序号, mshBody.Rows - 1, tmpItem.序号) = tmpItem.粗体
                            If tmpItem.前景 <> 0 Then mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.序号, mshBody.Rows - 1, tmpItem.序号) = tmpItem.前景
                        Case 8 '横向分类
                            If tmpItem.Relations.count > 0 Then
                                mshBody.Cell(flexcpForeColor, tmpItem.序号, mshBody.FixedCols, tmpItem.序号, mshBody.Cols - 1) = &HFF0001
                                mshBody.Cell(flexcpFontUnderline, tmpItem.序号, mshBody.FixedCols, tmpItem.序号, mshBody.Cols - 1) = True
                                mshBody.Cell(flexcpData, tmpItem.序号, mshBody.FixedCols, tmpItem.序号, mshBody.Cols - 1) = tmpItem
                            End If
                            '设置表格部分字体颜色和加粗（优先）
                            mshBody.Cell(flexcpFontBold, tmpItem.序号, mshBody.FixedCols, tmpItem.序号, mshBody.Cols - 1) = tmpItem.粗体
                            If tmpItem.前景 <> 0 Then mshBody.Cell(flexcpForeColor, tmpItem.序号, mshBody.FixedCols, tmpItem.序号, mshBody.Cols - 1) = tmpItem.前景
                        Case 9 '统计项
                            For j = mshBody.FixedCols To mshBody.Cols - 1 Step lngStatistics
                                On Error Resume Next
                                If tmpItem.Relations.count > 0 Then
                                    mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.序号 + j, mshBody.Rows - 1, tmpItem.序号 + j) = &HFF0001
                                    mshBody.Cell(flexcpFontUnderline, mshBody.FixedRows, tmpItem.序号 + j, mshBody.Rows - 1, tmpItem.序号 + j) = True
                                End If
                                '优化；只在特定行绑定链报表对象
                                mshBody.Cell(flexcpData, mshBody.FixedRows, tmpItem.序号 + j, mshBody.FixedRows, tmpItem.序号 + j) = tmpItem

'                                 '设置表格部分字体颜色和加粗（优先）
'                                mshBody.Cell(flexcpFontBold, mshBody.FixedRows, tmpItem.序号 + j, mshBody.Rows - 1, tmpItem.序号 + j) = tmpItem.粗体
'                                If tmpItem.前景 <> 0 Then mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.序号 + j, mshBody.Rows - 1, tmpItem.序号 + j) = tmpItem.前景
                                On Error GoTo 0
                            Next
                    End Select
                Next
                '去除函数行的超链接
                For i = 0 To mshBody.FixedCols - 1
                    For j = 0 To mshBody.Rows - 1
                        If Decode(Trim(mshBody.TextMatrix(j, i)), "合计", 1, "平均值", 2, "最大值", 3, "最小值", 4, "记录数", 5, 0) > 0 Then
                            mshBody.Cell(flexcpForeColor, j, i, j, mshBody.Cols - 1) = mshBody.ForeColor
                            mshBody.Cell(flexcpFontUnderline, j, i, j, mshBody.Cols - 1) = False
                            mshBody.Cell(flexcpData, j, i, j, mshBody.Cols - 1) = Empty
                            mshBody.Cell(flexcpFontBold, j, i, j, mshBody.Cols - 1) = False
                        End If
                    Next
                Next
                For j = 0 To mshBody.FixedRows - 1
                    For i = 0 To mshBody.Cols - 1
                        If Decode(Trim(mshBody.TextMatrix(j, i)), "合计", 1, "平均值", 2, "最大值", 3, "最小值", 4, "记录数", 5, 0) > 0 Then
                            mshBody.Cell(flexcpForeColor, j, i, mshBody.Rows - 1, i) = mshBody.ForeColor
                            mshBody.Cell(flexcpFontUnderline, j, i, mshBody.Rows - 1, i) = False
                            mshBody.Cell(flexcpData, j, i, mshBody.Rows - 1, i) = Empty
                            mshBody.Cell(flexcpFontBold, j, i, mshBody.Rows - 1, i) = False
                        End If
                    Next
                Next
                
                '处理单统计项表格式
                If Z = 1 And Y > 0 Then
                    For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1
                        For j = mshBody.FixedRows - 1 To Y Step -1
                            mshBody.TextMatrix(j, i) = mshBody.TextMatrix(Y - 1, i)
                        Next
                    Next
                Else
                    blnHide = False
                End If
                
                '独立表格的"表头"属性存放各个表格最后列数,用于"缺省列宽"功能
                objItem.表头 = objItem.表头 & "|" & .id & "," & mshBody.Cols
                strTopRow = strTopRow & "|" & objCurItem.内容 & "," & mshBody.Cols
                                
                '当前表格所处理到的列数
                lngCurCols = mshBody.Cols
            Else
                blnHide = False
                Call SetHeadCenter(mshBody)
                Exit For '如果主表格没有数据,则左联接的表格也不用处理了
            End If
        End With
    Next
    
    objItem.表头 = Mid(objItem.表头, 2)
    
    '多个表联接时,加一顶行(该段可取消,用SQL实现)
'    strTopRow = Mid(strTopRow, 2)
'    strTmp = ""
'    For i = 0 To UBound(Split(strTopRow, "|"))
'        If InStr(strTmp & "|", "|" & Split(Split(strTopRow, "|")(i), ",")(0) & "|") = 0 Then
'            strTmp = strTmp & "|" & Split(Split(strTopRow, "|")(i), ",")(0)
'        End If
'    Next
'    If UBound(Split(Mid(strTmp, 2), "|")) > 0 Then
'        '插入行
'        mshBody.AddItem "", mshBody.FixedRows
'        mshBody.FixedRows = mshBody.FixedRows + 1
'        For i = mshBody.FixedRows - 1 To 1 Step -1
'            For j = 0 To mshBody.Cols - 1
'                mshBody.TextMatrix(i, j) = mshBody.TextMatrix(i - 1, j)
'                mshBody.RowHeight(i) = mshBody.RowHeight(i - 1)
'                mshBody.RowData(i) = mshBody.RowData(i - 1)
'            Next
'        Next
'        mshBody.RowData(0) = 0
'        mshBody.RowHeight(0) = objItem.行高
'        mshBody.MergeRow(0) = True
'        For j = mshBody.FixedCols To mshBody.Cols - 1
'            mshBody.TextMatrix(0, j) = ""
'        Next
'
'        '填写内容
'        For i = 0 To UBound(Split(strTopRow, "|"))
'            If i = 0 Then
'                lngColB = mshBody.FixedCols
'            Else
'                lngColB = lngColE + 1
'            End If
'            lngColE = CLng(Split(Split(strTopRow, "|")(i), ",")(1)) - 1
'            For j = lngColB To lngColE
'                mshBody.TextMatrix(0, j) = CStr(Split(Split(strTopRow, "|")(i), ",")(0))
'            Next
'        Next
'    End If
    
    '固定行列合并
    For j = 0 To mshBody.Cols - 1
        mshBody.MergeCol(j) = True
    Next
    For i = 0 To mshBody.FixedRows - 2
        mshBody.MergeRow(i) = True
    Next
    
    
    '处理表头内容(单元对齐、行高、内容,列宽)
    For Each tmpID In objItem.SubIDs
        Set tmpItem = mobjReport.Items("_" & tmpID.id)
        arrHead = Split(tmpItem.表头, "|")
        For i = 0 To UBound(arrHead) '对齐^高度^内容
            mshBody.RowHeight(i) = CLng(Split(arrHead(i), "^")(1))
        Next
    Next
    
    '行高(含汇总行)
    For i = mshBody.FixedRows To mshBody.Rows - 1
        mshBody.RowHeight(i) = objItem.行高
    Next
    
    '隐藏单项统计项表头行
    '------此段可换为下段-----------------
    blnHide = True
    For i = mshBody.FixedRows - 1 To 1 Step -1
        For j = 0 To mshBody.Cols - 1
            If mshBody.TextMatrix(i, j) <> mshBody.TextMatrix(i - 1, j) Then
                blnHide = False: Exit For
            End If
        Next
        If blnHide Then
            mshBody.RowHeight(i) = 0
        Else
            Exit For
        End If
    Next
    '------此段可替换上段-----------------
    'If blnHide Then mshBody.RowHeight(mshBody.FixedRows - 1) = 0
    
    '固定行中对齐
    mshBody.Cell(flexcpAlignment, 0, 0, mshBody.FixedRows - 1, mshBody.Cols - 1) = flexAlignCenterCenter
    mshBody.Cell(flexcpAlignment, 0, 0, mshBody.Rows - 1, mshBody.FixedCols - 1) = flexAlignCenterCenter
    
    '固定列左对齐(非合计列)
    For i = mshBody.FixedRows To mshBody.Rows - 1
        If mshBody.RowData(i) = 0 Then
            mshBody.Row = i
            For j = 0 To mshBody.FixedCols - 1
                mshBody.Col = j
                mshBody.CellAlignment = 1
            Next
        End If
    Next
    
    mshBody.WordWrap = True
    
    mshBody.MergeCells = flexMergeFree
    mshBody.ScrollBars = flexScrollBarBoth
    mshBody.Row = mshBody.FixedRows
    mshBody.Col = mshBody.FixedCols
    mshBody.Redraw = flexRDBuffered
    mshBody.ZOrder
    mshBody.Visible = True
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Function GetGridColWidth(ByVal objGrid As Object _
    , Optional ByRef intPageLastCol As Integer _
    , Optional ByVal lngMaxWidth As Long = 0) As Long
'功能：获取一个表格各列宽度之和

    Dim i As Integer
    Dim lngW As Long, lngGridWidth As Long
    
    intPageLastCol = -1
    lngW = 0
    For i = 0 To objGrid.Cols - 1
        If lngMaxWidth > 0 Then
            If lngMaxWidth >= lngW + objGrid.ColWidth(i) Then
                lngW = lngW + objGrid.ColWidth(i)
                intPageLastCol = i
            Else
                Exit For
            End If
        Else
            lngW = lngW + objGrid.ColWidth(i)
        End If
    Next
    GetGridColWidth = lngW
End Function

Private Sub SetGridAlign(Optional ByVal bytMode As Byte = 0)
'功能：自动调整附加表格体列宽(以最宽表格为准对齐)
'说明：只处理附加表体,且按设计尺寸处理,只在报表显示前调用一次
    
    Dim tmpMsh As VSFlexGrid, tmpBody As VSFlexGrid
    Dim tmpItem As RPTItem
    Dim lngMaxW As Long, lngCurW As Long, lngWidth As Long, lngSum As Long
    Dim strIDs As String
    Dim i As Integer, j As Integer, intCurID As Integer, intCol As Integer
    Dim arrIDs As Variant
    Dim sngRate As Single

    If Not mobjReport.blnLoad Then Exit Sub

    On Error GoTo hErr

    For Each tmpItem In mobjReport.Items
        If tmpItem.格式号 = bytFormat _
            And (tmpItem.类型 = Val("4-自由表格") Or tmpItem.类型 = Val("5-汇总表格")) _
            And tmpItem.参照 = "" And tmpItem.性质 = 0 Then

            '判断是否存在附加表
            If GridHaveApp(tmpItem.id) Then
                '附加表（1..n）
                strIDs = GetGridAppIDs(tmpItem.名称)
                strIDs = tmpItem.id & "," & strIDs
                arrIDs = Split(strIDs, ",")

                '获取参考表格各列的总宽度（首页，不计算超页宽的其他列）
                On Error Resume Next
                Set tmpMsh = Nothing
                Set tmpMsh = msh(Val(arrIDs(0)))
                On Error GoTo hErr

                lngMaxW = -1
                If Not tmpMsh Is Nothing Then
                    If bytMode = Val("1-补齐") Then
                        GoSub makPro
                    End If
                    lngMaxW = GetGridColWidth(tmpMsh, , tmpItem.W)              '主表格首页的宽度
                    lngWidth = GetGridColWidth(tmpMsh)                          '表格各列的总宽度
                    
                    '有附加表格的主表格总列宽小于表格设计宽时，采用设计宽按比例调整汇总表格的列宽
                    If tmpItem.W > lngWidth Then
                        'lngMaxW调整为整除像素的缇值
                        lngMaxW = Round(tmpItem.W / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
                        sngRate = lngMaxW / IIF(lngWidth = 0, 1, lngWidth)
                        lngSum = 0
                        For j = 0 To tmpMsh.Cols - 1
                            '确保整除像素的缇值
                            tmpMsh.ColWidth(j) = Round(tmpMsh.ColWidth(j) * sngRate / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
                            If lngSum + tmpMsh.ColWidth(j) >= lngMaxW Then
                                '超表格设计宽度
                                tmpMsh.ColWidth(j) = lngMaxW - lngSum
                            End If
                            lngSum = lngSum + tmpMsh.ColWidth(j)
                        Next
                        
                        '自由表格的表头
                        If tmpItem.类型 = Val("4-自由表格") And Not tmpMsh Is Nothing Then
                            Set tmpMsh = msh(Val(tmpMsh.Tag))
                            If Not tmpMsh Is Nothing Then
                                For j = 0 To tmpMsh.Cols - 1
                                    '确保整除像素的缇值
                                    tmpMsh.ColWidth(j) = Round(tmpMsh.ColWidth(j) * sngRate / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
                                Next
                            End If
                        End If
                        
                        '调整总宽度
                        lngMaxW = lngSum
                    End If
                End If

                If lngMaxW > -1 Then
                    '调整附加表格最后一列的宽度与参考表格宽度对齐。附加表只能是自由表格，不允许汇总表格
                    For i = 1 To UBound(arrIDs)
                        intCurID = Val(arrIDs(i))
                        Set tmpBody = Nothing
                        If mobjReport.Items("_" & intCurID).类型 = Val("4-任意表") Then
                            Set tmpMsh = msh(CInt(msh(intCurID).Tag))           '表头
                            Set tmpBody = msh(intCurID)                         '表体
                            
                            tmpMsh.Redraw = False
                            lngCurW = GetGridColWidth(tmpMsh, intCol, lngMaxW)  '附加表格表头首页的宽度
                            
                            '附加表格参考主表格
                            If intCol > -1 Then
                                If intCol < tmpMsh.Cols - 1 Then
                                    '超页宽处理（平均列宽计算）
                                    For j = intCol + 1 To tmpMsh.Cols - 1
                                        tmpMsh.ColWidth(j) = (lngMaxW - lngCurW) \ (tmpMsh.Cols - 1 - intCol)
                                        tmpBody.ColWidth(j) = tmpMsh.ColWidth(j)
                                    Next
                                Else
                                    If lngMaxW >= lngCurW Then
                                        tmpMsh.ColWidth(tmpMsh.Cols - 1) = tmpMsh.ColWidth(tmpMsh.Cols - 1) + lngMaxW - lngCurW
                                        tmpBody.ColWidth(tmpBody.Cols - 1) = tmpMsh.ColWidth(tmpMsh.Cols - 1)
                                    Else
                                        Debug.Print ""
                                    End If
                                End If
                            End If
                            
                            tmpMsh.Redraw = True
                        End If
                    Next
                    Erase arrIDs
                End If
                
                '重整自由表格的行高
                If tmpItem.类型 = Val("4-任意表") Then
                    Call AdjustRowHight(tmpItem.id)
                End If
                
            ElseIf bytMode = Val("1-补齐") Then
                Set tmpMsh = msh(CInt(msh(tmpItem.id).Tag))     '表头
                GoSub makPro
                Set tmpMsh = msh(tmpItem.id)                    '表体
                GoSub makPro
            End If
        End If
    Next
    
    Exit Sub

hErr:
    Call ErrCenter
    Exit Sub
    
makPro:
    lngWidth = GetGridColWidth(tmpMsh)
    sngRate = (tmpMsh.Width - 300) / IIF(lngWidth = 0, 1, lngWidth)
    For j = 0 To tmpMsh.Cols - 1
        tmpMsh.ColWidth(j) = tmpMsh.ColWidth(j) * sngRate
    Next
    Return
End Sub

Private Function GetGridAppIDs(strName As String) As String
'功能：获取参照表格的附加表格的索引们
'参数：strName=参照名
    Dim tmpItem As RPTItem
    Dim strIDs As String
    
    For Each tmpItem In mobjReport.Items
        If tmpItem.格式号 = bytFormat And tmpItem.类型 = 4 _
            And tmpItem.性质 = 1 And tmpItem.参照 = strName Then
            strIDs = strIDs & "," & tmpItem.id
        End If
    Next
    GetGridAppIDs = Mid(strIDs, 2)
End Function

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txt(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
        '强行输入匹配
        If txt(Index).Text <> "" Then
            If cmd(Index).Enabled And cmd(Index).Visible Then
                blnMatch = True
                Call cmd_Click(Index)
            End If
            Cancel = True
        End If
    End If
End Sub

Private Sub ReplaceSysNo(objReport As Report)
    Dim i As Integer, j As Integer
    For i = 1 To objReport.Datas.count
        objReport.Datas(i).SQL = Replace(objReport.Datas(i).SQL, "[系统]", IIF(mlngSys <> 0, mlngSys, objReport.系统))
        For j = 1 To objReport.Datas(i).Pars.count
            objReport.Datas(i).Pars(j).明细SQL = Replace(objReport.Datas(i).Pars(j).明细SQL, "[系统]", IIF(mlngSys <> 0, mlngSys, objReport.系统))
            objReport.Datas(i).Pars(j).分类SQL = Replace(objReport.Datas(i).Pars(j).分类SQL, "[系统]", IIF(mlngSys <> 0, mlngSys, objReport.系统))
        Next
    Next
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Function GetStatGridData(ByVal intIndex As Integer, ByVal strFiled As String, ByVal lngRow As Long, ByVal lngCol As Long)
'功能：取汇总表对应纵向分类下指定行的其他字段的值
    Dim str纵向分类 As String
    Dim i As Long, strFiledTmp As String
    
    With msh(intIndex)
        strFiledTmp = Mid(strFiled, InStr(strFiled, ".") + 1)
        str纵向分类 = .TextMatrix(0, lngCol)
        
        For i = lngCol To .Cols - 1
            If str纵向分类 = .TextMatrix(0, i) And strFiledTmp = Trim(.TextMatrix(.FixedRows - 1, i)) Then
                GetStatGridData = .TextMatrix(lngRow, i)
                Exit Function
            End If
        Next
        
        For i = lngCol To .FixedCols Step -1
            If str纵向分类 = .TextMatrix(0, i) And strFiledTmp = Trim(.TextMatrix(.FixedRows - 1, i)) Then
                GetStatGridData = .TextMatrix(lngRow, i)
                Exit Function
            End If
        Next
        GetStatGridData = strFiled
    End With
End Function

Private Sub FindItem(ByVal strFind As String, Optional ByVal blnNext As Boolean)
'功能：查找界面上的关键字
'参数：blnNext=查找下一个
    Static lngindex As Long
    Static lngMshRow As Long
    Static lngMshcol As Long
    Static strFindLast As String
    Dim objControl As Object
    Dim blntmp As Boolean
    Dim i As Long, j As Long, k As Long
    
    If Trim(strFind) = "" Then Exit Sub
    If strFindLast <> strFind Then lngindex = 0
    strFindLast = strFind
    If lngCurInx <> 0 And lbl(lngCurInx).BackColor = CON_SETFOCES Then lbl(lngCurInx).BackColor = lngTmpColor
    For Each objControl In Me.Controls
        i = i + 1
        '只查找标签和表格
        If i >= lngindex Then
            If objControl.name = "lbl" Then
                If i > lngindex Then
                    If objControl.Caption Like "*" & strFind & "*" Then
                        lngCurInx = objControl.Index
                        lngTmpColor = objControl.BackColor
                        objControl.BackColor = CON_SETFOCES
                        lngindex = i
                        blntmp = True
                        Exit Sub
                    End If
                End If
            ElseIf objControl.name = "msh" Then
                If lngindex <> i Then lngMshRow = 0: lngMshcol = 0
                If lngMshRow < objControl.Rows - 1 Or lngMshcol < objControl.Cols - 1 Then
                    For j = objControl.FixedRows To objControl.Rows - 1
                        For k = objControl.FixedCols To objControl.Cols - 1
                            If j = lngMshRow And k > lngMshcol Or j > lngMshRow Then
                                If objControl.TextMatrix(j, k) Like "*" & strFind & "*" Then
                                    objControl.Row = j: objControl.Col = k
                                    objControl.ShowCell j, k
                                    lngindex = i
                                    blntmp = True
                                    lngMshRow = j: lngMshcol = k
                                    objControl.SetFocus
                                    Exit Sub
                                End If
                            End If
                            
                        Next
                    Next
                End If
            End If
        End If
    Next
    If blntmp = False Then
        If lngindex <> 0 Then
            MsgBox "已经全部查找完成。", vbInformation, App.Title
        Else
            MsgBox "没有查找到相关的文字。", vbInformation, App.Title
        End If
        lngindex = 0
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call FindItem(txtFind.Text)
    End If
End Sub

Private Sub vsfRelations_LostFocus()
    vsfRelations.Visible = False
End Sub

Private Sub vsfRelations_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strRelationReport As String
    If Button = 2 Then
        Exit Sub
    End If
    
    strRelationReport = vsfRelations.TextMatrix(vsfRelations.Row, vsfRelations.ColIndex("ID"))
    If strRelationReport <> "" Then
        mlngRelationReport = Val(strRelationReport)
        If mbytType = 0 Then
            Call msh_Click(mintGridIndex)
        ElseIf mbytType = 1 Then
            Call lbl_Click(mintLblIndex)
        End If
        mlngRelationReport = 0
        vsfRelations.Visible = False
    End If
End Sub

Private Sub LoadCondsMenu()
    Dim strSQL As String
    Dim i As Integer
    Dim rsPara As ADODB.Recordset
    Dim blnRetry As Boolean
    
    If mlngRPTID = 0 Then Exit Sub
    
    On Error GoTo hErr
    
    '删除条件菜单
    For i = mnuPop_Cond.count - 1 To 1 Step -1
        Unload mnuPop_Cond(i)
    Next
    
    blnRetry = True
    strSQL = "Select Distinct 条件号, 条件名称 From zlRptConds Where 报表ID=[1] Order by 条件号"
    Set rsPara = OpenSQLRecord(strSQL, "获取报表参数的保存条件", mlngRPTID)
    blnRetry = False
    
    With rsPara
        If .RecordCount = 0 Then
            mnuPop_Split1.Visible = False
            mnuPop_Del.Enabled = False
            mintCurCondID = 0
            mintCurMenuIndex = 0
        Else
            mnuPop_Split1.Visible = True
            mnuPop_Del.Enabled = mintCurCondID > 0
            Do While .EOF = False
                i = .AbsolutePosition
                Load mnuPop_Cond(i)
                mnuPop_Cond(i).Caption = Nvl(!条件名称) & "(&" & i & ")"
                mnuPop_Cond(i).Visible = True
                mnuPop_Cond(i).Tag = Nvl(!条件号, 0)
                
                If mintCurCondID = Nvl(!条件号, 0) Then
                    mnuPop_Cond(i).Checked = True
                Else
                    mnuPop_Cond(i).Checked = False
                End If
                
                .MoveNext
            Loop
        End If
        .Close
    End With
            
    mnuPop_Default.Checked = mintCurCondID = 0
    
    Exit Sub
    
hErr:
    If blnRetry Then
        If ErrCenter = 1 Then Resume
    Else
        Call ErrCenter
    End If
End Sub

Public Function GetReportForm(objParent As Object, objCurDLL As clsReport, LibDatas As Object, arrPars As Variant, ByVal bytStyle As Byte) As Object
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    
    On Error Resume Next
        If Not LibDatas Is Nothing Then Set mLibDatas = LibDatas
        Load Me
    If Err.Number = 0 Then
        Set mobjfrmShowDock = New frmPreviewDock
        If Not mobjReport.blnLoad Then Exit Function
    
        If mobjReport.Items.count = 0 Then Exit Function
        
        If Not InitPrinter(Me) Then
            gblnError = True
            MsgBox "设备初始化失败.可能是系统没有安装打印机或与当前设置不兼容！", vbInformation, App.Title: Exit Function
        End If
        
        If Not CalcCellPage Then
            gblnError = True
            MsgBox "无法处理的表格格式,操作不能继续！", vbInformation, App.Title: Exit Function
        End If
        If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
            lbl(lngCurInx).BackColor = lngTmpColor
            lngCurInx = 0: lngTmpColor = 0
        End If
        mobjfrmShowDock.BorderStyle = FormBorderStyleConstants.vbBSNone '设置为无边框
        mobjfrmShowDock.Caption = mobjfrmShowDock.Caption       '重点是这一句
        Set mobjfrmShowDock.frmParent = Me
        Load mobjfrmShowDock
        mobjfrmShowDock.LoadForm 1
        Set LibDatas = mLibDatas
        Set GetReportForm = mobjfrmShowDock
    ElseIf Err.Number <> 0 Then
        '364:对象已卸载(在Form_Load内部Unload,如取消条件窗体)
        Err.Clear
    End If
End Function

Public Sub PrintReportForRec(objParent As Object, objCurDLL As clsReport, LibDatas As Object, arrPars As Variant, ByVal bytStyle As Byte)
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    
    On Error Resume Next
    
    If mbytStyle <> 0 Then
        Set mLibDatas = LibDatas
        Load Me
        If Err.Number = 0 Then
            If mbytStyle = 1 Then       '自动预览
                mnuFile_Preview_Click
            ElseIf mbytStyle = 2 Then   '自动打印
                mnuFile_Print_Click
            ElseIf mbytStyle = 3 Then   '输出到Excel
                mnuFile_Excel_Click
            ElseIf mbytStyle = 4 Then   '固定输出到PDF
                mnuFile_Print_Click
            End If
        ElseIf Err.Number <> 0 Then
            '364:对象已卸载(在Form_Load内部Unload,如取消条件窗体)
            Err.Clear
        End If
        Unload Me
    Else
        '先尝试以非模态显示报表
        If frmParent Is Nothing Then
            Me.Show
        ElseIf frmParent.name = "frmDesign" Then
            Me.Show 1, frmParent
        Else
            Me.Show , frmParent
        End If
        
        '两种情况不能以非模态显示
        If Err.Number = 373 Or Err.Number = 401 Then
            '373:不支持编译和设计环境部件的内部操作(源程序调用zlReport.dll,不支持加父窗体)
            '401:当打开有模式窗体时不能显示无模式窗体
            '已自动Load，再显示时不会再激活Form_Load事件
            Err.Clear: Me.Show 1
        ElseIf Err.Number = 364 Then
            '364:对象已卸载(在Form_Load内部Unload,如取消条件窗体)
            Err.Clear
        ElseIf Err.Number <> 0 Then
            Err.Clear: Unload Me '已自动Load，未知错误时卸载窗体
        End If
    End If
End Sub

Private Function SQLExistLOB(ByVal clsData As RPTData) As Boolean
'功能：判断数据源的SQL是否存在LOB字段类型
    
    Dim arrField As Variant
    Dim i As Integer, intType As Integer
    
    SQLExistLOB = False
    arrField = Split(clsData.字段, "|")
    For i = 0 To UBound(arrField)
        intType = Val(Split(arrField(i), ",")(1))
        Select Case intType
        Case adBinary, adVarBinary, adLongVarBinary
            SQLExistLOB = True
            Exit For
        End Select
    Next
End Function

