VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmTendFileStyle 
   Appearance      =   0  'Flat
   Caption         =   "护理文件样式"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12225
   Icon            =   "frmTendStyle.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12225
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSum 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   600
      ScaleHeight     =   1875
      ScaleWidth      =   6885
      TabIndex        =   132
      Top             =   4920
      Width           =   6885
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1020
         Left            =   435
         TabIndex        =   133
         Top             =   390
         Width           =   1995
         _cx             =   3519
         _cy             =   1799
         Appearance      =   2
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
   Begin VB.PictureBox picCloumn 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   270
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   2
      Top             =   1020
      Width           =   6885
      Begin VB.PictureBox picAlign 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4440
         ScaleHeight     =   315
         ScaleWidth      =   1005
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1005
         Begin VB.OptionButton optAlign 
            Height          =   315
            Index           =   0
            Left            =   0
            Picture         =   "frmTendStyle.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   0
            Width           =   345
         End
         Begin VB.OptionButton optAlign 
            Height          =   315
            Index           =   1
            Left            =   330
            Picture         =   "frmTendStyle.frx":0910
            Style           =   1  'Graphical
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   0
            Width           =   345
         End
         Begin VB.OptionButton optAlign 
            Height          =   315
            Index           =   2
            Left            =   660
            Picture         =   "frmTendStyle.frx":0CA0
            Style           =   1  'Graphical
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "对角线"
         Height          =   210
         Left            =   5610
         TabIndex        =   66
         Top             =   2985
         Width           =   1020
      End
      Begin VB.TextBox txtColumnPrefix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4425
         TabIndex        =   61
         Top             =   2175
         Width           =   2295
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "应用(&Y)"
         Height          =   300
         Index           =   2
         Left            =   2760
         TabIndex        =   96
         Top             =   1635
         Width           =   1100
      End
      Begin VB.TextBox txtColumnPostfix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4425
         TabIndex        =   63
         Top             =   2565
         Width           =   2295
      End
      Begin VB.ListBox lstColumnUsed 
         Height          =   1680
         Left            =   3990
         TabIndex        =   59
         Top             =   450
         Width           =   2730
      End
      Begin MSComCtl2.UpDown udColumnNo 
         Height          =   300
         Left            =   4605
         TabIndex        =   58
         Top             =   120
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtColumnNo"
         BuddyDispid     =   196618
         OrigLeft        =   5400
         OrigTop         =   75
         OrigRight       =   5640
         OrigBottom      =   375
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtColumnNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4215
         MaxLength       =   2
         TabIndex        =   57
         Text            =   "1"
         Top             =   120
         Width           =   405
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "删除(&E)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2760
         TabIndex        =   55
         Top             =   1185
         Width           =   1100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "选用(&S)"
         Height          =   300
         Index           =   0
         Left            =   2760
         TabIndex        =   54
         Top             =   885
         Width           =   1100
      End
      Begin VB.ListBox lstColumnItems 
         Height          =   2760
         Left            =   240
         TabIndex        =   53
         Top             =   465
         Width           =   2370
      End
      Begin VB.Label lbl列对齐 
         AutoSize        =   -1  'True
         Caption         =   "对齐"
         Height          =   180
         Left            =   3990
         TabIndex        =   64
         Top             =   3000
         Width           =   360
      End
      Begin VB.Label lblColumnPostfix 
         AutoSize        =   -1  'True
         Caption         =   "后缀"
         Height          =   180
         Left            =   3990
         TabIndex        =   62
         Top             =   2610
         Width           =   360
      End
      Begin VB.Label lblColumnPrefix 
         AutoSize        =   -1  'True
         Caption         =   "前缀"
         Height          =   180
         Left            =   3990
         TabIndex        =   60
         Top             =   2235
         Width           =   360
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         Caption         =   "第        列内容项目:"
         Height          =   180
         Left            =   4005
         TabIndex        =   56
         Top             =   180
         Width           =   1890
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         Caption         =   "可选护理记录项目:"
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   180
         Width           =   1530
      End
   End
   Begin VB.PictureBox picLabel 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   90
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   1
      Top             =   840
      Width           =   6885
      Begin VB.CheckBox chkLabelCrLf 
         Caption         =   "换行"
         Height          =   195
         Left            =   3870
         TabIndex        =   48
         Top             =   3000
         Width           =   780
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2700
         TabIndex        =   45
         Top             =   1185
         Width           =   1100
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "选用(&U)"
         Height          =   300
         Index           =   0
         Left            =   2700
         TabIndex        =   44
         Top             =   885
         Width           =   1100
      End
      Begin VB.TextBox txtLabelPrefix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5670
         TabIndex        =   50
         Top             =   2940
         Width           =   975
      End
      Begin VB.ListBox lstLabelUsed 
         Height          =   2400
         Left            =   3870
         TabIndex        =   47
         Top             =   465
         Width           =   2775
      End
      Begin VB.ListBox lstLabelItems 
         Height          =   2760
         Left            =   240
         TabIndex        =   43
         Top             =   465
         Width           =   2370
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "应用(&A)"
         Height          =   300
         Index           =   2
         Left            =   2700
         TabIndex        =   51
         Top             =   1635
         Width           =   1100
      End
      Begin VB.Label lblLabelPrefix 
         AutoSize        =   -1  'True
         Caption         =   "前缀文本"
         Height          =   180
         Left            =   4905
         TabIndex        =   49
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label lblLabelUsed 
         AutoSize        =   -1  'True
         Caption         =   "已设置的标签:"
         Height          =   180
         Left            =   3885
         TabIndex        =   46
         Top             =   180
         Width           =   1170
      End
      Begin VB.Label lblLabelItems 
         AutoSize        =   -1  'True
         Caption         =   "可选标签项目:"
         Height          =   180
         Left            =   240
         TabIndex        =   42
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.PictureBox picFoot 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   870
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   129
      Top             =   1500
      Width           =   6885
      Begin VB.CheckBox chk重算数据行 
         Caption         =   "重算数据行"
         Height          =   195
         Left            =   5280
         TabIndex        =   131
         Top             =   3090
         Width           =   1395
      End
      Begin VB.CheckBox chk页码 
         Caption         =   "打印页码"
         Height          =   195
         Left            =   270
         TabIndex        =   107
         Top             =   3090
         Width           =   1155
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2160
         Picture         =   "frmTendStyle.frx":0FF9
         Style           =   1  'Graphical
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   3030
         Width           =   345
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1830
         Picture         =   "frmTendStyle.frx":1352
         Style           =   1  'Graphical
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   3030
         Width           =   345
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1500
         Picture         =   "frmTendStyle.frx":16E2
         Style           =   1  'Graphical
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   3030
         Width           =   345
      End
      Begin VB.CommandButton cmd同步 
         Caption         =   "同步(&G)"
         Height          =   350
         Left            =   5730
         TabIndex        =   105
         ToolTipText     =   "所有护理文件的页眉页脚与当前文件的页眉页脚格式一致"
         Top             =   1350
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "插图(&I)"
         Height          =   350
         Left            =   150
         TabIndex        =   99
         Top             =   1350
         Width           =   1200
      End
      Begin VB.CheckBox chkI 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5250
         Picture         =   "frmTendStyle.frx":1A68
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "斜体(Alt+I)"
         Top             =   1350
         Width           =   345
      End
      Begin VB.CheckBox chkU 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4920
         Picture         =   "frmTendStyle.frx":82BA
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "下划线(Alt+U)"
         Top             =   1350
         Width           =   345
      End
      Begin VB.CheckBox chkB 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4590
         Picture         =   "frmTendStyle.frx":EB0C
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "粗体(Alt+B)"
         Top             =   1350
         Width           =   345
      End
      Begin VB.ComboBox cboFSize 
         Height          =   300
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   1380
         Width           =   750
      End
      Begin RichTextLib.RichTextBox rtbHead 
         Height          =   1200
         Left            =   30
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   30
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2117
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmTendStyle.frx":1535E
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
      Begin RichTextLib.RichTextBox rtbFoot 
         Height          =   1200
         Left            =   30
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   1770
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2117
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmTendStyle.frx":153FB
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
      Begin VB.ComboBox cboFont 
         Height          =   300
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   1380
         Width           =   1905
      End
      Begin VB.Label lbl字体 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "字体"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1515
         TabIndex        =   130
         Top             =   1440
         Width           =   360
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   8265
      Top             =   6345
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   125
      Top             =   7230
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendStyle.frx":15498
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17674
            Text            =   "可以根据医院实际情况，设置调整护理记录的查看与输出样式。"
            TextSave        =   "可以根据医院实际情况，设置调整护理记录的查看与输出样式。"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
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
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4005
      Index           =   0
      Left            =   8640
      ScaleHeight     =   4005
      ScaleWidth      =   6645
      TabIndex        =   121
      Top             =   360
      Width           =   6645
      Begin XtremeSuiteControls.TabControl tbcStyle 
         Height          =   3930
         Left            =   600
         TabIndex        =   124
         Top             =   210
         Width           =   5460
         _Version        =   589884
         _ExtentX        =   9631
         _ExtentY        =   6932
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1665
      Index           =   1
      Left            =   8700
      ScaleHeight     =   1665
      ScaleWidth      =   7860
      TabIndex        =   122
      Top             =   4560
      Width           =   7860
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   1425
         Left            =   0
         TabIndex        =   123
         Top             =   0
         Width           =   3810
         _cx             =   6720
         _cy             =   2514
         Appearance      =   2
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   12632256
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483644
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   8
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTendStyle.frx":15D2C
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
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
   Begin VB.PictureBox picPaper 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   690
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   3
      Top             =   1350
      Width           =   6885
      Begin VB.Frame fraPaper 
         Height          =   30
         Index           =   0
         Left            =   210
         TabIndex        =   97
         Top             =   2475
         Width           =   6390
      End
      Begin VB.ComboBox cboPaperKind 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   435
         Width           =   5355
      End
      Begin VB.TextBox txtHeight 
         Height          =   300
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   74
         Text            =   "297.08"
         Top             =   825
         Width           =   975
      End
      Begin VB.TextBox txtWidth 
         Height          =   300
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   70
         Text            =   "210.05"
         Top             =   825
         Width           =   945
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   3
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   91
         Text            =   "19"
         Top             =   1665
         Width           =   615
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   2
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   87
         Text            =   "19"
         Top             =   1665
         Width           =   615
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   1
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   83
         Text            =   "25"
         Top             =   1260
         Width           =   615
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "纵向"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   5745
         TabIndex        =   94
         Top             =   1275
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "横向"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   5745
         TabIndex        =   95
         Top             =   1680
         Width           =   750
      End
      Begin MSComCtl2.UpDown udHeight 
         Height          =   300
         Left            =   4575
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196656
         OrigLeft        =   4170
         OrigTop         =   900
         OrigRight       =   4425
         OrigBottom      =   1185
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udWidth 
         Height          =   300
         Left            =   2235
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196657
         OrigLeft        =   1830
         OrigTop         =   893
         OrigRight       =   2070
         OrigBottom      =   1178
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   0
         Left            =   1905
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1275
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(0)"
         BuddyDispid     =   196658
         BuddyIndex      =   0
         OrigLeft        =   1785
         OrigTop         =   1410
         OrigRight       =   2025
         OrigBottom      =   1710
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   1
         Left            =   3615
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1275
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(1)"
         BuddyDispid     =   196658
         BuddyIndex      =   1
         OrigLeft        =   3780
         OrigTop         =   1410
         OrigRight       =   4020
         OrigBottom      =   1710
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   2
         Left            =   1905
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(2)"
         BuddyDispid     =   196658
         BuddyIndex      =   2
         OrigLeft        =   1785
         OrigTop         =   1815
         OrigRight       =   2025
         OrigBottom      =   2115
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   3
         Left            =   3615
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(3)"
         BuddyDispid     =   196658
         BuddyIndex      =   3
         OrigLeft        =   3780
         OrigTop         =   1815
         OrigRight       =   4020
         OrigBottom      =   2115
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   0
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   79
         Text            =   "25"
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblPrinter 
         AutoSize        =   -1  'True
         Caption         =   "打印机:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   270
         TabIndex        =   111
         Top             =   135
         Width           =   690
      End
      Begin VB.Label lblOrient 
         AutoSize        =   -1  'True
         Caption         =   "纸张方向:"
         Height          =   180
         Left            =   4830
         TabIndex        =   120
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label lblPaperHint 
         AutoSize        =   -1  'True
         Caption         =   "注意:  如果实际打印机和当前打印机不符，可能导致纸张设置失效！"
         Height          =   180
         Left            =   270
         TabIndex        =   112
         Top             =   2115
         Width           =   5490
      End
      Begin VB.Label lblPaper 
         AutoSize        =   -1  'True
         Caption         =   "纸张种类"
         Height          =   180
         Left            =   270
         TabIndex        =   67
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   0
         Left            =   2535
         TabIndex        =   72
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblHeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "高"
         Height          =   180
         Left            =   3375
         TabIndex        =   73
         Top             =   885
         Width           =   180
      End
      Begin VB.Label lblWidth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宽"
         Height          =   180
         Left            =   1065
         TabIndex        =   69
         Top             =   885
         Width           =   180
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   1
         Left            =   4890
         TabIndex        =   76
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblRound 
         AutoSize        =   -1  'True
         Caption         =   "页边距:"
         Height          =   180
         Left            =   270
         TabIndex        =   77
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "右"
         Height          =   180
         Index           =   3
         Left            =   2790
         TabIndex        =   90
         Top             =   1725
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "左"
         Height          =   180
         Index           =   2
         Left            =   1065
         TabIndex        =   86
         Top             =   1725
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下"
         Height          =   180
         Index           =   1
         Left            =   2790
         TabIndex        =   82
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上"
         Height          =   180
         Index           =   0
         Left            =   1065
         TabIndex        =   78
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   2
         Left            =   2190
         TabIndex        =   81
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   3
         Left            =   3915
         TabIndex        =   85
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   4
         Left            =   2190
         TabIndex        =   89
         Top             =   1725
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   5
         Left            =   3915
         TabIndex        =   93
         Top             =   1725
         Width           =   360
      End
   End
   Begin VB.PictureBox picTable 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   450
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   0
      Top             =   1170
      Width           =   6885
      Begin VB.OptionButton optTabTiers 
         Caption         =   "三(&3)"
         Height          =   180
         Index           =   2
         Left            =   2790
         TabIndex        =   12
         Top             =   750
         Width           =   780
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   3
         Left            =   4380
         TabIndex        =   119
         Top             =   1905
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   2
         Left            =   4380
         TabIndex        =   117
         Top             =   255
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   1
         Left            =   1020
         TabIndex        =   114
         Top             =   2355
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   0
         Left            =   1020
         TabIndex        =   113
         Top             =   225
         Width           =   2235
      End
      Begin VB.TextBox txtHeadText 
         Height          =   885
         Left            =   4980
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   2070
         Width           =   1710
      End
      Begin MSComCtl2.UpDown udHeadCol 
         Height          =   300
         Left            =   4335
         TabIndex        =   39
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtHeadCol"
         BuddyDispid     =   196677
         OrigLeft        =   5985
         OrigTop         =   2085
         OrigRight       =   6225
         OrigBottom      =   2370
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtHeadCol 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4020
         MaxLength       =   2
         TabIndex        =   38
         Text            =   "1"
         Top             =   2430
         Width           =   330
      End
      Begin VB.TextBox txtHeadRow 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4035
         MaxLength       =   1
         TabIndex        =   35
         Text            =   "2"
         Top             =   2070
         Width           =   330
      End
      Begin MSComCtl2.UpDown udRecordTo 
         Height          =   285
         Left            =   5730
         TabIndex        =   28
         Top             =   465
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   8
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRecordTo"
         BuddyDispid     =   196684
         OrigLeft        =   5985
         OrigTop         =   405
         OrigRight       =   6225
         OrigBottom      =   705
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRecordFrom 
         Height          =   285
         Left            =   4365
         TabIndex        =   25
         Top             =   465
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   18
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRecordFrom"
         BuddyDispid     =   196685
         OrigLeft        =   4440
         OrigTop         =   405
         OrigRight       =   4680
         OrigBottom      =   705
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTabCols 
         Height          =   285
         Left            =   1530
         TabIndex        =   6
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   5
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtTabCols"
         BuddyDispid     =   196692
         OrigLeft        =   1530
         OrigTop         =   105
         OrigRight       =   1770
         OrigBottom      =   360
         Max             =   30
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdRecordFont 
         Caption         =   "条件字体(&N)"
         Height          =   300
         Left            =   3810
         TabIndex        =   30
         Top             =   840
         Width           =   1185
      End
      Begin VB.CommandButton cmdRecordColor 
         Caption         =   "条件颜色(&L)"
         Height          =   300
         Left            =   3810
         TabIndex        =   29
         Top             =   1200
         Width           =   1185
      End
      Begin VB.TextBox txtRecordTo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5385
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "8"
         Top             =   465
         Width           =   345
      End
      Begin VB.TextBox txtRecordFrom 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4020
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "18"
         Top             =   465
         Width           =   585
      End
      Begin VB.CommandButton cmdTitleFont 
         Caption         =   "标题字体(&T)"
         Height          =   300
         Left            =   435
         TabIndex        =   20
         Top             =   2910
         Width           =   1185
      End
      Begin VB.TextBox txtTitleText 
         Height          =   300
         Left            =   435
         TabIndex        =   19
         Text            =   "特别护理记录单"
         Top             =   2535
         Width           =   2790
      End
      Begin VB.CommandButton cmdTabGridColor 
         Caption         =   "表格颜色(&G)"
         Height          =   300
         Left            =   435
         TabIndex        =   17
         Top             =   1785
         Width           =   1185
      End
      Begin VB.CommandButton cmdTabTextColor 
         Caption         =   "文本颜色(&R)"
         Height          =   300
         Left            =   435
         TabIndex        =   15
         Top             =   1425
         Width           =   1185
      End
      Begin VB.CommandButton cmdTabFont 
         Caption         =   "文本字体(&F)"
         Height          =   300
         Left            =   435
         TabIndex        =   13
         Top             =   1065
         Width           =   1185
      End
      Begin VB.TextBox txtTabRowHeight 
         Height          =   300
         Left            =   2730
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "300"
         Top             =   345
         Width           =   510
      End
      Begin VB.OptionButton optTabTiers 
         Caption         =   "双(&2)"
         Height          =   180
         Index           =   1
         Left            =   1995
         TabIndex        =   11
         Top             =   750
         Width           =   780
      End
      Begin VB.OptionButton optTabTiers 
         Caption         =   "单(&1)"
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   750
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox txtTabCols 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "5"
         Top             =   345
         Width           =   300
      End
      Begin MSComCtl2.UpDown udHeadRow 
         Height          =   285
         Left            =   4335
         TabIndex        =   36
         Top             =   2070
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtHeadRow"
         BuddyDispid     =   196678
         OrigLeft        =   4920
         OrigTop         =   2085
         OrigRight       =   5160
         OrigBottom      =   2385
         Max             =   3
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTabHint 
         AutoSize        =   -1  'True
         Caption         =   "相同内容的同行相邻单元自动合并。"
         Height          =   180
         Index           =   1
         Left            =   3810
         TabIndex        =   118
         Top             =   3030
         Width           =   2880
      End
      Begin VB.Label lblTabHint 
         AutoSize        =   -1  'True
         Caption         =   "记录符合时间条件时使用该格式。"
         Height          =   180
         Index           =   0
         Left            =   3810
         TabIndex        =   116
         Top             =   1605
         Width           =   2700
      End
      Begin VB.Label lblBasic 
         AutoSize        =   -1  'True
         Caption         =   "基本形态"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   115
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblHeadText 
         AutoSize        =   -1  'True
         Caption         =   "文本"
         Height          =   180
         Left            =   4620
         TabIndex        =   40
         Top             =   2115
         Width           =   360
      End
      Begin VB.Label lblHeadCol 
         AutoSize        =   -1  'True
         Caption         =   "列号"
         Height          =   180
         Left            =   3630
         TabIndex        =   37
         Top             =   2490
         Width           =   360
      End
      Begin VB.Label lblHeadRow 
         AutoSize        =   -1  'True
         Caption         =   "层号"
         Height          =   180
         Left            =   3630
         TabIndex        =   34
         Top             =   2130
         Width           =   360
      End
      Begin VB.Label lblHeadSet 
         AutoSize        =   -1  'True
         Caption         =   "表头单元"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3600
         TabIndex        =   33
         Top             =   1845
         Width           =   780
      End
      Begin VB.Label lblRecordFont 
         Caption         =   "宋体,9"
         Height          =   180
         Left            =   5010
         TabIndex        =   32
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblRecordColor 
         Caption         =   "条件颜色"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   5010
         TabIndex        =   31
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lblRecordTo 
         AutoSize        =   -1  'True
         Caption         =   "次日       点"
         Height          =   180
         Left            =   4995
         TabIndex        =   26
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label lblRecordFrom 
         AutoSize        =   -1  'True
         Caption         =   "从       点至"
         Height          =   180
         Left            =   3810
         TabIndex        =   23
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label lblRecordStyle 
         AutoSize        =   -1  'True
         Caption         =   "条件样式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3600
         TabIndex        =   22
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lblTitleFont 
         Caption         =   "宋体,20"
         Height          =   180
         Left            =   1635
         TabIndex        =   21
         Top             =   3015
         Width           =   1605
      End
      Begin VB.Label lblTitleText 
         AutoSize        =   -1  'True
         Caption         =   "标题文本"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   18
         Top             =   2265
         Width           =   780
      End
      Begin VB.Shape shpTabGridColor 
         Height          =   180
         Left            =   1635
         Top             =   1890
         Width           =   1605
      End
      Begin VB.Label lblTabTextColor 
         Caption         =   "文本颜色"
         Height          =   180
         Left            =   1635
         TabIndex        =   16
         Top             =   1530
         Width           =   1605
      End
      Begin VB.Label lblTabFont 
         Caption         =   "宋体,9"
         Height          =   180
         Left            =   1635
         TabIndex        =   14
         Top             =   1185
         Width           =   1605
      End
      Begin VB.Label lblTabRowHeight 
         AutoSize        =   -1  'True
         Caption         =   "最小行高"
         Height          =   180
         Left            =   1935
         TabIndex        =   7
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lblTabTiers 
         AutoSize        =   -1  'True
         Caption         =   "表头层数"
         Height          =   180
         Left            =   435
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblTabCols 
         AutoSize        =   -1  'True
         Caption         =   "表格列数"
         Height          =   180
         Left            =   435
         TabIndex        =   4
         Top             =   405
         Width           =   720
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmTendStyle.frx":15E14
      Left            =   75
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendFileStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conRowHeight = 300        '基准行高度

'打印纸张常量(256=自定义)
Const PageSize1 = "信笺， 8 1/2×11 英寸"
Const PageSize2 = "+A611 小型信笺， 8 1/2×11 英寸"
Const PageSize3 = "小型报， 11×17 英寸"
Const PageSize4 = "分类帐， 17×11 英寸"
Const PageSize5 = "法律文件， 8 1/2×14 英寸"
Const PageSize6 = "声明书，5 1/2×8 1/2 英寸"
Const PageSize7 = "行政文件，7 1/2×10 1/2 英寸"
Const PageSize8 = "A3, 297×420 毫米"
Const PageSize9 = "A4, 210×297 毫米"
Const PageSize10 = "A4小号， 210×297 毫米"
Const PageSize11 = "A5, 148×210 毫米"
Const PageSize12 = "B4, 250×354 毫米"
Const PageSize13 = "B5, 182×257 毫米"
Const PageSize14 = "对开本， 8 1/2×13 英寸"
Const PageSize15 = "四开本， 215×275 毫米"
Const PageSize16 = "10×14 英寸"
Const PageSize17 = "11×17 英寸"
Const PageSize18 = "便条，8 1/2×11 英寸"
Const PageSize19 = "#9 信封， 3 7/8×8 7/8 英寸"
Const PageSize20 = "#10 信封， 4 1/8×9 1/2 英寸"
Const PageSize21 = "#11 信封， 4 1/2×10 3/8 英寸"
Const PageSize22 = "#12 信封， 4 1/2×11 英寸"
Const PageSize23 = "#14 信封， 5×11 1/2 英寸"
Const PageSize24 = "C 尺寸工作单"
Const PageSize25 = "D 尺寸工作单"
Const PageSize26 = "E 尺寸工作单"
Const PageSize27 = "DL 型信封， 110×220 毫米"
Const PageSize28 = "C5 型信封， 162×229 毫米"
Const PageSize29 = "C3 型信封， 324×458 毫米"
Const PageSize30 = "C4 型信封， 229×324 毫米"
Const PageSize31 = "C6 型信封， 114×162 毫米"
Const PageSize32 = "C65 型信封，114×229 毫米"
Const PageSize33 = "B4 型信封， 250×353 毫米"
Const PageSize34 = "B5 型信封，176×250 毫米"
Const PageSize35 = "B6 型信封， 176×125 毫米"
Const PageSize36 = "信封， 110×230 毫米"
Const PageSize37 = "信封大王， 3 7/8×7 1/2 英寸"
Const PageSize38 = "信封， 3 5/8×6 1/2 英寸"
Const PageSize39 = "U.S. 标准复写簿， 14 7/8×11 英寸"
Const PageSize40 = "德国标准复写簿， 8 1/2×12 英寸"
Const PageSize41 = "德国法律复写簿， 8 1/2×13 英寸"

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Private mlngFileID As Long       '被编辑的记录ID，修改、查阅时由上级程序通过ShowMe传递进入,新增时为0，
Private mblnRTBFoot As Boolean
Private mblnOK As Boolean        '是否完成编辑退出
Private mlngPageRow As Long     '当前页面格式定义的有效数据行

'页眉页脚相关
'######################################################################################################
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'矩形
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'包含用于格式化指定设备的相关信息
Private Type FORMATRANGE
    hDC As Long             '渲染设备
    hdcTarget As Long       '目标设备
    rc As RECT              '渲染区域，单位：缇。
    rcPage As RECT          '渲染设备的整体区域，单位：缇。
    chrg As CHARRANGE       '用于格式化的文本范围。
End Type

Private Type PageInfo
    PageNumber As Long      '页码
    Start As Long           '字符起始位置
    End As Long             '字符终止位置
    ActualHeight As Long    '本页实际打印高度
End Type
Private AllPages() As PageInfo   '页信息
Private Const WM_PASTE = &H302&              '粘贴
Private Const WM_USER = &H400                '通常用 WM_USER + X 来自定义消息
Private Const EM_FORMATRANGE = (WM_USER + 57)    '为某一设备格式化指定范围的文本。
Private Const EM_SETTARGETDEVICE = (WM_USER + 72) '设置用于所见即所得的目标设备和行宽。
Private Const EM_HIDESELECTION = (WM_USER + 63)  '显示/隐藏文本。
Private Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Private Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '获取中英文混合字符串长度
'######################################################################################################

'临时变量
Private rsItems As New ADODB.Recordset
Private rsTemp As New ADODB.Recordset
Private lngCount As Long
Private strTemp As String
Private lngCurColor As Long
Private strCurFont As String
Private objFont As StdFont
Private mblnChanged As Boolean
Private rtbThis As Object

'项目性质：录入项，选择，单选，多选
'一列一个项目的，程序不控制
'一列有两个项目的，项目之间必须加前缀或后缀加以识别且项目性质必须相同，只能是录入项或选择项
'一列有两个以上项目的，只能是录入项
'规则修改:如果列头文本含/且只绑定两个项目,项目的分隔符也是/,则类型必须相同,只能是录入或选择项;否则不控制
'
'日期提供一种格式：日/月，如8/6


Private Property Let DataChanged(vData As Boolean)
    
    mblnChanged = vData
        
    If mblnChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If
    
End Property

Private Property Get DataChanged() As Boolean
    
    DataChanged = mblnChanged

End Property

Private Function GetPaperName(ByVal intSize As Integer) As String
    '功能： 根据当前打印机的设置，获取纸张名称
    '返回： 纸张名称
    If intSize = 256 Then
        GetPaperName = "用户自定义 ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        GetPaperName = Switch( _
            intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41)
    Else
        GetPaperName = "不可测的纸张 ..."
    End If
End Function

Private Sub LoadPaper()
    '---------------------------------------------------
    '功能：装入当前打印机可用的纸张
    '---------------------------------------------------
    Dim intCurPaper As Integer
    Dim strDevice As String
        
    With Me.cboPaperKind
        .AddItem GetPaperName(256)
        .ItemData(.NewIndex) = 256
        If Not ExistsPrinter Then
            .ListIndex = 0
            .Enabled = False
            Exit Sub
        End If
        
        strDevice = GetSetting("ZLSOFT", "公共模块\zl9PrintMode\Default", "DeviceName", Printer.DeviceName)
        For lngCount = 0 To Printers.Count - 1
            If Printers(lngCount).DeviceName = strDevice Then
                Set Printer = Printers(lngCount)
                Exit For
            End If
        Next
        Me.lblPrinter.Caption = "当前打印机: " & Printer.DeviceName
        
        intCurPaper = Printer.PaperSize
        .Enabled = True
        For lngCount = 1 To 41
            Err = 0: On Error Resume Next
            Printer.PaperSize = lngCount
            Err = 0: On Error GoTo 0
            If Printer.PaperSize = lngCount Then
                .AddItem GetPaperName(lngCount)
                .ItemData(.NewIndex) = lngCount
                If lngCount = intCurPaper Then .ListIndex = .NewIndex
            End If
        Next
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub

Public Function ShowMe(ByVal frmParent As Object, Optional ByVal lngFileID As Long) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '返回：确定返回新增或修改的ID；取消返回0
    '---------------------------------------------------
    Dim strTitle As String
    
    mlngFileID = lngFileID

    Err = 0: On Error GoTo errHand
    
    If RefreshData = False Then
        DataChanged = False
        Unload Me
        Exit Function
    End If
    
    DataChanged = False
    
    '---------------------------------------------------
    '窗体显示
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
'    DataChanged = False
    
    ShowMe = mblnOK
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function

Private Function RefreshData() As Boolean


    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '返回：确定返回新增或修改的ID；取消返回0
    '---------------------------------------------------
    Dim strTitle As String
    
    
    '
    With vfgThis
        .Cols = 6
        .Cell(flexcpText, 1, 1, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = ""
    End With
    '---------------------------------------------------
    '根据当前打印机，装入可选纸张
    '---------------------------------------------------
    Call LoadPaper
    Call PreparePage
    '---------------------------------------------------
    '基础数据获取
    '---------------------------------------------------
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select l.编号, l.名称, l.说明 From 病历文件列表 l Where l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    Me.Caption = "护理记录样式 - " & rsTemp!名称
    strTitle = rsTemp!名称
    
    gstrSQL = "Select i.中文名" & _
            " From 诊治所见项目 i, 诊治所见分类 k" & _
            " Where k.Id = i.分类id And k.编码 In ('02', '05', '06') And i.替换域 = 1" & _
            " Order By i.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lstLabelItems.Clear
        Do While Not .EOF
            Me.lstLabelItems.AddItem "" & !中文名
            .MoveNext
        Loop
        If Me.lstLabelItems.ListCount > 0 Then Me.lstLabelItems.ListIndex = 0
    End With
    
    gstrSQL = "Select 项目序号,项目名称,项目类型,项目表示 From 护理记录项目 Order By 项目序号"
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsItems
        Me.lstColumnItems.Clear
        Me.lstColumnItems.AddItem "日期"
        Me.lstColumnItems.AddItem "时间"
        Do While Not .EOF
            Me.lstColumnItems.AddItem "" & !项目名称
            .MoveNext
        Loop
        Me.lstColumnItems.AddItem "护士"
        Me.lstColumnItems.AddItem "签名人"
        Me.lstColumnItems.AddItem "签名时间"
        Me.lstColumnItems.ListIndex = 0
        .MoveFirst
    End With
    
    '---------------------------------------------------
    '定义样式获取
    '---------------------------------------------------
    '空表格时未能设置其表头层次,导致缺省表头层数是单层,但表头层数计数器最小值为2,在点击该表头时报错
    Me.optTabTiers(0).Value = True
    Call optTabTiers_Click(0)
    
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    mlngPageRow = 0
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数"
                If Val("" & !内容文本) = 1 Then
                    Me.optTabTiers(0).Value = True
                    Call optTabTiers_Click(0)
                ElseIf Val("" & !内容文本) = 2 Then
                    Me.optTabTiers(1).Value = True
                    Call optTabTiers_Click(1)
                Else
                    Me.optTabTiers(2).Value = True
                    Call optTabTiers_Click(2)
                End If
            Case "总列数":  Me.udTabCols.Value = Val("" & !内容文本)
            Case "最小行高"
                Me.txtTabRowHeight.Text = Val("" & !内容文本)
                Call txtTabRowHeight_Change
            Case "文本字体"
                Me.lblTabFont.Caption = "" & !内容文本
                strCurFont = Me.lblTabFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set Me.vfgThis.Font = objFont
            Case "文本颜色"
                Me.lblTabTextColor.ForeColor = Val("" & !内容文本)
                Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
            Case "表格颜色"
                Me.shpTabGridColor.BorderColor = Val("" & !内容文本)
                With Me.vfgThis
                    .GridColor = Me.shpTabGridColor.BorderColor
                    .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
                    .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
                End With
            
            Case "标题文本"
                Me.txtTitleText.Text = "" & !内容文本
                Call txtTitleText_Change
            Case "标题字体"
                Me.lblTitleFont.Caption = "" & !内容文本
                strCurFont = Me.lblTitleFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                With Me.vfgThis
                    Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
                    .ROWHEIGHT(1) = objFont.Size * 20 + 150
                End With
            
            Case "开始时间": Me.udRecordFrom.Value = Val("" & !内容文本)
            Case "终止时间": Me.udRecordTo.Value = Val("" & !内容文本)
            Case "条件字体"
                Me.lblRecordFont.Caption = "" & !内容文本
                strCurFont = Me.lblRecordFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                With Me.vfgThis
                    Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
                End With
            Case "条件颜色"
                Me.lblRecordColor.ForeColor = Val("" & !内容文本)
                With Me.vfgThis
                    .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
                End With
            Case "有效数据行"
                mlngPageRow = Val(!内容文本)
            End Select
            .MoveNext
        Loop
    End With
    chk重算数据行.Value = (IIf(mlngPageRow = 0, 1, 0))
    
    Dim strPaper As String, blnHead As Boolean, blnFoot As Boolean
    gstrSQL = "Select 种类||'-'||编号 AS KEY,格式,页脚 From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If Not rsTemp.EOF Then
        strPaper = "" & rsTemp!格式:
        blnHead = ReadPageHead(rtbHead, rsTemp!Key)
        blnFoot = ReadPageFoot(rtbFoot, rsTemp!Key)
        cmd同步.Enabled = blnHead Or blnFoot
        picFoot.Tag = rsTemp!Key
        
        chk页码.Value = IIf(Val(NVL(rsTemp!页脚, 0)) > 0, 1, 0)
        If chk页码.Value = 1 Then optPageAlign(Val(NVL(rsTemp!页脚, 0)) - 1).Value = True
    End If
    
    If UBound(Split(strPaper, ";")) >= 0 Then
        For lngCount = 0 To Me.cboPaperKind.ListCount - 1
            If Me.cboPaperKind.ItemData(lngCount) = Val(Split(strPaper, ";")(0)) Then Me.cboPaperKind.ListIndex = lngCount: Exit For
        Next
        If Me.cboPaperKind.ListIndex = 0 Then
            If UBound(Split(strPaper, ";")) >= 2 Then Me.txtHeight.Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(2)), vbTwips, vbMillimeters), 2)
            If UBound(Split(strPaper, ";")) >= 3 Then Me.txtWidth.Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(3)), vbTwips, vbMillimeters), 2)
        End If
    End If
    If UBound(Split(strPaper, ";")) >= 1 Then
        If Val(Split(strPaper, ";")(1)) = 2 Then
            Me.optOrient(1).Value = True
        Else
            Me.optOrient(0).Value = True
        End If
    End If
    If UBound(Split(strPaper, ";")) >= 4 Then Me.txtMarjin(2).Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 5 Then Me.txtMarjin(3).Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 6 Then Me.txtMarjin(0).Text = Round(Me.ScaleX(Val(Split(strPaper, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 7 Then Me.txtMarjin(1).Text = Round(Me.ScaleX(Val(Split(strPaper, ";")(7)), vbTwips, vbMillimeters), 2)
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Me.lstLabelUsed.Clear
        Do While Not .EOF
            Me.lstLabelUsed.AddItem !内容文本 & "{" & !要素名称 & "}"
            Me.lstLabelUsed.ItemData(Me.lstLabelUsed.NewIndex) = !是否换行
            .MoveNext
        Loop
        If Me.lstLabelUsed.ListCount > 0 Then
            Me.lstLabelUsed.ListIndex = 0
            Me.cmdLabel(1).Enabled = True
            Me.chkLabelCrLf.Enabled = True
            Me.txtLabelPrefix.Enabled = True
        Else
            Me.cmdLabel(1).Enabled = False
            Me.chkLabelCrLf.Enabled = False: Me.chkLabelCrLf.Value = vbUnchecked
            Me.txtLabelPrefix.Enabled = False: Me.txtLabelPrefix.Text = ""
        End If
        Call cmdLabel_Click(2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            Me.vfgThis.TextMatrix(!内容行次 + 2, !对象序号) = "" & !内容文本
            .MoveNext
        Loop
        Call udHeadCol_Change
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Me.lstColumnUsed.Clear
        Do While Not .EOF
            Me.vfgThis.ColWidth(!对象序号) = Val(Split("" & !对象属性, "`")(0))
            If InStr(1, "" & !对象属性, "`") <> 0 Then
                vfgThis.Cell(flexcpAlignment, 6, !对象序号, 7, !对象序号) = Val(Split("" & !对象属性, "`")(1))
            Else
                vfgThis.Cell(flexcpAlignment, 6, !对象序号, 7, !对象序号) = flexAlignLeftCenter
            End If
            If Me.udColumnNo.Value <> !对象序号 Then Me.udColumnNo.Value = !对象序号
            Me.lstColumnUsed.AddItem !内容文本 & "{" & !要素名称 & "}" & !要素单位
            Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = zlCommFun.NVL(!要素表示, 0)
            
            .MoveNext
            If .EOF Then
                Call cmdColumn_Click(2)
            ElseIf Me.udColumnNo.Value <> !对象序号 Then
                Call cmdColumn_Click(2)
            End If
        Loop
        Me.udColumnNo.Value = Me.vfgThis.Col
    End With
    
        '汇总时段
    '------------------------------------------------------------------------------------------------------------------
    Dim aryTmp As Variant
    
    
    gstrSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '汇总时段'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With vsf
        If rsTemp.BOF = False Then
            Do While Not rsTemp.EOF
                strTemp = zlCommFun.NVL(rsTemp!内容文本)
                If strTemp <> "" Then
                    
                    aryTmp = Split(strTemp, ",")
                    
                    If UBound(aryTmp) >= 2 Then
                        If .TextMatrix(.Rows - 1, 1) <> "" And .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
                        
                        .TextMatrix(.Rows - 1, 1) = Trim(aryTmp(0))
                        .TextMatrix(.Rows - 1, 2) = Trim(aryTmp(1))
                        .TextMatrix(.Rows - 1, 3) = Trim(aryTmp(2))
                    End If
                End If
                rsTemp.MoveNext
            Loop
            mclsVsf.AppendRows = True
        End If
    End With
    
    '再按列合并
    For lngCount = 1 To vfgThis.Cols - 1
        vfgThis.MergeCol(lngCount) = True
    Next
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    
    RefreshData = True
    
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

    
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False
    
    cbsThis.Icons = frmPubIcons.imgPublic.Icons
        With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '快键绑定
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, "保存并退出"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存"): cbrControl.ToolTipText = "保存已更改的数据(Ctrl+S,F2)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "恢复"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "恢复到上次保存时的数据状态"
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "帮助(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): cbrControl.ToolTipText = "退出当前的设计窗体(Esc)"

    End With
        
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_File_Exit
        
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlFontSet(strTitle As String, strFont As String)
    With Me.dlgThis
        .flags = &H3 Or &H400 Or &H200 Or &H10000
        .DialogTitle = strTitle
        .FontName = Split(strFont, ",")(0)
        .FontSize = Val(Split(strFont, ",")(1))
        If InStr(1, strFont, "粗") > 0 Then
            .FontBold = True
        Else
            .FontBold = False
        End If
        If InStr(1, strFont, "斜") > 0 Then
            .FontItalic = True
        Else
            .FontItalic = False
        End If
        Err = 0: On Error Resume Next
        .ShowFont
        .flags = 0
        If Err.Number <> 0 Then Exit Sub
        strFont = .FontName & "," & .FontSize
        If .FontBold Or .FontItalic Then
            strFont = strFont & "," & IIf(.FontBold, "粗", "") & IIf(.FontItalic, "斜", "")
        End If
    End With
End Sub

Private Sub zlColorSet(strTitle As String, lngColor As Long)
    With Me.dlgThis
        .DialogTitle = strTitle
        .COLOR = lngColor
        Err = 0: On Error Resume Next
        .ShowColor
        If Err.Number <> 0 Then Exit Sub
        lngColor = .COLOR
    End With
End Sub

Private Sub cboPaperKind_Click()
    If Me.cboPaperKind.ListIndex <= 0 Then
        Me.txtWidth.Enabled = True: Me.udWidth.Enabled = True
        Me.txtHeight.Enabled = True: Me.udHeight.Enabled = True
        Me.optOrient(0).Value = True
        Me.optOrient(0).Enabled = False: Me.optOrient(1).Enabled = False
    Else
        Me.txtWidth.Enabled = False: Me.udWidth.Enabled = False
        Me.txtHeight.Enabled = False: Me.udHeight.Enabled = False
        Me.optOrient(0).Enabled = True: Me.optOrient(1).Enabled = True
        Err = 0: On Error Resume Next
        Printer.PaperSize = Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex)
        Me.txtWidth.Text = Me.ScaleX(Printer.Width, vbTwips, vbMillimeters)
        Me.txtHeight.Text = Me.ScaleY(Printer.Height, vbTwips, vbMillimeters)
        If Printer.Orientation = 1 Then
            Me.optOrient(0).Value = True
        Else
            Me.optOrient(1).Value = True
        End If
    End If
    DataChanged = True
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        If SaveData Then
            DataChanged = False
            Unload Me
        End If
        
    Case conMenu_Edit_Transf_Save
        
        If SaveData Then
            DataChanged = False
        End If
        
    Case conMenu_Edit_Transf_Cancle
                
        Call RefreshData
        DataChanged = False
        
    Case conMenu_File_Exit
        
        mblnOK = False
        Unload Me
        
    Case conMenu_Help_Help
        
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Save
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Cancle
                
        Control.Enabled = DataChanged
        
    End Select
End Sub

Private Sub chk_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = chk.Value
    End With
    
    vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value) = Replace(vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value), IIf(chk.Value = 0, "`1", "`0"), "`" & lstColumnUsed.ItemData(lstColumnUsed.ListIndex))
    mblnChanged = True
End Sub

Private Sub chkLabelCrLf_Click()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = Me.chkLabelCrLf.Value
    End With
End Sub

Private Sub chk页码_Click()
    optPageAlign(0).Enabled = chk页码.Value
    optPageAlign(1).Enabled = chk页码.Value
    optPageAlign(2).Enabled = chk页码.Value
    If chk页码.Value = 1 Then
        If Not optPageAlign(0).Value Then
            If Not optPageAlign(1).Value Then
                If Not optPageAlign(2).Value Then optPageAlign(0).Value = True
            End If
        End If
    End If
    mblnChanged = True
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim strTmp As String
    Dim blnSplit As Boolean                         '多个项目时检查,如果前面的项目无后缀且后一个项目无前缀,则blnSplit=False,不允许继续
    Dim intType As Integer, intFace As Integer      '项目类型
    Dim strFaces As String                          '绑定两项目,只能是录入项目0与单选项目4;绑定两个以上项目,只能是录入项目0
    Dim strName As String                           '项目名称
    
    With Me.lstColumnUsed
        Select Case Index
        Case 0
            If Me.lstColumnItems.ListIndex = -1 Then Exit Sub
            .AddItem "{" & Me.lstColumnItems.List(Me.lstColumnItems.ListIndex) & "}"
            .ListIndex = .NewIndex
            Me.cmdColumn(1).Enabled = True
            Me.txtColumnPrefix.Enabled = True
            Me.txtColumnPostfix.Enabled = True
            chk.Enabled = True
        Case 1
            If .ListIndex = -1 Then Exit Sub
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                .ListIndex = -1
                Me.cmdColumn(1).Enabled = False
                Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
                Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
            End If
            
            chk.Enabled = True
        Case 2
            '当一列绑定2个项目时，项目之间必须存在前缀/或后缀符号加以区分
            '当一列绑定多个项目时，项目类型必须是录入型项目
            '单选与多选项目不能与其它项目一起绑定于某列
            '系统固定的项目，如签名人，日期，时间等，一列只能绑定一个
            strTemp = ""
            strTmp = ""
            
            If .ListCount = 1 Then
                strFaces = "0,1,2,3,4,5"
            ElseIf .ListCount = 2 Then
                strFaces = "0,4,5"
            Else
                strFaces = "0"
            End If
            For lngCount = 0 To .ListCount - 1
                strTemp = strTemp & Space(1) & .List(lngCount) & "`" & .ItemData(lngCount)
                strTmp = strTmp & Space(1) & .List(lngCount)
                strName = Mid(.List(lngCount), InStr(1, .List(lngCount), "{") + 1)
                strName = Mid(strName, 1, InStr(1, strName, "}") - 1)
                
                If lngCount > 0 Then
                    '检查项目间是否存在分隔符
                    If Not blnSplit Then
                        If Trim(Split(.List(lngCount), "{")(0) = "") Then
                            MsgBox "当一列绑定多个项目时，项目之间必须要存在前缀或后缀符号加以区分！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                    '检查项目类型是否一致
                    rsItems.Filter = "项目名称='" & strName & "'"
                    If rsItems.RecordCount <> 0 Then
                        If Not (intType = rsItems!项目类型 And intFace = rsItems!项目表示) Then
                            MsgBox "当一列绑定多个项目时，项目的类型必须一致！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    '只需要取第一个项目的属性即可
                    rsItems.Filter = "项目名称='" & strName & "'"
                    If rsItems.RecordCount <> 0 Then
                        intType = rsItems!项目类型
                        intFace = rsItems!项目表示
                        If .ListCount > 1 Then
                            If intType = 1 And intFace = 0 Then
                                MsgBox "一列要绑定多个项目时,不允许绑定文本项目！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If intFace = 3 Then
                                MsgBox "多选项只能单独绑定！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    Else
                        If .ListCount > 1 Then
                            MsgBox "固定项目不能与其它项目绑定在一起！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                '固定取上一个项目的后缀
                blnSplit = Trim(Split(.List(lngCount), "}")(1) <> "")
                '检查项目表示(绑定两项目,只能是录入项目0与单选项目4;绑定两个以上项目,只能是录入项目0)
'                If rsItems.RecordCount <> 0 Then
'                    If InStr(1, strFaces, rsItems!项目表示) = 0 Then
'                        If .ListCount > 2 Then
'                            MsgBox "当一列绑定多个项目时，只能选择录入型的项目！", vbInformation, gstrSysName
'                        Else
'                            MsgBox "当一列绑定两个项目时，只能选择录入型或选择型的项目！", vbInformation, gstrSysName
'                        End If
'                        Exit Sub
'                    End If
'                End If
            Next
            strTemp = Trim(strTemp)
            strTmp = Trim(strTmp)
            rsItems.Filter = 0
            
            With vfgThis
                .TextMatrix(6, Me.udColumnNo.Value) = strTmp
                '根据对齐方式设置其内容
                Select Case .Cell(flexcpAlignment, 6, Me.udColumnNo.Value)
                Case 4
                    .TextMatrix(7, Me.udColumnNo.Value) = " " & strTmp & " "
                Case 7
                    .TextMatrix(7, Me.udColumnNo.Value) = " " & strTmp
                Case Else
                    .TextMatrix(7, Me.udColumnNo.Value) = strTmp & " "
                End Select
                
                .Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value) = strTemp
                .Cell(flexcpData, 7, udColumnNo.Value, 7, udColumnNo.Value) = strTemp & " "
                
            End With
            DataChanged = True
        End Select
    End With
End Sub

Private Sub cmdLabel_Click(Index As Integer)
    With Me.lstLabelUsed
        Select Case Index
        Case 0
            If Me.lstLabelItems.ListIndex = -1 Then Exit Sub
            .AddItem Me.lstLabelItems.List(Me.lstLabelItems.ListIndex) & "：{" & Me.lstLabelItems.List(Me.lstLabelItems.ListIndex) & "}"
            .ListIndex = .NewIndex
            Me.cmdLabel(1).Enabled = True
            Me.chkLabelCrLf.Enabled = True
            Me.txtLabelPrefix.Enabled = True
        Case 1
            If .ListIndex = -1 Then Exit Sub
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                .ListIndex = -1
                Me.cmdLabel(1).Enabled = False
                Me.chkLabelCrLf.Enabled = False: Me.chkLabelCrLf.Value = vbUnchecked
                Me.txtLabelPrefix.Enabled = False: Me.txtLabelPrefix.Text = ""
            End If
        Case 2
            Dim intCrLf As Integer
            intCrLf = 0
            strTemp = ""
            For lngCount = 0 To .ListCount - 1
                If .ItemData(lngCount) <> 0 Then intCrLf = intCrLf + 1
                strTemp = strTemp & Space(1) & IIf(.ItemData(lngCount) = 0, "", vbCrLf) & .List(lngCount)
            Next
            strTemp = Trim(strTemp)
            For lngCount = Me.vfgThis.FixedCols To Me.vfgThis.Cols - 1
                Me.vfgThis.TextMatrix(2, lngCount) = strTemp
            Next
            Me.vfgThis.ROWHEIGHT(2) = Me.vfgThis.FontSize * 20 * (intCrLf + 1) + 150
            DataChanged = True
        End Select
    End With
End Sub

Private Function CheckData() As Boolean
    Dim arrData
    Dim intType As Integer, intFace As Integer                  '两个项目的表示方式一样才允许
    Dim bln护士 As Boolean, bln日期 As Boolean, bln时间 As Boolean
    Dim strText As String, strItem As String
    Dim lngCOL As Long, lngCount As Long
    Dim intDo As Integer, intHead As Integer, intRow As Integer
    
    '每种护理记录单都必须要有一列绑定护士才行
    lngCount = vfgThis.Cols - 1
    For lngCOL = 1 To lngCount
        If InStr(1, "{护士},{签名人}", vfgThis.TextMatrix(6, lngCOL)) <> 0 Then
            bln护士 = True
        End If
    Next
    If Not bln护士 Then
        MsgBox "必须有一列绑定护士项目或签名人项目！", vbInformation, gstrSysName
        Exit Function
    End If
    If vfgThis.TextMatrix(6, 1) <> "{日期}" Then
        MsgBox "第一列必须绑定日期项目！", vbInformation, gstrSysName
        Exit Function
    End If
    If vfgThis.TextMatrix(6, 2) <> "{时间}" Then
        MsgBox "第二列必须绑定时间项目！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '只有绑定两个项目的列，项目的分隔符也是/，且项目类型必须相同,只能是录入、选择项或汇总项才允许设置列格式为对角线（特例：日期设置对角线表示日期简写）
    '只取列最近的列头（3列表头取5，2列表头取4，1列表头取3)
    If optTabTiers(0).Value Then
        intHead = 3
    ElseIf optTabTiers(1).Value Then
        intHead = 4
    Else
        intHead = 5
    End If
    For lngCOL = 1 To lngCount
        If vfgThis.Cell(flexcpData, 6, lngCOL) <> "" Then
            strText = Val(Split(vfgThis.Cell(flexcpData, 6, lngCOL), "`")(1))
            
            If strText = 1 Then
                '格式：{A}{B}，按}分解，2列分解出来的数组就=2
                strText = vfgThis.TextMatrix(6, lngCOL)
                If UBound(Split(strText, "}")) <> 2 Then
                    If strText <> "{日期}" Then
                        MsgBox "第" & lngCOL & "列 " & vfgThis.TextMatrix(intHead, lngCOL) & " 不允许设置列对角线！" & vbCrLf & "[只有日期列、绑定两个项目的列才允许设置列对角线]", vbInformation, gstrSysName
                        Exit Function
                    Else
                        GoTo ntloop
                    End If
                End If
                
                '两个项目的分隔符也必须是/
                If Trim(Mid(strText, InStr(1, strText, "}") + 1, InStr(InStr(1, strText, "{") + 1, strText, "{") - InStr(1, strText, "}") - 1)) <> "/" Then
                    MsgBox "第" & lngCOL & "列 " & vfgThis.TextMatrix(intHead, lngCOL) & " 设置了列对角线，要求绑定的项目格式应该是:A/B", vbInformation, gstrSysName
                    Exit Function
                End If
                
                '两个项目的项目类型必须一致
                For intDo = 0 To 1
                    strItem = Trim(GetItemName(strText, intDo))
                    rsItems.Filter = "项目名称='" & strItem & "'"
                    '录入的时候检查过了,系统项如日期,时间等不允许与固定项绑定,所以,此时不存在找不到的情况
                    If rsItems.RecordCount = 0 Then
                        MsgBox "项目:" & strItem & "已改名或已经被删除！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If intDo > 0 Then
                        If Not (intFace = rsItems!项目表示 And intType = rsItems!项目类型) Then
                            MsgBox "第" & lngCOL & "列 " & vfgThis.TextMatrix(intHead, lngCOL) & " 绑定的两个项目的编辑方式必须一致！", vbInformation, gstrSysName
                            Exit Function
                        End If
                    Else
                        intFace = rsItems!项目表示
                        intType = rsItems!项目类型
                        If InStr(1, "0,2,4,5", intFace) = 0 Then
                            MsgBox "第" & lngCOL & "列 " & vfgThis.TextMatrix(intHead, lngCOL) & " 绑定的两个项目必须都是数值型、选择项、单选项或汇总项！", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
ntloop:
    Next
    
     With vsf
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) <> "" And .TextMatrix(intRow, 2) <> "" And .TextMatrix(intRow, 3) <> "" Then
                
                If CheckTime(.TextMatrix(intRow, 2)) = False Then Exit Function
                If CheckTime(.TextMatrix(intRow, 3)) = False Then Exit Function
                
            End If
        Next
    End With
    
    rsItems.Filter = 0
    CheckData = True
End Function

Private Function GetItemName(ByVal strData As String, ByVal intOrder As Integer) As String
    Dim intDo As Integer, intPos As Integer
    '获取指定格式串中指定序号的项目名称，格式如：{收缩压}/ {舒张压}mmHg
    
    intPos = InStr(1, strData, "{")
    If intOrder > 0 Then intPos = InStr(intPos + 1, strData, "{")
    strData = Mid(strData, intPos + 1)
    strData = Mid(strData, 1, InStr(1, strData, "}") - 1)
    GetItemName = strData
End Function

Private Function CheckTime(ByVal strText As String) As Boolean

    Dim intRow As Integer
    Dim intPos As Integer
    Dim strTmp As String
        
    intPos = InStr(strText, ":")
    
    If intPos > 0 Then
        
        strTmp = Mid(strText, 1, intPos - 1)
        If Val(strTmp) < 0 Or Val(strTmp) > 23 Then
            MsgBox "小时只能在0-23之间！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strText = Mid(strText, intPos + 1)
        intPos = InStr(strText, ":")
        If intPos > 0 Then
            strTmp = Mid(strText, 1, intPos - 1)
            If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                MsgBox "分钟只能在0-59之间！", vbInformation, gstrSysName
                Exit Function
            End If
            
            If InStr(Mid(strText, intPos + 1), ":") > 0 Then
                
                MsgBox "时间格式不正确！", vbInformation, gstrSysName
                Exit Function
            Else
                strTmp = Mid(strText, intPos + 1)
                If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                    MsgBox "秒钟只能在0-59之间！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            strTmp = strText
            If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                MsgBox "分钟只能在0-59之间！", vbInformation, gstrSysName
                Exit Function
            End If
            
        End If
    Else
        If Val(strText) < 0 Or Val(strText) > 23 Then
            MsgBox "小时只能在0-23之间！", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    CheckTime = True
    
End Function

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean
    Dim blnReCalc As Boolean    '因每页的可打印数据行发生变化，对所有未打印完的文件进行重算
    Dim strCol As String
    Dim strInput As String
    Dim intPageAlign As Integer
    Dim lngRows As Long, lngFixedRows As Long
    If CheckData = False Then Exit Function
    
    '保存数据
    If Me.optOrient(0).Value = True Then
        If Val(Me.txtMarjin(0).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "上边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(1).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "下边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(2).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "左边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(3).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "右边距太大！", vbExclamation, gstrSysName: Exit Function
    Else
        If Val(Me.txtMarjin(0).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "上边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(1).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "下边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(2).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "左边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(3).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "右边距太大！", vbExclamation, gstrSysName: Exit Function
    End If
    
    '计算有效数据行(纸张|纸向|高|宽|上边距|下边距|左边距|右边距|行高|固定行数|标题栏字体名|标题栏字体大小|标题文本|表上项字体名|表上项字体大小|表上项文本)
    If optTabTiers(0).Value Then
        lngFixedRows = 1
    ElseIf optTabTiers(1).Value Then
        lngFixedRows = 2
    Else
        lngFixedRows = 3
    End If
    strInput = Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex) & "|" & IIf(Me.optOrient(0).Value, 1, 2) & "|" & _
               Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleX(Val(Me.txtWidth.Text), vbMillimeters, vbTwips)) & "|" & _
               Int(Me.ScaleY(Val(Me.txtMarjin(0).Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleY(Val(Me.txtMarjin(1).Text), vbMillimeters, vbTwips)) & "|" & _
               Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips)) & "|" & _
               Val(txtTabRowHeight.Text) & "|" & lngFixedRows & "|" & Split(lblTitleFont.Caption, ",")(0) & "|" & Split(lblTitleFont.Caption, ",")(1) & "|" & _
               txtTitleText.Text & "|" & Split(lblTabFont.Caption, ",")(0) & "|" & Split(lblTabFont.Caption, ",")(1) & "|" & strTemp
    lngRows = frmTendFilePreview.ShowMe(Me, strInput)
    If chk重算数据行.Value = 1 Then
        If lngRows <> mlngPageRow And mlngPageRow > 0 Then
            '数据行发生变化，对已打印数据有影响，提示
            If MsgBox("    本次修改导致每页可打印的数据行发生变化，已打印的数据若修改或重打将会导致打印混乱，是否继续？" & vbCrLf & "    原来每页可打印" & mlngPageRow & "行；现在每页可打印" & lngRows & "行", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnReCalc = True
        End If
    Else
        lngRows = mlngPageRow
    End If
    
    '如果超出页面宽度则禁止保存
    If OverRun Then
        MsgBox "表格的宽度超出了纸张有效打印范围！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not PageHeadTest Then Exit Function
    
    '拼串
    If chk页码.Value = 1 Then
        If optPageAlign(0).Value Then
            intPageAlign = 1
        ElseIf optPageAlign(1).Value Then
            intPageAlign = 2
        Else
            intPageAlign = 3
        End If
    End If
    If Me.optTabTiers(0).Value Then
        gstrSQL = mlngFileID & ",1," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    ElseIf Me.optTabTiers(1).Value Then
        gstrSQL = mlngFileID & ",2," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    Else
        gstrSQL = mlngFileID & ",3," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    End If

    gstrSQL = gstrSQL & ",'" & Me.lblTabFont.Caption & "'," & Me.lblTabTextColor.ForeColor & "," & Me.shpTabGridColor.BorderColor
    gstrSQL = gstrSQL & ",'" & Trim(Me.txtTitleText.Text) & "','" & Me.lblTitleFont.Caption & "'"
    gstrSQL = gstrSQL & "," & Me.udRecordFrom.Value & "," & Me.udRecordTo.Value & ",'" & Me.lblRecordFont.Caption & "'," & Me.lblRecordColor.ForeColor & "," & lngRows
    
    gstrSQL = gstrSQL & ",'" & Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex)
    gstrSQL = gstrSQL & ";" & IIf(Me.optOrient(0).Value, 1, 2)
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtWidth.Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtMarjin(0).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtMarjin(1).Text), vbMillimeters, vbTwips)) & "'"
    gstrSQL = gstrSQL & ",NULL," & intPageAlign
    
    With Me.vfgThis
        gstrSQL = gstrSQL & ",'" & Replace(Replace(.TextMatrix(2, .FixedCols), " ", "|"), vbCrLf, "'||Chr(13)||Chr(10)||'") & "'"
        strTemp = ""
        For lngCount = .FixedCols To .Cols - 1
            If .RowHidden(3) = False Then strTemp = strTemp & "|" & lngCount & ",1," & Trim(.TextMatrix(3, lngCount))
            If .RowHidden(4) = False Then strTemp = strTemp & "|" & lngCount & ",2," & Trim(.TextMatrix(4, lngCount))
            If .RowHidden(5) = False Then strTemp = strTemp & "|" & lngCount & ",3," & Trim(.TextMatrix(5, lngCount))
'            strTemp = strTemp & "|" & lngCount & ",3," & Trim(.TextMatrix(5, lngCount))
        Next
        gstrSQL = gstrSQL & ",'" & Mid(strTemp, 2) & "'"
        
        strCol = ""
        For lngCount = .FixedCols To .Cols - 1
            strCol = strCol & "|" & lngCount & "," & .ColWidth(lngCount) & "`" & .Cell(flexcpAlignment, 6, lngCount) & "," & Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount))
'            strTemp = strTemp & "|" & lngCount & "," & .ColWidth(lngCount) & "," & Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount))
        Next
        gstrSQL = gstrSQL & ",'" & Mid(strCol, 2) & "'"
    End With
    
    '填写汇总时段
    '------------------------------------------------------------------------------------------------------------------
    With vsf
        strTemp = ""
        For lngCount = 1 To .Rows - 1
            If .TextMatrix(lngCount, 1) <> "" And .TextMatrix(lngCount, 2) <> "" And .TextMatrix(lngCount, 3) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(lngCount, 1)) & "," & Trim(.TextMatrix(lngCount, 2)) & "," & Trim(.TextMatrix(lngCount, 3))
            End If
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        gstrSQL = gstrSQL & ",'" & strTemp & "'"
    End With
    
    gstrSQL = "Zl_护理文件样式_Update(" & gstrSQL & ")"
    
    Err = 0: On Error GoTo errHand
    gcnOracle.BeginTrans
    blnTrans = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If Not SavePageHead(picFoot.Tag) Then GoTo ErrExit
    If Not SavePageFoot(picFoot.Tag) Then GoTo ErrExit
    gcnOracle.CommitTrans
    blnTrans = False
    
    SaveData = True
    mblnOK = True
    mlngPageRow = lngRows
    cmd同步.Enabled = True
    
    '因数据行发生变化，进行数据重算
    If blnReCalc Then
        If MsgBox("可打印数据行发生变化，是否对使用该格式的所有文件进行重算？" & vbCrLf & "（更新打印数据可能会占用你几分钟时间）", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If frmTendFilePreview.AnaliseData(Me, mlngFileID, strInput) Then MsgBox "打印数据已更新！", vbInformation, gstrSysName
        End If
    End If
    
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
ErrExit:
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Private Function SavePageHead(ByVal strKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:产生文件并压缩;True:已产生压缩文件
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Head_S.rtf"
        rtbHead.SaveFile strFile
        strZip = zlFileZip(strFile)
    Else
        strZip = strZipFile
    End If
    SavePageHead = zlBlobSave(12, strKey, strZip)
    If strZipFile = "" Then
        gobjFSO.DeleteFile strFile, True
        gobjFSO.DeleteFile strZip, True
    End If
End Function

Private Function SavePageFoot(ByVal strKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:产生文件并压缩;True:已产生压缩文件
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Foot_S.rtf"
        rtbFoot.SaveFile strFile
        strZip = zlFileZip(strFile)
    Else
        strZip = strZipFile
    End If
    SavePageFoot = zlBlobSave(13, strKey, strZip)
    If strZipFile = "" Then
        gobjFSO.DeleteFile strFile, True
        gobjFSO.DeleteFile strZip, True
    End If
End Function

Private Function PreparePage()
    Dim sFont As String, i As Integer
    
    For i = 0 To Screen.FontCount - 1
       sFont = Screen.Fonts(i)
       cboFont.AddItem sFont
       If sFont = "宋体" Then cboFont.ListIndex = i
    Next i
    With cboFSize
        .AddItem "初号"
        .AddItem "小初"
        .AddItem "一号"
        .AddItem "小一"
        .AddItem "二号"
        .AddItem "小二"
        .AddItem "三号"
        .AddItem "小三"
        .AddItem "四号"
        .AddItem "小四"
        .AddItem "五号"
        .AddItem "小五"
        .AddItem "六号"
        .AddItem "小六"
        .AddItem "七号"
        .AddItem "八号"
        .AddItem 5
        .AddItem 5.5
        .AddItem 6.5
        .AddItem 7.5
        .AddItem 8
        .AddItem 9
        .AddItem 10
        .AddItem 10.5
        .AddItem 11
        .AddItem 12
        .AddItem 14
        .AddItem 16
        .AddItem 18
        .AddItem 20
        .AddItem 22
        .AddItem 24
        .AddItem 26
        .AddItem 28
        .AddItem 36
        .AddItem 48
        .AddItem 72
        .ListIndex = 10
    End With
End Function

Private Function OverRun() As Boolean
    Dim intPageMargin As Integer    '边距
    Dim intPageWidth As Integer     '纸张宽度
    Dim intTrimSize As Integer      '纸张实际宽度
    Dim dblTableWidth As Double     '表格宽度
    Dim lngCOL As Long, lngCols As Long
    '检查表格的总体宽度是否超过了右边距
    
    lngCols = Me.vfgThis.Cols - 1
    For lngCOL = 1 To lngCols
        dblTableWidth = dblTableWidth + Val(vfgThis.ColWidth(lngCOL))
    Next
    
    '检查表格宽度是否超出页面有效打印范围
    Printer.Orientation = IIf(optOrient(0).Value, 1, 2)
    intPageWidth = Printer.ScaleWidth
    intPageMargin = Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips)) + Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips))
    intTrimSize = intPageWidth - intPageMargin - 100
    If dblTableWidth > intTrimSize Then
        OverRun = True
        Exit Function
    End If
End Function

Private Sub cmdRecordColor_Click()
    lngCurColor = Me.lblRecordColor.ForeColor
    Call zlColorSet("条件样式颜色", lngCurColor)
    If lngCurColor = Me.lblRecordColor.ForeColor Then Exit Sub
    Me.lblRecordColor.ForeColor = lngCurColor
    With Me.vfgThis
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
    End With
    DataChanged = True
End Sub

Private Sub cmdRecordFont_Click()
    strCurFont = Me.lblRecordFont.Caption
    Call zlFontSet("条件样式字体", strCurFont)
    If strCurFont = Me.lblRecordFont.Caption Then Exit Sub
    Me.lblRecordFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
        If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
    End With
    With Me.vfgThis
        Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
    End With
    DataChanged = True
End Sub

Private Sub cmdTabFont_Click()
    strCurFont = Me.lblTabFont.Caption
    Call zlFontSet("文本字体", strCurFont)
    If strCurFont = Me.lblTabFont.Caption Then Exit Sub
    Me.lblTabFont.Caption = strCurFont
'    Me.lblRecordFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
         .Bold = False: .Italic = False
        If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
        If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
    End With
    Set Me.vfgThis.Font = objFont
    DataChanged = True
End Sub

Private Sub cmdTabGridColor_Click()
    lngCurColor = Me.shpTabGridColor.BorderColor
    Call zlColorSet("表格颜色", lngCurColor)
    If lngCurColor = Me.shpTabGridColor.BorderColor Then Exit Sub
    Me.shpTabGridColor.BorderColor = lngCurColor
    With Me.vfgThis
        .GridColor = Me.shpTabGridColor.BorderColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
    End With
    DataChanged = True
End Sub

Private Sub cmdTabTextColor_Click()
    lngCurColor = Me.lblTabTextColor.ForeColor
    Call zlColorSet("文本颜色", lngCurColor)
    If lngCurColor = Me.lblTabTextColor.ForeColor Then Exit Sub
    Me.lblTabTextColor.ForeColor = lngCurColor
    Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
    DataChanged = True
End Sub

Private Sub cmdTitleFont_Click()
    strCurFont = Me.lblTitleFont.Caption
    Call zlFontSet("标题字体", strCurFont)
    If strCurFont = Me.lblTitleFont.Caption Then Exit Sub
    Me.lblTitleFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
        If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
    End With
    With Me.vfgThis
        Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
        .ROWHEIGHT(1) = objFont.Size * 20 + 150
    End With
    DataChanged = True
End Sub

Private Sub cmd同步_Click()
    Dim intPageAlign As Integer
    Dim strZIPHead As String, strZIPFoot As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    On Error GoTo errHand
    '将当前格式应用到所有护理文件
    
    gstrSQL = " Select 种类||'-'||编号 AS KEY From 病历文件列表 Where 种类=3 And 保留<>-1 and ID<>[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件", mlngFileID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "当前只有一份护理文件，不需要执行同步功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("请再次确认：" & vbCrLf & "        执行该功能后，所有护理文件的页眉页脚格式将统一与当前文件设置保存一致！", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '先将当前文件的页眉页脚压缩包读出来
    If chk页码.Value = 1 Then
        If optPageAlign(0).Value Then
            intPageAlign = 1
        ElseIf optPageAlign(1).Value Then
            intPageAlign = 2
        Else
            intPageAlign = 3
        End If
    End If
    strZIPHead = ReadPageHeadFile(picFoot.Tag)
    strZIPFoot = ReadPageFootFile(picFoot.Tag)
    
    gcnOracle.BeginTrans
    blnTrans = True
    '循环写入数据库
    With rsTemp
        Do While Not .EOF
            If Not SavePageHead(!Key, strZIPHead) Then GoTo ErrExit
            If Not SavePageFoot(!Key, strZIPFoot) Then GoTo ErrExit
            .MoveNext
        Loop
    End With
    Call zlDatabase.ExecuteProcedure("ZL_病历页面格式_页码(" & intPageAlign & ",'" & Split(picFoot.Tag, "-")(1) & "')", "更新页脚")
    gcnOracle.CommitTrans
    blnTrans = False
    '删除临时文件
    
    gobjFSO.DeleteFile strZIPHead, True
    gobjFSO.DeleteFile strZIPFoot, True
    
    MsgBox "同步成功！", vbInformation, gstrSysName
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
ErrExit:
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hwnd
    Case 2
        Item.Handle = picPane(1).hwnd
    End Select
End Sub

Private Sub Form_Load()
    Me.picTable.BackColor = Me.BackColor
    Me.picLabel.BackColor = Me.BackColor
    Me.picCloumn.BackColor = Me.BackColor
    Me.picPaper.BackColor = Me.BackColor
    Me.picFoot.BackColor = Me.BackColor

    '---------------------------------------------------
    '属性页面组织
    With Me.tbcStyle
        .Left = Me.ScaleLeft: .Top = Me.ScaleTop
        With .PaintManager
'            .Appearance = xtpTabAppearancePropertyPage
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameBorder
        End With
        .InsertItem 0, "基本属性", Me.picTable.hwnd, 0
        .InsertItem 1, "表上标签", Me.picLabel.hwnd, 0
        .InsertItem 2, "表列设置", Me.picCloumn.hwnd, 0
        .InsertItem 3, "打印设置", Me.picPaper.hwnd, 0
        .InsertItem 4, "页面格式", Me.picFoot.hwnd, 0
        .InsertItem 5, "汇总时段", Me.picSum.hwnd, 0
        .Item(0).Selected = True
    End With
    With Me.vfgThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.tbcStyle.Top + Me.tbcStyle.Height + 45: .Height = Me.ScaleHeight - .Top
    End With
'    With Me.picVBar
'        .BackColor = Me.BackColor
'        .Left = Me.ScaleWidth - .Width
'        .Top = 0: .Height = Me.Height
'    End With
    
    '---------------------------------------------------
    '默认样式设置
    Err = 0: On Error GoTo 0
    With Me.vfgThis
        .Rows = 9
        .TextMatrix(1, 0) = "标题文本"
        .TextMatrix(2, 0) = "表上标签"
        
        .TextMatrix(3, 0) = "表头单元"
        .TextMatrix(4, 0) = "表头单元"
        .TextMatrix(5, 0) = "表头单元"
        
        .TextMatrix(6, 0) = "表列内容"
        .TextMatrix(7, 0) = "条件格式"
        .TextMatrix(8, 0) = "表下标签"
        
        .ColWidth(0) = 1200
        .MergeCol(0) = True
        .RowHidden(3) = True
        .RowHidden(4) = True
        
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
'            .TextMatrix(.Rows - 1, lngCount) = "打印时间：[打印时间]                            第[页码]页"
        Next
        .MergeRow(1) = True
        .MergeRow(2) = True: .Cell(flexcpAlignment, 2, 1, 2, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(.Rows - 1) = True: .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(3) = True
        .MergeRow(4) = True
        .MergeRow(5) = True
        
        .Cell(flexcpAlignment, 6, 1, 7, .Cols - 1) = flexAlignGeneralCenter
        
        Call udTabCols_Change
        Call txtTabRowHeight_Change
        strCurFont = Me.lblTabFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
        End With
        
        .ForeColor = Me.lblTabTextColor.ForeColor
        
        .GridColor = Me.shpTabGridColor.BorderColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
            
        Call txtTitleText_Change
        strCurFont = Me.lblTitleFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
        .ROWHEIGHT(1) = objFont.Size * 20 + 150
            
        strCurFont = Me.lblRecordFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
        
        .RowHidden(3) = True
        .RowHidden(4) = True
        
    End With
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    
    Set mclsVsf = New clsVsf
    With mclsVsf
        
        Call .Initialize(Me.Controls, vsf, True, True)
        Call .ClearColumn

        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        
        Call .AppendColumn("时段名称", 2100, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("开始时间", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("结束时间", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
        
        Call .InitializeEdit(True, True, True)
        Call .InitializeEditColumn(mclsVsf.ColIndex("时段名称"), True, vbVsfEditText)
        Call .InitializeEditColumn(mclsVsf.ColIndex("开始时间"), True, vbVsfEditText)
        Call .InitializeEditColumn(mclsVsf.ColIndex("结束时间"), True, vbVsfEditText)
                
        .AppendRows = True
    End With
    
    vsf.ColHidden(0) = True
    
    Dim objPane As Pane
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitMenuBar
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.UseSplitterTracker = False '实时拖动
    dkpMan.Options.AlphaDockingContext = True
    dkpMan.Options.CloseGroupOnButtonClick = True
    dkpMan.Options.HideClient = True
    dkpMan.SetCommandBars cbsThis

    Set objPane = dkpMan.CreatePane(1, 100, 100, DockTopOf, Nothing): objPane.Title = "示例": objPane.Options = PaneNoCaption
    Set objPane = dkpMan.CreatePane(2, 100, 200, DockBottomOf, objPane): objPane.Title = "设计": objPane.Options = PaneNoCaption
                
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    rtbHead.Width = picFoot.Width - 500
    rtbFoot.Width = rtbHead.Width
    
    Call SetPaneRange(dkpMan, 1, 15, 255, Me.ScaleWidth, 255)
    dkpMan.RecalcLayout
    
    rtbHead.Width = picFoot.Width - 500
    rtbFoot.Width = rtbHead.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If DataChanged Then
        Cancel = (MsgBox("更改后的设计必须保存后才生效，是否放弃保存？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    
    If Cancel Then Exit Sub
    
    DataChanged = False
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lstColumnItems_DblClick()
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnUsed_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        Me.txtColumnPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
        Me.txtColumnPostfix.Text = Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "}") + 1)
        chk.Value = .ItemData(.ListIndex)
    End With
End Sub

Private Sub lstLabelItems_DblClick()
    Call cmdLabel_Click(0)
End Sub

Private Sub lstLabelUsed_Click()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        Me.chkLabelCrLf.Value = IIf(.ItemData(.ListIndex) = 0, vbUnchecked, vbChecked)
        Me.txtLabelPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
    End With
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf
        Cancel = (.TextMatrix(Row, 1) = "" Or .TextMatrix(Row, 2) = "" Or .TextMatrix(Row, 3) = "")
    End With
End Sub

Private Sub optAlign_Click(Index As Integer)
    Dim intAlign As Integer
    
    Select Case Index
    Case 0
        intAlign = flexAlignLeftCenter
        vfgThis.TextMatrix(7, vfgThis.Col) = vfgThis.TextMatrix(6, vfgThis.Col) & " "
    Case 1
        intAlign = flexAlignCenterCenter
        vfgThis.TextMatrix(7, vfgThis.Col) = " " & vfgThis.TextMatrix(6, vfgThis.Col) & " "
    Case 2
        intAlign = flexAlignRightCenter
        vfgThis.TextMatrix(7, vfgThis.Col) = " " & vfgThis.TextMatrix(6, vfgThis.Col)
    End Select
    vfgThis.Cell(flexcpAlignment, 6, vfgThis.Col, 7, vfgThis.Col) = intAlign
    DataChanged = True
    
    On Error Resume Next
    chk.SetFocus
End Sub

Private Sub optOrient_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub optPageAlign_Click(Index As Integer)
    mblnChanged = True
End Sub

Private Sub optTabTiers_Click(Index As Integer)
    
    With vfgThis
        If optTabTiers(0).Value Then
            If .Row = 5 Or .Row = 4 Then .Row = 3
            .RowHidden(3) = False
            .RowHidden(4) = True
            .RowHidden(5) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 1
        ElseIf optTabTiers(1).Value Then
            If .Row = 5 Then .Row = 4
            .RowHidden(3) = False
            .RowHidden(4) = False
            .RowHidden(5) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 2
        Else
            .RowHidden(3) = False
            .RowHidden(4) = False
            .RowHidden(5) = False
            udHeadRow.Min = 1
            udHeadRow.Max = 3
        End If

    End With
    DataChanged = True
End Sub

Private Sub optTabTiers_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picPane_Resize(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
    Case 0
        tbcStyle.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
'        cmdOK.Left = tbcStyle.Left + tbcStyle.Width + 90
'        cmdCancel.Left = tbcStyle.Left + tbcStyle.Width + 90
'        imgNote.Left = tbcStyle.Left + tbcStyle.Width + 90
'        lblNote.Left = tbcStyle.Left + tbcStyle.Width + 90
    Case 1
        vfgThis.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub picSum_Resize()
    On Error Resume Next
    
    vsf.Move 15, 15, picSum.Width - 30, picSum.Height - 30
    mclsVsf.AppendRows = True
    
End Sub

Private Sub rtbFoot_Change()
    mblnChanged = True
End Sub

Private Sub rtbFoot_GotFocus()
    mblnRTBFoot = True
End Sub

Private Sub rtbHead_Change()
    mblnChanged = True
End Sub

Private Sub rtbHead_GotFocus()
    mblnRTBFoot = False
End Sub

Private Sub tbcStyle_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Caption = "页面格式" Then
        On Error Resume Next
        
        Printer.PaperSize = cboPaperKind.ItemData(cboPaperKind.ListIndex)
        Printer.Orientation = IIf(optOrient(0).Value, 1, 2)
        Call SendMessage(rtbHead.hwnd, EM_SETTARGETDEVICE, Me.hDC, ByVal CLng(Printer.ScaleWidth))
        SendMessageLong rtbHead.hwnd, EM_HIDESELECTION, 0, 0
        Call SendMessage(rtbFoot.hwnd, EM_SETTARGETDEVICE, Me.hDC, ByVal CLng(Printer.ScaleWidth))
        SendMessageLong rtbFoot.hwnd, EM_HIDESELECTION, 0, 0
    
        rtbHead.Width = picFoot.Width - 500
        rtbFoot.Width = rtbHead.Width
    End If
End Sub

Private Sub txtColumnPostfix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "}")) & Me.txtColumnPostfix.Text
    End With
End Sub

Private Sub txtColumnPostfix_GotFocus()
    Me.txtColumnPostfix.SelStart = 0: Me.txtColumnPostfix.SelLength = 4000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtColumnPostfix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtColumnPrefix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Me.txtColumnPrefix.Text & Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "{"))
    End With
End Sub

Private Sub txtColumnPrefix_GotFocus()
    Me.txtColumnPrefix.SelStart = 0: Me.txtColumnPrefix.SelLength = 4000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtColumnPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtHeadRow_Change()
    DataChanged = True
End Sub

Private Sub txtHeadText_Change()
    Dim strInput As String
    Dim lngStart As Long, lngRow As Long, lngCOL As Long
    Dim blnExist As Boolean
    
    strInput = Trim(Me.txtHeadText.Text)
    '检查,如果有相邻的四个单元格的值相同,则不允许设置(有可能需要和左上,左下,右上,右下进行检查)
    lngRow = udHeadRow.Value + 2
    lngCOL = Me.udHeadCol.Value
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 2, lngCOL) = strInput
    
    If lngRow <= 4 Then
        If (vfgThis.TextMatrix(3, lngCOL) = vfgThis.TextMatrix(4, lngCOL) And vfgThis.TextMatrix(3, lngCOL) <> "") Then
            If lngCOL > 1 Then
                '左上
                If vfgThis.TextMatrix(lngRow, lngCOL - 1) = vfgThis.TextMatrix(lngRow, lngCOL) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCOL - 1) = vfgThis.TextMatrix(lngRow, lngCOL) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCOL < vfgThis.Cols - 1 Then
                '右上
                If vfgThis.TextMatrix(lngRow, lngCOL) = vfgThis.TextMatrix(lngRow, lngCOL + 1) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCOL + 1) = vfgThis.TextMatrix(lngRow, lngCOL) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
        End If
    End If
    If lngRow >= 4 Then
        If (vfgThis.TextMatrix(4, lngCOL) = vfgThis.TextMatrix(5, lngCOL) And vfgThis.TextMatrix(4, lngCOL) <> "") Then
            If lngCOL > 1 Then
                '左下
                If vfgThis.TextMatrix(lngRow, lngCOL - 1) = vfgThis.TextMatrix(lngRow, lngCOL) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCOL - 1) = vfgThis.TextMatrix(lngRow, lngCOL) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCOL < vfgThis.Cols - 1 Then
                '右下
                If vfgThis.TextMatrix(lngRow, lngCOL) = vfgThis.TextMatrix(lngRow, lngCOL + 1) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCOL + 1) = vfgThis.TextMatrix(lngRow, lngCOL) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
        End If
    End If
Limit:
    If blnExist Then strInput = strInput & "_1"

WriteIt:
    Me.txtHeadText.Text = strInput
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 2, lngCOL) = strInput
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    DataChanged = True
End Sub

Private Sub txtHeadText_GotFocus()
    Me.txtHeadText.SelStart = 0: Me.txtHeadText.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtHeadText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtHeight_Change()
    DataChanged = True
End Sub

Private Sub txtLabelPrefix_Change()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Me.txtLabelPrefix.Text & Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "{"))
    End With
End Sub

Private Sub txtLabelPrefix_GotFocus()
    Me.txtLabelPrefix.SelStart = 0: Me.txtLabelPrefix.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtLabelPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtMarjin_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txtRecordFrom_Change()
    DataChanged = True
End Sub

Private Sub txtRecordTo_Change()
    DataChanged = True
End Sub

Private Sub txtTabCols_Change()
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_Change()
    Me.vfgThis.RowHeightMin = Val(Me.txtTabRowHeight.Text)
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_GotFocus()
    Me.txtTabRowHeight.SelStart = 0: Me.txtTabRowHeight.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtTabRowHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtTitleText_Change()
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(1, lngCount) = Trim(Me.txtTitleText.Text)
        Next
    End With
    
    DataChanged = True
End Sub

Private Sub txtTitleText_GotFocus()
    Me.txtTitleText.SelStart = 0: Me.txtTitleText.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtTitleText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*()[]{}_+|=-`;'"":/\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtWidth_Change()
    DataChanged = True
End Sub

Private Sub udColumnNo_Change()
    Me.lstColumnUsed.Clear
'    strTemp = Trim(Me.vfgThis.TextMatrix(5, Me.udColumnNo.Value))
'
    strTemp = vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value)
    
    If strTemp = "" Then
        Me.lstColumnUsed.ListIndex = -1
        Me.cmdColumn(1).Enabled = False
        Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
        Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
    Else
        Dim aryCol() As String
        aryCol = Split(strTemp, Space(1))
        For lngCount = 0 To UBound(aryCol)
            If InStr(aryCol(lngCount), "`") > 0 Then
                Me.lstColumnUsed.AddItem Mid(aryCol(lngCount), 1, InStr(aryCol(lngCount), "`") - 1)
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = Val(Mid(aryCol(lngCount), InStr(aryCol(lngCount), "`") + 1))
            Else
                Me.lstColumnUsed.AddItem aryCol(lngCount)
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = 0
            End If
        Next
        Me.lstColumnUsed.ListIndex = 0
        Me.cmdColumn(1).Enabled = True
        Me.txtColumnPrefix.Enabled = True
        Me.txtColumnPostfix.Enabled = True
        
        chk.Enabled = True
        
    End If
End Sub

Private Sub udHeadCol_Change()
    
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    txtHeadText.Text = vfgThis.TextMatrix(udHeadRow.Value + 2, udHeadCol.Value)

    DataChanged = blnSvrChanged
End Sub

Private Sub udHeadRow_Change()

    Call udHeadCol_Change
    
End Sub

Private Sub udRecordFrom_Change()
    If Me.udRecordFrom.Value > Me.udRecordTo.Value Then
        Me.lblRecordTo.Caption = "次日" & Space(7) & "点"
    Else
        Me.lblRecordTo.Caption = "当日" & Space(7) & "点"
    End If
End Sub

Private Sub udRecordTo_Change()
    Call udRecordFrom_Change
End Sub

Private Sub udTabCols_Change()
    Me.vfgThis.Cols = Me.udTabCols.Value + 1
    Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
    Me.udHeadCol.Max = Me.udTabCols.Value
    If Val(Me.txtHeadCol.Text) > Me.udHeadCol.Max Then Me.txtHeadCol.Text = Me.udHeadCol.Max
    Me.udColumnNo.Max = Me.udTabCols.Value
    If Val(Me.txtColumnNo.Text) > Me.udColumnNo.Max Then Me.txtColumnNo.Text = Me.udColumnNo.Max
    
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .TextMatrix(1, lngCount) = .TextMatrix(1, .FixedCols)
            .TextMatrix(2, lngCount) = .TextMatrix(2, .FixedCols)
            .TextMatrix(.Rows - 1, lngCount) = .TextMatrix(.Rows - 1, .FixedCols)
        Next
        '.Cell(flexcpAlignment, 6, 1, 7, .Cols - 1) = flexAlignGeneralCenter
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, Me.shpTabGridColor.BorderColor, 1, 1, 1, 1, 1, 1
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
    End With
End Sub

Private Sub vfgThis_AfterMoveColumn(ByVal Col As Long, Position As Long)
    DataChanged = True
End Sub

Private Sub vfgThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    Me.udHeadCol.Value = NewCol
    Me.udColumnNo.Value = NewCol
    If NewRow >= 3 And NewRow <= 5 Then udHeadRow.Value = NewRow - 2
    
    Dim intAlign As Integer
    Select Case vfgThis.Cell(flexcpAlignment, 6, NewCol)
    Case Is >= 6    '右
        intAlign = 2
    Case Is >= 2    '中
        intAlign = 1
    Case Is <= 2    '左
        intAlign = 0
    End Select
    optAlign(intAlign).Value = True
    
    DataChanged = blnSvrChanged
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    DataChanged = True
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    '编辑处理
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    '编辑处理
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick()
    '编辑处理
    Call mclsVsf.DbClick
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    '编辑处理
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    
    '编辑处理,最后调用
    If KeyAscii <> vbKeyReturn Then
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        Select Case vsf.Col
        Case vsf.ColIndex("开始时间"), vsf.ColIndex("结束时间")
            If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
        End Select
        
        
    End If
    Call mclsVsf.KeyPress(KeyAscii)

End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    
    If KeyAscii <> vbKeyReturn Then
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        Select Case vsf.Col
        Case vsf.ColIndex("开始时间"), vsf.ColIndex("结束时间")
            If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
        End Select

    End If
    Call mclsVsf.KeyPressEdit(KeyAscii)
    
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf.MouseRow, vsf.MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub

Private Sub GetrtbObject()
    If mblnRTBFoot Then
        Set rtbThis = rtbFoot
    Else
        Set rtbThis = rtbHead
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim picTemp As StdPicture
    
    With Me.dlgThis
        .DialogTitle = "标志图选择"
        .Filename = ""
        .Filter = "图像|*.jpg;*.bmp;*.ico;*.gif"
        .CancelError = True
        On Error Resume Next
        .ShowOpen
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
    End With
    Set picTemp = Nothing
    Set picTemp = LoadPicture(Me.dlgThis.Filename)
    If picTemp Is Nothing Then MsgBox "不是有效的图片文件！", vbExclamation, Me.Caption: Exit Sub
    
    Clipboard.Clear
    Clipboard.SetData picTemp
    
    Call GetrtbObject
    SendMessageLong rtbThis.hwnd, WM_PASTE, 0, 0
    mblnChanged = True
End Sub

Private Sub cboFont_Click()
    Call GetrtbObject
    rtbThis.SelFontName = cboFont.List(cboFont.ListIndex)
    mblnChanged = True
End Sub

Private Sub cboFSize_Click()
    Call GetrtbObject
    rtbThis.SelFontSize = GetFontSizeNumber(cboFSize.List(cboFSize.ListIndex))
    mblnChanged = True
End Sub

Private Sub chkB_Click()
    Call GetrtbObject
    If chkB.Value = vbChecked Then
        rtbThis.SelBold = True
    Else
        rtbThis.SelBold = True
    End If
    mblnChanged = True
End Sub

Private Sub chkI_Click()
    Call GetrtbObject
    If chkI.Value = vbChecked Then
        rtbThis.SelItalic = True
    Else
        rtbThis.SelItalic = False
    End If
    mblnChanged = True
End Sub

Private Sub chkU_Click()
    Call GetrtbObject
    If chkU.Value = vbChecked Then
        rtbThis.SelUnderline = True
    Else
        rtbThis.SelUnderline = False
    End If
    mblnChanged = True
End Sub

Private Function PageHeadTest() As Boolean
    '超过上边距返回假
    Dim fr As FORMATRANGE           '格式化的文本范围
    Dim rcDrawTo As RECT            '目标文字区域
    Dim rcPage As RECT              '目标页面区域
    Dim gTargetDC As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
'    Dim lngOffsetWidth As Long
'    Dim lngOffsetHeight As Long
    Dim lngNextPos As Long, lngLen As Long, lngTMP As Long, lngPageCount As Long
    
    lngLen = lstrlen(rtbHead.Text)
    'printer.Duplex = vbPRDPHorizontal
    'printer.ScaleMode = vbTwips
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
'    lngOffsetWidth = Printer.ScaleWidth
'    lngOffsetHeight = Printer.ScaleHeight
    
    gTargetDC = hDC
    With rcPage
        .Left = 0
        .Top = 0
        .Right = Printer.Width
        .Bottom = Printer.Height
    End With
    With rcDrawTo
        .Left = lngOffsetLeft
        .Top = lngOffsetTop
        .Right = Printer.Width - lngOffsetLeft
        .Bottom = Printer.ScaleX(txtMarjin(0).Text, vbMillimeters, vbTwips)
    End With
    With fr
        .hDC = Printer.hDC
        .hdcTarget = gTargetDC
        .rc = rcDrawTo
        .rcPage = rcPage
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    
    Do
        lngNextPos = SendMessage(rtbHead.hwnd, EM_FORMATRANGE, 0, fr)
        
        lngPageCount = lngPageCount + 1             ' 页数＋1
        '记录分页信息
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          '实际打印高度
        AllPages(lngPageCount).Start = lngTMP
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos
        If lngNextPos <= lngTMP Or lngNextPos >= lngLen Then Exit Do      ' 完成所有页面的分页
        lngTMP = lngNextPos
    Loop
    Call SendMessage(rtbHead.hwnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
    If fr.rc.Bottom > rcDrawTo.Bottom Or lngPageCount > 1 Then
        MsgBox "设计的页眉内容超过上边距！", vbInformation, gstrSysName
        Exit Function
    End If
    PageHeadTest = True
End Function

Private Function ReadPageHead(objHead As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, strKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Private Function ReadPageFoot(objFoot As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, strKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

Private Function ReadPageHeadFile(ByVal strKey As String) As String
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(12, strKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageHeadFile = strZip
    End If
End Function

Private Function ReadPageFootFile(ByVal strKey As String) As String
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(13, strKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageFootFile = strZip
    End If
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Private Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
    
    Exit Function
    
errHand:
    Call SaveErrLog
End Function
