VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmOpsStation 
   Caption         =   "手术室工作站"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11760
   Icon            =   "frmOpsStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3435
      Top             =   5895
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5010
      TabIndex        =   32
      ToolTipText     =   "快捷键：F3"
      Top             =   5475
      Width           =   1320
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   8
      Left            =   3195
      ScaleHeight     =   1695
      ScaleWidth      =   8175
      TabIndex        =   10
      Top             =   795
      Width           =   8175
      Begin VB.Frame fra 
         Height          =   1035
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   8430
         Begin VB.PictureBox pic 
            BorderStyle     =   0  'None
            Height          =   870
            Left            =   30
            ScaleHeight     =   870
            ScaleWidth      =   6990
            TabIndex        =   12
            Top             =   135
            Width           =   6990
            Begin VB.PictureBox picState 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   6270
               ScaleHeight     =   360
               ScaleWidth      =   585
               TabIndex        =   13
               Top             =   420
               Visible         =   0   'False
               Width           =   585
               Begin VB.Shape shpState 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   330
                  Left            =   60
                  Top             =   30
                  Width           =   405
               End
               Begin VB.Label lblState 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "收"
                  BeginProperty Font 
                     Name            =   "黑体"
                     Size            =   15
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   300
                  Left            =   90
                  TabIndex        =   14
                  Top             =   30
                  Width           =   315
               End
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   7
               Left            =   3120
               TabIndex        =   36
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   2475
               TabIndex        =   35
               Top             =   45
               Width           =   630
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   6
               Left            =   2115
               TabIndex        =   34
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "床号:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   1635
               TabIndex        =   33
               Top             =   45
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "申请手术:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   11
               Left            =   45
               TabIndex        =   28
               Top             =   615
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "申 请 人:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   12
               Left            =   30
               TabIndex        =   27
               Top             =   345
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "申请科室:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   13
               Left            =   1620
               TabIndex        =   26
               Top             =   345
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "申请时间:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   14
               Left            =   3855
               TabIndex        =   25
               Top             =   345
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病人姓名:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   20
               Left            =   45
               TabIndex        =   24
               Top             =   45
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病人来源:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   21
               Left            =   3855
               TabIndex        =   23
               Top             =   45
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病人科室:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   22
               Left            =   5250
               TabIndex        =   22
               Top             =   45
               Width           =   810
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   0
               Left            =   4680
               TabIndex        =   21
               Top             =   345
               Width           =   90
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   1
               Left            =   870
               TabIndex        =   20
               Top             =   345
               Width           =   90
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   2
               Left            =   2460
               TabIndex        =   19
               Top             =   345
               Width           =   90
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   3
               Left            =   870
               TabIndex        =   18
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   4
               Left            =   4695
               TabIndex        =   17
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   5
               Left            =   6090
               TabIndex        =   16
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   10
               Left            =   870
               TabIndex        =   15
               Top             =   615
               Width           =   90
            End
         End
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2160
      Index           =   7
      Left            =   7140
      ScaleHeight     =   2160
      ScaleWidth      =   4395
      TabIndex        =   8
      Top             =   2670
      Width           =   4395
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   1215
         Left            =   270
         TabIndex        =   9
         Top             =   90
         Width           =   2475
         _Version        =   589884
         _ExtentX        =   4366
         _ExtentY        =   2143
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   3
      Left            =   675
      ScaleHeight     =   1125
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   6150
      Width           =   1935
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   2
         Left            =   555
         TabIndex        =   7
         Top             =   390
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   2
      Left            =   225
      ScaleHeight     =   1125
      ScaleWidth      =   1935
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   1
         Left            =   390
         TabIndex        =   6
         Top             =   180
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1005
      Index           =   1
      Left            =   210
      ScaleHeight     =   1005
      ScaleWidth      =   2565
      TabIndex        =   2
      Top             =   4065
      Width           =   2565
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3390
      Index           =   0
      Left            =   120
      ScaleHeight     =   3390
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   405
      Width           =   3000
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   675
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   60
         Width           =   2130
      End
      Begin XtremeSuiteControls.TabControl tbc 
         Height          =   2550
         Left            =   105
         TabIndex        =   1
         Top             =   420
         Width           =   2820
         _Version        =   589884
         _ExtentX        =   4974
         _ExtentY        =   4498
         _StockProps     =   64
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手术室"
         Height          =   180
         Left            =   30
         TabIndex        =   30
         Top             =   105
         Width           =   540
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   7455
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14896
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmOpsStation.frx":1CFA
      Left            =   615
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmOpsStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'窗体级变量定义
'######################################################################################################################
Private mstrPrivs As String
Private mblnDataChanged As Boolean
Private mblnStartUp As Boolean
Private mblnAllowClose As Boolean
Private mlngSvrDept As Long
Private mintIndex As Integer
Private mlng手术部门id As Long
Private mstr手术部门 As String
Private mstrCondition As String
Private mstrFindKey As String
Private mlngTmp As Long
Private mlngCountTmr As Long
Private mobjFindKey As CommandBarControl
Private mclsVsf(2) As clsVsf

Private WithEvents mfrmChildStationOutLine As frmChildStationOutLine
Attribute mfrmChildStationOutLine.VB_VarHelpID = -1
Private WithEvents mfrmChildStationDrug As frmChildStationDrug
Attribute mfrmChildStationDrug.VB_VarHelpID = -1
Private WithEvents mfrmChildStationCure As frmChildStationCure
Attribute mfrmChildStationCure.VB_VarHelpID = -1
Private WithEvents mfrmChildStationMaterial As frmChildStationMaterial
Attribute mfrmChildStationMaterial.VB_VarHelpID = -1
Private WithEvents mfrmChildStationInEPR As frmChildStationInEPR
Attribute mfrmChildStationInEPR.VB_VarHelpID = -1
Private WithEvents mfrmChildStationOutEPR As frmChildStationOutEPR
Attribute mfrmChildStationOutEPR.VB_VarHelpID = -1
Private WithEvents mclsInAdvices As zlCISKernel.clsDockInAdvices
Attribute mclsInAdvices.VB_VarHelpID = -1
Private WithEvents mclsOutAdvices As zlCISKernel.clsDockOutAdvices
Attribute mclsOutAdvices.VB_VarHelpID = -1
Private WithEvents mclsExpenses As zlCISKernel.clsDockExpense
Attribute mclsExpenses.VB_VarHelpID = -1
'
''自定义过程或函数
''######################################################################################################################
'

Private Property Let AutoRefresh(vData As Boolean)
    '
    '功能:自动刷新
    '
    tmr.Enabled = vData
    
    If vData = True Then
        mlngCountTmr = 0
        tmr.Tag = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "自动刷新间隔", 0))
        tmr.Enabled = (Val(tmr.Tag) > 0)
    End If
End Property

Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildStationOutLine.DataChanged = blnData
    mfrmChildStationDrug.DataChanged = blnData
    mfrmChildStationCure.DataChanged = blnData
    mfrmChildStationMaterial.DataChanged = blnData

    If mfrmChildStationOutLine.DataChanged Or mfrmChildStationDrug.DataChanged Or mfrmChildStationMaterial.DataChanged And mfrmChildStationCure.DataChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmChildStationOutLine Is Nothing) And Not (mfrmChildStationDrug Is Nothing) And Not (mfrmChildStationMaterial Is Nothing) And Not (mfrmChildStationCure Is Nothing) Then
        DataChanged = mfrmChildStationOutLine.DataChanged Or mfrmChildStationDrug.DataChanged Or mfrmChildStationMaterial.DataChanged Or mfrmChildStationCure.DataChanged
    End If
End Property

Private Function ExchangeAdvice(ByVal blnClinic As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：切换显示门诊/住院医嘱页卡
    '参数：blnClinic=是否显示门诊医嘱页
    '返回：是否进行了切换选择
    '******************************************************************************************************************
    Dim blnSel As Boolean
    Dim blnOld As Boolean
    Dim intIdx As Integer

    If Not tbcPage.Selected Is Nothing Then
        blnSel = tbcPage.Selected.Tag Like "*医嘱"
    End If

    For intIdx = 0 To tbcPage.ItemCount - 1
        If tbcPage(intIdx).Tag = "门诊医嘱" Then
            If tbcPage(intIdx).Visible <> blnClinic Then
                tbcPage(intIdx).Visible = blnClinic
                If blnSel And blnClinic Then
                    tbcPage(intIdx).Selected = True
                    ExchangeAdvice = True
                End If
            End If
        ElseIf tbcPage(intIdx).Tag = "住院医嘱" Then
            If tbcPage(intIdx).Visible <> Not blnClinic Then
                tbcPage(intIdx).Visible = Not blnClinic
                If blnSel And Not blnClinic Then
                    tbcPage(intIdx).Selected = True
                    ExchangeAdvice = True
                End If
            End If
        End If
    Next
End Function

Private Function ExchangeEPRS(ByVal blnClinic As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：切换显示门诊/住院病历页卡
    '参数：blnClinic=是否显示门诊病历页
    '返回：是否进行了切换选择
    '******************************************************************************************************************
    Dim blnSel As Boolean
    Dim blnOld As Boolean
    Dim intIdx As Integer

    If Not tbcPage.Selected Is Nothing Then
        blnSel = tbcPage.Selected.Tag Like "*病历"
    End If

    For intIdx = 0 To tbcPage.ItemCount - 1
        If tbcPage(intIdx).Tag = "门诊病历" Then
            If tbcPage(intIdx).Visible <> blnClinic Then
                tbcPage(intIdx).Visible = blnClinic
                If blnSel And blnClinic Then
                    tbcPage(intIdx).Selected = True
                    ExchangeEPRS = True
                End If
            End If
        ElseIf tbcPage(intIdx).Tag = "住院病历" Then
            If tbcPage(intIdx).Visible <> Not blnClinic Then
                tbcPage(intIdx).Visible = Not blnClinic
                If blnSel And Not blnClinic Then
                    tbcPage(intIdx).Selected = True
                    ExchangeEPRS = True
                End If
            End If
        End If
    Next
End Function

Private Sub zlDefCommandBars(ByVal strMenuKind As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Select Case strMenuKind
    '------------------------------------------------------------------------------------------------------------------
    Case "概要", "药品", "材料", "治疗"
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", objMenu.Index + 1, False)
        objMenu.ID = conMenu_EditPopup
        Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "保存更改(&S)", True)
        Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消更改(&C)")
    End Select


    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain(2)

    For Each objControl In objBar.Controls  '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "保存", True, , , , objControl.Index + 1)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消", , , , , objControl.Index + 1)

    '命令的快键绑定:公共部份主界面已处理
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save              '保存
    End With

End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsMain.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsMain.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
    '******************************************************************************************************************
    '功能：刷新子窗体菜单及工具条
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim cbrCustom As CommandBarControlCustom
    
    On Error GoTo errHand
    
    '记录现有菜单样式
    '------------------------------------------------------------------------------------------------------------------
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        blnShowBar = cbsMain(2).Visible
        bytStyle = cbsMain(2).Controls(1).STYLE
    End If

    '刷新子窗口菜单
    '------------------------------------------------------------------------------------------------------------------
    Call LockWindowUpdate(Me.hWnd)

    '删除现在的工具栏及顶级菜单项
    '------------------------------------------------------------------------------------------------------------------
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    
    For lngCount = cbsMain.Count To 2 Step -1
        If lngCount <> 3 Then
            cbsMain(lngCount).Controls.DeleteAll
        End If
    Next

    '主窗口重新加入
    '------------------------------------------------------------------------------------------------------------------

    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap

    '文件
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)", , , "预览数据")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)", , , "打印数据")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel", , , "输出到Excel")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BillPrintView, "预览通知单", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BillPrint, "打印通知单")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "参数设置(&M)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Option, "执行间设置(&R)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)


    '手术
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "执行(&C)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Request, "补录登记(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Audit, "手术审核(&K)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_UnAudit, "取消审核(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Arrange, "手术安排(&M)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_UnArrange, "取消安排(&X)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Plan, "执行报到(&L)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Logout, "取消报到(&R)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Complete, "执行完成(&I)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Undone, "取消完成(&U)")

    '查看
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Filter, "过滤(&F)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)")

    '帮助
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & ParamInfo.产品名称)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.产品名称 & "主页(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.产品名称 & "论坛(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)

    '工具栏定义:包括公共部份
    '------------------------------------------------------------------------------------------------------------------
    If cbsMain.Count < 2 Then
        Set objBar = cbsMain.Add("标准", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagHideWrap
    Else
        Set objBar = cbsMain(2)
    End If
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Audit, "审核", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Arrange, "安排")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Plan, "报到")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Complete, "完成")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")

    '定位工具栏
    '------------------------------------------------------------------------------------------------------------------
    If cbsMain.Count < 3 Then
        Set objExtendedBar = cbsMain.Add("定位", xtpBarTop)
        
        objExtendedBar.ContextMenuPresent = False
        objExtendedBar.ShowTextBelowIcons = False
        objExtendedBar.EnableDocking xtpFlagHideWrap

        mstrFindKey = Trim(GetRegister(私有模块, Me.Name, "定位依据", "姓名"))
        If mstrFindKey = "" Then mstrFindKey = "姓名"

        Set mobjFindKey = NewToolBar(objExtendedBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , , "快捷键:F4")
        mobjFindKey.IconId = conMenu_View_Find

        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.姓名"): objControl.Parameter = "姓名"
        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.床号"): objControl.Parameter = "床号"
        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&3.住院号"): objControl.Parameter = "住院号"
        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&4.门诊号"): objControl.Parameter = "门诊号"
        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&5.手术名称"): objControl.Parameter = "手术名称"
        
        Set cbrCustom = NewToolBar(objExtendedBar, xtpControlCustom, conMenu_View_Location, "")
        cbrCustom.Handle = txtLocation.hWnd
        
        Set objControl = NewToolBar(objExtendedBar, xtpControlButton, conMenu_View_Forward, "前一条", , , , "快捷键:Ctrl+Left")
        Set objControl = NewToolBar(objExtendedBar, xtpControlButton, conMenu_View_Backward, "后一条", , , , "快捷键:Ctrl+Right")

        Call SetDockRight(objExtendedBar, objBar)
    End If
    
    '命令的快键绑定:公共部份主界面已处理
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF12, conMenu_File_Parameter        '参数设置
        .Add 0, vbKeyF5, conMenu_View_Refresh           '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help              '帮助
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save              '保存
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add 0, vbKeyF4, conMenu_View_Option                '选择定位依据
        .Add 0, vbKeyF6, conMenu_View_Jump           '
        .Add FCONTROL, vbKeyP, conMenu_File_Print       '打印
        .Add FCONTROL, vbKeyV, conMenu_File_Preview
        .Add FCONTROL, vbKeyN, conMenu_Manage_Request     '申请
        .Add FCONTROL, vbKeyK, conMenu_Manage_Audit     '审核
        .Add FCONTROL, vbKeyM, conMenu_Manage_Arrange     '安排
        .Add FCONTROL, vbKeyL, conMenu_Manage_Plan     '报到
        .Add FCONTROL, vbKeyI, conMenu_Manage_Complete     '完成
        .Add FCONTROL, vbKeyF, conMenu_View_Filter        '过滤
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
    End With
    

    '子窗口重新加入
    '------------------------------------------------------------------------------------------------------------------
    Select Case objItem.Tag
    Case "概要", "药品", "材料", "治疗"
        Call zlDefCommandBars(objItem.Tag)
    Case "附费"
        Call mclsExpenses.zlDefCommandBars(Me, cbsMain)
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_EditPopup)
'        Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Edit_Preferences, "方案费用生成(R)", True)
    Case "门诊医嘱"
        Call mclsOutAdvices.zlDefCommandBars(Me, cbsMain, 2)
    Case "住院医嘱"
        Call mclsInAdvices.zlDefCommandBars(Me, cbsMain, 2)
    Case "门诊病历"
        Call mfrmChildStationOutEPR.zlDefCommandBars(cbsMain)
    Case "住院病历"
        Call mfrmChildStationInEPR.zlDefCommandBars(cbsMain)
    End Select

    '恢复及固定的一些菜单设置
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel And objControl.Type <> xtpControlEdit Then
                objControl.STYLE = bytStyle
            End If
        Next
        cbsMain(lngCount).Visible = blnShowBar
    Next

    '如果用了RecalcLayout反而不正常
    '------------------------------------------------------------------------------------------------------------------
    Call LockWindowUpdate(0)
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    objPane.Title = "病人手术列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable

    Set objPane = dkpMain.CreatePane(2, 350, 300, DockRightOf, Nothing)
    objPane.Title = "申请"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 350, 150, DockBottomOf, objPane)
    objPane.Title = "业务"
    objPane.Options = PaneNoCaption

    dkpMain.SetCommandBars cbsMain
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.HideClient = True
End Sub

Private Sub InitTabControl()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************

    '左边的手术的状态分类页面
    '------------------------------------------------------------------------------------------------------------------
    With tbc
        With .PaintManager

            .Appearance = xtpTabAppearanceVisio
            .ClientFrame = xtpTabFrameSingleLine
            .COLOR = xtpTabColorOffice2003
            .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
            .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
            .ShowIcons = True
        End With
        Set .Icons = frmPubIcons.imgPublic.Icons

        .InsertItem 0, "等待手术", picPane(1).hWnd, 0
        .InsertItem 1, "正在手术", picPane(2).hWnd, 0
        .InsertItem 2, "已完手术", picPane(3).hWnd, 0

        .Item(0).Selected = True

    End With

    '右边的具体内容页面
    '------------------------------------------------------------------------------------------------------------------
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
        End With
        Set .Icons = frmPubIcons.imgPublic.Icons

        If mfrmChildStationDrug Is Nothing Then Set mfrmChildStationDrug = New frmChildStationDrug
        If mfrmChildStationMaterial Is Nothing Then Set mfrmChildStationMaterial = New frmChildStationMaterial
        If mfrmChildStationOutLine Is Nothing Then Set mfrmChildStationOutLine = New frmChildStationOutLine
        If mfrmChildStationCure Is Nothing Then Set mfrmChildStationCure = New frmChildStationCure
        
        If mfrmChildStationInEPR Is Nothing Then Set mfrmChildStationInEPR = New frmChildStationInEPR
        If mfrmChildStationOutEPR Is Nothing Then Set mfrmChildStationOutEPR = New frmChildStationOutEPR

        Call mfrmChildStationInEPR.InitData(Me, False)
        Call mfrmChildStationOutEPR.InitData(Me, False)
        
        Set mclsOutAdvices = New zlCISKernel.clsDockOutAdvices
        Set mclsInAdvices = New zlCISKernel.clsDockInAdvices
        Set mclsExpenses = New zlCISKernel.clsDockExpense

        .InsertItem(0, "基本信息", mfrmChildStationOutLine.hWnd, 0).Tag = "概要"
        .InsertItem(1, "用药单", mfrmChildStationDrug.hWnd, 0).Tag = "药品"
        .InsertItem(2, "材料单", mfrmChildStationMaterial.hWnd, 0).Tag = "材料"
        .InsertItem(3, "治疗单", mfrmChildStationCure.hWnd, 0).Tag = "治疗"
        
        If GetInsidePrivs(p医嘱附费管理, True) <> "" Then
            .InsertItem(4, "费用 ", mclsExpenses.zlGetForm.hWnd, 0).Tag = "附费"
        End If

        If GetInsidePrivs(p门诊医嘱下达, True) <> "" Then
            .InsertItem(5, " 医嘱 ", mclsOutAdvices.zlGetForm.hWnd, 0).Tag = "门诊医嘱"
        End If

        If GetInsidePrivs(p住院医嘱下达, True) <> "" Then
            .InsertItem(6, " 医嘱 ", mclsInAdvices.zlGetForm.hWnd, 0).Tag = "住院医嘱"
        End If

        .InsertItem(7, " 病历 ", mfrmChildStationInEPR.hWnd, 0).Tag = "住院病历"
        .InsertItem(8, " 病历 ", mfrmChildStationOutEPR.hWnd, 0).Tag = "门诊病历"

        Call mfrmChildStationOutLine.InitData(Me, False)
        Call mfrmChildStationDrug.InitData(Me, False)
        Call mfrmChildStationMaterial.InitData(Me, False)
        Call mfrmChildStationCure.InitData(Me, False)
        
        .Item(0).Selected = True

        Call SubWinDefCommandBar(.Item(0))

    End With
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    Dim intCount As Integer
    Dim int类型 As Integer
    Dim intTmp As Integer
    Dim blnZero As Boolean

    On Error GoTo errHand
    
    AutoRefresh = False
    
    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        For intCount = 0 To 2
            Set mclsVsf(intCount) = New clsVsf
            With mclsVsf(intCount)
                Call .Initialize(Me.Controls, vsf(intCount), True, True, frmPubResource.GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("图标", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
                Call .AppendColumn("紧急标志", 255, flexAlignCenterCenter, flexDTString, "", "[紧急标志]", False)
                Call .AppendColumn("手术名称", 1800, flexAlignLeftCenter, flexDTString, "", "医嘱内容", True)
                Call .AppendColumn("姓名", 990, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("病人来源", 990, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("病人科室", 1080, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("申请时间", 1680, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", "开嘱时间", True)
                Call .AppendColumn("床号", 600, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("住院号", 810, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("门诊号", 810, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("发送号", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)

                Call .AppendColumn("当前病区id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("当前科室id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("医嘱id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("状态", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("挂号单", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("执行状态", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("出院日期", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("手术状态", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("诊疗项目id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", "", True)
                .AppendRows = True

            End With
        Next
        
        Call InitCommandBar
        Call InitDockPannel
        Call InitTabControl

        fra.BackColor = COLOR_NativeXpPlain.SpecialGroupClient
        pic.BackColor = COLOR_NativeXpPlain.SpecialGroupClient
        picState.BackColor = COLOR_NativeXpPlain.SpecialGroupClient
        picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        mlngSvrDept = 0
        mintIndex = 0

        strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & UserInfo.数据库用户 & "\" & App.ProductName & "\" & Me.Name, "待手术时间范围", "今  天"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & UserInfo.数据库用户 & "\" & App.ProductName & "\" & Me.Name, "待手术时间范围", "今  天"), 2)
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
        
        strEnd = Format(strEnd, "yyyy-MM-dd") & " 23:59:59"
        
        mstrCondition = strStart & ";" & strEnd & ";;;;;;"

        strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & UserInfo.数据库用户 & "\" & App.ProductName & "\" & Me.Name, "已完手术时间范围", "今  天"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & UserInfo.数据库用户 & "\" & App.ProductName & "\" & Me.Name, "已完手术时间范围", "今  天"), 2)
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
        strEnd = Format(strEnd, "yyyy-MM-dd") & " 23:59:59"
        mstrCondition = mstrCondition & ";" & strStart & ";" & strEnd & ";0;"

        '获取体检部门
        '--------------------------------------------------------------------------------------------------------------
        cboDept.Clear
        If IsPrivs(mstrPrivs, "所有手术室") Then
            gstrSQL = GetPublicSQL(SQL.手术部门清单, "所有")
        Else
            gstrSQL = GetPublicSQL(SQL.手术部门清单, "")
        End If

        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID)
        If rs.BOF = False Then
            Do While Not rs.EOF
                cboDept.AddItem rs("编码").Value & " - " & rs("名称").Value
                cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
                rs.MoveNext
            Loop
        Else
            ShowSimpleMsg "没有手术部门信息，请在人员管理中检查人员的科室属性及权限！"
            GoTo errEnd
        End If

        On Error Resume Next
        zlControl.CboLocate cboDept, UserInfo.部门ID, True
        On Error GoTo errHand
        If cboDept.ListIndex < 0 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0

        mlng手术部门id = cboDept.ItemData(cboDept.ListIndex)
        mstr手术部门 = zlCommFun.GetNeedName(cboDept.List(cboDept.ListIndex))

    '--------------------------------------------------------------------------------------------------------------
    Case "控件状态"

        If tbc.Enabled <> Not DataChanged Then
            tbc.Enabled = Not DataChanged
            vsf(0).Enabled = Not DataChanged
            vsf(0).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
            vsf(1).Enabled = Not DataChanged
            vsf(1).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
            vsf(2).Enabled = Not DataChanged
            vsf(2).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
        End If
        stbThis.Panels(3).Enabled = DataChanged

    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"

        Call ExecuteCommand("装载等待手术")
        Call ExecuteCommand("装载正在手术")
        Call ExecuteCommand("装载已完手术")

        With vsf(mintIndex)
            Call vsf_AfterRowColChange(mintIndex, 0, 0, .Row, .Col)
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "本地参数设置"

        ExecuteCommand = frmOpsStationPara.ShowPara(Me)
        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "手术间设置"

        ExecuteCommand = frmOpsStationRoom.ShowEdit(Me, mlng手术部门id)
        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "直接手术申请"
        ExecuteCommand = frmOpsStationRequest.ShowEdit(Me, mlng手术部门id)
        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "读取申请信息"

        gstrSQL = "SELECT Distinct B.当前床号,b.门诊号,b.住院号,A.ID,B.姓名,C.发送号,A.医嘱内容,A.开嘱医生,A.开嘱时间,Decode(A.病人来源,1,'门诊',2,'住院',3,'外来') AS 来源,Decode(A.病人来源,1,B.就诊诊室,2,E.名称) AS 病人科室,G.名称 AS 申请科室 " & _
                "FROM 病人医嘱记录 A, 病人信息 B,病人医嘱发送 C,部门表 E,部门表 G " & _
                "Where A.病人id = B.病人id AND A.ID=C.医嘱id(+) AND E.ID(+)=B.当前科室ID  " & _
                    "AND A.ID=[1] AND A.开嘱科室ID=G.ID "

        With vsf(mintIndex)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(.TextMatrix(.Row, .ColIndex("医嘱id"))))
            If rs.BOF = False Then

                lblValue(1).Caption = zlCommFun.NVL(rs("开嘱医生"))
                lblValue(0).Caption = Format(zlCommFun.NVL(rs("开嘱时间")), "YYYY-MM-DD HH:MM")
                lblValue(2).Caption = zlCommFun.NVL(rs("申请科室"))
                lblValue(3).Caption = zlCommFun.NVL(rs("姓名"))
                lblValue(4).Caption = zlCommFun.NVL(rs("来源"))
                lblValue(5).Caption = zlCommFun.NVL(rs("病人科室"))
                lblValue(6).Caption = zlCommFun.NVL(rs("当前床号"))
                If zlCommFun.NVL(rs("来源")) = "住院" Then
                    lbl(1).Caption = "住院号"
                    lblValue(7).Caption = zlCommFun.NVL(rs("住院号"))
                Else
                    lbl(1).Caption = "门诊号"
                    lblValue(7).Caption = zlCommFun.NVL(rs("门诊号"))
                End If
                
                lblValue(10).Caption = zlCommFun.NVL(rs("医嘱内容"))

                picState.Visible = CheckChargeState(Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), zlCommFun.NVL(rs("发送号").Value, 0))
            Else
                picState.Visible = False

                lblValue(0).Caption = ""
                lblValue(1).Caption = ""
                lblValue(2).Caption = ""
                lblValue(3).Caption = ""
                lblValue(4).Caption = ""
                lblValue(5).Caption = ""
                lblValue(6).Caption = ""
                lblValue(7).Caption = ""
                lblValue(10).Caption = ""
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "装载等待手术"

        mclsVsf(0).ClearGrid

        strTmp = ""

        If Split(mstrCondition, ";")(0) <> "" Then
            strTmp = " AND a.开嘱时间 BETWEEN [2] AND [3] "
        End If

        If Trim(Split(mstrCondition, ";")(2)) <> "" Then
            strTmp = strTmp & " AND b.姓名 LIKE [4] "
        End If

        If Trim(Split(mstrCondition, ";")(3)) <> "" Then
            strTmp = strTmp & " AND b.住院号 = [5] "
        End If

        If Trim(Split(mstrCondition, ";")(4)) <> "" Then
            strTmp = strTmp & " AND b.当前床号 = [6] "
        End If

        If Trim(Split(mstrCondition, ";")(5)) <> "" Then
            strTmp = strTmp & " AND b.门诊号 = [7] "
        End If

        If Val(Trim(Split(mstrCondition, ";")(7))) > 0 Then
            strTmp = strTmp & " AND a.诊疗项目ID = [8] "
        End If

        gstrSQL = "Select   Decode(e.ID,Null,-100,e.ID) As ID,Decode(e.手术状态,1,'审核',2,'安排',3,'手术',4,'完成','申请') As 图标,a.Id As 医嘱id," & vbNewLine & _
                    "       Decode(a.紧急标志,1,'紧急','') As 紧急标志," & vbNewLine & _
                    "       DECODE(a.病人来源,1,'门诊',2,'住院',4,'体检','外来') AS 病人来源," & vbNewLine & _
                    "       Decode(a.诊疗项目id,Null,a.医嘱内容,f.名称) As 医嘱内容," & vbNewLine & _
                    "       a.开嘱时间," & vbNewLine & _
                    "       b.姓名," & vbNewLine & _
                    "       b.门诊号," & vbNewLine & _
                    "       b.住院号,b.当前床号 As 床号," & vbNewLine & _
                    "       c.名称 As 病人科室," & vbNewLine & _
                    "       d.名称 As 开单科室," & vbNewLine & _
                    "       a.开嘱医生 As 开单人," & vbNewLine & _
                    "       a.医嘱状态," & vbNewLine & _
                    "       a.病人id," & vbNewLine & _
                    "       a.主页id," & vbNewLine & _
                    "       a.诊疗项目id," & vbNewLine & _
                    "       e.手术状态,g.发送号,g.执行状态,a.挂号单,0 As 状态,b.出院时间 As 出院日期,b.当前病区id,b.当前科室id "
        gstrSQL = gstrSQL & _
                    "From 病人医嘱记录 a," & vbNewLine & _
                    "     病人信息 b," & vbNewLine & _
                    "     部门表 c," & vbNewLine & _
                    "     部门表 d," & vbNewLine & _
                    "     病人手术记录 e,诊疗项目目录 f,病人医嘱发送 g " & vbNewLine & _
                    "Where (a.诊疗类别='F' Or a.诊疗类别 Is Null)" & vbNewLine & _
                    "      And a.相关id Is Null" & vbNewLine & _
                    "      And a.医嘱状态<>4 " & vbNewLine & strTmp & _
                    "      And a.执行科室id+0=[1]" & vbNewLine & _
                    "      And b.病人id=a.病人id" & vbNewLine & _
                    "      And c.Id=a.病人科室id" & vbNewLine & _
                    "      And d.Id=a.开嘱科室id And f.Id(+)=a.诊疗项目id " & vbNewLine & _
                    "      And a.Id=e.医嘱id(+)  And Nvl(e.手术状态,0)<=2 And a.ID=g.医嘱id(+) "


        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng手术部门id, _
                                                                CDate(Split(mstrCondition, ";")(0)), _
                                                                CDate(Split(mstrCondition, ";")(1)), _
                                                                "%" & Trim(Split(mstrCondition, ";")(2)) & "%", _
                                                                IIf(IsNumeric(Split(mstrCondition, ";")(3)), Split(mstrCondition, ";")(3), "0"), _
                                                                CStr(Trim(Split(mstrCondition, ";")(4))), _
                                                                IIf(IsNumeric(Split(mstrCondition, ";")(5)), Split(mstrCondition, ";")(5), "0"), _
                                                                Val(Trim(Split(mstrCondition, ";")(7))))
        If rs.BOF = False Then Call mclsVsf(0).LoadGrid(rs)

    '------------------------------------------------------------------------------------------------------------------
    Case "装载正在手术"

        mclsVsf(1).ClearGrid

        gstrSQL = GetPublicSQL(SQL.正在手术记录, mstrCondition)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng手术部门id, _
                                                                "%" & Trim(Split(mstrCondition, ";")(2)) & "%", _
                                                                IIf(IsNumeric(Split(mstrCondition, ";")(3)), Split(mstrCondition, ";")(3), "0"), _
                                                                CStr(Trim(Split(mstrCondition, ";")(4))), _
                                                                IIf(IsNumeric(Split(mstrCondition, ";")(5)), Split(mstrCondition, ";")(5), "0"), _
                                                                Val(Trim(Split(mstrCondition, ";")(7))), CStr(Trim(Split(mstrCondition, ";")(11))))
        If rs.BOF = False Then Call mclsVsf(1).LoadGrid(rs)

    '------------------------------------------------------------------------------------------------------------------
    Case "装载已完手术"

        mclsVsf(2).ClearGrid

        gstrSQL = GetPublicSQL(SQL.手术申请记录, mstrCondition)

        If Val(Trim(Split(mstrCondition, ";")(10))) = 1 Then
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng手术部门id, _
                                                                    4, _
                                                                    CDate(Split(mstrCondition, ";")(8)), _
                                                                    CDate(Split(mstrCondition, ";")(9)), _
                                                                    "%" & Trim(Split(mstrCondition, ";")(2)) & "%", _
                                                                    IIf(IsNumeric(Split(mstrCondition, ";")(3)), Split(mstrCondition, ";")(3), "0"), _
                                                                    CStr(Trim(Split(mstrCondition, ";")(4))), _
                                                                    IIf(IsNumeric(Split(mstrCondition, ";")(5)), Split(mstrCondition, ";")(5), "0"), _
                                                                    Val(Trim(Split(mstrCondition, ";")(7))))
        Else
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng手术部门id, _
                                                                    4, _
                                                                    CDate(Split(mstrCondition, ";")(0)), _
                                                                    CDate(Split(mstrCondition, ";")(1)), _
                                                                    "%" & Trim(Split(mstrCondition, ";")(2)) & "%", _
                                                                    IIf(IsNumeric(Split(mstrCondition, ";")(3)), Split(mstrCondition, ";")(3), "0"), _
                                                                    CStr(Trim(Split(mstrCondition, ";")(4))), _
                                                                    IIf(IsNumeric(Split(mstrCondition, ";")(5)), Split(mstrCondition, ";")(5), "0"), _
                                                                    Val(Trim(Split(mstrCondition, ";")(7))))
        End If
        If rs.BOF = False Then Call mclsVsf(2).LoadGrid(rs)

    '------------------------------------------------------------------------------------------------------------------
    Case "刷新手术数据"

        Call ExecuteCommand("读取手术概要")
        Call ExecuteCommand("读取手术用药")
        Call ExecuteCommand("读取手术材料")
        Call ExecuteCommand("读取手术治疗")
        Call ExecuteCommand("读取手术费用")
        Call ExecuteCommand("读取病人医嘱")
        Call ExecuteCommand("读取手术病历")

    '------------------------------------------------------------------------------------------------------------------
    Case "读取手术概要"

        With vsf(mintIndex)
            Call mfrmChildStationOutLine.RefreshData(Val(.RowData(.Row)), (Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3 And IsPrivs(mstrPrivs, "概要登记")), mlng手术部门id)
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "读取手术用药"

        With vsf(mintIndex)
            Call mfrmChildStationDrug.RefreshData(Val(.RowData(.Row)), _
                                        (Val(.TextMatrix(.Row, .ColIndex("手术状态"))) > 1 And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) < 4), _
                                        IIf(Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 2, 1, 2), _
                                        .TextMatrix(.Row, .ColIndex("病人来源")), _
                                        Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), _
                                        mstrPrivs)
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "读取手术材料"
        With vsf(mintIndex)
            Call mfrmChildStationMaterial.RefreshData(Val(.RowData(.Row)), (Val(.TextMatrix(.Row, .ColIndex("手术状态"))) > 1 And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) < 4), IIf(Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 2, 1, 2), .TextMatrix(.Row, .ColIndex("病人来源")), _
                                        Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), _
                                        mstrPrivs)
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "读取手术治疗"
        With vsf(mintIndex)
            Call mfrmChildStationCure.RefreshData(Val(.RowData(.Row)), (Val(.TextMatrix(.Row, .ColIndex("手术状态"))) > 1 And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) < 4), IIf(Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 2, 1, 2), .TextMatrix(.Row, .ColIndex("病人来源")), _
                                        Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), _
                                        mstrPrivs)
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取手术费用"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("发送号"))) > 0 Then
                Call mclsExpenses.zlRefresh(mlng手术部门id, Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), Val(.TextMatrix(.Row, .ColIndex("发送号"))), False)
            Else
                Call mclsExpenses.zlRefresh(0, 0, 0)
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "读取病人医嘱"

        With vsf(mintIndex)
            Select Case .TextMatrix(.Row, .ColIndex("病人来源"))
            Case "门诊"
                If Val(.TextMatrix(.Row, .ColIndex("病人id"))) > 0 Then
                    Call mclsOutAdvices.zlRefresh(Val(.TextMatrix(.Row, .ColIndex("病人id"))), .TextMatrix(.Row, .ColIndex("挂号单")), Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3, False, 0)
                Else
                    Call mclsOutAdvices.zlRefresh(0, "", False)
                End If
            Case "住院"
                If Val(.TextMatrix(.Row, .ColIndex("病人id"))) > 0 Then

                    If .TextMatrix(.Row, .ColIndex("出院日期")) = "" Then
                        If Val(.TextMatrix(.Row, .ColIndex("状态"))) <> 2 Then
                            int类型 = 0 '在院
                        Else
                            int类型 = 1 '预出院
                        End If
                    Else
                        int类型 = 2 '出院
                    End If
                    Call mclsInAdvices.zlRefresh(Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), Val(.TextMatrix(.Row, .ColIndex("当前病区id"))), Val(.TextMatrix(.Row, .ColIndex("当前科室id"))), IIf(Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 3, 4, int类型), False, 0, Val(.TextMatrix(.Row, .ColIndex("执行状态"))))
                Else
                    Call mclsInAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
                End If
            End Select
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "读取手术病历"

        With vsf(mintIndex)
            Select Case .TextMatrix(.Row, .ColIndex("病人来源"))
            Case "门诊"

                gstrSQL = "Select ID From 病人挂号记录 Where No=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .TextMatrix(.Row, .ColIndex("挂号单")))
                If rs.BOF = False Then
                    Call mfrmChildStationOutEPR.RefreshData(Val(.RowData(.Row)), Val(.TextMatrix(.Row, .ColIndex("病人id"))), _
                                                            Val(rs("ID").Value), _
                                                            mlng手术部门id, _
                                                            Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), _
                                                            (Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3))
                Else
                    Call mfrmChildStationOutEPR.RefreshData(0, 0, 0, 0, 0, (Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3))
                End If

            Case "住院"

                Call mfrmChildStationInEPR.RefreshData(Val(.RowData(.Row)), Val(.TextMatrix(.Row, .ColIndex("病人id"))), _
                                                        Val(.TextMatrix(.Row, .ColIndex("主页id"))), _
                                                        mlng手术部门id, _
                                                        Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), _
                                                        (Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3))
            End Select
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "审核手术申请"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 0 Then GoTo errEnd
            ExecuteCommand = frmOpsStationAuditing.ShowEdit(Me, Val(.TextMatrix(.Row, .ColIndex("医嘱id"))))
        End With

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "取消手术审核"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 1 Then GoTo errEnd

            If MsgBox("真的要撤消“" & .TextMatrix(.Row, .ColIndex("手术名称")) & "”审核吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "ZL_病人诊断记录_DELETE2(" & Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) & ",8)"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            gstrSQL = "zl_病人手术记录_AduitCancel(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "安排手术申请"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 1 Then GoTo errEnd
            ExecuteCommand = frmOpsStationArrange.ShowEdit(Me, Val(.RowData(.Row)), mlng手术部门id)
        End With

        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "取消手术安排"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 2 Then GoTo errEnd

            If MsgBox("真的要撤消“" & .TextMatrix(.Row, .ColIndex("手术名称")) & "”的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_病人手术记录_ArrangeCancel(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "手术报到"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 2 Then GoTo errEnd

            If MsgBox("真的要报到“" & .TextMatrix(.Row, .ColIndex("手术名称")) & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_病人手术记录_Register(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "取消报到"
        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Or mintIndex <> 1 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 3 Then GoTo errEnd

            If MsgBox("真的要取消报到“" & .TextMatrix(.Row, .ColIndex("手术名称")) & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_病人手术记录_RegisterCancel(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "手术完成"
        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Or mintIndex <> 1 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 3 Then GoTo errEnd

            If MsgBox("真的要完成“" & .TextMatrix(.Row, .ColIndex("手术名称")) & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_病人手术记录_Complete(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "取消完成"
        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Or mintIndex <> 2 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("手术状态"))) <> 4 Then GoTo errEnd

            If MsgBox("真的要取消完成“" & .TextMatrix(.Row, .ColIndex("手术名称")) & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_病人手术记录_CompleteCancel(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd
        
'    '------------------------------------------------------------------------------------------------------------------
'    Case "生成费用收费单", "生成费用记帐单", "生成费用零费单"
'
''        Dim blnZero As Boolean
''        Dim intTmp As Integer
'
'        With vsf(mintIndex)
'            gstrSQL = GetPublicSQL(SQL.手术费用选择)
'
'            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(.RowData(.Row)))
'
'            If ShowPubSelect(Me, Nothing, 3, "编码,900,0,;名称,2400,0,;规格,900,0,;数量,1200,2,;单位,810,0,", Me.Name & "\手术费用选择", "请从下面左边列表中选择手术费用参考", rsData, rs, 8790, 4500, False, , , True) = 1 Then
'
'                blnZero = False
'                Select Case strCommand
'                Case "生成费用收费单"
'                    intTmp = 1
'                Case "生成费用记帐单"
'                    intTmp = 2
'                Case "生成费用零费单"
'                    intTmp = 2
'                    blnZero = True
'                End Select
'
'                If frmOpsStationCharge.ShowEdit(Me, Val(rs("ID").Value), Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), mstrPrivs, intTmp, blnZero, IIf(.TextMatrix(.Row, .ColIndex("病人来源")) = "住院", 2, 1)) Then
'
'                    '刷新费用列表
'                    If Not (mclsExpenses Is Nothing) Then
'                        Call mclsExpenses.zlRefresh(mlng手术部门id, Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), Val(.TextMatrix(.Row, .ColIndex("发送号"))), False)
'                    End If
'
'                End If
'            End If
'        End With
'
'        GoTo errEnd
        
    '------------------------------------------------------------------------------------------------------------------
    Case "定位数据"
        
        Dim lngRow As Long
        Dim intCol As Integer

        lngRow = -1

        With vsf(mintIndex)

            intCol = .ColIndex(mstrFindKey)

            lngRow = mclsVsf(mintIndex).FindRow(UCase(varParam(0)), intCol, 2, .Row + 1)
            If lngRow = -1 Then
                lngRow = mclsVsf(mintIndex).FindRow(UCase(varParam(0)), intCol, 2)
            End If
            If lngRow > 0 And .Row <> lngRow Then
                .Row = lngRow
                .ShowCell .Row, .Col
            End If

        End With

'        Call LocationObj(txtLocation)
    
    '------------------------------------------------------------------------------------------------------------------
    Case "恢复数据"

        '1.恢复概要记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationOutLine.DataChanged Then
            If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 And vsf(mintIndex).Rows > 2 Then
                vsf(mintIndex).Rows = vsf(mintIndex).Rows - 1
                vsf(mintIndex).Row = vsf(mintIndex).Rows - 1
            End If

            Call ExecuteCommand("读取手术概要")
            mfrmChildStationOutLine.DataChanged = False
        End If

        '2.恢复用药记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationDrug.DataChanged Then
            Call ExecuteCommand("读取手术用药")
            mfrmChildStationDrug.DataChanged = False
        End If

        '3.恢复材料记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationMaterial.DataChanged Then
            Call ExecuteCommand("读取手术材料")
            mfrmChildStationMaterial.DataChanged = False
        End If

        '4.恢复治疗记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationCure.DataChanged Then
            Call ExecuteCommand("读取手术治疗")
            mfrmChildStationCure.DataChanged = False
        End If
        
'        mblnNew = False
    '------------------------------------------------------------------------------------------------------------------
    Case "校验数据"

        '1.校验概要记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationOutLine.DataChanged Then
            If mfrmChildStationOutLine.ValidData = False Then GoTo errEnd
        End If

        '2.校验用药记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationDrug.DataChanged Then
            If mfrmChildStationDrug.ValidData = False Then GoTo errEnd
        End If

        '3.校验材料记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationMaterial.DataChanged Then
            If mfrmChildStationMaterial.ValidData = False Then GoTo errEnd
        End If

        '4.校验治疗记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationCure.DataChanged Then
            If mfrmChildStationCure.ValidData = False Then GoTo errEnd
        End If
        
        ExecuteCommand = True

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "保存数据"

        mlngTmp = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))

        '1.保存概要记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationOutLine.DataChanged Then

            If mfrmChildStationOutLine.SaveData(rsSQL) = False Then GoTo errEnd

        End If

        '2.保存用药记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationDrug.DataChanged Then
            If mfrmChildStationDrug.SaveData(rsSQL) = False Then GoTo errEnd
        End If

        '3.保存材料记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationMaterial.DataChanged Then
            If mfrmChildStationMaterial.SaveData(rsSQL) = False Then GoTo errEnd
        End If

        '4.保存治疗记录
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationCure.DataChanged Then
            If mfrmChildStationCure.SaveData(rsSQL) = False Then GoTo errEnd
        End If
        
        ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)

        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "前一条"
        With vsf(mintIndex)
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
                Call vsf_AfterRowColChange(mintIndex, 0, 0, .Row, .Col)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "后一条"
        With vsf(mintIndex)
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
                Call vsf_AfterRowColChange(mintIndex, 0, 0, .Row, .Col)
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "读注册表"

        If Val(GetRegister(私有全局, "", "使用个性化风格", "0")) = 1 Then
            '使用个性化设置

            dkpMain.LoadStateFromString GetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, "")

            mstrFindKey = Trim(GetRegister(私有模块, Me.Name, "定位依据", "姓名"))
            mclsVsf(0).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(0)) & "A", ""))
            mclsVsf(1).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(1)) & "B", ""))
            mclsVsf(2).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(2)) & "C", ""))
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "写注册表"

        If Val(GetRegister(私有全局, "", "使用个性化风格", "0")) = 1 Then
            '使用个性化设置
            Call SetRegister(私有模块, Me.Name, "定位依据", mstrFindKey)
        End If
        Call SetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        Call SetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(0)) & "A", mclsVsf(0).SaveStateToString)
        Call SetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(1)) & "B", mclsVsf(1).SaveStateToString)
        Call SetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(2)) & "C", mclsVsf(2).SaveStateToString)

    End Select

    ExecuteCommand = True
    
    AutoRefresh = True
    
    Exit Function

    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
errEnd:
    AutoRefresh = True
    
End Function

'控件或窗体事件
'######################################################################################################################

'Private Sub cbsDept_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'
'    Select Case Control.ID
'    Case conMenu_View_Dept
'
'        If mlng手术部门id <> mobjPeisDept.ItemData(mobjPeisDept.ListIndex) Then
'            mlng手术部门id = mobjPeisDept.ItemData(mobjPeisDept.ListIndex)
'            mstr手术部门 = zlCommFun.GetNeedName(mobjPeisDept.List(mobjPeisDept.ListIndex))
'            Call ExecuteCommand("刷新数据")
'        End If
'
'    End Select
'
'End Sub

'Private Sub cbsDept_Resize()
'    Dim lngLeft As Long
'    Dim lngTop  As Long
'    Dim lngRight  As Long
'    Dim lngBottom  As Long
'
'    Call cbsDept.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'
'    On Error Resume Next
'
'    '窗体其它控件Resize处理
'    tbc.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
'
'End Sub

Private Sub cboDept_Click()

    If mlng手术部门id <> cboDept.ItemData(cboDept.ListIndex) Then
        mlng手术部门id = cboDept.ItemData(cboDept.ListIndex)
        mstr手术部门 = zlCommFun.GetNeedName(cboDept.List(cboDept.ListIndex))
        Call ExecuteCommand("刷新数据")
    End If

End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strTmp As String
    Dim lngLoop As Long
    Dim objControl As Object

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BillPrintView
        
        With vsf(mintIndex)
            If mintIndex <> 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Then Exit Sub
            ReportOpen gcnOracle, ParamInfo.系统号, "ZL1_BILL_1804", Me, "医嘱id=" & Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), 1
        End With
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BillPrint
        
        With vsf(mintIndex)
            If mintIndex <> 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) = 0 Then Exit Sub
            ReportOpen gcnOracle, ParamInfo.系统号, "ZL1_BILL_1804", Me, "医嘱id=" & Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), 2
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Parameter                                                         '参数设置

        If ExecuteCommand("本地参数设置") Then
            AutoRefresh = True
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option

        Call ExecuteCommand("手术间设置")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Request

        If ExecuteCommand("直接手术申请") Then
            Call ExecuteCommand("装载等待手术", "刷新其他信息")
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Audit
        If ExecuteCommand("审核手术申请") Then
            Call ExecuteCommand("装载等待手术")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_UnAudit
        If ExecuteCommand("取消手术审核") Then
            Call ExecuteCommand("装载等待手术")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Arrange
        If ExecuteCommand("安排手术申请") Then
            Call ExecuteCommand("装载等待手术")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_UnArrange
        If ExecuteCommand("取消手术安排") Then
            Call ExecuteCommand("装载等待手术")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Plan
        If ExecuteCommand("手术报到") Then
            Call ExecuteCommand("装载等待手术")
            Call ExecuteCommand("装载正在手术")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Logout
        If ExecuteCommand("取消报到") Then
            Call ExecuteCommand("装载等待手术")
            Call ExecuteCommand("装载正在手术")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Complete
        If ExecuteCommand("手术完成") Then
            Call ExecuteCommand("装载正在手术")
            Call ExecuteCommand("装载已完手术")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Undone
        If ExecuteCommand("取消完成") Then
            Call ExecuteCommand("装载正在手术")
            Call ExecuteCommand("装载已完手术")
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save                  '保存手术方案

        If ExecuteCommand("校验数据") And DataChanged Then
            If ExecuteCommand("保存数据") Then
                DataChanged = False
                Call ExecuteCommand("刷新指定数据")
'                mblnNew = False
            End If
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle                  '恢复手术方案

        Call ExecuteCommand("恢复数据")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        If frmOpsStationFilter.ShowSearch(Me, mstrCondition, mlng手术部门id) Then
            Call ExecuteCommand("刷新数据")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh

        Call ExecuteCommand("刷新数据")
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump

        If tbcPage.Selected.Index + 1 <= tbcPage.ItemCount - 1 Then
            tbcPage.Item(tbcPage.Selected.Index + 1).Selected = True
        Else
            tbcPage.Item(0).Selected = True
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Exit, conMenu_Help_About, conMenu_Help_Web_Mail, conMenu_Help_Web_Home, conMenu_View_StatusBar, _
        conMenu_View_ToolBar_Size, conMenu_View_ToolBar_Text, conMenu_View_ToolBar_Button, conMenu_File_PrintSet, _
        conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel, conMenu_Help_Web_Forum, conMenu_Help_Help

        Call CommandBarExecutePublic(Control, Me, vsf(mintIndex), "病人手术清单")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward

        Call ExecuteCommand("前一条")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward

        Call ExecuteCommand("后一条")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem

        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
        
        Call ExecuteCommand("定位数据", txtLocation.Text)
        
    '--------------------------------------------------------------------------------------------------------------
    Case Else

        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            'Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
        Else
            Select Case tbcPage.Selected.Tag
            Case "附费"
                
                Select Case Control.ID
                '------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2# + 1
                
                    Call ExecuteCommand("生成费用收费单")
            
                '------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2# + 2
            
                    Call ExecuteCommand("生成费用记帐单")
                    
                '------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2# + 3
                            
                    Call ExecuteCommand("生成费用零费单")
                Case Else
                    Call mclsExpenses.zlExecuteCommandBars(Control)
                End Select
                
            Case "住院医嘱"
                Call mclsInAdvices.zlExecuteCommandBars(Control)
            Case "门诊医嘱"
                Call mclsOutAdvices.zlExecuteCommandBars(Control)
            Case "住院病历"
                Call mfrmChildStationInEPR.zlExecuteCommandBars(Control)
            Case "门诊病历"
                Call mfrmChildStationOutEPR.zlExecuteCommandBars(Control)
            End Select
            
        End If

    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As Object
    
    If CommandBar.Parent Is Nothing Then Exit Sub

    Select Case tbcPage.Selected.Tag
    Case "附费"
        Call mclsExpenses.zlPopupCommandBars(CommandBar)
        
        If CommandBar.Parent.ID = conMenu_Edit_Preferences Then
    
            With CommandBar.Controls
    
                .DeleteAll
    
                Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 1, "生成收费单据(&1)")
                Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 2, "生成记帐单据(&2)")
                Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 3, "生成零耗费用(&3)")
                With cbsMain.KeyBindings
                    .Add FCONTROL, vbKeyN, conMenu_Edit_MakeCharge * 2 + 1
                    .Add FCONTROL, vbKeyB, conMenu_Edit_MakeCharge * 2 + 2
                End With
                
            End With
        End If
        
    Case "门诊医嘱"
        Call mclsOutAdvices.zlPopupCommandBars(CommandBar)
    Case "住院医嘱"
        Call mclsInAdvices.zlPopupCommandBars(CommandBar)
    Case "门诊病历"
        Call mfrmChildStationOutEPR.zlPopupCommandBars(CommandBar)
    Case "住院病历"
        Call mfrmChildStationInEPR.zlPopupCommandBars(CommandBar)
    End Select

End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo errHand

    With vsf(mintIndex)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel

            Control.Enabled = (Val(.RowData(.Row)) > 0)
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_BillPrintView, conMenu_File_BillPrint
            
            With vsf(mintIndex)
                
                Control.Visible = IsPrivs(mstrPrivs, "手术通知单")
                
                Control.Enabled = (Val(.TextMatrix(.Row, .ColIndex("医嘱id"))) > 0 And mintIndex = 0 And Control.Visible And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) >= 2)
                
            End With
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Option
            Control.Visible = IsPrivs(mstrPrivs, "执行间设置")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_ManagePopup

            Control.Visible = IsPrivs(mstrPrivs, "审核手术") Or _
                                IsPrivs(mstrPrivs, "审核取消") Or _
                                IsPrivs(mstrPrivs, "安排手术") Or _
                                IsPrivs(mstrPrivs, "安排取消") Or _
                                IsPrivs(mstrPrivs, "执行报到") Or _
                                IsPrivs(mstrPrivs, "报到取消") Or _
                                IsPrivs(mstrPrivs, "完成手术") Or _
                                IsPrivs(mstrPrivs, "补录申请") Or _
                                IsPrivs(mstrPrivs, "完成取消")

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Request            '补录申请
        
            Control.Visible = IsPrivs(mstrPrivs, "补录申请")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Audit            '审核

            Control.Visible = IsPrivs(mstrPrivs, "审核手术")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 0 And Val(.RowData(.Row)) <> 0 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_UnAudit            '审核取消

            Control.Visible = IsPrivs(mstrPrivs, "审核取消")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 1 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Arrange               '安排手术
'
            Control.Visible = IsPrivs(mstrPrivs, "安排手术")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 1 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_UnArrange             '安排取消

            Control.Visible = IsPrivs(mstrPrivs, "安排取消")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 2 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Plan                '报到

            Control.Visible = IsPrivs(mstrPrivs, "执行报到")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 2 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Complete            '手术完成

            Control.Visible = IsPrivs(mstrPrivs, "完成手术")
            Control.Enabled = (mintIndex = 1 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Logout              '报到取消

            Control.Visible = IsPrivs(mstrPrivs, "报到取消")
            Control.Enabled = (mintIndex = 1 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Undone                            '完成取消

            Control.Visible = IsPrivs(mstrPrivs, "完成取消")
            Control.Enabled = (mintIndex = 2 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 4 And Control.Visible)

        '------------------------------------------------------------------------------------------------------------------
        Case conMenu_EditPopup

            Control.Visible = IsPrivs(mstrPrivs, "用药准备") Or _
                                IsPrivs(mstrPrivs, "材料准备") Or _
                                IsPrivs(mstrPrivs, "材料登记") Or _
                                IsPrivs(mstrPrivs, "概要登记") Or _
                                IsPrivs(mstrPrivs, "用药登记")
            
        '------------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle

            Control.Visible = IsPrivs(mstrPrivs, "用药准备") Or _
                                IsPrivs(mstrPrivs, "材料准备") Or _
                                IsPrivs(mstrPrivs, "材料登记") Or _
                                IsPrivs(mstrPrivs, "概要登记") Or _
                                IsPrivs(mstrPrivs, "用药登记")
            Control.Enabled = (DataChanged And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Filter, conMenu_View_Refresh, conMenu_Edit_Request
            Control.Enabled = (DataChanged = False And Control.Visible)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button            '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = cbsMain(2).Visible
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Text              '图标文字
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Size              '大图标
            Control.Checked = cbsMain.Options.LargeIcons
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_StatusBar                 '状态栏
            Control.Checked = stbThis.Visible
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Forward
            Control.Enabled = (.Row > 1 And DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Backward
            Control.Enabled = (.Row < .Rows - 1 And DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem        '
            Control.Checked = (mstrFindKey = Control.Parameter)
            Control.Enabled = (DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Location
        
             Control.Enabled = (DataChanged = False)
         
        '--------------------------------------------------------------------------------------------------------------
        Case Else
            Select Case tbcPage.Selected.Tag
            Case "附费"
                
                Select Case Control.ID
                '--------------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_Preferences
                    
                    Control.Visible = IsPrivs(mstrPrivs, "补录附费")
                    Control.Enabled = Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3 And Control.Visible
                    
                '--------------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2 + 1
                    
                    Control.Visible = (.TextMatrix(.Row, .ColIndex("病人来源")) = "门诊" And IsPrivs(mstrPrivs, "补录附费"))
                    Control.Enabled = Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3 And Control.Visible
                    
                '--------------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2 + 2, conMenu_Edit_MakeCharge * 2 + 3
                    
                    Control.Visible = IsPrivs(mstrPrivs, "补录附费")
                    Control.Enabled = Val(.TextMatrix(.Row, .ColIndex("手术状态"))) = 3 And Control.Visible
                '--------------------------------------------------------------------------------------------------------------
                Case Else
                    Call mclsExpenses.zlUpdateCommandBars(Control)
                End Select

            Case "住院医嘱"
                Call mclsInAdvices.zlUpdateCommandBars(Control)
            Case "门诊医嘱"
                Call mclsOutAdvices.zlUpdateCommandBars(Control)
            Case "住院病历"
                Call mfrmChildStationInEPR.zlUpdateCommandBars(Control)
            Case "门诊病历"
                Call mfrmChildStationOutEPR.zlUpdateCommandBars(Control)
            End Select
        End Select
    End With

errHand:

End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
'    Case 2
'        Item.Handle = picPane(6).hWnd
    Case 2
        Item.Handle = picPane(8).hWnd
    Case 3
        Item.Handle = picPane(7).hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    DoEvents

    If ExecuteCommand("初始数据") = False Then GoTo errHand

    Call ExecuteCommand("刷新数据")
    
    AutoRefresh = True
    
    mblnAllowClose = True
    Exit Sub

    '------------------------------------------------------------------------------------------------------------------
errHand:
    mblnAllowClose = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mblnAllowClose = False

    mstrPrivs = UserInfo.模块权限
'    mstrPrivs = "所有手术室"

    Call ExecuteCommand("初始控件")
    Call ExecuteCommand("读注册表")

    Call RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, ParamInfo.系统号, ParamInfo.模块号, UserInfo.模块权限)

End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Call SetPaneRange(dkpMain, 1, 100, 100, 200, Me.ScaleHeight)
'    Call SetPaneRange(dkpMain, 2, 15, 26, Me.ScaleWidth, 26)
    Call SetPaneRange(dkpMain, 2, 100, 70, Me.ScaleWidth, 70)
'
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    Call ExecuteCommand("写注册表")

    Call SaveWinState(Me, App.ProductName)

    Set mclsVsf(0) = Nothing
    Set mclsVsf(1) = Nothing
    Set mclsVsf(2) = Nothing


    Set mfrmChildStationOutLine = Nothing
    Set mfrmChildStationDrug = Nothing
    Set mfrmChildStationMaterial = Nothing
    Set mfrmChildStationInEPR = Nothing
    Set mfrmChildStationOutEPR = Nothing
    Set mfrmChildStationCure = Nothing
    
    Set mobjFindKey = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsInAdvices = Nothing
    Set mclsExpenses = Nothing

End Sub

Private Sub mfrmChildStationCure_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildStationDrug_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildStationMaterial_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildStationOutLine_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
'        mobjPeisDept.Width = picPane(Index).Width / 15 - 55
        cboDept.Move cboDept.Left, cboDept.Top, picPane(Index).Width - cboDept.Left - 30
        tbc.Move 0, cboDept.Top + cboDept.Height + 30, picPane(Index).Width, picPane(Index).Height - (cboDept.Top + cboDept.Height + 30)
    Case 1
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(0).AppendRows = True
    Case 2
        vsf(1).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(1).AppendRows = True
    Case 3
        vsf(2).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(2).AppendRows = True
    Case 7
        tbcPage.Move 15, 0, picPane(Index).Width - 30, picPane(Index).Height
    Case 8
        fra.Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        pic.Move 30, 120, fra.Width - 60, fra.Height - 150
        picState.Move pic.Width - picState.Width - 90, picState.Top
    End Select
End Sub

Private Sub tbc_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mintIndex = Item.Index

    If mblnStartUp Then Exit Sub

    Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值

    Call SubWinDefCommandBar(Item)

End Sub

Private Sub tmr_Timer()
    Dim strSvrKey As String
    
    mlngCountTmr = mlngCountTmr + 1
    
    If mlngCountTmr >= Val(tmr.Tag) Then
    
        '时间到了，开始触发
        mlngCountTmr = 0
                
        Call ExecuteCommand("刷新数据")
                
    End If
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim bytMatch As Byte

    If KeyAscii = vbKeyReturn Then

        lngRow = -1
        bytMatch = 2

        With vsf(mintIndex)

            intCol = .ColIndex(mstrFindKey)

            lngRow = mclsVsf(mintIndex).FindRow(UCase(txtLocation.Text), intCol, bytMatch, .Row + 1)
            If lngRow = -1 Then
                lngRow = mclsVsf(mintIndex).FindRow(UCase(txtLocation.Text), intCol, bytMatch)
            End If
            If lngRow > 0 And .Row <> lngRow Then
                .Row = lngRow
                .ShowCell .Row, .Col
            End If

        End With

        Call LocationObj(txtLocation)
    End If
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Call mclsVsf(Index).AfterMoveColumn(Col, Position)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnChange As Boolean
    
    On Error Resume Next
    
    With vsf(mintIndex)
        blnChange = ExchangeAdvice(.TextMatrix(NewRow, .ColIndex("病人来源")) = "门诊")
        blnChange = blnChange Or ExchangeEPRS(.TextMatrix(NewRow, .ColIndex("病人来源")) = "门诊")
    End With

    Call ExecuteCommand("读取申请信息")
    Call ExecuteCommand("刷新手术数据")

End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey)
    vsf(Index).ShowCell vsf(Index).Row, vsf(Index).Col
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mclsVsf(Index).SaveKey = Val(vsf(Index).RowData(vsf(Index).Row))
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col < 3)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar

    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        cbrPopupBar.ShowPopup
    End Select
End Sub
