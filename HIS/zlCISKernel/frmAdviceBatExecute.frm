VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceBatExecute 
   Caption         =   "医嘱批量执行登记"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13260
   Icon            =   "frmAdviceBatExecute.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   13260
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2115
      ScaleHeight     =   180
      ScaleWidth      =   285
      TabIndex        =   49
      Top             =   315
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   7110
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceBatExecute.frx":6852
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceBatExecute.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceBatExecute.frx":7386
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceBatExecute.frx":76D8
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceBatExecute.frx":DF3A
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceBatExecute.frx":1479C
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceBatExecute.frx":14D36
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picExecuted 
      BorderStyle     =   0  'None
      Height          =   2160
      Left            =   3210
      ScaleHeight     =   2160
      ScaleWidth      =   13095
      TabIndex        =   22
      Top             =   3300
      Width           =   13095
      Begin VSFlex8Ctl.VSFlexGrid vsgExecAdvice 
         Height          =   1185
         Left            =   0
         TabIndex        =   26
         Top             =   480
         Width           =   8505
         _cx             =   15002
         _cy             =   2090
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAdviceBatExecute.frx":152D0
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
         ExplorerBar     =   7
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
   Begin VB.PictureBox picPati 
      BorderStyle     =   0  'None
      Height          =   8985
      Left            =   135
      ScaleHeight     =   8985
      ScaleWidth      =   2985
      TabIndex        =   2
      Top             =   900
      Width           =   2985
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   3150
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2655
         _Version        =   589884
         _ExtentX        =   4683
         _ExtentY        =   5556
         _StockProps     =   0
         BorderStyle     =   2
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picFitter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3300
         Left            =   150
         ScaleHeight     =   3300
         ScaleWidth      =   2655
         TabIndex        =   4
         Top             =   3975
         Width           =   2655
         Begin VB.ComboBox cboExecutePoeple 
            Height          =   300
            Index           =   1
            Left            =   720
            TabIndex        =   27
            Top             =   1050
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker dpkReqTime 
            Height          =   300
            Index           =   0
            Left            =   720
            TabIndex        =   8
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   237109251
            CurrentDate     =   40945
         End
         Begin MSComCtl2.DTPicker dpkReqTime 
            Height          =   300
            Index           =   1
            Left            =   720
            TabIndex        =   10
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   237109251
            CurrentDate     =   40945.9999884259
         End
         Begin VB.PictureBox picPanel 
            BorderStyle     =   0  'None
            Height          =   1800
            Left            =   105
            ScaleHeight     =   1800
            ScaleWidth      =   2490
            TabIndex        =   48
            Top             =   1440
            Width           =   2490
            Begin VB.CheckBox chkType 
               Caption         =   "中药"
               Height          =   255
               Index           =   9
               Left            =   1800
               TabIndex        =   51
               Top             =   1185
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "输血"
               Height          =   255
               Index           =   7
               Left            =   960
               TabIndex        =   28
               Top             =   1185
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "其它医嘱"
               Height          =   255
               Index           =   6
               Left            =   105
               TabIndex        =   21
               Top             =   1470
               Width           =   1035
            End
            Begin VB.CheckBox chkType 
               Caption         =   "采集"
               Height          =   255
               Index           =   5
               Left            =   1800
               TabIndex        =   20
               Top             =   900
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "治疗"
               Height          =   255
               Index           =   4
               Left            =   960
               TabIndex        =   19
               Top             =   900
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "皮试"
               Height          =   255
               Index           =   3
               Left            =   105
               TabIndex        =   18
               Top             =   900
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "口服"
               Height          =   255
               Index           =   2
               Left            =   1800
               TabIndex        =   17
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "注射"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   16
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "输液"
               Height          =   255
               Index           =   0
               Left            =   105
               TabIndex        =   15
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chk期效 
               Caption         =   "临嘱"
               Height          =   180
               Index           =   1
               Left            =   1800
               TabIndex        =   13
               Top             =   0
               Width           =   735
            End
            Begin VB.CheckBox chk期效 
               Caption         =   "长嘱"
               Height          =   180
               Index           =   0
               Left            =   960
               TabIndex        =   12
               Top             =   0
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "其它给药途径"
               Height          =   255
               Index           =   8
               Left            =   105
               TabIndex        =   1
               Top             =   1185
               Width           =   1395
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "类别"
               Height          =   180
               Index           =   5
               Left            =   105
               TabIndex        =   14
               Top             =   330
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "期效"
               Height          =   180
               Index           =   4
               Left            =   105
               TabIndex        =   11
               Top             =   0
               Width           =   360
            End
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "执行人"
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   50
            Top             =   1095
            Width           =   540
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "到"
            Height          =   180
            Index           =   3
            Left            =   450
            TabIndex        =   9
            Top             =   750
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "从"
            Height          =   180
            Index           =   2
            Left            =   450
            TabIndex        =   7
            Top             =   390
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "要求时间"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   90
            Width           =   720
         End
      End
      Begin VB.Frame fraBaby 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   29
         Top             =   3615
         Visible         =   0   'False
         Width           =   2600
         Begin VB.OptionButton optBaby 
            Caption         =   "病人"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   32
            Top             =   0
            Width           =   660
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "所有医嘱"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "婴儿"
            Height          =   180
            Index           =   2
            Left            =   1815
            TabIndex        =   30
            Top             =   0
            Width           =   660
         End
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "当前病区："
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox picWaitExecute 
      BorderStyle     =   0  'None
      Height          =   2160
      Left            =   3330
      ScaleHeight     =   2160
      ScaleWidth      =   11535
      TabIndex        =   24
      Top             =   825
      Width           =   11535
      Begin VB.PictureBox picExec 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   0
         ScaleHeight     =   720
         ScaleWidth      =   10935
         TabIndex        =   33
         Top             =   0
         Width           =   10935
         Begin VB.ComboBox cboExecuteResult 
            Height          =   300
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   125
            Width           =   975
         End
         Begin VB.Frame fraExecutePeople 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   300
            Left            =   1080
            TabIndex        =   39
            Top             =   420
            Width           =   5415
            Begin VB.OptionButton optExecutePeople 
               Caption         =   "上次执行人"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   43
               Top             =   60
               Width           =   1215
            End
            Begin VB.OptionButton optExecutePeople 
               Caption         =   "指定人员"
               Height          =   180
               Index           =   1
               Left            =   1320
               TabIndex        =   42
               Top             =   60
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.ComboBox cboExecutePoeple 
               Height          =   300
               Index           =   0
               Left            =   2450
               TabIndex        =   41
               Top             =   0
               Width           =   1815
            End
            Begin VB.OptionButton optExecutePeople 
               Caption         =   "本人"
               Height          =   180
               Index           =   2
               Left            =   4560
               TabIndex        =   40
               Top             =   60
               Width           =   855
            End
         End
         Begin VB.Frame frmExecuteTime 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   310
            Left            =   1080
            TabIndex        =   35
            Top             =   105
            Width           =   4335
            Begin VB.OptionButton optExecuteTime 
               Caption         =   "要求时间"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   37
               Top             =   60
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optExecuteTime 
               Caption         =   "指定时间"
               Height          =   180
               Index           =   1
               Left            =   1320
               TabIndex        =   36
               Top             =   60
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker dpkExecuteTime 
               Height          =   300
               Left            =   2450
               TabIndex        =   38
               Top             =   0
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   237109251
               CurrentDate     =   40945
            End
         End
         Begin VB.CommandButton cmdBatUpdate 
            Caption         =   "批量修改(&M)"
            Height          =   300
            Left            =   7800
            TabIndex        =   34
            Top             =   100
            Width           =   1215
         End
         Begin VB.Label lblInfo 
            Caption         =   "执行时间"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   47
            Top             =   170
            Width           =   855
         End
         Begin VB.Label lblInfo 
            Caption         =   "执行结果"
            Height          =   255
            Index           =   7
            Left            =   5640
            TabIndex        =   46
            Top             =   170
            Width           =   855
         End
         Begin VB.Label lblInfo 
            Caption         =   "执行人"
            Height          =   255
            Index           =   8
            Left            =   315
            TabIndex        =   45
            Top             =   480
            Width           =   615
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsgWaitExecAdvice 
         Height          =   1035
         Left            =   0
         TabIndex        =   25
         Top             =   885
         Width           =   8505
         _cx             =   15002
         _cy             =   1826
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAdviceBatExecute.frx":1536B
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
         ExplorerBar     =   7
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   795
      Left            =   4470
      TabIndex        =   0
      Top             =   -180
      Width           =   2235
      _Version        =   589884
      _ExtentX        =   3942
      _ExtentY        =   1402
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   9870
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   635
      SimpleText      =   $"frmAdviceBatExecute.frx":15406
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceBatExecute.frx":1544D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18309
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   465
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAdviceBatExecute.frx":15CE1
      Left            =   1410
      Top             =   285
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmAdviceBatExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PatiCol
    COL_病人ID = 0
    COL_主页ID = 1
    COL_选择 = 2
    COL_床号 = 3
    COL_姓名 = 4
    COL_性别 = 5
    COL_住院号 = 6
End Enum

Private Enum AdviceCol
    col选择 = 0
    COL床位 = 1
    col姓名 = 2
    col性别 = 3
    col住院号 = 4
    col住院次数 = 5
    col期效 = 6
    col医嘱内容 = 7
    col总量 = 8
    col单量 = 9
    col本次数次 = 10
    COL给药途径 = 11
    col要求时间 = 12
    col执行时间 = 13
    col执行人 = 14
    col核对时间 = 15
    COL核对人 = 16
    col执行结果 = 17
    COL未执行原因 = 18
    col备注 = 19
    col医嘱ID = 20
    col相关ID = 21
    col诊疗类别 = 22
    Col病人ID = 23
    COL主页ID = 24
    COL频率 = 25
    col发送号 = 26
    col待执行ID = 27
    COL医嘱状态 = 28
    col在院 = 29
    col登记人 = 30
    COL原始执行时间 = 31
    col是否修改 = 32
    COL最后执行人 = 33
    col操作类型 = 34
    col执行分类 = 35
    col病区ID = 36
    col执行状态 = 37
    col开始执行时间 = 38
    col检查时间 = 39
End Enum

Private Enum ClientType
    chk长嘱 = 0
    chk临嘱 = 1
    
    Type输液 = 0
    Type注射 = 1
    Type口服 = 2
    Type皮试 = 3
    Type治疗 = 4
    Type采集 = 5
    Type其它医嘱 = 6
    Type输血 = 7
    Type其它给药途径 = 8
    Type中药服法 = 9
    lbl病区 = 0
    lbl时间范围 = 1
    lbl期效 = 4
    lbl类别 = 5
    lbl执行人 = 11
    
    dpk开始日期 = 0
    dpk结束日期 = 1
    
    opt要求时间 = 0
    opt指定时间 = 1
    
    opt上次执行人 = 0
    opt指定人员 = 1
    opt本人 = 2
    
    cbo执行人登记 = 0
    cbo执行人取消 = 1
    
End Enum

Private Type FilterCond
    str病人IDs As String '病人选择情况，"病人ID:主页ID,..." 或者 "*" 表示全选/全不选
    datB As Date '开始日期
    datE As Date '终止日期
    str人员 As String     '人员姓名
End Type
Private mvarCond As FilterCond

Private mlng病区ID As Long
Private mlng病人ID As Long
Private mrsDefine As ADODB.Recordset
Private mblnWaitIsUpdate As Boolean
Private mblnExecIsUpdate As Boolean
Private mint场合 As Integer '0-医生站调用,1-护士站调用
Private mbytUseType As Integer        '调用模式   1-医嘱核对，0-医嘱执行
Private mint医嘱处理范围 As Integer    '医嘱处理范围   0-所有医嘱,1-病人医嘱,2-婴儿医嘱
Private mlng医护科室ID As Long
Private mlng婴儿科室ID As Long
Private mlng婴儿病区ID As Long

Private Enum UseType
    T医嘱执行 = 0
    T医嘱核对 = 1
End Enum

Public Sub ShowMe(ByVal intType As Integer, ByRef frmParent As Object, ByVal lng病区ID As Long, ByVal lng病人ID As Long, ByVal int场合 As Integer, ByVal bytUseType As Byte, Optional ByVal lng医护科室ID As Long, _
    Optional ByVal lng婴儿科室ID As Long, Optional ByVal lng婴儿病区ID As Long)
    mlng病区ID = lng病区ID
    mlng病人ID = lng病人ID
    mint场合 = int场合
    mbytUseType = bytUseType
    mlng医护科室ID = lng医护科室ID
    mlng婴儿科室ID = lng婴儿科室ID
    mlng婴儿病区ID = lng婴儿病区ID
    Me.lblInfo(lbl病区).Caption = IIF(mint场合 = 1, "当前病区：", IIF(Val(zlDatabase.GetPara("部门显示方式", glngSys, p住院医生站)) = 1, "当前病区：", "当前科室：")) & Sys.RowValue("部门表", IIF(mlng婴儿病区ID <> 0 And (mlng婴儿科室ID = mlng医护科室ID Or mlng婴儿病区ID = mlng医护科室ID), lng婴儿病区ID, lng病区ID), "名称")
    
    Me.Show intType, frmParent
End Sub

Private Sub cboExecutePoeple_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim i As Long
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    
    Call Cbo.SetIndex(cboExecutePoeple(Index).Hwnd, Cbo.MatchIndex(cboExecutePoeple(Index).Hwnd, KeyAscii))
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    tbcSub.SetFocus
    Select Case Control.ID
    
    Case conMenu_View_Refresh
        If tbcSub.Selected.Tag = "待执行医嘱" Then
            Call LoadAdvice
        Else
            Call LoadAdvice(True)
        End If
    Case conMenu_Edit_Save
        If tbcSub.Selected.Tag = "待执行医嘱" Then
            Call FuncSaveExecute
        Else
            Call FuncSaveUpdate
        End If
    Case conMenu_Manage_ThingAudit
        Call FuncExecAuditBatch
    Case conMenu_Manage_ThingDelAudit
        Call FuncExecDelAuditBatch
    Case conMenu_Manage_Undone
        Call FuncCancleExec
    Case conMenu_File_Exit
        Unload Me
    
    End Select
End Sub

Private Function FuncExecAuditBatch() As Boolean
'功能：批量核对
    Dim strSQL As String
    Dim str核对人 As String
    Dim i As Long
    Dim arrSQL As Variant
    Dim rsTmp As Recordset
    Dim blnTrans As Boolean
    Dim strMsgNameSame As String
    Dim strMsg As String
    Dim strXML As String
    Dim blnDo As Boolean
    Dim str核对时间 As String
    
    On err GoTo errH
    arrSQL = Array()
    str核对时间 = zlDatabase.Currentdate '获取当前服务器时间
    With vsgWaitExecAdvice
        For i = 0 To .Rows - 1
            If .Cell(flexcpData, i, col选择) = "1" And .TextMatrix(i, col相关ID) = "" Then
                If str核对人 = "" Then str核对人 = zlDatabase.UserIdentifyByUser(Me, "在核对执行情况前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "执行情况登记", , True)
                If str核对人 = "" Then Exit Function

                If str核对人 = .TextMatrix(i, col执行人) & "" Then
                    strMsgNameSame = strMsgNameSame & "," & .TextMatrix(i, col医嘱内容)
                Else
                    '调用核对前外挂接口
                    If Not gobjPlugIn Is Nothing Then
                        On Error Resume Next
                        blnDo = gobjPlugIn.AdvcieBeforToReview(glngSys, IIF(mint场合 = 0, p住院医生站, p住院护士站), Val(.TextMatrix(i, Col病人ID)), Val(.TextMatrix(i, COL主页ID)), Val(.TextMatrix(i, col医嘱ID)), Val(.TextMatrix(i, col发送号)), str核对人, str核对时间, .TextMatrix(i, col执行人) & "", strXML)
                        Call zlPlugInErrH(err, "AdvcieBeforToReview")
                        If 0 = err.Number Then '接口没有出错的情况下再判断接口的返回值
                            If blnDo Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱核对_Insert(" & Val(.TextMatrix(i, col医嘱ID)) & "," & Val(.TextMatrix(i, col发送号)) & ",'" & str核对人 & "',Null,To_Date('" & str核对时间 & "','YYYY-MM-DD HH24:MI:SS'))"
                            End If
                        End If
                        If err.Number <> 0 Then err.Clear
                        On Error GoTo 0
                    Else
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱核对_Insert(" & Val(.TextMatrix(i, col医嘱ID)) & "," & Val(.TextMatrix(i, col发送号)) & ",'" & str核对人 & "',Null,To_Date('" & str核对时间 & "','YYYY-MM-DD HH24:MI:SS'))"
                    End If
                    
                End If
            End If
        Next
        
    
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        If strMsgNameSame <> "" Then
            strMsg = strMsg & "以下医嘱的审核人和执行人为同一个人：" & vbCrLf & Mid(strMsgNameSame, 2) & "。" & vbCrLf
        End If
        
        If UBound(arrSQL) < 0 Then
            MsgBox "您勾选的项目未核对成功，其中：" & vbCrLf & strMsg, vbInformation, "医嘱核对"
        Else
            If strMsg <> "" Then
                MsgBox "共核对了" & UBound(arrSQL) + 1 & "个项目，其中：" & vbCrLf & strMsg, vbInformation, "医嘱核对"
            End If
        End If
    End With
    '显示执行情况
    Call LoadAdvice
    FuncExecAuditBatch = True

    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncExecDelAuditBatch() As Boolean
'功能：批量取消核对
    Dim bln输血皮试 As Boolean
    Dim strSQL As String
    Dim str核对人 As String
    Dim i As Long
    Dim arrSQL As Variant
    Dim strMsg As String
    Dim rsTmp As Recordset
    Dim blnTrans As Boolean
    Dim blnIsTwo As Boolean   '判断是否存在有两个以上的核对人
    Dim strTmp As String
    Dim bln核对人 As Boolean
    Dim strMsgNoRecord As String
    Dim strNoExec As String
    Dim datCur As Date
    
    On err GoTo errH
    arrSQL = Array()
    datCur = zlDatabase.Currentdate
    With vsgExecAdvice
        For i = 0 To .Rows - 1
            If .Cell(flexcpData, i, col选择) = "1" And .TextMatrix(i, col相关ID) = "" Then
                If Val(.TextMatrix(i, col执行状态) & "") = 3 Then
                    If strTmp <> "" And strTmp <> .TextMatrix(i, COL核对人) & "" Then
                        blnIsTwo = True
                    Else
                        strTmp = .TextMatrix(i, COL核对人) & ""
                    End If
                    If CanUnExec(CDate(.TextMatrix(i, col核对时间)), datCur) Then
                        If .TextMatrix(i, COL核对人) & "" <> UserInfo.姓名 Then
                            If str核对人 = "" Then str核对人 = zlDatabase.UserIdentifyByUser(Me, "在取消核对前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "执行情况登记", , True)
                            If str核对人 = "" Then Exit Function
                            
                            If str核对人 = .TextMatrix(i, COL核对人) & "" Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱核对_Delete(" & Val(.TextMatrix(i, col医嘱ID)) & "," & Val(.TextMatrix(i, col发送号)) & ")"
                            Else
                                bln核对人 = True
                                str核对人 = .TextMatrix(i, COL核对人) & ""
                            End If
                        Else
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱核对_Delete(" & Val(.TextMatrix(i, col医嘱ID)) & "," & Val(.TextMatrix(i, col发送号)) & ")"
                        End If
                    Else
                        strNoExec = strNoExec & "," & .TextMatrix(i, col医嘱内容)
                    End If
                Else
                    strMsgNoRecord = strMsgNoRecord & "," & .TextMatrix(i, col医嘱内容)
                End If
            End If
        Next
    
        If blnIsTwo Then
            MsgBox "不能同时取消多个人核对的项目，请选择同一个人所核对的项目。", vbInformation, "取消核对"
            '显示执行情况
            Call LoadAdvice(True)
            Exit Function
        End If
        
        If bln核对人 Then
            MsgBox "只能取消自己核对的医嘱，当前选择的医嘱核对人是""" & str核对人 & """。", vbInformation, "取消核对"
            '显示执行情况
            Call LoadAdvice(True)
            Exit Function
        End If
    End With

    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "取消核对")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If strMsgNoRecord <> "" Then
        strMsg = strMsg & "以下医嘱已经执行完成，请取消完成后再试：" & vbCrLf & Mid(strMsgNoRecord, 2) & "。" & vbCrLf
    End If
    
    If strNoExec <> "" Then
        strMsg = strMsg & "以下医嘱已经执行完成，请取消完成后再试：" & vbCrLf & Mid(strNoExec, 2) & "。" & vbCrLf
    End If
    
    If UBound(arrSQL) < 0 Then
        MsgBox "您勾选的项目未取消成功，其中：" & vbCrLf & strMsg, vbInformation, "取消核对"
    Else
        If strMsg <> "" Then
            MsgBox "共取消核对了" & UBound(arrSQL) + 1 & "个项目,其中：" & vbCrLf & strMsg, vbInformation, "取消核对"
        End If
    End If
    '显示执行情况
    Call LoadAdvice(True)
    FuncExecDelAuditBatch = True

    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngLW As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
        
    'TabControl
    tbcSub.Left = lngLeft
    tbcSub.Top = lngTop
    tbcSub.Width = lngRight - lngLeft
    tbcSub.Height = Me.Height - stbThis.Height - 560 - lngTop
    
    picPati.Height = tbcSub.Height - 400
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    
    Case conMenu_Manage_Undone
        If tbcSub.Selected.Tag = "待执行医嘱" Or mbytUseType = T医嘱核对 Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Case conMenu_Edit_Save
        If mbytUseType = T医嘱核对 Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Case conMenu_Manage_ThingAudit
        If mbytUseType = T医嘱核对 And tbcSub.Selected.Tag = "待执行医嘱" Then
            Control.Visible = True
        Else
            Control.Visible = False
        End If
    Case conMenu_Manage_ThingDelAudit
        If mbytUseType = T医嘱核对 And tbcSub.Selected.Tag <> "待执行医嘱" Then
            Control.Visible = True
        Else
            Control.Visible = False
        End If
    End Select
End Sub

Private Sub chkType_Click(Index As Integer)
    Dim i As Long, blnIsSelect As Boolean
    
    For i = 0 To chkType.Count - 1
        If chkType(i).value = 1 Then blnIsSelect = True: Exit For
    Next
    If Not blnIsSelect Then chkType(Index).value = 1
End Sub

Private Sub chk期效_Click(Index As Integer)
    If chk期效(chk长嘱).value = 0 And chk期效(chk临嘱).value = 0 Then
        chk期效(Index).value = 1
    End If
End Sub

Private Sub cmdBatUpdate_Click()
    Call BatUpdate
    mblnWaitIsUpdate = True
End Sub

Private Sub BatUpdate()
'功能：批量修改
    Dim i As Long, lngRow As Long
    Dim blnSame As Boolean
    
    With vsgWaitExecAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, col选择) = "1" Then
                '执行时间
                If optExecuteTime(opt要求时间).value Then
                    .TextMatrix(i, col执行时间) = .TextMatrix(i, col要求时间)
                    .Cell(flexcpData, i, col执行时间) = .TextMatrix(i, col执行时间)
                ElseIf optExecuteTime(opt指定时间).value Then
                    '如果时间超出了开始执行时间，则取要求时间
                    If Format(dpkExecuteTime.value, "yyyy-MM-dd HH:mm") >= Format(.TextMatrix(i, col开始执行时间), "yyyy-MM-dd HH:mm") Then
                         '如果有同一条医嘱同次发送的，一起执行时，处理为要求时间
                        For lngRow = 1 To .Rows - 1
                            If .Cell(flexcpData, lngRow, col选择) = "1" Then
                                If .TextMatrix(lngRow, col要求时间) <> .TextMatrix(i, col要求时间) And .TextMatrix(lngRow, col医嘱ID) = .TextMatrix(i, col医嘱ID) And .TextMatrix(lngRow, col发送号) = .TextMatrix(i, col发送号) Then
                                    blnSame = True
                                    Exit For
                                End If
                            End If
                        Next
                        If blnSame Then
                            .TextMatrix(i, col执行时间) = .TextMatrix(i, col要求时间)
                        Else
                            .TextMatrix(i, col执行时间) = Format(dpkExecuteTime.value, "yyyy-MM-dd HH:mm")
                        End If
                        blnSame = False
                    Else
                       .TextMatrix(i, col执行时间) = .TextMatrix(i, col要求时间)
                    End If
                    .Cell(flexcpData, i, col执行时间) = .TextMatrix(i, col执行时间)
                End If
                '执行人
                If optExecutePeople(opt上次执行人).value Then
                    .TextMatrix(i, col执行人) = .TextMatrix(i, COL最后执行人)
                    .Cell(flexcpData, i, col执行人) = .TextMatrix(i, col执行人)
                ElseIf optExecutePeople(opt指定人员).value Then
                    .TextMatrix(i, col执行人) = cboExecutePoeple(0).Text
                    .Cell(flexcpData, i, col执行人) = .TextMatrix(i, col执行人)
                ElseIf optExecutePeople(opt本人).value Then
                    .TextMatrix(i, col执行人) = UserInfo.姓名
                    .Cell(flexcpData, i, col执行人) = .TextMatrix(i, col执行人)
                End If
                '执行结果
                If .TextMatrix(i, col执行结果) = "未执行" And cboExecuteResult.Text = "完成" Then
                    .TextMatrix(i, COL未执行原因) = ""
                    .Cell(flexcpData, i, COL未执行原因) = .TextMatrix(i, COL未执行原因)
                End If
                .TextMatrix(i, col执行结果) = cboExecuteResult.Text
                .Cell(flexcpData, i, col执行结果) = .TextMatrix(i, col执行结果)
                
            End If
        Next
    End With
End Sub

Private Sub FuncCancleExec()
'功能：取消执行
    Dim arrSQL() As Variant
    Dim i As Long, j As Long
    Dim arrAdvcie() As Variant
    Dim lngBegin As Long, lngEnd As Long
    Dim blnTrans As Boolean
    Dim datCur As Date
    
    arrSQL = Array()
    arrAdvcie = Array()
    datCur = zlDatabase.Currentdate
    With vsgExecAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col医嘱ID) <> "" And .RowData(i) = "Begin" And .Cell(flexcpData, i, col选择) = "1" Then
                If CanUnExec(CDate(.TextMatrix(i, col检查时间)), datCur) Then
                    '检查是否核对
                    If .TextMatrix(i, col核对时间) <> "" Then
                        MsgBox "医嘱：" & .TextMatrix(i, col医嘱内容) & " 已经核对，请取消核对后再试。", vbInformation, "取消执行"
                        Exit Sub
                    Else
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱执行_Delete(" & .TextMatrix(i, col待执行ID) & "," & .TextMatrix(i, col发送号) & ",To_date('" & .TextMatrix(i, COL原始执行时间) & "','YYYY-MM-DD HH24:MI:ss'),0,1," & mlng病区ID & ")"
                        ReDim Preserve arrAdvcie(UBound(arrAdvcie) + 1)
                            arrAdvcie(UBound(arrAdvcie)) = i
                    End If
                End If
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    Screen.MousePointer = 0
    If vsgExecAdvice.TextMatrix(1, col医嘱ID) = "" Then
        stbThis.Panels(2).Text = "没有可取消执行的医嘱。"
    Else
        If UBound(arrSQL) = -1 Then
            MsgBox "请勾选您需要取消执行的医嘱。", vbInformation, Me.Caption
            Exit Sub
        End If
        stbThis.Panels(2).Text = "取消成功，本次共取消执行了 " & UBound(arrSQL) + 1 & " 条医嘱。"
    End If
    '删除取消的医嘱
    For i = UBound(arrAdvcie) To 0 Step -1
        If vsgExecAdvice.TextMatrix(Val(arrAdvcie(i)), col医嘱ID) <> "" Then
            If Not RowIn一并给药(Val(arrAdvcie(i)), lngBegin, lngEnd, vsgExecAdvice) Then
                lngBegin = Val(arrAdvcie(i)): lngEnd = Val(arrAdvcie(i))
            End If
            
            For j = lngEnd To lngBegin Step -1
                vsgExecAdvice.RemoveItem j
            Next
        End If
    Next
    If vsgExecAdvice.Rows = 1 Then vsgExecAdvice.AddItem ""
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncSaveExecute()
'功能：保存执行待执行的医嘱
    Dim arrSQL() As Variant
    Dim i As Long, j As Long, s As Long
    Dim blnTrans As Boolean
    
    If Not CheckData(vsgWaitExecAdvice) Then Exit Sub
    arrSQL = Array()
    With vsgWaitExecAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col医嘱ID) <> "" And .RowData(i) = "Begin" And .Cell(flexcpData, i, col选择) = "1" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱执行_Insert(" & .TextMatrix(i, col待执行ID) & "," & .TextMatrix(i, col发送号) & ",To_date('" & .TextMatrix(i, col要求时间) & "','YYYY-MM-DD HH24:MI')," & _
                IIF(Val(.TextMatrix(i, col本次数次)) = 0, 1, Val(.TextMatrix(i, col本次数次))) & ",'" & .TextMatrix(i, col备注) & "','" & .TextMatrix(i, col执行人) & "',To_date('" & .TextMatrix(i, col执行时间) & _
                "','YYYY-MM-DD HH24:MI'),0,1," & IIF(.TextMatrix(i, col执行结果) = "未执行", "0", "1") & ",'" & .TextMatrix(i, COL未执行原因) & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlng病区ID & ")"
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        j = j + 1
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    Screen.MousePointer = 0
    mblnWaitIsUpdate = False
    cboExecutePoeple(0).Text = ""
    If vsgWaitExecAdvice.TextMatrix(1, col医嘱ID) = "" Then
        stbThis.Panels(2).Text = "没有可执行的医嘱。"
    Else
        stbThis.Panels(2).Text = "保存成功，本次共执行了 " & UBound(arrSQL) + 1 & " 条医嘱。"
    End If
    '清空医嘱列表
    vsgWaitExecAdvice.Rows = 1
    vsgWaitExecAdvice.AddItem ""
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    For i = 1 To vsgWaitExecAdvice.Rows - 1
        If vsgWaitExecAdvice.TextMatrix(i, col医嘱ID) <> "" And vsgWaitExecAdvice.RowData(i) = "Begin" And vsgWaitExecAdvice.Cell(flexcpData, i, col选择) = "1" Then
            s = s + 1
            If s = j Then
                vsgWaitExecAdvice.Row = i: vsgWaitExecAdvice.ShowCell i, COL_姓名
                Me.Refresh
                Exit For
            End If
        End If
    Next
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncSaveUpdate()
'功能：保存已修改的医嘱
    Dim arrSQL() As Variant
    Dim i As Long
    Dim blnTrans As Boolean
    Dim datCur As Date
    
    If Not CheckData(vsgExecAdvice) Then Exit Sub
    arrSQL = Array()
    datCur = zlDatabase.Currentdate
    With vsgExecAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col医嘱ID) <> "" And .RowData(i) = "Begin" And .TextMatrix(i, col是否修改) = "1" Then
                If CanUnExec(CDate(.TextMatrix(i, col检查时间)), datCur) Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱执行_Update(" & "To_date('" & .TextMatrix(i, COL原始执行时间) & "','YYYY-MM-DD HH24:MI')," & .TextMatrix(i, col待执行ID) & "," & .TextMatrix(i, col发送号) & ",To_date('" & .TextMatrix(i, col要求时间) & "','YYYY-MM-DD HH24:MI')," & _
                        IIF(Val(.TextMatrix(i, col本次数次)) = 0, 1, Val(.TextMatrix(i, col本次数次))) & ",'" & .TextMatrix(i, col备注) & "','" & .TextMatrix(i, col执行人) & "',To_date('" & .TextMatrix(i, col执行时间) & _
                        "','YYYY-MM-DD HH24:MI')," & IIF(.TextMatrix(i, col执行结果) = "未执行", "0", "1") & ",'" & .TextMatrix(i, COL未执行原因) & "',0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlng病区ID & ")"
                End If
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    Screen.MousePointer = 0
    mblnExecIsUpdate = False
    If vsgExecAdvice.TextMatrix(1, col医嘱ID) = "" Then
        stbThis.Panels(2).Text = "没有可修改的已执行医嘱。"
    Else
        stbThis.Panels(2).Text = "修改成功，本次共修改了 " & UBound(arrSQL) + 1 & " 条医嘱。"
    End If
    '恢复前景颜色
    vsgExecAdvice.Cell(flexcpForeColor, 1, col选择, vsgExecAdvice.Rows - 1, col选择) = vbBlack
    vsgExecAdvice.Cell(flexcpForeColor, 1, col执行时间, vsgExecAdvice.Rows - 1, col备注) = vbBlack
    '恢复修改的标识和原执行时间
    With vsgExecAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col医嘱ID) <> "" And .RowData(i) = "Begin" And .TextMatrix(i, col是否修改) = "1" Then
                If CanUnExec(CDate(.TextMatrix(i, col检查时间)), datCur) Then
                    .TextMatrix(i, col是否修改) = ""
                    .TextMatrix(i, COL原始执行时间) = .TextMatrix(i, col执行时间)
                End If
            End If
        Next
    End With
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        If tbcSub.Selected.Tag = "待执行医嘱" Then
            LoadAdvice
        Else
            LoadAdvice True
        End If
    ElseIf KeyCode = vbKey1 And Shift = 4 Then
        tbcSub.Item(0).Selected = True
    ElseIf KeyCode = vbKey2 And Shift = 4 Then
        tbcSub.Item(1).Selected = True
    End If
End Sub

Private Sub Form_Load()
    Dim strHead As String
    Dim strTbc As String
    Dim objPane As Pane
    
    mblnWaitIsUpdate = False
    mblnExecIsUpdate = False
    
    'commandbar
    '-----------------------------------------------------
    Call InitCommandBar
    
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        If mbytUseType = T医嘱核对 Then
            strTbc = "核对"
        Else
            strTbc = "执行"
        End If
        .InsertItem(0, "待" & strTbc & "医嘱(&1)", picWaitExecute.Hwnd, 0).Tag = "待执行医嘱"
        .InsertItem(1, "已" & strTbc & "医嘱(&2)", picExecuted.Hwnd, 0).Tag = "已执行医嘱"
        
        .Item(0).Selected = True
    End With
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 180, 400, DockLeftOf, Nothing)
    objPane.Title = "病人列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    'VSFlexGrid
    '-----------------------------------------------------
    strHead = ",400,1;床位,750,1;姓名,850,1;性别,450,1;住院号,1200,1;住院次数,950,1;期效,450,1;医嘱内容,2500,1;总量;单量,700,1;本次数次,840,7;给药途径;要求时间,1800,1;执行时间,1800,1;执行人,950,1;核对时间,1550,1;核对人,950,1;执行结果,950,1;未执行原因,1000,1;备注,2000,1;医嘱ID;相关ID;诊疗类别;病人ID;主页ID;频率;总量;发送号;待执行ID;医嘱状态;在院;登记人;原始执行时间;是否修改;最后执行人;操作类型;执行分类;病区ID;执行状态;开始执行时间;检查时间"
    Call InitTable(vsgWaitExecAdvice, strHead)
    vsgWaitExecAdvice.ExtendLastCol = True
    vsgWaitExecAdvice.ExplorerBar = flexExSortShow
    
    Call InitTable(vsgExecAdvice, strHead)
    vsgExecAdvice.ExtendLastCol = True
    vsgExecAdvice.ExplorerBar = flexExSortShow
    
    Set mrsDefine = InitAdviceDefine
    Call InitPageData
    Call LoadPatiInfo
    
    Call RestoreWinState(Me, App.ProductName)
    
    '医嘱核对
    If mbytUseType = T医嘱核对 Then
        Me.Caption = "医嘱批量核对"
        lblInfo(lbl时间范围).Caption = "执行时间"
        picExec.Visible = False
        lblInfo(11).Caption = "核对人"
        lblInfo(lbl期效).Visible = False
        chk期效(chk长嘱).Visible = False
        chk期效(chk临嘱).Visible = False
        chkType(Type输液).Visible = False
        chkType(Type注射).Visible = False
        chkType(Type口服).Visible = False
        chkType(Type采集).Visible = False
        chkType(Type其它医嘱).Visible = False
        chkType(Type其它给药途径).Visible = False
        chkType(Type治疗).Visible = False
        chkType(Type中药服法).Visible = False
        chkType(Type输血).Top = lblInfo(lbl类别).Top
        chkType(Type皮试).Top = chkType(Type输血).Top
        picFitter.Height = chkType(Type皮试).Top + 350
        lblInfo(lbl类别).Top = lblInfo(lbl期效).Top
    Else
        chkType(Type输血).Visible = False
    End If
    
    If DeptIsWoman(0, Get科室IDs(mlng病区ID)) Then
        fraBaby.Visible = True
        '医嘱处理范围
        mint医嘱处理范围 = Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0"))
        optBaby(mint医嘱处理范围).value = True
    End If
    If Not (mbytUseType = T医嘱核对 Or Val(zlDatabase.GetPara(51)) = 1) Then
        optExecutePeople(opt本人).value = True
        cboExecutePoeple(cbo执行人取消).Visible = False
        lblInfo(lbl执行人).Visible = False
    End If
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
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, " 查询(&Q)"): objControl.BeginGroup = True
        objControl.ToolTipText = "读取待执行/已执行的数据"

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " 保存(&S)")
        objControl.BeginGroup = True
        objControl.ToolTipText = "对已经勾选的医嘱进行执行/对已经修改的内容进行保存。"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消执行(&C)")
        objControl.ToolTipText = "对已经勾选的医嘱进行取消执行的操作。"
        objControl.IconId = 3651
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "核对(&A)")
        objControl.ToolTipText = "对已经勾选的皮试或输血医嘱进行核对的操作。"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit, "取消核对(&C)")
        objControl.ToolTipText = "对已经勾选的皮试或输血医嘱进行取消核对的操作。"
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " 退出(&E)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With

End Sub

Private Sub LoadAdvice(Optional ByVal blnIsExecute As Boolean)
'功能：加载医嘱
'参数：blnIsExecute=true加载已执行/已核对医嘱,false为加载待执行/待核对医嘱
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim lngID As Long       '用于定位
    Dim strFormat As String
    Dim strTmp As String
    Dim strFitter As String
    Dim strPatis As Variant
    Dim blnDo As Boolean
    Dim lngCount As Long   '需要执行的医嘱数(一并给药算1条)
    Dim strExecPeople As String
    Dim strFace As String  '当前刷新界面 00－待执行，01－待核对，10－已执行，11－已核对。
    Dim lngCntH As Long, lngCntN As Long
    Dim blnALL As Boolean
    Dim strWhere As String
    Dim str本科本病区 As String
    
    On Error GoTo errH
    '病人
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record.Tag = "1" Then
                strPatis = strPatis & "," & rptPati.Rows(i).Record(COL_病人ID).value & ":" & rptPati.Rows(i).Record(COL_主页ID).value
                lngCntH = lngCntH + 1
            Else
                lngCntN = lngCntN + 1
            End If
        End If
    Next
    
    If lngCntH * lngCntN = 0 Then
        blnALL = True
    End If
    strPatis = Mid(strPatis, 2)
    If strPatis = "" Then
        MsgBox "请选择需要查询的病人。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    strFace = IIF(blnIsExecute, "1", "0") & IIF(mbytUseType <> T医嘱核对, "0", "1")

    If strFace = "00" Or strFace = "10" Then
        '登记功能，执行和取消
        If chk期效(chk长嘱).value = 1 And chk期效(chk临嘱).value <> 1 Then
            strWhere = " And a.医嘱期效=0 "
        ElseIf chk期效(chk长嘱).value <> 1 And chk期效(chk临嘱).value = 1 Then
            strWhere = " And a.医嘱期效=1 "
        End If
        '诊疗类别
        strFitter = IIF(chkType(Type输液).value, " a.诊疗类别='E' And c.操作类型='2' and c.执行分类=1", "")
        strFitter = strFitter & IIF(chkType(Type注射).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='E' And c.操作类型='2' and c.执行分类=2", "")
        strFitter = strFitter & IIF(chkType(Type口服).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='E' And c.操作类型='2' and c.执行分类=4", "")
        strFitter = strFitter & IIF(chkType(Type其它给药途径).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='E' And c.操作类型='2' and c.执行分类=0", "")
        strFitter = strFitter & IIF(chkType(Type治疗).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='E' And c.操作类型 in('0','5')", "")
        strFitter = strFitter & IIF(chkType(Type采集).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='E' And c.操作类型='6' ", "")
        strFitter = strFitter & IIF(chkType(Type中药服法).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='E' And c.操作类型='4' ", "")
        '其他
        strFitter = strFitter & IIF(chkType(Type其它医嘱).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='4' " & _
                    " Or a.诊疗类别='H' And c.操作类型='0'" & _
                    " Or a.诊疗类别='D'" & _
                    " Or a.诊疗类别='I'" & _
                    " Or a.诊疗类别='L'" & _
                    " Or a.诊疗类别='Z' And c.操作类型='0'" & _
                    " Or a.诊疗类别='K'" _
                    , "")
        strFitter = strFitter & IIF(chkType(Type皮试).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='E' And c.操作类型='1' and c.执行分类=3", "")
    Else
        '核对功能，执行和取消   期效,核对跳过期效  诊疗类别-核对只包含输血和皮试
        strFitter = strFitter & IIF(chkType(Type输血).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='K' ", "")
        If chkType(Type皮试).value = 0 And chkType(Type输血).value = 0 Then
            strFitter = strFitter & IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='K' Or a.诊疗类别='E' And c.操作类型='1'"
        End If
        strFitter = strFitter & IIF(chkType(Type皮试).value, IIF(strFitter <> "", " Or ", "") & " a.诊疗类别='E' And c.操作类型='1' and c.执行分类=3", "")
    End If
    
    strWhere = strWhere & IIF(strFitter <> "", " And (" & strFitter & ")", "")
    If blnALL Then
        '如果是全选，则使用病区ID访问
        strWhere = strWhere & " And (a.病人ID,A.主页ID) In (Select h.病人id,h.主页id From 病案主页 H,在院病人 R Where h.病人id=r.病人id " & _
            IIF(mint场合 = 1, " And h.当前病区ID+0=R.病区ID  And (R.病区ID=[5] Or h.婴儿病区ID=[5])", "  And h.出院科室id+0=R.科室ID And (R.科室ID=[5] Or h.婴儿科室ID=[5])") & ")"
    Else
        strWhere = strWhere & " And (a.病人ID,A.主页ID) In(Select /*+cardinality(x,10)*/  x.C1 As 病人ID,x.C2 As 主页ID From Table(f_Num2list2([6])) x )"
    End If
    
    str本科本病区 = " Exists (Select 1 From 病案主页 H Where a.病人id=h.病人id And a.主页id=h.主页id And (h.出院科室id=g.执行部门id Or h.当前病区id=g.执行部门id))"
    
    Select Case strFace
    Case "00"
        If mblnWaitIsUpdate Then
            If MsgBox("待执行医嘱的内容未保存，是否要刷新？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                vsgWaitExecAdvice.Redraw = flexRDDirect
                Exit Sub
            End If
        End If
        strSQL = "Select b.要求时间, b.医嘱id, b.发送号,g.执行状态,Decode(A.执行频次,'一次性',NVL(A.总给予量,1),'需要时',NVL(A.总给予量,1),'持续性'," & _
            "NVL(A.单次用量,1),'必要时',NVL(A.单次用量,1),'不定时',NVL(A.单次用量,1),Decode(A.医嘱期效,0,NVL(A.单次用量,1),1,Decode(A.上次执行时间,b.要求时间,Decode(mod(NVL(A.总给予量,1),NVL(A.单次用量,1)),0,NVL(A.单次用量,1),mod(NVL(A.总给予量,1),NVL(A.单次用量,1))),NVL(A.单次用量,1)),1)) as 本次数次" & _
            " From 医嘱执行时间 B,病人医嘱记录 A, 诊疗项目目录 C,病人医嘱发送 G" & vbNewLine & _
            " Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And g.发送号=b.发送号 And b.医嘱ID=g.医嘱ID And (a.医嘱期效=1 or (a.医嘱期效=0 and (b.要求时间<=a.执行终止时间 or a.执行终止时间 is null)))" & _
            " And Not Exists (Select 1 From 病人医嘱执行 Where b.要求时间 = 要求时间 And b.医嘱id = 医嘱id And b.发送号 = 发送号) " & _
            " And (g.执行部门id=[4] Or Exists (Select 1 From 病人医嘱记录 D Where a.Id = d.相关id And d.执行科室id = [4]) Or " & str本科本病区 & " or A.执行性质=5 and a.执行标记=0) And b.要求时间+0 Between [1] And [2]"
    Case "01"
        '核对功能只有护士站才有，过滤当前病区和病区对应的科室的医嘱。
        strSQL = "Select b.要求时间, b.医嘱id, b.发送号, b.执行人, b.执行时间, b.执行结果, b.说明, b.执行摘要, b.登记人,b.核对人,b.核对时间,g.执行状态,B.本次数次,Decode(g.执行状态,1,g.完成时间,2,g.完成时间, B.登记时间) 检查时间 " & _
            " From 病人医嘱执行 B, 病人医嘱记录 A, 诊疗项目目录 C,病人医嘱发送 G " & _
            " Where a.Id = b.医嘱id And b.医嘱ID=g.医嘱ID And g.发送号=b.发送号 And a.诊疗项目id = c.Id " & _
            " And (g.执行部门id=[4] Or Exists (Select 1 From 病人医嘱记录 D Where a.Id = d.相关id And d.执行科室id = [4]) Or " & str本科本病区 & ") And b.核对时间 Is Null And b.执行时间 Between [1] And [2]"
    Case "10"
         '如果已经有修改了,则提示是否继续
        If mblnExecIsUpdate Then
            If MsgBox("已执行医嘱的内容未保存，是否要刷新？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                vsgExecAdvice.Redraw = flexRDDirect
                Exit Sub
            End If
        End If
        strSQL = "Select b.医嘱id, b.发送号, b.要求时间,B.本次数次, b.执行人, b.执行时间, b.执行结果, b.说明, b.执行摘要, b.登记人,b.核对人,b.核对时间 ,g.执行状态,Decode(g.执行状态,1,g.完成时间,2,g.完成时间, B.登记时间) 检查时间 " & vbNewLine & _
            " From 病人医嘱执行 B,医嘱执行时间 E,病人医嘱发送 G,病人医嘱记录 A,诊疗项目目录 C" & vbNewLine & _
            " Where e.要求时间=b.要求时间 and e.医嘱id=b.医嘱id and e.发送号=b.发送号 And g.发送号=b.发送号 And b.医嘱ID=g.医嘱ID And b.医嘱ID=a.ID and a.诊疗项目ID=c.ID" & _
            " and (g.执行部门id=[4] Or Exists (Select 1 From 病人医嘱记录 D Where a.Id = d.相关id And d.执行科室id = [4]) Or " & str本科本病区 & " or a.执行性质=5 and a.执行标记=0) And e.要求时间 Between [1] And [2] "
        If cboExecutePoeple(1).ListCount >= 0 Then
            strSQL = strSQL & IIF(cboExecutePoeple(1).Text = "", "", " And B.执行人=[3] ")
            strExecPeople = cboExecutePoeple(1).Text
        End If
    Case "11"
        '执行登记功能医生和护士在用要区分场合，医生站调用时只过当前面的科室范围内的医嘱。
        strSQL = "Select b.医嘱id, b.发送号, b.要求时间,B.本次数次, b.执行人, b.执行时间, b.执行结果, b.说明, b.执行摘要, b.登记人,b.核对人,b.核对时间 ,g.执行状态,Decode(g.执行状态,1,g.完成时间,2,g.完成时间, B.登记时间) 检查时间 " & vbNewLine & _
            " From 病人医嘱执行 B,医嘱执行时间 E,病人医嘱发送 G,病人医嘱记录 A,诊疗项目目录 C" & vbNewLine & _
            " Where e.要求时间=b.要求时间 and e.医嘱id=b.医嘱id and e.发送号=b.发送号 And g.发送号=b.发送号 And b.医嘱ID=g.医嘱ID And b.医嘱ID=a.ID and a.诊疗项目ID=c.ID and (" & _
            " g.执行部门id=[4] Or Exists (Select 1 From 病人医嘱记录 D Where a.Id = d.相关id And d.执行科室id = [4]) Or " & str本科本病区 & ") And b.核对时间 Is Not Null And b.执行时间  Between [1] And [2] "
         '如果没有代行执行的权限，也不允许取消他人的执行记录
        If Val(zlDatabase.GetPara(51)) = 0 Then
            strSQL = strSQL & " And B.执行人=[3] "
            strExecPeople = UserInfo.姓名
        Else
            If cboExecutePoeple(1).ListCount >= 0 Then
                strSQL = strSQL & IIF(cboExecutePoeple(1).Text = "", "", " And B.核对人=[3] ")
                strExecPeople = cboExecutePoeple(1).Text
            End If
        End If
    End Select
    strSQL = strSQL & strWhere
    If blnIsExecute Or mbytUseType = T医嘱核对 Then
        strSQL = "Select b.执行人, to_char(b.执行时间,'YYYY-MM-DD HH24:MI') as 执行时间, b.执行结果, b.说明 as 未执行原因, b.执行摘要 as 备注,b.登记人,b.执行时间 as 原始执行时间,b.核对人,to_char(b.核对时间,'YYYY-MM-DD HH24:MI') as 核对时间," & _
            " a.Id, b.发送号,b.医嘱id as 待执行ID,a.相关id, a.诊疗类别,b.执行状态,a.开始执行时间, A.姓名, f.出院病床 As 床号, A.性别, Decode(Nvl(a.医嘱期效, 0), 0, '长嘱', '临嘱') As 期效,a.医嘱状态,Decode(f.出院日期,NULL,1,0) as 在院,a.总给予量,a.单次用量," & vbNewLine & _
            " Decode(a.单次用量, Null, Null, decode(sign(1-A.单次用量),1,'0'||A.单次用量,A.单次用量) || c.计算单位) As 单量,  Decode(a.相关id,Null,a.医嘱内容 || ' ' || a.执行频次  ,a.医嘱内容) as 医嘱内容, to_char(b.要求时间,'YYYY-MM-DD HH24:MI') as 要求时间,f.当前病区ID,B.本次数次," & _
            " a.执行频次 As 频率, a.病人id, a.主页id, a.诊疗项目id,c.操作类型,c.执行分类,Decode(a.总给予量, Null, Null," & _
            " Round(a.总给予量 / Decode(a.病人来源, 2, d.住院包装, d.门诊包装), 5) || Decode(a.病人来源, 2, d.住院单位, d.门诊单位)) As 总量,B.检查时间,F.住院号,G.住院次数,a.医生嘱托" & vbNewLine & _
            " From (" & strSQL & ") B, 病人医嘱记录 A,病案主页 F, 诊疗项目目录 C, 药品规格 D,病人信息 G" & vbNewLine & _
            " Where (a.Id = b.医嘱id " & IIF(mbytUseType = T医嘱核对, "", "Or a.相关id = b.医嘱id") & ") And f.病人id = a.病人id And f.主页id = a.主页id And F.病人ID=G.病人ID And a.诊疗项目id = c.Id And a.收费细目id = d.药品id(+) And a.诊疗类别 Not In('C','7') And Not (a.诊疗类别='E' And c.操作类型='3') " & _
            " " & decode(mint医嘱处理范围, 1, " And nvl(a.婴儿,0) = 0 ", 2, " And nvl(a.婴儿,0) <> 0 ", "") & _
            " And (F.婴儿科室ID is null or F.婴儿科室ID is not null and (F.婴儿病区ID=[5] or F.婴儿科室ID=[5]) and NVL(A.婴儿,0)<>0 or F.婴儿科室ID is not null and (F.婴儿病区ID<>[5] and f.婴儿科室ID<>[5]) and NVL(A.婴儿,0)=0) "
    Else
        strSQL = "Select distinct" & _
            " a.Id, b.发送号,b.医嘱id as 待执行ID,a.相关id, a.诊疗类别,b.执行状态,a.开始执行时间, A.姓名, f.出院病床 As 床号, A.性别, Decode(Nvl(a.医嘱期效, 0), 0, '长嘱', '临嘱') As 期效,a.医嘱状态,Decode(f.出院日期,NULL,1,0) as 在院,a.总给予量,a.单次用量," & vbNewLine & _
            " Decode(a.单次用量, Null, Null, decode(sign(1-A.单次用量),1,'0'||A.单次用量,A.单次用量) || c.计算单位) As 单量,  Decode(a.相关id,Null,a.医嘱内容 || ' ' || a.执行频次  ,a.医嘱内容) as 医嘱内容, to_char(b.要求时间,'YYYY-MM-DD HH24:MI') as 要求时间,f.当前病区ID,B.本次数次," & _
            " a.执行频次 As 频率, a.病人id, a.主页id, a.诊疗项目id,c.操作类型,c.执行分类,Decode(a.总给予量, Null, Null," & _
            " Round(a.总给予量 / Decode(a.病人来源, 2, d.住院包装, d.门诊包装), 5) || Decode(a.病人来源, 2, d.住院单位, d.门诊单位)) As 总量,first_value(e.执行人) Over(partition By e.医嘱id Order By e.执行时间 DESC) As 最后执行人,b.要求时间, a.序号,F.住院号,G.住院次数,a.医生嘱托" & vbNewLine & _
            " From (" & strSQL & ") B, 病人医嘱记录 A,病案主页 F, 诊疗项目目录 C, 药品规格 D ,病人医嘱执行 E,病人信息 G" & vbNewLine & _
            " Where (a.Id = b.医嘱id Or a.相关id = b.医嘱id) And e.医嘱id(+) = a.Id And f.病人id = a.病人id And f.主页id = a.主页id And a.诊疗项目id = c.Id And F.病人ID=G.病人ID And a.收费细目id = d.药品id(+) And a.诊疗类别 Not In('C','7') And Not (a.诊疗类别='E' And c.操作类型='3') " & _
            " " & decode(mint医嘱处理范围, 1, " And nvl(a.婴儿,0) = 0 ", 2, " And nvl(a.婴儿,0) <> 0 ", "") & _
            " And (F.婴儿科室ID is null or F.婴儿科室ID is not null and (F.婴儿病区ID=[5] or F.婴儿科室ID=[5]) and NVL(A.婴儿,0)<>0 or F.婴儿科室ID is not null and (F.婴儿病区ID<>[5] and f.婴儿科室ID<>[5]) and NVL(A.婴儿,0)=0) "
    End If
    strSQL = strSQL & " Order By A.姓名,b.要求时间,Nvl(a.相关id,a.Id),a.id,a.序号"
    
    If strFace = "11" Or strFace = "10" Then
        vsgExecAdvice.Cell(flexcpPicture, 0, col选择) = img16.ListImages("UnCheck").Picture
        vsgExecAdvice.Cell(flexcpPictureAlignment, 0, col选择) = flexPicAlignCenterCenter
        vsgExecAdvice.ColData(col选择) = ""
    Else
        vsgWaitExecAdvice.Cell(flexcpPicture, 0, col选择) = img16.ListImages("AllCheck").Picture
        vsgWaitExecAdvice.Cell(flexcpPictureAlignment, 0, col选择) = flexPicAlignCenterCenter
        vsgWaitExecAdvice.ColData(col选择) = "Check"
        '保存过滤条件
        Call zlDatabase.SetPara("医嘱执行期效", chk期效(chk长嘱).value & chk期效(chk临嘱).value & "", glngSys, p住院医嘱发送)
        strTmp = chkType(Type输液).value & chkType(Type注射).value & chkType(Type口服).value & chkType(Type皮试).value & chkType(Type治疗).value & _
            chkType(Type采集).value & chkType(Type其它医嘱).value & chkType(Type输血).value & chkType(Type其它给药途径).value & chkType(Type中药服法).value
        Call zlDatabase.SetPara("医嘱执行范围", strTmp, glngSys, p住院医嘱发送)
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(dpkReqTime(dpk开始日期).value), CDate(dpkReqTime(dpk结束日期).value), strExecPeople, mlng病区ID, mlng医护科室ID, strPatis)

    With IIF(blnIsExecute, vsgExecAdvice, vsgWaitExecAdvice)
        .Redraw = flexRDNone
        .Rows = 1
        .ExplorerBar = 7
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                If .ColData(col选择) = "Check" Then
                    .Cell(flexcpPicture, i, col选择) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, col选择) = 1
                    .Cell(flexcpPictureAlignment, i, col选择) = flexPicAlignCenterCenter
                End If
                .TextMatrix(i, col姓名) = rsTmp!姓名 & ""
                .TextMatrix(i, col住院号) = rsTmp!住院号 & ""
                .TextMatrix(i, col住院次数) = rsTmp!住院次数 & ""
                .TextMatrix(i, col期效) = rsTmp!期效 & ""
                .TextMatrix(i, col单量) = rsTmp!单量 & ""
                .TextMatrix(i, col医嘱ID) = rsTmp!ID & ""
                .TextMatrix(i, col相关ID) = rsTmp!相关ID & ""
                .TextMatrix(i, col性别) = rsTmp!性别 & ""
                .TextMatrix(i, COL床位) = rsTmp!床号 & ""
                .TextMatrix(i, Col病人ID) = rsTmp!病人ID & ""
                .TextMatrix(i, COL主页ID) = rsTmp!主页ID & ""
                .TextMatrix(i, col诊疗类别) = rsTmp!诊疗类别 & ""
                .TextMatrix(i, col总量) = rsTmp!总量 & ""
                '输血医嘱的计量医嘱特殊处理，因为输血医嘱临嘱按频率次数发送的
                If rsTmp!期效 & "" = "临嘱" And (rsTmp!诊疗类别 & "" = "K" Or rsTmp!诊疗类别 & "" = "E" And rsTmp!操作类型 & "" = "8" And rsTmp!相关ID & "" <> "") Then
                   If rsTmp!诊疗类别 & "" = "K" Then
                        .TextMatrix(i, col本次数次) = Get输血本次数次(Val(rsTmp!ID & ""), Val(rsTmp!发送号 & ""), CDate(Format(rsTmp!要求时间 & "", "yyyy-mm-dd HH:mm:ss")), Val(rsTmp!总给予量 & ""), Val(rsTmp!单次用量 & ""))
                   Else
                        .TextMatrix(i, col本次数次) = 1
                   End If
                Else
                   .TextMatrix(i, col本次数次) = rsTmp!本次数次 & ""
                End If
                .TextMatrix(i, col发送号) = rsTmp!发送号 & ""
                .TextMatrix(i, col待执行ID) = rsTmp!待执行ID & ""
                .TextMatrix(i, COL医嘱状态) = rsTmp!医嘱状态 & ""
                .TextMatrix(i, col在院) = rsTmp!在院 & ""
                .TextMatrix(i, col操作类型) = rsTmp!操作类型 & ""
                .TextMatrix(i, col执行分类) = rsTmp!执行分类 & ""
                .TextMatrix(i, col病区ID) = rsTmp!当前病区ID & ""
                .TextMatrix(i, COL频率) = rsTmp!频率 & ""
                .TextMatrix(i, col开始执行时间) = rsTmp!开始执行时间 & ""
                .RowData(i) = IIF(.TextMatrix(i, col相关ID) = "", "Begin", "")
                '显示简洁模式下的医嘱内容
                strFormat = rsTmp!医嘱内容 & ""
                
                If .TextMatrix(i, col诊疗类别) & .TextMatrix(i, col操作类型) & .TextMatrix(i, col执行分类) = "E21" Then
                    If rsTmp!医生嘱托 & "" <> "" Then
                        strFormat = strFormat & " " & rsTmp!医生嘱托
                    End If
                End If
                
                If .TextMatrix(i, COL频率) <> "一次性" Then
                    blnDo = True
                    If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[总量]") = 0
                    If blnDo Then
                        strTmp = .TextMatrix(i, col总量)
                        If strTmp <> "" Then strFormat = strFormat & ",共" & strTmp
                    End If
                End If
                .TextMatrix(i, col医嘱内容) = strFormat
                .TextMatrix(i, col要求时间) = rsTmp!要求时间 & ""
                '可编辑列颜色
                .Cell(flexcpBackColor, i, col选择, i, col选择) = COLEditBackColor
                If mbytUseType <> T医嘱核对 Then
                    .Cell(flexcpBackColor, i, col执行时间, i, col执行人) = COLEditBackColor
                    .Cell(flexcpBackColor, i, col执行结果, i, col备注) = COLEditBackColor
                End If
                If blnIsExecute Or mbytUseType = T医嘱核对 Then
                    .TextMatrix(i, col执行时间) = rsTmp!执行时间 & ""
                    .Cell(flexcpData, i, col执行时间) = .TextMatrix(i, col执行时间)
                    '记录原始执行时间，用于删除执行
                    .TextMatrix(i, COL原始执行时间) = rsTmp!原始执行时间 & ""
                    .TextMatrix(i, col执行人) = rsTmp!执行人 & ""
                    .Cell(flexcpData, i, col执行人) = .TextMatrix(i, col执行人)
                    .TextMatrix(i, col执行结果) = IIF(rsTmp!执行结果 & "" = "0", "未执行", "完成")
                    .Cell(flexcpData, i, col执行结果) = .TextMatrix(i, col执行结果)
                    .TextMatrix(i, COL未执行原因) = rsTmp!未执行原因 & ""
                    .Cell(flexcpData, i, COL未执行原因) = .TextMatrix(i, COL未执行原因)
                    .TextMatrix(i, col备注) = rsTmp!备注 & ""
                    .Cell(flexcpData, i, col备注) = .TextMatrix(i, col备注)
                    .TextMatrix(i, col登记人) = rsTmp!登记人 & ""
                    .TextMatrix(i, col检查时间) = Format(rsTmp!检查时间, "yyyy-MM-dd HH:mm")
                    If rsTmp!在院 & "" <> "1" Then
                        '出院病人的用浅灰色显示
                        .Cell(flexcpBackColor, i, col选择, i, col选择) = &HE0E0E0
                        .Cell(flexcpBackColor, i, col执行时间, i, col备注) = &HE0E0E0
                    End If
                    If blnIsExecute Then
                        .TextMatrix(i, COL核对人) = rsTmp!核对人 & ""
                        .Cell(flexcpData, i, COL核对人) = .TextMatrix(i, COL核对人)
                        .TextMatrix(i, col核对时间) = rsTmp!核对时间 & ""
                        .Cell(flexcpData, i, col核对时间) = .TextMatrix(i, col核对时间)
                        .TextMatrix(i, col执行状态) = rsTmp!执行状态 & ""
                    Else
                        .ColHidden(COL核对人) = True
                        .ColHidden(col核对时间) = True
                    End If
                Else
                    .TextMatrix(i, COL最后执行人) = rsTmp!最后执行人 & ""
                    Call BatUpdate
                    .ColHidden(COL核对人) = True
                    .ColHidden(col核对时间) = True
                End If
                
                '需要执行的医嘱组数
                If rsTmp!相关ID & "" = "" Then lngCount = lngCount + 1
                
                rsTmp.MoveNext
                i = i + 1
            Loop
        Else
            .AddItem ""
        End If
                
        If blnIsExecute Then
            stbThis.Panels(2).Text = "共有 " & lngCount & " 条医嘱已经" & IIF(mbytUseType <> T医嘱核对, "执行", "核对") & "。"
            mblnExecIsUpdate = False
        Else
            stbThis.Panels(2).Text = "共有 " & lngCount & " 条医嘱需要" & IIF(mbytUseType <> T医嘱核对, "执行", "核对") & "。"
            mblnWaitIsUpdate = False
        End If
        '自动调整行高
        .AutoSize col医嘱内容
        .Redraw = flexRDDirect
        '恢复前景色
        .Cell(flexcpForeColor, 1, col选择, .Rows - 1, col选择) = vbBlack
        .Cell(flexcpForeColor, 1, col执行时间, .Rows - 1, col备注) = vbBlack
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get输血本次数次(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal dat要求时间 As Date, ByVal dbl总量 As Double, ByVal dbl单量 As Double) As Double
'功能：根据医嘱信息执行时间，查出输血医嘱本次数次
    Dim strSQL As String, rsTmp As Recordset
    Dim lng当前次数 As Long, i As Long
    Dim dbl总量Tmp As Double, dbl数量 As Double
    
    strSQL = "Select 要求时间 From 医嘱执行时间 Where 医嘱id = [1] And 发送号 = [2] Order By 要求时间"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get输血本次数次", lng医嘱ID, lng发送号)
    dbl总量Tmp = dbl总量
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp.RecordCount = 1 Then
                dbl数量 = dbl总量
            Else
                If i = rsTmp.RecordCount Then
                    dbl数量 = dbl总量Tmp
                Else
                    If dbl总量Tmp >= dbl单量 Then
                        dbl数量 = dbl单量
                    Else
                        dbl数量 = dbl总量Tmp
                    End If
                    dbl总量Tmp = dbl总量Tmp - dbl数量
                End If
            End If
            If CDate(Format(rsTmp!要求时间 & "", "YYYY-MM-DD HH:mm:ss")) = dat要求时间 Then
                Get输血本次数次 = dbl数量
                Exit For
            End If
            rsTmp.MoveNext
        Next
    Else
        Get输血本次数次 = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadPatiInfo()
'功能：加载病人列表
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long, lngUnitID As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngSelectRow As Long
        
    On Error GoTo errH
    lngUnitID = mlng病区ID
    If mlng婴儿病区ID <> 0 Then
        If mlng婴儿科室ID = mlng医护科室ID Or mlng婴儿病区ID = mlng医护科室ID Then
            lngUnitID = mlng婴儿病区ID
        End If
    End If
    
    str病人IDs = zlDatabase.GetPara("发送病人", glngSys, p住院医嘱发送)
    If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
        lng病区ID = Val(Split(str病人IDs, ":")(0))
        str病人IDs = Split(str病人IDs, ":")(1)
    End If
            
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng病人ID, False, False, False)
    lngSelectRow = -1
    With rptPati
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!审核标志 & "") < 1 Or gbyt病人审核方式 <> 1 Then
                Set objRecord = .Records.Add()
                objRecord.Tag = "0"
                Set objItem = objRecord.AddItem(rsTmp!病人ID & "")
                Set objItem = objRecord.AddItem(rsTmp!主页ID & "")
                Set objItem = objRecord.AddItem("")
                Set objItem = objRecord.AddItem(rsTmp!床号 & "")
                Set objItem = objRecord.AddItem(rsTmp!姓名 & "")
                    objItem.Icon = img16.ListImages.Item(IIF(rsTmp!性别 & "" = "男", "Man", "Woman")).Index - 1
                Set objItem = objRecord.AddItem(rsTmp!性别 & "")
                Set objItem = objRecord.AddItem(rsTmp!住院号 & "")
                
                
                '病人颜色
                objRecord.Item(0).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!病人类型))
                For j = 1 To objRecord.Childs.Count - 1
                    objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
                Next
                
                '上次是否选择
                If lngUnitID = lng病区ID And str病人IDs <> "" Then
                    If InStr("," & str病人IDs & ",", "," & rsTmp!病人ID & ",") > 0 Or str病人IDs = "ALL" Then
                        objRecord.Item(COL_选择).Icon = img16.ListImages.Item("Check").Index - 1
                        objRecord.Tag = "1"
                        lngSelectRow = objRecord.Index
                    End If
                ElseIf rsTmp!病人ID = mlng病人ID Then
                    objRecord.Item(COL_选择).Icon = img16.ListImages.Item("Check").Index - 1
                    objRecord.Tag = "1"
                    lngSelectRow = objRecord.Index
                End If
            End If
            rsTmp.MoveNext
        Next
        .Populate
        If lngSelectRow <> -1 Then Set .FocusedRow = .Rows(lngSelectRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitTable(vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long

    With rptPati
        
        Set objCol = .Columns.Add(COL_病人ID, "病人ID", 0, False)
        Set objCol = .Columns.Add(COL_主页ID, "主页ID", 0, False)
        Set objCol = .Columns.Add(COL_选择, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(COL_床号, "床号", 45, True)
        Set objCol = .Columns.Add(COL_姓名, "姓名", 80, True)
        Set objCol = .Columns.Add(COL_性别, "性别", 30, True)
        Set objCol = .Columns.Add(COL_住院号, "住院号", 60, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub

Private Sub InitPageData()
'功能：初始化界面
    Dim curDate As Date
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim i As Long
    Dim strTmp As String
    Dim objCbo As Object
    
    curDate = zlDatabase.Currentdate
    
    dpkExecuteTime.value = curDate
    dpkReqTime(dpk开始日期).value = Format(curDate, "yyyy-MM-dd 00:00:00")
    dpkReqTime(dpk结束日期).value = Format(curDate, "yyyy-MM-dd 23:59:59")
    
    cboExecuteResult.AddItem "完成"
    cboExecuteResult.AddItem "未执行"
    cboExecuteResult.ListIndex = 0
    vsgWaitExecAdvice.ColComboList(col执行结果) = "完成|未执行"
    vsgExecAdvice.ColComboList(col执行结果) = "完成|未执行"
    
    strSQL = "Select 名称 From 医嘱未执行原因"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        vsgWaitExecAdvice.ColComboList(COL未执行原因) = vsgWaitExecAdvice.ColComboList(COL未执行原因) & "|" & rsTmp!名称 & ""
        vsgExecAdvice.ColComboList(COL未执行原因) = vsgExecAdvice.ColComboList(COL未执行原因) & "|" & rsTmp!名称 & ""
        rsTmp.MoveNext
    Loop
    
    vsgWaitExecAdvice.ColComboList(COL未执行原因) = Mid(vsgWaitExecAdvice.ColComboList(COL未执行原因), 2)
    vsgExecAdvice.ColComboList(COL未执行原因) = Mid(vsgExecAdvice.ColComboList(COL未执行原因), 2)
    
    strSQL = "Select Distinct a.Id, a.编号, a.姓名" & vbNewLine & _
            "From 人员表 A, 部门人员 B, 人员性质说明 C" & vbNewLine & _
            "Where a.Id = b.人员id And a.Id = c.人员id And (c.人员性质 = '护士' or c.人员性质 = '医生') And b.部门id = [1]" & _
            " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) Order By a.姓名"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病区ID)
    
    cboExecutePoeple(cbo执行人登记).AddItem ""
    cboExecutePoeple(cbo执行人取消).AddItem ""
    
    Do While Not rsTmp.EOF
        For i = 0 To 1
            With cboExecutePoeple(i)
                .AddItem rsTmp!姓名 & ""
                .ItemData(.NewIndex) = rsTmp!ID & ""
                If rsTmp!姓名 = UserInfo.姓名 Then
                    .ListIndex = .ListCount - 1
                End If
            End With
        Next
        rsTmp.MoveNext
    Loop
    
    '过滤条件
    strTmp = zlDatabase.GetPara("医嘱执行期效", glngSys, p住院医嘱发送, "11")
    chk期效(chk长嘱).value = Val(Mid(strTmp, 1, 1))
    chk期效(chk临嘱).value = Val(Mid(strTmp, 2, 1))
    strTmp = zlDatabase.GetPara("医嘱执行范围", glngSys, p住院医嘱发送, "1111111111")
    chkType(Type输液).value = Val(Mid(strTmp, 1, 1))
    chkType(Type注射).value = Val(Mid(strTmp, 2, 1))
    chkType(Type口服).value = Val(Mid(strTmp, 3, 1))
    chkType(Type皮试).value = Val(Mid(strTmp, 4, 1))
    chkType(Type治疗).value = Val(Mid(strTmp, 5, 1))
    chkType(Type采集).value = Val(Mid(strTmp, 6, 1))
    chkType(Type其它医嘱).value = Val(Mid(strTmp, 7, 1))
    chkType(Type输血).value = Val(Mid(strTmp, 8, 1))
    chkType(Type其它给药途径).value = Val(Mid(strTmp, 9, 1))
    chkType(Type中药服法).value = Val(Mid(strTmp, 10, 1))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim str病人IDs As String
    
    If mblnWaitIsUpdate Then
        tbcSub.Item(0).Selected = True
        If MsgBox("待执行医嘱的内容未保存，是否要退出？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    If mblnExecIsUpdate Then
        tbcSub.Item(1).Selected = True
        If MsgBox("已执行医嘱的内容未保存，是否要退出？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    '保存报表病人设置
    str病人IDs = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            str病人IDs = str病人IDs & "," & rptPati.Rows(i).Record(COL_病人ID).value
        End If
    Next
    str病人IDs = Mid(str病人IDs, 2)
    If str病人IDs <> "" Then
        If UBound(Split(str病人IDs, ",")) = 0 And Val(str病人IDs) = mlng病人ID Then
            Call zlDatabase.SetPara("发送病人", "", glngSys, p住院医嘱发送)
        Else
            Call zlDatabase.SetPara("发送病人", mlng病区ID & ":" & str病人IDs, glngSys, p住院医嘱发送)
        End If
    End If

    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optBaby_Click(Index As Integer)
    mint医嘱处理范围 = Index
End Sub

Private Sub optExecutePeople_Click(Index As Integer)
    cboExecutePoeple(cbo执行人登记).Enabled = optExecutePeople(opt指定人员).value
End Sub

Private Sub optExecuteTime_Click(Index As Integer)
    If Index = 1 Then
        dpkExecuteTime.Enabled = True
    Else
        dpkExecuteTime.Enabled = False
    End If
End Sub

Private Sub picExecuted_Resize()
    On Error Resume Next
    vsgExecAdvice.Top = 0
    vsgExecAdvice.Left = 0
    vsgExecAdvice.Width = picExecuted.Width
    vsgExecAdvice.Height = picExecuted.Height - vsgExecAdvice.Top
End Sub

Private Sub picWaitExecute_Resize()
    On Error Resume Next
    vsgWaitExecAdvice.Left = 0
    vsgWaitExecAdvice.Top = IIF(mbytUseType <> T医嘱核对, picExec.Height, 0)
    vsgWaitExecAdvice.Height = picWaitExecute.Height - vsgWaitExecAdvice.Top
    vsgWaitExecAdvice.Width = picWaitExecute.Width
End Sub

Private Sub picPati_Resize()
    Dim lngTmp As Long
    
    On Error Resume Next
    
    picPati.Height = tbcSub.Height - 400
    If tbcSub.Selected.Tag = "待执行医嘱" Then
        If mbytUseType <> T医嘱核对 Then
            picFitter.Height = 2850
        Else
            picFitter.Height = 1800
        End If
    Else
        If mbytUseType <> T医嘱核对 Then
            picFitter.Height = 3170
        Else
            picFitter.Height = 2120
        End If
    End If

    lblInfo(lbl病区).Top = 60
    lblInfo(lbl病区).Left = 120
    rptPati.Left = 120
    rptPati.Top = 400
    rptPati.Width = picPati.Width - rptPati.Left
    
    rptPati.Height = picPati.Height - rptPati.Top - picFitter.Height - 300
    fraBaby.Top = rptPati.Top + rptPati.Height + 100
    fraBaby.Left = rptPati.Left + 20
    picFitter.Top = picPati.Height - picFitter.Height
    picFitter.Left = rptPati.Left - 80
    
    lngTmp = dpkReqTime(dpk结束日期).Top + dpkReqTime(dpk结束日期).Height + 50
    picPanel.Left = 0
    If tbcSub.Selected.Tag = "待执行医嘱" Then
        lblInfo(lbl执行人).Visible = False
        cboExecutePoeple(cbo执行人取消).Visible = False
        picPanel.Top = lngTmp + 50
    Else
        lblInfo(lbl执行人).Visible = True
        cboExecutePoeple(cbo执行人取消).Visible = True
        cboExecutePoeple(cbo执行人取消).Top = lngTmp
        picPanel.Top = cboExecutePoeple(cbo执行人取消).Top + cboExecutePoeple(cbo执行人取消).Height + 70
    End If
    
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptPati.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptPati_RowDblClick(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(COL_选择))
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '如果点击表头的图片，就选中全部
    If Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COL_选择 Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptPati.Columns(COL_选择).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(COL_选择).Icon = img16.ListImages("Check").Index - 1
                            rptPati.Rows(i).Record.Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptPati.Columns(COL_选择).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(COL_选择).Icon = -1
                            rptPati.Rows(i).Record.Tag = "0"
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(COL_选择).Icon = -1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(COL_选择).Icon = img16.ListImages.Item("Check").Index - 1
        Row.Record.Tag = "1"
    End If
    rptPati.Populate
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If .TextMatrix(lngRow, col诊疗类别) = "" Then Exit Function
        If .TextMatrix(lngRow, col诊疗类别) = "诊疗类别" Then Exit Function
        '加要求时间是因为有可能同一条医嘱的多条执行记录在一起
        If .TextMatrix(lngRow - 1, col要求时间) = .TextMatrix(lngRow, col要求时间) And (Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 Or Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col医嘱ID)) Or Val(.TextMatrix(lngRow - 1, col医嘱ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow - 1, col医嘱ID)) <> 0) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If .TextMatrix(lngRow + 1, col要求时间) = .TextMatrix(lngRow, col要求时间) And (Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow + 1, col相关ID)) <> 0 Or Val(.TextMatrix(lngRow + 1, col医嘱ID)) = Val(.TextMatrix(lngRow, col相关ID)) Or Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.TextMatrix(lngRow, col医嘱ID))) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If .TextMatrix(i, col要求时间) = .TextMatrix(lngRow, col要求时间) And (Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 And Val(.TextMatrix(i, col医嘱ID)) <> Val(.TextMatrix(lngRow, col医嘱ID)) Or Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col医嘱ID)) Or Val(.TextMatrix(i, col医嘱ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(i, col医嘱ID)) <> 0) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If .TextMatrix(i, col要求时间) = .TextMatrix(lngRow, col要求时间) And (Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 And Val(.TextMatrix(i, col医嘱ID)) <> Val(.TextMatrix(lngRow, col医嘱ID)) Or Val(.TextMatrix(i, col医嘱ID)) = Val(.TextMatrix(lngRow, col相关ID)) Or Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col医嘱ID))) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        Else
            .RowData(lngRow) = "Begin"
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Me.Visible Then Exit Sub
    If Item.Tag = "待执行医嘱" Then
        'Call LoadAdvice
    Else
'        Call LoadAdvice(True)
    End If
End Sub

Private Sub vsgExecAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsgExecAdvice.Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor And vsgExecAdvice.RowData(NewRow) = "Begin" Then
        If NewCol = col执行人 Then
            If (Val(zlDatabase.GetPara(51)) = 1 And mbytUseType = T医嘱执行 Or mbytUseType = T医嘱核对) Then
                vsgExecAdvice.ComboList = "..."
                vsgExecAdvice.Editable = flexEDKbdMouse
            Else
                vsgExecAdvice.ComboList = ""
                vsgExecAdvice.Editable = flexEDNone
            End If
            Exit Sub
        Else
            vsgExecAdvice.ComboList = ""
        End If
        If NewCol = COL未执行原因 And vsgExecAdvice.TextMatrix(NewRow, col执行结果) = "完成" Then
            vsgExecAdvice.FocusRect = flexFocusNone
            vsgExecAdvice.Editable = flexEDNone
        Else
            vsgExecAdvice.FocusRect = flexFocusHeavy
            If NewCol <> col选择 Then
                vsgExecAdvice.Editable = flexEDKbdMouse
            Else
                vsgExecAdvice.Editable = flexEDNone
            End If
        End If
    Else
        vsgExecAdvice.FocusRect = flexFocusNone
        vsgExecAdvice.Editable = flexEDNone
        vsgExecAdvice.ComboList = ""
    End If
End Sub

Private Sub vsgExecAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim vPoint As PointAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset
    
    If Col = col执行人 Then
        With vsgExecAdvice
            strSQL = "Select a.Id, a.编号, a.姓名" & vbNewLine & _
                        "From 人员表 A, 部门人员 B, 人员性质说明 C" & vbNewLine & _
                        "Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = '护士' And b.部门id = [1]" & _
                        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
            
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病区护士", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    mlng病区ID)
                    
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    If rsTmp!姓名 & "" <> .Cell(flexcpData, Row, Col) Then
                        .TextMatrix(Row, Col) = rsTmp!姓名 & ""
                        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                        mblnExecIsUpdate = True
                        .TextMatrix(Row, col是否修改) = "1"
                        '设置颜色为深蓝色字体
                        .Cell(flexcpForeColor, Row, Col) = &HFF0000
                    End If
                Else
                    MsgBox "当前病区下没有可选的护士。", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Sub vsgExecAdvice_DblClick()
    With vsgExecAdvice
        If .MouseCol = col选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsgExecAdvice_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsgExecAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsgExecAdvice
        lngLeft = col选择: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col本次数次: lngRight = col备注
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowIn一并给药(Row, lngBegin, lngEnd, vsgExecAdvice) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If .TextMatrix(Row, col相关ID) = "" Then
            vRect.Top = Bottom - 1 '相关ID为空的行文字保留
            vRect.Bottom = Bottom - 1
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Between(Col, col选择, col选择) Or (Between(Col, col执行时间, col执行人) Or Between(Col, col执行结果, col备注)) And mbytUseType <> T医嘱核对 Then
                SetBkColor hDC, OS.SysColor2RGB(COLEditBackColor)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsgExecAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsgExecAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsgExecAdvice_KeyPress(KeyAscii As Integer)
    With vsgExecAdvice
        If .Col = col备注 Or .Col = col执行人 Then
            If KeyAscii = Asc("*") And .Col = col执行人 Then
                KeyAscii = 0
                Call vsgExecAdvice_CellButtonClick(.Row, .Col)
                Exit Sub
            End If
            .ComboList = "" '使按钮状态进入输入状态
        ElseIf .Col = col选择 And KeyAscii = vbKeySpace Then
            Call ExecCheck(vsgExecAdvice)
        End If
    End With
End Sub

Private Sub vsgExecAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vPoint As PointAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset
    
    With vsgExecAdvice
        '检查是否是当前操作员登记的
        If .TextMatrix(Row, col登记人) <> UserInfo.姓名 And .EditText <> .TextMatrix(Row, Col) Then
            MsgBox "只能取消和修改自己登记的医嘱。", vbInformation, Me.Caption
            .EditText = .Cell(flexcpData, Row, Col)
            Cancel = True
            Exit Sub
        End If
        '检查是否核对
        If .TextMatrix(Row, col核对时间) <> "" And .EditText <> .TextMatrix(Row, Col) Then
            MsgBox "该医嘱已经核对，请取消核对后在试。", vbInformation, Me.Caption
            .EditText = .Cell(flexcpData, Row, Col)
            Cancel = True
            Exit Sub
        End If
        If Col = col执行时间 Then
            If Not IsDate(.EditText) Then
                MsgBox "请输入正确的日期。", vbInformation, Me.Caption
                .EditText = .Cell(flexcpData, Row, Col)
                Cancel = True
                Exit Sub
            End If
        ElseIf Col = col备注 Then
            If zlCommFun.ActualLen(.EditText) > 200 Then
                MsgBox "备注文字不能超出100个汉字或200个字母。", vbInformation, Me.Caption
                .EditText = .Cell(flexcpData, Row, Col)
                Cancel = True
                Exit Sub
            End If
        ElseIf Col = col执行结果 Then
            If .Cell(flexcpData, Row, Col) = "未执行" And .EditText = "完成" Then
                .TextMatrix(Row, COL未执行原因) = ""
                .Cell(flexcpData, Row, COL未执行原因) = ""
            End If
        ElseIf Col = col执行人 Then
            strSQL = "Select a.Id, a.编号, a.姓名" & vbNewLine & _
                        " From 人员表 A, 部门人员 B, 人员性质说明 C" & vbNewLine & _
                        " Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = '护士'" & _
                        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " And b.部门id = [1] And (a.编号=[2] or a.姓名 Like [3] or a.简码 Like [4])"
            
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病区护士", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    mlng病区ID, .EditText, gstrLike & .EditText & "%", gstrLike & UCase(.EditText) & "%")
                    
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    If Not rsTmp Is Nothing Then
                        .EditText = rsTmp!姓名 & ""
                    Else
                        MsgBox "没查找到您指定的护士。", vbInformation, Me.Caption
                        .EditText = .Cell(flexcpData, Row, Col)
                        Cancel = True
                        Exit Sub
                    End If
                End If
        End If
        
        If .EditText <> .Cell(flexcpData, Row, Col) Then
            .TextMatrix(Row, Col) = .EditText
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            mblnExecIsUpdate = True
            .TextMatrix(Row, col是否修改) = "1"
            '设置颜色为深蓝色字体
            .Cell(flexcpForeColor, Row, Col) = &HFF0000
        End If
    End With
End Sub

Private Sub vsgWaitExecAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsgWaitExecAdvice.Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor And vsgWaitExecAdvice.RowData(NewRow) = "Begin" Then
        If NewCol = col执行人 Then
            If (Val(zlDatabase.GetPara(51)) = 1 And mbytUseType = T医嘱执行 Or mbytUseType = T医嘱核对) Then
                vsgWaitExecAdvice.ComboList = "..."
                vsgExecAdvice.Editable = flexEDKbdMouse
            Else
                vsgWaitExecAdvice.ComboList = ""
                vsgExecAdvice.Editable = flexEDNone
            End If
            Exit Sub
        Else
            vsgWaitExecAdvice.ComboList = ""
        End If
        If NewCol = COL未执行原因 And vsgWaitExecAdvice.TextMatrix(NewRow, col执行结果) = "完成" Then
            vsgWaitExecAdvice.FocusRect = flexFocusNone
            vsgWaitExecAdvice.Editable = flexEDNone
        Else
            vsgWaitExecAdvice.FocusRect = flexFocusHeavy
            If NewCol <> col选择 Then
                vsgWaitExecAdvice.Editable = flexEDKbdMouse
            Else
                vsgWaitExecAdvice.Editable = flexEDNone
            End If
        End If
    Else
        vsgWaitExecAdvice.FocusRect = flexFocusNone
        vsgWaitExecAdvice.Editable = flexEDNone
        vsgWaitExecAdvice.ComboList = ""
    End If
End Sub

Private Sub vsgWaitExecAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim vPoint As PointAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset
    
    If Col = col执行人 Then
        With vsgWaitExecAdvice
            strSQL = "Select a.Id, a.编号, a.姓名" & vbNewLine & _
                        "From 人员表 A, 部门人员 B, 人员性质说明 C" & vbNewLine & _
                        "Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = '护士' And b.部门id = [1]" & _
                        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
            
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病区护士", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    mlng病区ID)
                    
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    If rsTmp!姓名 & "" <> .Cell(flexcpData, Row, Col) Then
                        .TextMatrix(Row, Col) = rsTmp!姓名 & ""
                        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                        mblnWaitIsUpdate = True
                    End If
                Else
                    MsgBox "当前病区下没有可选的护士。", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Sub vsgWaitExecAdvice_DblClick()
    With vsgWaitExecAdvice
        If .MouseCol = col选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsgWaitExecAdvice_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsgWaitExecAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsgWaitExecAdvice
        lngLeft = col选择: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col本次数次: lngRight = col备注
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowIn一并给药(Row, lngBegin, lngEnd, vsgWaitExecAdvice) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If .TextMatrix(Row, col相关ID) = "" Then
            vRect.Top = Bottom - 1 '相关ID为空的行文字保留
            vRect.Bottom = Bottom - 1
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Between(Col, col选择, col选择) Or Between(Col, col执行时间, col备注) And mbytUseType <> T医嘱核对 Then
                SetBkColor hDC, OS.SysColor2RGB(COLEditBackColor)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsgWaitExecAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsgWaitExecAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsgWaitExecAdvice_KeyPress(KeyAscii As Integer)
    With vsgWaitExecAdvice
        If .Col = col备注 Or .Col = col执行人 Then
            If KeyAscii = Asc("*") And .Col = col执行人 Then
                KeyAscii = 0
                Call vsgWaitExecAdvice_CellButtonClick(.Row, .Col)
                Exit Sub
            End If
            .ComboList = "" '使按钮状态进入输入状态
        ElseIf .Col = col选择 And KeyAscii = vbKeySpace Then
            Call ExecCheck(vsgWaitExecAdvice)
        End If
    End With
End Sub

Private Sub ExecCheck(ByRef objVsg As VSFlexGrid)
'功能：同步选择一组医嘱
'参数：表格
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    
    With objVsg
        If .TextMatrix(.Row, col医嘱ID) = "" Then Exit Sub
        If Not RowIn一并给药(.Row, lngBegin, lngEnd, objVsg) Then
            lngBegin = .Row: lngEnd = .Row
        End If
        
        For i = lngBegin To lngEnd
            If .Cell(flexcpData, i, col选择) = 1 Then
                Set .Cell(flexcpPicture, i, col选择) = Nothing
                .Cell(flexcpData, i, col选择) = 0
            Else
                If objVsg.Name = "vsgExecAdvice" Then
                    '检查是否出院
                    If .TextMatrix(i, col在院) <> "1" Then
                        MsgBox "该病人已经出院，不能取消执行。", vbInformation, Me.Caption
                        Exit Sub
                    End If
                    '检查是否是当前操作员登记的
                    If .TextMatrix(i, col登记人) <> UserInfo.姓名 Then
                        MsgBox "只能取消执行自己登记的医嘱。", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End If
                .Cell(flexcpPicture, i, col选择) = img16.ListImages("Check").Picture
                .Cell(flexcpData, i, col选择) = 1
                .Cell(flexcpPictureAlignment, i, col选择) = flexPicAlignCenterCenter
            End If
        Next
    End With
End Sub

Private Sub vsgWaitExecAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vPoint As PointAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset
    
    With vsgWaitExecAdvice
        If Col = col执行时间 Then
            If Not IsDate(.EditText) Then
                MsgBox "请输入正确的日期。", vbInformation, Me.Caption
                .EditText = .Cell(flexcpData, Row, Col)
                Cancel = True
                Exit Sub
            End If
        ElseIf Col = col备注 Then
            If zlCommFun.ActualLen(.EditText) > 200 Then
                MsgBox "备注文字不能超出100个汉字或200个字母。", vbInformation, Me.Caption
                .EditText = .Cell(flexcpData, Row, Col)
                Cancel = True
                Exit Sub
            End If
        ElseIf Col = col执行结果 Then
            If .Cell(flexcpData, Row, Col) = "未执行" And .EditText = "完成" Then
                .TextMatrix(Row, COL未执行原因) = ""
                .Cell(flexcpData, Row, COL未执行原因) = ""
            End If
        ElseIf Col = col执行人 Then
            strSQL = "Select a.Id, a.编号, a.姓名" & vbNewLine & _
                        " From 人员表 A, 部门人员 B, 人员性质说明 C" & vbNewLine & _
                        " Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = '护士'" & _
                        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " And b.部门id = [1] And (a.编号=[2] or a.姓名 Like [3] or a.简码 Like [4])"
            
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病区护士", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    mlng病区ID, .EditText, gstrLike & .EditText & "%", gstrLike & UCase(.EditText) & "%")
                    
            If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                Cancel = True
            Else
                If Not rsTmp Is Nothing Then
                    .EditText = rsTmp!姓名 & ""
                Else
                    MsgBox "没查找到您指定的护士。", vbInformation, Me.Caption
                    .EditText = .Cell(flexcpData, Row, Col)
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
        
        If .EditText <> .Cell(flexcpData, Row, Col) Then
            .TextMatrix(Row, Col) = .EditText
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            mblnWaitIsUpdate = True
        End If
    End With
End Sub

Private Function CheckData(ByRef objVsg As VSFlexGrid) As Boolean
'功能：检查医嘱
    Dim i As Long
    
    With objVsg
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col医嘱ID) <> "" And .RowData(i) = "Begin" And (.Cell(flexcpData, i, col选择) = "1" Or objVsg.Name = "vsgExecAdvice" And objVsg.TextMatrix(i, col是否修改) = "1") Then
                '执行人不能为空
                If .TextMatrix(i, col执行人) = "" Then
                    .Row = i: .Col = col执行人
                    Call ShowMessage(vsgWaitExecAdvice, "执行人不能为空。")
                    Exit Function
                End If
                
                '结果为未执行的，须添原因
                If .TextMatrix(i, col执行结果) = "未执行" And .TextMatrix(i, COL未执行原因) = "" Then
                    .Row = i: .Col = COL未执行原因
                    Call ShowMessage(vsgWaitExecAdvice, "未执行的医嘱必须填写未执行原因。")
                    Exit Function
                End If
            End If
        Next
    End With
    
    CheckData = True
End Function

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long

    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then objTmp.SetFocus
    Me.Refresh
End Function

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.Hwnd
    End If
End Sub

Private Sub vsgExecAdvice_BeforeSort(ByVal Col As Long, Order As Integer)
    Call VSColumnClick(vsgExecAdvice, Col, Order)
End Sub

Private Sub vsgWaitExecAdvice_BeforeSort(ByVal Col As Long, Order As Integer)
    Call VSColumnClick(vsgWaitExecAdvice, Col, Order)
End Sub

Private Sub VSColumnClick(ByRef objVs As Object, ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    Select Case Col
        Case col姓名, COL床位, col性别, col要求时间, col执行时间 '排序的列
        Case Else
            Order = 0
    End Select
  
    If Col = col选择 Then
        With objVs
            If .TextMatrix(1, col医嘱ID) = "" Then Exit Sub
            If .ColData(col选择) = "Check" Then
                .Cell(flexcpPicture, 0, col选择) = img16.ListImages("UnCheck").Picture
                .ColData(col选择) = ""
            Else
                .Cell(flexcpPicture, 0, col选择) = img16.ListImages("AllCheck").Picture
                .ColData(col选择) = "Check"
            End If
            For i = 1 To .Rows - 1
                If .TextMatrix(i, col医嘱ID) = "" Then Exit For
                If .ColData(col选择) = "Check" Then
                    .Cell(flexcpPicture, i, col选择) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, col选择) = 1
                    .Cell(flexcpPictureAlignment, i, col选择) = flexPicAlignCenterCenter
                Else
                    Set .Cell(flexcpPicture, i, col选择) = Nothing
                    .Cell(flexcpData, i, col选择) = 0
                End If
            Next
        End With
    End If
End Sub
