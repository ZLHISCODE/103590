VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceBatExecute 
   Caption         =   "ҽ������ִ�еǼ�"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13260
   Icon            =   "frmAdviceBatExecute.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   13260
   StartUpPosition =   3  '����ȱʡ
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
            Key             =   "ǩ��"
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
            Name            =   "����"
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
               Caption         =   "��ҩ"
               Height          =   255
               Index           =   9
               Left            =   1800
               TabIndex        =   51
               Top             =   1185
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "��Ѫ"
               Height          =   255
               Index           =   7
               Left            =   960
               TabIndex        =   28
               Top             =   1185
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "����ҽ��"
               Height          =   255
               Index           =   6
               Left            =   105
               TabIndex        =   21
               Top             =   1470
               Width           =   1035
            End
            Begin VB.CheckBox chkType 
               Caption         =   "�ɼ�"
               Height          =   255
               Index           =   5
               Left            =   1800
               TabIndex        =   20
               Top             =   900
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "����"
               Height          =   255
               Index           =   4
               Left            =   960
               TabIndex        =   19
               Top             =   900
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "Ƥ��"
               Height          =   255
               Index           =   3
               Left            =   105
               TabIndex        =   18
               Top             =   900
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "�ڷ�"
               Height          =   255
               Index           =   2
               Left            =   1800
               TabIndex        =   17
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "ע��"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   16
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "��Һ"
               Height          =   255
               Index           =   0
               Left            =   105
               TabIndex        =   15
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chk��Ч 
               Caption         =   "����"
               Height          =   180
               Index           =   1
               Left            =   1800
               TabIndex        =   13
               Top             =   0
               Width           =   735
            End
            Begin VB.CheckBox chk��Ч 
               Caption         =   "����"
               Height          =   180
               Index           =   0
               Left            =   960
               TabIndex        =   12
               Top             =   0
               Width           =   735
            End
            Begin VB.CheckBox chkType 
               Caption         =   "������ҩ;��"
               Height          =   255
               Index           =   8
               Left            =   105
               TabIndex        =   1
               Top             =   1185
               Width           =   1395
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "���"
               Height          =   180
               Index           =   5
               Left            =   105
               TabIndex        =   14
               Top             =   330
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "��Ч"
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
            Caption         =   "ִ����"
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   50
            Top             =   1095
            Width           =   540
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   450
            TabIndex        =   9
            Top             =   750
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   450
            TabIndex        =   7
            Top             =   390
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Ҫ��ʱ��"
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
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   32
            Top             =   0
            Width           =   660
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "Ӥ��"
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
         Caption         =   "��ǰ������"
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
               Caption         =   "�ϴ�ִ����"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   43
               Top             =   60
               Width           =   1215
            End
            Begin VB.OptionButton optExecutePeople 
               Caption         =   "ָ����Ա"
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
               Caption         =   "����"
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
               Caption         =   "Ҫ��ʱ��"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   37
               Top             =   60
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optExecuteTime 
               Caption         =   "ָ��ʱ��"
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
            Caption         =   "�����޸�(&M)"
            Height          =   300
            Left            =   7800
            TabIndex        =   34
            Top             =   100
            Width           =   1215
         End
         Begin VB.Label lblInfo 
            Caption         =   "ִ��ʱ��"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   47
            Top             =   170
            Width           =   855
         End
         Begin VB.Label lblInfo 
            Caption         =   "ִ�н��"
            Height          =   255
            Index           =   7
            Left            =   5640
            TabIndex        =   46
            Top             =   170
            Width           =   855
         End
         Begin VB.Label lblInfo 
            Caption         =   "ִ����"
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
            Name            =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
    COL_����ID = 0
    COL_��ҳID = 1
    COL_ѡ�� = 2
    COL_���� = 3
    COL_���� = 4
    COL_�Ա� = 5
    COL_סԺ�� = 6
End Enum

Private Enum AdviceCol
    colѡ�� = 0
    COL��λ = 1
    col���� = 2
    col�Ա� = 3
    colסԺ�� = 4
    colסԺ���� = 5
    col��Ч = 6
    colҽ������ = 7
    col���� = 8
    col���� = 9
    col�������� = 10
    COL��ҩ;�� = 11
    colҪ��ʱ�� = 12
    colִ��ʱ�� = 13
    colִ���� = 14
    col�˶�ʱ�� = 15
    COL�˶��� = 16
    colִ�н�� = 17
    COLδִ��ԭ�� = 18
    col��ע = 19
    colҽ��ID = 20
    col���ID = 21
    col������� = 22
    Col����ID = 23
    COL��ҳID = 24
    COLƵ�� = 25
    col���ͺ� = 26
    col��ִ��ID = 27
    COLҽ��״̬ = 28
    col��Ժ = 29
    col�Ǽ��� = 30
    COLԭʼִ��ʱ�� = 31
    col�Ƿ��޸� = 32
    COL���ִ���� = 33
    col�������� = 34
    colִ�з��� = 35
    col����ID = 36
    colִ��״̬ = 37
    col��ʼִ��ʱ�� = 38
    col���ʱ�� = 39
End Enum

Private Enum ClientType
    chk���� = 0
    chk���� = 1
    
    Type��Һ = 0
    Typeע�� = 1
    Type�ڷ� = 2
    TypeƤ�� = 3
    Type���� = 4
    Type�ɼ� = 5
    Type����ҽ�� = 6
    Type��Ѫ = 7
    Type������ҩ;�� = 8
    Type��ҩ���� = 9
    lbl���� = 0
    lblʱ�䷶Χ = 1
    lbl��Ч = 4
    lbl��� = 5
    lblִ���� = 11
    
    dpk��ʼ���� = 0
    dpk�������� = 1
    
    optҪ��ʱ�� = 0
    optָ��ʱ�� = 1
    
    opt�ϴ�ִ���� = 0
    optָ����Ա = 1
    opt���� = 2
    
    cboִ���˵Ǽ� = 0
    cboִ����ȡ�� = 1
    
End Enum

Private Type FilterCond
    str����IDs As String '����ѡ�������"����ID:��ҳID,..." ���� "*" ��ʾȫѡ/ȫ��ѡ
    datB As Date '��ʼ����
    datE As Date '��ֹ����
    str��Ա As String     '��Ա����
End Type
Private mvarCond As FilterCond

Private mlng����ID As Long
Private mlng����ID As Long
Private mrsDefine As ADODB.Recordset
Private mblnWaitIsUpdate As Boolean
Private mblnExecIsUpdate As Boolean
Private mint���� As Integer '0-ҽ��վ����,1-��ʿվ����
Private mbytUseType As Integer        '����ģʽ   1-ҽ���˶ԣ�0-ҽ��ִ��
Private mintҽ������Χ As Integer    'ҽ������Χ   0-����ҽ��,1-����ҽ��,2-Ӥ��ҽ��
Private mlngҽ������ID As Long
Private mlngӤ������ID As Long
Private mlngӤ������ID As Long

Private Enum UseType
    Tҽ��ִ�� = 0
    Tҽ���˶� = 1
End Enum

Public Sub ShowMe(ByVal intType As Integer, ByRef frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int���� As Integer, ByVal bytUseType As Byte, Optional ByVal lngҽ������ID As Long, _
    Optional ByVal lngӤ������ID As Long, Optional ByVal lngӤ������ID As Long)
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mint���� = int����
    mbytUseType = bytUseType
    mlngҽ������ID = lngҽ������ID
    mlngӤ������ID = lngӤ������ID
    mlngӤ������ID = lngӤ������ID
    Me.lblInfo(lbl����).Caption = IIF(mint���� = 1, "��ǰ������", IIF(Val(zlDatabase.GetPara("������ʾ��ʽ", glngSys, pסԺҽ��վ)) = 1, "��ǰ������", "��ǰ���ң�")) & Sys.RowValue("���ű�", IIF(mlngӤ������ID <> 0 And (mlngӤ������ID = mlngҽ������ID Or mlngӤ������ID = mlngҽ������ID), lngӤ������ID, lng����ID), "����")
    
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
        If tbcSub.Selected.Tag = "��ִ��ҽ��" Then
            Call LoadAdvice
        Else
            Call LoadAdvice(True)
        End If
    Case conMenu_Edit_Save
        If tbcSub.Selected.Tag = "��ִ��ҽ��" Then
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
'���ܣ������˶�
    Dim strSQL As String
    Dim str�˶��� As String
    Dim i As Long
    Dim arrSQL As Variant
    Dim rsTmp As Recordset
    Dim blnTrans As Boolean
    Dim strMsgNameSame As String
    Dim strMsg As String
    Dim strXML As String
    Dim blnDo As Boolean
    Dim str�˶�ʱ�� As String
    
    On err GoTo errH
    arrSQL = Array()
    str�˶�ʱ�� = zlDatabase.Currentdate '��ȡ��ǰ������ʱ��
    With vsgWaitExecAdvice
        For i = 0 To .Rows - 1
            If .Cell(flexcpData, i, colѡ��) = "1" And .TextMatrix(i, col���ID) = "" Then
                If str�˶��� = "" Then str�˶��� = zlDatabase.UserIdentifyByUser(Me, "�ں˶�ִ�����ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "ִ������Ǽ�", , True)
                If str�˶��� = "" Then Exit Function

                If str�˶��� = .TextMatrix(i, colִ����) & "" Then
                    strMsgNameSame = strMsgNameSame & "," & .TextMatrix(i, colҽ������)
                Else
                    '���ú˶�ǰ��ҽӿ�
                    If Not gobjPlugIn Is Nothing Then
                        On Error Resume Next
                        blnDo = gobjPlugIn.AdvcieBeforToReview(glngSys, IIF(mint���� = 0, pסԺҽ��վ, pסԺ��ʿվ), Val(.TextMatrix(i, Col����ID)), Val(.TextMatrix(i, COL��ҳID)), Val(.TextMatrix(i, colҽ��ID)), Val(.TextMatrix(i, col���ͺ�)), str�˶���, str�˶�ʱ��, .TextMatrix(i, colִ����) & "", strXML)
                        Call zlPlugInErrH(err, "AdvcieBeforToReview")
                        If 0 = err.Number Then '�ӿ�û�г������������жϽӿڵķ���ֵ
                            If blnDo Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_����ҽ���˶�_Insert(" & Val(.TextMatrix(i, colҽ��ID)) & "," & Val(.TextMatrix(i, col���ͺ�)) & ",'" & str�˶��� & "',Null,To_Date('" & str�˶�ʱ�� & "','YYYY-MM-DD HH24:MI:SS'))"
                            End If
                        End If
                        If err.Number <> 0 Then err.Clear
                        On Error GoTo 0
                    Else
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_����ҽ���˶�_Insert(" & Val(.TextMatrix(i, colҽ��ID)) & "," & Val(.TextMatrix(i, col���ͺ�)) & ",'" & str�˶��� & "',Null,To_Date('" & str�˶�ʱ�� & "','YYYY-MM-DD HH24:MI:SS'))"
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
            strMsg = strMsg & "����ҽ��������˺�ִ����Ϊͬһ���ˣ�" & vbCrLf & Mid(strMsgNameSame, 2) & "��" & vbCrLf
        End If
        
        If UBound(arrSQL) < 0 Then
            MsgBox "����ѡ����Ŀδ�˶Գɹ������У�" & vbCrLf & strMsg, vbInformation, "ҽ���˶�"
        Else
            If strMsg <> "" Then
                MsgBox "���˶���" & UBound(arrSQL) + 1 & "����Ŀ�����У�" & vbCrLf & strMsg, vbInformation, "ҽ���˶�"
            End If
        End If
    End With
    '��ʾִ�����
    Call LoadAdvice
    FuncExecAuditBatch = True

    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncExecDelAuditBatch() As Boolean
'���ܣ�����ȡ���˶�
    Dim bln��ѪƤ�� As Boolean
    Dim strSQL As String
    Dim str�˶��� As String
    Dim i As Long
    Dim arrSQL As Variant
    Dim strMsg As String
    Dim rsTmp As Recordset
    Dim blnTrans As Boolean
    Dim blnIsTwo As Boolean   '�ж��Ƿ�������������ϵĺ˶���
    Dim strTmp As String
    Dim bln�˶��� As Boolean
    Dim strMsgNoRecord As String
    Dim strNoExec As String
    Dim datCur As Date
    
    On err GoTo errH
    arrSQL = Array()
    datCur = zlDatabase.Currentdate
    With vsgExecAdvice
        For i = 0 To .Rows - 1
            If .Cell(flexcpData, i, colѡ��) = "1" And .TextMatrix(i, col���ID) = "" Then
                If Val(.TextMatrix(i, colִ��״̬) & "") = 3 Then
                    If strTmp <> "" And strTmp <> .TextMatrix(i, COL�˶���) & "" Then
                        blnIsTwo = True
                    Else
                        strTmp = .TextMatrix(i, COL�˶���) & ""
                    End If
                    If CanUnExec(CDate(.TextMatrix(i, col�˶�ʱ��)), datCur) Then
                        If .TextMatrix(i, COL�˶���) & "" <> UserInfo.���� Then
                            If str�˶��� = "" Then str�˶��� = zlDatabase.UserIdentifyByUser(Me, "��ȡ���˶�ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "ִ������Ǽ�", , True)
                            If str�˶��� = "" Then Exit Function
                            
                            If str�˶��� = .TextMatrix(i, COL�˶���) & "" Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_����ҽ���˶�_Delete(" & Val(.TextMatrix(i, colҽ��ID)) & "," & Val(.TextMatrix(i, col���ͺ�)) & ")"
                            Else
                                bln�˶��� = True
                                str�˶��� = .TextMatrix(i, COL�˶���) & ""
                            End If
                        Else
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_����ҽ���˶�_Delete(" & Val(.TextMatrix(i, colҽ��ID)) & "," & Val(.TextMatrix(i, col���ͺ�)) & ")"
                        End If
                    Else
                        strNoExec = strNoExec & "," & .TextMatrix(i, colҽ������)
                    End If
                Else
                    strMsgNoRecord = strMsgNoRecord & "," & .TextMatrix(i, colҽ������)
                End If
            End If
        Next
    
        If blnIsTwo Then
            MsgBox "����ͬʱȡ������˺˶Ե���Ŀ����ѡ��ͬһ�������˶Ե���Ŀ��", vbInformation, "ȡ���˶�"
            '��ʾִ�����
            Call LoadAdvice(True)
            Exit Function
        End If
        
        If bln�˶��� Then
            MsgBox "ֻ��ȡ���Լ��˶Ե�ҽ������ǰѡ���ҽ���˶�����""" & str�˶��� & """��", vbInformation, "ȡ���˶�"
            '��ʾִ�����
            Call LoadAdvice(True)
            Exit Function
        End If
    End With

    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "ȡ���˶�")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If strMsgNoRecord <> "" Then
        strMsg = strMsg & "����ҽ���Ѿ�ִ����ɣ���ȡ����ɺ����ԣ�" & vbCrLf & Mid(strMsgNoRecord, 2) & "��" & vbCrLf
    End If
    
    If strNoExec <> "" Then
        strMsg = strMsg & "����ҽ���Ѿ�ִ����ɣ���ȡ����ɺ����ԣ�" & vbCrLf & Mid(strNoExec, 2) & "��" & vbCrLf
    End If
    
    If UBound(arrSQL) < 0 Then
        MsgBox "����ѡ����Ŀδȡ���ɹ������У�" & vbCrLf & strMsg, vbInformation, "ȡ���˶�"
    Else
        If strMsg <> "" Then
            MsgBox "��ȡ���˶���" & UBound(arrSQL) + 1 & "����Ŀ,���У�" & vbCrLf & strMsg, vbInformation, "ȡ���˶�"
        End If
    End If
    '��ʾִ�����
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
        If tbcSub.Selected.Tag = "��ִ��ҽ��" Or mbytUseType = Tҽ���˶� Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Case conMenu_Edit_Save
        If mbytUseType = Tҽ���˶� Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Case conMenu_Manage_ThingAudit
        If mbytUseType = Tҽ���˶� And tbcSub.Selected.Tag = "��ִ��ҽ��" Then
            Control.Visible = True
        Else
            Control.Visible = False
        End If
    Case conMenu_Manage_ThingDelAudit
        If mbytUseType = Tҽ���˶� And tbcSub.Selected.Tag <> "��ִ��ҽ��" Then
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

Private Sub chk��Ч_Click(Index As Integer)
    If chk��Ч(chk����).value = 0 And chk��Ч(chk����).value = 0 Then
        chk��Ч(Index).value = 1
    End If
End Sub

Private Sub cmdBatUpdate_Click()
    Call BatUpdate
    mblnWaitIsUpdate = True
End Sub

Private Sub BatUpdate()
'���ܣ������޸�
    Dim i As Long, lngRow As Long
    Dim blnSame As Boolean
    
    With vsgWaitExecAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, colѡ��) = "1" Then
                'ִ��ʱ��
                If optExecuteTime(optҪ��ʱ��).value Then
                    .TextMatrix(i, colִ��ʱ��) = .TextMatrix(i, colҪ��ʱ��)
                    .Cell(flexcpData, i, colִ��ʱ��) = .TextMatrix(i, colִ��ʱ��)
                ElseIf optExecuteTime(optָ��ʱ��).value Then
                    '���ʱ�䳬���˿�ʼִ��ʱ�䣬��ȡҪ��ʱ��
                    If Format(dpkExecuteTime.value, "yyyy-MM-dd HH:mm") >= Format(.TextMatrix(i, col��ʼִ��ʱ��), "yyyy-MM-dd HH:mm") Then
                         '�����ͬһ��ҽ��ͬ�η��͵ģ�һ��ִ��ʱ������ΪҪ��ʱ��
                        For lngRow = 1 To .Rows - 1
                            If .Cell(flexcpData, lngRow, colѡ��) = "1" Then
                                If .TextMatrix(lngRow, colҪ��ʱ��) <> .TextMatrix(i, colҪ��ʱ��) And .TextMatrix(lngRow, colҽ��ID) = .TextMatrix(i, colҽ��ID) And .TextMatrix(lngRow, col���ͺ�) = .TextMatrix(i, col���ͺ�) Then
                                    blnSame = True
                                    Exit For
                                End If
                            End If
                        Next
                        If blnSame Then
                            .TextMatrix(i, colִ��ʱ��) = .TextMatrix(i, colҪ��ʱ��)
                        Else
                            .TextMatrix(i, colִ��ʱ��) = Format(dpkExecuteTime.value, "yyyy-MM-dd HH:mm")
                        End If
                        blnSame = False
                    Else
                       .TextMatrix(i, colִ��ʱ��) = .TextMatrix(i, colҪ��ʱ��)
                    End If
                    .Cell(flexcpData, i, colִ��ʱ��) = .TextMatrix(i, colִ��ʱ��)
                End If
                'ִ����
                If optExecutePeople(opt�ϴ�ִ����).value Then
                    .TextMatrix(i, colִ����) = .TextMatrix(i, COL���ִ����)
                    .Cell(flexcpData, i, colִ����) = .TextMatrix(i, colִ����)
                ElseIf optExecutePeople(optָ����Ա).value Then
                    .TextMatrix(i, colִ����) = cboExecutePoeple(0).Text
                    .Cell(flexcpData, i, colִ����) = .TextMatrix(i, colִ����)
                ElseIf optExecutePeople(opt����).value Then
                    .TextMatrix(i, colִ����) = UserInfo.����
                    .Cell(flexcpData, i, colִ����) = .TextMatrix(i, colִ����)
                End If
                'ִ�н��
                If .TextMatrix(i, colִ�н��) = "δִ��" And cboExecuteResult.Text = "���" Then
                    .TextMatrix(i, COLδִ��ԭ��) = ""
                    .Cell(flexcpData, i, COLδִ��ԭ��) = .TextMatrix(i, COLδִ��ԭ��)
                End If
                .TextMatrix(i, colִ�н��) = cboExecuteResult.Text
                .Cell(flexcpData, i, colִ�н��) = .TextMatrix(i, colִ�н��)
                
            End If
        Next
    End With
End Sub

Private Sub FuncCancleExec()
'���ܣ�ȡ��ִ��
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
            If .TextMatrix(i, colҽ��ID) <> "" And .RowData(i) = "Begin" And .Cell(flexcpData, i, colѡ��) = "1" Then
                If CanUnExec(CDate(.TextMatrix(i, col���ʱ��)), datCur) Then
                    '����Ƿ�˶�
                    If .TextMatrix(i, col�˶�ʱ��) <> "" Then
                        MsgBox "ҽ����" & .TextMatrix(i, colҽ������) & " �Ѿ��˶ԣ���ȡ���˶Ժ����ԡ�", vbInformation, "ȡ��ִ��"
                        Exit Sub
                    Else
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_����ҽ��ִ��_Delete(" & .TextMatrix(i, col��ִ��ID) & "," & .TextMatrix(i, col���ͺ�) & ",To_date('" & .TextMatrix(i, COLԭʼִ��ʱ��) & "','YYYY-MM-DD HH24:MI:ss'),0,1," & mlng����ID & ")"
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
    If vsgExecAdvice.TextMatrix(1, colҽ��ID) = "" Then
        stbThis.Panels(2).Text = "û�п�ȡ��ִ�е�ҽ����"
    Else
        If UBound(arrSQL) = -1 Then
            MsgBox "�빴ѡ����Ҫȡ��ִ�е�ҽ����", vbInformation, Me.Caption
            Exit Sub
        End If
        stbThis.Panels(2).Text = "ȡ���ɹ������ι�ȡ��ִ���� " & UBound(arrSQL) + 1 & " ��ҽ����"
    End If
    'ɾ��ȡ����ҽ��
    For i = UBound(arrAdvcie) To 0 Step -1
        If vsgExecAdvice.TextMatrix(Val(arrAdvcie(i)), colҽ��ID) <> "" Then
            If Not RowInһ����ҩ(Val(arrAdvcie(i)), lngBegin, lngEnd, vsgExecAdvice) Then
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
'���ܣ�����ִ�д�ִ�е�ҽ��
    Dim arrSQL() As Variant
    Dim i As Long, j As Long, s As Long
    Dim blnTrans As Boolean
    
    If Not CheckData(vsgWaitExecAdvice) Then Exit Sub
    arrSQL = Array()
    With vsgWaitExecAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colҽ��ID) <> "" And .RowData(i) = "Begin" And .Cell(flexcpData, i, colѡ��) = "1" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����ҽ��ִ��_Insert(" & .TextMatrix(i, col��ִ��ID) & "," & .TextMatrix(i, col���ͺ�) & ",To_date('" & .TextMatrix(i, colҪ��ʱ��) & "','YYYY-MM-DD HH24:MI')," & _
                IIF(Val(.TextMatrix(i, col��������)) = 0, 1, Val(.TextMatrix(i, col��������))) & ",'" & .TextMatrix(i, col��ע) & "','" & .TextMatrix(i, colִ����) & "',To_date('" & .TextMatrix(i, colִ��ʱ��) & _
                "','YYYY-MM-DD HH24:MI'),0,1," & IIF(.TextMatrix(i, colִ�н��) = "δִ��", "0", "1") & ",'" & .TextMatrix(i, COLδִ��ԭ��) & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlng����ID & ")"
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
    If vsgWaitExecAdvice.TextMatrix(1, colҽ��ID) = "" Then
        stbThis.Panels(2).Text = "û�п�ִ�е�ҽ����"
    Else
        stbThis.Panels(2).Text = "����ɹ������ι�ִ���� " & UBound(arrSQL) + 1 & " ��ҽ����"
    End If
    '���ҽ���б�
    vsgWaitExecAdvice.Rows = 1
    vsgWaitExecAdvice.AddItem ""
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    For i = 1 To vsgWaitExecAdvice.Rows - 1
        If vsgWaitExecAdvice.TextMatrix(i, colҽ��ID) <> "" And vsgWaitExecAdvice.RowData(i) = "Begin" And vsgWaitExecAdvice.Cell(flexcpData, i, colѡ��) = "1" Then
            s = s + 1
            If s = j Then
                vsgWaitExecAdvice.Row = i: vsgWaitExecAdvice.ShowCell i, COL_����
                Me.Refresh
                Exit For
            End If
        End If
    Next
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncSaveUpdate()
'���ܣ��������޸ĵ�ҽ��
    Dim arrSQL() As Variant
    Dim i As Long
    Dim blnTrans As Boolean
    Dim datCur As Date
    
    If Not CheckData(vsgExecAdvice) Then Exit Sub
    arrSQL = Array()
    datCur = zlDatabase.Currentdate
    With vsgExecAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colҽ��ID) <> "" And .RowData(i) = "Begin" And .TextMatrix(i, col�Ƿ��޸�) = "1" Then
                If CanUnExec(CDate(.TextMatrix(i, col���ʱ��)), datCur) Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_����ҽ��ִ��_Update(" & "To_date('" & .TextMatrix(i, COLԭʼִ��ʱ��) & "','YYYY-MM-DD HH24:MI')," & .TextMatrix(i, col��ִ��ID) & "," & .TextMatrix(i, col���ͺ�) & ",To_date('" & .TextMatrix(i, colҪ��ʱ��) & "','YYYY-MM-DD HH24:MI')," & _
                        IIF(Val(.TextMatrix(i, col��������)) = 0, 1, Val(.TextMatrix(i, col��������))) & ",'" & .TextMatrix(i, col��ע) & "','" & .TextMatrix(i, colִ����) & "',To_date('" & .TextMatrix(i, colִ��ʱ��) & _
                        "','YYYY-MM-DD HH24:MI')," & IIF(.TextMatrix(i, colִ�н��) = "δִ��", "0", "1") & ",'" & .TextMatrix(i, COLδִ��ԭ��) & "',0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlng����ID & ")"
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
    If vsgExecAdvice.TextMatrix(1, colҽ��ID) = "" Then
        stbThis.Panels(2).Text = "û�п��޸ĵ���ִ��ҽ����"
    Else
        stbThis.Panels(2).Text = "�޸ĳɹ������ι��޸��� " & UBound(arrSQL) + 1 & " ��ҽ����"
    End If
    '�ָ�ǰ����ɫ
    vsgExecAdvice.Cell(flexcpForeColor, 1, colѡ��, vsgExecAdvice.Rows - 1, colѡ��) = vbBlack
    vsgExecAdvice.Cell(flexcpForeColor, 1, colִ��ʱ��, vsgExecAdvice.Rows - 1, col��ע) = vbBlack
    '�ָ��޸ĵı�ʶ��ԭִ��ʱ��
    With vsgExecAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colҽ��ID) <> "" And .RowData(i) = "Begin" And .TextMatrix(i, col�Ƿ��޸�) = "1" Then
                If CanUnExec(CDate(.TextMatrix(i, col���ʱ��)), datCur) Then
                    .TextMatrix(i, col�Ƿ��޸�) = ""
                    .TextMatrix(i, COLԭʼִ��ʱ��) = .TextMatrix(i, colִ��ʱ��)
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
        If tbcSub.Selected.Tag = "��ִ��ҽ��" Then
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
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        If mbytUseType = Tҽ���˶� Then
            strTbc = "�˶�"
        Else
            strTbc = "ִ��"
        End If
        .InsertItem(0, "��" & strTbc & "ҽ��(&1)", picWaitExecute.Hwnd, 0).Tag = "��ִ��ҽ��"
        .InsertItem(1, "��" & strTbc & "ҽ��(&2)", picExecuted.Hwnd, 0).Tag = "��ִ��ҽ��"
        
        .Item(0).Selected = True
    End With
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 180, 400, DockLeftOf, Nothing)
    objPane.Title = "�����б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    'VSFlexGrid
    '-----------------------------------------------------
    strHead = ",400,1;��λ,750,1;����,850,1;�Ա�,450,1;סԺ��,1200,1;סԺ����,950,1;��Ч,450,1;ҽ������,2500,1;����;����,700,1;��������,840,7;��ҩ;��;Ҫ��ʱ��,1800,1;ִ��ʱ��,1800,1;ִ����,950,1;�˶�ʱ��,1550,1;�˶���,950,1;ִ�н��,950,1;δִ��ԭ��,1000,1;��ע,2000,1;ҽ��ID;���ID;�������;����ID;��ҳID;Ƶ��;����;���ͺ�;��ִ��ID;ҽ��״̬;��Ժ;�Ǽ���;ԭʼִ��ʱ��;�Ƿ��޸�;���ִ����;��������;ִ�з���;����ID;ִ��״̬;��ʼִ��ʱ��;���ʱ��"
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
    
    'ҽ���˶�
    If mbytUseType = Tҽ���˶� Then
        Me.Caption = "ҽ�������˶�"
        lblInfo(lblʱ�䷶Χ).Caption = "ִ��ʱ��"
        picExec.Visible = False
        lblInfo(11).Caption = "�˶���"
        lblInfo(lbl��Ч).Visible = False
        chk��Ч(chk����).Visible = False
        chk��Ч(chk����).Visible = False
        chkType(Type��Һ).Visible = False
        chkType(Typeע��).Visible = False
        chkType(Type�ڷ�).Visible = False
        chkType(Type�ɼ�).Visible = False
        chkType(Type����ҽ��).Visible = False
        chkType(Type������ҩ;��).Visible = False
        chkType(Type����).Visible = False
        chkType(Type��ҩ����).Visible = False
        chkType(Type��Ѫ).Top = lblInfo(lbl���).Top
        chkType(TypeƤ��).Top = chkType(Type��Ѫ).Top
        picFitter.Height = chkType(TypeƤ��).Top + 350
        lblInfo(lbl���).Top = lblInfo(lbl��Ч).Top
    Else
        chkType(Type��Ѫ).Visible = False
    End If
    
    If DeptIsWoman(0, Get����IDs(mlng����ID)) Then
        fraBaby.Visible = True
        'ҽ������Χ
        mintҽ������Χ = Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0"))
        optBaby(mintҽ������Χ).value = True
    End If
    If Not (mbytUseType = Tҽ���˶� Or Val(zlDatabase.GetPara(51)) = 1) Then
        optExecutePeople(opt����).value = True
        cboExecutePoeple(cboִ����ȡ��).Visible = False
        lblInfo(lblִ����).Visible = False
    End If
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, " ��ѯ(&Q)"): objControl.BeginGroup = True
        objControl.ToolTipText = "��ȡ��ִ��/��ִ�е�����"

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " ����(&S)")
        objControl.BeginGroup = True
        objControl.ToolTipText = "���Ѿ���ѡ��ҽ������ִ��/���Ѿ��޸ĵ����ݽ��б��档"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ��ִ��(&C)")
        objControl.ToolTipText = "���Ѿ���ѡ��ҽ������ȡ��ִ�еĲ�����"
        objControl.IconId = 3651
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "�˶�(&A)")
        objControl.ToolTipText = "���Ѿ���ѡ��Ƥ�Ի���Ѫҽ�����к˶ԵĲ�����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit, "ȡ���˶�(&C)")
        objControl.ToolTipText = "���Ѿ���ѡ��Ƥ�Ի���Ѫҽ������ȡ���˶ԵĲ�����"
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " �˳�(&E)"): objControl.BeginGroup = True
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
'���ܣ�����ҽ��
'������blnIsExecute=true������ִ��/�Ѻ˶�ҽ��,falseΪ���ش�ִ��/���˶�ҽ��
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim lngID As Long       '���ڶ�λ
    Dim strFormat As String
    Dim strTmp As String
    Dim strFitter As String
    Dim strPatis As Variant
    Dim blnDo As Boolean
    Dim lngCount As Long   '��Ҫִ�е�ҽ����(һ����ҩ��1��)
    Dim strExecPeople As String
    Dim strFace As String  '��ǰˢ�½��� 00����ִ�У�01�����˶ԣ�10����ִ�У�11���Ѻ˶ԡ�
    Dim lngCntH As Long, lngCntN As Long
    Dim blnALL As Boolean
    Dim strWhere As String
    Dim str���Ʊ����� As String
    
    On Error GoTo errH
    '����
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record.Tag = "1" Then
                strPatis = strPatis & "," & rptPati.Rows(i).Record(COL_����ID).value & ":" & rptPati.Rows(i).Record(COL_��ҳID).value
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
        MsgBox "��ѡ����Ҫ��ѯ�Ĳ��ˡ�", vbInformation, Me.Caption
        Exit Sub
    End If
    
    strFace = IIF(blnIsExecute, "1", "0") & IIF(mbytUseType <> Tҽ���˶�, "0", "1")

    If strFace = "00" Or strFace = "10" Then
        '�Ǽǹ��ܣ�ִ�к�ȡ��
        If chk��Ч(chk����).value = 1 And chk��Ч(chk����).value <> 1 Then
            strWhere = " And a.ҽ����Ч=0 "
        ElseIf chk��Ч(chk����).value <> 1 And chk��Ч(chk����).value = 1 Then
            strWhere = " And a.ҽ����Ч=1 "
        End If
        '�������
        strFitter = IIF(chkType(Type��Һ).value, " a.�������='E' And c.��������='2' and c.ִ�з���=1", "")
        strFitter = strFitter & IIF(chkType(Typeע��).value, IIF(strFitter <> "", " Or ", "") & " a.�������='E' And c.��������='2' and c.ִ�з���=2", "")
        strFitter = strFitter & IIF(chkType(Type�ڷ�).value, IIF(strFitter <> "", " Or ", "") & " a.�������='E' And c.��������='2' and c.ִ�з���=4", "")
        strFitter = strFitter & IIF(chkType(Type������ҩ;��).value, IIF(strFitter <> "", " Or ", "") & " a.�������='E' And c.��������='2' and c.ִ�з���=0", "")
        strFitter = strFitter & IIF(chkType(Type����).value, IIF(strFitter <> "", " Or ", "") & " a.�������='E' And c.�������� in('0','5')", "")
        strFitter = strFitter & IIF(chkType(Type�ɼ�).value, IIF(strFitter <> "", " Or ", "") & " a.�������='E' And c.��������='6' ", "")
        strFitter = strFitter & IIF(chkType(Type��ҩ����).value, IIF(strFitter <> "", " Or ", "") & " a.�������='E' And c.��������='4' ", "")
        '����
        strFitter = strFitter & IIF(chkType(Type����ҽ��).value, IIF(strFitter <> "", " Or ", "") & " a.�������='4' " & _
                    " Or a.�������='H' And c.��������='0'" & _
                    " Or a.�������='D'" & _
                    " Or a.�������='I'" & _
                    " Or a.�������='L'" & _
                    " Or a.�������='Z' And c.��������='0'" & _
                    " Or a.�������='K'" _
                    , "")
        strFitter = strFitter & IIF(chkType(TypeƤ��).value, IIF(strFitter <> "", " Or ", "") & " a.�������='E' And c.��������='1' and c.ִ�з���=3", "")
    Else
        '�˶Թ��ܣ�ִ�к�ȡ��   ��Ч,�˶�������Ч  �������-�˶�ֻ������Ѫ��Ƥ��
        strFitter = strFitter & IIF(chkType(Type��Ѫ).value, IIF(strFitter <> "", " Or ", "") & " a.�������='K' ", "")
        If chkType(TypeƤ��).value = 0 And chkType(Type��Ѫ).value = 0 Then
            strFitter = strFitter & IIF(strFitter <> "", " Or ", "") & " a.�������='K' Or a.�������='E' And c.��������='1'"
        End If
        strFitter = strFitter & IIF(chkType(TypeƤ��).value, IIF(strFitter <> "", " Or ", "") & " a.�������='E' And c.��������='1' and c.ִ�з���=3", "")
    End If
    
    strWhere = strWhere & IIF(strFitter <> "", " And (" & strFitter & ")", "")
    If blnALL Then
        '�����ȫѡ����ʹ�ò���ID����
        strWhere = strWhere & " And (a.����ID,A.��ҳID) In (Select h.����id,h.��ҳid From ������ҳ H,��Ժ���� R Where h.����id=r.����id " & _
            IIF(mint���� = 1, " And h.��ǰ����ID+0=R.����ID  And (R.����ID=[5] Or h.Ӥ������ID=[5])", "  And h.��Ժ����id+0=R.����ID And (R.����ID=[5] Or h.Ӥ������ID=[5])") & ")"
    Else
        strWhere = strWhere & " And (a.����ID,A.��ҳID) In(Select /*+cardinality(x,10)*/  x.C1 As ����ID,x.C2 As ��ҳID From Table(f_Num2list2([6])) x )"
    End If
    
    str���Ʊ����� = " Exists (Select 1 From ������ҳ H Where a.����id=h.����id And a.��ҳid=h.��ҳid And (h.��Ժ����id=g.ִ�в���id Or h.��ǰ����id=g.ִ�в���id))"
    
    Select Case strFace
    Case "00"
        If mblnWaitIsUpdate Then
            If MsgBox("��ִ��ҽ��������δ���棬�Ƿ�Ҫˢ�£�", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                vsgWaitExecAdvice.Redraw = flexRDDirect
                Exit Sub
            End If
        End If
        strSQL = "Select b.Ҫ��ʱ��, b.ҽ��id, b.���ͺ�,g.ִ��״̬,Decode(A.ִ��Ƶ��,'һ����',NVL(A.�ܸ�����,1),'��Ҫʱ',NVL(A.�ܸ�����,1),'������'," & _
            "NVL(A.��������,1),'��Ҫʱ',NVL(A.��������,1),'����ʱ',NVL(A.��������,1),Decode(A.ҽ����Ч,0,NVL(A.��������,1),1,Decode(A.�ϴ�ִ��ʱ��,b.Ҫ��ʱ��,Decode(mod(NVL(A.�ܸ�����,1),NVL(A.��������,1)),0,NVL(A.��������,1),mod(NVL(A.�ܸ�����,1),NVL(A.��������,1))),NVL(A.��������,1)),1)) as ��������" & _
            " From ҽ��ִ��ʱ�� B,����ҽ����¼ A, ������ĿĿ¼ C,����ҽ������ G" & vbNewLine & _
            " Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And g.���ͺ�=b.���ͺ� And b.ҽ��ID=g.ҽ��ID And (a.ҽ����Ч=1 or (a.ҽ����Ч=0 and (b.Ҫ��ʱ��<=a.ִ����ֹʱ�� or a.ִ����ֹʱ�� is null)))" & _
            " And Not Exists (Select 1 From ����ҽ��ִ�� Where b.Ҫ��ʱ�� = Ҫ��ʱ�� And b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ�) " & _
            " And (g.ִ�в���id=[4] Or Exists (Select 1 From ����ҽ����¼ D Where a.Id = d.���id And d.ִ�п���id = [4]) Or " & str���Ʊ����� & " or A.ִ������=5 and a.ִ�б��=0) And b.Ҫ��ʱ��+0 Between [1] And [2]"
    Case "01"
        '�˶Թ���ֻ�л�ʿվ���У����˵�ǰ�����Ͳ�����Ӧ�Ŀ��ҵ�ҽ����
        strSQL = "Select b.Ҫ��ʱ��, b.ҽ��id, b.���ͺ�, b.ִ����, b.ִ��ʱ��, b.ִ�н��, b.˵��, b.ִ��ժҪ, b.�Ǽ���,b.�˶���,b.�˶�ʱ��,g.ִ��״̬,B.��������,Decode(g.ִ��״̬,1,g.���ʱ��,2,g.���ʱ��, B.�Ǽ�ʱ��) ���ʱ�� " & _
            " From ����ҽ��ִ�� B, ����ҽ����¼ A, ������ĿĿ¼ C,����ҽ������ G " & _
            " Where a.Id = b.ҽ��id And b.ҽ��ID=g.ҽ��ID And g.���ͺ�=b.���ͺ� And a.������Ŀid = c.Id " & _
            " And (g.ִ�в���id=[4] Or Exists (Select 1 From ����ҽ����¼ D Where a.Id = d.���id And d.ִ�п���id = [4]) Or " & str���Ʊ����� & ") And b.�˶�ʱ�� Is Null And b.ִ��ʱ�� Between [1] And [2]"
    Case "10"
         '����Ѿ����޸���,����ʾ�Ƿ����
        If mblnExecIsUpdate Then
            If MsgBox("��ִ��ҽ��������δ���棬�Ƿ�Ҫˢ�£�", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                vsgExecAdvice.Redraw = flexRDDirect
                Exit Sub
            End If
        End If
        strSQL = "Select b.ҽ��id, b.���ͺ�, b.Ҫ��ʱ��,B.��������, b.ִ����, b.ִ��ʱ��, b.ִ�н��, b.˵��, b.ִ��ժҪ, b.�Ǽ���,b.�˶���,b.�˶�ʱ�� ,g.ִ��״̬,Decode(g.ִ��״̬,1,g.���ʱ��,2,g.���ʱ��, B.�Ǽ�ʱ��) ���ʱ�� " & vbNewLine & _
            " From ����ҽ��ִ�� B,ҽ��ִ��ʱ�� E,����ҽ������ G,����ҽ����¼ A,������ĿĿ¼ C" & vbNewLine & _
            " Where e.Ҫ��ʱ��=b.Ҫ��ʱ�� and e.ҽ��id=b.ҽ��id and e.���ͺ�=b.���ͺ� And g.���ͺ�=b.���ͺ� And b.ҽ��ID=g.ҽ��ID And b.ҽ��ID=a.ID and a.������ĿID=c.ID" & _
            " and (g.ִ�в���id=[4] Or Exists (Select 1 From ����ҽ����¼ D Where a.Id = d.���id And d.ִ�п���id = [4]) Or " & str���Ʊ����� & " or a.ִ������=5 and a.ִ�б��=0) And e.Ҫ��ʱ�� Between [1] And [2] "
        If cboExecutePoeple(1).ListCount >= 0 Then
            strSQL = strSQL & IIF(cboExecutePoeple(1).Text = "", "", " And B.ִ����=[3] ")
            strExecPeople = cboExecutePoeple(1).Text
        End If
    Case "11"
        'ִ�еǼǹ���ҽ���ͻ�ʿ����Ҫ���ֳ��ϣ�ҽ��վ����ʱֻ����ǰ��Ŀ��ҷ�Χ�ڵ�ҽ����
        strSQL = "Select b.ҽ��id, b.���ͺ�, b.Ҫ��ʱ��,B.��������, b.ִ����, b.ִ��ʱ��, b.ִ�н��, b.˵��, b.ִ��ժҪ, b.�Ǽ���,b.�˶���,b.�˶�ʱ�� ,g.ִ��״̬,Decode(g.ִ��״̬,1,g.���ʱ��,2,g.���ʱ��, B.�Ǽ�ʱ��) ���ʱ�� " & vbNewLine & _
            " From ����ҽ��ִ�� B,ҽ��ִ��ʱ�� E,����ҽ������ G,����ҽ����¼ A,������ĿĿ¼ C" & vbNewLine & _
            " Where e.Ҫ��ʱ��=b.Ҫ��ʱ�� and e.ҽ��id=b.ҽ��id and e.���ͺ�=b.���ͺ� And g.���ͺ�=b.���ͺ� And b.ҽ��ID=g.ҽ��ID And b.ҽ��ID=a.ID and a.������ĿID=c.ID and (" & _
            " g.ִ�в���id=[4] Or Exists (Select 1 From ����ҽ����¼ D Where a.Id = d.���id And d.ִ�п���id = [4]) Or " & str���Ʊ����� & ") And b.�˶�ʱ�� Is Not Null And b.ִ��ʱ��  Between [1] And [2] "
         '���û�д���ִ�е�Ȩ�ޣ�Ҳ������ȡ�����˵�ִ�м�¼
        If Val(zlDatabase.GetPara(51)) = 0 Then
            strSQL = strSQL & " And B.ִ����=[3] "
            strExecPeople = UserInfo.����
        Else
            If cboExecutePoeple(1).ListCount >= 0 Then
                strSQL = strSQL & IIF(cboExecutePoeple(1).Text = "", "", " And B.�˶���=[3] ")
                strExecPeople = cboExecutePoeple(1).Text
            End If
        End If
    End Select
    strSQL = strSQL & strWhere
    If blnIsExecute Or mbytUseType = Tҽ���˶� Then
        strSQL = "Select b.ִ����, to_char(b.ִ��ʱ��,'YYYY-MM-DD HH24:MI') as ִ��ʱ��, b.ִ�н��, b.˵�� as δִ��ԭ��, b.ִ��ժҪ as ��ע,b.�Ǽ���,b.ִ��ʱ�� as ԭʼִ��ʱ��,b.�˶���,to_char(b.�˶�ʱ��,'YYYY-MM-DD HH24:MI') as �˶�ʱ��," & _
            " a.Id, b.���ͺ�,b.ҽ��id as ��ִ��ID,a.���id, a.�������,b.ִ��״̬,a.��ʼִ��ʱ��, A.����, f.��Ժ���� As ����, A.�Ա�, Decode(Nvl(a.ҽ����Ч, 0), 0, '����', '����') As ��Ч,a.ҽ��״̬,Decode(f.��Ժ����,NULL,1,0) as ��Ժ,a.�ܸ�����,a.��������," & vbNewLine & _
            " Decode(a.��������, Null, Null, decode(sign(1-A.��������),1,'0'||A.��������,A.��������) || c.���㵥λ) As ����,  Decode(a.���id,Null,a.ҽ������ || ' ' || a.ִ��Ƶ��  ,a.ҽ������) as ҽ������, to_char(b.Ҫ��ʱ��,'YYYY-MM-DD HH24:MI') as Ҫ��ʱ��,f.��ǰ����ID,B.��������," & _
            " a.ִ��Ƶ�� As Ƶ��, a.����id, a.��ҳid, a.������Ŀid,c.��������,c.ִ�з���,Decode(a.�ܸ�����, Null, Null," & _
            " Round(a.�ܸ����� / Decode(a.������Դ, 2, d.סԺ��װ, d.�����װ), 5) || Decode(a.������Դ, 2, d.סԺ��λ, d.���ﵥλ)) As ����,B.���ʱ��,F.סԺ��,G.סԺ����,a.ҽ������" & vbNewLine & _
            " From (" & strSQL & ") B, ����ҽ����¼ A,������ҳ F, ������ĿĿ¼ C, ҩƷ��� D,������Ϣ G" & vbNewLine & _
            " Where (a.Id = b.ҽ��id " & IIF(mbytUseType = Tҽ���˶�, "", "Or a.���id = b.ҽ��id") & ") And f.����id = a.����id And f.��ҳid = a.��ҳid And F.����ID=G.����ID And a.������Ŀid = c.Id And a.�շ�ϸĿid = d.ҩƷid(+) And a.������� Not In('C','7') And Not (a.�������='E' And c.��������='3') " & _
            " " & decode(mintҽ������Χ, 1, " And nvl(a.Ӥ��,0) = 0 ", 2, " And nvl(a.Ӥ��,0) <> 0 ", "") & _
            " And (F.Ӥ������ID is null or F.Ӥ������ID is not null and (F.Ӥ������ID=[5] or F.Ӥ������ID=[5]) and NVL(A.Ӥ��,0)<>0 or F.Ӥ������ID is not null and (F.Ӥ������ID<>[5] and f.Ӥ������ID<>[5]) and NVL(A.Ӥ��,0)=0) "
    Else
        strSQL = "Select distinct" & _
            " a.Id, b.���ͺ�,b.ҽ��id as ��ִ��ID,a.���id, a.�������,b.ִ��״̬,a.��ʼִ��ʱ��, A.����, f.��Ժ���� As ����, A.�Ա�, Decode(Nvl(a.ҽ����Ч, 0), 0, '����', '����') As ��Ч,a.ҽ��״̬,Decode(f.��Ժ����,NULL,1,0) as ��Ժ,a.�ܸ�����,a.��������," & vbNewLine & _
            " Decode(a.��������, Null, Null, decode(sign(1-A.��������),1,'0'||A.��������,A.��������) || c.���㵥λ) As ����,  Decode(a.���id,Null,a.ҽ������ || ' ' || a.ִ��Ƶ��  ,a.ҽ������) as ҽ������, to_char(b.Ҫ��ʱ��,'YYYY-MM-DD HH24:MI') as Ҫ��ʱ��,f.��ǰ����ID,B.��������," & _
            " a.ִ��Ƶ�� As Ƶ��, a.����id, a.��ҳid, a.������Ŀid,c.��������,c.ִ�з���,Decode(a.�ܸ�����, Null, Null," & _
            " Round(a.�ܸ����� / Decode(a.������Դ, 2, d.סԺ��װ, d.�����װ), 5) || Decode(a.������Դ, 2, d.סԺ��λ, d.���ﵥλ)) As ����,first_value(e.ִ����) Over(partition By e.ҽ��id Order By e.ִ��ʱ�� DESC) As ���ִ����,b.Ҫ��ʱ��, a.���,F.סԺ��,G.סԺ����,a.ҽ������" & vbNewLine & _
            " From (" & strSQL & ") B, ����ҽ����¼ A,������ҳ F, ������ĿĿ¼ C, ҩƷ��� D ,����ҽ��ִ�� E,������Ϣ G" & vbNewLine & _
            " Where (a.Id = b.ҽ��id Or a.���id = b.ҽ��id) And e.ҽ��id(+) = a.Id And f.����id = a.����id And f.��ҳid = a.��ҳid And a.������Ŀid = c.Id And F.����ID=G.����ID And a.�շ�ϸĿid = d.ҩƷid(+) And a.������� Not In('C','7') And Not (a.�������='E' And c.��������='3') " & _
            " " & decode(mintҽ������Χ, 1, " And nvl(a.Ӥ��,0) = 0 ", 2, " And nvl(a.Ӥ��,0) <> 0 ", "") & _
            " And (F.Ӥ������ID is null or F.Ӥ������ID is not null and (F.Ӥ������ID=[5] or F.Ӥ������ID=[5]) and NVL(A.Ӥ��,0)<>0 or F.Ӥ������ID is not null and (F.Ӥ������ID<>[5] and f.Ӥ������ID<>[5]) and NVL(A.Ӥ��,0)=0) "
    End If
    strSQL = strSQL & " Order By A.����,b.Ҫ��ʱ��,Nvl(a.���id,a.Id),a.id,a.���"
    
    If strFace = "11" Or strFace = "10" Then
        vsgExecAdvice.Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("UnCheck").Picture
        vsgExecAdvice.Cell(flexcpPictureAlignment, 0, colѡ��) = flexPicAlignCenterCenter
        vsgExecAdvice.ColData(colѡ��) = ""
    Else
        vsgWaitExecAdvice.Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("AllCheck").Picture
        vsgWaitExecAdvice.Cell(flexcpPictureAlignment, 0, colѡ��) = flexPicAlignCenterCenter
        vsgWaitExecAdvice.ColData(colѡ��) = "Check"
        '�����������
        Call zlDatabase.SetPara("ҽ��ִ����Ч", chk��Ч(chk����).value & chk��Ч(chk����).value & "", glngSys, pסԺҽ������)
        strTmp = chkType(Type��Һ).value & chkType(Typeע��).value & chkType(Type�ڷ�).value & chkType(TypeƤ��).value & chkType(Type����).value & _
            chkType(Type�ɼ�).value & chkType(Type����ҽ��).value & chkType(Type��Ѫ).value & chkType(Type������ҩ;��).value & chkType(Type��ҩ����).value
        Call zlDatabase.SetPara("ҽ��ִ�з�Χ", strTmp, glngSys, pסԺҽ������)
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(dpkReqTime(dpk��ʼ����).value), CDate(dpkReqTime(dpk��������).value), strExecPeople, mlng����ID, mlngҽ������ID, strPatis)

    With IIF(blnIsExecute, vsgExecAdvice, vsgWaitExecAdvice)
        .Redraw = flexRDNone
        .Rows = 1
        .ExplorerBar = 7
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                If .ColData(colѡ��) = "Check" Then
                    .Cell(flexcpPicture, i, colѡ��) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, colѡ��) = 1
                    .Cell(flexcpPictureAlignment, i, colѡ��) = flexPicAlignCenterCenter
                End If
                .TextMatrix(i, col����) = rsTmp!���� & ""
                .TextMatrix(i, colסԺ��) = rsTmp!סԺ�� & ""
                .TextMatrix(i, colסԺ����) = rsTmp!סԺ���� & ""
                .TextMatrix(i, col��Ч) = rsTmp!��Ч & ""
                .TextMatrix(i, col����) = rsTmp!���� & ""
                .TextMatrix(i, colҽ��ID) = rsTmp!ID & ""
                .TextMatrix(i, col���ID) = rsTmp!���ID & ""
                .TextMatrix(i, col�Ա�) = rsTmp!�Ա� & ""
                .TextMatrix(i, COL��λ) = rsTmp!���� & ""
                .TextMatrix(i, Col����ID) = rsTmp!����ID & ""
                .TextMatrix(i, COL��ҳID) = rsTmp!��ҳID & ""
                .TextMatrix(i, col�������) = rsTmp!������� & ""
                .TextMatrix(i, col����) = rsTmp!���� & ""
                '��Ѫҽ���ļ���ҽ�����⴦����Ϊ��Ѫҽ��������Ƶ�ʴ������͵�
                If rsTmp!��Ч & "" = "����" And (rsTmp!������� & "" = "K" Or rsTmp!������� & "" = "E" And rsTmp!�������� & "" = "8" And rsTmp!���ID & "" <> "") Then
                   If rsTmp!������� & "" = "K" Then
                        .TextMatrix(i, col��������) = Get��Ѫ��������(Val(rsTmp!ID & ""), Val(rsTmp!���ͺ� & ""), CDate(Format(rsTmp!Ҫ��ʱ�� & "", "yyyy-mm-dd HH:mm:ss")), Val(rsTmp!�ܸ����� & ""), Val(rsTmp!�������� & ""))
                   Else
                        .TextMatrix(i, col��������) = 1
                   End If
                Else
                   .TextMatrix(i, col��������) = rsTmp!�������� & ""
                End If
                .TextMatrix(i, col���ͺ�) = rsTmp!���ͺ� & ""
                .TextMatrix(i, col��ִ��ID) = rsTmp!��ִ��ID & ""
                .TextMatrix(i, COLҽ��״̬) = rsTmp!ҽ��״̬ & ""
                .TextMatrix(i, col��Ժ) = rsTmp!��Ժ & ""
                .TextMatrix(i, col��������) = rsTmp!�������� & ""
                .TextMatrix(i, colִ�з���) = rsTmp!ִ�з��� & ""
                .TextMatrix(i, col����ID) = rsTmp!��ǰ����ID & ""
                .TextMatrix(i, COLƵ��) = rsTmp!Ƶ�� & ""
                .TextMatrix(i, col��ʼִ��ʱ��) = rsTmp!��ʼִ��ʱ�� & ""
                .RowData(i) = IIF(.TextMatrix(i, col���ID) = "", "Begin", "")
                '��ʾ���ģʽ�µ�ҽ������
                strFormat = rsTmp!ҽ������ & ""
                
                If .TextMatrix(i, col�������) & .TextMatrix(i, col��������) & .TextMatrix(i, colִ�з���) = "E21" Then
                    If rsTmp!ҽ������ & "" <> "" Then
                        strFormat = strFormat & " " & rsTmp!ҽ������
                    End If
                End If
                
                If .TextMatrix(i, COLƵ��) <> "һ����" Then
                    blnDo = True
                    If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                    If blnDo Then
                        strTmp = .TextMatrix(i, col����)
                        If strTmp <> "" Then strFormat = strFormat & ",��" & strTmp
                    End If
                End If
                .TextMatrix(i, colҽ������) = strFormat
                .TextMatrix(i, colҪ��ʱ��) = rsTmp!Ҫ��ʱ�� & ""
                '�ɱ༭����ɫ
                .Cell(flexcpBackColor, i, colѡ��, i, colѡ��) = COLEditBackColor
                If mbytUseType <> Tҽ���˶� Then
                    .Cell(flexcpBackColor, i, colִ��ʱ��, i, colִ����) = COLEditBackColor
                    .Cell(flexcpBackColor, i, colִ�н��, i, col��ע) = COLEditBackColor
                End If
                If blnIsExecute Or mbytUseType = Tҽ���˶� Then
                    .TextMatrix(i, colִ��ʱ��) = rsTmp!ִ��ʱ�� & ""
                    .Cell(flexcpData, i, colִ��ʱ��) = .TextMatrix(i, colִ��ʱ��)
                    '��¼ԭʼִ��ʱ�䣬����ɾ��ִ��
                    .TextMatrix(i, COLԭʼִ��ʱ��) = rsTmp!ԭʼִ��ʱ�� & ""
                    .TextMatrix(i, colִ����) = rsTmp!ִ���� & ""
                    .Cell(flexcpData, i, colִ����) = .TextMatrix(i, colִ����)
                    .TextMatrix(i, colִ�н��) = IIF(rsTmp!ִ�н�� & "" = "0", "δִ��", "���")
                    .Cell(flexcpData, i, colִ�н��) = .TextMatrix(i, colִ�н��)
                    .TextMatrix(i, COLδִ��ԭ��) = rsTmp!δִ��ԭ�� & ""
                    .Cell(flexcpData, i, COLδִ��ԭ��) = .TextMatrix(i, COLδִ��ԭ��)
                    .TextMatrix(i, col��ע) = rsTmp!��ע & ""
                    .Cell(flexcpData, i, col��ע) = .TextMatrix(i, col��ע)
                    .TextMatrix(i, col�Ǽ���) = rsTmp!�Ǽ��� & ""
                    .TextMatrix(i, col���ʱ��) = Format(rsTmp!���ʱ��, "yyyy-MM-dd HH:mm")
                    If rsTmp!��Ժ & "" <> "1" Then
                        '��Ժ���˵���ǳ��ɫ��ʾ
                        .Cell(flexcpBackColor, i, colѡ��, i, colѡ��) = &HE0E0E0
                        .Cell(flexcpBackColor, i, colִ��ʱ��, i, col��ע) = &HE0E0E0
                    End If
                    If blnIsExecute Then
                        .TextMatrix(i, COL�˶���) = rsTmp!�˶��� & ""
                        .Cell(flexcpData, i, COL�˶���) = .TextMatrix(i, COL�˶���)
                        .TextMatrix(i, col�˶�ʱ��) = rsTmp!�˶�ʱ�� & ""
                        .Cell(flexcpData, i, col�˶�ʱ��) = .TextMatrix(i, col�˶�ʱ��)
                        .TextMatrix(i, colִ��״̬) = rsTmp!ִ��״̬ & ""
                    Else
                        .ColHidden(COL�˶���) = True
                        .ColHidden(col�˶�ʱ��) = True
                    End If
                Else
                    .TextMatrix(i, COL���ִ����) = rsTmp!���ִ���� & ""
                    Call BatUpdate
                    .ColHidden(COL�˶���) = True
                    .ColHidden(col�˶�ʱ��) = True
                End If
                
                '��Ҫִ�е�ҽ������
                If rsTmp!���ID & "" = "" Then lngCount = lngCount + 1
                
                rsTmp.MoveNext
                i = i + 1
            Loop
        Else
            .AddItem ""
        End If
                
        If blnIsExecute Then
            stbThis.Panels(2).Text = "���� " & lngCount & " ��ҽ���Ѿ�" & IIF(mbytUseType <> Tҽ���˶�, "ִ��", "�˶�") & "��"
            mblnExecIsUpdate = False
        Else
            stbThis.Panels(2).Text = "���� " & lngCount & " ��ҽ����Ҫ" & IIF(mbytUseType <> Tҽ���˶�, "ִ��", "�˶�") & "��"
            mblnWaitIsUpdate = False
        End If
        '�Զ������и�
        .AutoSize colҽ������
        .Redraw = flexRDDirect
        '�ָ�ǰ��ɫ
        .Cell(flexcpForeColor, 1, colѡ��, .Rows - 1, colѡ��) = vbBlack
        .Cell(flexcpForeColor, 1, colִ��ʱ��, .Rows - 1, col��ע) = vbBlack
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get��Ѫ��������(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal datҪ��ʱ�� As Date, ByVal dbl���� As Double, ByVal dbl���� As Double) As Double
'���ܣ�����ҽ����Ϣִ��ʱ�䣬�����Ѫҽ����������
    Dim strSQL As String, rsTmp As Recordset
    Dim lng��ǰ���� As Long, i As Long
    Dim dbl����Tmp As Double, dbl���� As Double
    
    strSQL = "Select Ҫ��ʱ�� From ҽ��ִ��ʱ�� Where ҽ��id = [1] And ���ͺ� = [2] Order By Ҫ��ʱ��"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get��Ѫ��������", lngҽ��ID, lng���ͺ�)
    dbl����Tmp = dbl����
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp.RecordCount = 1 Then
                dbl���� = dbl����
            Else
                If i = rsTmp.RecordCount Then
                    dbl���� = dbl����Tmp
                Else
                    If dbl����Tmp >= dbl���� Then
                        dbl���� = dbl����
                    Else
                        dbl���� = dbl����Tmp
                    End If
                    dbl����Tmp = dbl����Tmp - dbl����
                End If
            End If
            If CDate(Format(rsTmp!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:mm:ss")) = datҪ��ʱ�� Then
                Get��Ѫ�������� = dbl����
                Exit For
            End If
            rsTmp.MoveNext
        Next
    Else
        Get��Ѫ�������� = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadPatiInfo()
'���ܣ����ز����б�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long, lngUnitID As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngSelectRow As Long
        
    On Error GoTo errH
    lngUnitID = mlng����ID
    If mlngӤ������ID <> 0 Then
        If mlngӤ������ID = mlngҽ������ID Or mlngӤ������ID = mlngҽ������ID Then
            lngUnitID = mlngӤ������ID
        End If
    End If
    
    str����IDs = zlDatabase.GetPara("���Ͳ���", glngSys, pסԺҽ������)
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
            
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng����ID, False, False, False)
    lngSelectRow = -1
    With rptPati
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!��˱�־ & "") < 1 Or gbyt������˷�ʽ <> 1 Then
                Set objRecord = .Records.Add()
                objRecord.Tag = "0"
                Set objItem = objRecord.AddItem(rsTmp!����ID & "")
                Set objItem = objRecord.AddItem(rsTmp!��ҳID & "")
                Set objItem = objRecord.AddItem("")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                    objItem.Icon = img16.ListImages.Item(IIF(rsTmp!�Ա� & "" = "��", "Man", "Woman")).Index - 1
                Set objItem = objRecord.AddItem(rsTmp!�Ա� & "")
                Set objItem = objRecord.AddItem(rsTmp!סԺ�� & "")
                
                
                '������ɫ
                objRecord.Item(0).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
                For j = 1 To objRecord.Childs.Count - 1
                    objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
                Next
                
                '�ϴ��Ƿ�ѡ��
                If lngUnitID = lng����ID And str����IDs <> "" Then
                    If InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") > 0 Or str����IDs = "ALL" Then
                        objRecord.Item(COL_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
                        objRecord.Tag = "1"
                        lngSelectRow = objRecord.Index
                    End If
                ElseIf rsTmp!����ID = mlng����ID Then
                    objRecord.Item(COL_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
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
        
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(COL_��ҳID, "��ҳID", 0, False)
        Set objCol = .Columns.Add(COL_ѡ��, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(COL_����, "����", 45, True)
        Set objCol = .Columns.Add(COL_����, "����", 80, True)
        Set objCol = .Columns.Add(COL_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(COL_סԺ��, "סԺ��", 60, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub

Private Sub InitPageData()
'���ܣ���ʼ������
    Dim curDate As Date
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim i As Long
    Dim strTmp As String
    Dim objCbo As Object
    
    curDate = zlDatabase.Currentdate
    
    dpkExecuteTime.value = curDate
    dpkReqTime(dpk��ʼ����).value = Format(curDate, "yyyy-MM-dd 00:00:00")
    dpkReqTime(dpk��������).value = Format(curDate, "yyyy-MM-dd 23:59:59")
    
    cboExecuteResult.AddItem "���"
    cboExecuteResult.AddItem "δִ��"
    cboExecuteResult.ListIndex = 0
    vsgWaitExecAdvice.ColComboList(colִ�н��) = "���|δִ��"
    vsgExecAdvice.ColComboList(colִ�н��) = "���|δִ��"
    
    strSQL = "Select ���� From ҽ��δִ��ԭ��"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        vsgWaitExecAdvice.ColComboList(COLδִ��ԭ��) = vsgWaitExecAdvice.ColComboList(COLδִ��ԭ��) & "|" & rsTmp!���� & ""
        vsgExecAdvice.ColComboList(COLδִ��ԭ��) = vsgExecAdvice.ColComboList(COLδִ��ԭ��) & "|" & rsTmp!���� & ""
        rsTmp.MoveNext
    Loop
    
    vsgWaitExecAdvice.ColComboList(COLδִ��ԭ��) = Mid(vsgWaitExecAdvice.ColComboList(COLδִ��ԭ��), 2)
    vsgExecAdvice.ColComboList(COLδִ��ԭ��) = Mid(vsgExecAdvice.ColComboList(COLδִ��ԭ��), 2)
    
    strSQL = "Select Distinct a.Id, a.���, a.����" & vbNewLine & _
            "From ��Ա�� A, ������Ա B, ��Ա����˵�� C" & vbNewLine & _
            "Where a.Id = b.��Աid And a.Id = c.��Աid And (c.��Ա���� = '��ʿ' or c.��Ա���� = 'ҽ��') And b.����id = [1]" & _
            " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) Order By a.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    
    cboExecutePoeple(cboִ���˵Ǽ�).AddItem ""
    cboExecutePoeple(cboִ����ȡ��).AddItem ""
    
    Do While Not rsTmp.EOF
        For i = 0 To 1
            With cboExecutePoeple(i)
                .AddItem rsTmp!���� & ""
                .ItemData(.NewIndex) = rsTmp!ID & ""
                If rsTmp!���� = UserInfo.���� Then
                    .ListIndex = .ListCount - 1
                End If
            End With
        Next
        rsTmp.MoveNext
    Loop
    
    '��������
    strTmp = zlDatabase.GetPara("ҽ��ִ����Ч", glngSys, pסԺҽ������, "11")
    chk��Ч(chk����).value = Val(Mid(strTmp, 1, 1))
    chk��Ч(chk����).value = Val(Mid(strTmp, 2, 1))
    strTmp = zlDatabase.GetPara("ҽ��ִ�з�Χ", glngSys, pסԺҽ������, "1111111111")
    chkType(Type��Һ).value = Val(Mid(strTmp, 1, 1))
    chkType(Typeע��).value = Val(Mid(strTmp, 2, 1))
    chkType(Type�ڷ�).value = Val(Mid(strTmp, 3, 1))
    chkType(TypeƤ��).value = Val(Mid(strTmp, 4, 1))
    chkType(Type����).value = Val(Mid(strTmp, 5, 1))
    chkType(Type�ɼ�).value = Val(Mid(strTmp, 6, 1))
    chkType(Type����ҽ��).value = Val(Mid(strTmp, 7, 1))
    chkType(Type��Ѫ).value = Val(Mid(strTmp, 8, 1))
    chkType(Type������ҩ;��).value = Val(Mid(strTmp, 9, 1))
    chkType(Type��ҩ����).value = Val(Mid(strTmp, 10, 1))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim str����IDs As String
    
    If mblnWaitIsUpdate Then
        tbcSub.Item(0).Selected = True
        If MsgBox("��ִ��ҽ��������δ���棬�Ƿ�Ҫ�˳���", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    If mblnExecIsUpdate Then
        tbcSub.Item(1).Selected = True
        If MsgBox("��ִ��ҽ��������δ���棬�Ƿ�Ҫ�˳���", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    '���汨��������
    str����IDs = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            str����IDs = str����IDs & "," & rptPati.Rows(i).Record(COL_����ID).value
        End If
    Next
    str����IDs = Mid(str����IDs, 2)
    If str����IDs <> "" Then
        If UBound(Split(str����IDs, ",")) = 0 And Val(str����IDs) = mlng����ID Then
            Call zlDatabase.SetPara("���Ͳ���", "", glngSys, pסԺҽ������)
        Else
            Call zlDatabase.SetPara("���Ͳ���", mlng����ID & ":" & str����IDs, glngSys, pסԺҽ������)
        End If
    End If

    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optBaby_Click(Index As Integer)
    mintҽ������Χ = Index
End Sub

Private Sub optExecutePeople_Click(Index As Integer)
    cboExecutePoeple(cboִ���˵Ǽ�).Enabled = optExecutePeople(optָ����Ա).value
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
    vsgWaitExecAdvice.Top = IIF(mbytUseType <> Tҽ���˶�, picExec.Height, 0)
    vsgWaitExecAdvice.Height = picWaitExecute.Height - vsgWaitExecAdvice.Top
    vsgWaitExecAdvice.Width = picWaitExecute.Width
End Sub

Private Sub picPati_Resize()
    Dim lngTmp As Long
    
    On Error Resume Next
    
    picPati.Height = tbcSub.Height - 400
    If tbcSub.Selected.Tag = "��ִ��ҽ��" Then
        If mbytUseType <> Tҽ���˶� Then
            picFitter.Height = 2850
        Else
            picFitter.Height = 1800
        End If
    Else
        If mbytUseType <> Tҽ���˶� Then
            picFitter.Height = 3170
        Else
            picFitter.Height = 2120
        End If
    End If

    lblInfo(lbl����).Top = 60
    lblInfo(lbl����).Left = 120
    rptPati.Left = 120
    rptPati.Top = 400
    rptPati.Width = picPati.Width - rptPati.Left
    
    rptPati.Height = picPati.Height - rptPati.Top - picFitter.Height - 300
    fraBaby.Top = rptPati.Top + rptPati.Height + 100
    fraBaby.Left = rptPati.Left + 20
    picFitter.Top = picPati.Height - picFitter.Height
    picFitter.Left = rptPati.Left - 80
    
    lngTmp = dpkReqTime(dpk��������).Top + dpkReqTime(dpk��������).Height + 50
    picPanel.Left = 0
    If tbcSub.Selected.Tag = "��ִ��ҽ��" Then
        lblInfo(lblִ����).Visible = False
        cboExecutePoeple(cboִ����ȡ��).Visible = False
        picPanel.Top = lngTmp + 50
    Else
        lblInfo(lblִ����).Visible = True
        cboExecutePoeple(cboִ����ȡ��).Visible = True
        cboExecutePoeple(cboִ����ȡ��).Top = lngTmp
        picPanel.Top = cboExecutePoeple(cboִ����ȡ��).Top + cboExecutePoeple(cboִ����ȡ��).Height + 70
    End If
    
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptPati.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptPati_RowDblClick(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(COL_ѡ��))
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COL_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptPati.Columns(COL_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(COL_ѡ��).Icon = img16.ListImages("Check").Index - 1
                            rptPati.Rows(i).Record.Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptPati.Columns(COL_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(COL_ѡ��).Icon = -1
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
        Row.Record.Item(COL_ѡ��).Icon = -1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(COL_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
        Row.Record.Tag = "1"
    End If
    rptPati.Populate
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If .TextMatrix(lngRow, col�������) = "" Then Exit Function
        If .TextMatrix(lngRow, col�������) = "�������" Then Exit Function
        '��Ҫ��ʱ������Ϊ�п���ͬһ��ҽ���Ķ���ִ�м�¼��һ��
        If .TextMatrix(lngRow - 1, colҪ��ʱ��) = .TextMatrix(lngRow, colҪ��ʱ��) And (Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 Or Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, colҽ��ID)) Or Val(.TextMatrix(lngRow - 1, colҽ��ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow - 1, colҽ��ID)) <> 0) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If .TextMatrix(lngRow + 1, colҪ��ʱ��) = .TextMatrix(lngRow, colҪ��ʱ��) And (Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow + 1, col���ID)) <> 0 Or Val(.TextMatrix(lngRow + 1, colҽ��ID)) = Val(.TextMatrix(lngRow, col���ID)) Or Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.TextMatrix(lngRow, colҽ��ID))) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If .TextMatrix(i, colҪ��ʱ��) = .TextMatrix(lngRow, colҪ��ʱ��) And (Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 And Val(.TextMatrix(i, colҽ��ID)) <> Val(.TextMatrix(lngRow, colҽ��ID)) Or Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, colҽ��ID)) Or Val(.TextMatrix(i, colҽ��ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(i, colҽ��ID)) <> 0) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If .TextMatrix(i, colҪ��ʱ��) = .TextMatrix(lngRow, colҪ��ʱ��) And (Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 And Val(.TextMatrix(i, colҽ��ID)) <> Val(.TextMatrix(lngRow, colҽ��ID)) Or Val(.TextMatrix(i, colҽ��ID)) = Val(.TextMatrix(lngRow, col���ID)) Or Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, colҽ��ID))) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        Else
            .RowData(lngRow) = "Begin"
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Me.Visible Then Exit Sub
    If Item.Tag = "��ִ��ҽ��" Then
        'Call LoadAdvice
    Else
'        Call LoadAdvice(True)
    End If
End Sub

Private Sub vsgExecAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsgExecAdvice.Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor And vsgExecAdvice.RowData(NewRow) = "Begin" Then
        If NewCol = colִ���� Then
            If (Val(zlDatabase.GetPara(51)) = 1 And mbytUseType = Tҽ��ִ�� Or mbytUseType = Tҽ���˶�) Then
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
        If NewCol = COLδִ��ԭ�� And vsgExecAdvice.TextMatrix(NewRow, colִ�н��) = "���" Then
            vsgExecAdvice.FocusRect = flexFocusNone
            vsgExecAdvice.Editable = flexEDNone
        Else
            vsgExecAdvice.FocusRect = flexFocusHeavy
            If NewCol <> colѡ�� Then
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
    
    If Col = colִ���� Then
        With vsgExecAdvice
            strSQL = "Select a.Id, a.���, a.����" & vbNewLine & _
                        "From ��Ա�� A, ������Ա B, ��Ա����˵�� C" & vbNewLine & _
                        "Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = '��ʿ' And b.����id = [1]" & _
                        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������ʿ", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    mlng����ID)
                    
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    If rsTmp!���� & "" <> .Cell(flexcpData, Row, Col) Then
                        .TextMatrix(Row, Col) = rsTmp!���� & ""
                        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                        mblnExecIsUpdate = True
                        .TextMatrix(Row, col�Ƿ��޸�) = "1"
                        '������ɫΪ����ɫ����
                        .Cell(flexcpForeColor, Row, Col) = &HFF0000
                    End If
                Else
                    MsgBox "��ǰ������û�п�ѡ�Ļ�ʿ��", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Sub vsgExecAdvice_DblClick()
    With vsgExecAdvice
        If .MouseCol = colѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsgExecAdvice_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsgExecAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsgExecAdvice
        lngLeft = colѡ��: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col��������: lngRight = col��ע
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd, vsgExecAdvice) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If .TextMatrix(Row, col���ID) = "" Then
            vRect.Top = Bottom - 1 '���IDΪ�յ������ֱ���
            vRect.Bottom = Bottom - 1
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Between(Col, colѡ��, colѡ��) Or (Between(Col, colִ��ʱ��, colִ����) Or Between(Col, colִ�н��, col��ע)) And mbytUseType <> Tҽ���˶� Then
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
        '���ֱ�����뺺�ֵ�����
        Call vsgExecAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsgExecAdvice_KeyPress(KeyAscii As Integer)
    With vsgExecAdvice
        If .Col = col��ע Or .Col = colִ���� Then
            If KeyAscii = Asc("*") And .Col = colִ���� Then
                KeyAscii = 0
                Call vsgExecAdvice_CellButtonClick(.Row, .Col)
                Exit Sub
            End If
            .ComboList = "" 'ʹ��ť״̬��������״̬
        ElseIf .Col = colѡ�� And KeyAscii = vbKeySpace Then
            Call ExecCheck(vsgExecAdvice)
        End If
    End With
End Sub

Private Sub vsgExecAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vPoint As PointAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset
    
    With vsgExecAdvice
        '����Ƿ��ǵ�ǰ����Ա�Ǽǵ�
        If .TextMatrix(Row, col�Ǽ���) <> UserInfo.���� And .EditText <> .TextMatrix(Row, Col) Then
            MsgBox "ֻ��ȡ�����޸��Լ��Ǽǵ�ҽ����", vbInformation, Me.Caption
            .EditText = .Cell(flexcpData, Row, Col)
            Cancel = True
            Exit Sub
        End If
        '����Ƿ�˶�
        If .TextMatrix(Row, col�˶�ʱ��) <> "" And .EditText <> .TextMatrix(Row, Col) Then
            MsgBox "��ҽ���Ѿ��˶ԣ���ȡ���˶Ժ����ԡ�", vbInformation, Me.Caption
            .EditText = .Cell(flexcpData, Row, Col)
            Cancel = True
            Exit Sub
        End If
        If Col = colִ��ʱ�� Then
            If Not IsDate(.EditText) Then
                MsgBox "��������ȷ�����ڡ�", vbInformation, Me.Caption
                .EditText = .Cell(flexcpData, Row, Col)
                Cancel = True
                Exit Sub
            End If
        ElseIf Col = col��ע Then
            If zlCommFun.ActualLen(.EditText) > 200 Then
                MsgBox "��ע���ֲ��ܳ���100�����ֻ�200����ĸ��", vbInformation, Me.Caption
                .EditText = .Cell(flexcpData, Row, Col)
                Cancel = True
                Exit Sub
            End If
        ElseIf Col = colִ�н�� Then
            If .Cell(flexcpData, Row, Col) = "δִ��" And .EditText = "���" Then
                .TextMatrix(Row, COLδִ��ԭ��) = ""
                .Cell(flexcpData, Row, COLδִ��ԭ��) = ""
            End If
        ElseIf Col = colִ���� Then
            strSQL = "Select a.Id, a.���, a.����" & vbNewLine & _
                        " From ��Ա�� A, ������Ա B, ��Ա����˵�� C" & vbNewLine & _
                        " Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = '��ʿ'" & _
                        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " And b.����id = [1] And (a.���=[2] or a.���� Like [3] or a.���� Like [4])"
            
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������ʿ", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    mlng����ID, .EditText, gstrLike & .EditText & "%", gstrLike & UCase(.EditText) & "%")
                    
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    If Not rsTmp Is Nothing Then
                        .EditText = rsTmp!���� & ""
                    Else
                        MsgBox "û���ҵ���ָ���Ļ�ʿ��", vbInformation, Me.Caption
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
            .TextMatrix(Row, col�Ƿ��޸�) = "1"
            '������ɫΪ����ɫ����
            .Cell(flexcpForeColor, Row, Col) = &HFF0000
        End If
    End With
End Sub

Private Sub vsgWaitExecAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsgWaitExecAdvice.Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor And vsgWaitExecAdvice.RowData(NewRow) = "Begin" Then
        If NewCol = colִ���� Then
            If (Val(zlDatabase.GetPara(51)) = 1 And mbytUseType = Tҽ��ִ�� Or mbytUseType = Tҽ���˶�) Then
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
        If NewCol = COLδִ��ԭ�� And vsgWaitExecAdvice.TextMatrix(NewRow, colִ�н��) = "���" Then
            vsgWaitExecAdvice.FocusRect = flexFocusNone
            vsgWaitExecAdvice.Editable = flexEDNone
        Else
            vsgWaitExecAdvice.FocusRect = flexFocusHeavy
            If NewCol <> colѡ�� Then
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
    
    If Col = colִ���� Then
        With vsgWaitExecAdvice
            strSQL = "Select a.Id, a.���, a.����" & vbNewLine & _
                        "From ��Ա�� A, ������Ա B, ��Ա����˵�� C" & vbNewLine & _
                        "Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = '��ʿ' And b.����id = [1]" & _
                        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������ʿ", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    mlng����ID)
                    
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    If rsTmp!���� & "" <> .Cell(flexcpData, Row, Col) Then
                        .TextMatrix(Row, Col) = rsTmp!���� & ""
                        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                        mblnWaitIsUpdate = True
                    End If
                Else
                    MsgBox "��ǰ������û�п�ѡ�Ļ�ʿ��", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Sub vsgWaitExecAdvice_DblClick()
    With vsgWaitExecAdvice
        If .MouseCol = colѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsgWaitExecAdvice_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsgWaitExecAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsgWaitExecAdvice
        lngLeft = colѡ��: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col��������: lngRight = col��ע
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd, vsgWaitExecAdvice) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If .TextMatrix(Row, col���ID) = "" Then
            vRect.Top = Bottom - 1 '���IDΪ�յ������ֱ���
            vRect.Bottom = Bottom - 1
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Between(Col, colѡ��, colѡ��) Or Between(Col, colִ��ʱ��, col��ע) And mbytUseType <> Tҽ���˶� Then
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
        '���ֱ�����뺺�ֵ�����
        Call vsgWaitExecAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsgWaitExecAdvice_KeyPress(KeyAscii As Integer)
    With vsgWaitExecAdvice
        If .Col = col��ע Or .Col = colִ���� Then
            If KeyAscii = Asc("*") And .Col = colִ���� Then
                KeyAscii = 0
                Call vsgWaitExecAdvice_CellButtonClick(.Row, .Col)
                Exit Sub
            End If
            .ComboList = "" 'ʹ��ť״̬��������״̬
        ElseIf .Col = colѡ�� And KeyAscii = vbKeySpace Then
            Call ExecCheck(vsgWaitExecAdvice)
        End If
    End With
End Sub

Private Sub ExecCheck(ByRef objVsg As VSFlexGrid)
'���ܣ�ͬ��ѡ��һ��ҽ��
'���������
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    
    With objVsg
        If .TextMatrix(.Row, colҽ��ID) = "" Then Exit Sub
        If Not RowInһ����ҩ(.Row, lngBegin, lngEnd, objVsg) Then
            lngBegin = .Row: lngEnd = .Row
        End If
        
        For i = lngBegin To lngEnd
            If .Cell(flexcpData, i, colѡ��) = 1 Then
                Set .Cell(flexcpPicture, i, colѡ��) = Nothing
                .Cell(flexcpData, i, colѡ��) = 0
            Else
                If objVsg.Name = "vsgExecAdvice" Then
                    '����Ƿ��Ժ
                    If .TextMatrix(i, col��Ժ) <> "1" Then
                        MsgBox "�ò����Ѿ���Ժ������ȡ��ִ�С�", vbInformation, Me.Caption
                        Exit Sub
                    End If
                    '����Ƿ��ǵ�ǰ����Ա�Ǽǵ�
                    If .TextMatrix(i, col�Ǽ���) <> UserInfo.���� Then
                        MsgBox "ֻ��ȡ��ִ���Լ��Ǽǵ�ҽ����", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End If
                .Cell(flexcpPicture, i, colѡ��) = img16.ListImages("Check").Picture
                .Cell(flexcpData, i, colѡ��) = 1
                .Cell(flexcpPictureAlignment, i, colѡ��) = flexPicAlignCenterCenter
            End If
        Next
    End With
End Sub

Private Sub vsgWaitExecAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vPoint As PointAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset
    
    With vsgWaitExecAdvice
        If Col = colִ��ʱ�� Then
            If Not IsDate(.EditText) Then
                MsgBox "��������ȷ�����ڡ�", vbInformation, Me.Caption
                .EditText = .Cell(flexcpData, Row, Col)
                Cancel = True
                Exit Sub
            End If
        ElseIf Col = col��ע Then
            If zlCommFun.ActualLen(.EditText) > 200 Then
                MsgBox "��ע���ֲ��ܳ���100�����ֻ�200����ĸ��", vbInformation, Me.Caption
                .EditText = .Cell(flexcpData, Row, Col)
                Cancel = True
                Exit Sub
            End If
        ElseIf Col = colִ�н�� Then
            If .Cell(flexcpData, Row, Col) = "δִ��" And .EditText = "���" Then
                .TextMatrix(Row, COLδִ��ԭ��) = ""
                .Cell(flexcpData, Row, COLδִ��ԭ��) = ""
            End If
        ElseIf Col = colִ���� Then
            strSQL = "Select a.Id, a.���, a.����" & vbNewLine & _
                        " From ��Ա�� A, ������Ա B, ��Ա����˵�� C" & vbNewLine & _
                        " Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = '��ʿ'" & _
                        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " And b.����id = [1] And (a.���=[2] or a.���� Like [3] or a.���� Like [4])"
            
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������ʿ", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    mlng����ID, .EditText, gstrLike & .EditText & "%", gstrLike & UCase(.EditText) & "%")
                    
            If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                Cancel = True
            Else
                If Not rsTmp Is Nothing Then
                    .EditText = rsTmp!���� & ""
                Else
                    MsgBox "û���ҵ���ָ���Ļ�ʿ��", vbInformation, Me.Caption
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
'���ܣ����ҽ��
    Dim i As Long
    
    With objVsg
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colҽ��ID) <> "" And .RowData(i) = "Begin" And (.Cell(flexcpData, i, colѡ��) = "1" Or objVsg.Name = "vsgExecAdvice" And objVsg.TextMatrix(i, col�Ƿ��޸�) = "1") Then
                'ִ���˲���Ϊ��
                If .TextMatrix(i, colִ����) = "" Then
                    .Row = i: .Col = colִ����
                    Call ShowMessage(vsgWaitExecAdvice, "ִ���˲���Ϊ�ա�")
                    Exit Function
                End If
                
                '���Ϊδִ�еģ�����ԭ��
                If .TextMatrix(i, colִ�н��) = "δִ��" And .TextMatrix(i, COLδִ��ԭ��) = "" Then
                    .Row = i: .Col = COLδִ��ԭ��
                    Call ShowMessage(vsgWaitExecAdvice, "δִ�е�ҽ��������дδִ��ԭ��")
                    Exit Function
                End If
            End If
        Next
    End With
    
    CheckData = True
End Function

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'���ܣ���ʾ��ʾ��Ϣ����λ��������Ŀ��
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
        Case col����, COL��λ, col�Ա�, colҪ��ʱ��, colִ��ʱ�� '�������
        Case Else
            Order = 0
    End Select
  
    If Col = colѡ�� Then
        With objVs
            If .TextMatrix(1, colҽ��ID) = "" Then Exit Sub
            If .ColData(colѡ��) = "Check" Then
                .Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("UnCheck").Picture
                .ColData(colѡ��) = ""
            Else
                .Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("AllCheck").Picture
                .ColData(colѡ��) = "Check"
            End If
            For i = 1 To .Rows - 1
                If .TextMatrix(i, colҽ��ID) = "" Then Exit For
                If .ColData(colѡ��) = "Check" Then
                    .Cell(flexcpPicture, i, colѡ��) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, colѡ��) = 1
                    .Cell(flexcpPictureAlignment, i, colѡ��) = flexPicAlignCenterCenter
                Else
                    Set .Cell(flexcpPicture, i, colѡ��) = Nothing
                    .Cell(flexcpData, i, colѡ��) = 0
                End If
            Next
        End With
    End If
End Sub
