VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmOpsStation 
   Caption         =   "�����ҹ���վ"
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
      ToolTipText     =   "��ݼ���F3"
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
                  Caption         =   "��"
                  BeginProperty Font 
                     Name            =   "����"
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
               Caption         =   "סԺ��:"
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
               Caption         =   "����:"
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
               Caption         =   "��������:"
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
               Caption         =   "�� �� ��:"
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
               Caption         =   "�������:"
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
               Caption         =   "����ʱ��:"
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
               Caption         =   "��������:"
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
               Caption         =   "������Դ:"
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
               Caption         =   "���˿���:"
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
         Caption         =   "������"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "�༭"
            TextSave        =   "�༭"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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

'���弶��������
'######################################################################################################################
Private mstrPrivs As String
Private mblnDataChanged As Boolean
Private mblnStartUp As Boolean
Private mblnAllowClose As Boolean
Private mlngSvrDept As Long
Private mintIndex As Integer
Private mlng��������id As Long
Private mstr�������� As String
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
''�Զ�����̻���
''######################################################################################################################
'

Private Property Let AutoRefresh(vData As Boolean)
    '
    '����:�Զ�ˢ��
    '
    tmr.Enabled = vData
    
    If vData = True Then
        mlngCountTmr = 0
        tmr.Tag = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�Զ�ˢ�¼��", 0))
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
    '���ܣ��л���ʾ����/סԺҽ��ҳ��
    '������blnClinic=�Ƿ���ʾ����ҽ��ҳ
    '���أ��Ƿ�������л�ѡ��
    '******************************************************************************************************************
    Dim blnSel As Boolean
    Dim blnOld As Boolean
    Dim intIdx As Integer

    If Not tbcPage.Selected Is Nothing Then
        blnSel = tbcPage.Selected.Tag Like "*ҽ��"
    End If

    For intIdx = 0 To tbcPage.ItemCount - 1
        If tbcPage(intIdx).Tag = "����ҽ��" Then
            If tbcPage(intIdx).Visible <> blnClinic Then
                tbcPage(intIdx).Visible = blnClinic
                If blnSel And blnClinic Then
                    tbcPage(intIdx).Selected = True
                    ExchangeAdvice = True
                End If
            End If
        ElseIf tbcPage(intIdx).Tag = "סԺҽ��" Then
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
    '���ܣ��л���ʾ����/סԺ����ҳ��
    '������blnClinic=�Ƿ���ʾ���ﲡ��ҳ
    '���أ��Ƿ�������л�ѡ��
    '******************************************************************************************************************
    Dim blnSel As Boolean
    Dim blnOld As Boolean
    Dim intIdx As Integer

    If Not tbcPage.Selected Is Nothing Then
        blnSel = tbcPage.Selected.Tag Like "*����"
    End If

    For intIdx = 0 To tbcPage.ItemCount - 1
        If tbcPage(intIdx).Tag = "���ﲡ��" Then
            If tbcPage(intIdx).Visible <> blnClinic Then
                tbcPage(intIdx).Visible = blnClinic
                If blnSel And blnClinic Then
                    tbcPage(intIdx).Selected = True
                    ExchangeEPRS = True
                End If
            End If
        ElseIf tbcPage(intIdx).Tag = "סԺ����" Then
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
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Select Case strMenuKind
    '------------------------------------------------------------------------------------------------------------------
    Case "��Ҫ", "ҩƷ", "����", "����"
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", objMenu.Index + 1, False)
        objMenu.ID = conMenu_EditPopup
        Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "�������(&S)", True)
        Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ������(&C)")
    End Select


    '����������:���ļ�������˵������ť֮��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain(2)

    For Each objControl In objBar.Controls  '�����ǰ������һ��Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "����", True, , , , objControl.Index + 1)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��", , , , , objControl.Index + 1)

    '����Ŀ����:���������������Ѵ���
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save              '����
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
    '���ܣ�ˢ���Ӵ���˵���������
    '������
    '���أ�
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
    
    '��¼���в˵���ʽ
    '------------------------------------------------------------------------------------------------------------------
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        blnShowBar = cbsMain(2).Visible
        bytStyle = cbsMain(2).Controls(1).STYLE
    End If

    'ˢ���Ӵ��ڲ˵�
    '------------------------------------------------------------------------------------------------------------------
    Call LockWindowUpdate(Me.hWnd)

    'ɾ�����ڵĹ������������˵���
    '------------------------------------------------------------------------------------------------------------------
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    
    For lngCount = cbsMain.Count To 2 Step -1
        If lngCount <> 3 Then
            cbsMain(lngCount).Controls.DeleteAll
        End If
    Next

    '���������¼���
    '------------------------------------------------------------------------------------------------------------------

    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap

    '�ļ�
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)", , , "Ԥ������")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)", , , "��ӡ����")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel", , , "�����Excel")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BillPrintView, "Ԥ��֪ͨ��", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BillPrint, "��ӡ֪ͨ��")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "��������(&M)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Option, "ִ�м�����(&R)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)


    '����
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "ִ��(&C)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Request, "��¼�Ǽ�(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Audit, "�������(&K)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_UnAudit, "ȡ�����(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Arrange, "��������(&M)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_UnArrange, "ȡ������(&X)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Plan, "ִ�б���(&L)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Logout, "ȡ������(&R)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Complete, "ִ�����(&I)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Undone, "ȡ�����(&U)")

    '�鿴
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Filter, "����(&F)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Jump, "������ת(&J)")

    '����
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & ParamInfo.��Ʒ����)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.��Ʒ���� & "��ҳ(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.��Ʒ���� & "��̳(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)

    '����������:������������
    '------------------------------------------------------------------------------------------------------------------
    If cbsMain.Count < 2 Then
        Set objBar = cbsMain.Add("��׼", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagHideWrap
    Else
        Set objBar = cbsMain(2)
    End If
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Audit, "���", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Arrange, "����")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Plan, "����")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Complete, "���")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")

    '��λ������
    '------------------------------------------------------------------------------------------------------------------
    If cbsMain.Count < 3 Then
        Set objExtendedBar = cbsMain.Add("��λ", xtpBarTop)
        
        objExtendedBar.ContextMenuPresent = False
        objExtendedBar.ShowTextBelowIcons = False
        objExtendedBar.EnableDocking xtpFlagHideWrap

        mstrFindKey = Trim(GetRegister(˽��ģ��, Me.Name, "��λ����", "����"))
        If mstrFindKey = "" Then mstrFindKey = "����"

        Set mobjFindKey = NewToolBar(objExtendedBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , , "��ݼ�:F4")
        mobjFindKey.IconId = conMenu_View_Find

        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.����"): objControl.Parameter = "����"
        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.����"): objControl.Parameter = "����"
        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&3.סԺ��"): objControl.Parameter = "סԺ��"
        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&4.�����"): objControl.Parameter = "�����"
        Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&5.��������"): objControl.Parameter = "��������"
        
        Set cbrCustom = NewToolBar(objExtendedBar, xtpControlCustom, conMenu_View_Location, "")
        cbrCustom.Handle = txtLocation.hWnd
        
        Set objControl = NewToolBar(objExtendedBar, xtpControlButton, conMenu_View_Forward, "ǰһ��", , , , "��ݼ�:Ctrl+Left")
        Set objControl = NewToolBar(objExtendedBar, xtpControlButton, conMenu_View_Backward, "��һ��", , , , "��ݼ�:Ctrl+Right")

        Call SetDockRight(objExtendedBar, objBar)
    End If
    
    '����Ŀ����:���������������Ѵ���
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF12, conMenu_File_Parameter        '��������
        .Add 0, vbKeyF5, conMenu_View_Refresh           'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help              '����
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save              '����
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add 0, vbKeyF4, conMenu_View_Option                'ѡ��λ����
        .Add 0, vbKeyF6, conMenu_View_Jump           '
        .Add FCONTROL, vbKeyP, conMenu_File_Print       '��ӡ
        .Add FCONTROL, vbKeyV, conMenu_File_Preview
        .Add FCONTROL, vbKeyN, conMenu_Manage_Request     '����
        .Add FCONTROL, vbKeyK, conMenu_Manage_Audit     '���
        .Add FCONTROL, vbKeyM, conMenu_Manage_Arrange     '����
        .Add FCONTROL, vbKeyL, conMenu_Manage_Plan     '����
        .Add FCONTROL, vbKeyI, conMenu_Manage_Complete     '���
        .Add FCONTROL, vbKeyF, conMenu_View_Filter        '����
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
    End With
    

    '�Ӵ������¼���
    '------------------------------------------------------------------------------------------------------------------
    Select Case objItem.Tag
    Case "��Ҫ", "ҩƷ", "����", "����"
        Call zlDefCommandBars(objItem.Tag)
    Case "����"
        Call mclsExpenses.zlDefCommandBars(Me, cbsMain)
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_EditPopup)
'        Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Edit_Preferences, "������������(R)", True)
    Case "����ҽ��"
        Call mclsOutAdvices.zlDefCommandBars(Me, cbsMain, 2)
    Case "סԺҽ��"
        Call mclsInAdvices.zlDefCommandBars(Me, cbsMain, 2)
    Case "���ﲡ��"
        Call mfrmChildStationOutEPR.zlDefCommandBars(cbsMain)
    Case "סԺ����"
        Call mfrmChildStationInEPR.zlDefCommandBars(cbsMain)
    End Select

    '�ָ����̶���һЩ�˵�����
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
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

    '�������RecalcLayout����������
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
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    objPane.Title = "���������б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable

    Set objPane = dkpMain.CreatePane(2, 350, 300, DockRightOf, Nothing)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 350, 150, DockBottomOf, objPane)
    objPane.Title = "ҵ��"
    objPane.Options = PaneNoCaption

    dkpMain.SetCommandBars cbsMain
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.HideClient = True
End Sub

Private Sub InitTabControl()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************

    '��ߵ�������״̬����ҳ��
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

        .InsertItem 0, "�ȴ�����", picPane(1).hWnd, 0
        .InsertItem 1, "��������", picPane(2).hWnd, 0
        .InsertItem 2, "��������", picPane(3).hWnd, 0

        .Item(0).Selected = True

    End With

    '�ұߵľ�������ҳ��
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

        .InsertItem(0, "������Ϣ", mfrmChildStationOutLine.hWnd, 0).Tag = "��Ҫ"
        .InsertItem(1, "��ҩ��", mfrmChildStationDrug.hWnd, 0).Tag = "ҩƷ"
        .InsertItem(2, "���ϵ�", mfrmChildStationMaterial.hWnd, 0).Tag = "����"
        .InsertItem(3, "���Ƶ�", mfrmChildStationCure.hWnd, 0).Tag = "����"
        
        If GetInsidePrivs(pҽ�����ѹ���, True) <> "" Then
            .InsertItem(4, "���� ", mclsExpenses.zlGetForm.hWnd, 0).Tag = "����"
        End If

        If GetInsidePrivs(p����ҽ���´�, True) <> "" Then
            .InsertItem(5, " ҽ�� ", mclsOutAdvices.zlGetForm.hWnd, 0).Tag = "����ҽ��"
        End If

        If GetInsidePrivs(pסԺҽ���´�, True) <> "" Then
            .InsertItem(6, " ҽ�� ", mclsInAdvices.zlGetForm.hWnd, 0).Tag = "סԺҽ��"
        End If

        .InsertItem(7, " ���� ", mfrmChildStationInEPR.hWnd, 0).Tag = "סԺ����"
        .InsertItem(8, " ���� ", mfrmChildStationOutEPR.hWnd, 0).Tag = "���ﲡ��"

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
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    Dim intCount As Integer
    Dim int���� As Integer
    Dim intTmp As Integer
    Dim blnZero As Boolean

    On Error GoTo errHand
    
    AutoRefresh = False
    
    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        For intCount = 0 To 2
            Set mclsVsf(intCount) = New clsVsf
            With mclsVsf(intCount)
                Call .Initialize(Me.Controls, vsf(intCount), True, True, frmPubResource.GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("ͼ��", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
                Call .AppendColumn("������־", 255, flexAlignCenterCenter, flexDTString, "", "[������־]", False)
                Call .AppendColumn("��������", 1800, flexAlignLeftCenter, flexDTString, "", "ҽ������", True)
                Call .AppendColumn("����", 990, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("������Դ", 990, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("���˿���", 1080, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("����ʱ��", 1680, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", "����ʱ��", True)
                Call .AppendColumn("����", 600, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("סԺ��", 810, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("�����", 810, flexAlignLeftCenter, flexDTString, "", "", True)
                Call .AppendColumn("���ͺ�", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)

                Call .AppendColumn("��ǰ����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("��ǰ����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("ҽ��id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("״̬", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("�Һŵ�", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("ִ��״̬", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("��Ժ����", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("����״̬", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("������Ŀid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
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
    Case "��ʼ����"

        mlngSvrDept = 0
        mintIndex = 0

        strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & Me.Name, "������ʱ�䷶Χ", "��  ��"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & Me.Name, "������ʱ�䷶Χ", "��  ��"), 2)
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
        
        strEnd = Format(strEnd, "yyyy-MM-dd") & " 23:59:59"
        
        mstrCondition = strStart & ";" & strEnd & ";;;;;;"

        strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & Me.Name, "��������ʱ�䷶Χ", "��  ��"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & Me.Name, "��������ʱ�䷶Χ", "��  ��"), 2)
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
        strEnd = Format(strEnd, "yyyy-MM-dd") & " 23:59:59"
        mstrCondition = mstrCondition & ";" & strStart & ";" & strEnd & ";0;"

        '��ȡ��첿��
        '--------------------------------------------------------------------------------------------------------------
        cboDept.Clear
        If IsPrivs(mstrPrivs, "����������") Then
            gstrSQL = GetPublicSQL(SQL.���������嵥, "����")
        Else
            gstrSQL = GetPublicSQL(SQL.���������嵥, "")
        End If

        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID)
        If rs.BOF = False Then
            Do While Not rs.EOF
                cboDept.AddItem rs("����").Value & " - " & rs("����").Value
                cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
                rs.MoveNext
            Loop
        Else
            ShowSimpleMsg "û������������Ϣ��������Ա�����м����Ա�Ŀ������Լ�Ȩ�ޣ�"
            GoTo errEnd
        End If

        On Error Resume Next
        zlControl.CboLocate cboDept, UserInfo.����ID, True
        On Error GoTo errHand
        If cboDept.ListIndex < 0 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0

        mlng��������id = cboDept.ItemData(cboDept.ListIndex)
        mstr�������� = zlCommFun.GetNeedName(cboDept.List(cboDept.ListIndex))

    '--------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"

        If tbc.Enabled <> Not DataChanged Then
            tbc.Enabled = Not DataChanged
            vsf(0).Enabled = Not DataChanged
            vsf(0).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
            vsf(1).Enabled = Not DataChanged
            vsf(1).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
            vsf(2).Enabled = Not DataChanged
            vsf(2).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
        End If
        stbThis.Panels(3).Enabled = DataChanged

    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"

        Call ExecuteCommand("װ�صȴ�����")
        Call ExecuteCommand("װ����������")
        Call ExecuteCommand("װ����������")

        With vsf(mintIndex)
            Call vsf_AfterRowColChange(mintIndex, 0, 0, .Row, .Col)
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "���ز�������"

        ExecuteCommand = frmOpsStationPara.ShowPara(Me)
        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "����������"

        ExecuteCommand = frmOpsStationRoom.ShowEdit(Me, mlng��������id)
        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "ֱ����������"
        ExecuteCommand = frmOpsStationRequest.ShowEdit(Me, mlng��������id)
        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ������Ϣ"

        gstrSQL = "SELECT Distinct B.��ǰ����,b.�����,b.סԺ��,A.ID,B.����,C.���ͺ�,A.ҽ������,A.����ҽ��,A.����ʱ��,Decode(A.������Դ,1,'����',2,'סԺ',3,'����') AS ��Դ,Decode(A.������Դ,1,B.��������,2,E.����) AS ���˿���,G.���� AS ������� " & _
                "FROM ����ҽ����¼ A, ������Ϣ B,����ҽ������ C,���ű� E,���ű� G " & _
                "Where A.����id = B.����id AND A.ID=C.ҽ��id(+) AND E.ID(+)=B.��ǰ����ID  " & _
                    "AND A.ID=[1] AND A.��������ID=G.ID "

        With vsf(mintIndex)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))))
            If rs.BOF = False Then

                lblValue(1).Caption = zlCommFun.NVL(rs("����ҽ��"))
                lblValue(0).Caption = Format(zlCommFun.NVL(rs("����ʱ��")), "YYYY-MM-DD HH:MM")
                lblValue(2).Caption = zlCommFun.NVL(rs("�������"))
                lblValue(3).Caption = zlCommFun.NVL(rs("����"))
                lblValue(4).Caption = zlCommFun.NVL(rs("��Դ"))
                lblValue(5).Caption = zlCommFun.NVL(rs("���˿���"))
                lblValue(6).Caption = zlCommFun.NVL(rs("��ǰ����"))
                If zlCommFun.NVL(rs("��Դ")) = "סԺ" Then
                    lbl(1).Caption = "סԺ��"
                    lblValue(7).Caption = zlCommFun.NVL(rs("סԺ��"))
                Else
                    lbl(1).Caption = "�����"
                    lblValue(7).Caption = zlCommFun.NVL(rs("�����"))
                End If
                
                lblValue(10).Caption = zlCommFun.NVL(rs("ҽ������"))

                picState.Visible = CheckChargeState(Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), zlCommFun.NVL(rs("���ͺ�").Value, 0))
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
    Case "װ�صȴ�����"

        mclsVsf(0).ClearGrid

        strTmp = ""

        If Split(mstrCondition, ";")(0) <> "" Then
            strTmp = " AND a.����ʱ�� BETWEEN [2] AND [3] "
        End If

        If Trim(Split(mstrCondition, ";")(2)) <> "" Then
            strTmp = strTmp & " AND b.���� LIKE [4] "
        End If

        If Trim(Split(mstrCondition, ";")(3)) <> "" Then
            strTmp = strTmp & " AND b.סԺ�� = [5] "
        End If

        If Trim(Split(mstrCondition, ";")(4)) <> "" Then
            strTmp = strTmp & " AND b.��ǰ���� = [6] "
        End If

        If Trim(Split(mstrCondition, ";")(5)) <> "" Then
            strTmp = strTmp & " AND b.����� = [7] "
        End If

        If Val(Trim(Split(mstrCondition, ";")(7))) > 0 Then
            strTmp = strTmp & " AND a.������ĿID = [8] "
        End If

        gstrSQL = "Select   Decode(e.ID,Null,-100,e.ID) As ID,Decode(e.����״̬,1,'���',2,'����',3,'����',4,'���','����') As ͼ��,a.Id As ҽ��id," & vbNewLine & _
                    "       Decode(a.������־,1,'����','') As ������־," & vbNewLine & _
                    "       DECODE(a.������Դ,1,'����',2,'סԺ',4,'���','����') AS ������Դ," & vbNewLine & _
                    "       Decode(a.������Ŀid,Null,a.ҽ������,f.����) As ҽ������," & vbNewLine & _
                    "       a.����ʱ��," & vbNewLine & _
                    "       b.����," & vbNewLine & _
                    "       b.�����," & vbNewLine & _
                    "       b.סԺ��,b.��ǰ���� As ����," & vbNewLine & _
                    "       c.���� As ���˿���," & vbNewLine & _
                    "       d.���� As ��������," & vbNewLine & _
                    "       a.����ҽ�� As ������," & vbNewLine & _
                    "       a.ҽ��״̬," & vbNewLine & _
                    "       a.����id," & vbNewLine & _
                    "       a.��ҳid," & vbNewLine & _
                    "       a.������Ŀid," & vbNewLine & _
                    "       e.����״̬,g.���ͺ�,g.ִ��״̬,a.�Һŵ�,0 As ״̬,b.��Ժʱ�� As ��Ժ����,b.��ǰ����id,b.��ǰ����id "
        gstrSQL = gstrSQL & _
                    "From ����ҽ����¼ a," & vbNewLine & _
                    "     ������Ϣ b," & vbNewLine & _
                    "     ���ű� c," & vbNewLine & _
                    "     ���ű� d," & vbNewLine & _
                    "     ����������¼ e,������ĿĿ¼ f,����ҽ������ g " & vbNewLine & _
                    "Where (a.�������='F' Or a.������� Is Null)" & vbNewLine & _
                    "      And a.���id Is Null" & vbNewLine & _
                    "      And a.ҽ��״̬<>4 " & vbNewLine & strTmp & _
                    "      And a.ִ�п���id+0=[1]" & vbNewLine & _
                    "      And b.����id=a.����id" & vbNewLine & _
                    "      And c.Id=a.���˿���id" & vbNewLine & _
                    "      And d.Id=a.��������id And f.Id(+)=a.������Ŀid " & vbNewLine & _
                    "      And a.Id=e.ҽ��id(+)  And Nvl(e.����״̬,0)<=2 And a.ID=g.ҽ��id(+) "


        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��������id, _
                                                                CDate(Split(mstrCondition, ";")(0)), _
                                                                CDate(Split(mstrCondition, ";")(1)), _
                                                                "%" & Trim(Split(mstrCondition, ";")(2)) & "%", _
                                                                IIf(IsNumeric(Split(mstrCondition, ";")(3)), Split(mstrCondition, ";")(3), "0"), _
                                                                CStr(Trim(Split(mstrCondition, ";")(4))), _
                                                                IIf(IsNumeric(Split(mstrCondition, ";")(5)), Split(mstrCondition, ";")(5), "0"), _
                                                                Val(Trim(Split(mstrCondition, ";")(7))))
        If rs.BOF = False Then Call mclsVsf(0).LoadGrid(rs)

    '------------------------------------------------------------------------------------------------------------------
    Case "װ����������"

        mclsVsf(1).ClearGrid

        gstrSQL = GetPublicSQL(SQL.����������¼, mstrCondition)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��������id, _
                                                                "%" & Trim(Split(mstrCondition, ";")(2)) & "%", _
                                                                IIf(IsNumeric(Split(mstrCondition, ";")(3)), Split(mstrCondition, ";")(3), "0"), _
                                                                CStr(Trim(Split(mstrCondition, ";")(4))), _
                                                                IIf(IsNumeric(Split(mstrCondition, ";")(5)), Split(mstrCondition, ";")(5), "0"), _
                                                                Val(Trim(Split(mstrCondition, ";")(7))), CStr(Trim(Split(mstrCondition, ";")(11))))
        If rs.BOF = False Then Call mclsVsf(1).LoadGrid(rs)

    '------------------------------------------------------------------------------------------------------------------
    Case "װ����������"

        mclsVsf(2).ClearGrid

        gstrSQL = GetPublicSQL(SQL.���������¼, mstrCondition)

        If Val(Trim(Split(mstrCondition, ";")(10))) = 1 Then
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��������id, _
                                                                    4, _
                                                                    CDate(Split(mstrCondition, ";")(8)), _
                                                                    CDate(Split(mstrCondition, ";")(9)), _
                                                                    "%" & Trim(Split(mstrCondition, ";")(2)) & "%", _
                                                                    IIf(IsNumeric(Split(mstrCondition, ";")(3)), Split(mstrCondition, ";")(3), "0"), _
                                                                    CStr(Trim(Split(mstrCondition, ";")(4))), _
                                                                    IIf(IsNumeric(Split(mstrCondition, ";")(5)), Split(mstrCondition, ";")(5), "0"), _
                                                                    Val(Trim(Split(mstrCondition, ";")(7))))
        Else
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��������id, _
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
    Case "ˢ����������"

        Call ExecuteCommand("��ȡ������Ҫ")
        Call ExecuteCommand("��ȡ������ҩ")
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("��ȡ����ҽ��")
        Call ExecuteCommand("��ȡ��������")

    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ������Ҫ"

        With vsf(mintIndex)
            Call mfrmChildStationOutLine.RefreshData(Val(.RowData(.Row)), (Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3 And IsPrivs(mstrPrivs, "��Ҫ�Ǽ�")), mlng��������id)
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ������ҩ"

        With vsf(mintIndex)
            Call mfrmChildStationDrug.RefreshData(Val(.RowData(.Row)), _
                                        (Val(.TextMatrix(.Row, .ColIndex("����״̬"))) > 1 And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) < 4), _
                                        IIf(Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 2, 1, 2), _
                                        .TextMatrix(.Row, .ColIndex("������Դ")), _
                                        Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), _
                                        mstrPrivs)
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"
        With vsf(mintIndex)
            Call mfrmChildStationMaterial.RefreshData(Val(.RowData(.Row)), (Val(.TextMatrix(.Row, .ColIndex("����״̬"))) > 1 And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) < 4), IIf(Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 2, 1, 2), .TextMatrix(.Row, .ColIndex("������Դ")), _
                                        Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), _
                                        mstrPrivs)
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"
        With vsf(mintIndex)
            Call mfrmChildStationCure.RefreshData(Val(.RowData(.Row)), (Val(.TextMatrix(.Row, .ColIndex("����״̬"))) > 1 And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) < 4), IIf(Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 2, 1, 2), .TextMatrix(.Row, .ColIndex("������Դ")), _
                                        Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), _
                                        mstrPrivs)
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("���ͺ�"))) > 0 Then
                Call mclsExpenses.zlRefresh(mlng��������id, Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), Val(.TextMatrix(.Row, .ColIndex("���ͺ�"))), False)
            Else
                Call mclsExpenses.zlRefresh(0, 0, 0)
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����ҽ��"

        With vsf(mintIndex)
            Select Case .TextMatrix(.Row, .ColIndex("������Դ"))
            Case "����"
                If Val(.TextMatrix(.Row, .ColIndex("����id"))) > 0 Then
                    Call mclsOutAdvices.zlRefresh(Val(.TextMatrix(.Row, .ColIndex("����id"))), .TextMatrix(.Row, .ColIndex("�Һŵ�")), Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3, False, 0)
                Else
                    Call mclsOutAdvices.zlRefresh(0, "", False)
                End If
            Case "סԺ"
                If Val(.TextMatrix(.Row, .ColIndex("����id"))) > 0 Then

                    If .TextMatrix(.Row, .ColIndex("��Ժ����")) = "" Then
                        If Val(.TextMatrix(.Row, .ColIndex("״̬"))) <> 2 Then
                            int���� = 0 '��Ժ
                        Else
                            int���� = 1 'Ԥ��Ժ
                        End If
                    Else
                        int���� = 2 '��Ժ
                    End If
                    Call mclsInAdvices.zlRefresh(Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), Val(.TextMatrix(.Row, .ColIndex("��ǰ����id"))), Val(.TextMatrix(.Row, .ColIndex("��ǰ����id"))), IIf(Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 3, 4, int����), False, 0, Val(.TextMatrix(.Row, .ColIndex("ִ��״̬"))))
                Else
                    Call mclsInAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
                End If
            End Select
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"

        With vsf(mintIndex)
            Select Case .TextMatrix(.Row, .ColIndex("������Դ"))
            Case "����"

                gstrSQL = "Select ID From ���˹Һż�¼ Where No=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .TextMatrix(.Row, .ColIndex("�Һŵ�")))
                If rs.BOF = False Then
                    Call mfrmChildStationOutEPR.RefreshData(Val(.RowData(.Row)), Val(.TextMatrix(.Row, .ColIndex("����id"))), _
                                                            Val(rs("ID").Value), _
                                                            mlng��������id, _
                                                            Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), _
                                                            (Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3))
                Else
                    Call mfrmChildStationOutEPR.RefreshData(0, 0, 0, 0, 0, (Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3))
                End If

            Case "סԺ"

                Call mfrmChildStationInEPR.RefreshData(Val(.RowData(.Row)), Val(.TextMatrix(.Row, .ColIndex("����id"))), _
                                                        Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), _
                                                        mlng��������id, _
                                                        Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), _
                                                        (Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3))
            End Select
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "�����������"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 0 Then GoTo errEnd
            ExecuteCommand = frmOpsStationAuditing.ShowEdit(Me, Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))))
        End With

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ���������"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 1 Then GoTo errEnd

            If MsgBox("���Ҫ������" & .TextMatrix(.Row, .ColIndex("��������")) & "�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "ZL_������ϼ�¼_DELETE2(" & Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) & ",8)"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            gstrSQL = "zl_����������¼_AduitCancel(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "������������"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 1 Then GoTo errEnd
            ExecuteCommand = frmOpsStationArrange.ShowEdit(Me, Val(.RowData(.Row)), mlng��������id)
        End With

        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ����������"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 2 Then GoTo errEnd

            If MsgBox("���Ҫ������" & .TextMatrix(.Row, .ColIndex("��������")) & "���İ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_����������¼_ArrangeCancel(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "��������"

        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Or mintIndex <> 0 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 2 Then GoTo errEnd

            If MsgBox("���Ҫ������" & .TextMatrix(.Row, .ColIndex("��������")) & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_����������¼_Register(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ������"
        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Or mintIndex <> 1 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 3 Then GoTo errEnd

            If MsgBox("���Ҫȡ��������" & .TextMatrix(.Row, .ColIndex("��������")) & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_����������¼_RegisterCancel(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Or mintIndex <> 1 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 3 Then GoTo errEnd

            If MsgBox("���Ҫ��ɡ�" & .TextMatrix(.Row, .ColIndex("��������")) & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_����������¼_Complete(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ�����"
        With vsf(mintIndex)
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Or mintIndex <> 2 Then GoTo errEnd
            If Val(.TextMatrix(.Row, .ColIndex("����״̬"))) <> 4 Then GoTo errEnd

            If MsgBox("���Ҫȡ����ɡ�" & .TextMatrix(.Row, .ColIndex("��������")) & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo errEnd

            gstrSQL = "zl_����������¼_CompleteCancel(" & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)

            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End With

        GoTo errEnd
        
'    '------------------------------------------------------------------------------------------------------------------
'    Case "���ɷ����շѵ�", "���ɷ��ü��ʵ�", "���ɷ�����ѵ�"
'
''        Dim blnZero As Boolean
''        Dim intTmp As Integer
'
'        With vsf(mintIndex)
'            gstrSQL = GetPublicSQL(SQL.��������ѡ��)
'
'            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(.RowData(.Row)))
'
'            If ShowPubSelect(Me, Nothing, 3, "����,900,0,;����,2400,0,;���,900,0,;����,1200,2,;��λ,810,0,", Me.Name & "\��������ѡ��", "�����������б���ѡ���������òο�", rsData, rs, 8790, 4500, False, , , True) = 1 Then
'
'                blnZero = False
'                Select Case strCommand
'                Case "���ɷ����շѵ�"
'                    intTmp = 1
'                Case "���ɷ��ü��ʵ�"
'                    intTmp = 2
'                Case "���ɷ�����ѵ�"
'                    intTmp = 2
'                    blnZero = True
'                End Select
'
'                If frmOpsStationCharge.ShowEdit(Me, Val(rs("ID").Value), Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), mstrPrivs, intTmp, blnZero, IIf(.TextMatrix(.Row, .ColIndex("������Դ")) = "סԺ", 2, 1)) Then
'
'                    'ˢ�·����б�
'                    If Not (mclsExpenses Is Nothing) Then
'                        Call mclsExpenses.zlRefresh(mlng��������id, Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), Val(.TextMatrix(.Row, .ColIndex("���ͺ�"))), False)
'                    End If
'
'                End If
'            End If
'        End With
'
'        GoTo errEnd
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��λ����"
        
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
    Case "�ָ�����"

        '1.�ָ���Ҫ��¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationOutLine.DataChanged Then
            If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 And vsf(mintIndex).Rows > 2 Then
                vsf(mintIndex).Rows = vsf(mintIndex).Rows - 1
                vsf(mintIndex).Row = vsf(mintIndex).Rows - 1
            End If

            Call ExecuteCommand("��ȡ������Ҫ")
            mfrmChildStationOutLine.DataChanged = False
        End If

        '2.�ָ���ҩ��¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationDrug.DataChanged Then
            Call ExecuteCommand("��ȡ������ҩ")
            mfrmChildStationDrug.DataChanged = False
        End If

        '3.�ָ����ϼ�¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationMaterial.DataChanged Then
            Call ExecuteCommand("��ȡ��������")
            mfrmChildStationMaterial.DataChanged = False
        End If

        '4.�ָ����Ƽ�¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationCure.DataChanged Then
            Call ExecuteCommand("��ȡ��������")
            mfrmChildStationCure.DataChanged = False
        End If
        
'        mblnNew = False
    '------------------------------------------------------------------------------------------------------------------
    Case "У������"

        '1.У���Ҫ��¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationOutLine.DataChanged Then
            If mfrmChildStationOutLine.ValidData = False Then GoTo errEnd
        End If

        '2.У����ҩ��¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationDrug.DataChanged Then
            If mfrmChildStationDrug.ValidData = False Then GoTo errEnd
        End If

        '3.У����ϼ�¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationMaterial.DataChanged Then
            If mfrmChildStationMaterial.ValidData = False Then GoTo errEnd
        End If

        '4.У�����Ƽ�¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationCure.DataChanged Then
            If mfrmChildStationCure.ValidData = False Then GoTo errEnd
        End If
        
        ExecuteCommand = True

        GoTo errEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"

        mlngTmp = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))

        '1.�����Ҫ��¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationOutLine.DataChanged Then

            If mfrmChildStationOutLine.SaveData(rsSQL) = False Then GoTo errEnd

        End If

        '2.������ҩ��¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationDrug.DataChanged Then
            If mfrmChildStationDrug.SaveData(rsSQL) = False Then GoTo errEnd
        End If

        '3.������ϼ�¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationMaterial.DataChanged Then
            If mfrmChildStationMaterial.SaveData(rsSQL) = False Then GoTo errEnd
        End If

        '4.�������Ƽ�¼
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildStationCure.DataChanged Then
            If mfrmChildStationCure.SaveData(rsSQL) = False Then GoTo errEnd
        End If
        
        ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)

        GoTo errEnd

    '------------------------------------------------------------------------------------------------------------------
    Case "ǰһ��"
        With vsf(mintIndex)
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
                Call vsf_AfterRowColChange(mintIndex, 0, 0, .Row, .Col)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "��һ��"
        With vsf(mintIndex)
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
                Call vsf_AfterRowColChange(mintIndex, 0, 0, .Row, .Col)
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "��ע���"

        If Val(GetRegister(˽��ȫ��, "", "ʹ�ø��Ի����", "0")) = 1 Then
            'ʹ�ø��Ի�����

            dkpMain.LoadStateFromString GetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, "")

            mstrFindKey = Trim(GetRegister(˽��ģ��, Me.Name, "��λ����", "����"))
            mclsVsf(0).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(0)) & "A", ""))
            mclsVsf(1).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(1)) & "B", ""))
            mclsVsf(2).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(2)) & "C", ""))
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "дע���"

        If Val(GetRegister(˽��ȫ��, "", "ʹ�ø��Ի����", "0")) = 1 Then
            'ʹ�ø��Ի�����
            Call SetRegister(˽��ģ��, Me.Name, "��λ����", mstrFindKey)
        End If
        Call SetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        Call SetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(0)) & "A", mclsVsf(0).SaveStateToString)
        Call SetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(1)) & "B", mclsVsf(1).SaveStateToString)
        Call SetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(2)) & "C", mclsVsf(2).SaveStateToString)

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

'�ؼ������¼�
'######################################################################################################################

'Private Sub cbsDept_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'
'    Select Case Control.ID
'    Case conMenu_View_Dept
'
'        If mlng��������id <> mobjPeisDept.ItemData(mobjPeisDept.ListIndex) Then
'            mlng��������id = mobjPeisDept.ItemData(mobjPeisDept.ListIndex)
'            mstr�������� = zlCommFun.GetNeedName(mobjPeisDept.List(mobjPeisDept.ListIndex))
'            Call ExecuteCommand("ˢ������")
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
'    '���������ؼ�Resize����
'    tbc.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
'
'End Sub

Private Sub cboDept_Click()

    If mlng��������id <> cboDept.ItemData(cboDept.ListIndex) Then
        mlng��������id = cboDept.ItemData(cboDept.ListIndex)
        mstr�������� = zlCommFun.GetNeedName(cboDept.List(cboDept.ListIndex))
        Call ExecuteCommand("ˢ������")
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
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Then Exit Sub
            ReportOpen gcnOracle, ParamInfo.ϵͳ��, "ZL1_BILL_1804", Me, "ҽ��id=" & Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), 1
        End With
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BillPrint
        
        With vsf(mintIndex)
            If mintIndex <> 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) = 0 Then Exit Sub
            ReportOpen gcnOracle, ParamInfo.ϵͳ��, "ZL1_BILL_1804", Me, "ҽ��id=" & Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))), 2
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Parameter                                                         '��������

        If ExecuteCommand("���ز�������") Then
            AutoRefresh = True
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option

        Call ExecuteCommand("����������")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Request

        If ExecuteCommand("ֱ����������") Then
            Call ExecuteCommand("װ�صȴ�����", "ˢ��������Ϣ")
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Audit
        If ExecuteCommand("�����������") Then
            Call ExecuteCommand("װ�صȴ�����")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_UnAudit
        If ExecuteCommand("ȡ���������") Then
            Call ExecuteCommand("װ�صȴ�����")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Arrange
        If ExecuteCommand("������������") Then
            Call ExecuteCommand("װ�صȴ�����")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_UnArrange
        If ExecuteCommand("ȡ����������") Then
            Call ExecuteCommand("װ�صȴ�����")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Plan
        If ExecuteCommand("��������") Then
            Call ExecuteCommand("װ�صȴ�����")
            Call ExecuteCommand("װ����������")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Logout
        If ExecuteCommand("ȡ������") Then
            Call ExecuteCommand("װ�صȴ�����")
            Call ExecuteCommand("װ����������")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Complete
        If ExecuteCommand("�������") Then
            Call ExecuteCommand("װ����������")
            Call ExecuteCommand("װ����������")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Undone
        If ExecuteCommand("ȡ�����") Then
            Call ExecuteCommand("װ����������")
            Call ExecuteCommand("װ����������")
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save                  '������������

        If ExecuteCommand("У������") And DataChanged Then
            If ExecuteCommand("��������") Then
                DataChanged = False
                Call ExecuteCommand("ˢ��ָ������")
'                mblnNew = False
            End If
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle                  '�ָ���������

        Call ExecuteCommand("�ָ�����")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        If frmOpsStationFilter.ShowSearch(Me, mstrCondition, mlng��������id) Then
            Call ExecuteCommand("ˢ������")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh

        Call ExecuteCommand("ˢ������")
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

        Call CommandBarExecutePublic(Control, Me, vsf(mintIndex), "���������嵥")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward

        Call ExecuteCommand("ǰһ��")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward

        Call ExecuteCommand("��һ��")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem

        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
        
        Call ExecuteCommand("��λ����", txtLocation.Text)
        
    '--------------------------------------------------------------------------------------------------------------
    Case Else

        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            'Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
        Else
            Select Case tbcPage.Selected.Tag
            Case "����"
                
                Select Case Control.ID
                '------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2# + 1
                
                    Call ExecuteCommand("���ɷ����շѵ�")
            
                '------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2# + 2
            
                    Call ExecuteCommand("���ɷ��ü��ʵ�")
                    
                '------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2# + 3
                            
                    Call ExecuteCommand("���ɷ�����ѵ�")
                Case Else
                    Call mclsExpenses.zlExecuteCommandBars(Control)
                End Select
                
            Case "סԺҽ��"
                Call mclsInAdvices.zlExecuteCommandBars(Control)
            Case "����ҽ��"
                Call mclsOutAdvices.zlExecuteCommandBars(Control)
            Case "סԺ����"
                Call mfrmChildStationInEPR.zlExecuteCommandBars(Control)
            Case "���ﲡ��"
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
    Case "����"
        Call mclsExpenses.zlPopupCommandBars(CommandBar)
        
        If CommandBar.Parent.ID = conMenu_Edit_Preferences Then
    
            With CommandBar.Controls
    
                .DeleteAll
    
                Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 1, "�����շѵ���(&1)")
                Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 2, "���ɼ��ʵ���(&2)")
                Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 3, "������ķ���(&3)")
                With cbsMain.KeyBindings
                    .Add FCONTROL, vbKeyN, conMenu_Edit_MakeCharge * 2 + 1
                    .Add FCONTROL, vbKeyB, conMenu_Edit_MakeCharge * 2 + 2
                End With
                
            End With
        End If
        
    Case "����ҽ��"
        Call mclsOutAdvices.zlPopupCommandBars(CommandBar)
    Case "סԺҽ��"
        Call mclsInAdvices.zlPopupCommandBars(CommandBar)
    Case "���ﲡ��"
        Call mfrmChildStationOutEPR.zlPopupCommandBars(CommandBar)
    Case "סԺ����"
        Call mfrmChildStationInEPR.zlPopupCommandBars(CommandBar)
    End Select

End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo errHand

    With vsf(mintIndex)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel

            Control.Enabled = (Val(.RowData(.Row)) > 0)
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_BillPrintView, conMenu_File_BillPrint
            
            With vsf(mintIndex)
                
                Control.Visible = IsPrivs(mstrPrivs, "����֪ͨ��")
                
                Control.Enabled = (Val(.TextMatrix(.Row, .ColIndex("ҽ��id"))) > 0 And mintIndex = 0 And Control.Visible And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) >= 2)
                
            End With
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Option
            Control.Visible = IsPrivs(mstrPrivs, "ִ�м�����")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_ManagePopup

            Control.Visible = IsPrivs(mstrPrivs, "�������") Or _
                                IsPrivs(mstrPrivs, "���ȡ��") Or _
                                IsPrivs(mstrPrivs, "��������") Or _
                                IsPrivs(mstrPrivs, "����ȡ��") Or _
                                IsPrivs(mstrPrivs, "ִ�б���") Or _
                                IsPrivs(mstrPrivs, "����ȡ��") Or _
                                IsPrivs(mstrPrivs, "�������") Or _
                                IsPrivs(mstrPrivs, "��¼����") Or _
                                IsPrivs(mstrPrivs, "���ȡ��")

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Request            '��¼����
        
            Control.Visible = IsPrivs(mstrPrivs, "��¼����")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Audit            '���

            Control.Visible = IsPrivs(mstrPrivs, "�������")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 0 And Val(.RowData(.Row)) <> 0 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_UnAudit            '���ȡ��

            Control.Visible = IsPrivs(mstrPrivs, "���ȡ��")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 1 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Arrange               '��������
'
            Control.Visible = IsPrivs(mstrPrivs, "��������")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 1 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_UnArrange             '����ȡ��

            Control.Visible = IsPrivs(mstrPrivs, "����ȡ��")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 2 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Plan                '����

            Control.Visible = IsPrivs(mstrPrivs, "ִ�б���")
            Control.Enabled = (mintIndex = 0 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 2 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Complete            '�������

            Control.Visible = IsPrivs(mstrPrivs, "�������")
            Control.Enabled = (mintIndex = 1 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Logout              '����ȡ��

            Control.Visible = IsPrivs(mstrPrivs, "����ȡ��")
            Control.Enabled = (mintIndex = 1 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3 And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Undone                            '���ȡ��

            Control.Visible = IsPrivs(mstrPrivs, "���ȡ��")
            Control.Enabled = (mintIndex = 2 And DataChanged = False And Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 4 And Control.Visible)

        '------------------------------------------------------------------------------------------------------------------
        Case conMenu_EditPopup

            Control.Visible = IsPrivs(mstrPrivs, "��ҩ׼��") Or _
                                IsPrivs(mstrPrivs, "����׼��") Or _
                                IsPrivs(mstrPrivs, "���ϵǼ�") Or _
                                IsPrivs(mstrPrivs, "��Ҫ�Ǽ�") Or _
                                IsPrivs(mstrPrivs, "��ҩ�Ǽ�")
            
        '------------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle

            Control.Visible = IsPrivs(mstrPrivs, "��ҩ׼��") Or _
                                IsPrivs(mstrPrivs, "����׼��") Or _
                                IsPrivs(mstrPrivs, "���ϵǼ�") Or _
                                IsPrivs(mstrPrivs, "��Ҫ�Ǽ�") Or _
                                IsPrivs(mstrPrivs, "��ҩ�Ǽ�")
            Control.Enabled = (DataChanged And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Filter, conMenu_View_Refresh, conMenu_Edit_Request
            Control.Enabled = (DataChanged = False And Control.Visible)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button            '������
            If cbsMain.Count >= 2 Then
                Control.Checked = cbsMain(2).Visible
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Text              'ͼ������
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Size              '��ͼ��
            Control.Checked = cbsMain.Options.LargeIcons
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_StatusBar                 '״̬��
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
            Case "����"
                
                Select Case Control.ID
                '--------------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_Preferences
                    
                    Control.Visible = IsPrivs(mstrPrivs, "��¼����")
                    Control.Enabled = Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3 And Control.Visible
                    
                '--------------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2 + 1
                    
                    Control.Visible = (.TextMatrix(.Row, .ColIndex("������Դ")) = "����" And IsPrivs(mstrPrivs, "��¼����"))
                    Control.Enabled = Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3 And Control.Visible
                    
                '--------------------------------------------------------------------------------------------------------------
                Case conMenu_Edit_MakeCharge * 2 + 2, conMenu_Edit_MakeCharge * 2 + 3
                    
                    Control.Visible = IsPrivs(mstrPrivs, "��¼����")
                    Control.Enabled = Val(.TextMatrix(.Row, .ColIndex("����״̬"))) = 3 And Control.Visible
                '--------------------------------------------------------------------------------------------------------------
                Case Else
                    Call mclsExpenses.zlUpdateCommandBars(Control)
                End Select

            Case "סԺҽ��"
                Call mclsInAdvices.zlUpdateCommandBars(Control)
            Case "����ҽ��"
                Call mclsOutAdvices.zlUpdateCommandBars(Control)
            Case "סԺ����"
                Call mfrmChildStationInEPR.zlUpdateCommandBars(Control)
            Case "���ﲡ��"
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

    If ExecuteCommand("��ʼ����") = False Then GoTo errHand

    Call ExecuteCommand("ˢ������")
    
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

    mstrPrivs = UserInfo.ģ��Ȩ��
'    mstrPrivs = "����������"

    Call ExecuteCommand("��ʼ�ؼ�")
    Call ExecuteCommand("��ע���")

    Call RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, ParamInfo.ϵͳ��, ParamInfo.ģ���, UserInfo.ģ��Ȩ��)

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

    Call ExecuteCommand("дע���")

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
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildStationDrug_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildStationMaterial_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildStationOutLine_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
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

    If Item.Tag = "" Then Exit Sub '��ʼ��ʱ,��û��ֵ

    Call SubWinDefCommandBar(Item)

End Sub

Private Sub tmr_Timer()
    Dim strSvrKey As String
    
    mlngCountTmr = mlngCountTmr + 1
    
    If mlngCountTmr >= Val(tmr.Tag) Then
    
        'ʱ�䵽�ˣ���ʼ����
        mlngCountTmr = 0
                
        Call ExecuteCommand("ˢ������")
                
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
        blnChange = ExchangeAdvice(.TextMatrix(NewRow, .ColIndex("������Դ")) = "����")
        blnChange = blnChange Or ExchangeEPRS(.TextMatrix(NewRow, .ColIndex("������Դ")) = "����")
    End With

    Call ExecuteCommand("��ȡ������Ϣ")
    Call ExecuteCommand("ˢ����������")

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
    Case 2          '�����˵�����
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        cbrPopupBar.ShowPopup
    End Select
End Sub
