VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Object = "*\A..\ZlPatiAddress\ZlPatiAddress.vbp"
Begin VB.Form frmPatiCureCardEdit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人发卡管理"
   ClientHeight    =   9960
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   11625
   Icon            =   "frmPatiCureCardEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCertificate 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   10710
      ScaleHeight     =   3105
      ScaleWidth      =   5925
      TabIndex        =   164
      Top             =   7380
      Width           =   5925
      Begin VSFlex8Ctl.VSFlexGrid vsCertificate 
         Height          =   3015
         Left            =   15
         TabIndex        =   165
         Top             =   0
         Width           =   5895
         _cx             =   10398
         _cy             =   5318
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
         BackColorBkg    =   -2147483643
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picDrugAllergy 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   11820
      ScaleHeight     =   3255
      ScaleWidth      =   6840
      TabIndex        =   158
      Top             =   1200
      Width           =   6840
      Begin VB.CommandButton cmdSelDrug 
         Caption         =   "…"
         Height          =   300
         Left            =   600
         TabIndex        =   159
         Top             =   540
         Visible         =   0   'False
         Width           =   300
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDrug 
         Height          =   3060
         Left            =   -30
         TabIndex        =   160
         Top             =   240
         Width           =   5895
         _cx             =   10398
         _cy             =   5397
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
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
   Begin VB.PictureBox picTaskPanelOther 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   10350
      ScaleHeight     =   1125
      ScaleWidth      =   1215
      TabIndex        =   156
      Top             =   2910
      Visible         =   0   'False
      Width           =   1215
      Begin XtremeSuiteControls.TaskPanel wndTaskPanelOther 
         Height          =   945
         Left            =   0
         TabIndex        =   157
         Top             =   0
         Width           =   1035
         _Version        =   589884
         _ExtentX        =   1826
         _ExtentY        =   1667
         _StockProps     =   64
         VisualTheme     =   7
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin VB.PictureBox pic预交余额 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   8000
      ScaleHeight     =   225
      ScaleWidth      =   2205
      TabIndex        =   154
      Top             =   7380
      Visible         =   0   'False
      Width           =   2200
      Begin VB.Label lbl预交余额 
         Caption         =   "预交余额:0元"
         ForeColor       =   &H000000FF&
         Height          =   220
         Left            =   0
         TabIndex        =   155
         Top             =   0
         Width           =   2200
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   106
      Top             =   9600
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   176
            Picture         =   "frmPatiCureCardEdit.frx":000C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15531
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   10680
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOtherInfo 
      BorderStyle     =   0  'None
      Height          =   4080
      Left            =   10500
      ScaleHeight     =   4080
      ScaleWidth      =   10110
      TabIndex        =   130
      Top             =   4620
      Width           =   10110
      Begin VB.CommandButton cmdMedicalWarning 
         Caption         =   "…"
         Height          =   330
         Left            =   9465
         TabIndex        =   147
         Top             =   0
         Width           =   315
      End
      Begin VB.ComboBox cboBloodType 
         Height          =   300
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   0
         Width           =   1410
      End
      Begin VB.ComboBox cboBH 
         Height          =   300
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   0
         Width           =   1410
      End
      Begin VB.TextBox txtMedicalWarning 
         Height          =   350
         Left            =   5535
         Locked          =   -1  'True
         TabIndex        =   134
         Top             =   0
         Width           =   4260
      End
      Begin VB.TextBox txtOtherWaring 
         Height          =   350
         Left            =   1275
         MaxLength       =   100
         TabIndex        =   133
         Top             =   375
         Width           =   8535
      End
      Begin VB.Frame frameLinkMan 
         BackColor       =   &H80000004&
         Height          =   105
         Left            =   1065
         TabIndex        =   132
         Top             =   840
         Width           =   8895
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         Height          =   105
         Left            =   885
         TabIndex        =   131
         Top             =   2160
         Width           =   9135
      End
      Begin VSFlex8Ctl.VSFlexGrid vsLinkMan 
         Height          =   975
         Left            =   60
         TabIndex        =   139
         Top             =   1080
         Width           =   9705
         _cx             =   17119
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
         BackColorSel    =   -2147483634
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
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
      Begin VSFlex8Ctl.VSFlexGrid vsOtherInfo 
         Height          =   1380
         Left            =   60
         TabIndex        =   140
         Top             =   2400
         Width           =   9705
         _cx             =   17119
         _cy             =   2434
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPatiCureCardEdit.frx":08A0
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
      Begin VB.Label lblBloodType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "血型"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   525
         TabIndex        =   146
         Top             =   45
         Width           =   1020
      End
      Begin VB.Label lblBH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2535
         TabIndex        =   145
         Top             =   45
         Width           =   885
      End
      Begin VB.Label lblMedicalWarning 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医学警示"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4215
         TabIndex        =   144
         Top             =   45
         Width           =   1860
      End
      Begin VB.Label lblOtherWaring 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "其他医学警示"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -270
         TabIndex        =   143
         Top             =   420
         Width           =   1860
      End
      Begin VB.Label lblLinkman 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "联系人信息"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -360
         TabIndex        =   142
         Top             =   840
         Width           =   1860
      End
      Begin VB.Label lblOtherInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "其他信息"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   -450
         TabIndex        =   141
         Top             =   2145
         Width           =   1860
      End
   End
   Begin VB.PictureBox picInoculate 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   120
      ScaleHeight     =   3105
      ScaleWidth      =   5925
      TabIndex        =   128
      Top             =   9240
      Width           =   5925
      Begin VSFlex8Ctl.VSFlexGrid vsInoculate 
         Height          =   3015
         Left            =   540
         TabIndex        =   129
         Top             =   210
         Width           =   5895
         _cx             =   10398
         _cy             =   5318
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
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
   Begin VB.CommandButton cmd余额退款 
      Caption         =   "退款(&B)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10365
      TabIndex        =   119
      Top             =   1845
      Width           =   1100
   End
   Begin VB.CommandButton cmd充值 
      Caption         =   "充值(&Z)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10365
      TabIndex        =   118
      Top             =   1425
      Width           =   1100
   End
   Begin VB.PictureBox picTittle 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   150
      ScaleHeight     =   495
      ScaleWidth      =   9945
      TabIndex        =   107
      Top             =   240
      Width           =   9945
      Begin VB.TextBox txtFact 
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
         IMEMode         =   3  'DISABLE
         Left            =   5370
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   60
         Width           =   1575
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9405
         Style           =   1  'Graphical
         TabIndex        =   112
         TabStop         =   0   'False
         ToolTipText     =   "热键：F8"
         Top             =   15
         Width           =   405
      End
      Begin VB.Frame fraSplit 
         Caption         =   "Frame1"
         Height          =   150
         Left            =   -750
         TabIndex        =   108
         Top             =   345
         Width           =   12990
      End
      Begin VB.ComboBox cboNO 
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
         Left            =   7770
         Locked          =   -1  'True
         TabIndex        =   114
         ToolTipText     =   "热键:F12"
         Top             =   45
         Width           =   1620
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票号"
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
         Left            =   4860
         TabIndex        =   168
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -645
         TabIndex        =   117
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
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
         Left            =   7080
         TabIndex        =   116
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.PictureBox picCard 
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   1635
      Left            =   90
      ScaleHeight     =   1635
      ScaleWidth      =   9975
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   7650
      Width           =   9975
      Begin VB.Frame fraCard 
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
         Height          =   1545
         Left            =   30
         TabIndex        =   153
         Top             =   30
         Width           =   9795
         Begin VB.TextBox txt余额 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6420
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   111
            Top             =   1110
            Width           =   3210
         End
         Begin VB.TextBox txt合计 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   8340
            MaxLength       =   16
            TabIndex        =   91
            Tag             =   "合计"
            Top             =   650
            Width           =   1260
         End
         Begin VB.TextBox txt病历费 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   3090
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   87
            TabStop         =   0   'False
            Tag             =   "病历费"
            Top             =   660
            Width           =   705
         End
         Begin VB.CheckBox chk病历费 
            Caption         =   "收病历费"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1980
            TabIndex        =   86
            Top             =   690
            Width           =   1140
         End
         Begin VB.CommandButton cmdReadCard 
            Caption         =   "读卡"
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
            Left            =   4845
            TabIndex        =   75
            TabStop         =   0   'False
            Tag             =   "读卡"
            Top             =   215
            Width           =   615
         End
         Begin VB.TextBox txt卡号 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   74
            Tag             =   "卡号"
            Top             =   205
            Width           =   3780
         End
         Begin VB.TextBox txt卡费 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   1095
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   85
            Tag             =   "卡费"
            Top             =   660
            Width           =   800
         End
         Begin VB.TextBox txtAudi 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   8355
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   81
            Tag             =   "验证"
            Top             =   205
            Width           =   1260
         End
         Begin VB.CheckBox chk记帐 
            Caption         =   "记帐"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3870
            TabIndex        =   88
            Top             =   690
            Width           =   885
         End
         Begin VB.TextBox txtPass 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6420
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   79
            Tag             =   "密码"
            Top             =   205
            Width           =   1125
         End
         Begin VB.ComboBox cbo支付方式 
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
            Left            =   6420
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   660
            Width           =   1935
         End
         Begin VB.TextBox txt操作员 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1085
            Locked          =   -1  'True
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   1100
            Width           =   1080
         End
         Begin VB.TextBox txt变动原因 
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
            Left            =   1100
            MaxLength       =   100
            TabIndex        =   83
            Tag             =   "变动原因"
            Top             =   660
            Visible         =   0   'False
            Width           =   8535
         End
         Begin VB.TextBox txt原卡密码 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6420
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   136
            Tag             =   "密码"
            Top             =   205
            Visible         =   0   'False
            Width           =   1125
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   3075
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   1095
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm"
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin zlIDKind.IDKindNew IDKindPay 
            Height          =   360
            Left            =   500
            TabIndex        =   151
            Top             =   200
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   $"frmPatiCureCardEdit.frx":0902
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
            OnlyReadCardNo  =   0   'False
            BackColor       =   -2147483633
         End
         Begin VB.TextBox txt刷卡卡号 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6420
            TabIndex        =   77
            Tag             =   "刷卡卡号"
            Top             =   210
            Width           =   3210
         End
         Begin VB.ComboBox cbo挂失方式 
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
            Left            =   6420
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   215
            Visible         =   0   'False
            Width           =   3225
         End
         Begin zlIDKind.IDKindNew IDKindPayMode 
            Height          =   360
            Left            =   5535
            TabIndex        =   166
            Top             =   1095
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
            ShowSortName    =   0   'False
            IDKindStr       =   "应收|应收|0|0|0|0|0|0|0|0|0;充值|充值|0|0|0|0|0|0|0|0|0"
            CaptionAlignment=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   10.5
            FontName        =   "宋体"
            IDKind          =   -1
            DefaultCardType =   "0"
            NotAutoAppendKind=   -1  'True
            OnlyReadCardNo  =   0   'False
            BackColor       =   -2147483633
         End
         Begin VB.Label lbl支付方式 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "缴款"
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
            Left            =   5925
            TabIndex        =   89
            Top             =   720
            Width           =   420
         End
         Begin VB.Label lbl卡号 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "卡号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   70
            TabIndex        =   73
            Top             =   260
            Width           =   450
         End
         Begin VB.Label lbl验证 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "验证"
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
            Left            =   7890
            TabIndex        =   80
            Top             =   270
            Width           =   420
         End
         Begin VB.Label lbl卡费 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "卡费"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   590
            TabIndex        =   84
            Top             =   720
            Width           =   420
         End
         Begin VB.Label lbl密码 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "密码"
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
            Left            =   5940
            TabIndex        =   78
            Top             =   270
            Width           =   420
         End
         Begin VB.Label lbl发卡人 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发卡人"
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
            Left            =   420
            TabIndex        =   115
            Top             =   1170
            Width           =   615
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发卡时间"
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
            Left            =   2205
            TabIndex        =   113
            Top             =   1170
            Width           =   840
         End
         Begin VB.Label lbl刷卡验证 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " 刷卡验证"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5430
            TabIndex        =   76
            Top             =   270
            Width           =   945
         End
         Begin VB.Label lbl原卡密码 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "原卡密码"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5520
            TabIndex        =   138
            Top             =   270
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin XtremeSuiteControls.TabControl tbPageDo 
         Height          =   240
         Left            =   180
         TabIndex        =   152
         Top             =   330
         Width           =   420
         _Version        =   589884
         _ExtentX        =   741
         _ExtentY        =   423
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picBasePati 
      BorderStyle     =   0  'None
      Height          =   2280
      Left            =   90
      ScaleHeight     =   2280
      ScaleWidth      =   9990
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   765
      Width           =   9990
      Begin VB.Frame fra 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Left            =   60
         TabIndex        =   98
         Top             =   -15
         Width           =   9840
         Begin VB.TextBox txt手机 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6120
            MaxLength       =   20
            TabIndex        =   9
            Tag             =   "手机号"
            Top             =   601
            Width           =   1590
         End
         Begin ZlPatiAddress.PatiAddress padd户口地址 
            Height          =   330
            Left            =   1170
            TabIndex        =   20
            Tag             =   "户口地址"
            Top             =   1830
            Visible         =   0   'False
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padd家庭地址 
            Height          =   330
            Left            =   1170
            TabIndex        =   17
            Tag             =   "现住址"
            Top             =   1425
            Visible         =   0   'False
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   100
         End
         Begin VB.CommandButton cmd户口地址 
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5610
            TabIndex        =   161
            TabStop         =   0   'False
            Tag             =   "户口地址"
            Top             =   1845
            Width           =   300
         End
         Begin VB.TextBox txt户口地址 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1170
            TabIndex        =   19
            Tag             =   "户口地址"
            Top             =   1830
            Width           =   4755
         End
         Begin VB.TextBox txt户口地址邮编 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6930
            MaxLength       =   6
            TabIndex        =   21
            Tag             =   "户口地址邮编"
            Top             =   1820
            Width           =   780
         End
         Begin VB.TextBox txt年龄 
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
            Left            =   3765
            TabIndex        =   13
            Text            =   "年龄"
            Top             =   1019
            Width           =   555
         End
         Begin VB.ComboBox cbo民族 
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
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   1260
         End
         Begin VB.TextBox txt家庭邮编 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6930
            MaxLength       =   6
            TabIndex        =   18
            Tag             =   "家庭地址邮编"
            Top             =   1418
            Width           =   780
         End
         Begin VB.PictureBox picPatient 
            Height          =   1500
            Left            =   7815
            ScaleHeight     =   1440
            ScaleWidth      =   1815
            TabIndex        =   127
            Top             =   180
            Width           =   1875
            Begin VB.Image imgPatient 
               Height          =   1425
               Left            =   15
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1800
            End
         End
         Begin VB.CommandButton cmdPicCollect 
            Caption         =   "采集"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   8445
            TabIndex        =   125
            Top             =   1710
            Width           =   600
         End
         Begin VB.CommandButton cmdPicFile 
            Caption         =   "文件"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   7815
            TabIndex        =   124
            Top             =   1710
            Width           =   585
         End
         Begin VB.CommandButton cmdPicClear 
            Caption         =   "清除"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   9090
            TabIndex        =   123
            Top             =   1710
            Width           =   600
         End
         Begin VB.TextBox txtPatient 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1170
            TabIndex        =   0
            Tag             =   "姓名"
            Top             =   180
            Width           =   1965
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   555
            TabIndex        =   121
            Top             =   180
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   $"frmPatiCureCardEdit.frx":0991
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
            DefaultCardType =   "0"
            OnlyReadCardNo  =   0   'False
            BackColor       =   -2147483633
         End
         Begin VB.TextBox txt门诊号 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6120
            MaxLength       =   18
            TabIndex        =   4
            Tag             =   "门诊号"
            Top             =   180
            Width           =   1590
         End
         Begin VB.ComboBox cbo年龄单位 
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
            Left            =   4305
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1019
            Width           =   690
         End
         Begin VB.ComboBox cbo性别 
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
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   200
            Width           =   1260
         End
         Begin VB.TextBox txt身份证号 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   1170
            MaxLength       =   18
            TabIndex        =   6
            Tag             =   "身份证号"
            Text            =   "012345678901234567"
            Top             =   601
            Width           =   1965
         End
         Begin VB.TextBox txt家庭电话 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   6120
            MaxLength       =   20
            TabIndex        =   15
            Tag             =   "家庭电话"
            Top             =   1012
            Width           =   1590
         End
         Begin VB.CommandButton cmd家庭地址 
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5610
            TabIndex        =   22
            TabStop         =   0   'False
            Tag             =   "现住址"
            ToolTipText     =   "热键：F3"
            Top             =   1443
            Width           =   300
         End
         Begin MSMask.MaskEdBox txt出生时间 
            Height          =   345
            Left            =   2280
            TabIndex        =   12
            Tag             =   "出生时间"
            Top             =   1012
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt出生日期 
            Height          =   345
            Left            =   1170
            TabIndex        =   11
            Tag             =   "出生日期"
            Top             =   1012
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "YYYY-MM-DD"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txt家庭地址 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   16
            Tag             =   "现住址"
            Top             =   1418
            Width           =   4755
         End
         Begin VB.Label lbl手机 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "手机号"
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
            Left            =   5415
            TabIndex        =   8
            Top             =   675
            Width           =   630
         End
         Begin VB.Label lbl户口地址邮编 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "户口邮编"
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
            Left            =   6060
            TabIndex        =   163
            Top             =   1890
            Width           =   840
         End
         Begin VB.Label lbl户口地址 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "户口地址"
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
            Left            =   315
            TabIndex        =   162
            Top             =   1890
            Width           =   840
         End
         Begin VB.Label lbl民族 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "民族"
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
            Left            =   3285
            TabIndex        =   149
            Top             =   671
            Width           =   420
         End
         Begin VB.Label lbl家庭邮编 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "邮编"
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
            Left            =   6480
            TabIndex        =   148
            Top             =   1488
            Width           =   420
         End
         Begin VB.Label lbl姓名 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名"
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
            Left            =   120
            TabIndex        =   126
            Top             =   255
            Width           =   420
         End
         Begin VB.Label lbl门诊号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊号"
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
            Left            =   5415
            TabIndex        =   3
            Top             =   255
            Width           =   630
         End
         Begin VB.Label lbl出生日期 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生日期"
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
            Left            =   315
            TabIndex        =   10
            Top             =   1079
            Width           =   840
         End
         Begin VB.Label lbl年龄 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄"
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
            Left            =   3300
            TabIndex        =   23
            Top             =   1079
            Width           =   420
         End
         Begin VB.Label lbl性别 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
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
            Left            =   3300
            TabIndex        =   1
            Top             =   255
            Width           =   420
         End
         Begin VB.Label lbl身份证号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份证号"
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
            Left            =   315
            TabIndex        =   5
            Top             =   671
            Width           =   840
         End
         Begin VB.Label lbl家庭电话 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家庭电话"
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
            Left            =   5205
            TabIndex        =   25
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lbl家庭地址 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "现住址"
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
            Left            =   525
            TabIndex        =   24
            Top             =   1485
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox picExpend 
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
      Height          =   4250
      Left            =   75
      ScaleHeight     =   4245
      ScaleWidth      =   10005
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   3135
      Width           =   10005
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   390
         Left            =   30
         TabIndex        =   122
         Top             =   240
         Width           =   270
         _Version        =   589884
         _ExtentX        =   476
         _ExtentY        =   688
         _StockProps     =   64
      End
      Begin VB.Frame fraBase 
         Height          =   3825
         Left            =   90
         TabIndex        =   99
         Top             =   60
         Width           =   9855
         Begin VB.ComboBox cbo联系人关系 
            Height          =   300
            Left            =   7770
            TabIndex        =   67
            Tag             =   "联系人关系"
            Top             =   3120
            Width           =   1950
         End
         Begin VB.TextBox txt其他关系 
            Height          =   300
            Left            =   8730
            MaxLength       =   30
            TabIndex        =   68
            Tag             =   "联系人其他关系"
            Top             =   3120
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.TextBox txt联系人身份证号 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   1365
            MaxLength       =   18
            TabIndex        =   65
            Tag             =   "联系人身份证号"
            Top             =   3075
            Width           =   2490
         End
         Begin VB.CommandButton cmd出生地点 
            Caption         =   "…"
            Height          =   255
            Left            =   4320
            TabIndex        =   51
            TabStop         =   0   'False
            Tag             =   "出生地点"
            ToolTipText     =   "热键：F3"
            Top             =   1958
            Width           =   285
         End
         Begin VB.TextBox txt单位帐户 
            Height          =   300
            Left            =   1155
            MaxLength       =   100
            TabIndex        =   62
            Tag             =   "单位帐户"
            Top             =   2730
            Width           =   3480
         End
         Begin VB.TextBox txt单位开户行 
            Height          =   300
            Left            =   5835
            MaxLength       =   100
            TabIndex        =   60
            Tag             =   "单位开户行"
            Top             =   2340
            Width           =   3885
         End
         Begin VB.CommandButton cmd区域 
            Caption         =   "…"
            Height          =   255
            Left            =   9420
            TabIndex        =   48
            TabStop         =   0   'False
            Tag             =   "区域"
            ToolTipText     =   "热键：F3"
            Top             =   1545
            Width           =   285
         End
         Begin VB.TextBox txt其他证件 
            Height          =   300
            Left            =   1155
            MaxLength       =   20
            TabIndex        =   45
            Tag             =   "其他证件"
            Top             =   1530
            Width           =   3480
         End
         Begin VB.ComboBox cbo费别 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Tag             =   "费别"
            Top             =   720
            Width           =   1485
         End
         Begin VB.ComboBox cbo身份 
            Height          =   300
            Left            =   8250
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   705
            Width           =   1470
         End
         Begin VB.ComboBox cbo职业 
            Height          =   300
            Left            =   5835
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1125
            Width           =   3885
         End
         Begin VB.ComboBox cbo国籍 
            Height          =   300
            Left            =   3150
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Tag             =   "国籍"
            Top             =   690
            Width           =   1485
         End
         Begin VB.ComboBox cbo学历 
            Height          =   300
            Left            =   3150
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Tag             =   "学历"
            Top             =   1125
            Width           =   1485
         End
         Begin VB.ComboBox cbo婚姻状况 
            Height          =   300
            Left            =   5835
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   705
            Width           =   1485
         End
         Begin VB.CommandButton cmd合同单位 
            Caption         =   "…"
            Height          =   255
            Left            =   9420
            TabIndex        =   54
            TabStop         =   0   'False
            Tag             =   "工作单位"
            ToolTipText     =   "热键：F3"
            Top             =   1950
            Width           =   285
         End
         Begin VB.CommandButton cmd联系人地址 
            Caption         =   "…"
            Height          =   255
            Left            =   9405
            TabIndex        =   70
            TabStop         =   0   'False
            Tag             =   "联系人地址"
            ToolTipText     =   "热键：F3"
            Top             =   3480
            Width           =   285
         End
         Begin VB.ComboBox cbo医疗付款 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Tag             =   "医疗付款"
            Top             =   1125
            Width           =   1485
         End
         Begin VB.TextBox txt工作单位 
            Height          =   300
            Left            =   5835
            MaxLength       =   100
            TabIndex        =   53
            Tag             =   "工作单位"
            Top             =   1935
            Width           =   3885
         End
         Begin VB.TextBox txt出生地点 
            Height          =   300
            Left            =   1155
            MaxLength       =   30
            TabIndex        =   50
            Tag             =   "出生地点"
            Top             =   1935
            Width           =   3480
         End
         Begin VB.TextBox txt单位电话 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1155
            MaxLength       =   20
            TabIndex        =   56
            Tag             =   "单位电话"
            Top             =   2340
            Width           =   1485
         End
         Begin VB.TextBox txt联系人电话 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5175
            MaxLength       =   20
            TabIndex        =   66
            Tag             =   "联系人电话"
            Top             =   3120
            Width           =   1365
         End
         Begin VB.TextBox txt单位邮编 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3585
            MaxLength       =   6
            TabIndex        =   58
            Tag             =   "单位邮编"
            Top             =   2340
            Width           =   1035
         End
         Begin VB.TextBox txt区域 
            Height          =   300
            Left            =   5835
            MaxLength       =   30
            TabIndex        =   47
            Tag             =   "区域"
            Top             =   1530
            Width           =   3885
         End
         Begin VB.TextBox txt医保号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1155
            MaxLength       =   30
            TabIndex        =   27
            Tag             =   "医保号"
            Top             =   285
            Width           =   3480
         End
         Begin VB.TextBox txt联系人姓名 
            Height          =   300
            Left            =   5835
            MaxLength       =   64
            TabIndex        =   64
            Tag             =   "联系人姓名"
            Top             =   2730
            Width           =   3870
         End
         Begin VB.TextBox txt联系人地址 
            Height          =   300
            Left            =   1170
            MaxLength       =   64
            TabIndex        =   69
            Tag             =   "联系人地址"
            Top             =   3465
            Width           =   8535
         End
         Begin VB.TextBox txt验证医保号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5835
            MaxLength       =   30
            TabIndex        =   29
            Tag             =   "验证医保号"
            Top             =   285
            Width           =   3870
         End
         Begin VB.Label lbl联系人身份证号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系人身份证号"
            Height          =   180
            Left            =   45
            TabIndex        =   120
            Top             =   3165
            Width           =   1260
         End
         Begin VB.Label lbl医保号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "验证医保号"
            Height          =   180
            Index           =   1
            Left            =   4845
            TabIndex        =   28
            Top             =   345
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位帐户"
            Height          =   180
            Left            =   390
            TabIndex        =   61
            Top             =   2790
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位开户行"
            Height          =   180
            Left            =   4860
            TabIndex        =   59
            Top             =   2400
            Width           =   900
         End
         Begin VB.Label lbl备注 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备注"
            Height          =   180
            Left            =   5220
            TabIndex        =   104
            Top             =   3840
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label lblPatiColor 
            Height          =   255
            Left            =   9060
            TabIndex        =   103
            Top             =   2700
            Width           =   105
         End
         Begin VB.Label lbl其他证件 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "其他证件"
            Height          =   180
            Left            =   390
            TabIndex        =   44
            Top             =   1590
            Width           =   720
         End
         Begin VB.Label lbl费别 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费别"
            Height          =   180
            Left            =   750
            TabIndex        =   30
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl出生地点 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生地点"
            Height          =   180
            Left            =   390
            TabIndex        =   49
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label lbl身份 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份"
            Height          =   180
            Left            =   7860
            TabIndex        =   36
            Top             =   765
            Width           =   360
         End
         Begin VB.Label lbl职业 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "职业"
            Height          =   180
            Left            =   5400
            TabIndex        =   42
            Top             =   1185
            Width           =   360
         End
         Begin VB.Label lbl国籍 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "国籍"
            Height          =   180
            Left            =   2730
            TabIndex        =   32
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl学历 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "学历"
            Height          =   180
            Left            =   2730
            TabIndex        =   40
            Top             =   1185
            Width           =   360
         End
         Begin VB.Label lbl婚姻状况 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "婚姻"
            Height          =   180
            Left            =   5385
            TabIndex        =   34
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl联系人姓名 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系人姓名"
            Height          =   180
            Left            =   4845
            TabIndex        =   63
            Top             =   2790
            Width           =   900
         End
         Begin VB.Label lbl联系人关系 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系人关系"
            Height          =   180
            Left            =   6840
            TabIndex        =   72
            Top             =   3180
            Width           =   1260
         End
         Begin VB.Label lbl联系人地址 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系人地址"
            Height          =   180
            Left            =   210
            TabIndex        =   102
            Top             =   3525
            Width           =   900
         End
         Begin VB.Label lbl联系人电话 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系人电话"
            Height          =   180
            Left            =   4185
            TabIndex        =   71
            Top             =   3180
            Width           =   900
         End
         Begin VB.Label lbl工作单位 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "工作单位"
            Height          =   180
            Left            =   5025
            TabIndex        =   52
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label lbl单位电话 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位电话"
            Height          =   180
            Left            =   390
            TabIndex        =   55
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label lbl单位邮编 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位邮编"
            Height          =   180
            Left            =   2760
            TabIndex        =   57
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label lbl单位开户行 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位开户行"
            Height          =   180
            Left            =   135
            TabIndex        =   101
            Top             =   4200
            Width           =   900
         End
         Begin VB.Label lbl单位帐号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位帐号"
            Height          =   180
            Left            =   4860
            TabIndex        =   100
            Top             =   4200
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医疗付款"
            Height          =   180
            Index           =   1
            Left            =   390
            TabIndex        =   38
            Top             =   1185
            Width           =   720
         End
         Begin VB.Label lbl区域 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "区域"
            Height          =   180
            Left            =   5385
            TabIndex        =   46
            Top             =   1590
            Width           =   360
         End
         Begin VB.Label lbl医保号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医保号"
            Height          =   180
            Index           =   0
            Left            =   570
            TabIndex        =   26
            Top             =   345
            Width           =   540
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10365
      TabIndex        =   93
      Top             =   585
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10350
      TabIndex        =   94
      Top             =   7590
      Width           =   1100
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   8925
      Left            =   180
      TabIndex        =   95
      Top             =   0
      Width           =   10125
      _Version        =   589884
      _ExtentX        =   17859
      _ExtentY        =   15743
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10365
      TabIndex        =   92
      Top             =   150
      Width           =   1100
   End
   Begin VB.CommandButton cmdCreateCard 
      Caption         =   "制卡(&M)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10365
      TabIndex        =   150
      Top             =   1005
      Width           =   1100
   End
   Begin MSCommLib.MSComm com 
      Left            =   11040
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmPatiCureCardEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
'入口参数
Private mstrPrivs As String, mlngModule As Long
Private mlngCardTypeID As Long, mstrCardNo As String
Public Enum gCardType
    Cr_发卡 = 0
    Cr_退卡 = 1
    Cr_绑定卡 = 2
    Cr_取消绑定 = 3
    Cr_换卡 = 4
    Cr_补卡 = 5
    Cr_挂失 = 6
    Cr_查询 = 7
    Cr_调整病人信息 = 8
End Enum
Private mEditType As gCardType
Private mEditTypeOld As gCardType
Private mstrBillNo  As String, mint记录状态   As Integer
Private mblnNOMoved As Boolean  '历史数据转移
Private mblnNotClick As Boolean
Private mblnUnLoad As Boolean
Private mstrPrepayPrivs As String
Private mstrIDImageFile As String
'---------------------------------------------------------------------------------------
'模块变量
Private mintSucces As Integer
Private Enum mTaskPancel_ID
      idx_TP_Tittle = 1
      Idx_TP_PatiBase = 2
      Idx_TP_PatiExpend = 3
      Idx_TP_PatiCard = 4
End Enum
Private Const mFormMaxHeight = 11330 '问题号:51071;问题号:56599
Private mblnChange As Boolean
Private Type Ty_ParaData
        blnSeekName As Boolean  '是否通过姓名进行模糊查找
        intNameDays As Integer     '模糊查找的天数
        blnShowExpend As Boolean '显示病人的扩展信息
        int退卡模式 As Integer  '0-不进行刷卡;1-刷卡退卡;2-单据号后再验证刷卡;3-1和2的共用模式
        bln记帐 As Boolean
        strControl As String  '输入项控制
End Type
Private mParaData As Ty_ParaData
Private mrsInfo As ADODB.Recordset
Private WithEvents mobjIDCard As zlIDCard.clsIDCard   '身份证读卡
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard   'IC卡接口
Attribute mobjICCard.VB_VarHelpID = -1
Private mobjReadCard As Object    '三方机构接口或读卡接口
Private mlng缺省卡号长度 As Long
Private mblnICCard As Boolean
Private mlng病人ID As Long
Private mblnNotChange As Boolean
Private mstr年龄 As String ' 记录年龄是否变化
Private mstr年龄单位 As String '同上
Private mstrCboSplit As String
Private Type Ty_CardProperty
       lng卡类别ID As Long
       str卡名称  As String
       lng卡号长度 As Long
       lng结算方式 As String
       bln自制卡 As Boolean
       bln严格控制 As Boolean
       lng领用ID As Long
       lng共用批次 As Long
       bln变价 As Boolean
       bln就诊卡 As Boolean
       str卡号密文 As String
       int密码长度 As Integer
       int密码长度限制 As Integer
       int密码规则 As Integer
       blnOneCard As Boolean '是否启用了一卡通接口,此模式下,票号严格管理;票号范围外的发卡和绑定卡不收费
       rs医疗卡费 As ADODB.Recordset
       dbl应收金额 As Double
       dbl实收金额 As Double
       bln是否制卡 As Boolean
       bln是否发卡 As Boolean
       bln是否写卡 As Boolean
       bln是否院外发卡 As Boolean
       lng发卡性质 As Long '0-不限制,1-同一个病人只允许发一张卡,2-同一个病人可以发多张卡,但需要提醒 问题号:57326
       bln是否重复使用 As Boolean
       str读卡性质 As String
       str特定项目 As String
       byt发卡控制 As Byte '0-卡号必须达到卡号长度，不足禁止；1-允许卡号小于等于卡号长度；2-发卡卡号小于卡号长度时检查并提醒
End Type
Private mCardType As Ty_CardProperty
Private mlngBillCardTypeID As Long

Private Type Ty_InsurePara
        lng外挂式医保险类 As Long   '外挂式医保的险类
End Type

Private Type TY_PayMoney
    lng医疗卡类别ID As Long
    bln消费卡 As Boolean
    str结算方式 As String
    str名称 As String
    str刷卡卡号 As String
    str刷卡密码 As String
    strNO As String
    lngID As Long '预交ID
    lng结帐ID As Long
End Type
Private mblnStructAdress As Boolean  '病人地址结构化录入
Private mblnShowTown As Boolean      '乡镇地址结构化录入

Private mCurPayMoney As TY_PayMoney
Private mInsurePara As Ty_InsurePara
Private mblnFirst As Boolean
Private mobjCardObject As clsCardObject
Private mcolPayMode As Collection
Private mstrBrushCardNo As String, mstrBrushPassWord As String
Private mcolBillBalance As Collection '退号的三方结算信息
Private mobjDelObject As clsCardObject
Private mintTabIndex卡号 As Integer '卡号的TabIndex
Private mintTabIndex刷卡卡号 As Integer '刷卡验证的TabIndex
Private mobjKeyboard As Object '建盘输入对象

Private mblnPassInputCardNo As Boolean  '是否密文输入卡号
Private mblnDefaultPassInputCardNo As Boolean '缺省刷卡是否密文输入卡号
Private mlng医疗卡长度  As Boolean
'问题号:56599
Private Enum mPageIndex
    常用 = 1
    病人证件 = 2
    药物过敏 = 3
    接种信息 = 4
    其他信息 = 5
    附加信息 = 6 '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
End Enum
Private mobjPlugIn As Object '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
Private mblnPlugin As Boolean
Private mrsEMPIOut As ADODB.Recordset
Private mlngPlugInHwnd As Long
Private mobjPubPatient As Object
Private mbln医嘱业务 As Boolean  '是否发生了医嘱业务
Private mstrPrivsPubPatient As String
Private mbln基本信息调整 As Boolean
Private mbln病历费 As Boolean '是否可以收取病历工本费
Private mbln存在门诊号 As Boolean  '该病人是否存在门诊号(针对建档病人)

'问题号:56599
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const C_ColumHeader = "过敏药物,1,1500,1;过敏反映,4,3000,1;过敏药物ID,1,100,0" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_InoculateHeader = "接种日期,4,2000,1;接种名称,4,2700,1;接种日期,4,2000,1;接种名称,4,1900,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_LinkManColumHeader = "姓名,4,1000,1;关系,4,2700,1;身份证号,4,2000,1;电话,4,1200,1;附加信息,4,2000,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_OtherInfoColumHeader = "信息名,4,2000,1;信息值,4,2700,1;信息名,4,2000,1;信息值,4,1900,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_CertificateHeader = "证件类型,4,2000,1;证件号码,4,2700,1;证件类型,4,2000,1;证件号码,4,1900,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
'Private Const C_血型 = "A型,B型,O型,AB型,不详"
Private Const C_BH = "阴,阳,不详,未查"
Private mdic医疗卡属性 As New Dictionary
Private mstr采集图片 As String '采集图片本地保存路径
Private mlng图像操作 As Long '指明当前对病人图像操作的类型(1-文件 2-采集 3-清除 4-身份证提取)
Private mblnAddPage As Boolean '是否显示发卡/绑定卡分页控件
Private mblnFromCardMgr As Boolean '是否从发卡界面进入
Private mstrTitle As String '用于窗体个性化保存的窗体名
Private mblnTab As Boolean
Private mstr必输项目 As String '发卡(绑定卡)界面必须输入项目
Private mbln自动门诊号 As Boolean
Private mstrFirstCode As String '第一种证件类型的编码
Private Type Ty_FeeProperty
       bln变价 As Boolean
       rs病历费 As ADODB.Recordset
       dbl应收金额 As Double
       dbl实收金额 As Double
       bln是否全退 As Boolean
       bln是否退现 As Boolean
End Type
Private mFeeType As Ty_FeeProperty
Private mstrPriceGrade As String, mstrPrePriceGrade As String
'------------------------预交变量-------------------------------------
Private mFactProperty As Ty_FactProperty
Private mblnBill预交 As Boolean '是否严格票据管理
Private mbyt预交 As Byte '票据号码长度
Private mstrRedFact As String '预交红票
Private mlng领用ID As Long '预交领用ID
Private mblnPrepayPrint As Boolean '是否打印预交票据
Private mstrPrepayInvioce As String '预交票据号
Private mlng预交ID As Long '生成预交记录的ID
Private mstrPrePayNo As String
Private mlng预交病人ID As Long
Private mdat预交时间 As Date
'-----------------------收费发票-------------------------------------
Private Type Ty_PrintProperty
    bytPrintType As Byte '发卡票据打印方式
    bytPrintFormat As Byte '发卡打印格式:发卡|绑定卡
    strUseType As String '使用类别
    blnPrint As Boolean '是否打印票据
    lng领用ID As Long '本次票据领用ID
    strBackInvoice As String  '回收票据
    dtPrintdate  As Date '打印时间
End Type
Private mPrint As Ty_PrintProperty
Private mblnSaveDeposit As Boolean              '病人缴款余额是否存为预交款
Private mblnGetBirth As Boolean '判断是否允许通过年龄计算生日

'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function LoadCardData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载卡片数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-12 11:03:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, strCardPass As String, strErrMsg As String
    If mEditType = Cr_绑定卡 Or mEditType = Cr_发卡 Then LoadCardData = True: Exit Function
    If mlngCardTypeID = 0 Then Exit Function
    
    If mstrCardNo <> "" Then
        If GetPatiID(mlngCardTypeID, mstrCardNo, False, lng病人ID, strCardPass, strErrMsg) = False Then Exit Function
        If lng病人ID <= 0 Then
           Exit Function
        End If
    Else
        lng病人ID = mlng病人ID  '修改病人
    End If
    If lng病人ID = 0 Then LoadCardData = True: Exit Function
    If GetPatient("-" & lng病人ID, False, True) = False Then Exit Function
    
    Call LoadPatiInfor: Call zlQueryEMPIPatiInfo
    If mEditType = Cr_退卡 Then
        Me.txt卡号.Text = GetCardNODencode(Trim(mstrCardNo), mlngCardTypeID)
        Me.lbl卡号.Tag = mstrCardNo
    End If

End Function
Private Function InitCardType() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡类别
    '返回:初始卡不成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-06 17:03:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, i As Long
    Dim str批次 As String, varData As Variant, varTemp As Variant, lng就诊卡ID As Long
    
    Err = 0: On Error GoTo errHandle
    '问题号:57326
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    '76505,冉俊明,2014-8-14,终止问题回退修改,注：修改医疗卡类别后必须重启系统才生效
    Set rsTemp = zlGet医疗卡类别
    
    rsTemp.Filter = "ID=" & mlngCardTypeID & " And 是否启用=1"
    
    '74449,冉俊明,2014-6-25,“上次发卡类别”不存在或被停用时无法提取其它卡类别
    If rsTemp.EOF Then Exit Function
    
    With mCardType
        .str卡名称 = Nvl(rsTemp!名称)
        .lng卡号长度 = Val(Nvl(rsTemp!卡号长度))
        .lng结算方式 = Trim(Nvl(rsTemp!结算方式))
        .bln自制卡 = Val(Nvl(rsTemp!是否自制)) = 1
        .bln严格控制 = Val(Nvl(rsTemp!是否严格控制)) = 1
        .str卡号密文 = Nvl(rsTemp!卡号密文)
        .int密码长度 = Val(Nvl(rsTemp!密码长度))
        .int密码长度限制 = Val(Nvl(rsTemp!密码长度限制))
        .int密码规则 = Val(Nvl(rsTemp!密码规则))
        .bln是否制卡 = Val(Nvl(rsTemp!是否制卡)) = 1 '问题号:56599
        .bln是否发卡 = Val(Nvl(rsTemp!是否发卡)) = 1
        .bln是否写卡 = Val(Nvl(rsTemp!是否写卡)) = 1
        .bln是否院外发卡 = (InStr(mstrPrivs, ";发卡;") > 0 And .bln自制卡 = False And .bln是否发卡 = True) '问题号:56599
        .lng发卡性质 = Val(Nvl(rsTemp!发卡性质)) '问题号:57326
        .lng卡类别ID = Val(Nvl(rsTemp!id)) '问题号:57326
        .bln是否重复使用 = Val(Nvl(rsTemp!是否重复使用))
        .bln就诊卡 = .str卡名称 = "就诊卡" And Val(Nvl(rsTemp!是否固定)) = 1 And Nvl(rsTemp!部件) = "" '问题号:57962
        .blnOneCard = False
        .str读卡性质 = Nvl(rsTemp!读卡性质, "1000")
        .str特定项目 = Trim(Nvl(rsTemp!特定项目))
        .byt发卡控制 = Val(Nvl(rsTemp!发卡控制)) '104238
        If .str特定项目 <> "" Then
            Set .rs医疗卡费 = zlGetSpecialItemFee(.str特定项目, mstrPriceGrade)
            '问题:48090
            '问题号:56599
            If (.bln自制卡 = True Or .bln是否院外发卡) And .rs医疗卡费 Is Nothing Then
                MsgBox .str卡名称 & "未设置对应的卡费,请到【医疗卡类别管理】中设置!", vbInformation + vbOKOnly, gstrSysName
                mblnUnLoad = True
                mblnChange = False
                Exit Function
            End If
            If .bln就诊卡 Then .blnOneCard = GetOneCard.RecordCount > 0
        Else
            Set .rs医疗卡费 = Nothing
        End If
        str批次 = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModule, "0")
        '领用ID,卡类别ID|...
        varData = Split(str批次, "|")
        For i = 0 To UBound(varData)
             varTemp = Split(varData(i), ",")
             If Val(varTemp(0)) <> 0 Then
                If Val(varTemp(1)) = mlngCardTypeID Then
                    .lng共用批次 = Val(varTemp(0)): Exit For
                End If
             End If
        Next
        txtPass.MaxLength = .int密码长度
        txtAudi.MaxLength = .int密码长度
        txt卡号.PasswordChar = IIf(.str卡号密文 <> "", "*", "")
        txt刷卡卡号.PasswordChar = IIf(.str卡号密文 <> "", "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End With
    InitCardType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Init病历费()
    Dim strSQL As String
    If Not mbln病历费 Then
        chk病历费.Enabled = False
        Set mFeeType.rs病历费 = Nothing
        Exit Sub
    End If
    
    On Error GoTo Errhand
    Set mFeeType.rs病历费 = zlGetSpecialItemFee("病历费", mstrPriceGrade)
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitInsurePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2011-07-07 02:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, i As Long
    With mInsurePara
        .lng外挂式医保险类 = 0
        varTemp = Split(GetSetting("ZLSOFT", "公共全局", "本地支持的医保", ""), ",")
        For i = 0 To UBound(varTemp)
            If IsNumeric(varTemp(i)) Then
                If zlCheckMCOutMode(Val(varTemp(i))) Then .lng外挂式医保险类 = Val(varTemp(i)): Exit For
            End If
        Next
    End With
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2011-07-01 11:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mParaData
        .blnSeekName = zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModule) = "1"
        .intNameDays = Val(zlDatabase.GetPara("姓名查找天数", glngSys, mlngModule))
        .blnShowExpend = Val(zlDatabase.GetPara("显示扩展信息", glngSys, mlngModule))
        .int退卡模式 = Val(zlDatabase.GetPara("退卡刷卡", glngSys, mlngModule))
        '0-不进行刷卡;1-刷卡退卡;2-单据号后再验证刷卡;3-1和2的共用模式
        .bln记帐 = Val(zlDatabase.GetPara("卡费记帐", glngSys, mlngModule, , Array(chk记帐), InStr(1, mstrPrivs, ";参数设置;"))) = "1"
        .strControl = zlDatabase.GetPara("输入项控制", glngSys, mlngModule)
    End With
    mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '病人地址结构化录入
    mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '乡镇地址结构化录入
    '94941:李南春,2016/4/7,是否自动产生门诊号
    mbln自动门诊号 = Val(zlDatabase.GetPara("自动门诊号", glngSys, mlngModule)) = 1
    '95809:李南春,2016/8/19,是否允许收取病历费
    mbln病历费 = Val(zlDatabase.GetPara("收取病历费", glngSys, mlngModule, , Array(chk病历费), InStr(1, mstrPrivs, ";参数设置;"))) = 1
    
    '104726:李南春,2017/4/17,收费发票打印发卡票据
    mPrint.bytPrintType = Val(zlDatabase.GetPara("发卡打印方式", glngSys, mlngModule))
    mPrint.bytPrintFormat = Val(Split(zlDatabase.GetPara("医疗卡收据格式", glngSys, mlngModule) & "|", "|")(IIf(mEditType = Cr_绑定卡, 1, 0)))
End Sub
Private Sub SetDefaultLen()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的编辑长度
    '编制:刘兴洪
    '日期:2011-06-28 02:50:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
        strSQL = " " & _
    "   Select A.姓名,A.门诊号,A.身份证号,A.年龄,A.家庭地址,A.家庭电话,A.医保号,A.家庭地址, " & _
    "          A.其他证件,A.家庭地址邮编,A.区域,A.出生地点,A.工作单位,A.单位电话,A.户口地址,A.户口地址邮编," & _
    "          A.单位邮编,A.单位开户行,A.单位帐号,A.联系人姓名,A.联系人地址,A.联系人电话,B.卡号,B.密码,A.手机号" & _
    "   From 病人信息 A,病人医疗卡信息 B" & _
    "   Where a.病人id = 0 and a.病人ID=b.病人ID and B.卡类别ID=0 " & _
    " "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    txtPatient.MaxLength = rsTemp.Fields("姓名").DefinedSize
    txt身份证号.MaxLength = rsTemp.Fields("身份证号").DefinedSize
    txt门诊号.MaxLength = rsTemp.Fields("门诊号").DefinedSize - 1
    txt年龄.MaxLength = rsTemp.Fields("年龄").DefinedSize
    txt家庭地址.MaxLength = rsTemp.Fields("家庭地址").DefinedSize
    padd家庭地址.MaxLength = rsTemp.Fields("家庭地址").DefinedSize
    txt家庭电话.MaxLength = rsTemp.Fields("家庭电话").DefinedSize
    txt医保号.MaxLength = rsTemp.Fields("医保号").DefinedSize
    txt家庭邮编.MaxLength = rsTemp.Fields("家庭地址邮编").DefinedSize
    txt户口地址.MaxLength = rsTemp.Fields("户口地址").DefinedSize
    padd户口地址.MaxLength = rsTemp.Fields("户口地址").DefinedSize
    txt户口地址邮编.MaxLength = rsTemp.Fields("户口地址邮编").DefinedSize
    txt其他证件.MaxLength = rsTemp.Fields("其他证件").DefinedSize
    txt区域.MaxLength = rsTemp.Fields("区域").DefinedSize
    txt出生地点.MaxLength = rsTemp.Fields("出生地点").DefinedSize
    txt工作单位.MaxLength = rsTemp.Fields("工作单位").DefinedSize
    txt单位电话.MaxLength = rsTemp.Fields("单位电话").DefinedSize
    txt单位邮编.MaxLength = rsTemp.Fields("单位邮编").DefinedSize
    txt单位开户行.MaxLength = rsTemp.Fields("单位开户行").DefinedSize
    txt单位帐户.MaxLength = rsTemp.Fields("单位帐号").DefinedSize
    txt联系人姓名.MaxLength = rsTemp.Fields("联系人姓名").DefinedSize
    txt联系人地址.MaxLength = rsTemp.Fields("联系人地址").DefinedSize
    txt联系人电话.MaxLength = rsTemp.Fields("联系人电话").DefinedSize
    txtPass.MaxLength = rsTemp.Fields("密码").DefinedSize
    txtAudi.MaxLength = rsTemp.Fields("密码").DefinedSize
    txt手机.MaxLength = rsTemp.Fields("手机号").DefinedSize
    If mCardType.lng卡号长度 = 0 Then mCardType.lng卡号长度 = rsTemp.Fields("卡号").DefinedSize
    txt卡号.MaxLength = mCardType.lng卡号长度
    If mInsurePara.lng外挂式医保险类 = 920 Then '北京外挂
        txt医保号.MaxLength = 12
    Else
        txt医保号.MaxLength = 30
    End If
    txt医保号.ToolTipText = "最大长度" & txt医保号.MaxLength & "位"
    txt验证医保号.MaxLength = txt医保号.MaxLength
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入数据是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-01 10:35:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtl As Control, lngLen As Long, strMCAccount As String, lngTmp As Long
    Dim strTemp As String, i As Long
    Dim blnNewPati As Boolean, str出生时间 As String
    Dim strBirthday As String, strAge As String, strSex As String, strErrInfo As String, strInfo As String
    
    blnNewPati = False
    If mrsInfo Is Nothing Then
        blnNewPati = True
    ElseIf mrsInfo.State <> 1 Then
        blnNewPati = True
    End If
  
    For Each objCtl In Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("TextBox")  '文本
            lngLen = objCtl.MaxLength
            If lngLen <> 0 Then
                If zlCommFun.ActualLen(objCtl.Text) > lngLen Then
                    MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "最多只能输入" & lngLen \ 2 & "个汉字或" & lngLen & "个字符,请检查", vbInformation + vbOKOnly, gstrSysName
                    If InStr(1, ",姓名,门诊号,身份证号,现住址,户口地址,家庭电话,手机号,卡号,密码,验证,", "," & objCtl.Tag & ",") > 0 Then
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    Else
                        If wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded = False Then
                            wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded = True
                        End If
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    End If
                End If
            End If
            If Trim(objCtl.Text) = "" And InStr(1, ",姓名,门诊号,卡号," & mstr必输项目, "," & objCtl.Tag & ",") > 0 Then
                '必输项目
                If Not (mEditType = Cr_调整病人信息 And objCtl.Tag = "卡号") Then
                    If Not (blnNewPati = False And objCtl.Tag = "门诊号") Then
                        If objCtl.Enabled And objCtl.Visible Then
                            MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "必须输入,请检查", vbInformation + vbOKOnly, gstrSysName
                            If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            '90568:李南春,2015/11/18,联系人身份证号只检查身份证有效性
            If objCtl.Tag = "联系人身份证号" And Trim(objCtl.Text) <> "" And Not mobjPubPatient Is Nothing Then
                If Not mobjPubPatient.CheckPatiIdcard(Trim(objCtl.Text), strBirthday, strAge, strSex, strErrInfo) Then
                    MsgBox strErrInfo, vbInformation, gstrSysName
                    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                    Exit Function
                End If
            End If
            '81103,冉俊明,2014-12-29,录入身份证号后,出生日期、年龄、性别的同步关联检查和调整
            If objCtl.Tag = "身份证号" And Trim(objCtl.Text) <> "" Then
                If Not mobjPubPatient Is Nothing Then
                    'CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, _
                    '    Optional strAge As String, Optional strSex As String, Optional strErrInfo As String) As Boolean
                    '功能：身份证号码合法性校验
                    '入参：strIdCard 身份证号码
                    '出参：strBirthday  函数返回True为出生日期
                    '         strAge 函数返回True为年龄
                    '         strSex 函数返回True为性别
                    '         strErrInfo 函数返回False为错误信息
                    '返回：True/False  身份证合法返回True(可从strBirthday，strSex获取出生日期和性别)，
                    '       否则返回False(可从strErrInfo获取详细错误信息)
                    If mobjPubPatient.CheckPatiIdcard(Trim(objCtl.Text), strBirthday, strAge, strSex, strErrInfo) Then
                        '新病人或调整无业务数据的已有病人信息时提示是否调整不一致的基本信息
                        If blnNewPati Or mEditType = Cr_调整病人信息 Then
                            If strSex <> zlstr.NeedName(cbo性别.Text) Then strInfo = "性别"
                            If strAge <> Trim(txt年龄.Text) & cbo年龄单位 Then strInfo = strInfo & IIf(strInfo = "", "年龄", "、年龄")
                            If Format(strBirthday, "yyyy-mm-dd") <> txt出生日期.Text Then strInfo = strInfo & IIf(strInfo = "", "出生日期", "、出生日期")
                            
                            If strInfo <> "" Then
                                If InStr(mstrPrivsPubPatient, ";基本信息调整;") > 0 Then
                                    If MsgBox("输入的" & strInfo & "与身份证号的" & strInfo & "不一致，" & _
                                            "将根据身份证号修改" & strInfo & "，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                                        Call zlControl.CboLocate(cbo性别, strSex)
                                        txt出生日期.Text = Format(strBirthday, "yyyy-mm-dd")
                                        '只有病人发生医嘱业务，操作员有“基本信息调整”权限，且基础信息不一致时操作员选择继续，才单独调用SavePatiBaseInfo接口
                                    Else
                                        Exit Function
                                    End If
                                Else
                                    If MsgBox("输入的" & strInfo & "与身份证号的" & strInfo & "不一致，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                                         Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Else
                        MsgBox strErrInfo, vbInformation, gstrSysName
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    End If
                End If
            End If
        '89242:李南春,2015/12/10,结构化地址信息检查
        Case UCase("Patiaddress")
            If mblnStructAdress And objCtl.Enabled Then
                lngLen = objCtl.MaxLength
                If lngLen <> 0 Then
                    If zlCommFun.ActualLen(objCtl.value) > lngLen Then
                        MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "最多只能输入" & lngLen \ 2 & "个汉字,请检查。", vbInformation + vbOKOnly, gstrSysName
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    End If
                End If
                If objCtl.CheckNullValue(, True, False) <> "" Then
                    MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "的" & objCtl.CheckNullValue(, True, False) & "尚未输入,请检查。", vbInformation + vbOKOnly, gstrSysName
                    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                    Exit Function
                End If
            End If
        Case Else
        End Select
    Next
    
    '123098:李南春，2018/3/20，新病人，或者建档病人没有门诊号要自动生成门诊的情况
    If Not blnNewPati And (Not mbln自动门诊号 Or mbln存在门诊号 Or (mEditType <> Cr_绑定卡 And mEditType <> Cr_发卡)) Then isValied = True: Exit Function
    
    If Not IsNumeric(txt门诊号.Text) And txt门诊号.Text <> "" Then
        MsgBox "不是有效的门诊号,请检查！", vbInformation, gstrSysName
        If txt门诊号.Enabled And txt门诊号.Visible Then txt门诊号.SetFocus
        Exit Function
    End If
    If IsNumeric(txt门诊号.Text) Then
        If ExistClinicNO(txt门诊号.Text) Then
            If mbln自动门诊号 Then
                If txt门诊号.Text <> lbl门诊号.Tag Then
                    MsgBox "发现该病人的病人门诊号[" & txt门诊号.Text & "]已经被其它病人使用,系统将自动更换一个不重复的号码！", vbInformation, gstrSysName
                    txt门诊号.Text = zlGet门诊号: lbl门诊号.Tag = txt门诊号.Text
                    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
                    Exit Function
                Else
                    '自动产生的号如果没有修改，则直接再次自动产生即可
                    txt门诊号.Text = zlGet门诊号: lbl门诊号.Tag = txt门诊号.Text
                End If
            Else
                MsgBox "发现该病人的病人门诊号[" & txt门诊号.Text & "]已经被其它病人使用,请更换一个不重复的号码！", vbInformation, gstrSysName
                txt门诊号.Text = "": lbl门诊号.Tag = ""
                If txt门诊号.Enabled And txt门诊号.Visible Then txt门诊号.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Not blnNewPati Then isValied = True: Exit Function
    
    If txt医保号.Text <> "" Or txt验证医保号.Text <> "" Then
        If txt医保号.Text <> txt验证医保号.Text And txt验证医保号.Visible Then
            MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
            If txt医保号.Visible And txt医保号.Enabled Then txt医保号.SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(txt医保号.Text) > txt医保号.MaxLength Then
            MsgBox "请检查,医保号最大长度不能超过" & txt医保号.MaxLength & "个字符！", vbInformation, gstrSysName
            If txt医保号.Visible And txt医保号.Enabled Then txt医保号.SetFocus
            Exit Function
        End If
    End If
        
    
    strMCAccount = Trim(txt医保号.Text)
    If mInsurePara.lng外挂式医保险类 = 920 And strMCAccount <> lbl医保号(0).Tag And strMCAccount <> "" Then
        strMCAccount = UCase(strMCAccount)
        If CheckExistsMCNO(strMCAccount) Then
            If txt医保号.Visible And txt医保号.Enabled Then txt医保号.SetFocus
            Exit Function
        End If
    End If
    If Not IsDate(txt出生日期.Text) Then
        MsgBox "必须正确输入出生日期！", vbInformation, gstrSysName
        If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
        Exit Function
    End If
    If Trim(txt年龄.Text) = "" Then
        MsgBox "必须输入[年龄]！", vbExclamation, gstrSysName
        If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
        Exit Function
    End If
    '69026,冉俊明,2014-8-11,年龄有效性检查
    '76703,冉俊明,2014-8-15
    If txt年龄.Enabled And txt年龄.Visible Then
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), _
                IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
                IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
            If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus: Exit Function
        End If
    End If
    If IsDate(txt出生日期.Text) Then
        '76669，李南春,2014-8-15,年龄与出生日期检查
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        If CDate(str出生时间) > zlDatabase.Currentdate Then
            If MsgBox("出生时间：" & str出生时间 & " 超过了当前系统时间。" & _
                vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性 ，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If cbo费别.ListIndex = -1 Then
        MsgBox "必须确定[费别]！", vbExclamation, gstrSysName
        If cbo费别.Enabled And cbo费别.Visible Then cbo费别.SetFocus
        Exit Function
    End If
    If cbo国籍.ListIndex = -1 Then
        MsgBox "必须确定[国籍]！", vbExclamation, gstrSysName
        If cbo国籍.Enabled And cbo国籍.Visible Then cbo国籍.SetFocus
        Exit Function
    End If
    
    If cbo民族.ListIndex = -1 Then
        MsgBox "必须确定[民族]！", vbExclamation, gstrSysName
        If cbo民族.Enabled And cbo民族.Visible Then cbo民族.SetFocus
        Exit Function
    End If
    
    '检查相似的病人,已免重复
    If mrsInfo Is Nothing Then
        strTemp = SimilarIDs(zlstr.NeedName(cbo国籍.Text), zlstr.NeedName(cbo民族), CDate(IIf(IsDate(txt出生日期.Text), txt出生日期.Text, #1/1/1900#)), zlstr.NeedName(cbo性别), txtPatient.Text, txt身份证号.Text)
        If strTemp <> "" Then
            i = UBound(Split(strTemp, "|")) + 1
            strTemp = Replace(strTemp, "|", vbCrLf)
            If i > 20 Then strTemp = Mid(strTemp, 1, 200) & "..."
            
            If MsgBox("在已有的病人信息中发现 " & i & " 个信息相似的病人(国籍,民族,性别,姓名,出生日期相同或身份证号相同): " & vbCrLf & vbCrLf & _
                strTemp & vbCrLf & vbCrLf & "确实要保存该病人的信息吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            Else
                MsgBox "该病人的相似记录可以使用""合并""功能处理！", vbInformation, gstrSysName
            End If
        End If
    End If
    isValied = True
End Function
Public Function zlShowCard(ByVal frmMain As Object, lngModule As Long, strPrivs As String, _
     EditType As gCardType, lngCardTypeID As Long, _
     Optional strCardNo As String = "", _
     Optional lng病人ID As Long, _
     Optional strBillNo As String, _
     Optional int记录状态 As Integer, _
     Optional blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示卡片
    '入参:frmMain-调用的主窗体
    '       EditType:=编辑类型
    '       lngCardTypeID-卡类别Id
    '       strCardNo-卡号
    '出参:
    '返回:成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2011-06-28 22:29:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditTypeOld = EditType
    mEditType = EditType: mlngModule = lngModule: mstrPrivs = strPrivs
    mlngCardTypeID = lngCardTypeID: mstrCardNo = strCardNo: mintSucces = 0
    mstrBillNo = strBillNo: mint记录状态 = int记录状态: mblnNOMoved = blnNOMoved
    mlng病人ID = lng病人ID: mblnUnLoad = False
    mblnFromCardMgr = False
    If frmMain.Caption = "医疗卡发放管理" Then mblnFromCardMgr = True
    Me.Show 1, frmMain
    zlShowCard = mintSucces > 0
End Function
Private Sub InitTaskPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化InitTaskPancel
    '编制:刘兴洪
    '日期:2011-06-30 18:20:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    Call wndTaskPanel.SetGroupInnerMargins(2, 0, 2, 0)
    Call wndTaskPanel.SetGroupOuterMargins(2, -10, 2, -21)
    Call wndTaskPanel.SetMargins(2, 16, 2, 10, 30)
     
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    
    If Not (mEditType = Cr_挂失 Or mEditType = Cr_调整病人信息 Or mEditType = Cr_补卡 Or mEditType = Cr_换卡) Or ((mEditType = Cr_补卡 Or mEditType = Cr_换卡) And gbln收费发票) Then
        Set tkpGroup = wndTaskPanel.Groups.Add(idx_TP_Tittle, "")
        Set Item = tkpGroup.Items.Add(idx_TP_Tittle, "", xtpTaskItemTypeControl)
        Set Item.Control = picTittle
        fraSplit.BackColor = Item.BackColor
        picTittle.BackColor = Item.BackColor
        tkpGroup.Expandable = False
        Call Item.SetMargins(0, -19, 0, 0)
    End If

    Set tkpGroup = wndTaskPanel.Groups.Add(Idx_TP_PatiBase, "病人基本信息")
    Set Item = tkpGroup.Items.Add(Idx_TP_PatiBase, "", xtpTaskItemTypeControl)
    Set Item.Control = picBasePati
    fra.BackColor = Item.BackColor
    picBasePati.BackColor = Item.BackColor
    tkpGroup.Expandable = False
    
    Set tkpGroup = wndTaskPanel.Groups.Add(Idx_TP_PatiExpend, "更多病人信息")
    tkpGroup.Tooltip = "按CTRL+E 显示更多的病人信息"
    Set Item = tkpGroup.Items.Add(Idx_TP_PatiExpend, "", xtpTaskItemTypeControl)
    Set Item.Control = picExpend
    picExpend.BackColor = Item.BackColor
    fraBase.BackColor = picExpend.BackColor
    If mEditType = Cr_调整病人信息 Then
        tkpGroup.Expandable = False
        tkpGroup.Expanded = True
    Else
        tkpGroup.Expanded = mParaData.blnShowExpend
    End If
    
    If mEditType <> Cr_调整病人信息 Then
        Set tkpGroup = wndTaskPanel.Groups.Add(Idx_TP_PatiCard, IIf(mCardType.str卡名称 = "", "医疗卡", mCardType.str卡名称))
        tkpGroup.Expandable = False
        tkpGroup.Expanded = True
        picCard.Height = 2205: If mEditType <> Cr_绑定卡 And mEditType <> Cr_发卡 Then picCard.Height = 1705
        Set Item = tkpGroup.Items.Add(Idx_TP_PatiCard, "", xtpTaskItemTypeControl)
        Set Item.Control = picCard
        picCard.BackColor = Item.BackColor
        fraCard.BackColor = Item.BackColor
        tkpGroup.Expanded = True
    End If
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date
    Dim intReDelt As Integer
    Dim blnNotShowMsg As Boolean
    If cboNO.Locked Then Exit Sub
    
    '转换成大写(汉字不可处理)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        'Call SetNOInputLimit(cboNO, KeyAscii)
        Exit Sub
    End If
    If Not (cboNO.Text <> "" And Not cboNO.Locked) Then Exit Sub
    
    cboNO.Text = GetFullNO(cboNO.Text, 16)
    '是否已转入后备数据表中
    If zlDatabase.NOMoved("住院费用记录", cboNO.Text, , "5") Then
        If Not ReturnMovedExes(cboNO.Text, 5, Me.Caption) Then Exit Sub
        mblnNOMoved = False
    End If
    '单据权限
    If Not ReadBillInfo(2, cboNO.Text, 5, strOper, vDate) Then
        txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Sub
    End If
    If Not BillOperCheck(8, strOper, vDate, "退卡") Then
        txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Sub
    End If
    '读取要退卡的记录(由NO)
    '过程中提示了错误信息后，外层不再提示
    intReDelt = ReadBill(cboNO.Text, blnNotShowMsg)
    If blnNotShowMsg Then
        txtPatient.Text = ""
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Exit Sub
    End If
    Select Case intReDelt
        Case -1
            If InStr(1, "13", mParaData.int退卡模式) > 0 Then
                If txt刷卡卡号.Visible And txt刷卡卡号.Enabled Then txt刷卡卡号.SetFocus
            Else
               If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
            End If
        Case 0
            If glngSys Like "8??" Then
                MsgBox "读取该会员卡发放记录失败！", vbExclamation, gstrSysName
            Else
                MsgBox "读取该医疗卡发放记录失败！", vbExclamation, gstrSysName
            End If
            txtPatient.Text = "": cboNO.SetFocus
        Case 1
            If glngSys Like "8??" Then
                MsgBox "该会员卡发放记录不存在！", vbExclamation, gstrSysName
            Else
                MsgBox "该医疗卡发放记录不存在！", vbExclamation, gstrSysName
            End If
            txtPatient.Text = "": cboNO.SetFocus
        Case 2
            If glngSys Like "8??" Then
                MsgBox "该会员卡发放记录已经退除！", vbExclamation, gstrSysName
            Else
                MsgBox "该医疗卡发放记录已经退除！", vbExclamation, gstrSysName
            End If
            txtPatient.Text = ""
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    End Select
End Sub

Private Sub cbo费别_Change()
    mblnChange = True
End Sub

Private Sub cbo费别_Click()
    mblnChange = True
    If mblnNotChange Then Exit Sub
    Call LoadCardFee
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    SearchCombox cbo费别, KeyAscii
End Sub

Private Sub cbo费别_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(cbo费别)
End Sub

Private Sub cbo国籍_Change()
    mblnChange = True
End Sub

Private Sub cbo国籍_Click()
    mblnChange = True
End Sub

Private Sub cbo国籍_KeyPress(KeyAscii As Integer)
    SearchCombox cbo国籍, KeyAscii
End Sub

Private Sub cbo国籍_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(cbo国籍)
End Sub

Private Sub cbo婚姻状况_Change()
    mblnChange = True
End Sub

Private Sub cbo婚姻状况_Click()
    mblnChange = True
End Sub

Private Sub cbo婚姻状况_KeyPress(KeyAscii As Integer)
    SearchCombox cbo婚姻状况, KeyAscii
End Sub

Private Sub cbo联系人关系_Change()
        mblnChange = True
End Sub

Private Sub cbo联系人关系_Click()
    mblnChange = True
    With cbo联系人关系
        If .ListIndex = 8 And txt其他关系.Visible = False Then
            .Width = 950: txt其他关系.Visible = True
        Else
            .Width = 1950: txt其他关系.Visible = False: txt其他关系.Text = ""
        End If
    End With
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("关系") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("关系")) = zlstr.NeedName(cbo联系人关系.Text)
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("附加信息")) = zlstr.NeedName(txt其他关系.Text)
    End If
End Sub

Private Sub cbo联系人关系_KeyPress(KeyAscii As Integer)

    SearchCombox cbo联系人关系, KeyAscii
End Sub

Private Sub cbo联系人关系_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(cbo联系人关系)
End Sub

Private Sub cbo民族_Change()
    mblnChange = True
End Sub

Private Sub cbo民族_KeyPress(KeyAscii As Integer)
    SearchCombox cbo民族, KeyAscii
End Sub

Private Sub cbo年龄单位_Click()
    mblnChange = True
End Sub

Private Sub cbo年龄单位_LostFocus()
    '69026,冉俊明,2014-8-8,检查输入年龄
    '76703,冉俊明,2014-8-15
    '111836:李南春，2017/7/21，年龄反算
    Dim strBirth As String
    
    If mobjPubPatient Is Nothing Then Exit Sub
    If cbo年龄单位.Text <> mstr年龄单位 Then
        mblnNotChange = True
        If mblnGetBirth Then
            '103807:李南春，2016/12/20，年龄反算精确到小时
            If mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), strBirth) Then
                txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                txt出生时间.Text = Format(strBirth, "hh:mm")
            End If
        End If
        mblnNotChange = False
        mstr年龄单位 = cbo年龄单位.Text
    End If
    
    If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & cbo年龄单位.Text, _
            IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
            IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
        If txt年龄.Visible And txt年龄.Enabled Then txt年龄.SetFocus: Exit Sub
    End If
End Sub

Private Sub cbo身份_Change()
    mblnChange = True
End Sub

Private Sub cbo身份_Click()
    mblnChange = True
End Sub

Private Sub cbo身份_KeyPress(KeyAscii As Integer)
    SearchCombox cbo身份, KeyAscii
End Sub

Private Sub cbo学历_Change()
    mblnChange = True
End Sub

Private Sub cbo学历_Click()
    mblnChange = True
End Sub

Private Sub cbo学历_KeyPress(KeyAscii As Integer)
  SearchCombox cbo学历, KeyAscii
End Sub

Private Sub cbo学历_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(cbo学历)
End Sub

Private Sub cbo医疗付款_Change()
    mblnChange = True
End Sub

Private Sub cbo医疗付款_Click()
    On Error GoTo ErrHandler
    If gintPriceGradeStartType < 2 Then Exit Sub
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlstr.NeedName(cbo医疗付款.Text), , , mstrPriceGrade)
    If mstrPrePriceGrade = mstrPriceGrade Then Exit Sub
    mstrPrePriceGrade = mstrPriceGrade

    If mCardType.str特定项目 <> "" Then
        Set mCardType.rs医疗卡费 = zlGetSpecialItemFee(mCardType.str特定项目, mstrPriceGrade)
    Else
        Set mCardType.rs医疗卡费 = Nothing
    End If

    If mbln病历费 Then
        Set mFeeType.rs病历费 = zlGetSpecialItemFee("病历费", mstrPriceGrade)
    Else
        Set mFeeType.rs病历费 = Nothing
    End If

    Call LoadCardFee
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo医疗付款_KeyPress(KeyAscii As Integer)
     SearchCombox cbo医疗付款, KeyAscii
End Sub

Private Sub cbo医疗付款_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(cbo医疗付款)
End Sub

Private Sub cbo支付方式_Change()
    mblnChange = True
End Sub

Private Sub cbo支付方式_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    
    mblnChange = True
    If mblnNotClick = True Then Exit Sub
    
    With mCurPayMoney
            .lng医疗卡类别ID = 0
            .bln消费卡 = False
            .str结算方式 = ""
            .str名称 = ""
            .str刷卡卡号 = ""
            .str刷卡密码 = ""
            .strNO = ""
            .lngID = 0
            .lng结帐ID = 0
     End With
     
    If Not (mEditType = Cr_发卡 Or mEditType = Cr_补卡 Or mEditType = Cr_绑定卡 Or mEditType = Cr_换卡 Or mEditType = Cr_退卡) Then Exit Sub
    With cbo支付方式
        If .ListIndex = -1 Then Exit Sub
        lngIndex = .ListIndex + 1
    End With
    '短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not mcolPayMode Is Nothing Then
        With mCurPayMoney
            .lng医疗卡类别ID = Val(mcolPayMode(lngIndex)(3))
            .bln消费卡 = Val(mcolPayMode(lngIndex)(5)) = 1
            .str结算方式 = Trim(mcolPayMode(lngIndex)(6))
            .str名称 = Trim(mcolPayMode(lngIndex)(1))
         End With
    Else
        '86578:李南春,2015/7/16,刷卡结算方式
        mCurPayMoney.str结算方式 = zlstr.NeedName(cbo支付方式.Text)
    End If
    If Val(txt合计.Text) - Val(txt合计.Tag) >= 0 Then
        If cbo支付方式.Text = "支票" Then
            IDKindPayMode.Cards(1).名称 = "退支票"
        Else
            IDKindPayMode.Cards(1).名称 = "找补"
        End If
    End If
    Call txt余额_Change
End Sub

Private Sub cbo支付方式_KeyPress(KeyAscii As Integer)
     SearchCombox cbo支付方式, KeyAscii
End Sub
Private Sub cbo职业_Change()
    mblnChange = True
End Sub

Private Sub cbo职业_Click()
    mblnChange = True
End Sub

Private Sub cbo职业_KeyPress(KeyAscii As Integer)
    SearchCombox cbo职业, KeyAscii
End Sub

Private Sub chkCancel_Click()
    stbThis.Panels(2).Text = ""
    If mEditType <> Cr_发卡 And mEditType <> Cr_退卡 Then Exit Sub
    If mblnNotClick Then Exit Sub
    mblnNotClick = True
    If mEditType = Cr_退卡 Then chkCancel.value = 1
    mblnNotClick = False
    Load支付方式 (chkCancel.value = 1)
    lbl预交余额.Caption = "预交余额:0元"
    If mEditType <> Cr_退卡 Then pic预交余额.Top = wndTaskPanel.Height - picCard.Height - pic预交余额.Height - 180
    Call SetControlEnable: Call SetControlVisitble
    chkCancel.ForeColor = IIf(chkCancel.value = 1, &HFF, 0) '退为红色
    Call ClearData
    If chkCancel.value = Checked Then
        '待输入退款的单据号
        cboNO.Text = "": cboNO.Tag = "": cboNO.Locked = False
        chk病历费.Caption = "退病历费"
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
        If txt刷卡卡号.Visible And txt刷卡卡号.Enabled Then txt刷卡卡号.SetFocus
    Else
        Call LoadCardFee
        txtPatient.Text = ""
        chk病历费.Caption = "收病历费"
        txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm")
        cboNO.Locked = True
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    If chkCancel.value = 1 Then
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    End If
End Sub

Private Sub chk病历费_Click()
    Dim blnEnabled As Boolean
    If txt卡费.Visible = False And chk病历费.value = Unchecked Then
        blnEnabled = False
        txt合计.Text = "": txt合计.Tag = ""
        txt余额.Text = ""
    Else
        blnEnabled = True
        Call txt余额_Change
    End If
    
    If mEditType = Cr_退卡 Or chkCancel.value = Checked Then Exit Sub
    Call SetCardEditEnabled
    
    If gblnLED And chk记帐.value = 0 Then zl9LedVoice.Speak "#21 " & txt合计.Tag
End Sub

Private Sub chk记帐_Click()
    mblnChange = True
    cbo支付方式.Enabled = Not chk记帐.value = Checked
    cbo支付方式.BackColor = IIf(cbo支付方式.Enabled, &H80000005, &H8000000F)
    If mEditType = Cr_退卡 Or chkCancel.value = Checked Then Exit Sub
    txt合计.Enabled = Not chk记帐.value = Checked
    txt合计.BackColor = IIf(txt合计.Enabled, &H80000005, &H8000000F)
    IDKindPayMode.Enabled = Not chk记帐.value = Checked
    txt余额.Enabled = Not chk记帐.value = Checked
    txt余额.BackColor = IIf(txt余额.Enabled, &H80000005, &H8000000F)
    
    If chk记帐.value = Checked Then
        txt合计.Text = "": txt合计.Tag = ""
        txt余额.Text = ""
    Else
        Call txt余额_Change
    End If
    If gblnLED And chk记帐.value = 0 Then zl9LedVoice.Speak "#21 " & txt合计.Tag
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-07 03:47:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, blnNewPati As Boolean, Curdate As Date, lng结帐ID As Long
    Dim cllPro As Collection, cllUpdateSwap As Collection, cllThree As Collection
    Dim strErrMsg As String, strSQL As String
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then
        blnNewPati = True
    ElseIf mrsInfo.State <> 1 Then
        blnNewPati = True
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If

    Set cllPro = New Collection
    Set cllUpdateSwap = New Collection: Set cllThree = New Collection
    Curdate = zlDatabase.Currentdate
    If blnNewPati Then
        lng病人ID = zlDatabase.GetNextNo(1)
        Call AddNewPatiSQL(lng病人ID, Curdate, cllPro)
    Else
        If Not mbln存在门诊号 Then
            If txt门诊号.Text = "" Then
                '123098:李南春，2018/3/20，自动生成门诊号
                If mbln自动门诊号 Then
                    If MsgBox("病人门诊号为空,是否自动生成门诊号?", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
                        strSQL = "Zl_病人信息_绑定门诊号(" & lng病人ID & "," & zlGet门诊号 & ",to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & ")"
                        zlAddArray cllPro, strSQL
                    End If
                End If
            Else
                strSQL = "Zl_病人信息_绑定门诊号(" & lng病人ID & "," & txt门诊号.Text & ",to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & ")"
                zlAddArray cllPro, strSQL
            End If
        End If
    End If
    If AddCardDataSQL(lng病人ID, Curdate, cllPro, lng结帐ID) = False Then Exit Function
    If IDKindPayMode.IDKind = 2 And Val(txt余额.Text) > 0 Then Call AddDepositSQL(lng病人ID, Curdate, cllPro, lng结帐ID)
    '问题号:56599
    If lng病人ID > 0 Then Call Add健康卡相关信息(lng病人ID, cllPro)
    
    '90875:李南春,2016/1/22,保存病人证件信息
    If lng病人ID > 0 Then Call AddCertificate(lng病人ID, cllPro, Curdate)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    '110269:李南春,2016/10/10,保存HIS数据要提交EMPI数据，失败后所有数据都要回退
    If zlSaveEMPIPatiInfo(blnNewPati, lng病人ID, 0, strErrMsg) = False Then
        gcnOracle.RollbackTrans
        If strErrMsg = "" Then strErrMsg = "向EMPI平台上传病人信息失败！"
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    If zlInterfacePrayMoney(cllUpdateSwap, cllThree) = False Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    zlExecuteProcedureArrAy cllUpdateSwap, Me.Caption, False, True
    '问题号:53408
    '问题号:54172
    '问题号:54208
    If mCardType.str卡名称 = "二代身份证" And Not mrsInfo Is Nothing Then
        If mrsInfo.State <> 0 Then
            SaveModifyPati '修改病人信息（主要是为了跟新下身份证）
        End If
    End If
    Err = 0: On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllThree, Me.Caption
    '照片
    Select Case mlng图像操作
        Case 1 '文件
            SavePatPicture lng病人ID, cmdialog.FileName
        Case 2 '采集
            SavePatPicture lng病人ID, mstr采集图片
            mstr采集图片 = ""
        Case 4 '二代身份证
            mstrIDImageFile = App.Path & "\SFZIMG.bmp"
            SavePicture imgPatient.Picture, mstrIDImageFile
            SavePatPicture lng病人ID, mstrIDImageFile
        Case 3 '消除
            DeletePatPicture lng病人ID
    End Select
    '问题号:56599
    '院外发卡情况下需要进行写卡操作
    If mCardType.bln是否写卡 Then Call WriteCard(lng病人ID)
        
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '保存插件附加信息
        On Error Resume Next
        Call mobjPlugIn.PatiInfoSaveAfter(mlng病人ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    mbln存在门诊号 = False
    
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    Call ErrCenter
   
    Exit Function
ErrOthers:
    If ErrCenter = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    gcnOracle.CommitTrans
End Function
 
Private Function AddCardDataSQL(ByVal lng病人ID As Long, ByVal dtCurdate As Date, _
    ByRef cllPro As Collection, ByRef lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:就诊卡发放处理
    '入参:lng病人ID
    '编制:刘兴洪
    '日期:2011-07-07 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt操作类型 As Byte, strNO As String, str划价单 As String, strPassWord As String, strSQL As String
    Dim str原卡号 As String, str年龄 As String, strCard As String, str变动原因 As String
    Dim strICCard As String, lngBrushCardTypeID As Long, str结算方式 As String, strBrushCardNo As String
    Dim bln消费卡 As Boolean, blnInRange As Boolean   '范围内的卡
    Dim lngIndex As Long, lng执行部门ID As Long
    Dim byt变动类型 As Byte, dbl实收 As Double
    
     blnInRange = True
     lng结帐ID = 0
    
    If mCardType.blnOneCard And mCardType.bln严格控制 Then blnInRange = mCardType.lng领用ID > 0
    Select Case mEditType
    Case Cr_绑定卡
         blnInRange = False: byt操作类型 = 0
         byt变动类型 = 11
    Case Cr_发卡
         byt操作类型 = 0: byt变动类型 = 1
         If mCardType.rs医疗卡费 Is Nothing Then
             blnInRange = False
         End If
    Case Cr_补卡
         byt操作类型 = 1: byt变动类型 = 3
    Case Cr_换卡
        byt操作类型 = 2: blnInRange = False: byt变动类型 = 2
        '如果原卡,是存在卡费的,在换卡时,需要调用调用过程处理相应的,票据明细
        If Not mCardType.rs医疗卡费 Is Nothing Then
            blnInRange = True
        End If
    Case Else
        AddCardDataSQL = True: Exit Function
    End Select
    strCard = Trim(txt卡号.Text): strICCard = IIf(mblnICCard, strCard, "")
    
   
    
    str原卡号 = Trim(txt刷卡卡号.Text)
    lblNo.Tag = ""
    strPassWord = zlCommFun.zlStringEncode(Trim(txtPass.Text))
    mPrint.dtPrintdate = dtCurdate
    If blnInRange = False Then
          'Zl_医疗卡变动_Insert
           strSQL = "Zl_医疗卡变动_Insert("
          '      变动类型_In   Number,
          '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
          strSQL = strSQL & "" & byt变动类型 & ","
          '      病人id_In     住院费用记录.病人id%Type,
          strSQL = strSQL & "" & lng病人ID & ","
          '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
          strSQL = strSQL & "" & mlngCardTypeID & ","
          '      原卡号_In     病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & str原卡号 & "',"
          '      医疗卡号_In   病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      变动原因_In   病人医疗卡变动.变动原因%Type,
          '      --变动原因_In:如果密码调整，变动原因为密码.加密的
          strSQL = strSQL & "'" & str变动原因 & "',"
          '      密码_In       病人信息.卡验证码%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      操作员姓名_In 住院费用记录.操作员姓名%Type,
          strSQL = strSQL & "'" & UserInfo.姓名 & "',"
          '      变动时间_In   住院费用记录.登记时间%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
          strSQL = strSQL & "NULL)"
    Else
        If mEditType = Cr_换卡 Then
            strNO = GetCardFeeNo(mlngCardTypeID, Trim(txt刷卡卡号.Text), lng病人ID)
        Else
            strNO = zlDatabase.GetNextNo(16)  '医疗卡
        End If
        str划价单 = ""
        If gSystemPara.bln免挂号模式 And mEditType <> Cr_换卡 Then
           '免挂号模式，只存为划价单
            
           str划价单 = zlDatabase.GetNextNo(13)
           With mCardType.rs医疗卡费
              lng执行部门ID = zlGetCardFeeExcuteDeptID(Val(Nvl(!收费细目ID)), Val(Nvl(!科室标志)), UserInfo.部门ID)
           
              strSQL = "zl_门诊划价记录_Insert('" & str划价单 & "',1," & lng病人ID & ",NULL," & txt门诊号.Text & "," & _
                       "NULL,'" & txtPatient.Text & "','" & zlstr.NeedName(cbo性别.Text) & "','" & txt年龄.Text & cbo年龄单位.Text & "'," & _
                       "'" & zlstr.NeedName(cbo费别.Text) & "',0," & UserInfo.部门ID & "," & _
                       UserInfo.部门ID & ",'" & UserInfo.姓名 & "',NULL," & !收费细目ID & "," & _
                       "'" & !收费类别 & "','" & !计算单位 & "',NULL,1,1,0," & lng执行部门ID & ",NULL," & _
                       !收入项目ID & ",'" & !收据费目 & "'," & Format(!现价, "0.000") & "," & _
                       Format(!现价, "0.00") & "," & IIf(mCardType.bln变价 = False, Format(!现价, "0.00"), Val(txt卡费.Text)) & "," & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & "," & _
                       "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & ",NULL,'" & UserInfo.姓名 & "','" & strNO & "')"
            End With
            
            zlAddArray cllPro, strSQL
        End If
        
        lblNo.Tag = strNO
        If chk记帐.value = 0 Then
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
        End If
        
        mCurPayMoney.lng结帐ID = lng结帐ID
        mCurPayMoney.strNO = strNO
        
        If cbo支付方式.ItemData(cbo支付方式.ListIndex) < 0 Then
            lngIndex = cbo支付方式.ListIndex + 1
            lngBrushCardTypeID = mcolPayMode(lngIndex)(3)
            '短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
            lngBrushCardTypeID = Val(mcolPayMode(lngIndex)(3))
            bln消费卡 = Val(mcolPayMode(lngIndex)(5)) = 1
        Else
            bln消费卡 = False
        End If
        
        '103980:李南春，2016/12/15，保存医疗卡费用信息时保存年龄信息
        str年龄 = Trim(txt年龄.Text)
        If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
        
        dbl实收 = Val(txt卡费.Text)
        
        If gSystemPara.bln免挂号模式 Then
            dbl实收 = 0: str结算方式 = "现金"
        ElseIf mEditType <> Cr_换卡 Then
            '86578:李南春,2015/7/16,刷卡结算方式
            str结算方式 = mcolPayMode(cbo支付方式.ListIndex + 1)(6)
            If str结算方式 = "" Then str结算方式 = zlstr.NeedName(cbo支付方式.Text)
            If Not cbo支付方式.Enabled Then
                str结算方式 = ""
            ElseIf cbo支付方式.Text <> mCurPayMoney.str名称 Then
                MsgBox "支付方式错误，请重新选择支付方式。", vbInformation, gstrSysName
                zlControl.ControlSetFocus cbo支付方式: Exit Function
            End If
        End If
        
        strSQL = zlGetSaveCardFeeSQL(mlngCardTypeID, byt操作类型, strNO, lng病人ID, 0, UserInfo.部门ID, UserInfo.部门ID, 0, _
        zlstr.NeedName(cbo费别.Text), str原卡号, Trim(txtPatient.Text), zlstr.NeedName(cbo性别.Text), str年龄, _
        strCard, strPassWord, str变动原因, IIf(mCardType.bln变价 = False, mCardType.dbl应收金额, Val(txt卡费.Text)), dbl实收, str结算方式, _
         dtCurdate, mCardType.lng领用ID, mCardType.rs医疗卡费, strICCard, mCurPayMoney.lng医疗卡类别ID, mCurPayMoney.bln消费卡, mCurPayMoney.str刷卡卡号, mCurPayMoney.lng结帐ID, _
         str划价单, IIf(chk病历费.value, Val(txt合计.Tag), 0))
    End If
    
    zlAddArray cllPro, strSQL
    
    '95809
    If chk病历费.value And Not mFeeType.rs病历费 Is Nothing Then
        If strNO = "" Then strNO = zlDatabase.GetNextNo(16) '医疗卡
        lblNo.Tag = strNO
        If lng结帐ID = 0 And chk记帐.value = 0 Then
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
            
            str结算方式 = mcolPayMode(cbo支付方式.ListIndex + 1)(6)
            If str结算方式 = "" Then str结算方式 = zlstr.NeedName(cbo支付方式.Text)
            If str结算方式 <> mCurPayMoney.str名称 Then
                MsgBox "支付方式错误，请重新选择支付方式。", vbInformation, gstrSysName
                zlControl.ControlSetFocus cbo支付方式: Exit Function
            End If
        End If
        mCurPayMoney.strNO = strNO
        mCurPayMoney.lng结帐ID = lng结帐ID
        
        strSQL = zlGetSaveCardFeeSQL(0, IIf(mEditType = Cr_绑定卡 Or mEditType = Cr_换卡, 9, 8), strNO, lng病人ID, 0, UserInfo.部门ID, UserInfo.部门ID, 0, _
        zlstr.NeedName(cbo费别.Text), "", Trim(txtPatient.Text), zlstr.NeedName(cbo性别.Text), str年龄, _
        "", "", "", IIf(mCardType.bln变价 = False, mFeeType.dbl应收金额, Val(txt病历费.Text)), Val(txt病历费.Text), str结算方式, _
        dtCurdate, 0, mFeeType.rs病历费, "", mCurPayMoney.lng医疗卡类别ID, mCurPayMoney.bln消费卡, mCurPayMoney.str刷卡卡号, mCurPayMoney.lng结帐ID)
        
        zlAddArray cllPro, strSQL
    End If
    AddCardDataSQL = True
End Function

Private Function GetCardFeeNo(ByVal lng卡类别ID As Long, ByVal strCard As String, ByVal lng病人ID As Long) As String
    '获取指定卡号的费用NO
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select No From 住院费用记录 Where 记录性质 = 5 And 记录状态 = 1 And Nvl(结论,0) = [1] And 实际票号 = [2] And 病人ID = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "发卡记录", lng卡类别ID, strCard, lng病人ID)
    If Not rsTmp.EOF Then GetCardFeeNo = Nvl(rsTmp!NO)
End Function
 
Private Sub AddPrintSQL(ByVal byt操作类型 As Byte, ByVal strNO As String, ByVal dtCurdate As Date, ByRef cllPro As Collection)
    Dim strSQL As String
    If gbln收费发票 = False Then Exit Sub '使用门诊收据
    If mPrint.bytPrintType = 0 Then Exit Sub '票据允许打印
    If Trim(txtFact.Text) = "" Then Exit Sub '没有打印票据
    
    strSQL = "Zl_病人发卡票据_Print("
    '  No_In           Varchar2,
    strSQL = strSQL & "'" & strNO & "'" & ","
    '  票据号_In       票据使用明细.号码%Type,
    strSQL = strSQL & "'" & Trim(txtFact.Text) & "',"
    '  领用id_In       票据使用明细.领用id%Type,
    strSQL = strSQL & "" & ZVal(mPrint.lng领用ID) & ","
    '  使用人_In       票据使用明细.使用人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  使用时间_In     票据使用明细.使用时间%Type,
    strSQL = strSQL & "To_Date('" & Format(dtCurdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  操作类型_In     Number
    strSQL = strSQL & byt操作类型 & ","
    '  票据张数_In     Number := 1,
    strSQL = strSQL & "" & 1 & ")"
    zlAddArray cllPro, strSQL
End Sub

Private Function AddNewPatiSQL(ByVal lng病人ID As Long, ByVal dtCurdate As Date, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存新病人数据
    '出参:cllPro-过程集
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-07 04:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str年龄 As String, str出生日期 As String
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
       
    '  Zl_病人信息_Insert
    strSQL = "Zl_病人信息_Insert("
    '  病人id_In     病人信息.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  门诊号_In     病人信息.门诊号%Type,
    strSQL = strSQL & "" & IIf(Trim(txt门诊号.Text) <> "", Val(txt门诊号.Text), "NULL") & ","
    '  费别_In       病人信息.费别%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo费别.Text) & "',"
    '  医疗付款_In   病人信息.医疗付款方式%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo医疗付款.Text) & "',"
    '  姓名_In       病人信息.姓名%Type,
    strSQL = strSQL & "'" & txtPatient.Text & "',"
    '  性别_In       病人信息.性别%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo性别.Text) & "',"
    '  年龄_In       病人信息.年龄%Type,
    strSQL = strSQL & "'" & str年龄 & "',"
    '  出生日期_In   病人信息.出生日期%Type,
    strSQL = strSQL & "" & str出生日期 & ","
    '  出生地点_In   病人信息.出生地点%Type,
    strSQL = strSQL & "'" & txt出生地点.Text & "',"
    '  身份证号_In   病人信息.身份证号%Type,
    strSQL = strSQL & "'" & txt身份证号.Text & "',"
    '  身份_In       病人信息.身份%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo身份.Text) & "',"
    '  职业_In       病人信息.职业%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo职业.Text, mstrCboSplit) & "',"
    '  民族_In       病人信息.民族%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo民族.Text) & "',"
    '  国籍_In       病人信息.国籍%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo国籍.Text) & "',"
    '  学历_In       病人信息.学历%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo学历.Text) & "',"
    '  婚姻_In       病人信息.婚姻状况%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo婚姻状况.Text) & "',"
    '  家庭地址_In   病人信息.家庭地址%Type,
    strSQL = strSQL & "'" & IIf(mblnStructAdress, padd家庭地址.value, txt家庭地址.Text) & "',"
    '  家庭电话_In   病人信息.家庭电话%Type,
    strSQL = strSQL & "'" & txt家庭电话.Text & "',"
    '  家庭地址邮编_In   病人信息.家庭地址邮编%Type,
    strSQL = strSQL & "'" & txt家庭邮编.Text & "',"
    '  联系人姓名_In 病人信息.联系人姓名%Type,
    strSQL = strSQL & "'" & txt联系人姓名.Text & "',"
    '  联系人关系_In 病人信息.联系人关系%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo联系人关系.Text) & "',"
    '  联系人地址_In 病人信息.联系人地址%Type,
    strSQL = strSQL & "'" & txt联系人地址.Text & "',"
    '  联系人电话_In 病人信息.联系人电话%Type,
    strSQL = strSQL & "'" & txt联系人电话.Text & "',"
    '  合同单位id_In 病人信息.合同单位id%Type,
    strSQL = strSQL & "" & IIf(Val(lbl工作单位.Tag) = 0, "NULL", Val(lbl工作单位.Tag)) & ","
    '  工作单位_In   病人信息.工作单位%Type,
    strSQL = strSQL & "'" & txt工作单位.Text & "',"
    '  单位电话_In   病人信息.单位电话%Type,
    strSQL = strSQL & "'" & txt单位电话.Text & "',"
    '  单位邮编_In   病人信息.单位邮编%Type,
    strSQL = strSQL & "'" & txt单位邮编.Text & "',"
    '  单位开户行_In 病人信息.单位开户行%Type,
    strSQL = strSQL & "'" & txt单位开户行.Text & "',"
    '  单位帐号_In   病人信息.单位帐号%Type,
    strSQL = strSQL & "'" & txt单位帐户.Text & "',"
    '  担保人_In     病人信息.担保人%Type,
    strSQL = strSQL & "null,"
    '  担保额_In     病人信息.担保额%Type,
    strSQL = strSQL & "null,"
    '  险类_In       病人信息.险类%Type,
    strSQL = strSQL & "null,"
    '  登记时间_In   病人信息.登记时间%Type,
    strSQL = strSQL & "To_Date('" & Format(dtCurdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  区域_In       病人信息.区域%Type := Null,
    strSQL = strSQL & "'" & zlstr.NeedName(txt区域.Text) & "',"
    '  担保性质_In   病人信息.担保性质%Type := Null,
    strSQL = strSQL & "null,"
    '  操作员编号_In 病人担保记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人担保记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  医保号_In     病人信息.医保号%Type := Null,
    strSQL = strSQL & "" & IIf(Trim(txt医保号.Text) = "", "NULL", "'" & Trim(txt医保号.Text) & "'") & ","
    '  其他证件_In   病人信息.其他证件%Type := Null
    strSQL = strSQL & "'" & txt其他证件.Text & "',"
    '问题号:51071
    '  籍贯_In   病人信息.籍贯%Type := Null
    strSQL = strSQL & "'',"
    '  户口地址_In   病人信息.户口地址%Type := Null
    strSQL = strSQL & "'" & IIf(mblnStructAdress, Trim(padd户口地址.value), Trim(txt户口地址.Text)) & "',"
    '  户口地址邮编_In   病人信息.户口地址邮编%Type := Null
    strSQL = strSQL & "'" & Trim(txt户口地址邮编.Text) & "',"
    '  联系人身份证号_In   病人信息.联系人身份证号%Type := Null
    strSQL = strSQL & "'" & Trim(txt联系人身份证号.Text) & "',"
    '  病人类型_In   病人信息.病人类型%Type := Null
    strSQL = strSQL & "'',"
    '  监护人_In     病人信息.监护人%Type := Null
    strSQL = strSQL & "'',"
    '  手机号_In     病人信息.手机号%Type := Null
    strSQL = strSQL & "'" & txt手机.Text & "')"
    zlAddArray cllPro, strSQL
    
    '89242:李南春,2015/12/11,更新病人地址信息
    If Not mblnStructAdress Then Exit Function
    If padd家庭地址.Enabled Then
        If padd家庭地址.value <> "" Then
           strSQL = "zl_病人地址信息_update(1," & lng病人ID & ",NULL,3,'" & padd家庭地址.value省 & "','" & _
               padd家庭地址.value市 & "','" & padd家庭地址.value区县 & "','" & padd家庭地址.value乡镇 & "','" & _
               padd家庭地址.value详细地址 & "','" & padd家庭地址.Code & "')"
        Else
           strSQL = "zl_病人地址信息_update(2," & lng病人ID & ",NULL,3)"
        End If
        zlAddArray cllPro, strSQL
    End If
    If padd户口地址.Enabled Then
        If padd户口地址.value <> "" Then
           strSQL = "zl_病人地址信息_update(1," & lng病人ID & ",NULL,4,'" & padd户口地址.value省 & "','" & _
               padd户口地址.value市 & "','" & padd户口地址.value区县 & "','" & padd户口地址.value乡镇 & "','" & _
               padd户口地址.value详细地址 & "','" & padd户口地址.Code & "')"
        Else
           strSQL = "zl_病人地址信息_update(2," & lng病人ID & ",NULL,4)"
        End If
        zlAddArray cllPro, strSQL
    End If
End Function
Private Function SaveModifyPati() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改病人信息
    '返回:修改成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-07 03:48:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str年龄 As String, str出生日期 As String, str其他关系 As String
    Dim blnTrans As Boolean, strErrMsg As String
    Dim str家庭地址 As String, str户口地址 As String
    On Error GoTo errHandle
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
    
    '    Zl_病人信息_Update
    strSQL = "Zl_病人信息_Update("
    '      病人id_In     病人信息.病人id%Type,
    strSQL = strSQL & "" & mrsInfo!病人ID & ","
    '      门诊号_In     病人信息.门诊号%Type,
    strSQL = strSQL & "" & IIf(Trim(txt门诊号.Text) <> "", Val(txt门诊号.Text), "NULL") & ","
    '      住院号_In     病人信息.住院号%Type,
    strSQL = strSQL & "" & IIf(Val(Nvl(mrsInfo!住院号)) = 0, "NULL", Val(Nvl(mrsInfo!住院号))) & ","
    '      费别_In       病人信息.费别%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo费别.Text) & "',"
    '      医疗付款_In   病人信息.医疗付款方式%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo医疗付款.Text) & "',"
    '      姓名_In       病人信息.姓名%Type,
    strSQL = strSQL & "'" & txtPatient.Text & "',"
    '      性别_In       病人信息.性别%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo性别.Text) & "',"
    '      年龄_In       病人信息.年龄%Type,
    strSQL = strSQL & "'" & str年龄 & "',"
    '      出生日期_In   病人信息.出生日期%Type,
    strSQL = strSQL & "" & str出生日期 & ","
    '      出生地点_In   病人信息.出生地点%Type,
    strSQL = strSQL & "'" & txt出生地点.Text & "',"
    '      身份证号_In   病人信息.身份证号%Type,
    strSQL = strSQL & "'" & txt身份证号.Text & "',"
    '      身份_In       病人信息.身份%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo身份.Text) & "',"
    '      职业_In       病人信息.职业%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo职业.Text, mstrCboSplit) & "',"
    '      民族_In       病人信息.民族%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo民族.Text) & "',"
    '      国籍_In       病人信息.国籍%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo国籍.Text) & "',"
    '      学历_In       病人信息.学历%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo学历.Text) & "',"
    '      婚姻_In       病人信息.婚姻状况%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo婚姻状况.Text) & "',"
    '      家庭地址_In   病人信息.家庭地址%Type,
    strSQL = strSQL & "'" & IIf(mblnStructAdress, padd家庭地址.value, txt家庭地址.Text) & "',"
    '      家庭电话_In   病人信息.家庭电话%Type,
    strSQL = strSQL & "'" & txt家庭电话.Text & "',"
    '      家庭地址邮编_In   病人信息.家庭地址邮编%Type,
    strSQL = strSQL & "'" & txt家庭邮编.Text & "',"
    '      联系人姓名_In 病人信息.联系人姓名%Type,
    strSQL = strSQL & "'" & txt联系人姓名.Text & "',"
    '      联系人关系_In 病人信息.联系人关系%Type,
    strSQL = strSQL & "'" & zlstr.NeedName(cbo联系人关系.Text) & "',"
    '      联系人地址_In 病人信息.联系人地址%Type,
    strSQL = strSQL & "'" & txt联系人地址.Text & "',"
    '      联系人电话_In 病人信息.联系人电话%Type,
    strSQL = strSQL & "'" & txt联系人电话.Text & "',"
    '      合同单位id_In 病人信息.合同单位id%Type,
    strSQL = strSQL & "" & IIf(Val(lbl工作单位.Tag) = 0, "NULL", Val(lbl工作单位.Tag)) & ","
    '      工作单位_In   病人信息.工作单位%Type,
    strSQL = strSQL & "'" & txt工作单位.Text & "',"
    '      单位电话_In   病人信息.单位电话%Type,
    strSQL = strSQL & "'" & txt单位电话.Text & "',"
    '      单位邮编_In   病人信息.单位邮编%Type,
    strSQL = strSQL & "'" & txt单位邮编.Text & "',"
    '      单位开户行_In 病人信息.单位开户行%Type,
    strSQL = strSQL & "'" & txt单位开户行.Text & "',"
    '      单位帐号_In   病人信息.单位帐号%Type,
    strSQL = strSQL & "'" & txt单位帐户.Text & "',"
    '      担保人_In     病人信息.担保人%Type,
    strSQL = strSQL & "'" & Nvl(mrsInfo!担保人) & "',"
    '      担保额_In     病人信息.担保额%Type,
    strSQL = strSQL & "" & Val(Nvl(mrsInfo!担保额)) & ","
    '      险类_In       病人信息.险类%Type,
    strSQL = strSQL & "" & IIf(Val(Nvl(mrsInfo!险类)) = 0, "NULL", Val(Nvl(mrsInfo!险类))) & ","
    '      住院费别_In   Number := 0, --是否修改的是病人的住院费别
    strSQL = strSQL & "0,"
    '      医保号_In     保险帐户.医保号%Type := Null,
    strSQL = strSQL & "" & IIf(Trim(txt医保号.Text) = "", "NULL", "'" & Trim(txt医保号.Text) & "'") & ","
    '      区域_In       病人信息.区域%Type := Null,
    strSQL = strSQL & "'" & zlstr.NeedName(txt区域.Text) & "',"
    '      担保性质_In   病人信息.担保性质%Type := Null,
    strSQL = strSQL & "" & Val(Nvl(mrsInfo!担保性质)) & ","
    '      操作员编号_In 病人担保记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '      操作员姓名_In 病人担保记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '      其他证件_In   病人信息.其他证件%Type := Null,
    strSQL = strSQL & "'" & txt其他证件.Text & "',"
    '      病人类型_In   病案主页.病人类型%Type := Null,
    strSQL = strSQL & "'" & Nvl(mrsInfo!病人类型) & "',"
    '      备注_In       病案主页.备注%Type := Null
     strSQL = strSQL & "'" & Nvl(mrsInfo!备注) & "',"
    '问题号:51071
    '  籍贯_In   病人信息.籍贯%Type := Null
    strSQL = strSQL & "'',"
    '  户口地址_In   病人信息.户口地址%Type := Null
    strSQL = strSQL & "'" & IIf(mblnStructAdress, Trim(padd户口地址.value), Trim(txt户口地址.Text)) & "',"
    '  户口地址邮编_In   病人信息.户口地址邮编%Type := Null
    strSQL = strSQL & "'" & Trim(txt户口地址邮编.Text) & "',"
     '     联系人身份证号_In       病人信息.联系人身份证号%Type := Null WJ
    strSQL = strSQL & "'" & Trim(txt联系人身份证号.Text) & "',"
    '   模块号_In         Number := 0 --修改病人姓名、性别、年龄、出生日期的模块
    strSQL = strSQL & "" & mlngModule & ","
    '  监护人_In         病人信息.监护人%Type :=Null
    strSQL = strSQL & "" & "'',"
    '  手机号_In         病人信息.手机号%Type :=Null
    strSQL = strSQL & "'" & txt手机.Text & "')"
    
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    '其他关系
    If txt联系人姓名.Text <> "" And txt其他关系.Visible Then
        str其他关系 = "Zl_病人信息从表_Update("
        '病人ID_In 病人信息从表.病人Id%Type
        str其他关系 = str其他关系 & "" & mrsInfo!病人ID & ","
        '信息名_In 病人信息从表.信息名%Type
        str其他关系 = str其他关系 & "'联系人附加信息',"
        '信息值_In 病人信息从表.信息值%Type
        str其他关系 = str其他关系 & "'" & txt其他关系.Text & "',"
        '就诊Id_In 病人信息从表.就诊Id%Type
        str其他关系 = str其他关系 & "'')"
    End If
    
    '89242:李南春,2015/12/10,更新病人地址信息
    '家庭地址
     If mblnStructAdress And padd家庭地址.Enabled Then
        If padd家庭地址.value <> "" Then
           str家庭地址 = "zl_病人地址信息_update(1," & mrsInfo!病人ID & ",NULL,3,'" & padd家庭地址.value省 & "','" & _
               padd家庭地址.value市 & "','" & padd家庭地址.value区县 & "','" & padd家庭地址.value乡镇 & "','" & _
               padd家庭地址.value详细地址 & "','" & padd家庭地址.Code & "')"
        Else
           str家庭地址 = "zl_病人地址信息_update(2," & mrsInfo!病人ID & ",NULL,3)"
        End If
    End If
    '户口地址
    If mblnStructAdress And padd户口地址.Enabled Then
        If padd户口地址.value <> "" Then
           str户口地址 = "zl_病人地址信息_update(1," & mrsInfo!病人ID & ",NULL,4,'" & padd户口地址.value省 & "','" & _
               padd户口地址.value市 & "','" & padd户口地址.value区县 & "','" & padd户口地址.value乡镇 & "','" & _
               padd户口地址.value详细地址 & "','" & padd户口地址.Code & "')"
        Else
           str户口地址 = "zl_病人地址信息_update(2," & mrsInfo!病人ID & ",NULL,4)"
        End If
    End If
    
    '81103,冉俊明,2014-12-29,录入身份证号后,出生日期、年龄、性别的同步关联检查和调整
    '101929:李南春,2016/10/27,放在最开始执行，否则无法记录变动信息
    If mbln基本信息调整 Then
        Call mobjPubPatient.SavePatiBaseInfo(mrsInfo!病人ID, 0, Trim(txtPatient.Text), _
                                    zlstr.NeedName(cbo性别.Text), str年龄, txt出生日期.Text, "医疗卡管理", 1, strErrMsg, , Not mrsEMPIOut Is Nothing)
    End If
    
    blnTrans = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If str其他关系 <> "" Then zlDatabase.ExecuteProcedure str其他关系, Me.Caption
    If str家庭地址 <> "" Then zlDatabase.ExecuteProcedure str家庭地址, Me.Caption
    If str户口地址 <> "" Then zlDatabase.ExecuteProcedure str户口地址, Me.Caption
    
    '110269:李南春,2016/10/10,保存HIS数据要提交EMPI数据，失败后所有数据都要回退
    strErrMsg = ""
    If zlSaveEMPIPatiInfo(False, mrsInfo!病人ID, 0, strErrMsg) = False Then
        gcnOracle.RollbackTrans
        If strErrMsg = "" Then strErrMsg = "向EMPI平台上传病人信息失败！"
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    SaveModifyPati = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LoadSaveNotoCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载单号据号给Combox
    '编制:刘兴洪
    '日期:2011-07-12 18:38:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTmp As String
    If Not (mEditType = Cr_发卡 And lblNo.Tag <> "") Then Exit Sub
    '加入单据历史记录(所有类型单据)
    For i = 0 To cboNO.ListCount - 1
        strTmp = strTmp & "," & cboNO.List(i)
    Next
    strTmp = lblNo.Tag & strTmp
    stbThis.Panels(2).Text = "上次保存单据:" & lblNo.Tag
    cboNO.Clear
    For i = 0 To UBound(Split(strTmp, ","))
        cboNO.AddItem Split(strTmp, ",")(i)
        If i = 9 Then Exit For '只显示10个
    Next
End Sub
Private Function IsCheckCancelValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费时的数据有效性
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String
    Dim str验证卡号  As String
    
    On Error GoTo errHandle
    strName = IIf(glngSys \ 100 = 8, "会员卡", "医疗卡")
    
    If cboNO.Tag = "" Then
        MsgBox "该" & strName & "发放记录未正确读取,不能退卡！", vbExclamation, gstrSysName
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Exit Function
    End If
    
    If InStr(1, "12", mParaData.int退卡模式) > 0 And txt刷卡卡号.Visible Then
        str验证卡号 = Trim(txt卡号.Text)
        If Trim(txt刷卡卡号) = "" Or str验证卡号 <> Trim(lbl刷卡验证.Tag) Then
            If mParaData.int退卡模式 = 1 Then
                MsgBox "退卡验证失败，必须刷卡验证！", vbExclamation, gstrSysName
                If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
            Else
                MsgBox "退卡验证失败，请核对实际卡号与当前单据卡号是否一致！", vbExclamation, gstrSysName
                If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
            End If
            Exit Function
        End If
    End If
    IsCheckCancelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckCardDelValied() As Boolean
    On Error GoTo errHandle
    Dim bln消费卡 As Boolean, lng卡类别ID As Long
    Dim str验证卡号  As String, dblMoney As Double
   '问题:48249
    Dim strSQL As String, rsBill As Recordset, rsTemp As ADODB.Recordset, lngCardBill As Long
    Dim intStyle As Integer, bln退卡 As Boolean
    Dim str卡号 As String, str交易流水号 As String, str交易说明 As String, str结算信息 As String
    Dim strXMLExpend As String
    Dim cllSquareBalance As Collection, blnErrCount As Boolean
    
    On Error GoTo errHandle
    '81839: 李南春，2015/1/19，医疗卡退卡检查
    intStyle = Val(zlDatabase.GetPara("已结帐单据操作", 100))
    strSQL = "Select B.NO From 住院费用记录 a,病人结帐记录 b Where a.结帐id=b.id And a.记录性质 In (5,15) And a.记录状态=1 And b.记录状态=1 And a.no=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, cboNO.Text)
    If rsTemp.EOF Then bln退卡 = True
    Select Case intStyle
        Case 0
            bln退卡 = True
        Case 1
            If bln退卡 = False Then
                If MsgBox("卡号" & txt卡号.Text & "的记账单已做结账处理，是否继续退卡", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    bln退卡 = True
                End If
            End If
        Case 2
            If bln退卡 = False Then
                MsgBox "卡号" & txt卡号.Text & "的记账单已做结账处理，不允许退卡", vbInformation + vbOKOnly, gstrSysName
            End If
    End Select
    If bln退卡 = False Then Exit Function
    
    '如果选择了其他方式退费，不再调用接口
    If cbo支付方式.ItemData(cbo支付方式.ListIndex) < 7 Then CheckCardDelValied = True: Exit Function
    If mcolBillBalance Is Nothing Then CheckCardDelValied = True: Exit Function
    
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID,结算方式,消费卡ID
    lng卡类别ID = mcolBillBalance(1)(0)
    If lng卡类别ID = 0 Then CheckCardDelValied = True: Exit Function
    
    str卡号 = mcolBillBalance(1)(1)
    bln消费卡 = Val(mcolBillBalance(1)(2)) = 1
    str交易流水号 = mcolBillBalance(1)(3)
    str交易说明 = mcolBillBalance(1)(4)
    str结算信息 = "5|" & mcolBillBalance(1)(6)
    dblMoney = Val(txt卡费.Text) + IIf(chk病历费.value = Checked, Val(txt病历费.Text), 0)
    
    Set cllSquareBalance = New Collection
    'Array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
    cllSquareBalance.Add Array(lng卡类别ID, mcolBillBalance(1)(8), 0, str卡号, "", "", False, dblMoney)
    
    '不为零,需要获取相应的支付对象
    Set mobjDelObject = zlGetClsCardObject(lng卡类别ID, bln消费卡)
    '92895:李南春,2016/1/21,未启用对象是nothing
    If mobjDelObject Is Nothing Then
        MsgBox "你未启用发卡时使用的支付接口 ,不能在此工作站上进行退卡!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Not mobjDelObject.CardPreporty.启用 Then
        MsgBox "你未启用" & mobjDelObject.CardPreporty.名称 & "接口 ,不能在此工作站上进行退卡!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjDelObject.CardObject Is Nothing Then
        If zlCreatePatiCardObject(mobjDelObject.CardPreporty, mobjDelObject.CardObject) = False Then
            Exit Function
        End If
    End If
    If Not mobjDelObject.InitCompents Then
        If mobjDelObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") = False Then
              Exit Function
        End If
        mobjDelObject.InitCompents = True
    End If
     
    '4.3.3.2.6   zlReturnCheck:帐户回退交易前的检查
    'zlPaymentCheck帐户扣款交易检查
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  模块号
    'lngCardTypeID   Long    In  卡类别ID:医疗卡类别.ID
    'strCardNo   String  IN  卡号
    'strBalanceIDs:格式:收费类型( 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款)|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    'dblMoney    Double  IN  退款金额
    'strSwapNo   String  In  交易流水号(退款时检查)
    'strSwapMemo String  In  交易说明(退款时传入)
    '    Boolean 函数返回    True:调用成功,False:调用失败
    '说明:
    '在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,以便控制死锁情况。
    If mobjDelObject.CardObject.zlReturncheck(Me, mlngModule, lng卡类别ID, str卡号, str结算信息, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus: Exit Function
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Function
        Exit Function
    End If
    
    '100610:李南春,2016/10/13，预交退款和余额退款是否验证刷卡
    If mobjDelObject.CardPreporty.消费卡 = False And mobjDelObject.CardPreporty.是否退款验卡 Then
    '   zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-三方卡调用XML入参,目前格式如下:
        '       <IN>
        '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        '       </IN>
        Err = 0: On Error Resume Next
        If mobjDelObject.CardObject.zlBrushCard(Me, mlngModule, lng卡类别ID, _
         Nvl(mrsInfo!姓名), Nvl(mrsInfo!性别), Nvl(mrsInfo!年龄), dblMoney, _
         mCurPayMoney.str刷卡卡号, mCurPayMoney.str刷卡密码, "<IN><CZLX>2</CZLX></IN>") = False Then
            If Err = 450 Then
                Err = 0: On Error GoTo errHandle
                If mobjDelObject.CardObject.zlBrushCard(Me, mlngModule, lng卡类别ID, _
                 Nvl(mrsInfo!姓名), Nvl(mrsInfo!性别), Nvl(mrsInfo!年龄), Val(txt卡费.Text), mCurPayMoney.str刷卡卡号, mCurPayMoney.str刷卡密码) = False Then Exit Function
            Else
                Exit Function
            End If
        End If
    ElseIf mobjDelObject.CardPreporty.消费卡 And mobjDelObject.CardPreporty.自制卡 And gbln消费卡退费验卡 Then
        Err = 0: On Error Resume Next
        If IsEmpty(cllSquareBalance) Then   '57682
            Set cllSquareBalance = Nothing
        End If
        blnErrCount = cllSquareBalance.count
        If Err <> 0 Then
            Set cllSquareBalance = Nothing
            Err = 0: On Error GoTo 0
        End If
        If frmInputPass.zlBrushPay(Me, mlngModule, mobjDelObject, Nothing, _
            lng卡类别ID, True, Nvl(mrsInfo!姓名), Nvl(mrsInfo!性别), Nvl(mrsInfo!年龄), dblMoney, _
            mCurPayMoney.str刷卡卡号, mCurPayMoney.str刷卡密码, True, True, False, False, cllSquareBalance) = False Then Exit Function
    End If
    CheckCardDelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsCheckCancel退预交()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退卡时检查病人是否有预交款未退
     '返回:有效,返回true,否则返回False
    '编制:王吉
    '日期:2012-07-16 18:50:36
    '问题号:51537
    '问题号:50891
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim msgBoxResult As String
    Dim strSQL As String
    Dim blnOneCard As Boolean   '是否是唯一一张医疗卡
    Dim rsBill As Recordset, rsCard As Recordset
    '69483,刘尔旋,2014-01-15,病人医疗卡退卡退款处理
    strSQL = "Select Count(1) As 医疗卡数 From 病人医疗卡信息 Where 状态=0 And 病人ID=[1]"
    Set rsCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    strSQL = _
            "Select 预交余额,费用余额 From 病人余额 Where 性质=1 And 类型=1 And 病人ID=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    '问题:48249
    If InStr(1, mstrPrepayPrivs, ";预交退款;") > 0 Then
        If rsBill.RecordCount > 0 Then
            If Format(Nvl(rsBill!预交余额, 0) - Nvl(rsBill!费用余额, 0), "0.00") > 0 Then
                '问题号:51537
                '问题号:50891
                '108836：李南春，2017/6/28，调整退卡描述
                msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, vbNewLine & "该病人尚有预交余额未退,是否先进行余额退款再退卡?" & vbNewLine, "退余额再退卡,仅退卡,取消", Me, vbQuestion)
                If msgBoxResult = "退余额再退卡" Then '退预交余额操作
                   '检查该卡是否是记账收费(退卡退余额时应该把记账的费用算到病人余额中一起退给病人)
                    '病人余额退款
                    '问题号:112995,焦博,2017/10/13,退卡退费时提示病人退费金额
                     blnOneCard = IIf(rsCard!医疗卡数 = 1, True, False)
                     IsCheckCancel退预交 = zlPrepayFunc(2, mlng病人ID, blnOneCard)
                     Exit Function
                ElseIf msgBoxResult = "取消" Or msgBoxResult = "" Then
                     IsCheckCancel退预交 = False
                     Exit Function
                ElseIf msgBoxResult = "仅退卡" Then
                    If rsCard!医疗卡数 = 1 Then
                        MsgBox "该病人尚有预交余额，不能对病人唯一的医疗卡进行退卡操作!", vbInformation, gstrSysName
                        IsCheckCancel退预交 = False
                        Exit Function
                    End If
                End If
            Else
            '问题号:51537
            '问题号:50891
                msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "您确定要进行退卡操作吗?", "退卡,取消", Me, vbQuestion)
                If msgBoxResult = "" Or msgBoxResult = "取消" Then
                    IsCheckCancel退预交 = False
                    Exit Function
                End If
            End If
        Else
        '问题号:51537
        '问题号:50891
           msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "您确定要进行退卡操作吗?", "退卡,取消", Me, vbQuestion)
           If msgBoxResult = "" Or msgBoxResult = "取消" Then
                IsCheckCancel退预交 = False
                Exit Function
           End If
        End If
    Else
        If rsBill.RecordCount > 0 Then
            If Format(Nvl(rsBill!预交余额, 0) - Nvl(rsBill!费用余额, 0), "0.00") > 0 Then
                If rsCard!医疗卡数 = 1 Then
                    MsgBox "您没有预交退款权限，不能对病人唯一的医疗卡退卡操作!", vbInformation, gstrSysName
                    IsCheckCancel退预交 = False
                    Exit Function
                End If
            End If
        End If
        If MsgBox("您没有预交退款权限,是否继续进行退卡操作?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then IsCheckCancel退预交 = False: Exit Function
    End If
        IsCheckCancel退预交 = True
End Function

Private Function SaveDelete(strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退卡号
    '入参:strNO-具体的单据号
    '返回:退号成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-07-12 18:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, blnTrans As Boolean, bln消费卡 As Boolean, lng卡类别ID As Long
    Dim lng结帐ID As Long, blnOraclTrans As Boolean, Index As Integer, cllPro As Collection
    Dim dtCurdate As Date
    On Error GoTo errH
    Index = cbo支付方式.ListIndex + 1
    '104726:李南春,2017/4/24,退卡时收回门诊收据
    Set cllPro = New Collection
    dtCurdate = zlDatabase.Currentdate
    '如果不是原样退
    If mcolPayMode(Index)(6) <> mcolBillBalance(1)(7) Then
        strSQL = "zl_医疗卡记录_DELETE('" & strNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(chk病历费.value, 1, 0) & ",'" & mcolPayMode(Index)(6) & "')"
    Else
        strSQL = "zl_医疗卡记录_DELETE('" & strNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(chk病历费.value, 1, 0) & ")"
    End If
    zlAddArray cllPro, strSQL
    Call AddPrintSQL(2, strNO, dtCurdate, cllPro)
    gcnOracle.BeginTrans: blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
    If CallBackBalanceInterface(strNO, blnOraclTrans) = False Then
        If blnOraclTrans = False Then gcnOracle.RollbackTrans
        Exit Function
    End If
    If blnOraclTrans = False Then gcnOracle.CommitTrans
    blnTrans = False
    SaveDelete = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CallBackBalanceInterface(ByVal strNO As String, ByRef blnTrancs As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用回退接口
    '入参:blnTrancs-是否处理了事务
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String, strSwapGlideNO As String, strSwapMemo As String, str结算信息 As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, lng结帐ID As Long, cllPro As Collection
    Dim bln消费卡 As Boolean, lng卡类别ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim str交易信息 As String, strTemp As String, dblMoney As Double
    
    On Error GoTo errHandle
    blnTrancs = False
    '如果选择了其他方式退费，不再调用接口
    If cbo支付方式.ItemData(cbo支付方式.ListIndex) < 7 Then CallBackBalanceInterface = True: Exit Function
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    'mcolBillBalance.Add Array(Val(Nvl(rsTmp!卡类别ID)), Trim(Nvl(rsTmp!卡号)), IIf(Val(Nvl(rsTmp!结算卡序号)) <> 0, 1, 0), Trim(Nvl(rsTmp!交易流水号)), Trim(Nvl(rsTmp!交易说明))), strNO
    If mcolBillBalance Is Nothing Then CallBackBalanceInterface = True: Exit Function
    '92895:李南春,2016/1/21,消费卡标志在第3位
    bln消费卡 = Val(mcolBillBalance(1)(2)) = 1
    lng卡类别ID = mcolBillBalance(1)(0)
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    If lng卡类别ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    str卡号 = mcolBillBalance(1)(1)
    strSwapGlideNO = mcolBillBalance(1)(3)
    strSwapMemo = mcolBillBalance(1)(4)
    str结算信息 = "5|" & mcolBillBalance(1)(6)
    strSQL = "Select 结帐ID,记帐费用 From 住院费用记录  Where 记录性质=5 and NO=[1] and 记录状态=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then
        gcnOracle.RollbackTrans: blnTrancs = True
        MsgBox "未找到退卡信息，不能继续", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng结帐ID = Val(Nvl(rsTemp!结帐ID))
    '81489,冉俊明,2015-1-22,退费传入冲销ID
    strSwapExtendInfor = "5|" & lng结帐ID: strTemp = strSwapExtendInfor
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(收款时的交易流水号)
    '       strSwapMemo-交易说明(收款时的交易说明)
    '       strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    dblMoney = Val(txt卡费.Text) + IIf(chk病历费.value = Checked, Val(txt病历费.Text), 0)
    If mobjDelObject.CardObject.zlReturnMoney(Me, mlngModule, lng卡类别ID, str卡号, str结算信息, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans: blnTrancs = True
        Exit Function
    End If
    
    '更新交易信息
    '    Zl_三方接口更新_Update
    strSQL = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  消费卡_In     Number,
    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSQL = strSQL & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSQL = strSQL & "'" & lng结帐ID & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSQL = strSQL & "'" & strSwapGlideNO & "',"
    '  交易说明_In   病人预交记录.交易说明%Type
    strSQL = strSQL & "'" & strSwapMemo & "',"
    '  预交款缴款_In Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  退费标志_In   Number := 0,
    strSQL = strSQL & "" & 1 & ")"
    '  校对标志_In   Number := Null,
    '  发送标志_In   Number := 0,
    '  消费卡管理_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    '先提交,这样避免风险,再更新相关的交易信息
    
    If strTemp <> strSwapExtendInfor Then
        'strSwapExtendInfor:交易扩展信息,格式:项目名称|项目内容||...
        varData = Split(strSwapExtendInfor, "||")
        Set cllPro = New Collection
        For i = 0 To UBound(varData)
            If Trim(varData(i)) <> "" Then
                varTemp = Split(varData(i) & "|", "|")
                If varTemp(0) <> "" Then
                    strTemp = varTemp(0) & "|" & varTemp(1)
                    If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                        str交易信息 = Mid(str交易信息, 3)
                        'Zl_三方结算交易_Insert
                        strSQL = "Zl_三方结算交易_Insert("
                        '卡类别id_In 病人预交记录.卡类别id%Type,
                        strSQL = strSQL & "" & lng卡类别ID & ","
                        '消费卡_In   Number,
                        strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                        '卡号_In     病人预交记录.卡号%Type,
                        strSQL = strSQL & "'" & str卡号 & "',"
                        '结帐ids_In  Varchar2,
                        strSQL = strSQL & "'" & lng结帐ID & "',"
                        '交易信息_In Varchar2:交易项目|交易内容||...
                        strSQL = strSQL & "'" & str交易信息 & "')"
                        zlAddArray cllPro, strSQL
                        str交易信息 = ""
                    End If
                    str交易信息 = str交易信息 & "||" & strTemp
                End If
            End If
        Next
        If str交易信息 <> "" Then
            str交易信息 = Mid(str交易信息, 3)
            'Zl_三方结算交易_Insert
            strSQL = "Zl_三方结算交易_Insert("
            '卡类别id_In 病人预交记录.卡类别id%Type,
            strSQL = strSQL & "" & lng卡类别ID & ","
            '消费卡_In   Number,
            strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
            '卡号_In     病人预交记录.卡号%Type,
            strSQL = strSQL & "'" & str卡号 & "',"
            '结帐ids_In  Varchar2,
            strSQL = strSQL & "'" & lng结帐ID & "',"
            '交易信息_In Varchar2:交易项目|交易内容||...
            strSQL = strSQL & "'" & str交易信息 & "')"
            zlAddArray cllPro, strSQL
        End If
        Err = 0: On Error GoTo ErrOthers:
        zlExecuteProcedureArrAy cllPro, Me.Caption
    End If
    CallBackBalanceInterface = True: blnTrancs = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans: blnTrancs = True
    Call ErrCenter
    Exit Function
ErrOthers:
    '扩展信息,允许保存一部分,以便查证
    If ErrCenter() = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    CallBackBalanceInterface = True
    gcnOracle.CommitTrans: blnTrancs = True
End Function
Private Function IsCheckChangeCardValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查换卡的数据是否合法
    '返回:数据合法,返回True,否则返回False
    '编制:刘兴洪
    '日期:2011-07-14 11:06:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If lbl刷卡验证.Tag = "" Then
        If Trim(txt刷卡卡号.Text) = "" Then
            MsgBox "原始卡号未进行刷卡确认,不能换卡!", vbInformation + vbOKOnly, gstrSysName
            If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
            Exit Function
        End If
        '-1-成功;0-失败;1-该记录不存在
        Select Case ReadCardNo(Trim(txt刷卡卡号.Text), 2)
        Case 0
            Exit Function
        Case 2
            Exit Function
        Case 1
            MsgBox "未找到原始卡号的持有人,请检查!", vbInformation + vbOKOnly, gstrSysName
            If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
            Exit Function
        End Select
    End If
    If mrsInfo Is Nothing Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If

     '问题号:50893
    If CStr(txt原卡密码.Tag) <> zlCommFun.zlStringEncode(Trim(txt原卡密码.Text)) Then
        MsgBox "原卡密码输入错误,请重新输入密码!", vbInformation + vbOKOnly, gstrSysName
        If txt原卡密码.Enabled And txt原卡密码.Visible Then txt原卡密码.SetFocus
        Exit Function
    End If
    IsCheckChangeCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function IsCheckFillCardValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查补卡的数据是否合法
    '返回:数据合法,返回True,否则返回False
    '编制:刘兴洪
    '日期:2011-07-14 11:06:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
     If mrsInfo Is Nothing Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If

    IsCheckFillCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveChangeCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存换卡
    '返回:换卡成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-14 11:50:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, Curdate As Date, lng结帐ID As Long
    Dim cllPro As Collection, cllUpdateSwap As Collection, cllThree As Collection
    On Error GoTo errHandle
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    Set cllPro = New Collection
    Set cllUpdateSwap = New Collection
    Set cllThree = New Collection
    Curdate = zlDatabase.Currentdate
    If AddCardDataSQL(lng病人ID, Curdate, cllPro, lng结帐ID) = False Then Exit Function
    If IDKindPayMode.IDKind = 2 And Val(txt余额.Text) > 0 Then Call AddDepositSQL(lng病人ID, Curdate, cllPro, lng结帐ID)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If zlInterfacePrayMoney(cllUpdateSwap, cllThree) = False Then
        gcnOracle.RollbackTrans
    End If
    zlExecuteProcedureArrAy cllUpdateSwap, Me.Caption, False, True
    On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllThree, Me.Caption
    SaveChangeCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    If ErrCenter = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    gcnOracle.CommitTrans
End Function
Private Function SaveFillCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存补卡信息
    '返回:补卡成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-14 11:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim lng病人ID As Long, Curdate As Date, lng结帐ID As Long
   Dim cllPro As Collection, cllUpdateSwap As Collection, cllThree As Collection
    On Error GoTo errHandle
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    
    Set cllPro = New Collection
    Set cllUpdateSwap = New Collection: Set cllThree = New Collection
    Curdate = zlDatabase.Currentdate
    If AddCardDataSQL(lng病人ID, Curdate, cllPro, lng结帐ID) = False Then Exit Function
    If IDKindPayMode.IDKind = 2 And Val(txt余额.Text) > 0 Then Call AddDepositSQL(lng病人ID, Curdate, cllPro, lng结帐ID)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If zlInterfacePrayMoney(cllUpdateSwap, cllThree) = False Then
        gcnOracle.RollbackTrans
    End If
    zlExecuteProcedureArrAy cllUpdateSwap, Me.Caption, False, True
    On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllThree, Me.Caption
    SaveFillCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    If ErrCenter = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    gcnOracle.CommitTrans
End Function
Private Function isCheckLossValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查挂失数据的合法性
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-14 13:40:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
   If mrsInfo Is Nothing Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "病人信息未找到,请先确定病人信息!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    If lbl刷卡验证.Tag = "" Then
        If Trim(txt刷卡卡号.Text) = "" Then
            MsgBox "挂失的卡号未进行刷卡确认,不能挂失!", vbInformation + vbOKOnly, gstrSysName
            If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
            Exit Function
        End If
        
        '-1-成功;0-失败;1-该记录不存在;2-无操作权限
        Select Case ReadCardNo(Trim(txt刷卡卡号.Text), 2)
        Case 0
            Exit Function
        Case 2
            Exit Function
        Case 1
            MsgBox "未找到当前卡号的持有人,请检查!", vbInformation + vbOKOnly, gstrSysName
            If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
            Exit Function
        End Select
    End If
    isCheckLossValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function SaveLossCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存挂失信息
    '返回:补卡成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-14 11:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim lng病人ID As Long, Curdate As Date, lng结帐ID As Long, cllPro As Collection
   Dim strSQL As String
   On Error GoTo errHandle
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
      'Zl_医疗卡变动_Insert
       strSQL = "Zl_医疗卡变动_Insert("
      '      变动类型_In   Number,
      '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
      strSQL = strSQL & "" & 6 & ","
      '      病人id_In     住院费用记录.病人id%Type,
      strSQL = strSQL & "" & lng病人ID & ","
      '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
      strSQL = strSQL & "" & mlngCardTypeID & ","
      '      原卡号_In     病人医疗卡信息.卡号%Type,
      strSQL = strSQL & "'" & lbl刷卡验证.Tag & "',"
      '      医疗卡号_In   病人医疗卡信息.卡号%Type,
      strSQL = strSQL & "'" & lbl刷卡验证.Tag & "',"
      '      变动原因_In   病人医疗卡变动.变动原因%Type,
      '      --变动原因_In:如果密码调整，变动原因为密码.加密的
      strSQL = strSQL & "'" & Trim(txt变动原因.Text) & "',"
      '      密码_In       病人信息.卡验证码%Type,
      strSQL = strSQL & "NULL,"
      '      操作员姓名_In 住院费用记录.操作员姓名%Type,
      strSQL = strSQL & "'" & UserInfo.姓名 & "',"
      '      变动时间_In   住院费用记录.登记时间%Type,
      strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
      '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
      strSQL = strSQL & "NULL,"
      '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
      strSQL = strSQL & "'" & cbo挂失方式.Text & "')"
     Call zlAddArray(cllPro, strSQL)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveLossCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdCreateCard_Click()
    '问题号:56599
    Dim strExpend As String
    Dim blnFlag As Boolean
    Dim strOutPatiInforXml As String
    On Error GoTo errHandle
    
    If mrsInfo Is Nothing Then
        MsgBox "病人信息不存在或是未在本院建档,不能进行制卡操作！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Sub
    End If
    If mobjReadCard Is Nothing Then Exit Sub
    On Error Resume Next
    If mobjReadCard.zlMakeCard(Me, mlngModule, mlngCardTypeID, Get制卡XML(mrsInfo!病人ID), mstr采集图片, strOutPatiInforXml, strExpend) = False Then
        If Err = 438 Then
            MsgBox mCardType.str卡名称 & "没有编写制卡接口(zlMakeCard),制卡失败!", vbInformation, gstrSysName
            Err.Clear
        ElseIf Err <> 0 Then
            GoTo errHandle
        End If
        Exit Sub
    End If
    On Error GoTo errHandle
    If strOutPatiInforXml <> "" Then
        Call Clear健康档案
        LoadPati strOutPatiInforXml
    End If
    Exit Sub
    
errHandle:
   If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

 
Private Sub cmdOK_Click()
    Dim blnPrint As Boolean, blnPlugInCheck As Boolean, lng病人ID As Long
    Dim objfrmPrint As frmPrint
    
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    Set objfrmPrint = New frmPrint
    Load objfrmPrint
    Call txt余额_Change
    If IsCheck医疗卡 = False Then Exit Sub
    If CheckDepositFactValied = False Then Exit Sub
    If Check发卡性质(lng病人ID, mCardType.lng卡类别ID) = False Then Exit Sub
    If mEditType = Cr_退卡 Or chkCancel.value = 1 Then
       If IsCheckCancelValied = False Then Exit Sub
       If IsCheckCancel退预交 = False Then Exit Sub '问题号:51537
       If CheckCardDelValied() = False Then Exit Sub
       If SaveDelete(cboNO.Tag) = False Then Exit Sub
       
        mintSucces = 1
        If mEditType = Cr_退卡 Then
            mblnChange = False
            Unload Me: Exit Sub
        End If
        chkCancel.value = 0
        If Me.txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call ClearData
        mblnChange = False
        Exit Sub
    End If
    If mEditType = Cr_补卡 Then
        If IsCheckFillCardValied = False Then mEditType = mEditTypeOld: Exit Sub
        If CheckChargeFactValied = False Then mEditType = mEditTypeOld: Exit Sub
        '刷卡处理
        If CheckBrushCard = False Then Exit Sub
        If SaveFillCard = False Then Exit Sub
        Call objfrmPrint.PrintBill(mCurPayMoney.strNO, Trim(txt卡号.Text), Trim(txtFact.Text), _
                            mlngCardTypeID, mPrint.blnPrint, mEditType, mPrint.bytPrintFormat, _
                            mPrint.lng领用ID, mPrint.strUseType, mPrint.dtPrintdate, UserInfo.姓名, mblnPrepayPrint, _
                            mstrPrePayNo, mlng预交病人ID, mdat预交时间)
                            
        mintSucces = 1
        Call ClearData(True)
        mEditType = mEditTypeOld
        Unload Me: Exit Sub
    End If
    If mEditType = Cr_挂失 Then
        If isCheckLossValied = False Then Exit Sub
        If SaveLossCard = False Then Exit Sub
        Call ClearData
        mintSucces = 1: Unload Me: Exit Sub
    End If
    If Not isValied Then Exit Sub
    
    If mEditType = Cr_调整病人信息 Then
        If mrsInfo Is Nothing Then Exit Sub
        If mrsInfo.State <> 1 Then Exit Sub
        If LinkManValid = False Then Exit Sub
        '如果更改病人的基本信息,发生业务的,不能进行调整
        mbln基本信息调整 = False
        
        If IsCertificateCard(lng病人ID) = False Then Exit Sub
        
        If Nvl(mrsInfo!姓名) <> txtPatient.Text _
            Or Nvl(mrsInfo!性别) <> zlstr.NeedName(cbo性别.Text) _
            Or Nvl(mrsInfo!年龄) <> txt年龄.Text & cbo年龄单位 _
            Or Format(mrsInfo!出生日期, "yyyy-mm-dd") <> txt出生日期.Text Then
            If InStr(mstrPrivsPubPatient, ";基本信息调整;") = 0 Then
                MsgBox "该病人已经发生了医嘱业务数据,不能进行病人的基本信息调整(姓名,性别,年龄,出生日期等),请在『病人信息管理』中进行调整!", vbOKOnly + vbInformation, gstrSysName
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                zlControl.TxtSelAll txtPatient
                Exit Sub
            Else
                mbln基本信息调整 = True
            End If
        End If
        If SaveModifyPati = False Then Exit Sub
        mintSucces = 1
        Call ClearData
        Unload Me: Exit Sub
    End If
    
    If mEditType = Cr_换卡 Then
        If IsCheckChangeCardValied = False Then Exit Sub
        If CheckChargeFactValied = False Then Exit Sub
        If CheckBrushCard = False Then Exit Sub
        If SaveChangeCard = False Then Exit Sub
        Call objfrmPrint.PrintBill(lblNo.Tag, Trim(txt卡号.Text), Trim(txtFact.Text), _
                                mlngCardTypeID, mPrint.blnPrint, mEditType, mPrint.bytPrintFormat, _
                                mPrint.lng领用ID, mPrint.strUseType, mPrint.dtPrintdate, UserInfo.姓名)
        mintSucces = 1
        Call ClearData(True)
        Unload Me: Exit Sub
    End If
           
    If mEditType = Cr_绑定卡 Or mEditType = Cr_发卡 Then
    
        If CheckChargeFactValied = False Then Exit Sub
        
        '刷卡处理
        If CheckBrushCard = False Then Exit Sub
        '问题号:51072
        If Len(Trim(txtPass.Text)) = 0 Then '没有输入卡密码的情况
           If zl_Get设置默认发卡密码 = False Then Exit Sub
        End If
        
        '问题号56599
        If InoculateValid = False Then Exit Sub
        If LinkManValid = False Then Exit Sub

        '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
        If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '保存插件附加信息前的数据有效性检查
            On Error Resume Next
            blnPlugInCheck = mobjPlugIn.PatiInfoSaveBefore(mlng病人ID)
            Call zlPlugInErrH(Err, "PatiInfoSaveBefore")
            If Err = 0 And blnPlugInCheck = False Then
                Exit Sub '检查未通过终止保存
            End If
            Err.Clear
        End If

        If SaveData = False Then Exit Sub
        
        Call objfrmPrint.PrintBill(mCurPayMoney.strNO, Trim(txt卡号.Text), Trim(txtFact.Text), _
                            mlngCardTypeID, mPrint.blnPrint, mEditType, mPrint.bytPrintFormat, _
                            mPrint.lng领用ID, mPrint.strUseType, mPrint.dtPrintdate, UserInfo.姓名, mblnPrepayPrint, _
                            mstrPrePayNo, mlng预交病人ID, mdat预交时间)
        
        mintSucces = mintSucces + 1
        Call LoadSaveNotoCombox: Call ClearData(True)
        Call CheckBILL("")
        If txtPatient.Enabled And txtPass.Visible Then txtPatient.SetFocus
        mintSucces = mintSucces + 1
        Exit Sub
    End If
    mintSucces = mintSucces + 1
    Call ClearData
    Unload Me
End Sub

Private Function zl_Get设置默认发卡密码() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置默认发卡密码
    '返回:是否继续发卡操作
    '编制:王吉
    '日期:2012-07-06 15:53:14
    '问题号:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardType As clsCard
    Dim msgResult As VbMsgBoxResult
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    '59760
    If zlGetCards_YL(objYLCards) = False Then Exit Function
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    Set objCardType = objYLCards.Item("K" & mlngCardTypeID)
    If objCardType.是否缺省密码 = False Then '无限制
        Select Case objCardType.密码输入限制
            Case 0 '无限制
                zl_Get设置默认发卡密码 = True
                Exit Function
            Case 1 '未输入提醒
               msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续？", vbQuestion + vbYesNo, gstrSysName)
               zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
               Exit Function
            Case 2 '为输入禁止
                 MsgBox "未输入卡密码,不能进行发卡！", vbExclamation, gstrSysName
                zl_Get设置默认发卡密码 = False
                Exit Function
        End Select
    ElseIf objCardType.是否缺省密码 Then '缺省身份证后N位
        If Len(Trim(txt身份证号.Text)) > 0 Or Len(Trim(txt联系人身份证号.Text)) > 0 Then '输入了身份证或联系人身份证号
            If Len(Trim(txt身份证号.Text)) > 0 Then '有身份证优先用身份证
                   txtPass.Text = Right(Trim(txt身份证号.Text), objCardType.密码长度)
            Else '否则就用代办人身份证作为密码
                   txtPass.Text = Right(Trim(txt联系人身份证号.Text), objCardType.密码长度)
            End If
            txtAudi.Text = txtPass.Text
        Else '身份证与联系人身份证都没输入
            Select Case objCardType.密码输入限制
                Case 0 '无限制
                    zl_Get设置默认发卡密码 = True
                    Exit Function
                Case 1 '未输入提醒
                    msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续？", vbQuestion + vbYesNo, gstrSysName)
                    zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
                    Exit Function
                Case 2 '为输入禁止
                    MsgBox "未输入卡密码,不能进行发卡！", vbExclamation, gstrSysName
                    zl_Get设置默认发卡密码 = False
                    Exit Function
            End Select
        End If
    End If
    zl_Get设置默认发卡密码 = True
End Function

Private Function CheckBILL(Optional strCardNo As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查票据是否存在在
    '编制:刘兴洪
    '日期:2011-07-12 15:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '106010:李南春，2017/3/10，非严格控制发卡清空领用ID
    Dim strSQL As String
    Dim rsTemp As Recordset
    If Not mCardType.bln严格控制 Then mCardType.lng领用ID = 0: CheckBILL = True: Exit Function
    If mCardType.bln是否重复使用 Then
        mCardType.lng领用ID = 0
        strSQL = "Select b.领用Id" & vbNewLine & _
             "From 票据领用记录 A, 票据使用明细 B" & vbNewLine & _
             "Where a.Id = b.领用id And a.票种 = 5 And (Nvl(a.使用类别, 'LXH') = [1] Or a.使用类别 Is Null) And b.号码 = [2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, strCardNo)
        If rsTemp.RecordCount > 0 Then
            mCardType.lng领用ID = Nvl(rsTemp!领用Id, 0)
        Else
            mCardType.lng领用ID = CheckUsedBill(5, IIf(mCardType.lng领用ID > 0, mCardType.lng领用ID, mCardType.lng共用批次), strCardNo, mlngCardTypeID)
        End If
    Else
        mCardType.lng领用ID = CheckUsedBill(5, IIf(mCardType.lng领用ID > 0, mCardType.lng领用ID, mCardType.lng共用批次), strCardNo, mlngCardTypeID)
    End If
    If mCardType.lng领用ID <= 0 Then
        Select Case mCardType.lng领用ID
            Case 0 '操作失败
            Case -1
                If txt卡号.Text <> "" Then MsgBox "你已没有自用及共用的" & mCardType.str卡名称 & ",不能发放！" & vbCrLf & _
                    "请先在本地设置共用批次或领用一批新卡! ", vbExclamation, gstrSysName
                Exit Function
            Case -2
                If txt卡号.Text <> "" Then MsgBox "本地共用的" & mCardType.str卡名称 & "已用完,不能发放！" & vbCrLf & _
                    "请重新设置本地共用卡批次或领用一批新卡！", vbExclamation, gstrSysName
                Exit Function
            Case -3
                MsgBox "该张卡号不在有效范围内,请检查是否正确刷卡！", vbExclamation, gstrSysName
                If txt卡号.Enabled And txt卡号.Enabled Then txt卡号.SetFocus
                Exit Function
        End Select
    End If
    CheckBILL = True
End Function

Private Sub cmdPicClear_Click()
    '问题号:56599
    imgPatient.Picture = Nothing
    mlng图像操作 = 3
End Sub

Private Sub cmdPicCollect_Click()
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.PatiImageGatherer(Me, mstr采集图片) = False Then Exit Sub
    imgPatient.Picture = LoadPicture(mstr采集图片)
    mlng图像操作 = 2
End Sub

Private Sub cmdPicFile_Click()
    '问题号:56599
    Dim strFileDir As String
On Error GoTo ErrHanl:
    With cmdialog
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .Filter = "(*.bmp)|*.bmp"
        .FilterIndex = 2
        .ShowOpen
        strFileDir = .FileName
        If strFileDir = "" Then Exit Sub
        imgPatient.Picture = LoadPicture(strFileDir)
    End With
    mlng图像操作 = 1
    Exit Sub
ErrHanl:
    
End Sub

Private Sub cmdReadCard_Click()
    Call ReReadCard("")
End Sub

Private Function LoadPati(ByVal strPatiXML As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息,读取病人信息
    '编制:刘兴洪
    '日期:2011-09-08 21:52:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '问题号:56599
    Dim str过敏药物 As String, str过敏反应 As String '问题号:56599
    Dim str接种日期 As String, str接种名称 As String '问题号:56599
    Dim strABO血型 As String '问题号:56599
    Dim str信息名 As String, str信息值 As String '问题号:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '问题号:56599
    Dim str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, str地址 As String '问题号:56599
    Dim str其他关系 As String, strBirth As String
    On Error GoTo errHandle
    If Not (mEditType = Cr_绑定卡 Or mEditType = Cr_发卡) Then Exit Function
    If strPatiXML = "" Then Exit Function
    
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strPatiXML, False) = False Then Exit Function
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    Call zlXML_GetNodeValue("卡号", , strValue)
    '    姓名    Varchar2    100
    Call zlXML_GetNodeValue("姓名", , strValue)
    txtPatient.Text = strValue
    '    性别    Varchar2    4
    Call zlXML_GetNodeValue("性别", , strValue)
    If strValue <> "" Then
        Call zlControl.CboLocate(cbo性别, strValue)
        If cbo性别.ListIndex = -1 Then
            cbo性别.AddItem strValue
            cbo性别.ListIndex = cbo性别.NewIndex
        End If
    End If
    '    年龄    Varchar2    10
    Call zlXML_GetNodeValue("年龄", , strValue)
    If strValue <> "" Then
        Call LoadOldData(strValue, txt年龄, cbo年龄单位)
    End If
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("出生日期", , strValue)
    
    txt出生日期.Text = Format(IIf(strValue = "", "____-__-__", strValue), "YYYY-MM-DD")
    mblnNotChange = True
    If strValue <> "" Then
         txt年龄.Text = ReCalcOld(CDate(txt出生日期.Text), cbo年龄单位)      '修改的时候,根据出生日期重算年龄
         If CDate(txt出生日期.Text) - CDate(strValue) <> 0 Then txt出生时间.Text = Format(strValue, "HH:MM")
     Else
         '103807:李南春，2016/12/20，年龄反算精确到小时
        If Not mobjPubPatient Is Nothing Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), strBirth) Then
                txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                txt出生时间.Text = Format(strBirth, "hh:mm")
            End If
        End If
     End If
     mblnNotChange = False
    '    出生地点    Varchar2    50
    Call zlXML_GetNodeValue("出生地点", , strValue)
    '    身份证号    VARCHAR2    18
    Call zlXML_GetNodeValue("身份证号", , strValue)
    If strValue <> "" Then txt身份证号.Text = strValue
    '    其他证件    Varchar2    20
    Call zlXML_GetNodeValue("其他证件", , strValue)
    If strValue <> "" Then txt其他证件.Text = strValue
    '    职业    Varchar2    80
    Call zlXML_GetNodeValue("职业", , strValue)
    If strValue <> "" Then
        Call cbo.SeekIndex(cbo职业, strValue)
        If cbo职业.ListIndex = -1 Then
            cbo职业.AddItem strValue, 0
            cbo职业.ListIndex = cbo职业.NewIndex
        End If
    End If
    '    民族    Varchar2    20
    Call zlXML_GetNodeValue("民族", , strValue)
    Call cbo.SeekIndex(cbo民族, strValue, , True)
     If cbo民族.ListIndex = -1 And strValue <> "" Then
         cbo民族.AddItem strValue, 0
         cbo民族.ListIndex = cbo民族.NewIndex
     End If
    '    国籍    Varchar2    30
    Call zlXML_GetNodeValue("国籍", , strValue)
    Call cbo.SeekIndex(cbo国籍, strValue, , True)
     If cbo国籍.ListIndex = -1 And strValue <> "" Then
         cbo国籍.AddItem strValue, 0
         cbo国籍.ListIndex = cbo国籍.NewIndex
     End If
    '    学历    Varchar2    10
    Call zlXML_GetNodeValue("学历", , strValue)
    Call cbo.SeekIndex(cbo学历, strValue, , True)
    If cbo学历.ListIndex = -1 And strValue <> "" Then
        cbo学历.AddItem strValue, 0
        cbo学历.ListIndex = cbo学历.NewIndex
    End If
    '    婚姻状况    Varchar2    4
    Call zlXML_GetNodeValue("婚姻状况", , strValue)
    Call cbo.SeekIndex(cbo婚姻状况, strValue, , True)
     If cbo婚姻状况.ListIndex = -1 And strValue <> "" Then
         cbo婚姻状况.AddItem strValue, 0
         cbo婚姻状况.ListIndex = cbo婚姻状况.NewIndex
     End If
    '    区域    Varchar2    30
    Call zlXML_GetNodeValue("区域", , strValue)
    txt区域.Text = strValue
    '    家庭地址    Varchar2    50
    Call zlXML_GetNodeValue("家庭地址", , strValue)
   txt家庭地址.Text = strValue
   padd家庭地址.value = strValue
    '    户口地址    Varchar2    50
    Call zlXML_GetNodeValue("户口地址", , strValue)
    txt户口地址.Text = strValue
    padd户口地址.value = strValue
    '    家庭电话    Varchar2    20
    Call zlXML_GetNodeValue("家庭电话", , strValue)
   txt家庭电话.Text = strValue
    '    家庭地址邮编    Varchar2    6
    Call zlXML_GetNodeValue("家庭地址邮编", , strValue)
   txt家庭邮编.Text = strValue
   '    手机号    Varchar2    20
    Call zlXML_GetNodeValue("手机号", , strValue)
   txt手机.Text = strValue
    '    监护人  Varchar2    64
    Call zlXML_GetNodeValue("监护人", , strValue)
   'txt监护人.Text = strValue
'    '    联系人姓名  Varchar2    64
'    Call zlXML_GetNodeValue("联系人姓名", , strValue)
'    '    联系人关系  Varchar2    30
'    Call zlXML_GetNodeValue("联系人关系", , strValue)
'    '    联系人地址  Varchar2    50
'    Call zlXML_GetNodeValue("联系人地址", , strValue)
'    txt联系人姓名.Text = strValue
'    '    联系人电话  Varchar2    20
'    Call zlXML_GetNodeValue("联系人电话", , strValue)
'    txt联系人电话.Text = strValue
    '    工作单位    Varchar2    100
    Call zlXML_GetNodeValue("工作单位", , strValue)
    txt工作单位.Text = strValue
    lbl工作单位.Tag = ""
    '    单位电话    Varchar2    20
    Call zlXML_GetNodeValue("单位电话", , strValue)
   txt单位电话.Text = strValue
    '    单位邮编    Varchar2    6
    Call zlXML_GetNodeValue("单位邮编", , strValue)
   txt单位邮编.Text = strValue
    '    单位开户行  Varchar2    50
    Call zlXML_GetNodeValue("单位开户行", , strValue)
   txt单位开户行.Text = strValue
    '    单位帐号    Varchar2    20
    Call zlXML_GetNodeValue("单位帐号", , strValue)
   txt单位帐户.Text = strValue
    '问题号:56599
    '过敏情况
    Call zlXML_GetRows("药物名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("药物名称", i, str过敏药物)
        Call zlXML_GetNodeValue("药物反应", i, str过敏反应)
        SetDrugAllergy str过敏药物, str过敏反应
    Next
    lngCount = 0
    '免疫记录
    Call zlXML_GetRows("疫苗名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("疫苗名称", i, str接种名称)
        Call zlXML_GetNodeValue("接种时间", i, str接种日期)
        SetInoculate str接种日期, str接种名称
    Next
    lngCount = 0
    'ABO血型
    Call zlXML_GetNodeValue("ABO血型", , strABO血型)
    If strABO血型 <> "" Then
        For i = 0 To cboBloodType.ListCount - 1
            '76314,李南春，2014-08-06，病人信息正确获取
            If zlstr.NeedName(cboBloodType.List(i), ".") = zlstr.NeedName(strABO血型) Then cboBloodType.ListIndex = i
        Next
    End If
    'RH
    Call zlXML_GetNodeValue("RH", , strValue)
    If strValue <> "" Then
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = strValue Then cboBH.ListIndex = i
        Next
    End If
    '医学警示
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("临床基本信息")
    If Not xmlChildNodes Is Nothing Then
        If xmlChildNodes.length > 0 Then
            For i = 0 To xmlChildNodes.length - 1
                Set xmlChildNode = xmlChildNodes(i)
                If xmlChildNode.Text = "1" Then
                    strValue = strValue & ";" & Replace(xmlChildNode.nodeName, "标志", "")
                End If
            Next
        End If
    End If
    If strValue <> "" Then txtMedicalWarning.Text = Mid(strValue, 2)
    '其他医学警示
    Call zlXML_GetNodeValue("其他医学警示", , strValue)
    If strValue <> "" Then txtOtherWaring.Text = strValue
    '联系信息
    '    联系人地址  Varchar2    50
    Call zlXML_GetNodeValue("联系人地址", , str地址)
    txt联系人地址.Text = str地址
     '    联系人姓名  Varchar2    64
    Call zlXML_GetNodeValue("联系人姓名", , str姓名)
    '    联系人关系  Varchar2    30
    Call zlXML_GetNodeValue("联系人关系", , str关系)
    '    联系人电话  Varchar2    20
    Call zlXML_GetNodeValue("联系人电话", , str电话)
    '    联系人身份证 Varchar2   20
    Call zlXML_GetNodeValue("联系人身份证号", , str身份证号)
    '84313,李南春,2015/4/27,联系人关系以及其他关系
     '    联系人其他关系 Varchar2   30
    Call zlXML_GetNodeValue("联系人附加信息", , str其他关系)
    
    SetLinkInfo str姓名, str关系, str电话, str身份证号, str其他关系
    
    Call zlXML_GetRows("联系信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("联系信息", "姓名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("联系信息", "姓名", i, j, str姓名)
                Call zlXML_GetChildNodeValue("联系信息", "关系", i, j, str关系)
                Call zlXML_GetChildNodeValue("联系信息", "电话", i, j, str电话)
                Call zlXML_GetChildNodeValue("联系信息", "身份证号", i, j, str身份证号)
                Call zlXML_GetChildNodeValue("联系信息", "附加信息", i, j, str其他关系)
                SetLinkInfo str姓名, str关系, str电话, str身份证号, str其他关系
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '其他信息
    '健康档案编号
    Call zlXML_GetNodeValue("健康档案编号", , strValue)
    SetOtherInfo "健康档案编号", strValue
    
    '新农合证号
    Call zlXML_GetNodeValue("新农合证号", , strValue)
    SetOtherInfo "新农合证号", strValue

    '其他证件
    Call zlXML_GetRows("其他证件", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他证件", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他证件", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他证件", "信息值", i, j, str信息值)
                SetOtherInfo str信息名, str信息值
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '其他信息
    Call zlXML_GetRows("其他信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他信息", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他信息", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他信息", "信息值", i, j, str信息值)
                SetOtherInfo str信息名, str信息值
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '医疗卡属性
    Call zlXML_GetRows("医疗卡属性", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("医疗卡属性", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("医疗卡属性", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("医疗卡属性", "信息值", i, j, str信息值)
                If mdic医疗卡属性.Exists(str信息名) Then
                    mdic医疗卡属性.Item(str信息名) = str信息值
                Else
                    mdic医疗卡属性.Add str信息名, str信息值
                End If
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    
    LoadPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdSelDrug_Click()
    '问题号:56599
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select ID,nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
        " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试" & _
        " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码," & _
        " A.名称,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"

    '获取当前鼠标坐标值
    vRect = zlControl.GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (vsDrug.Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "过敏药物", False, "过敏药物选择器", "请从下面的药品中选择一项作为病人过敏药物", False, False, True, vRect.Left, vRect.Top, 0, True, False, True)

    If Not rsTemp Is Nothing Then
        vsDrug.TextMatrix(vsDrug.Row, vsDrug.Col) = rsTemp!名称
        vsDrug.TextMatrix(vsDrug.Row, 2) = rsTemp!id
        If vsDrug.Rows - 1 = vsDrug.Row Then vsDrug.Rows = vsDrug.Rows + 1
    End If
    If vsDrug.Visible = True And vsDrug.Enabled = True Then vsDrug.SetFocus
    Exit Sub
ErrHandl:
    MsgBox Err.Description
End Sub

Private Sub cmd充值_Click()
    '问题号:54208
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State <> 0 Then
            Call zlPrepayFunc(1, mrsInfo!病人ID)
        End If
    Else
        Call zlPrepayFunc(1, 0)
    End If
End Sub

Private Sub cmd户口地址_Click()
    Call SearchAddress("", txt户口地址)
End Sub

Private Sub cmd余额退款_Click()
    '问题号:50891
    Call zlPrepayFunc(2, mlng病人ID)
End Sub
Private Function zlPrepayFunc(ByVal intFunc As Integer, ByVal lng病人ID As Long, Optional blnOneCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:缴预存款
    '入参:intFunc-1-缴预存;2-退预款;3-作废,4-门诊转住院;5-住院转门诊;
    '编制:刘兴洪
    '日期:2011-07-24 18:25:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun As Object, int预交类型 As Integer
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Function
    'byt预交类型: 0-收预交款(缺省,可切换到退),1-浏览单据(1),2-作废状态(1); 3-余额退款(37770), 4-门诊转住院;5-住院转门诊
    Select Case intFunc
    Case 1  '1.缴预存
        int预交类型 = 0
    Case 2 '退款
        int预交类型 = 3
    Case 3: int预交类型 = 2
    Case 4: int预交类型 = 4
    Case 5: int预交类型 = 5
    End Select
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能： 调用预交款收款窗口
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前数据库登录用户名
    '  bytCallObject:刘兴洪加入(0-预交款调用(缺省的);1-病人费用查询调用,2-医疗卡调用)
    '  lng病人ID-缺省的病人ID
    '  lng主页ID-缺省的主页ID
    '  dblDefPrePayMoney-缺省的预付金额
    Set gfrmCardMgr = Me
    '问题:48249
    If objFun.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 2, lng病人ID, 0, 0, int预交类型, blnOneCard) = False Then
        zlPrepayFunc = False
        Set gfrmCardMgr = Nothing
        Exit Function
    End If
    Set gfrmCardMgr = Nothing
    zlPrepayFunc = True
End Function
Private Sub cmd出生地点_Click()
    If Select地区(txt出生地点, lbl出生地点, "") = False Then Exit Sub
End Sub
Private Sub cmd合同单位_Click()
    If Select合约单位("") = False Then Exit Sub
End Sub

Private Sub cmd家庭地址_Click()
    If Select地区(txt家庭地址, lbl家庭地址, "") = False Then Exit Sub
End Sub

Private Sub cmd联系人地址_Click()
    If Select地区(txt联系人地址, lbl联系人地址, "") = False Then Exit Sub
End Sub

Private Sub cmd区域_Click()
    If Select区域("") = False Then Exit Sub
End Sub

Private Sub Set权限()
    Dim strValue As String
    '票号严格控制
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    mblnBill预交 = Mid(strValue, 2, 1) = "1"
    
    '票据号码长度、就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    mbyt预交 = Val(Split(strValue, "|")(1))
    
    cmd余额退款.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "预交退款")
    cmd充值.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "预交退款")
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnLoad Then Unload Me: Exit Sub
    
    If gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    End If
    
    Call LoadCardFee: Call SetCtrlMove
    Call SetControlEnable
    Call SetCardEditEnabled
    '修改人:56599
    Call InitTabPage
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    Call InitTaskPanelOther
    If mstrCardNo <> "" Then
        If mEditType = Cr_查询 Then
            mint记录状态 = 1
            Call ReadCardNo(mstrCardNo, 2)
        Else
            Call ReReadCard(mstrCardNo)
        End If
    End If
    
    If mlng病人ID <> 0 Then
        If GetPatient("-" & mlng病人ID) Then
            Call LoadPatiInfor: Call zlQueryEMPIPatiInfo
            Call ReInitPatiInvoice
        End If
        If mEditType = Cr_挂失 Then
            txt刷卡卡号.Text = mstrCardNo
            If txt刷卡卡号.Text = "" Then
                If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
            Else
                If cbo挂失方式.Enabled And cbo挂失方式.Visible Then cbo挂失方式.SetFocus
            End If
        End If
    Else
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    If mEditType = Cr_换卡 Then
         If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
         txt刷卡卡号.Text = mstrCardNo
    End If
    If mEditType = Cr_退卡 Then
        '问题:47772
         chkCancel.value = 1
        '问题:48249
         mblnNotClick = True
         '0-不进行刷卡;1-刷卡退卡;2-单据号后再验证刷卡;3-1和2的共用模式
         Select Case mParaData.int退卡模式
         Case 0
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
         Case 1
             If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
         Case 2
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
         Case Else
             If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
         End Select
        mblnNotClick = False
    End If
    wndTaskPanel.Reposition
    mblnChange = False
    
       '问题号:56599
'    If mEditType <> Cr_绑定卡 And mEditType <> Cr_发卡 Then
'        NotVisibleImage
'    End If
End Sub
Private Sub BackCardReadCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退卡读卡
    '编制:刘兴洪
    '日期:2011-12-25 14:04:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strOutPut As String, strExpand As String, strOutXml As String, strCardNo As String, intFlag As Integer
    If Not (mEditType = Cr_退卡 Or chkCancel.value = 1) Then Exit Sub
    If mCardType.bln就诊卡 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt刷卡卡号.Text = mobjICCard.Read_Card()
            If Trim(txt刷卡卡号.Text) = "" Then Exit Sub
            intFlag = ReadCardNo(Trim(txt刷卡卡号.Text), 2)
            If intFlag = -1 Then
                If mEditType <> Cr_换卡 Then
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
                End If
            ElseIf intFlag = 2 Then
                Exit Sub
            Else
                Call zlControl.TxtSelAll(txt刷卡卡号)
                stbThis.Panels(2) = "没有发现该" & mCardType.str卡名称 & "的信息,可能未建档,请检查！"
                txt刷卡卡号.Text = ""
                Exit Sub
            End If
        End If
        Exit Sub
    End If
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Sub
    End If
    If mobjReadCard Is Nothing Then Exit Sub
    If mobjReadCard.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    txt刷卡卡号.Text = strCardNo
    If Trim(txt刷卡卡号.Text) = "" Then Exit Sub
    intFlag = ReadCardNo(Trim(txt刷卡卡号.Text), 2)
    If intFlag = -1 Then
        '成功
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
    ElseIf intFlag = 2 Then
        Exit Sub
    Else
        Call zlControl.TxtSelAll(txt刷卡卡号)
        stbThis.Panels(2) = "没有发现该" & mCardType.str卡名称 & "的信息,请检查！"
        txt刷卡卡号.Text = ""
        Exit Sub
    End If
End Sub

Private Function ReReadCard(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新读书
    '编制:刘兴洪
    '日期:2011-09-08 22:20:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPhotoFile As String
    Dim strOutPut As String, strExpand As String, strOutXml As String
    '问题:48249
    If (mEditType = Cr_退卡 Or chkCancel.value = 1) And strCardNo = "" Then
        '退卡读卡
        Call BackCardReadCard: Exit Function
    End If
    '问题号:57962
    If mEditType = Cr_换卡 Then
        txt刷卡卡号.Text = strCardNo '换卡时这里的Text对象代表的是原卡号输入框
    End If
    
    '问题:47914
    '问题:48079
    If Not (mEditType = Cr_发卡 Or mEditType = Cr_绑定卡 _
                                Or (mEditType = Cr_补卡 And (Mid(mCardType.str读卡性质, 3, 1) = 1 Or Mid(mCardType.str读卡性质, 4, 1) = 1)) _
                                Or (mEditType = Cr_换卡 And (Mid(mCardType.str读卡性质, 3, 1) = 1 Or Mid(mCardType.str读卡性质, 4, 1) = 1)) _
                                    ) Then Exit Function
   ' If mCardType.bln自制卡 Then Exit Sub
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Function
    End If
    
    If mobjReadCard Is Nothing Then Exit Function
    strExpand = mlngCardTypeID
    On Error Resume Next
    ReReadCard = mobjReadCard.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml, strPhotoFile)
    If Err <> 0 Then
        If Err <> 450 Then GoTo errHandle:
        '450-错误的参数号或无效的属性赋值
        '主要是歉容以前的
         Err = 0: On Error GoTo errHandle
         ReReadCard = mobjReadCard.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml)
    End If
    If Not ReReadCard Then Exit Function
    
    txt卡号.Text = Trim(strCardNo)
    If txt卡号.Text <> "" Then
        Call CheckFreeCard(txt卡号.Text)
        '问题:62821
        If strPhotoFile <> "" Then imgPatient.Picture = LoadPicture(strPhotoFile)
        Call Clear健康档案
        Call LoadPati(strOutXml)
    End If
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is txt卡号 Or Me.ActiveControl Is txtAudi Or Me.ActiveControl Is txtPass Then Exit Sub
        If Me.ActiveControl Is txt刷卡卡号 Then Exit Sub
        If Me.ActiveControl Is txt联系人地址 Then Exit Sub
        If Me.ActiveControl Is txt区域 Then Exit Sub
        If Me.ActiveControl Is txt家庭地址 Then Exit Sub
        If Me.ActiveControl Is txt工作单位 Then Exit Sub
        If Me.ActiveControl Is txt出生地点 Then Exit Sub
        If Me.ActiveControl Is txt年龄 Then Exit Sub
        '76609,冉俊明,2014-8-14,焦点定位问题
        If Me.ActiveControl Is txtPatient Then Exit Sub
        '78408:李南春,2014/10/9,光标跳转
        If Me.ActiveControl Is vsDrug Then Exit Sub
        If Me.ActiveControl Is vsInoculate Then Exit Sub
        If Me.ActiveControl Is vsCertificate Then Exit Sub
        If Me.ActiveControl Is txt卡费 Then Exit Sub
        
        '89242:李南春,2015/12/10,PatiAddress控件内部处理了跳转，外部不再处理
        If UCase(TypeName(Me.ActiveControl)) = UCase("PatiAddress") Then Exit Sub
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    Select Case KeyCode
    Case vbKeyE
        If Shift = vbCtrlMask Then
            If wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded Then
                wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded = False
            Else
                wndTaskPanel.Groups(mTaskPancel_ID.Idx_TP_PatiExpend).Expanded = True
            End If
        End If
    Case vbKeyF2
        If cmdOK.Enabled And cmdOK.Visible Then
            cmdOK.SetFocus: Call cmdOK_Click
        End If
    Case vbKeyF6
        If txtPatient.Enabled And txtPatient.Visible Then
            txtPatient.SetFocus
        End If
    Case vbKeyF8
        If mEditType = Cr_发卡 Then
            chkCancel.value = 1
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        End If
    Case vbKeyF12
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    Case vbKeyEscape
        If cmdCancel.Enabled And cmdCancel.Visible Then
            cmdCancel.SetFocus: Call cmdCancel_Click
        End If
    Case Else
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim intKind As Integer, strKey As String
    mstrCboSplit = "-" & Chr(30)
    mblnFirst = True
    mstrPrepayPrivs = ";" & GetPrivFunc(glngSys, 1103) & ";"
    mstrTitle = "病人发卡管理"
    RestoreWinState Me, App.ProductName, mstrTitle
    Call CreateObjectPlugIn '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    Call CreateObjectKeyboard
    '69026,冉俊明,2014-8-8,检查输入年龄
    If CreatePublicPatient = False Then
        mblnUnLoad = True: Exit Sub
    End If
    mstrPriceGrade = gstr普通价格等级
    mstrPrePriceGrade = ""
    
    Call InitFace
    Call InitTaskPancel '初始化面版
    Call SetControlVisitble: Call Set权限

    '获取病人信息公共模块权限
    mstrPrivsPubPatient = ";" & GetPrivFunc(glngSys, 9003) & ";"
    mbln基本信息调整 = False
    mblnSaveDeposit = Val(zlDatabase.GetPara("剩余款缺省处理方式", glngSys, mlngModule, 0, , InStr(1, mstrPrivs, ";参数设置;") > 0)) = 1
    '初始化LED
    If gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModule, gcnOracle
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strName As String
    
    '115193:李南春,2017/10/13,卸载窗体时，清空模块变量
    '问题号:56599
    strName = IIf(glngSys \ 100 = 8, "客户的会员卡", "病人的医疗卡")
    If Not mblnUnLoad Then
        If mEditType = Cr_查询 Then
        ElseIf chkCancel.value = Checked Then
            If mblnChange Then
                If MsgBox("确定要放弃退卡操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
            End If
        ElseIf Not mrsInfo Is Nothing And (mEditType = Cr_发卡 Or mEditType = Cr_绑定卡) Then
            If mrsInfo.State = adStateOpen Then
                If MsgBox("该" & strName & "尚未发放,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
            End If
        End If
        If mblnChange Then
             If MsgBox("卡片信息已经发生改变，但你还未确认，是否真的要退出吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
        End If
    End If
    mlng图像操作 = 0: mstr采集图片 = "": Set mdic医疗卡属性 = Nothing
    If Not mobjPubPatient Is Nothing Then Set mobjPubPatient = Nothing
    Call SaveRegInFor(g私有模块, Me.Name, "idkind", IDKind.IDKind)
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled False
        Set mobjICCard = Nothing
    End If
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
    End If
    
    Set mobjReadCard = Nothing
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If Not mobjPlugIn Is Nothing Then Set mobjPlugIn = Nothing
    mlngPlugInHwnd = 0: mblnPlugin = False
    
    zlDatabase.SetPara "显示扩展信息", IIf(mParaData.blnShowExpend, 1, 0), glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    If mEditType = Cr_发卡 Or mEditType = Cr_绑定卡 Then
        '保存参数
        zlDatabase.SetPara "上次发卡类别", mCardType.lng卡类别ID, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
    mblnGetBirth = False
    SaveWinState Me, App.ProductName, mstrTitle
    Call UnHookKBD
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String
    If IsCardType(IDKind, "IC卡号") Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If
    lng卡类别ID = IDKind.GetCurCard.接口序号
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
    'Call InitInterFacel(Me, mlngModule, lng卡类别ID, False, mobjCardObject)
    strExpand = lng卡类别ID
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNo
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
    Exit Sub
 
End Sub

 
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
     '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    Set gobjSquare.objCurCard = objCard
    mlng医疗卡长度 = objCard.卡号长度
    '105667:李南春，2017/5/23，卡号加密导致第一个汉字拼音不能触发输入法
    txtPatient.PasswordChar = ""
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Then Exit Sub  'Or Not Me.ActiveControl Is txtPatient Or txtPatient.Text <> ""
    mblnNotClick = True
    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtPatient.Text = objPatiInfor.卡号
    Call txtPatient_KeyPress(vbKeyReturn)
    '118878:李南春,2018/1/4,如果还是卡号，则没找到病人
    If txtPatient.Text = objPatiInfor.卡号 Or txtPatient.Text = "" Then Call FromKindLoadPati(objPatiInfor)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub IDKindPayMode_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnNotChange Then Exit Sub
    mblnNotChange = True
    
    If Val(txt合计.Text) - Val(txt合计.Tag) < 0 Then
        IDKindPayMode.IDKind = 1 '为负数时不能充值
    ElseIf cbo支付方式.ListIndex >= 0 Then
        If cbo支付方式.ItemData(cbo支付方式.ListIndex) = -1 Then IDKindPayMode.IDKind = 2 '三方卡或消费卡不能找补
    End If
    mblnNotChange = False
End Sub

Private Sub lbl刷卡验证_Click()
    Dim strOutPut As String, strExpand As String, strOutXml As String, strCardNo As String, intFlag As Integer
    If mCardType.bln就诊卡 = False Then Exit Sub
    If Not (mEditType = Cr_退卡 Or chkCancel.value = 1) Then Exit Sub
    If mCardType.bln就诊卡 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt刷卡卡号.Text = mobjICCard.Read_Card()
            If Trim(txt刷卡卡号.Text) = "" Then Exit Sub
            intFlag = ReadCardNo(Trim(txt刷卡卡号.Text), 2)
            If intFlag = -1 Then
                If mEditType <> Cr_换卡 Then
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
                End If
            ElseIf intFlag = 2 Then
                Exit Sub
            Else
                Call zlControl.TxtSelAll(txt刷卡卡号)
                stbThis.Panels(2) = "没有发现该" & mCardType.str卡名称 & "的信息,可能未建档,请检查！"
                txt刷卡卡号.Text = ""
                Exit Sub
            End If
        End If
        Exit Sub
    End If
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Sub
    End If
    If mobjReadCard Is Nothing Then Exit Sub
    If mobjReadCard.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    txt刷卡卡号.Text = strCardNo
    If Trim(txt刷卡卡号.Text) = "" Then Exit Sub
    intFlag = ReadCardNo(Trim(txt刷卡卡号.Text), 2)
    If intFlag = -1 Then
        If mEditType <> Cr_换卡 Then
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
        End If
    ElseIf intFlag = 2 Then
        Exit Sub
    Else
        Call zlControl.TxtSelAll(txt刷卡卡号)
        stbThis.Panels(2) = "没有发现该" & mCardType.str卡名称 & "的信息,可能未建档,请检查！"
        txt刷卡卡号.Text = ""
        Exit Sub
    End If
End Sub

Private Sub picCard_Resize()
    Err = 0: On Error Resume Next
    With picCard
        If mEditType = Cr_绑定卡 Or mEditType = Cr_发卡 Then
            tbPageDo.Move 0, 0, .ScaleWidth, .ScaleHeight
            fraCard.Move 0, 0, tbPageDo.ScaleWidth, tbPageDo.ScaleHeight
        Else
            fraCard.Move 0, 0, .ScaleWidth, .ScaleHeight
        End If
        
    End With
End Sub
Private Sub picDrugAllergy_Resize()
'问题号:56599
    vsDrug.Left = picDrugAllergy.Left - 80
    vsDrug.Top = picDrugAllergy.Top - 380
    vsDrug.Height = picDrugAllergy.ScaleHeight
    vsDrug.Width = picDrugAllergy.ScaleWidth
End Sub

Private Sub picExpend_Resize()
'修改人:56599
Err = 0: On Error Resume Next
    With picExpend
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分页控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Dim intEditType As Integer '进入窗体时的操作类型
    
    Err = 0: On Error GoTo Errhand:
    If mEditType <> Cr_调整病人信息 Then
        Set objItem = tbPage.InsertItem(mPageIndex.常用, "常用", fraBase.hWnd, 0)
        objItem.Tag = mPageIndex.常用
        
        If mEditType = Cr_绑定卡 Or mEditType = Cr_发卡 Then
            Set objItem = tbPage.InsertItem(mPageIndex.病人证件, "证件信息", picCertificate.hWnd, 0)
            objItem.Tag = mPageIndex.病人证件
            Call InitCertificate
            
            Set objItem = tbPage.InsertItem(mPageIndex.药物过敏, "药物过敏", picDrugAllergy.hWnd, 0)
            objItem.Tag = mPageIndex.药物过敏
            Call InitvsDrug
            
            Set objItem = tbPage.InsertItem(mPageIndex.接种信息, "接种信息", picInoculate.hWnd, 0)
            objItem.Tag = mPageIndex.接种信息
            Call InitVsInoculate
            
            Set objItem = tbPage.InsertItem(mPageIndex.其他信息, "其他信息", picOtherInfo.hWnd, 0)
            objItem.Tag = mPageIndex.其他信息
            Call InitVsOtherInfo
            Call InitCombox
            
            '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
            If Not mobjPlugIn Is Nothing Then
                On Error Resume Next
                mlngPlugInHwnd = mobjPlugIn.GetFormHwnd
                Call zlPlugInErrH(Err, "GetFormHwnd")
                Err.Clear: On Error GoTo 0
                If mlngPlugInHwnd <> 0 Then
                    picTaskPanelOther.Visible = True
                    Set objItem = tbPage.InsertItem(mPageIndex.附加信息, "附加信息", picTaskPanelOther.hWnd, 0)
                    objItem.Tag = mPageIndex.附加信息
                End If
            End If
        Else
            picDrugAllergy.Visible = False
            picInoculate.Visible = False
            picOtherInfo.Visible = False
            picCertificate.Visible = False
        End If
         
         With tbPage
            tbPage.Item(0).Selected = True
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
        End With
        
        '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
        If mEditType = Cr_绑定卡 Or mEditType = Cr_发卡 Then
            intEditType = mEditType '记录操作类型，防止创建页面时被更改
            Set objItem = tbPageDo.InsertItem(0, "发卡", fraCard.hWnd, 0): objItem.Tag = Cr_发卡
            Set objItem = tbPageDo.InsertItem(1, "绑定卡", fraCard.hWnd, 0): objItem.Tag = Cr_绑定卡
            If intEditType = Cr_绑定卡 Then
                tbPageDo(1).Selected = True
            Else
                tbPageDo(1).Selected = True: tbPageDo(0).Selected = True
            End If
            With tbPageDo
                Call SetCardPayOrBound
                .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
                .PaintManager.BoldSelected = True
                .PaintManager.Layout = xtpTabLayoutAutoSize
                .PaintManager.StaticFrame = True
                .PaintManager.ClientFrame = xtpTabFrameSingleLine
            End With
        End If
    Else
        picDrugAllergy.Visible = False
        picInoculate.Visible = False
        picOtherInfo.Visible = False
        picCertificate.Visible = False
        tbPage.Visible = False
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub
Private Sub picInoculate_Resize()
'问题号:56599
    vsInoculate.Left = picInoculate.Left - 80
    vsInoculate.Top = picInoculate.Top - 380
    vsInoculate.Height = picInoculate.ScaleHeight
    vsInoculate.Width = picInoculate.ScaleWidth
End Sub

Private Sub picCertificate_Resize()
'问题号:90875
    vsCertificate.Left = picCertificate.Left - 80
    vsCertificate.Top = picCertificate.Top - 380
    vsCertificate.Height = picCertificate.ScaleHeight
    vsCertificate.Width = picCertificate.ScaleWidth
End Sub

Private Sub picTaskPanelOther_Resize()
    wndTaskPanelOther.Move 0, 0, picTaskPanelOther.Width, picTaskPanelOther.Height
End Sub

Private Sub txtAudi_Change()
    mblnChange = True
End Sub

Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    zlCommFun.OpenIme False
    Call OpenPassKeyboard(txtAudi, True)
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)

    Call CheckInputPassWord(KeyAscii, mCardType.int密码规则 = 1)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0

    If Not txt卡费.Locked And txt卡费.TabStop And txt卡费.Enabled And txt卡费.Visible Then txt卡费.SetFocus: Exit Sub
    If chk病历费.Visible And chk病历费.Enabled Then chk病历费.SetFocus: Exit Sub
    If Not txt病历费.Locked And txt病历费.TabStop And txt病历费.Enabled And txt病历费.Visible Then txt病历费.SetFocus: Exit Sub
    If chk记帐.Visible And chk记帐.Enabled Then chk记帐.SetFocus: Exit Sub
    If cbo支付方式.Visible And cbo支付方式.Enabled Then cbo支付方式.SetFocus: Exit Sub
    Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub
Private Sub txtAudi_LostFocus()

    Call ClosePassKeyboard(txtAudi)
End Sub

Private Sub txtAudi_Validate(Cancel As Boolean)
    If txtPass.Text <> txtAudi.Text And txtAudi.Text <> "" Then
        MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
        Cancel = 1
        Call zlControl.TxtSelAll(txtAudi)
        If txtAudi.Enabled And txtAudi.Visible Then txtAudi.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtPass_Change()
    mblnChange = True
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    zlCommFun.OpenIme False
    txtPass.MaxLength = 0
    '108779:李南春,2017/5/8,密码限制规则为N位以上时，不能超过密码长度
    Select Case mCardType.int密码长度限制
        Case 0
        Case Else
            txtPass.MaxLength = mCardType.int密码长度
            txtAudi.MaxLength = mCardType.int密码长度
    End Select
    Call OpenPassKeyboard(txtPass, False)
    If gblnLED Then Call chk记帐_Click
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCardType.int密码规则 = 1)

    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    If Not (txtPass.Text = "" And txtAudi.Text = "") Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If Not txt卡费.Locked And txt卡费.TabStop And txt卡费.Enabled And txt卡费.Visible Then txt卡费.SetFocus: Exit Sub
    If chk病历费.Visible And chk病历费.Enabled Then chk病历费.SetFocus: Exit Sub
    If Not txt病历费.Locked And txt病历费.TabStop And txt病历费.Enabled And txt病历费.Visible Then txt病历费.SetFocus: Exit Sub
    If chk记帐.Visible And chk记帐.Enabled Then chk记帐.SetFocus: Exit Sub
    If cbo支付方式.Visible And cbo支付方式.Enabled Then cbo支付方式.SetFocus: Exit Sub
    Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub
Private Sub txtPass_LostFocus()
    Call ClosePassKeyboard(txtPass)
End Sub

Private Sub txtPatient_Change()
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    Call AutoBrushSet(IDKind, txtPatient.Text = "")
    mblnChange = True
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    Call ReInitPatiInvoice
End Sub

Private Sub txt病历费_Change()
    mblnChange = True
End Sub

Private Sub txt病历费_GotFocus()
    zlControl.TxtSelAll txt病历费
    zlCommFun.OpenIme False
End Sub
Private Sub txt病历费_KeyPress(KeyAscii As Integer)
    If txt病历费.Locked Then Exit Sub
    zlControl.TxtCheckKeyPress txt病历费, KeyAscii, m金额式
    If KeyAscii <> vbKeyReturn Then Exit Sub
    KeyAscii = 0
    If mFeeType.bln变价 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If mFeeType.rs病历费 Is Nothing Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If mFeeType.rs病历费!现价 <> 0 And Abs(CCur(txt病历费.Text)) > Abs(mFeeType.rs病历费!现价) Then
        MsgBox "病历费金额绝对值不能大于最高限价：" & Format(Abs(mFeeType.rs病历费!现价), "0.00"), vbExclamation, gstrSysName
        If txt病历费.Enabled And txt病历费.Visible Then txt病历费.SetFocus: Call zlControl.TxtSelAll(txt病历费): Exit Sub
    End If
    If mFeeType.rs病历费!原价 <> 0 And Abs(CCur(txt病历费.Text)) < Abs(mFeeType.rs病历费!原价) Then
        MsgBox "病历费金额绝对值不能小于最低限价：" & Format(Abs(mFeeType.rs病历费!原价), "0.00"), vbExclamation, gstrSysName
        If txt病历费.Enabled And txt病历费.Visible Then txt病历费.SetFocus: Call zlControl.TxtSelAll(txt病历费): Exit Sub
    End If
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt病历费_Validate(Cancel As Boolean)
    Call txt余额_Change
End Sub

Private Sub txt出生地点_Change()
    mblnChange = True: lbl出生地点.Tag = ""
End Sub

Private Sub txt出生地点_GotFocus()
    zlControl.TxtSelAll txt出生地点
    zlCommFun.OpenIme True
    cmd出生地点.CausesValidation = False
End Sub

Private Sub txt出生地点_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl出生地点.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt出生地点) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select地区(txt出生地点, lbl出生地点, Trim(txt出生地点)) = False Then Exit Sub
End Sub

Private Sub txt出生地点_LostFocus()
      zlCommFun.OpenIme False
End Sub

Private Sub txt出生地点_Validate(Cancel As Boolean)
    If Not Check必须输入项(txt出生地点) Then
        Cancel = True
    Else
        cmd出生地点.CausesValidation = True
    End If
End Sub

Private Sub txt出生日期_Change()
    Dim str出生时间 As String
    If IsDate(txt出生日期.Text) And Not mblnNotChange Then
        mblnNotChange = True
        txt出生日期.Text = Format(CDate(txt出生日期.Text), "yyyy-mm-dd")
        mblnNotChange = False
        
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        mstr年龄 = txt年龄.Text: mstr年龄单位 = cbo年龄单位.Text
        '111836:李南春，2017/7/21，年龄控件位置计算
        If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
            cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False: txt年龄.Width = 1220
        Else
            cbo年龄单位.Visible = True: txt年龄.Width = 550
            If cbo年龄单位.ListIndex = -1 Then cbo年龄单位.ListIndex = 0
        End If
        '年龄由生日产生后，不再允许由年龄产生生日
        mblnGetBirth = False
    End If
End Sub

Private Sub txt出生日期_LostFocus()
    If txt出生日期.Text <> "____-__-__" And Not IsDate(txt出生日期.Text) Then
        If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
    End If
End Sub

Private Sub txt出生时间_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt出生日期) Then
        KeyAscii = 0
        txt出生时间.Text = "__:__"
    End If
End Sub

 Private Sub txt出生时间_Validate(Cancel As Boolean)
    Dim str出生时间 As String
    '76669，李南春,2014-8-18,病人年龄更新
    If txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
        If txt出生时间.Enabled And txt出生时间.Visible Then txt出生时间.SetFocus
        Cancel = True
    ElseIf IsDate(txt出生日期.Text) Then
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        mstr年龄 = txt年龄.Text: mstr年龄单位 = cbo年龄单位.Text
    End If
End Sub

Private Sub txt单位电话_Change()
    mblnChange = True
End Sub

Private Sub txt单位电话_GotFocus()
    zlControl.TxtSelAll txt单位电话
    zlCommFun.OpenIme False
End Sub

Private Sub txt单位电话_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt单位电话_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(txt单位电话)
End Sub

Private Sub txt单位开户行_Change()
    mblnChange = True
End Sub

Private Sub txt单位开户行_GotFocus()
    zlControl.TxtSelAll txt单位开户行
    zlCommFun.OpenIme True
End Sub
Private Sub txt单位开户行_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt单位开户行_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(txt单位开户行)
End Sub

Private Sub txt单位邮编_Change()
    mblnChange = True
End Sub

Private Sub txt单位邮编_GotFocus()
    zlControl.TxtSelAll txt单位邮编
    zlCommFun.OpenIme False
End Sub

Private Sub txt单位邮编_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(txt单位邮编)
End Sub

Private Sub txt单位帐户_Change()
    mblnChange = True
End Sub

Private Sub txt单位帐户_GotFocus()
    zlControl.TxtSelAll txt单位帐户
    zlCommFun.OpenIme False
End Sub

Private Sub txt单位帐户_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(txt单位帐户)
End Sub

Private Sub txt工作单位_Change()
    mblnChange = True: lbl工作单位.Tag = ""
End Sub

Private Sub txt工作单位_GotFocus()
    zlControl.TxtSelAll txt工作单位
    zlCommFun.OpenIme True
    cmd合同单位.CausesValidation = False
End Sub

Private Sub txt工作单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl工作单位.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt工作单位) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select合约单位(Trim(txt工作单位.Text)) = False Then Exit Sub
End Sub

Private Sub txt工作单位_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt工作单位_Validate(Cancel As Boolean)
    If Not Check必须输入项(txt工作单位) Then
        Cancel = True
    Else
        cmd合同单位.CausesValidation = True
    End If
End Sub

Private Sub txt合计_Change()
    Call txt余额_Change
End Sub

Private Sub txt合计_GotFocus()
    zlControl.TxtSelAll txt合计
    zlCommFun.OpenIme False
End Sub
Private Sub txt合计_KeyPress(KeyAscii As Integer)
    If txt合计.Locked Or txt合计.Enabled = False Then Exit Sub
    zlControl.TxtCheckKeyPress txt合计, KeyAscii, m金额式
End Sub

Private Sub txt合计_LostFocus()
    If gblnLED And chk记帐.value = 0 And Val(txt合计.Text) > Val(txt合计.Tag) Then
        zl9LedVoice.DispCharge Format(txt合计.Tag, "0.00"), txt合计.Text, IIf(IDKindPayMode.IDKind = 2, 0, txt余额.Text)
        zl9LedVoice.Speak "#22 " & txt合计.Text
        zl9LedVoice.Speak "#23 " & IIf(IDKindPayMode.IDKind = 2, 0, txt余额.Text)
        zl9LedVoice.Speak "#3 "
    End If
End Sub

Private Sub txt合计_Validate(Cancel As Boolean)
    txt合计.Text = Format(txt合计.Text, "0.00")
End Sub

Private Sub txt户口地址_Change()
    mblnChange = True
    txt户口地址.Tag = ""
End Sub

Private Sub txt户口地址_GotFocus()
    Call zlControl.TxtSelAll(txt户口地址)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt户口地址_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Trim(txt户口地址.Text) <> "" Then
        Call SearchAddress(Trim(txt户口地址.Text), txt户口地址)
    End If
End Sub

Private Sub txt户口地址_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub SearchAddress(ByVal strInput As String, txtInput As Object)
    '--------------------------------------------------------------
    '功能:模糊查找，弹出地区选择列表
    '编制:冉俊明
    '日期:2014-5-23
    '参数:
    '   strInput:输入文本，若为空表示点击按钮进入
    '   txtInput:文本框对象
    '--------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
    vRect = zlControl.GetControlRect(txtInput.hWnd)
    If strInput = "" Then '点击按钮
        strSQL = "" & _
            "Select ID, 上级id, 编码, 名称, 末级 " & _
            "From (With 地区_t As" & _
            "    (Select Rownum As 行号, ID, 上级id, 末级, 编码, 名称" & _
            "     From (Select Distinct Substr(名称, 1, 2) As ID, Null As 上级id, 0 As 末级, Null As 编码, Substr(名称, 1, 2) As 名称" & _
            "            From 地区" & _
            "            Union All" & _
            "            Select 编码 As ID, Substr(名称, 1, 2) As 上级id, 1 As 末级, 编码, 名称 From 地区))" & _
            "   Select 行号 As ID, To_Number(上级id) As 上级id, 编码, 名称, 末级 From 地区_t Where 上级id Is Null" & _
            "   Union All" & _
            "   Select b.行号, a.行号, b.编码, b.名称, b.末级 From 地区_t A, 地区_t B Where a.Id = b.上级id Order By 编码)"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "地区", False, _
                       "", "", False, False, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        '去掉"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '输入全是数字时只匹配编码
                strWhere = " Where 编码 Like Upper([1])"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then '输入全是字母时只匹配简码
                strWhere = " Where 简码 Like Upper([1])"
            Else
                strWhere = " Where 编码 Like Upper([1]) Or 名称 Like [1] Or 简码 Like Upper([1])"
            End If
        End If
        
        strSQL = "" & _
            "Select Rownum As ID, 编码, 名称 From 地区 " & strWhere & " Order By 编码"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "地区", False, _
                       "", "", False, False, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False, strKey)
    End If
    If blnCancel Then txtInput.SetFocus: Exit Sub

    If rsTemp Is Nothing Then txtInput.SetFocus: Exit Sub
    If rsTemp.State <> 1 Then txtInput.SetFocus: Exit Sub
    
    txtInput.Text = Nvl(rsTemp!名称)
    txtInput.Tag = Nvl(rsTemp!id)
    txtInput.SelStart = Len(txtInput.Text)
    txtInput.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt户口地址邮编_Change()
    mblnChange = True
End Sub

Private Sub txt户口地址邮编_GotFocus()
    Call zlControl.TxtSelAll(txt户口地址邮编)
End Sub

Private Sub txt户口地址邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt户口地址邮编_Validate(Cancel As Boolean)
     Cancel = Not Check必须输入项(txt户口地址邮编)
End Sub

Private Sub txt家庭邮编_Change()
    mblnChange = True
End Sub

Private Sub txt家庭邮编_GotFocus()
    zlControl.TxtSelAll txt家庭邮编
    zlCommFun.OpenIme False
End Sub

Private Sub txt家庭地址_Change()
    mblnChange = True
    lbl家庭地址.Tag = ""
End Sub

Private Sub txt家庭地址_GotFocus()
    zlControl.TxtSelAll txt家庭地址
    zlCommFun.OpenIme True
End Sub

Private Sub txt家庭地址_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl家庭地址.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt家庭地址) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select地区(txt家庭地址, lbl家庭地址, Trim(txt家庭地址)) = False Then Exit Sub
End Sub
 

Private Sub txt家庭地址_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt家庭电话_Change()
    mblnChange = True
End Sub

Private Sub txt家庭电话_GotFocus()
    zlControl.TxtSelAll txt家庭电话
    zlCommFun.OpenIme False
End Sub

Private Sub txt卡费_Change()
    mblnChange = True
End Sub

Private Sub txt卡费_GotFocus()
    zlControl.TxtSelAll txt卡费
    zlCommFun.OpenIme False
End Sub
Private Sub txt卡费_KeyPress(KeyAscii As Integer)
    If txt卡费.Locked Then Exit Sub
    zlControl.TxtCheckKeyPress txt卡费, KeyAscii, m金额式
    If KeyAscii <> vbKeyReturn Then Exit Sub
    KeyAscii = 0
    If mCardType.bln变价 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If mCardType.rs医疗卡费 Is Nothing Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If mCardType.rs医疗卡费!现价 <> 0 And Abs(CCur(txt卡费.Text)) > Abs(mCardType.rs医疗卡费!现价) Then
        MsgBox mCardType.str卡名称 & "金额绝对值不能大于最高限价：" & Format(Abs(mCardType.rs医疗卡费!现价), "0.00"), vbExclamation, gstrSysName
        If txt卡费.Enabled And txt卡费.Visible Then txt卡费.SetFocus: Call zlControl.TxtSelAll(txt卡费): Exit Sub
    End If
    If mCardType.rs医疗卡费!原价 <> 0 And Abs(CCur(txt卡费.Text)) < Abs(mCardType.rs医疗卡费!原价) Then
        MsgBox mCardType.str卡名称 & "卡金额绝对值不能小于最低限价：" & Format(Abs(mCardType.rs医疗卡费!原价), "0.00"), vbExclamation, gstrSysName
        If txt卡费.Enabled And txt卡费.Visible Then txt卡费.SetFocus: Call zlControl.TxtSelAll(txt卡费): Exit Sub
    End If
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt卡费_Validate(Cancel As Boolean)
    Call txt余额_Change
End Sub

Private Sub txt卡号_Change()
    Dim rsTemp As Recordset

    mblnChange = True
    Call SetCardEditEnabled
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    Call AutoBrushSet(IDKindPay, txt卡号.Text = "")
    '问题号:53408
    If mCardType.str卡名称 = "二代身份证" Then
        Call OpenIDCard(txt卡号.Text = "")
        If Len(txt卡号.Text) = mCardType.lng卡号长度 Then
            Set rsTemp = zl是否已绑定(Trim(txt卡号.Text))
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount <= 0 Then Exit Sub
            If MsgBox("卡号为:" & txt身份证号.Text & "已经被病人:" & rsTemp!姓名 & "绑定,是否要取消已绑定的身份证号", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
                frmPaticurCardCancelBound.zlCancelBand Me, mlngModule, mlngCardTypeID, rsTemp!病人ID, txt卡号.Text, False
            End If
        End If

    End If
End Sub

Private Sub txt卡号_GotFocus()
    '76609,冉俊明,2014-8-14,刷卡机刷卡末尾可能存在有回车符焦点定位问题
    mblnTab = False
    If Not txt卡号.Enabled Then Exit Sub
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    Call AutoBrushSet(IDKindPay, txt卡号.Text = "")
    zlControl.TxtSelAll txt卡号
    zlCommFun.OpenIme False
    '问题号:53408
    If mCardType.str卡名称 = "二代身份证" Then
        Call OpenIDCard(txt卡号.Text = "")
    End If
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    '问题号:53408
    If mCardType.str卡名称 = "二代身份证" Or mCardType.str卡名称 = "IC卡" Then
        KeyAscii = 0
    End If

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> 13 Then
        If Len(txt卡号.Text) = mCardType.lng卡号长度 - 1 And KeyAscii <> 8 Then
            '76609,冉俊明,2014-8-14,刷卡机刷卡末尾可能存在有回车符焦点定位问题
            mblnTab = True
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Call EnableKBDHook
        End If
    ElseIf txt卡号.Text = "" Then
        KeyAscii = 0: cmdOK.SetFocus  '不发卡,直接跳过
    Else
        KeyAscii = 0: If Not mblnTab Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开身份证读卡器
    '编制:王吉
    '日期:2012-08-31 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '初始化对卡对象
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '打开读卡器
    mobjIDCard.SetEnabled (blnEnabled)
End Sub

Private Sub txt卡号_LostFocus()
    '76609,冉俊明,2014-8-14,刷卡机刷卡末尾可能存在有回车符焦点定位问题
    mblnTab = False
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    '97702,李南春,2016/6/28,焦点移除后关闭自动读卡
    Call AutoBrushSet(IDKindPay, False)
    Call zlCommFun.OpenIme(False)
    If mCardType.str卡名称 = "二代身份证" Then
        Call OpenIDCard(False)
    End If
    Call ReLoadCardFee
End Sub

Private Sub ReLoadCardFee(Optional ByVal blnFeedName As Boolean)
    'blnFeedName-是否姓名处检查，减少建档病人修改其它信息产生的调用
    '118124:李南春,2018/1/18,获取卡费
    Dim lng病人ID As Long, lng收费细目ID As Long
    Dim strSQL As String, str年龄 As String
    Dim rsTmp As ADODB.Recordset
    
    If (mEditType <> Cr_发卡 And mEditType <> Cr_绑定卡 And mEditType <> Cr_补卡) Or chkCancel.value = 1 Then Exit Sub
    If mCardType.rs医疗卡费 Is Nothing Then Exit Sub
    If mCardType.rs医疗卡费.RecordCount = 0 Then Exit Sub
    If mCardType.lng卡类别ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(txt卡号.Text) = "" Then Exit Sub
    If Trim(txt年龄.Text) = "" Then Exit Sub
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = mrsInfo!病人ID
    End If
    If blnFeedName = False And lng病人ID <> 0 Then Exit Sub
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    mCardType.rs医疗卡费.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as 收费细目ID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "卡费", mlngModule, mCardType.lng卡类别ID, Trim(txt卡号.Text), lng病人ID, _
                Trim(txtPatient.Text), zlstr.NeedName(cbo性别.Text), str年龄, txt身份证号.Text, Val(Nvl(mCardType.rs医疗卡费!收费细目ID)))
    If rsTmp.EOF Then Exit Sub
    
    lng收费细目ID = Val(Nvl(rsTmp!收费细目ID))
    Set rsTmp = zlGetSpecialItemFee(mCardType.str特定项目, mstrPriceGrade, lng收费细目ID)
    If Not rsTmp Is Nothing Then Set mCardType.rs医疗卡费 = rsTmp
    Call LoadCardFee
End Sub

Private Sub txt联系人地址_Change()
    mblnChange = True
End Sub

Private Sub txt联系人地址_GotFocus()
    zlControl.TxtSelAll txt联系人地址
    zlCommFun.OpenIme True
    cmd联系人地址.CausesValidation = False
End Sub
 

Private Sub txt联系人地址_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl联系人地址.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt联系人地址) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select地区(txt联系人地址, lbl联系人地址, Trim(txt联系人地址)) = False Then Exit Sub
End Sub

Private Sub txt联系人地址_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt联系人地址_Validate(Cancel As Boolean)
    If Not Check必须输入项(txt联系人地址) Then
        Cancel = True
    Else
        cmd联系人地址.CausesValidation = True
    End If
End Sub

Private Sub txt联系人电话_Change()
    mblnChange = True
End Sub

Private Sub txt联系人电话_GotFocus()
    zlControl.TxtSelAll txt联系人电话
    zlCommFun.OpenIme False
End Sub

Private Sub txt联系人电话_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(txt联系人电话)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("电话") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("电话")) = txt联系人电话.Text
    End If
End Sub

Private Sub txt联系人身份证号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt联系人身份证号_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(txt联系人身份证号)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("身份证号") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("身份证号")) = txt联系人身份证号.Text
    End If
End Sub

Private Sub txt联系人姓名_Change()
    mblnChange = True
End Sub

Private Sub txt联系人姓名_GotFocus()
    zlControl.TxtSelAll txt联系人姓名
    zlCommFun.OpenIme True
End Sub

Private Sub txt联系人姓名_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt联系人姓名_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(txt联系人姓名)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("姓名") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("姓名")) = txt联系人姓名.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt联系人姓名.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt门诊号_Change()
    mblnChange = True
End Sub

Private Sub txt门诊号_GotFocus()
    '94941:李南春,2016/4/7,修改门诊号权限
    If InStr(";" & mstrPrivs & ";", ";允许修改门诊号;") > 0 Then
        zlControl.TxtSelAll txt门诊号
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    '94941:李南春,2016/4/7,修改门诊号权限
    If KeyAscii = vbKeySpace Then
        txt门诊号.Text = zlGet门诊号: KeyAscii = 0: Exit Sub
    End If
    If InStr(";" & mstrPrivs & ";", ";允许修改门诊号;") <= 0 Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txt门诊号, KeyAscii, m数字式
End Sub
Private Sub txt年龄_Change()
    mblnChange = True
End Sub

Private Sub txt年龄_GotFocus()
    Call zlCommFun.OpenIme
    zlControl.TxtSelAll txt年龄
End Sub
Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo年龄单位.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
            Call cbo年龄单位.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt年龄.Text) And cbo年龄单位.Visible Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt年龄_Validate(Cancel As Boolean)
    Dim strBirth As String
     '111836:李南春，2017/7/21，年龄反算
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False: txt年龄.Width = 1220
    ElseIf cbo年龄单位.Visible = False Then
        cbo年龄单位.Visible = True: txt年龄.Width = 550
        If cbo年龄单位.ListIndex = -1 Then cbo年龄单位.ListIndex = 0
    End If
    '69026,冉俊明,2014-8-8,检查输入年龄
    '76703,冉俊明,2014-8-15
    If mobjPubPatient Is Nothing Then Exit Sub
    If txt年龄.Text <> mstr年龄 Then
        mstr年龄 = txt年龄.Text
        If Not IsDate(txt出生日期.Text) Then mblnGetBirth = True
        If cbo年龄单位.Visible Then mstr年龄单位 = "": Exit Sub
        mblnNotChange = True
        
        If mblnGetBirth Then
            '103807:李南春，2016/12/20，年龄反算精确到小时
            If mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), strBirth) Then
                txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                txt出生时间.Text = Format(strBirth, "hh:mm")
            End If
        End If
        mblnNotChange = False
    End If

    If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), _
            IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
            IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then Cancel = True: Exit Sub
End Sub

Private Sub txt其他关系_Change()
    mblnChange = True
End Sub

Private Sub txt其他关系_GotFocus()
    zlControl.TxtSelAll txt其他关系
    zlCommFun.OpenIme True
End Sub

Private Sub txt其他关系_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt其他关系_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("附加信息") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("关系")) = zlstr.NeedName(cbo联系人关系.Text)
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("附加信息")) = txt其他关系.Text
    End If
End Sub

Private Sub txt其他证件_Change()
    mblnChange = True
End Sub
Private Sub txt其他证件_GotFocus()
    zlControl.TxtSelAll txt其他证件
    zlCommFun.OpenIme False
End Sub

Private Sub txt其他证件_Validate(Cancel As Boolean)
    Cancel = Not Check必须输入项(txt其他证件)
End Sub

Private Sub txt区域_Change()
    mblnChange = True: lbl区域.Tag = ""
End Sub

Private Sub txt区域_GotFocus()
    zlControl.TxtSelAll txt区域
    zlCommFun.OpenIme True
End Sub

Private Sub txt区域_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lbl区域.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt区域) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select区域(Trim(txt区域)) = False Then Exit Sub
End Sub

Private Sub txt区域_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt身份证号_Change()
    Dim strDate As String
    mblnChange = True
    '当不能更改病人基本信息时,出生日期不能返算67184
    If Not mblnNotChange And txt出生日期.Enabled Then
        strDate = zlCommFun.GetIDCardDate(txt身份证号.Text)
        If IsDate(strDate) Then txt出生日期.Text = strDate
    End If
End Sub
Private Sub txt身份证号_GotFocus()
    zlControl.TxtSelAll txt身份证号
    zlCommFun.OpenIme False
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt手机_Change()
    mblnChange = True
End Sub

Private Sub txt手机_GotFocus()
    zlControl.TxtSelAll txt手机
    zlCommFun.OpenIme False
End Sub

Private Sub txt手机_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt手机, KeyAscii, m数字式)
End Sub

Private Sub txt手机_Validate(Cancel As Boolean)
    
    If CheckMobile(txt手机.Text) = False Then Cancel = True
End Sub

Private Sub txt刷卡卡号_Change()
    lbl刷卡验证.Tag = ""
End Sub

Private Sub txt刷卡卡号_GotFocus()
   zlControl.TxtSelAll txt刷卡卡号
End Sub

Private Sub txt刷卡卡号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng病人ID As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    If txt刷卡卡号.Text = "" Then
        If zlShowSelectCardNo(lng病人ID, "") = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt刷卡卡号_KeyPress(KeyAscii As Integer)
   Dim strCardNo As String, intFlag As Integer
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii <> 13 Then
        If Len(txt刷卡卡号.Text) = mCardType.lng卡号长度 - 1 And KeyAscii <> 8 Then
            stbThis.Panels(2) = ""
            txt刷卡卡号.Text = txt刷卡卡号.Text & Chr(KeyAscii)
             strCardNo = Trim(txt刷卡卡号)
             KeyAscii = 0:
            intFlag = ReadCardNo(strCardNo, 2)
            If intFlag = -1 Then
                If mEditType <> Cr_换卡 Then
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus: Exit Sub
                End If
            ElseIf intFlag = 2 Then
                Exit Sub
            Else
                Call zlControl.TxtSelAll(txt刷卡卡号)
                stbThis.Panels(2) = "没有发现该就诊卡的信息,可能未建档,请检查！"
                txt刷卡卡号.Text = ""
                Exit Sub
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        Exit Sub
    End If
    stbThis.Panels(2) = ""
    If lbl刷卡验证.Tag = Trim(txt刷卡卡号.Text) Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    strCardNo = Trim(txt刷卡卡号)
    intFlag = ReadCardNo(strCardNo, 2)
    If intFlag = -1 Then
        If mEditType <> Cr_换卡 Then
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        End If
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab: Exit Sub
    ElseIf intFlag = 2 Then
        Exit Sub
    Else
        If (chkCancel.value = 1 Or mEditType = Cr_退卡) And mParaData.int退卡模式 = 2 And Trim(cboNO.Text) = "" Then
            Call zlControl.TxtSelAll(cboNO)
           If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Else
            Call zlControl.TxtSelAll(txt刷卡卡号)
        End If
        stbThis.Panels(2) = "没有发现该就诊卡的信息,可能未建档,请检查！"
    End If
End Sub

Private Sub txt验证医保号_Validate(Cancel As Boolean)

    txt验证医保号.Text = UCase(Trim(txt验证医保号.Text))
    If cbo医疗付款.ListCount > 0 And txt验证医保号.Text <> "" Then cbo医疗付款.ListIndex = 0
    If txt验证医保号.Text <> txt医保号.Text Then
        MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
        Exit Sub
    End If
    If mInsurePara.lng外挂式医保险类 = 920 And txt医保号.Text <> lbl医保号(0).Tag And txt医保号.Text <> "" Then
        If CheckExistsMCNO(txt医保号.Text) Then
             'Cancel = True
        End If
    End If
    Cancel = Not Check必须输入项(txt验证医保号)
End Sub

Private Sub txt医保号_Change()
    mblnChange = True
End Sub
Private Sub txt医保号_GotFocus()
    zlControl.TxtSelAll txt医保号
    zlCommFun.OpenIme False
End Sub

Private Sub txt医保号_Validate(Cancel As Boolean)
    txt医保号.Text = UCase(Trim(txt医保号.Text))
    If cbo医疗付款.ListCount > 0 And txt医保号.Text <> "" Then cbo医疗付款.ListIndex = 0
    If mInsurePara.lng外挂式医保险类 = 920 And txt医保号.Text <> lbl医保号(0).Tag And txt医保号.Text <> "" Then
        If CheckExistsMCNO(txt医保号.Text) Then
             'Cancel = True
        End If
    End If
    Cancel = Not Check必须输入项(txt医保号)
End Sub

Private Sub txt余额_Change()
    If mblnNotChange = True Then Exit Sub
    If chk记帐.value = Checked Then txt余额.Text = "": Exit Sub
    mblnNotChange = True
    txt合计.Tag = IIf(txt卡费.Visible, Val(txt卡费.Text), 0) + IIf(chk病历费.value, Val(txt病历费.Text), 0)
    If mEditType = Cr_退卡 Or chkCancel.value = 1 Then txt合计.Text = Format(txt合计.Tag, "0.00")
    txt余额.Text = Format(Val(txt合计.Text) - Val(txt合计.Tag), "0.00")
    
    txt余额.ForeColor = IIf(Val(txt余额.Text) < 0, vbRed, &H80000008)
    If Val(txt余额.Text) < 0 Then
        IDKindPayMode.IDKind = 1
        IDKindPayMode.GetCurCard.名称 = "应收"
        txt余额.Text = Format(-1 * Val(txt余额.Text), "0.00")
    Else
        If cbo支付方式.Text = "支票" And IDKindPayMode.IDKind = 1 Then
            IDKindPayMode.GetCurCard.名称 = "退支票"
        ElseIf IDKindPayMode.IDKind = 1 And cbo支付方式.ListIndex >= 0 Then
            If cbo支付方式.ItemData(cbo支付方式.ListIndex) = -1 Then
                IDKindPayMode.IDKind = 2
            Else
                IDKindPayMode.GetCurCard.名称 = "找补"
            End If
        End If
        If mblnSaveDeposit And Val(txt合计.Text) - Val(txt合计.Tag) > 0 Then
            IDKindPayMode.IDKind = 2
        End If
    End If
    If Not IDKindPayMode.GetCurCard Is Nothing Then IDKindPayMode.IDKind = IDKindPayMode.GetKindIndex(IDKindPayMode.GetCurCard.名称)
    mblnNotChange = False
End Sub

Private Sub wndTaskPanel_GroupExpanded(ByVal Group As XtremeSuiteControls.ITaskPanelGroup)
        If Group.id = Idx_TP_PatiExpend Then
            mParaData.blnShowExpend = Group.Expanded
            Call SetCtrlMove
        End If
End Sub
Private Sub SetCtrlMove()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的缺省位置
    '编制:刘兴洪
    '日期:2011-07-12 08:45:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngTaskHeight As Single, sngWinHeight As Single
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    Dim vRectForm As RECT, vRect As RECT
    Dim sinW As Single, sinH As Single
    
    Err = 0: On Error Resume Next
    If mParaData.blnShowExpend Then
        sngTaskHeight = mFormMaxHeight - 100 - stbThis.Height
        sngWinHeight = mFormMaxHeight + 400
    Else
        sngTaskHeight = mFormMaxHeight - 100 - picExpend.Height - stbThis.Height
        sngWinHeight = mFormMaxHeight - picExpend.Height + 400
    End If
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    picCard.Height = 2155
    If mEditType <> Cr_绑定卡 And mEditType <> Cr_发卡 Then
        picCard.Height = 1550
        sngTaskHeight = sngTaskHeight - 1050
        sngWinHeight = sngWinHeight - 1100
        If mEditType = Cr_换卡 And mbln病历费 Then
            picCard.Height = 2300
            sngTaskHeight = sngTaskHeight + 500
            sngWinHeight = sngWinHeight + 500
        End If
    Else
        If mEditType = Cr_发卡 Then
            picCard.Height = picCard.Height - cbo支付方式.Height * 2 + 420
            sngTaskHeight = sngTaskHeight - cbo支付方式.Height
            sngWinHeight = sngWinHeight - cbo支付方式.Height
            If mbln病历费 Then
                chk病历费.Left = txt卡费.Left + txt卡费.Width + 100: txt病历费.Left = chk病历费.Left + chk病历费.Width
                chk记帐.Left = txt病历费.Left + txt病历费.Width + 100
            Else
                chk记帐.Left = txt卡费.Left + txt卡费.Width + 100
            End If
        Else
            '无相关的卡费信息
            If mbln病历费 Then
                picCard.Height = picCard.Height - cbo支付方式.Height * 2 + 420
                sngTaskHeight = sngTaskHeight - cbo支付方式.Height
                sngWinHeight = sngWinHeight - cbo支付方式.Height
            Else
                picCard.Height = picCard.Height - cbo支付方式.Height * 2 - 50
                sngTaskHeight = sngTaskHeight - cbo支付方式.Height - 400
                sngWinHeight = sngWinHeight - cbo支付方式.Height - 400
            End If
            If mbln病历费 Then
                chk病历费.Left = lbl卡费.Left + 215: txt病历费.Left = chk病历费.Left + chk病历费.Width
                chk记帐.Left = txt病历费.Left + txt病历费.Width + 100
            Else
                chk记帐.Left = txt卡费.Left + txt卡费.Width + 100
            End If
        End If
        If Not mblnAddPage Then
            If mEditType = Cr_发卡 Or mbln病历费 Then
                picCard.Height = picCard.Height - 350
                sngTaskHeight = sngTaskHeight - 350: sngWinHeight = sngWinHeight - 350
                
            Else
                picCard.Height = picCard.Height - 750
                sngTaskHeight = sngTaskHeight - 800: sngWinHeight = sngWinHeight - 850
            End If
        Else
            If mEditType = Cr_绑定卡 Then
                picCard.Height = picCard.Height - 400
                sngTaskHeight = sngTaskHeight - 400: sngWinHeight = sngWinHeight - 400
            End If
        End If
    End If
    '重新加载一次发卡页面
    wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Clear
    Set Item = wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Add(Idx_TP_PatiCard, "", xtpTaskItemTypeControl)
    Set Item.Control = picCard: tkpGroup.Expanded = True
    wndTaskPanel.Reposition
    
    If mEditType = Cr_换卡 Then
        lbl卡号.Top = lbl卡费.Top: lbl密码.Top = lbl卡费.Top: lbl验证.Top = lbl卡费.Top
        txt卡号.Top = txt卡费.Top: txtAudi.Top = txt卡号.Top: txtPass.Top = txt卡号.Top
        txt刷卡卡号.Left = txt卡费.Left: lbl刷卡验证.Left = txt刷卡卡号.Left - lbl刷卡验证.Width - 20
        txt刷卡卡号.Width = txt卡号.Width
        '问题号:50893
        lbl原卡密码.Top = lbl刷卡验证.Top: txt原卡密码.Top = lbl原卡密码.Top - (txt原卡密码.Height - lbl原卡密码.Height) / 2
        lbl原卡密码.Left = txt原卡密码.Left - lbl原卡密码.Width - 50
        
        If mbln病历费 Then
            chk病历费.Left = txt卡费.Left: txt病历费.Left = chk病历费.Left + chk病历费.Width
            chk记帐.Left = txt病历费.Left + txt病历费.Width + 100
                
            sinH = txt卡费.Top + 450
            chk病历费.Top = sinH + 50: txt病历费.Top = sinH
            chk记帐.Top = sinH + 50
            cbo支付方式.Top = sinH: txt合计.Top = sinH
            picCard.Height = picCard.Height - txt病历费.Height + 50
            
            sinH = lbl卡费.Top + 450
            lbl支付方式.Top = sinH
            
            sinH = txt病历费.Top + 450
            txt操作员.Top = sinH: txtDate.Top = sinH
            
            sinH = lbl支付方式.Top + 450
            lbl发卡人.Top = sinH: lblDate.Top = sinH
            
            IDKindPayMode.Top = sinH - 60: txt余额.Top = sinH - 50
            
            '重新加载一次发卡页面
            wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Clear
            Set Item = wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Add(Idx_TP_PatiCard, "", xtpTaskItemTypeControl)
            Set Item.Control = picCard: tkpGroup.Expanded = True
            wndTaskPanel.Reposition
            sngTaskHeight = sngTaskHeight - 50
            sngWinHeight = sngWinHeight - 50
        End If
    End If
    If mEditType = Cr_挂失 Then
        txt刷卡卡号.Left = txt卡费.Left: lbl刷卡验证.Left = txt刷卡卡号.Left - lbl刷卡验证.Width - 50
        txt刷卡卡号.Width = txt卡号.Width
    End If
    
    If mEditType = Cr_补卡 Or mEditType = Cr_退卡 Or mEditType = Cr_查询 Then
        If mbln病历费 Then
            chk病历费.Left = txt卡费.Left + txt卡费.Width + 100: txt病历费.Left = chk病历费.Left + chk病历费.Width
            chk记帐.Left = txt病历费.Left + txt病历费.Width + 100
        Else
            chk记帐.Left = txt卡费.Left + txt卡费.Width + 100
        End If
    End If
    
    '104726:李南春，2017/4/24，显示票据
    If mEditType <> Cr_退卡 And Not (gbln收费发票 And (mEditType = Cr_补卡 Or mEditType = Cr_换卡)) Then
        sngTaskHeight = sngTaskHeight - picTittle.Height + 150
        sngWinHeight = sngWinHeight - picTittle.Height + 150
    End If
    
    If gbln收费发票 And (mEditType = Cr_补卡 Or mEditType = Cr_换卡) Then
        cboNO.Visible = False: lblNo.Visible = False: chkCancel.Visible = False
        lblFact.Left = 7700: txtFact.Left = 8230
    End If
    
    If mEditType = Cr_调整病人信息 Then
        sngTaskHeight = sngTaskHeight - picCard.Height - picTittle.Height
        sngWinHeight = sngWinHeight - picCard.Height - picTittle.Height
    End If
    
    '免挂号模式
     If gSystemPara.bln免挂号模式 And ( _
            (mEditType = Cr_补卡 Or mEditType = Cr_发卡 Or mEditType = Cr_退卡 Or chkCancel.value = 1) Or (mbln病历费 And (mEditType = Cr_绑定卡 Or mEditType = Cr_换卡)) _
            ) Then
        txtDate.Top = cbo支付方式.Top
        txtDate.Width = picCard.ScaleWidth - txtPass.Left - 60
        txtDate.Left = picCard.ScaleWidth - txtDate.Width - 60
        
        lblDate.Top = lbl卡费.Top
        lblDate.Left = txtDate.Left - lblDate.Width - 50
        
        txt操作员.Top = txtDate.Top
        txt操作员.Left = txt卡号.Width + txt卡号.Left - txt操作员.Width
        
        lbl发卡人.Top = lblDate.Top
        lbl发卡人.Left = txt操作员.Left - lbl发卡人.Width - 50
        lbl卡费.Left = txt卡费.Left - lbl卡费.Width - 20
    End If
    If mEditType = Cr_查询 Then
        txt操作员.Top = txt变动原因.Top: txtDate.Top = txt操作员.Top
        lbl发卡人.Top = lbl卡费.Top: lblDate.Top = lbl卡费.Top
        picCard.Height = picCard.Height - txt操作员.Height - 50
        '重新加载一次发卡页面
        wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Clear
        Set Item = wndTaskPanel.Groups(Idx_TP_PatiCard).Items.Add(Idx_TP_PatiCard, "", xtpTaskItemTypeControl)
        Set Item.Control = picCard: tkpGroup.Expanded = True
        wndTaskPanel.Reposition
        sngTaskHeight = sngTaskHeight - 50
        sngWinHeight = sngWinHeight - 50

    End If
    '问题号:56599

    wndTaskPanel.Height = sngTaskHeight
    Me.Height = sngWinHeight
 
    cmdHelp.Top = ScaleHeight - cmdHelp.Height - 100 - stbThis.Height
    
    '73063,冉俊明,2014-5-20
    vRectForm = zlControl.GetControlRect(Me.hWnd)
    vRect = zlControl.GetControlRect(fraCard.hWnd)
    '计算边框宽度
    sinW = (vRectForm.Right - vRectForm.Left - Me.ScaleWidth) / 2
    '标题栏高度
    sinH = vRectForm.Bottom - vRectForm.Top - Me.ScaleHeight - sinW
    '定位
    pic预交余额.Top = vRect.Top - vRectForm.Top - sinH - IIf(mEditType = Cr_退卡, 120, 0)
'    pic预交余额.Top = wndTaskPanel.Height - picCard.Height - pic预交余额.Height - IIf(mEditType = Cr_退卡, 80, 180)
End Sub


Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2011-06-21 13:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKindStr As String, blnVisible As Boolean
    Dim intKind As Integer, strKey As String
    
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Call InitPara: Call ClearData: Call InitData:  Call InitDicts
    Call InitInsurePara
    '74449,冉俊明,2014-6-25,“上次发卡类别”不存在或被停用时无法提取其它卡类别
    Call InitIDKind
    Call InitCardType
    Call Init病历费
    '74539,冉俊明,2014-6-27,在收费处发院内卡后，在病人变动记录插入的变动性质为11（绑定卡），应该为1（发卡）
    Call SetCardPayOrBound '设置当前卡的操作类型
    Call SetDefaultLen
    'IDKind.IDKindStr = GetIDKindStr(IDKind.IDKindStr)
    
    mlng缺省卡号长度 = IDKind.GetDefaultCardNoLen
    mintTabIndex卡号 = txt卡号.TabIndex: mintTabIndex刷卡卡号 = txt刷卡卡号.TabIndex
    Call GetRegInFor(g私有模块, Me.Name, "idkind", strKey)
    intKind = Val(strKey)
     If intKind > 0 And intKind <= IDKind.ListCount Then IDKind.IDKind = intKind
     
    '取缺省的刷卡方式
    '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    '第7位后,就只能用索引,不然取不到数
    mblnDefaultPassInputCardNo = IDKind.ShowPassText
    Call SetBrushCardObject
    '94941:李南春,2016/4/7,修改门诊号权限
    txt门诊号.Locked = InStr(";" & mstrPrivs & ";", ";允许修改门诊号;") <= 0
    '初始化地址控件
    txt家庭地址.MaxLength = glngMax家庭地址: txt户口地址.MaxLength = glngMax户口地址
    txt联系人地址.MaxLength = glngMax联系人地址
    If Not mblnStructAdress Then Exit Sub
    padd家庭地址.Visible = mblnStructAdress: padd户口地址.Visible = mblnStructAdress
    padd家庭地址.ShowTown = mblnShowTown: padd户口地址.ShowTown = mblnShowTown
    txt家庭地址.Visible = False: cmd家庭地址.Visible = False
    padd家庭地址.Top = txt家庭地址.Top: padd家庭地址.Left = txt家庭地址.Left
    txt户口地址.Visible = False: cmd户口地址.Visible = False
    padd户口地址.Top = txt户口地址.Top: padd户口地址.Left = txt户口地址.Left
    padd家庭地址.MaxLength = glngMax家庭地址: padd户口地址.MaxLength = glngMax户口地址
End Sub
Private Function SetBrushCardObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡对象
    '编制:刘兴洪
    '日期:2011-07-08 11:06:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjReadCard Is Nothing Then
        Set mobjReadCard = zlGetComponentObject(mlngCardTypeID, False)
    End If
    If mobjReadCard Is Nothing Then Exit Function
    'zlInitComponents(ByVal frmMain As Object, _
    '    ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '    ByVal cnOracle As ADODB.Connection, _
    '    Optional blnDeviceSet As Boolean = False, _
    '    Optional strExpand As String
    If Not mobjReadCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") Then
        Set mobjReadCard = Nothing: Exit Function
    End If
    SetBrushCardObject = True
End Function
Private Function InitCompoent(ByVal lngCardTypeID As Long, bln消费卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化指定部件
    '入参:lngCardTypeID-初始化卡类别ID
    '        bln消费卡-消费卡
    '出参:
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-09 23:50:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Object
    Set objCard = zlGetComponentObject(lngCardTypeID, bln消费卡)
    If objCard Is Nothing Then Exit Function
    'zlInitComponents(ByVal frmMain As Object, _
    '    ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '    ByVal cnOracle As ADODB.Connection, _
    '    Optional blnDeviceSet As Boolean = False, _
    '    Optional strExpand As String
    If objCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") Then
         Exit Function
    End If
    InitCompoent = True
End Function
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2011-07-05 10:14:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    IDKind.Font = lbl姓名.Font
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    mblnChange = False: cbo年龄单位.ListIndex = 0: mblnChange = True
    '加载有效的支付类别
    Call Load支付方式
    If mEditType = Cr_挂失 Then
        strSQL = "Select 编码,名称,简码,有效天数,缺省标志 From 医疗卡挂失方式"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With cbo挂失方式
            .Clear
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Int(Val(Nvl(rsTemp!有效天数)) * 100)
                If Val(Nvl(rsTemp!缺省标志)) = 1 Then
                    .ListIndex = .NewIndex
                End If
                rsTemp.MoveNext
            Loop
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Load支付方式(Optional ByVal blnDel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select Nvl(A.缺省标志,0) as 缺省,B.编码,B.名称,B.性质" & _
    "   From 结算方式应用 A,结算方式 B" & _
    "   Where A.结算方式=B.名称 And A.应用场合=[1]" & _
    "           And Nvl(B.性质,1) IN(1,2,8)  " & _
    "   Order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "就诊卡")
    Set mcolPayMode = New Collection
    '短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡|是否退现|是否全退;…
    If Not blnDel Then strPayType = GetAvailabilityCardType
    varData = Split(strPayType, ";")
    With cbo支付方式
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind And rsTemp!性质 <> 8 Then
                .AddItem Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                mcolPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0, 1, 0), "K" & j
                If rsTemp!缺省 = 1 Then
                    .ListIndex = .NewIndex
                End If
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    With cbo支付方式
        For i = 0 To UBound(varData)
            '问题号:116175，焦博，2017/12/8，将医疗卡的缴款方式控制调整为受结算方式管理和设备启用共同控制
            rsTemp.Filter = "名称 ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                If InStr(1, varData(i), "|") <> 0 Then
                    varTemp = Split(varData(i), "|")
                    mcolPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1)
                    .ItemData(cbo支付方式.NewIndex) = -1
                    j = j + 1
                End If
            End If
        Next
    End With
    If cbo支付方式.ListCount > 0 And cbo支付方式.ListIndex < 0 Then cbo支付方式.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetControlVisitble()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Visible属性
    '编制:刘兴洪
    '日期:2011-07-07 00:20:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    Dim blnTmp As Boolean
    
    If mEditType = Cr_调整病人信息 Then
        picTittle.Visible = False
        picCard.Visible = False: Exit Sub
    End If
    '问题号:56599
    cmdCreateCard.Visible = (mEditType = Cr_发卡 Or mEditType = Cr_绑定卡) And InStr(1, mstrPrivs, ";制卡;") > 0 And mCardType.bln是否制卡
    
    If mEditType <> Cr_发卡 And mEditType <> Cr_退卡 And Not (gbln收费发票 And (mEditType = Cr_补卡 Or mEditType = Cr_换卡)) Then picTittle.Visible = False
    
    blnTmp = mEditType = Cr_补卡 Or mEditType = Cr_发卡 Or mEditType = Cr_退卡 Or chkCancel.value = 1
    
    txt卡费.Visible = blnTmp:  lbl卡费.Visible = blnTmp
    
    
    blnVisible = (blnTmp Or (mbln病历费 And (mEditType = Cr_绑定卡 Or mEditType = Cr_换卡))) And gSystemPara.bln免挂号模式 = False
    
    
    cbo支付方式.Visible = blnVisible: chk记帐.Visible = blnVisible
    lbl支付方式.Visible = blnVisible: txt合计.Visible = blnVisible
    IDKindPayMode.Visible = blnVisible: txt余额.Visible = blnVisible
    
    
    '95809:李南春,2016/8/24,存在病历费的时候，要显示结算方式
    blnVisible = mbln病历费 And blnVisible And gSystemPara.bln免挂号模式 = False
    chk病历费.Visible = blnVisible: txt病历费.Visible = blnVisible
    
    
    blnVisible = (Mid(mCardType.str读卡性质, 3, 1) = 1 Or Mid(mCardType.str读卡性质, 4, 1) = 1) And (blnTmp Or mEditType = Cr_绑定卡 Or mEditType = Cr_换卡)
    If mCardType.blnOneCard Or mCardType.str卡名称 = "二代身份证" Then  '问题号:53408
        cmdReadCard.Visible = False '不包含一卡通
    Else
        blnVisible = blnVisible And Not mCardType.bln就诊卡
        cmdReadCard.Visible = blnVisible And Not mCardType.bln就诊卡
        lbl卡号.BorderStyle = IIf(mCardType.bln就诊卡 And mEditType <> Cr_退卡, 1, 0) '问题号 ：57962
    End If
    
    txt刷卡卡号.TabIndex = mintTabIndex刷卡卡号: txt卡号.TabIndex = mintTabIndex卡号
    '退卡的一些设置
    If (mEditType = Cr_退卡 Or chkCancel.value = 1) _
        And InStr(1, "123", mParaData.int退卡模式) > 0 Then
        
        '0-不进行刷卡;1-刷卡退卡;2-单据号后再验证刷卡;3-1和2的共用模式
        cmdReadCard.Left = txt刷卡卡号.Left + txt刷卡卡号.Width - cmdReadCard.Width
        lbl密码.Visible = False: lbl验证.Visible = False
        txtPass.Visible = False: txtAudi.Visible = False
        lbl刷卡验证.Visible = True: txt刷卡卡号.Visible = True
        lbl刷卡验证.BorderStyle = IIf(mCardType.bln就诊卡, 1, 0)
        
        'lbl刷卡验证.Caption = "刷卡验证"
    ElseIf mEditType = Cr_换卡 Then
        lbl刷卡验证.Visible = True: txt刷卡卡号.Visible = True
        lbl刷卡验证.Caption = "原卡号"
        txt刷卡卡号.TabIndex = mintTabIndex卡号: txt卡号.TabIndex = mintTabIndex刷卡卡号
        '50893
        lbl原卡密码.Visible = True: txt原卡密码.Visible = True: txt原卡密码.TabIndex = txt刷卡卡号.TabIndex + 1
        txt卡号.TabIndex = txt原卡密码.TabIndex + 1
    ElseIf mEditType = Cr_挂失 Then
        lbl密码.Visible = True: cbo挂失方式.Visible = True
        lbl密码.Caption = "挂失方式"
        lbl刷卡验证.Visible = True: txt刷卡卡号.Visible = True: txt卡号.Visible = False
        lbl刷卡验证.Caption = "挂失卡号"
        lbl卡号.Visible = False: txtPass.Visible = False: txtAudi.Visible = False
        lbl卡费.Visible = True: txt变动原因.Visible = True: lbl卡费.Caption = "挂失原因"
        txt变动原因.Tag = "挂失原因"
        lbl发卡人.Caption = "挂失人": lblDate.Caption = "挂失时间"
    Else
        cmdReadCard.Left = txt卡号.Left + txt卡号.Width
        lbl密码.Visible = True: lbl验证.Visible = True
        txtPass.Visible = True: txtAudi.Visible = True
        lbl刷卡验证.Visible = False: txt刷卡卡号.Visible = False
        If mEditType = Cr_查询 Then
            cmdOK.Visible = False: cmdCancel.Top = cmdOK.Top
            cmdCancel.Caption = "退出(&C)"
        End If
        
    End If

    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    '118959:李南春，2018/1/2，补卡和换卡都需要用IDkind
    If (mEditType = Cr_发卡 Or mEditType = Cr_绑定卡 Or mEditType = Cr_换卡 Or mEditType = Cr_补卡) And chkCancel.value = 0 Then
    
        IDKindPay.Visible = True: IDKindPay.Enabled = True
        lbl卡号.BorderStyle = 0
        lbl卡号.Left = IDKindPay.Left - lbl卡号.Width
        IDKindPay.Top = txt卡号.Top
        cmdReadCard.Visible = False: fraCard.BorderStyle = 0
        If (mEditType = Cr_补卡 Or mEditType = Cr_发卡) And gSystemPara.bln免挂号模式 Then lbl卡费.Caption = "卡费(划价)"
    Else
        IDKindPay.Visible = False: IDKindPay.Enabled = False
        lbl卡号.Left = txt卡号.Left - lbl卡号.Width - 60
        fraCard.BorderStyle = IIf(mEditType = Cr_发卡 Or mEditType = Cr_绑定卡, 0, 1)
    End If
    
    '问题号:73063
    pic预交余额.Visible = mEditType = Cr_退卡 Or chkCancel.value = 1
    

    If mEditType = Cr_退卡 Or chkCancel.value = 1 Then
        IDKindPayMode.Visible = False: txt余额.Visible = False
    End If
    '104726:李南春,2017/4/17,收费发票打印发卡票据
    txtFact.Visible = gbln收费发票 And mPrint.bytPrintType <> 0
    lblFact.Visible = gbln收费发票 And mPrint.bytPrintType <> 0
End Sub

Private Sub SetControlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enable属性
    '编制:刘兴洪
    '日期:2011-07-05 10:19:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim objCtl As Control
   Dim blnEdit As Boolean
   '问题号:56599
   If mEditType <> Cr_发卡 And mEditType <> Cr_绑定卡 Then
        cmdPicFile.Enabled = False: cmdPicCollect.Enabled = False: cmdPicClear.Enabled = False
   End If
   blnEdit = ((mEditType = Cr_发卡) Or (mEditType = Cr_绑定卡)) And chkCancel.value = 0
    
   blnEdit = blnEdit And mrsInfo Is Nothing
   For Each objCtl In Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("TextBox")  '文本
'                If objCtl Is txt Then
'                    MsgBox 1
'                End If
                If objCtl.Tag = "姓名" Then
                    objCtl.Enabled = (mEditType = Cr_发卡 Or mEditType = Cr_绑定卡 Or mEditType = Cr_补卡 Or mEditType = Cr_换卡 Or mEditType = Cr_调整病人信息 Or mEditType = Cr_挂失) And chkCancel.value = 0
                ElseIf InStr(1, ",现住址,户口地址,", "," & objCtl.Tag & ",") > 0 Then
                    objCtl.Enabled = (blnEdit Or mEditType = Cr_调整病人信息) And Not mblnStructAdress
                Else
                    objCtl.Enabled = (blnEdit Or mEditType = Cr_调整病人信息)
                End If
                If InStr(1, ",卡号,密码,验证,", "," & objCtl.Tag & ",") > 0 Then
                    objCtl.Enabled = (mEditType = Cr_发卡 Or mEditType = Cr_换卡 Or mEditType = Cr_绑定卡 Or mEditType = Cr_补卡) And chkCancel.value = 0
                End If
                If "卡费" = objCtl.Tag Then
                    objCtl.Enabled = (mEditType = Cr_发卡 Or mEditType = Cr_补卡) And chkCancel.value = 0
                    If mCardType.rs医疗卡费 Is Nothing Then
                        objCtl.Enabled = False
                    ElseIf mCardType.rs医疗卡费.State <> 1 Then
                        objCtl.Enabled = False
                    ElseIf mCardType.rs医疗卡费.RecordCount = 0 Then
                        objCtl.Enabled = False
                    End If
                ElseIf objCtl Is txt病历费 Then
                    '95809
                    objCtl.Enabled = mEditType <> Cr_查询 And mbln病历费 And chkCancel.value = 0
                ElseIf objCtl Is txt合计 Then
                    objCtl.Enabled = (mEditType = Cr_发卡 Or mEditType = Cr_补卡 Or chk病历费.value = 1) And chkCancel.value = 0
                End If
                If InStr(1, ",工作单位,单位电话,单位邮编,单位开户行,单位帐号,", "," & objCtl.Tag & ",") > 0 Then
                    objCtl.Enabled = (blnEdit Or mEditType = Cr_调整病人信息) And InStr(mstrPrivs, ";合约病人登记;") > 0
                End If
                If InStr(1, ",刷卡卡号,", "," & objCtl.Tag & ",") > 0 Then
                    objCtl.Enabled = mEditType = Cr_退卡 Or mEditType = Cr_换卡 Or chkCancel.value = 1 Or mEditType = Cr_挂失
                End If
                If InStr(1, ",变动原因,挂失原因,", "," & objCtl.Tag & ",") > 0 Then
                      '变动原因和挂失原因是一个控件txt变动原因.tag
                      objCtl.Enabled = mEditType = Cr_挂失
                End If
                '问题号:56599
                If objCtl Is txtOtherWaring Then
                    objCtl.Enabled = True
                End If
                objCtl.BackColor = IIf(objCtl.Enabled, &H80000005, &H8000000F)
        Case UCase("ComboBox")
                If Not objCtl Is cbo支付方式 Then
                    If objCtl Is cboNO Then
                        objCtl.Enabled = mEditType <> Cr_查询
                    ElseIf objCtl Is cbo挂失方式 Then
                        objCtl.Enabled = mEditType = Cr_挂失
                    Else
                        objCtl.Enabled = (blnEdit Or mEditType = Cr_调整病人信息)
                    End If
                    objCtl.BackColor = IIf(objCtl.Enabled, &H80000005, &H8000000F)
                Else
                    objCtl.Enabled = chk记帐.value = 0 And chk记帐.Visible = True
                    If mCardType.rs医疗卡费 Is Nothing Then
                        objCtl.Enabled = False
                    ElseIf mCardType.rs医疗卡费.State <> 1 Then
                        objCtl.Enabled = False
                    ElseIf mCardType.rs医疗卡费.RecordCount = 0 Then
                        objCtl.Enabled = False
                    End If
                End If
                '问题号:56599
                If objCtl Is cboBloodType Or objCtl Is cboBH Then
                    objCtl.Enabled = True
                End If
        Case UCase("MaskEdBox")
                objCtl.Enabled = (blnEdit Or mEditType = Cr_调整病人信息)
                objCtl.BackColor = IIf(objCtl.Enabled, &H80000005, &H8000000F)
        Case UCase("CommandButton")
            If InStr(1, ",出生地点,区域,工作单位,现住址,户口地址,联系人地址,", "," & objCtl.Tag & ",") > 0 Then
                objCtl.Visible = (blnEdit Or mEditType = Cr_调整病人信息)
                If objCtl.Tag = "现住址" Then objCtl.Visible = objCtl.Visible And Not mblnStructAdress
                If objCtl.Tag = "户口地址" Then objCtl.Visible = objCtl.Visible And Not mblnStructAdress
                If objCtl.Tag = "工作单位" Then
                    objCtl.Visible = InStr(mstrPrivs, ";合约病人登记;") > 0 And blnEdit
                End If
            End If
        Case UCase("CheckBox")
            If chkCancel Is objCtl Then
                objCtl.Enabled = (mEditType = Cr_发卡 Or mEditType = Cr_退卡)
            ElseIf chk记帐 Is objCtl Then
                objCtl.Enabled = (mEditType = Cr_发卡 Or mEditType = Cr_补卡) And chkCancel.value = 0
                If mCardType.rs医疗卡费 Is Nothing Then
                    objCtl.Enabled = False
                ElseIf mCardType.rs医疗卡费.State <> 1 Then
                    objCtl.Enabled = False
                ElseIf mCardType.rs医疗卡费.RecordCount = 0 Then
                    objCtl.Enabled = False
                End If
            Else
                '95809
                objCtl.Enabled = mEditType <> Cr_查询 And mbln病历费
            End If
        Case UCase("PatiAddress")
            objCtl.Enabled = (blnEdit Or mEditType = Cr_调整病人信息) And mblnStructAdress
            objCtl.ControlLock = Not objCtl.Enabled
        End Select
    Next
    txtDate.Enabled = False
    If mEditType = Cr_调整病人信息 Then
    
        '不能更改病人姓名 67184
        If Not mrsInfo Is Nothing Then
            mbln医嘱业务 = zlExistOperationData(Nvl(mrsInfo!病人ID), "")
        ElseIf mlng病人ID <> 0 Then
            mbln医嘱业务 = zlExistOperationData(mlng病人ID, "")
        End If
        blnEdit = Not mbln医嘱业务 And InStr(mstrPrivsPubPatient, ";基本信息调整;") > 0
        
        cbo性别.Enabled = blnEdit
        txt年龄.Enabled = blnEdit
        cbo年龄单位.Enabled = blnEdit
        txt出生日期.Enabled = blnEdit
        txt出生时间.Enabled = blnEdit
        cbo性别.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
        txt年龄.BackColor = cbo性别.BackColor
        cbo年龄单位.BackColor = cbo性别.BackColor
        txt出生日期.BackColor = cbo性别.BackColor
        txt出生时间.BackColor = cbo性别.BackColor
    End If
    
    '104726:李南春,2017/4/18,收费发票打印发卡票据
    blnEdit = (mEditType = Cr_补卡 Or mEditType = Cr_发卡 Or mEditType = Cr_换卡) And chkCancel.value = 0
    txtFact.Enabled = blnEdit
    txtFact.Locked = Not (InStr(1, mstrPrivs, ";修改票据号;") > 0 And gbln收费发票)
    txtFact.BackColor = IIf(txtFact.Enabled, &H80000005, &H8000000F)
    
    Call SetCardEditEnabled
    '80503:李南春,2015/1/23,输入项目参数控制
    Call InitControl
End Sub
Public Sub ClearData(Optional ByVal blnSave As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除数据
   '编制:刘兴洪
    '日期:2011-07-03 10:14:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtl As Control
    Set mrsInfo = Nothing
    For Each objCtl In Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("TextBox")  '文本
            objCtl.Text = ""
        Case UCase("ComboBox")
            objCtl.ListIndex = -1
        Case UCase("MaskEdBox")
            If InStr(1, ",出生日期,出生时间,", "," & objCtl.Tag & ",") > 0 Then
                 objCtl.Text = IIf(objCtl.Tag = "出生日期", "____-__-__", "__:__")
            End If
        Case UCase("Command")
        Case UCase("Image") '问题号:56599
            objCtl.Picture = Nothing
        Case UCase("VSFlexGrid") '问题号:56599
            objCtl.Rows = 1
            objCtl.Rows = 2
        Case UCase("Patiaddress")
            objCtl.value = ""
        End Select
    Next
    Call SetDefaultValue
    chk记帐.value = IIf(mParaData.bln记帐, 1, 0)
    If mEditType = Cr_退卡 Or chkCancel.value = 1 Then
        lbl支付方式.Caption = "退款"
    Else
        lbl支付方式.Caption = "缴款"
    End If
    mblnChange = False
    mstr年龄 = ""
    mstr年龄单位 = ""
    If gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    If blnSave Then Call setFact
End Sub
Private Sub SetDefaultValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省值
    '编制:刘兴洪
    '日期:2011-07-28 09:00:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call SetCboDefault(cbo性别)
    Call SetCboDefault(cbo费别)
    Call SetCboDefault(cbo医疗付款)
    Call SetCboDefault(cbo国籍)
    Call SetCboDefault(cbo民族)
    Call SetCboDefault(cbo学历)
    Call SetCboDefault(cbo婚姻状况)
    Call SetCboDefault(cbo职业)
    Call SetCboDefault(cbo身份)
    Call SetCboDefault(cbo联系人关系)
    Call SetCboDefault(cbo支付方式)
    Call SetCboDefault(cbo年龄单位)
    If cbo年龄单位.ListIndex < 0 And cbo年龄单位.ListCount > 0 Then cbo年龄单位.ListIndex = 0
    'Call SetCboDefault(cbo病人类型)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:MM")
    txt门诊号.Text = zlGet门诊号
    txt操作员.Text = UserInfo.姓名
    '问题号:56599
    Set mdic医疗卡属性 = Nothing
    mstr采集图片 = ""
    mlng图像操作 = 0
    '初始化地址信息
    Call zlLoadDefaultAddr(padd家庭地址)
    Call zlLoadDefaultAddr(padd户口地址)
End Sub

Private Sub AutoBrushSet(ByVal objIDKind As IDKindNew, blnAutoRefrsh As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动刷新设置
    '编制:刘兴洪
    '日期:2011-06-20 13:31:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(blnAutoRefrsh)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(blnAutoRefrsh)
    Call objIDKind.SetAutoReadCard(blnAutoRefrsh)
End Sub

Private Sub txtPatient_GotFocus()
    If Not txtPatient.Enabled Or txtPatient.Locked Then Exit Sub
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    Call AutoBrushSet(IDKind, txtPatient.Text = "")
    zlControl.TxtSelAll txtPatient
    If IsCardType(IDKind, "姓名") Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub
Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False
    Call zlCommFun.OpenIme(False)
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
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim strCardNo As String, blnNotMsg As Boolean
    Dim blnPass As Boolean
    On Error GoTo errH
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
'    If Not mrsInfo Is Nothing And mEditType = Cr_调整病人信息 And KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    If IsCardType(IDKind, "姓名") Then
        '105567:李南春,2017/5/25,卡号加密导致第一个汉字拼音不能触发输入法
        blnPass = txtPatient.PasswordChar <> ""
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, mblnDefaultPassInputCardNo)
        txtPatient.IMEMode = 0
        blnPass = txtPatient.PasswordChar = "" And blnPass
        If blnPass Then
            If txtPatient.SelLength = Len(txtPatient.Text) Then
                txtPatient.Text = ""
            End If
            SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
        End If
    ElseIf IsCardType(IDKind, "门诊号") Or IsCardType(IDKind, "住院号") Or IsCardType(IDKind, "手机号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then
        '不是刷卡和回车,则退出
        Exit Sub
    End If

    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
        txtPatient.SelStart = Len(txtPatient.Text)
    End If
    blnNotMsg = mEditType = Cr_发卡 Or mEditType = Cr_绑定卡
    
    KeyAscii = 0
    strCardNo = Trim(txtPatient.Text)
    If Not GetPatient(txtPatient.Text, blnCard, blnNotMsg) Then
        '调整病人基本信息时,姓名也可能被调整,所以不能清除界面信息
        If Not mrsInfo Is Nothing And mEditType = Cr_调整病人信息 Then
            If mrsInfo.State = 1 Then Exit Sub
        End If
        strCardNo = Trim(txtPatient.Text): Call ClearData
        '10214:李南春,2016/11/14，姓名信息缓存
        If IDKind.IDKind = IDKind.GetKindIndex("姓名") Or blnCard Then
            '传入被清空的病人姓名
            Call zlQueryEMPIPatiInfo(strCardNo)
            If Not blnCard And Trim(txtPatient.Text) <> "" Then strCardNo = Trim(txtPatient.Text)
        End If
        If blnCard Then
             If mEditType = Cr_发卡 Or mEditType = Cr_绑定卡 Then
                If IDKindDefaultKind = mlngCardTypeID Then
                    txt卡号.Text = strCardNo
                End If
             End If
            zlControl.TxtSelAll txtPatient
        Else
            txtPatient.Text = strCardNo: zlControl.TxtSelAll txtPatient
        End If
        Call SetControlEnable
        lbl医保号(1).Visible = True: txt验证医保号.Visible = True
        If mInsurePara.lng外挂式医保险类 = 0 Or Not (mEditType = Cr_发卡 Or mEditType = Cr_绑定卡) Then
            lbl医保号(1).Visible = False
            txt验证医保号.Visible = False
        End If
        
        If InStr(1, "+*-", Left(txtPatient.Text & " ", 1)) > 0 Then
            KeyAscii = 0
            DoEvents
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            zlControl.TxtSelAll txtPatient
            
            Exit Sub
        End If
        Call Led欢迎信息
        '76609,冉俊明,2014-8-14,焦点定位问题
        If IDKind.GetCurCard.接口序号 = 0 And Not blnCard Then zlCommFun.PressKey vbKeyTab
Exit Sub
    End If
    If mEditType = Cr_换卡 Or mEditType = Cr_挂失 Then
        If blnCard Then txt刷卡卡号.Text = strCardNo
    End If
    txtPatient.Text = Nvl(mrsInfo!姓名)
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0

    Call LoadPatiInfor: SetControlEnable: Call zlQueryEMPIPatiInfo
    lbl医保号(1).Visible = True: txt验证医保号.Visible = True
    If mInsurePara.lng外挂式医保险类 = 0 Or mEditType <> Cr_调整病人信息 Then
        lbl医保号(1).Visible = False
        txt验证医保号.Visible = False
    End If
    '76609,冉俊明,2014-8-14,焦点定位问题
'    If blnCard Then
        zlCommFun.PressKey vbKeyTab
'    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function LoadPatiInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-04 11:51:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str费别 As String, str其他关系 As String, strBirth As String
    On Error GoTo errHandle
    Call LoadCardFee
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    Call zlControl.CboLocate(cbo民族, Nvl(mrsInfo!民族))
    txt门诊号.Text = Nvl(mrsInfo!门诊号)
    mbln存在门诊号 = txt门诊号.Text <> ""
    If Not mbln存在门诊号 Then txt门诊号.Text = zlGet门诊号
    lbl门诊号.Tag = txt门诊号.Text
    txtPatient.Text = mrsInfo!姓名
    txt医保号.Text = Nvl(mrsInfo!医保号)
    '问题号:51071
    txt联系人身份证号.Text = Nvl(mrsInfo!联系人身份证号)
    If mEditType = Cr_调整病人信息 Then
        '外挂医保,或在院非真实医保病人可以修改医保号
        txt医保号.Enabled = mInsurePara.lng外挂式医保险类 > 0 Or Not IsNull(mrsInfo!住院次数) And IsNull(mrsInfo!险类)
        lbl医保号(0).Tag = txt医保号.Text
        If mInsurePara.lng外挂式医保险类 > 0 Then txt验证医保号.Text = txt医保号.Text
    End If
    
    
    Call zlControl.CboLocate(cbo性别, Nvl(mrsInfo!性别))
    If cbo性别.ListIndex = -1 And Not IsNull(mrsInfo!性别) Then
        cbo性别.AddItem mrsInfo!性别, 0
        cbo性别.ListIndex = cbo性别.NewIndex
    End If
    Call LoadOldData("" & mrsInfo!年龄, txt年龄, cbo年龄单位)
    mblnNotChange = True
    txt出生日期.Text = Format(IIf(IsNull(mrsInfo!出生日期), "____-__-__", mrsInfo!出生日期), "YYYY-MM-DD")
    If Not IsNull(mrsInfo!出生日期) Then
         'txt年龄.Text = ReCalcOld(CDate(txt出生日期.Text), cbo年龄单位, Val(Nvl(mrsInfo!病人ID)))   '修改的时候,根据出生日期重算年龄
         'If CDate(txt出生日期.Text) - CDate(mrsInfo!出生日期) <> 0 Then txt出生时间.Text = Format(mrsInfo!出生日期, "HH:MM")
     Else
        '103807:李南春，2016/12/20，年龄反算精确到小时
        If Not mobjPubPatient Is Nothing Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), strBirth) Then
                txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                txt出生时间.Text = Format(strBirth, "hh:mm")
            End If
        End If
     End If
    txt身份证号.Text = Nvl(mrsInfo!身份证号)
    mblnNotChange = False
    '根据不同查看方式读取不同的费别
    str费别 = Nvl(mrsInfo!费别)
    Call cbo.SeekIndex(cbo费别, str费别, , True)
    If cbo费别.ListIndex = -1 And str费别 <> "" Then
        cbo费别.AddItem str费别, 0
        cbo费别.ListIndex = cbo费别.NewIndex
    End If
    
        
    Call cbo.SeekIndex(cbo医疗付款, Nvl(mrsInfo!医疗付款方式), , True)
    If cbo医疗付款.ListIndex = -1 And Not IsNull(mrsInfo!医疗付款方式) Then
        cbo医疗付款.AddItem mrsInfo!医疗付款方式, 0
        cbo医疗付款.ListIndex = cbo医疗付款.NewIndex
    End If
       
   Call cbo.SeekIndex(cbo国籍, Nvl(mrsInfo!国籍), , True)
   If cbo国籍.ListIndex = -1 And Not IsNull(mrsInfo!国籍) Then
       cbo国籍.AddItem mrsInfo!国籍, 0
       cbo国籍.ListIndex = cbo国籍.NewIndex
   End If
   
   Call cbo.SeekIndex(cbo民族, Nvl(mrsInfo!民族), , True)
   If cbo民族.ListIndex = -1 And Not IsNull(mrsInfo!民族) Then
       cbo民族.AddItem mrsInfo!民族, 0
       cbo民族.ListIndex = cbo民族.NewIndex
   End If
   
   txt区域.Text = Nvl(mrsInfo!区域)
   
   Call cbo.SeekIndex(cbo学历, Nvl(mrsInfo!学历), , True)
   If cbo学历.ListIndex = -1 And Not IsNull(mrsInfo!学历) Then
       cbo学历.AddItem mrsInfo!学历, 0
       cbo学历.ListIndex = cbo学历.NewIndex
   End If
   
   Call cbo.SeekIndex(cbo婚姻状况, Nvl(mrsInfo!婚姻状况), , True)
   If cbo婚姻状况.ListIndex = -1 And Not IsNull(mrsInfo!婚姻状况) Then
       cbo婚姻状况.AddItem mrsInfo!婚姻状况, 0
       cbo婚姻状况.ListIndex = cbo婚姻状况.NewIndex
   End If
   
   Call cbo.SeekIndex(cbo职业, Nvl(mrsInfo!职业))
   If cbo职业.ListIndex = -1 And Not IsNull(mrsInfo!职业) Then
       cbo职业.AddItem mrsInfo!职业, 0
       cbo职业.ListIndex = cbo职业.NewIndex
   End If
   
   Call cbo.SeekIndex(cbo身份, Nvl(mrsInfo!身份), , True)
   If cbo身份.ListIndex = -1 And Not IsNull(mrsInfo!身份) Then
       cbo身份.AddItem mrsInfo!身份, 0
       cbo身份.ListIndex = cbo身份.NewIndex
   End If
        
   txt出生地点.Text = Nvl(mrsInfo!出生地点)
   txt家庭地址.Text = Nvl(mrsInfo!家庭地址)
   '89242:李南春,2015/12/10,读取病人地址信息
    Call zlReadAddrInfo(padd家庭地址, Val(Nvl(mrsInfo!病人ID)), 0, 3, txt家庭地址.Text)
   txt家庭电话.Text = Nvl(mrsInfo!家庭电话)
   txt手机.Text = Nvl(mrsInfo!手机号)
   txt家庭邮编.Text = Nvl(mrsInfo!家庭地址邮编)
   txt户口地址.Text = Nvl(mrsInfo!户口地址)
   Call zlReadAddrInfo(padd户口地址, Val(Nvl(mrsInfo!病人ID)), 0, 4, txt户口地址.Text)
   txt户口地址邮编.Text = Nvl(mrsInfo!户口地址邮编)
   txt联系人姓名.Text = Nvl(mrsInfo!联系人姓名)
   '84313,李南春,2015/4/27,联系人关系以及其他关系
    Call cbo.SeekIndex(cbo联系人关系, Nvl(mrsInfo!联系人关系), , True)
    If cbo联系人关系.ListIndex = -1 And Not IsNull(mrsInfo!联系人关系) Then
        cbo联系人关系.ListIndex = 8
        txt其他关系.Text = mrsInfo!联系人关系
    ElseIf cbo联系人关系.ListIndex = 8 Then
        str其他关系 = Get其他关系(Val(Nvl(mrsInfo!病人ID)))
        txt其他关系.Text = str其他关系
    End If
   
   txt联系人地址.Text = Nvl(mrsInfo!联系人地址)
   txt联系人电话.Text = Nvl(mrsInfo!联系人电话)
   txt工作单位.Text = Nvl(mrsInfo!工作单位)
   lbl工作单位.Tag = Nvl(mrsInfo!合同单位id)
   txt单位电话.Text = Nvl(mrsInfo!单位电话)
   txt单位邮编.Text = Nvl(mrsInfo!单位邮编)
   txt单位开户行.Text = Nvl(mrsInfo!单位开户行)
   txt单位帐户.Text = Nvl(mrsInfo!单位帐号)
   txt其他证件.Text = "" & mrsInfo!其他证件
   '问题号111659,焦博,2017/07/25,刷卡后清空了刷卡信息，退卡验卡失败
   '114252,李南春,2017/11/7，不清空密码信息
   'txt备注.Text = IIf(IsNull(mrsInfo!备注), "", mrsInfo!备注)
   Call Clear健康档案
   '问题号:56599
    Load健康卡相关信息 Nvl(mrsInfo!病人ID)
    '90875:李南春,2016/1/22,读取病人证件信息
    LoadCertificate Nvl(mrsInfo!病人ID)
    Call Led欢迎信息
    mblnChange = False
    LoadPatiInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim IDkindIndex As Integer
    Dim bln签约 As Boolean
    Dim strErrMsg As String
    Dim bln允许签约 As Boolean '是否允许签约,以身份证上信息与获取到的病人信息 是否一致判断
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        '先找病人
        mblnNotClick = True
        IDkindIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        bln签约 = 是否已经签约(strID)
        If mCardType.str卡名称 = "二代身份证" Then
            '发二代身份证
            If bln签约 Then
                MsgBox "此身份证已经签约,无需再次签约!", vbInformation, Me.Caption
                Set mrsInfo = Nothing
                Call txtPatient_GotFocus
                Exit Sub
            End If
        End If
        If GetPatient(strID, False, True) Then
            If Not mrsInfo Is Nothing Then
                If mCardType.str卡名称 = "二代身份证" Then
                    '检查身份证是否一直12-10-29 lgf
                    bln允许签约 = Not (Nvl(mrsInfo!姓名) <> Trim(strName) Or Nvl(mrsInfo!性别) <> strSex _
                                      Or Format(Nvl(mrsInfo!出生日期, "00-00-00"), "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd"))

                    If Not bln允许签约 Then
                         If Nvl(mrsInfo!姓名) <> Trim(strName) Then
                             strErrMsg = strErrMsg & "," & "姓名"
                        End If

                        If Nvl(mrsInfo!性别) <> strSex Then

                             strErrMsg = strErrMsg & "," & "性别"
                        End If

                        If Format(txt出生日期.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
                             strErrMsg = strErrMsg & "," & "出生日期"
                        End If

                        strErrMsg = Mid(strErrMsg, 2)
                        strErrMsg = "当前病人信息与身份证上的[" & strErrMsg & "]等信息不一致," & vbCrLf & "不能进行身份证签约!"
                        Call MsgBox(strErrMsg, vbInformation, Me.Caption)
                        Set mrsInfo = Nothing
                        Call txtPatient_GotFocus
                        Exit Sub
                    End If
                    txt卡号.Text = strID
                End If
                Call LoadPatiInfor: SetControlEnable: Call zlQueryEMPIPatiInfo
                '75717,冉俊明,2014-7-22,挂号预约时读取新病人身份证照片
                If imgPatient.Picture = 0 Then Call LoadIDImage
                txt户口地址.Text = IIf(Trim(txt户口地址.Text) = "", strAddress, txt户口地址.Text)
                txtPatient.PasswordChar = ""
            End If
        Else
            '新病人
             txtPatient.Text = strName
             txt身份证号.Text = strID
             Call zlControl.CboLocate(cbo性别, strSex)
             Call zlControl.CboLocate(cbo民族, strNation)
             txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
             '问题号:57817
             txt家庭地址.Text = IIf(Trim(txt家庭地址.Text) = "", strAddress, txt家庭地址.Text)
             txt户口地址.Text = strAddress
             '89242:李南春,2015/12/10,读取病人地址信息
             padd家庭地址.value = IIf(Trim(padd家庭地址.value) = "", strAddress, padd家庭地址.value)
             padd户口地址.value = strAddress
             
             If mCardType.str卡名称 = "二代身份证" Then
                txt卡号.Text = strID
             End If
             Call LoadIDImage: Call Led欢迎信息
             '新病人,姓名明文显示
             txtPatient.PasswordChar = ""
             Call zlQueryEMPIPatiInfo
        End If
        IDKind.IDKind = IDkindIndex
        mblnNotClick = False
        
         '问题号53408
        If mCardType.str卡名称 = "二代身份证" Then
            txt身份证号.PasswordChar = IIf(mCardType.str卡号密文 <> "", "*", "")
        Else
            txt身份证号.PasswordChar = ""
        End If
        
        '问题号:58072
        'Call SetControlEnable
        zlCommFun.PressKey vbKeyTab
    End If
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
End Sub
Private Sub txt原卡密码_Change()
'问题号:50893
    mblnChange = True
    Call SetCardEditEnabled
End Sub

Private Sub txt原卡密码_GotFocus()
'问题号:50893
    zlControl.TxtSelAll txt原卡密码
    zlCommFun.OpenIme False
End Sub

Private Sub txt原卡密码_KeyPress(KeyAscii As Integer)
'问题号:50893
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub
Private Function GetPatient(ByVal strInput As String, Optional ByVal blnCard As Boolean, Optional blnNotMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=表示是否就诊卡刷卡
    '出参:
    '返回:病人读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-03 10:46:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim vRect As RECT, rsTemp As ADODB.Recordset
    Dim strSQL As String, strPati As String, strWhere As String, blnHavePass As Boolean
    Dim lng病人ID As Long, blnCancel As Boolean, blnICCard As Boolean
    Dim strPassWord As String, bln存在帐户 As Boolean, strErrMsg As String
    Dim strCardNo As String, lng卡类别ID As Long, blnIsMobileNO As Boolean
    
    txtPatient.ForeColor = &HFF0000
    strErrMsg = ""
    blnIsMobileNO = IDKind.IsMobileNo(strInput)
    If IsCardType(IDKind, "IC卡号") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   '刷卡或缺省的卡
        
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
        Else
            lng卡类别ID = -1
        End If
        
        '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If GetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , , , mlngCardTypeID) = False Then
            If blnIsMobileNO Then
                '手机号查找
                If GetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , , , mlngCardTypeID) = False Then GoTo NotFoundPati:
            Else
                GoTo NotFoundPati:
            End If
        End If
        
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strCardNo = strInput
        strInput = "-" & lng病人ID
        blnHavePass = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then   '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strWhere = strWhere & " And A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strWhere = strWhere & " And A.病人ID = (Select Nvl(Max(病人ID),0) As 病人ID From 病案主页 Where 住院号 = [1])"
    ElseIf IsCardType(IDKind, "姓名") And blnIsMobileNO Then
        '手机号查找
        If GetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , , , mlngCardTypeID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strInput = "-" & lng病人ID
    Else
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                If mrsInfo!姓名 = strInput Then
                    '74309:李南春，2014-7-7，病人姓名显示颜色处理
                    Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), txtPatient.ForeColor)
                    GetPatient = True: Exit Function
                    End If
            End If
        End If
        Select Case IDKind.GetCurCard.名称
            Case "姓名", "姓名或就诊卡"
                '通过姓名模糊查找病人(允许输入病人标识时)
                If Not mParaData.blnSeekName Or mEditType = Cr_调整病人信息 Then
                    If Not mEditType = Cr_调整病人信息 Then
                        Set mrsInfo = New ADODB.Recordset
                    End If
                    Exit Function
                End If
                strPati = _
                " Select 1 As 排序id, 0 As ID, 0 As 病人id, '[新病人]' As 姓名, '' As 性别, '' As 年龄, 0 * Null As 门诊号, 0 * Null As 住院号," & vbNewLine & _
                "       To_Date(Null) As 出生日期, Null As 身份证号, Null As 家庭地址, Null As 工作单位, Null As 卡号" & vbNewLine & _
                " From Dual" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select /*+Rule */ 2 As 排序id, a.病人id As ID, a.病人id, Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.门诊号) As 门诊号," & vbNewLine & _
                "      Max(a.住院号) As 住院号, Max(a.出生日期) As 出生日期, Max(a.身份证号) As 身份证号, Max(a.家庭地址) As 家庭地址, Max(a.工作单位) As 工作单位, Max(b.卡号) As 卡号" & vbNewLine & _
                " From 病人信息 A, 病人医疗卡信息 B" & vbNewLine & _
                " Where a.停用时间 Is Null And a.病人id = b.病人id(+) And b.卡类别id(+) = 1 And Rownum < 101 And a.姓名 Like [1] " & vbNewLine & _
                IIf(mParaData.intNameDays = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])") & vbNewLine & _
                " Group By a.病人id"

                strPati = strPati & " Order by  排序ID,姓名, 卡号"
                
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTemp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人选择", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", mParaData.intNameDays)
                If blnCancel Then
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
                If rsTemp Is Nothing Then GoTo NotFoundPati:
                If rsTemp.State <> 1 Then GoTo NotFoundPati:
                If rsTemp.RecordCount = 0 Then GoTo NotFoundPati:
                If Val(Nvl(rsTemp!病人ID)) = 0 Then GoTo NotFoundPati:
                
                strInput = "-" & rsTemp!病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & "  And A.医保号=[2]"
             Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                '问题号:54197
                 If GetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , InStr(mstrPrivs, ";合并病人信息;") > 0, , , , , mlngCardTypeID) = False Then lng病人ID = 0
                 strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "联系人身份证号", "联系人身份证" '问题号:51071
                strInput = UCase(strInput)
                 If GetPatiID("联系人身份证", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , , , mlngCardTypeID) = False Then lng病人ID = 0
                 strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If GetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , , , mlngCardTypeID) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的号码
                If Val(IDKind.GetCurCard.接口序号) > 0 Then
                    lng卡类别ID = IDKind.GetCurCard.接口序号
                    bln存在帐户 = IDKind.GetCurCard.是否存在帐户
                    If GetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , , , mlngCardTypeID) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                    strCardNo = strInput
                    blnHavePass = True
                Else
                    If GetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , , , mlngCardTypeID) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
        End Select
    End If
    On Error GoTo errH
    '读取病人信息
   strSQL = "" & _
    "   Select Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊," & _
    "        A.病人id,A.门诊号,A.住院号,A.就诊卡号,A.卡验证码,A.费别,A.医疗付款方式," & _
    "        A.姓名,A.性别,A.年龄,A.出生日期,A.出生地点,A.身份证号," & _
    "        A.其他证件,A.身份,A.职业,A.民族,A.国籍,A.区域,A.学历,A.婚姻状况,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.监护人,A.联系人姓名," & _
    "        A.联系人关系,A.联系人地址,A.联系人电话,A.合同单位id,A.工作单位,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人," & _
    "        A.担保额,A.担保性质,A.就诊时间,A.就诊状态,A.就诊诊室,A.住院次数,A.当前科室id,A.当前病区id,A.当前床号,A.入院时间,A.出院时间," & _
    "        A.在院,A.Ic卡号,A.健康号,A.医保号,A.险类,A.查询密码,A.登记时间,A.停用时间,A.锁定,A.联系人身份证号,A.户口地址,A.户口地址邮编," & _
    "        M.编码 as 付款方式编码, decode(B1.病人性质,NULL,0,1,1,0) as 留观,B1.备注, " & _
    "        Nvl(Nvl(A.病人类型,B1.病人类型),Decode(Nvl(A.险类,B1.险类),Null,'普通病人','医保病人')) 病人类型,B1.入院日期, C.名称 险类名称," & _
    "        A.手机号" & _
    "   From 病人信息 A,病案主页 B1,保险类别 C ,医疗付款方式 M" & _
    "   Where A.险类 = C.序号(+) And A.医疗付款方式=M.名称(+) " & _
    "               And A.病人ID=B1.病人ID(+) And A.主页ID=B1.主页ID(+) " & _
    "               And A.停用时间 is NULL" & strWhere
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then GoTo NotFoundPati:
    If Not blnHavePass Then
        strPassWord = Nvl(mrsInfo!卡验证码)
    End If
    '74309:李南春，2014-7-7，病人姓名显示颜色处理
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), txtPatient.ForeColor)
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    If strErrMsg <> "" Then Exit Function
    
    If (IDKind.IDKind = IDKind.GetKindIndex("姓名") Or blnCard) And blnNotMsg Then
        txt门诊号.Text = zlGet门诊号
        Exit Function
    Else
        If blnCard Then
            MsgBox "不能确定病人信息，请检查是否正确刷卡！    ", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Else
            MsgBox "病人信息未找到,请检查是否输入正确!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        End If
    End If
End Function

Private Function zlInitMEPIPati(ByRef rsPati As ADODB.Recordset) As Boolean
    Set rsPati = New ADODB.Recordset
    With rsPati
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "病人ID", adBigInt, , adFldIsNullable
            .Append "主页ID", adBigInt, , adFldIsNullable
            .Append "挂号ID", adBigInt, , adFldIsNullable
            .Append "门诊号", adVarChar, 18, adFldIsNullable
            .Append "住院号", adVarChar, 18, adFldIsNullable
            .Append "医保号", adVarChar, 30, adFldIsNullable
            .Append "身份证号", adVarChar, 18, adFldIsNullable
            .Append "其他证件", adVarChar, 20, adFldIsNullable
            .Append "姓名", adVarChar, 100, adFldIsNullable
            .Append "性别", adVarChar, 4, adFldIsNullable
            .Append "出生日期", adVarChar, 20, adFldIsNullable
            .Append "出生地点", adVarChar, 100, adFldIsNullable
            .Append "国籍", adVarChar, 30, adFldIsNullable
            .Append "民族", adVarChar, 20, adFldIsNullable
            .Append "学历", adVarChar, 10, adFldIsNullable
            .Append "职业", adVarChar, 80, adFldIsNullable
            .Append "工作单位", adVarChar, 100, adFldIsNullable
            .Append "邮箱", adVarChar, 30, adFldIsNullable
            .Append "婚姻状况", adVarChar, 4, adFldIsNullable
            .Append "家庭电话", adVarChar, 20, adFldIsNullable
            .Append "联系人电话", adVarChar, 20, adFldIsNullable
            .Append "单位电话", adVarChar, 20, adFldIsNullable
            .Append "家庭地址", adVarChar, 100, adFldIsNullable
            .Append "家庭地址邮编", adVarChar, 6, adFldIsNullable
            .Append "户口地址", adVarChar, 100, adFldIsNullable
            .Append "户口地址邮编", adVarChar, 6, adFldIsNullable
            .Append "单位邮编", adVarChar, 6, adFldIsNullable
            .Append "联系人地址", adVarChar, 100, adFldIsNullable
            .Append "联系人关系", adVarChar, 30, adFldIsNullable
            .Append "联系人姓名", adVarChar, 64, adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    zlInitMEPIPati = True
End Function

Public Sub zlQueryEMPIPatiInfo(Optional ByVal strPatiName As String)
    '功能：从EMPI平台获取病人信息
    '日期：2016/10/9 10:47:13
    '编制：李南春
    '说明：101170
    Dim rsTmp As ADODB.Recordset, lng病人ID As Long, strDiff As String, strMsgInfo As String
    Dim strSQL As String
    If mblnPlugin = False Then Exit Sub
    If mobjPlugIn Is Nothing Then Exit Sub
    If mEditType <> Cr_发卡 And mEditType <> Cr_绑定卡 And mEditType <> Cr_调整病人信息 Or chkCancel.value = 1 Then Exit Sub
    
    On Error GoTo Errhand
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State = 0 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    '建档病人在发卡时不会调整个人信息
    If lng病人ID <> 0 And mEditType <> Cr_调整病人信息 Then Exit Sub
    If zlInitMEPIPati(rsTmp) = False Then Exit Sub
    
    With rsTmp
        .AddNew
        !病人ID = lng病人ID
        !门诊号 = txt门诊号.Text
        !医保号 = txt医保号.Text
        !身份证号 = txt身份证号.Text
        !姓名 = IIf(strPatiName = "", txtPatient.Text, strPatiName)
        !性别 = zlstr.NeedName(cbo性别.Text)
        If IsDate(txt出生日期) Then
            !出生日期 = Format(txt出生日期 & " " & IIf(IsDate(txt出生时间), txt出生时间, "00:00"), "YYYY-MM-DD HH:MM")
        Else
            !出生日期 = ""
        End If
        !出生地点 = txt出生地点.Text
        !国籍 = zlstr.NeedName(cbo国籍.Text)
        !民族 = zlstr.NeedName(cbo民族.Text)
        !职业 = zlstr.NeedName(cbo职业.Text)
        !工作单位 = txt工作单位.Text
        !家庭电话 = txt家庭电话.Text
        !联系人电话 = txt联系人电话.Text
        !单位电话 = txt单位电话.Text
        !家庭地址 = txt家庭地址.Text
        !家庭地址邮编 = txt家庭邮编.Text
        !户口地址 = txt户口地址.Text
        !户口地址邮编 = txt户口地址邮编.Text
        !单位邮编 = txt单位邮编.Text
        !联系人姓名 = txt联系人姓名.Text
        !联系人关系 = zlstr.NeedName(cbo联系人关系.Text)
        .Update
    End With
    'EMPI没有找到病人信息,直接返回
    Dim rsOut As New ADODB.Recordset
    On Error Resume Next
    If mobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsTmp, rsOut) = False Then
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: Set mrsEMPIOut = Nothing: Exit Sub
    End If
    Err.Clear: On Error GoTo Errhand
    Set mrsEMPIOut = rsOut
    If mrsEMPIOut Is Nothing Then Exit Sub
    If mrsEMPIOut.RecordCount = 0 Then Exit Sub
    mrsEMPIOut.MoveFirst
    On Error Resume Next
    With mrsEMPIOut
        '104905:李南春，2017/2/16，根据EMPI传回的病人ID，查找病人
        If lng病人ID <> Val(Nvl(!病人ID)) And Val(Nvl(!病人ID)) <> 0 Then
            strSQL = "" & _
            "   Select Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊," & _
            "        A.病人id,A.门诊号,A.住院号,A.就诊卡号,A.卡验证码,A.费别,A.医疗付款方式," & _
            "        A.姓名,A.性别,A.年龄,A.出生日期,A.出生地点,A.身份证号," & _
            "        A.其他证件,A.身份,A.职业,A.民族,A.国籍,A.区域,A.学历,A.婚姻状况,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.监护人,A.联系人姓名," & _
            "        A.联系人关系,A.联系人地址,A.联系人电话,A.合同单位id,A.工作单位,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人," & _
            "        A.担保额,A.担保性质,A.就诊时间,A.就诊状态,A.就诊诊室,A.住院次数,A.当前科室id,A.当前病区id,A.当前床号,A.入院时间,A.出院时间," & _
            "        A.在院,A.Ic卡号,A.健康号,A.医保号,A.险类,A.查询密码,A.登记时间,A.停用时间,A.锁定,A.联系人身份证号,A.户口地址,A.户口地址邮编," & _
            "        M.编码 as 付款方式编码, decode(B1.病人性质,NULL,0,1,1,0) as 留观,B1.备注, " & _
            "        Nvl(Nvl(A.病人类型,B1.病人类型),Decode(Nvl(A.险类,B1.险类),Null,'普通病人','医保病人')) 病人类型,B1.入院日期, C.名称 险类名称" & _
            "   From 病人信息 A,病案主页 B1,保险类别 C ,医疗付款方式 M" & _
            "   Where A.险类 = C.序号(+) And A.医疗付款方式=M.名称(+) " & _
            "               And A.病人ID=B1.病人ID(+) And A.主页ID=B1.主页ID(+) " & _
            "               And A.停用时间 is NULL And A.病人ID = [1]"
            Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(!病人ID)))
            If mrsInfo.EOF Then
                lng病人ID = 0: Call ClearData
            Else
                lng病人ID = Val(Nvl(!病人ID))
                Call LoadPatiInfor: SetControlEnable
                '如果不是调整基本信息，则退出更新
                If mEditType <> Cr_调整病人信息 Then Exit Sub
            End If
        End If
        
        If Nvl(!医保号) <> "" Then txt医保号.Text = Nvl(!医保号): txt验证医保号.Text = txt医保号.Text
        If Nvl(!身份证号) <> "" Then txt身份证号.Text = Nvl(!身份证号)
        If InStr(mstrPrivsPubPatient, ";基本信息调整;") > 0 Or lng病人ID = 0 Then
            If Nvl(!姓名) <> "" Then txtPatient.Text = Nvl(!姓名)
            If Nvl(!性别) <> "" Then cbo性别.ListIndex = cbo.FindIndex(cbo性别, Nvl(!性别), True)
            If Nvl(!出生日期) <> "" Then
                txt出生日期.Text = Format(Nvl(!出生日期), "YYYY-MM-DD")
                txt出生时间.Text = Format(Nvl(!出生日期), "HH:MM")
            End If
        Else
            If Nvl(!姓名) <> "" And txtPatient.Text <> Nvl(!姓名) Then strDiff = ",姓名"
            If Nvl(!性别) <> "" And cbo性别.ListIndex <> cbo.FindIndex(cbo性别, Nvl(!性别), True) Then strDiff = strDiff & ",性别"
            If Nvl(!出生日期) <> "" And Format(Nvl(!出生日期), "YYYY-MM-DD HH:MM:SS") <> Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS") Then strDiff = strDiff & ",出生日期"
        End If
        If Not txt门诊号.Locked And ExistClinicNO(Nvl(!门诊号), lng病人ID) = False Then
            If Nvl(!门诊号) <> "" Then txt门诊号.Text = Nvl(!门诊号): lbl门诊号.Tag = Nvl(!门诊号)
        Else
            If Nvl(!姓名) <> "" And txt门诊号.Text <> Nvl(!门诊号) Then strDiff = strDiff & ",门诊号"
        End If
        If Nvl(!出生地点) <> "" Then txt出生地点.Text = Nvl(!出生地点)
        If Nvl(!国籍) <> "" Then cbo国籍.ListIndex = cbo.FindIndex(cbo国籍, Nvl(!国籍), True)
        If Nvl(!民族) <> "" Then cbo民族.ListIndex = cbo.FindIndex(cbo民族, Nvl(!民族), True)
        If Nvl(!职业) <> "" Then cbo职业.ListIndex = cbo.FindIndex(cbo职业, Nvl(!职业))
        If Nvl(!工作单位) <> "" Then txt工作单位.Text = Nvl(!工作单位)
        If Nvl(!家庭电话) <> "" Then txt家庭电话.Text = Nvl(!家庭电话)
        If Nvl(!联系人电话) <> "" Then txt联系人电话.Text = Nvl(!联系人电话)
        If Nvl(!单位电话) <> "" Then txt单位电话.Text = Nvl(!单位电话)
        If Nvl(!家庭地址) <> "" Then txt家庭地址.Text = Nvl(!家庭地址): padd家庭地址.value = Nvl(!家庭地址)
        If Nvl(!家庭地址邮编) <> "" Then txt家庭邮编.Text = Nvl(!家庭地址邮编)
        If Nvl(!户口地址) <> "" Then txt户口地址.Text = Nvl(!户口地址): padd户口地址.value = Nvl(!户口地址)
        If Nvl(!户口地址邮编) <> "" Then txt户口地址邮编.Text = Nvl(!户口地址邮编)
        If Nvl(!单位邮编) <> "" Then txt单位邮编.Text = Nvl(!单位邮编)
        If Nvl(!联系人姓名) <> "" Then txt联系人姓名.Text = Nvl(!联系人姓名)
        If Nvl(!联系人关系) <> "" Then cbo联系人关系.ListIndex = cbo.FindIndex(cbo联系人关系, Nvl(!联系人关系), True)
    End With
    Err = 0: On Error GoTo 0
    If lng病人ID <> 0 Then
        If strDiff <> "" Then strDiff = Mid(strDiff, 2)
        If strDiff <> "" Then
            strMsgInfo = "病人的 " & strDiff & " 与EMPI信息不一致，因以下某种原因：" & vbNewLine & _
                        "     病人发了医嘱业务;" & vbNewLine & _
                        "     与其他病人信息冲突;" & vbNewLine & _
                        "     您不具有相应的权限。" & vbNewLine & _
                        "本次不会进行更新。 "
        End If
        If strMsgInfo <> "" Then MsgBox strMsgInfo, vbInformation, gstrSysName
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub InitDicts()
    mblnNotChange = True
    Call ReadDict("性别", cbo性别)
    Call ReadDict("费别", cbo费别)
    Call ReadDict("医疗付款方式", cbo医疗付款)
    Call ReadDict("国籍", cbo国籍)
    Call ReadDict("民族", cbo民族)
    Call ReadDict("学历", cbo学历)
    Call ReadDict("婚姻状况", cbo婚姻状况)
    Call ReadDict("职业", cbo职业, , mstrCboSplit)
    Call ReadDict("身份", cbo身份)
    Call ReadDict("社会关系", cbo联系人关系)
    mblnNotChange = False
End Sub

Private Function ReadDict(strDict As String, cboTemp As ComboBox, _
    Optional strClass As String, Optional strSplit As String = "-") As Boolean
'功能：初始化指定词典
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngMaxW As Long
    On Error GoTo errHandle
     If strDict = "结算方式" Then
        strSQL = "" & _
        "   Select Nvl(A.缺省标志,0) as 缺省,B.编码,B.名称,B.性质" & _
        "   From 结算方式应用 A,结算方式 B" & _
        "   Where A.结算方式=B.名称 And A.应用场合=[1]" & _
        "           And Nvl(B.性质,1) IN(1,2) Order by B.编码"
    ElseIf strDict = "身份" Then
        strSQL = "Select 编码,名称,简码,Nvl(优先级,0) as 缺省 From " & strDict & " Order by 编码"
    ElseIf strDict = "费别" Then
        strSQL = _
        "   Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 " & _
        "   From 费别" & _
        "   Where 属性=1 And Nvl(仅限初诊,0)=0 And Nvl(服务对象,3) IN(1,3)" & _
        "               And  Sysdate Between NVL(有效开始,Sysdate-1) and NVL(有效结束,Sysdate+1)" & _
        "   Order by 编码"
    ElseIf strDict = "病人类型" Then
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省,颜色 From 病人类型 Order by 编码"
    Else
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strClass)
    cboTemp.Clear
    If Not rsTemp.EOF Then
        For i = 1 To rsTemp.RecordCount
            cboTemp.AddItem rsTemp!编码 & strSplit & rsTemp!名称
            If rsTemp!缺省 = 1 Then
                cboTemp.ListIndex = cboTemp.NewIndex
                cboTemp.ItemData(cboTemp.NewIndex) = 1
            End If
            If strDict = "结算方式" And strClass = "预交款" Then
                   cboTemp.ItemData(cboTemp.NewIndex) = Val(Nvl(rsTemp!性质))
                   cboTemp.Tag = cboTemp.NewIndex   '单独保存为缺省的性质索引
            End If
            If TextWidth(cboTemp.List(cboTemp.NewIndex) & "兴洪") > lngMaxW Then lngMaxW = TextWidth(cboTemp.List(cboTemp.NewIndex) & "兴洪")
            rsTemp.MoveNext
        Next
        If strDict = "结算方式" And strClass <> "预交款" Then cboTemp.Tag = cboTemp.Text
        
    ElseIf strDict = "结算方式" Then
        If glngSys Like "8??" Then
            MsgBox "会员卡场合没有可用的结算方式，不能发卡！" & vbCrLf & _
                "请先到结算方式管理中设置会员卡的结算方式。", vbInformation, gstrSysName
        Else
            MsgBox "医疗卡场合没有可用的结算方式，只能使用记帐方式发卡！" & vbCrLf & _
                "要使用结算发卡,请先到结算方式管理中设置就诊卡结算方式。", vbInformation, gstrSysName
            chk记帐.value = 1: chk记帐.Enabled = False: cboTemp.Enabled = False
            chk记帐.Tag = 1
        End If
    End If
    ReadDict = True
    If cbo.ListWidth(cboTemp.hWnd) < lngMaxW Then zlControl.CboSetWidth cboTemp.hWnd, lngMaxW / Screen.TwipsPerPixelX
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub lbl卡号_Click()
    If mCardType.bln就诊卡 = False Then Exit Sub
    If mEditType = Cr_退卡 Then Exit Sub '问题号:57962
    
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    If mEditType = Cr_发卡 Or mEditType = Cr_绑定卡 Or mEditType = Cr_换卡 Or mEditType = Cr_补卡 Then Exit Sub
    
    If mobjICCard Is Nothing Then
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Set mobjICCard.gcnOracle = gcnOracle
    End If

    If Not mobjICCard Is Nothing Then
        txt卡号.Text = mobjICCard.Read_Card()
        If txt卡号.Text <> "" Then
            mblnICCard = True
            Call CheckFreeCard(txt卡号.Text)
        End If
    End If
End Sub
Private Sub CheckFreeCard(ByVal strCard As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对一卡通模式下的卡号，严格控制票号时， 检查是否在票据领用范围内，范围之外的卡不收费
    '编制:刘兴洪
    '日期:2011-07-05 08:53:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If txt卡费.Visible = False Then Exit Sub
    If Not mCardType.rs医疗卡费 Is Nothing And Val(txt卡费.Text) = 0 Then  '先恢复
        txt卡费.Text = Format(IIf(mCardType.bln变价, mCardType.rs医疗卡费!缺省价格, mCardType.rs医疗卡费!现价), "0.00")
        lbl卡费.Tag = txt卡费.Text
    End If
    If mCardType.blnOneCard And mCardType.lng共用批次 Then
        mCardType.lng领用ID = CheckUsedBill(5, IIf(mCardType.lng领用ID > 0, mCardType.lng领用ID, mCardType.lng共用批次), strCard)
        If mCardType.lng领用ID <= 0 Then txt卡费.Text = "0.00": lbl卡费.Tag = txt卡费.Text
    End If
    If Not mCardType.rs医疗卡费 Is Nothing And Val(txt卡费.Text) <> 0 Then
        If mCardType.bln变价 = False Then
            txt卡费.Text = Format(GetActualMoney(zlstr.NeedName(cbo费别.Text), mCardType.rs医疗卡费!收入项目ID, mCardType.rs医疗卡费!现价, mCardType.rs医疗卡费!收费细目ID), "0.00")
            lbl卡费.Tag = txt卡费.Text
        End If
    End If
End Sub
Private Function Select合约单位(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择合约单位
    '编制:刘兴洪
    '日期:2011-07-05 09:34:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, vRect As RECT, bytStyle As Byte
    Dim strWhere As String, strKey As String, blnCancel As Boolean
    
    bytStyle = 1: strWhere = "": strKey = GetMatchingSting(strInput)
    If strInput <> "" Then
        bytStyle = 0
        strWhere = " And 末级=1 and (编码 like upper([1]) or 名称 like [1] or 简码 like upper([1]) )"
    End If
    strSQL = "" & _
    "   Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人  " & _
    "   From  合约单位" & _
    "   Where (撤档时间 IS NULL OR TO_CHAR(撤档时间, 'yyyy-MM-dd') = '3000-01-01') " & _
        strWhere & _
    "       Start With 上级ID is NULL Connect by Prior ID=上级ID"
    vRect = zlControl.GetControlRect(txt工作单位.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, bytStyle, "合约单位选择", 1, "", "请选择病人的合约单位", False, False, True, vRect.Left, vRect.Top, txt工作单位.Height, blnCancel, False, True, strKey)
    If blnCancel Then
        If txt工作单位.Enabled And txt工作单位.Visible Then txt工作单位.SetFocus
        zlControl.TxtSelAll txt工作单位
        Set rsTemp = Nothing: Exit Function
    End If
    
    lbl工作单位.Tag = ""
    If Not rsTemp Is Nothing Then
        txt工作单位.Text = rsTemp!名称
        lbl工作单位.Tag = rsTemp!id
        txt单位电话.Text = Trim(rsTemp!电话 & "")
        txt单位开户行.Text = Trim(rsTemp!开户银行 & "")
        txt单位帐户.Text = Trim(rsTemp!帐号 & "")
    End If
    If txt工作单位.Enabled And txt工作单位.Visible Then txt工作单位.SetFocus
    zlCommFun.PressKey vbKeyTab
    Select合约单位 = True
End Function
Private Function Select区域(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择区域
    '编制:刘兴洪
    '日期:2011-07-05 09:34:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, vRect As RECT, bytStyle As Byte
    Dim strWhere As String, strKey As String, blnCancel As Boolean
    
    bytStyle = 0: strWhere = "": strKey = GetMatchingSting(strInput)
    If strInput <> "" Then
        strWhere = "  And  (编码 like upper([1]) or 名称 like [1] or 简码 like upper([1]))  "
    End If
    strSQL = "" & _
    "   Select 编码 as ID,编码,名称,简码 " & _
    "   From 区域" & _
    "   Where Nvl(级数,0)<3 " & strWhere
    vRect = zlControl.GetControlRect(txt区域.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, bytStyle, "区域选择", 1, "", "请选择病人的区域", False, False, True, vRect.Left, vRect.Top, txt区域.Height, blnCancel, False, True, strKey)
    If blnCancel Then
        If txt区域.Enabled And txt区域.Visible Then txt区域.SetFocus
        zlControl.TxtSelAll txt区域
        Set rsTemp = Nothing: Exit Function
    End If
    lbl区域.Tag = ""
    If Not rsTemp Is Nothing Then
        txt区域.Text = rsTemp!名称
        lbl区域.Tag = rsTemp!名称
    End If
    If txt区域.Enabled And txt区域.Visible Then txt区域.SetFocus
    zlCommFun.PressKey vbKeyTab
    Select区域 = True
End Function

Private Function Select地区(ByVal objCtrl As Control, ByVal objCtrlTag As Control, _
    ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择地区
    '编制:刘兴洪
    '日期:2011-07-05 09:34:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, vRect As RECT, bytStyle As Byte
    Dim strWhere As String, strKey As String, blnCancel As Boolean
    bytStyle = 0: strWhere = "": strKey = GetMatchingSting(strInput)
    
    If strInput <> "" Then
        strSQL = "" & _
        "   Select 编码 as ID,编码,名称,简码 " & _
        "   From 地区" & _
        "   Where     (编码 like upper([1]) or 名称 like [1] or 简码 like upper([1]) )"
    Else
        bytStyle = 2
        strSQL = "" & _
        "   Select Distinct Substr(名称,1,2) as ID,NULL as 上级ID,0 as 末级,NULL as 编码," & _
        "           Substr(名称,1,2) as 名称  " & _
        "   From 地区" & _
        "   Union All" & _
        "   Select 编码 as ID,Substr(名称,1,2) as 上级ID,1 as 末级,编码,名称 " & _
        "   From 地区  " & _
        "   Order by 编码"
    End If
    vRect = zlControl.GetControlRect(objCtrl.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, bytStyle, "出生地点选择", 1, "", "请选择病人的出生地点", False, False, True, vRect.Left, vRect.Top, objCtrl.Height, blnCancel, False, True, strKey)
    If blnCancel Then
        If objCtrl.Enabled And objCtrl.Visible Then objCtrl.SetFocus
        zlControl.TxtSelAll objCtrl
        Set rsTemp = Nothing: Exit Function
    End If
    objCtrlTag.Tag = ""
    If Not rsTemp Is Nothing Then
        objCtrl.Text = rsTemp!名称
        objCtrlTag.Tag = rsTemp!名称
    End If
    If objCtrl.Enabled And objCtrl.Visible Then objCtrl.SetFocus
    zlCommFun.PressKey vbKeyTab
    Select地区 = True
End Function
Private Sub LoadCardFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载卡费
    '编制:刘兴洪
    '日期:2011-07-06 17:24:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mCardType.rs医疗卡费 Is Nothing Then
        txt卡费.Text = ""
        GoTo Medical
    End If
    If mCardType.rs医疗卡费.RecordCount = 0 Then
        txt卡费.Text = ""
        GoTo Medical
    End If
    With mCardType.rs医疗卡费
        mCardType.bln变价 = Val(Nvl(!是否变价)) = 1
        mCardType.dbl应收金额 = Format(IIf(mCardType.bln变价, !缺省价格, !现价), "0.00")
        mCardType.dbl实收金额 = mCardType.dbl应收金额
        If mCardType.bln变价 = False And Nvl(!屏蔽费别, 0) <> 1 Then
            mCardType.dbl实收金额 = Format(GetActualMoney(zlstr.NeedName(cbo费别.Text), !收入项目ID, mCardType.dbl应收金额, !收费细目ID), "0.00")
        End If
        txt卡费.Locked = Not mCardType.bln变价
        txt卡费.TabStop = mCardType.bln变价
        If mCardType.bln变价 And Val(txt卡费.Text) = 0 Or Not mCardType.bln变价 Then
            txt卡费.Text = Format(mCardType.dbl实收金额, "0.00")
            Call txt余额_Change
        End If
    End With
    
Medical:
    '95809
    If Not mbln病历费 Then
        chk病历费.Enabled = False
        Exit Sub
    End If
    With mFeeType.rs病历费
        mFeeType.bln变价 = Val(Nvl(!是否变价)) = 1
        mFeeType.dbl应收金额 = Format(IIf(mFeeType.bln变价, !缺省价格, !现价), "0.00")
        mFeeType.dbl实收金额 = mFeeType.dbl应收金额
        If mFeeType.bln变价 = False And Nvl(!屏蔽费别, 0) <> 1 Then
            mFeeType.dbl实收金额 = Format(GetActualMoney(zlstr.NeedName(cbo费别.Text), !收入项目ID, mFeeType.dbl应收金额, !收费细目ID), "0.00")
        End If
        If mFeeType.bln变价 And Val(txt病历费.Text) = 0 Or Not mFeeType.bln变价 Then
            txt病历费.Text = Format(mFeeType.dbl实收金额, "0.00")
            Call txt余额_Change
        End If
        
        txt病历费.Locked = Not mFeeType.bln变价
        txt病历费.TabStop = mFeeType.bln变价
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetCardEditEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置卡的相关控件的Enable属性
    '编制:刘兴洪
    '日期:2011-07-07 00:12:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, blnEditFee As Boolean
    Select Case mEditType
    Case Cr_发卡, Cr_补卡, Cr_换卡, Cr_绑定卡
        blnEdit = Trim(txt卡号.Text) <> ""
        If chkCancel.value = 1 Then Exit Sub
    Case Else
        Exit Sub
    End Select
    txtPass.Enabled = blnEdit: txtAudi.Enabled = blnEdit
    lbl密码.Enabled = txtPass.Enabled: lbl验证.Enabled = blnEdit
    txtPass.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    txtAudi.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    '135260:李南春,2018/12/7,存在费用才允许选择支付方式
    If mEditType = Cr_补卡 Or mEditType = Cr_发卡 Then
        If mCardType.rs医疗卡费 Is Nothing Then
            blnEdit = False
        ElseIf mCardType.rs医疗卡费.State <> 1 Then
            blnEdit = False
        ElseIf mCardType.rs医疗卡费.RecordCount = 0 Then
            blnEdit = False
        End If
    Else
        blnEdit = False
    End If
    '只有发卡和补卡才存在卡费
    txt卡费.Enabled = blnEdit: cbo支付方式.Enabled = blnEdit And chk记帐.value = 0
    chk记帐.Enabled = blnEdit
    txt卡费.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    cbo支付方式.BackColor = IIf(cbo支付方式.Enabled, &H80000005, &H8000000F)
    txt合计.Enabled = blnEdit And chk记帐.value = 0
    txt合计.BackColor = IIf(txt合计.Enabled, &H80000005, &H8000000F)
    txt余额.Enabled = blnEdit And chk记帐.value = 0
    txt余额.BackColor = IIf(txt余额.Enabled, &H80000005, &H8000000F)
    
    If chk病历费.value = 0 Then Exit Sub
    If mbln病历费 Then
        blnEditFee = True
        If mFeeType.rs病历费 Is Nothing Then
            blnEditFee = False
        ElseIf mFeeType.rs病历费.State <> 1 Then
            blnEditFee = False
        ElseIf mFeeType.rs病历费.RecordCount = 0 Then
            blnEditFee = False
        End If
        chk病历费.Enabled = blnEditFee: blnEditFee = blnEditFee And chk病历费.value
        txt病历费.Enabled = blnEditFee
        cbo支付方式.Enabled = (blnEditFee Or blnEdit) And chk记帐.value = 0
        chk记帐.Enabled = (blnEditFee Or blnEdit)
        cbo支付方式.BackColor = IIf(cbo支付方式.Enabled, &H80000005, &H8000000F)
        txt合计.Enabled = (blnEditFee Or blnEdit) And chk记帐.value = 0
        txt合计.BackColor = IIf(txt合计.Enabled, &H80000005, &H8000000F)
        txt余额.Enabled = (blnEditFee Or blnEdit) And chk记帐.value = 0
        txt余额.BackColor = IIf(txt余额.Enabled, &H80000005, &H8000000F)
    End If
End Sub

Private Sub SearchCombox(cbo As ComboBox, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动索引指定的项目值
    '编制:刘兴洪
    '日期:2011-07-07 00:53:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngIdx As Long
    lngIdx = zlControl.CboMatchIndex(cbo.hWnd, KeyAscii)
    If lngIdx = -1 And cbo.ListCount > 0 Then lngIdx = 0
    cbo.ListIndex = lngIdx
End Sub
Private Function CheckExistsMCNO(ByVal strMCNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医保是否存在
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-07 03:08:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 1 From 病人信息 Where 医保号 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strMCNO)
    If rsTemp.RecordCount > 0 Then
        MsgBox "请检查,输入的医保号已存在!", vbInformation, gstrSysName
        CheckExistsMCNO = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlCheckMCOutMode(ByVal lng险类 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的险类是否外挂医保
    '入参:lng险类
    '返回:是外挂医保,返回True
    '编制:刘兴洪
    '日期:2011-07-07 02:35:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    strSQL = "Select 1 From 保险类别 Where 外挂=1 And 序号=[1]"
    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng险类)
    zlCheckMCOutMode = rsTemp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOldAcademic(ByVal dt出生日期 As Date, ByVal str年龄单位 As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前的出生日期和年龄单位，计算理论上的年龄值
    '返回:年龄
    '编制:刘兴洪
    '日期:2011-07-07 03:21:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim dtCurrDate As Date, lngOld As Long, strInterval As String
    If dt出生日期 = CDate(0) Or InStr(" 岁月天", str年龄单位) < 2 Then Exit Function
    dtCurrDate = zlDatabase.Currentdate
    strInterval = Switch(str年龄单位 = "岁", "yyyy", str年龄单位 = "月", "m", str年龄单位 = "天", "d")
    lngOld = DateDiff(strInterval, dt出生日期, dtCurrDate)
    If DateAdd(strInterval, lngOld, dt出生日期) > dtCurrDate Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function
Private Function SimilarIDs(str国籍 As String, str民族 As String, dat出生日期 As Date, str性别 As String, str姓名 As String, str身份证号 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人是否存在相似信息
    '入参:
    '出参:
    '返回:相似记录的病人ID串,如"234,235,236"
    '编制:刘兴洪
    '日期:2011-07-07 03:34:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String, i As Integer
    On Error GoTo errH
    strSQL = _
        " Select 病人ID,门诊号,住院号,Nvl(身份证号,'未登记') 身份证号,Nvl(家庭地址,'未登记') 地址,To_Char(登记时间,'YYYY-MM-DD') 登记时间 " & _
        " From 病人信息 Where (国籍=[1] And 民族=[2] And 性别=[3] And 姓名=[4]" & _
        " And 出生日期=[6]) Or 身份证号=[5] " & _
        " Order by 病人ID Desc"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", str国籍, str民族, str性别, str姓名, str身份证号, CDate(Format(dat出生日期, "YYYY-MM-DD")))
    For i = 1 To rsTemp.RecordCount
        SimilarIDs = SimilarIDs & "|ID:" & rsTemp!病人ID & ",门诊号:" & Nvl(rsTemp!门诊号, "无") & ",住院号:" & Nvl(rsTemp!住院号, "无") & ",身份证号:" & rsTemp!身份证号 & ",地址:" & rsTemp!地址 & ",登记日期:" & rsTemp!登记时间
        rsTemp.MoveNext
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ExistClinicNO(str门诊号 As String, Optional lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定门诊号是否已经存在于数据库中
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-07 03:40:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    On Error GoTo errH
    strSQL = "Select 病人ID,门诊号 From 病人信息 Where 门诊号=[1] And 病人ID<>[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查门诊号是否存在", Val(str门诊号), lng病人ID)
    If rsTemp.RecordCount > 0 Then ExistClinicNO = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function
Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        If Not mCardType.rs医疗卡费 Is Nothing Then
            .AddNew
            !收费类别 = mCardType.rs医疗卡费!收费类别
            !金额 = StrToNum(txt卡费.Text)
            .Update
        End If
        
        If Not mFeeType.rs病历费 Is Nothing Then
            .AddNew
            !收费类别 = mFeeType.rs病历费!收费类别
            !金额 = StrToNum(txt病历费.Text)
            .Update
        End If
        
        If Val(txt余额.Text) > 0 And IDKindPayMode.IDKind = 2 Then
            .AddNew
            !收费类别 = "预交"
            !金额 = StrToNum(txt余额.Text)
            .Update
        End If
    End With
    zlGetClassMoney = True
End Function

Private Function SetBrushObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡对象
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-10 13:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, bln消费卡 As Boolean, lngIndex As Long
    If mCurPayMoney.lng医疗卡类别ID = 0 Then SetBrushObject = True: Exit Function
    
    Set mobjCardObject = zlGetClsCardObject(mCurPayMoney.lng医疗卡类别ID, mCurPayMoney.bln消费卡)
    If mobjCardObject Is Nothing Then
        MsgBox "注意:" & vbCrLf & "   未找到相关的三方接口,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Not mobjCardObject.InitCompents Then
        If mobjCardObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") = False Then
              Exit Function
        End If
        mobjCardObject.InitCompents = True
    End If
    SetBrushObject = True
End Function
Private Function ReadCardNo(ByVal strCardNo As String, ByVal intFlag As Integer) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡验证就诊卡退卡姓名一致性及刷卡取数
    '入参:strCardNo-卡号
    '        intFlag 标志 1 验证 2 取数
    '出参:
    '返回:-1-成功;0-失败;1-该记录不存在
    '编制:刘兴洪
    '日期:2011-07-12 17:08:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim strOper As String, vDate As Date
    Dim lng病人ID As Long, str单据号 As String, strPassWord As String, strErrMsg As String
    Dim lng卡类别ID As String
    Dim blnNotShowMsg As Boolean
    
    Err = 0: On Error GoTo errH:
    ReadCardNo = 0
    If GetPatiID(mlngCardTypeID, strCardNo, False, lng病人ID, strPassWord, strErrMsg) = False Then
        If lng病人ID = 0 Then ReadCardNo = 1
        Exit Function
    End If
    If lng病人ID = 0 Then ReadCardNo = 1: Exit Function
    lbl刷卡验证.Tag = strCardNo
    If intFlag = 1 Then
        ReadCardNo = -1
        rsTmp.Close
        Exit Function
    End If
    If mEditType = Cr_换卡 Then
        If Not mrsInfo Is Nothing Then
            If Val(Nvl(mrsInfo!病人ID)) <> lng病人ID Then
                If GetPatient("-" & lng病人ID) = False Then
                    ReadCardNo = 1: Exit Function
                End If
            End If
        Else
            If GetPatient("-" & lng病人ID) = False Then
                ReadCardNo = 1: Exit Function
            End If
        End If
        Call LoadPatiInfor
        txt刷卡卡号.Text = strCardNo: lbl刷卡验证.Tag = strCardNo
        '问题号:50893
        txt原卡密码.Tag = strPassWord
        ReadCardNo = -1
        Exit Function
    End If
     If mEditType = Cr_挂失 Then
        txt刷卡卡号.Text = strCardNo: lbl刷卡验证.Tag = strCardNo
        ReadCardNo = -1
        Exit Function
     End If
     
    If mCardType.str卡名称 = "就诊卡" Then
        lng卡类别ID = mlngCardTypeID
    End If
    '获取就诊卡在费用中的No
    strSQL = _
    " Select A.NO" & _
    " From 住院费用记录 A" & _
    " Where A.记录性质=5   And A.实际票号=[1] " & _
    "           And A.病人ID = [2]  And A.记录状态=1 And nvl(A.结论,[3])=[4] "
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCardNo, lng病人ID, CStr(lng卡类别ID), CStr(mlngCardTypeID))
    If rsTmp.EOF Then ReadCardNo = 1: Exit Function
    str单据号 = IIf(IsNull(rsTmp!NO), "", rsTmp!NO)
    '读卡退卡验证操作权限
    If Not ReadBillInfo(2, str单据号, 5, strOper, vDate) Then
        ReadCardNo = 2: txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Function
    End If
    If Not BillOperCheck(8, strOper, vDate, "退卡") Then
        ReadCardNo = 2: txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Function
    End If
    
    If mEditType = Cr_退卡 Or chkCancel.value = 1 Then
        If mParaData.int退卡模式 = 2 And Trim(cboNO.Text) = "" Then
            MsgBox "注意:" & vbCrLf & "  退卡时,必须先输入单据,后刷卡!", vbInformation + vbOKOnly, gstrSysName
            
            Exit Function
        Else
            If str单据号 <> Trim(cboNO.Text) And Trim(cboNO.Text) <> "" Then
                MsgBox "当前刷卡的单据号与指定的单据号不符,不能退卡", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                If Nvl(mrsInfo!病人ID, 0) <> lng病人ID Then
                    MsgBox "当前病人所持有的卡不符,不能退卡", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If

    If ReadBill(str单据号, blnNotShowMsg) = -1 Then
        If blnNotShowMsg Then
            ReadCardNo = 2: txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Function
        Else
            ReadCardNo = -1
        End If
        rsTmp.Close
        Exit Function
    End If
    rsTmp.Close
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBill(strNO As String, blnNotShowMsg As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:由单据号读取并显示就诊卡发放记录
    '入参:strNO-单据号
    '出参:
    '返回:-1-成功;0-失败;1-该记录不存在;2-该记录已经作废(当mblnViewCancel=False时有效)
    '编制:刘兴洪
    '日期:2011-07-12 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, rsCheck As ADODB.Recordset
    Dim strSQL As String, str结算方式 As String, strFullNO As String, intIndex As Integer
    Dim byt消费卡 As Byte, lng卡类别ID As Long
    Dim str摘要 As String
    On Error GoTo errH
    cmdOK.Enabled = True
    strFullNO = GetFullNO(strNO, 16)
    '因为就诊卡费用的结帐ID可能是记帐发卡后结帐时产生的ID,
    '所以与预交记录联接时一定要加记录性质=5限制
    '问题号:50891
    '110414：李南春，2017/6/22，退卡左链接查询记录
    gstrSQL = _
        "Select a.No, a.病人id, a.姓名, a.性别, a.年龄, a.实际票号, a.附加标志, a.记录状态, a.结论, a.实收金额, a.操作员姓名," & vbNewLine & _
        "       a.发生时间, b.卡验证码, a.结帐id, a.摘要, c.结算方式, c.卡类别id, c.卡号, c.交易说明, c.结算序号, c.结算卡序号," & vbNewLine & _
        "       c.交易流水号, d.预交余额, e.收费票据, n.消费卡id" & vbNewLine & _
        "From 住院费用记录 A, 病人预交记录 C, 病人信息 B, 病人余额 D," & vbNewLine & _
        "     (Select m.号码 As 收费票据, n.No" & vbNewLine & _
        "       From 票据打印内容 N, 票据使用明细 M" & vbNewLine & _
        "       Where n.数据性质 = 5 And n.Id = m.打印id And m.性质 = 1 And m.票种 = 1 And" & vbNewLine & _
        "             m.使用时间 = (Select Max(M2.使用时间)" & vbNewLine & _
        "                       From 票据打印内容 N2, 票据使用明细 M2" & vbNewLine & _
        "                       Where M2.打印id = N2.Id And n.数据性质 = 5 And M2.票种 = 1 And N2.No = [1]) And n.No = [1]" & vbNewLine & _
        "       Order By m.使用时间 Desc) E, 病人卡结算记录 N" & vbNewLine & _
        "Where a.结帐id = c.结帐id(+) And c.记录性质(+) = 5 And a.病人id = d.病人id(+) And c.No(+) = [1] And a.记录性质 = 5" & vbNewLine & _
        "      And a.病人id = b.病人id And a.No = [1] And d.类型(+) = 1 And a.No = e.No(+) And c.Id = n.结算id(+)" & vbNewLine & _
               IIf(mEditType = Cr_查询, "And A.记录状态=[2] ", "")
    If mblnNOMoved Then
        gstrSQL = Replace(gstrSQL, "住院费用记录", "H住院费用记录")
        gstrSQL = Replace(gstrSQL, "病人预交记录", "H病人预交记录")
        gstrSQL = Replace(gstrSQL, "票据打印内容", "H票据打印内容")
        gstrSQL = Replace(gstrSQL, "票据使用明细", "H票据使用明细")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFullNO, mint记录状态)
    rsTemp.Filter = " 附加标志 <> 8 "
    If rsTemp.EOF Then ReadBill = 1: Exit Function
    
    If mEditType <> Cr_查询 And (rsTemp!记录状态 = 3 Or rsTemp!记录状态 = 2) Then
        ReadBill = 2: Exit Function
    End If
    '113613：李南春，2018/1/18，退卡时检查当前卡是否允许退卡
    If Nvl(rsTemp!实际票号) <> "" And (mEditType = Cr_退卡 Or chkCancel.value = 1) Then
        strSQL = "Select zl1_EX_ReFundCard_Check([1],[2],[3],[4]) as 验证 From Dual "
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "是否临时卡", mlngModule, Val(Nvl(rsTemp!病人ID)), Val(Nvl(rsTemp!结论)), Nvl(rsTemp!实际票号))
        If rsCheck.RecordCount > 0 Then
            If Nvl(rsCheck!验证) <> "" Then
                MsgBox Nvl(rsCheck!验证), vbOKOnly + vbInformation, gstrSysName
                blnNotShowMsg = True
                Exit Function
            End If
        End If
    End If
    
    Call GetPatient("-" & rsTemp!病人ID)
    Call LoadPatiInfor
    '问题号:73063
    lbl预交余额.Caption = "预交余额:" & Nvl(rsTemp!预交余额, "0") & "元"
    Call SetCtrlMove '重新布局当前界面控件
    
    txtFact.Text = Nvl(rsTemp!收费票据)
    cboNO.Text = rsTemp!NO
    cboNO.Tag = rsTemp!NO
    txtPatient.Text = rsTemp!姓名
    txtPatient.PasswordChar = ""
    str摘要 = Nvl(rsTemp!摘要)
    
    Call zlControl.CboLocate(cbo性别, Nvl(mrsInfo!性别))
    If cbo性别.ListIndex = -1 And Not IsNull(rsTemp!性别) Then
        cbo性别.AddItem mrsInfo!性别, 0
        cbo性别.ListIndex = cbo性别.NewIndex
    End If
    Call LoadOldData("" & rsTemp!年龄, txt年龄, cbo年龄单位)
    mlngBillCardTypeID = Val(Nvl(rsTemp!结论))
    Set mcolBillBalance = New Collection
    
    byt消费卡 = IIf(Val(Nvl(rsTemp!结算卡序号)) <> 0, 1, 0)
    lng卡类别ID = IIf(byt消费卡 = 1, Val(Nvl(rsTemp!结算卡序号)), Val(Nvl(rsTemp!卡类别id)))
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID,结算方式,消费卡ID
    mcolBillBalance.Add Array(lng卡类别ID, Trim(Nvl(rsTemp!卡号)), byt消费卡, Trim(Nvl(rsTemp!交易流水号)), _
        Trim(Nvl(rsTemp!交易说明)), strNO, Val(Nvl(rsTemp!结帐ID)), Nvl(rsTemp!结算方式), Val(Nvl(rsTemp!消费卡ID))), strNO
    
    'Call Load支付方式(True)
    If IsNull(rsTemp!结算方式) Then
        chk记帐.value = Checked
    Else
        '95809:李南春,2016/8/23,根据结算方式获取结算名称
        str结算方式 = zlGet支付方式(lng卡类别ID, rsTemp!结算方式)
    
        chk记帐.value = Unchecked
        Call cbo.SeekIndex(cbo支付方式, Split(str结算方式, "|")(0), , True)
        
        If cbo支付方式.ListIndex = -1 Then
            mcolPayMode.Add Array("", Split(str结算方式, "|")(0), 0, 0, 0, 0, Split(str结算方式, "|")(1), 0, 0, Split(str结算方式, "|")(2), Split(str结算方式, "|")(3))
            cbo支付方式.AddItem Split(str结算方式, "|")(0)
            cbo支付方式.ItemData(cbo支付方式.NewIndex) = Val(Split(str结算方式, "|")(4))
            cbo支付方式.ListIndex = cbo支付方式.NewIndex
            intIndex = cbo支付方式.NewIndex + 1
        Else
            intIndex = cbo支付方式.ListIndex + 1
        End If
        cbo支付方式.Tag = ""
    End If
    
    txt卡号.Text = IIf(IsNull(rsTemp!实际票号), "", rsTemp!实际票号)
    txtPass.Text = IIf(IsNull(rsTemp!卡验证码), "", rsTemp!卡验证码)
    txtAudi.Text = txtPass.Text
    txt卡费.Text = Format(rsTemp!实收金额, "0.00")
    txt操作员.Text = rsTemp!操作员姓名
    txtDate.Text = Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm")
    
    rsTemp.Filter = " 附加标志 = 8 "
    If mEditType = Cr_查询 Then
        If rsTemp.RecordCount > 0 Then
            stbThis.Panels(2).Text = "此张单据同时收取了病历费"
        End If
        ReadBill = -1: Exit Function
    End If
    If rsTemp.RecordCount > 0 Then
        chk病历费.Enabled = Val(Nvl(rsTemp!记录状态)) = 1
        txt病历费.Text = Format(rsTemp!实收金额, "0.00")
        '使用三方结算方式只能全部一起退
        If lng卡类别ID > 0 Then chk病历费.value = Checked: chk病历费.Enabled = False
    Else
        chk病历费.Enabled = False
    End If
    
    If intIndex > 0 Then
        cbo支付方式.Enabled = mcolPayMode.Item(intIndex)(9) = 1
        If cbo支付方式.ItemData(cbo支付方式.ListIndex) = 1 Then cbo支付方式.Enabled = False
    End If

    rsTemp.Filter = " 附加标志 <> 8 "
    '问题:48249
    If mEditType = Cr_退卡 Or chkCancel.value = 1 Then
        cbo支付方式.Enabled = False
        mlng病人ID = 0
        mlng病人ID = rsTemp!病人ID
        '116278:李南春,2017/12/15，不支持部分退的三方卡，退号必须同时退卡,暂时不管消费卡
        If str结算方式 <> "" And Nvl(rsTemp!卡类别id) <> 0 And Nvl(rsTemp!结算卡序号, 0) = 0 Then
            If Val(Split(str结算方式 & "||||", "|")(2)) = 0 Then
                strSQL = "Select 1 From 门诊费用记录 Where 记录性质=4 And 记录状态=1 And (病人ID,登记时间) = " & _
                        " (Select 病人ID,登记时间 From 住院费用记录 Where 记录性质=5 And NO=[1] And Rownum=1)"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", cboNO.Text)
                If Not rsTemp.EOF Then
                    MsgBox "当前卡是与挂号费一起收取的，请到挂号窗口与挂号费一起退。", vbInformation + vbOKOnly, gstrSysName
                    cmdOK.Enabled = False: blnNotShowMsg = True: Exit Function
                End If
            End If
        End If
        '如果存在摘要,需要读取门诊费用记录
        If str摘要 <> "" Then
            strSQL = "Select NO,记录状态 From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 where 病人ID=[1] and 记录性质=1 and 摘要=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, cboNO.Text)
            If rsTemp.RecordCount > 0 Then
                If Nvl(rsTemp!记录状态, 0) = 1 Then
                    MsgBox "当前卡已划价收费，请退卡后到收费窗口退费。", vbInformation + vbOKOnly, gstrSysName
                    cmdOK.Enabled = False: blnNotShowMsg = True: Exit Function
                End If
            End If
        End If
    End If
    
    txt合计.Text = Format(IIf(txt卡费.Visible, Val(txt卡费.Text), 0) + IIf(chk病历费.value, Val(txt病历费.Text), 0), "0.00")
    txt合计.Tag = txt合计.Text
    ReadBill = -1
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function zlGet门诊号() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取门诊号
    '返回:门诊号
    '编制:刘兴洪
    '日期:2011-07-28 08:39:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbln自动门诊号 Then zlGet门诊号 = zlDatabase.GetNextNo(3)
End Function
Private Function CheckBrushCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str年龄 As String
    Dim dblMoney As Double, cllBalance As Collection
    Dim blnErrCount As Boolean
    Dim frmInput As frmInputPass, lng病人ID As Long
    
    If Not (mEditType = Cr_发卡 Or mEditType = Cr_补卡) And chk病历费.value = 0 Then CheckBrushCard = True: Exit Function
    If SetBrushObject = False Then Exit Function
    
    On Error GoTo errHandle
    If mCurPayMoney.lng医疗卡类别ID = 0 Then CheckBrushCard = True: Exit Function
    dblMoney = IIf(IDKindPayMode.IDKind = 2, StrToNum(txt合计.Text), StrToNum(txt合计.Tag))
    Call zlGetClassMoney(rsMoney)
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text

    Set cllBalance = Nothing '57682
    
    '消费卡
    If mCurPayMoney.bln消费卡 And mobjCardObject.自制卡 Then
        Err = 0: On Error Resume Next
        If IsEmpty(cllBalance) Then   '57682
            Set cllBalance = Nothing
        End If
        blnErrCount = cllBalance.count
        If Err <> 0 Then
            Set cllBalance = Nothing
            Err = 0: On Error GoTo 0
        End If
        '功能:根据指定支付类别,弹出刷卡窗口
        '入参:rsClassMoney:收费类别,金额
        '        lngCardTypeID-为零时,为老一卡通刷卡
        '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
        '58322
        '115668:李南春,2017/10/25,新建实例调用消费卡支付
        If Not mrsInfo Is Nothing Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
        Set frmInput = New frmInputPass
        CheckBrushCard = frmInput.zlBrushPay(Me, mlngModule, mobjCardObject, rsMoney, _
                mCurPayMoney.lng医疗卡类别ID, True, txtPatient.Text, zlstr.NeedName(cbo性别.Text), str年龄, _
                dblMoney, mCurPayMoney.str刷卡卡号, mCurPayMoney.str刷卡密码, False, True, False, True, cllBalance, _
                Val(txt余额.Text) > 0 And IDKindPayMode.IDKind = 2, , "1", lng病人ID)
        Unload frmInput
        Set frmInput = Nothing
        Exit Function
    End If
    
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String) As Boolean
    If mobjCardObject.CardObject.zlBrushCard(Me, mlngModule, mCurPayMoney.lng医疗卡类别ID, _
            txtPatient.Text, zlstr.NeedName(cbo性别.Text), str年龄, dblMoney, mCurPayMoney.str刷卡卡号, mCurPayMoney.str刷卡密码) = False Then Exit Function
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If mobjCardObject.CardObject.zlPaymentCheck(Me, mlngModule, mCurPayMoney.lng医疗卡类别ID, _
         mCurPayMoney.str刷卡卡号, dblMoney, "", "") = False Then Exit Function
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlInterfacePrayMoney(ByRef cllPro As Collection, ByRef cllThreeSwap As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim dblMoney As Double
    If mCurPayMoney.lng医疗卡类别ID = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cbo支付方式.ItemData(cbo支付方式.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    If Val(txt合计.Tag) <= 0 Then zlInterfacePrayMoney = True: Exit Function
    
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dblMoney = IIf(IDKindPayMode.IDKind = 2, StrToNum(txt合计.Text), StrToNum(txt合计.Tag))
    If mobjCardObject.CardObject.zlPaymentMoney(Me, mlngModule, mCurPayMoney.lng医疗卡类别ID, mCurPayMoney.str刷卡卡号, mCurPayMoney.lng结帐ID, mCurPayMoney.strNO, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    
    '更新三交交易数据
    '问题号:58536
    If Not mCurPayMoney.bln消费卡 Then
        Call zlAddUpdateSwapSQL(False, mCurPayMoney.lng结帐ID, mCurPayMoney.lng医疗卡类别ID, mCurPayMoney.bln消费卡, mCurPayMoney.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllPro)
    End If
    Call zlAddThreeSwapSQLToCollection(False, mCurPayMoney.lng结帐ID, mCurPayMoney.lng医疗卡类别ID, mCurPayMoney.bln消费卡, mCurPayMoney.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    
    If IDKindPayMode.IDKind = 2 And Val(StrToNum(txt余额.Text)) > 0 Then
        Call zlAddUpdateSwapSQL(True, mlng预交ID, mCurPayMoney.lng医疗卡类别ID, mCurPayMoney.bln消费卡, mCurPayMoney.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllPro)
        Call zlAddThreeSwapSQLToCollection(True, mlng预交ID, mCurPayMoney.lng医疗卡类别ID, mCurPayMoney.bln消费卡, mCurPayMoney.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlShowSelectCardNo(Optional ByVal lng病人ID As Long = 0, _
    Optional str卡号 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择指定病人的卡号
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-16 17:12:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, vRect As RECT, rsTemp As ADODB.Recordset, blnCancel As Boolean
    strSQL = "" & _
    "   Select RowNum as ID, A.卡号, A.发卡日期, A.发卡人,B.姓名, B.年龄, B.身份证号, B.出生日期, B.手机号, B.家庭电话,B.家庭地址,B.联系人姓名,B.联系人关系 " & _
    "   From 病人医疗卡信息 A, 病人信息 B " & _
    "   Where A.病人id = B.病人id And A.卡类别id = [1] and A.病人ID=[2]  " & IIf(str卡号 = "", "", " And A.卡号=[3]") & _
    "   Order by A.卡号"
    
    vRect = zlControl.GetControlRect(txt刷卡卡号.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "选择需要挂失的卡号", 1, "", "选择需要挂失的卡号", False, False, True, vRect.Left, vRect.Top, txt刷卡卡号.Height, blnCancel, False, True, mlngCardTypeID, lng病人ID, str卡号)
    If blnCancel = True Then GoTo GoCancel:
    If rsTemp Is Nothing Then
        MsgBox "未找到满足条件的卡号或该病人未持有卡!", vbOKOnly + vbInformation, gstrSysName
        GoTo GoCancel:
        Exit Function
    End If
    If rsTemp.State <> 1 Then GoTo GoCancel:
    txt刷卡卡号.Text = Nvl(rsTemp!卡号)
    lbl刷卡验证.Tag = txt刷卡卡号.Text
    
    zlShowSelectCardNo = True
    Exit Function
GoCancel:
    txt刷卡卡号.Text = ""
    If txt刷卡卡号.Enabled And txt刷卡卡号.Visible Then txt刷卡卡号.SetFocus
    zlControl.TxtSelAll txt刷卡卡号
End Function

Private Function zl是否已绑定(str卡号 As String) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查卡号是否已经被绑定
    '入参:需要检查的卡号
    '返回:绑定的病人信息
    '编制:王吉
    '日期:2012-09-5 17:12:38
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo ErrHandl:
        strSQL = "" & _
        "   Select A.病人ID,A.姓名,A.身份证号 From 病人信息 A,病人医疗卡信息 B Where A.病人ID = B.病人ID And B.卡号 = [1]"
        Set zl是否已绑定 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str卡号)
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Dim strKindStr As String, strCardType As String
    Dim blnFindDefaultCard  As Boolean
    Dim lngCurCardTypeID As Long
    
    If gobjSquare Is Nothing Then Exit Function
    lngCurCardTypeID = mlngCardTypeID
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, IDKind.IDKindStr, txtPatient)
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    
    '96809
    IDKindPayMode.IDKindStr = "应收|应收|0|0|0|0|0|0|0|0|0;充值|充值|0|0|0|0|0|0|0|0|0"
    IDKindPayMode.IDKind = 1
    
    
    '72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
    '118959:李南春，2018/1/2，补卡和换卡都需要用IDkind
    If mEditType <> Cr_发卡 And mEditType <> Cr_绑定卡 And mEditType <> Cr_补卡 And mEditType <> Cr_换卡 Then Exit Function
'    IDKindPay.NotAutoAppendKind = True '不自动加入卡类别
    Call IDKindPay.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txt卡号)
    
    blnFindDefaultCard = GetValidKindStr(mlngCardTypeID)
    
    Select Case mEditType
        Case Cr_发卡
            strCardType = "发卡"
        Case Cr_绑定卡
            strCardType = "绑定卡"
        Case Cr_换卡
            strCardType = "换卡"
        Case Cr_补卡
            strCardType = "补卡"
    End Select
    If mblnFromCardMgr Then
        If blnFindDefaultCard = False Then
            MsgBox "该卡设备未启用，您不能进行" & strCardType & "操作，请到【参数设置>设备配置】中启用！", vbInformation, gstrSysName
            mblnUnLoad = True: Exit Function
        End If
    End If
    
    If IDKindPay.Cards.count = 0 Then
        MsgBox "没有可用于" & strCardType & "的有效医疗卡类别，请检查！", vbOKOnly + vbInformation, gstrSysName
        mblnUnLoad = True: Exit Function
    End If
    
    '定位缺省默认卡类别
    If blnFindDefaultCard Then
        If lngCurCardTypeID <> 0 Then
            IDKindPay.DefaultCardType = lngCurCardTypeID
            IDKindPay.IDKind = IDKindPay.GetKindIndex(IDKindPay.GetfaultCard.名称)
        End If
    Else
        mlngCardTypeID = IDKindPay.GetfaultCard.接口序号
        IDKindPay.DefaultCardType = mlngCardTypeID
        IDKindPay.IDKind = IDKindPay.GetKindIndex(IDKindPay.GetfaultCard.名称)
    End If
    '85565,李南春,2015/7/10:读卡性质
    txt卡号.Locked = Not (IDKindPay.GetCurCard.是否刷卡 Or IDKindPay.GetCurCard.是否扫描)
End Function
'获取默认IDKind索引
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

 
'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case "住院号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "住院号"
     Case "手机号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "手机号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.名称
            Else
                If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
            End If
     End Select
End Function
                
Private Function 是否已经签约(strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查需要绑定的卡号是否已经签约
    '入参:绑定卡号
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand:
    strSQL = "" & _
    "   Select Count(1) as 是否签约 From 病人医疗卡信息 Where 卡号=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡绑定", strCardNo)
    是否已经签约 = rsTemp!是否签约 > 0
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub InitvsDrug()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsDrug
    '初始化列表属性
     vsDrug.Editable = flexEDKbdMouse
    '设置列头
        SetColumHeader vsDrug, C_ColumHeader
    End With

End Sub

Private Sub SetColumHeader(objList As Object, strColumHeader As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置列头
    '参数:objList - 设置对象,strColumHeader - 列表设置字符串
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varSet As Variant
    Dim varColum As Variant
    Dim i As Long
    varSet = Split(strColumHeader, ";")
    If UBound(varSet) = 0 Then Exit Sub
    
    For i = LBound(varSet) To UBound(varSet)
        varColum = Split(varSet(i), ",")
        Select Case TypeName(objList)
            Case "VSFlexGrid"
                With objList
                    .Cols = UBound(varSet) + 1
                    .Cell(flexcpText, 0, i) = varColum(0)
                    .ColKey(i) = varColum(0)
                    .ColAlignment(i) = varColum(1)
                    .ColWidth(i) = varColum(2)
                    .ColHidden(i) = Not (varColum(3) = 1)
                End With
            Case Else
            '暂不处理
        End Select
    Next
End Sub
Private Sub vsDrug_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '问题号:56599
    If vsDrug.Col = 1 Then  '过敏反应列编辑时需判断是否字数超过了200
        With vsDrug
           If Len(.TextMatrix(vsDrug.Row, vsDrug.Col)) > 200 Then
                MsgBox "过敏反应输入字符超出最大字符数200,多出的字符将被自动截除！", vbInformation, gstrSysName
                .TextMatrix(.Row, .Col) = Mid(.TextMatrix(.Row, .Col), 1, 200)
           End If
        End With
    End If
End Sub

Private Sub vsDrug_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim rsTemp As Recordset
    Dim strFliter As String
    On Error GoTo ErrHandl:
    If KeyAscii <> 13 Then
        If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf vsDrug.Col = 0 Then
        KeyAscii = 0
        datCurr = zlDatabase.Currentdate
        strSQL = _
        " Select Distinct A.ID,A.编码," & _
        " A.名称,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.ID=C.诊疗项目ID" & _
        " And (C.名称 like [1] OR A.编码 like [1] OR C.简码 like [1])" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
        
        strFliter = gstrLike & UCase(vsDrug.EditText) & "%"
        '转移一次焦点(记录集只有一条时会自动返回，此时对单元格的赋值无效)
        cmdSelDrug.SetFocus
        '获取当前鼠标坐标值
        vRect = zlControl.GetControlRect(vsDrug.hWnd)
        vRect.Top = vRect.Top + (vsDrug.Row - 1) * 300 + 150
        vRect.Left = vRect.Left + 30
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "过敏药物", False, "过敏药物选择器", "请从下面的药品中选择一项作为病人过敏药物", False, False, True, vRect.Left, vRect.Top, 0, True, False, True, strFliter)
        vsDrug.SetFocus
        If Not rsTemp Is Nothing Then
            vsDrug.TextMatrix(vsDrug.Row, vsDrug.Col) = rsTemp!名称
            vsDrug.TextMatrix(vsDrug.Row, 2) = rsTemp!id
            If vsDrug.Rows - 1 = vsDrug.Row Then vsDrug.Rows = vsDrug.Rows + 1
        End If
    End If
    Exit Sub
ErrHandl:
    MsgBox Err.Description
End Sub

Private Sub vsDrug_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetCmdCtrlMove
End Sub
Private Sub vsDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:56599
    If KeyCode = 27 And vsDrug.Rows = 2 Then vsDrug.TextMatrix(1, 0) = "": vsDrug.Cell(flexcpData, 1, 0) = "": vsDrug.TextMatrix(1, 1) = ""
    If KeyCode = 27 And vsDrug.Rows > 2 Then vsDrug.Rows = vsDrug.Rows - 1 'Esc

End Sub

Private Sub vsDrug_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cmdSelDrug.Visible = False
End Sub

Private Sub vsDrug_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then SetCmdCtrlMove
End Sub

Private Sub vsDrug_KeyPress(KeyAscii As Integer)
    '78408:李南春,2014/10/9,光标跳转
    If KeyAscii = 13 Then
        If vsDrug.Col = 0 Then
             zlCommFun.PressKey vbKeyRight
        ElseIf vsDrug.Rows > vsDrug.Row + 1 Then
            vsDrug.Row = vsDrug.Row + 1
            vsDrug.Col = 0
        End If
    End If
End Sub

Private Sub vsDrug_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsDrug.Col = 0 Then SetCmdCtrlMove
    End If
End Sub
Private Sub InitVsInoculate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsInoculate
    '初始化列表属性
     vsInoculate.Editable = flexEDKbdMouse
    '设置列头
        SetColumHeader vsInoculate, C_InoculateHeader
    '设置选择按钮
        .ColDataType(0) = flexDTDate
        .ColEditMask(0) = "####-##-##"
        .ColDataType(2) = flexDTDate
        .ColEditMask(2) = "####-##-##"
    End With

End Sub
Private Sub VsInoculate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '问题号:56599
    If Col = 1 Or Col = 3 Then '接种名称列编辑时需判断是否字数超过了100
        With vsInoculate
           If Len(.TextMatrix(Row, Col)) > 100 Then
                MsgBox "接种名称输入字符超出最大字符数100,多出的字符将被自动截除！", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 100)
           End If
        End With
        If Col = 3 And vsInoculate.Rows - 1 = Row And vsInoculate.TextMatrix(Row, Col) <> "" Then
                vsInoculate.Rows = vsInoculate.Rows + 1
        End If
    Else
        With vsInoculate
           If IsDate(.TextMatrix(Row, Col)) = False And .TextMatrix(Row, Col) <> "    -  -  " Then
                MsgBox "输入的日期格式不对或不是正确的日期！", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = ""
           ElseIf .TextMatrix(Row, Col) = "    -  -  " Then
                .TextMatrix(Row, Col) = ""
           End If
        End With
    End If
End Sub
Private Sub VsInoculate_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:56599
    If KeyCode = 27 And vsInoculate.Rows = 2 Then
        If vsInoculate.TextMatrix(1, 2) <> "    -  -  " And vsInoculate.TextMatrix(1, 3) <> "" Then
            vsInoculate.TextMatrix(1, 2) = "": vsInoculate.TextMatrix(1, 3) = ""
        Else
            vsInoculate.TextMatrix(1, 0) = "": vsInoculate.TextMatrix(1, 1) = ""
        End If
    End If
    If KeyCode = 27 And vsInoculate.Rows > 2 Then 'Esc
        If vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) <> "    -  -  " And vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) <> "" Or vsInoculate.TextMatrix(vsInoculate.Rows - 1, 3) <> "" Then
            vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) = "": vsInoculate.TextMatrix(vsInoculate.Rows - 1, 3) = ""
        Else
            vsInoculate.Rows = vsInoculate.Rows - 1
        End If
    End If
End Sub

Private Sub vsInoculate_KeyPress(KeyAscii As Integer)
    '78408:李南春,2014/10/9,光标跳转
    If KeyAscii = 13 Then
        If vsInoculate.Col = 3 And vsInoculate.Rows - 1 = vsInoculate.Row Then
            zlCommFun.PressKey vbKeyTab
        ElseIf vsInoculate.Col = 3 Then
            vsInoculate.Col = 0: vsInoculate.Row = vsInoculate.Row + 1
            zlCommFun.PressKey vbKeyReturn
        Else
            zlCommFun.PressKey vbKeyRight
        End If
    End If
End Sub

Public Function InoculateValid() As Boolean
    '问题号56599
    Dim i As Long
    
    With vsInoculate
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" And .TextMatrix(i, 1) = "" Then
                MsgBox "接种名称必须输入！", vbInformation, gstrSysName
                .Select i, 1
                InoculateValid = False
                Exit Function
            End If
            If .TextMatrix(i, 0) = "" And .TextMatrix(i, 1) <> "" Then
                MsgBox "接种日期必须输入！", vbInformation, gstrSysName
                .Select i, 0
                InoculateValid = False
                Exit Function
            End If
            If .TextMatrix(i, 2) <> "" And .TextMatrix(i, 3) = "" Then
                MsgBox "接种名称必须输入！", vbInformation, gstrSysName
                .Select i, 3
                InoculateValid = False
                Exit Function
            End If
            If .TextMatrix(i, 2) = "" And .TextMatrix(i, 3) <> "" Then
                MsgBox "接种日期必须输入！", vbInformation, gstrSysName
                .Select i, 2
                InoculateValid = False
                Exit Function
            End If
        Next
    End With
    InoculateValid = True
End Function
Private Sub ComboBox(objcbo As ComboBox, strSet As String)
    Dim varTemp As Variant
    Dim i As Long
    varTemp = Split(strSet, ",")
    With objcbo
        For i = LBound(varTemp) To UBound(varTemp)
            .AddItem varTemp(i)
        Next
    End With
    If objcbo.ListCount <> 0 Then objcbo.ListIndex = 0
End Sub
Private Sub InitVsOtherInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsLinkMan
    '初始化列表属性
        .Editable = flexEDNone
    '设置列头
        SetColumHeader vsLinkMan, C_LinkManColumHeader
    End With
    With vsOtherInfo
         .Editable = flexEDNone
    '设置列头
        SetColumHeader vsOtherInfo, C_OtherInfoColumHeader
    End With
End Sub
Private Sub InitCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化ComBox控件
    '编制:56599
    '日期:2012-12-07 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '66743:刘尔旋,2013-11-25,血型与RH默认值的问题
    'ComboBox cboBloodType, C_血型
    zlComboxLoadFromSQL "Select 编码,名称,缺省标志 From 血型", cboBloodType
    ComboBox cboBH, C_BH
    If cboBH.ListCount <> 0 Then cboBH.ListIndex = -1
End Sub

Private Sub cmdMedicalWarning_Click()
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim strTemp As String
    
    vRect = zlControl.GetControlRect(txtMedicalWarning.hWnd)
    strSQL = "" & _
    "       Select 编码 as ID,名称,简码 From 医学警示 Where 名称 Not Like '其他%'"
    Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "医学警示", False, "", "", False, False, False, vRect.Left, vRect.Top - 180, 500, True, False, True)
    If Not rsTemp Is Nothing Then
      While rsTemp.EOF = False
        strTemp = strTemp & ";" & rsTemp!名称
        rsTemp.MoveNext
      Wend
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then txtMedicalWarning.Text = strTemp
End Sub
Private Sub SetDrugAllergy(str过敏药物 As String, str过敏反应 As String, Optional lng过敏ID = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置过敏药物
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str过敏药物 Then
                    .TextMatrix(i, 1) = str过敏反应
                    If lng过敏ID <> 0 Then .TextMatrix(i, 2) = lng过敏ID
                    Exit Sub
                End If

            Next
        End If
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str过敏药物
        .TextMatrix(.Rows - 1, 1) = str过敏反应
        If lng过敏ID <> 0 Then .TextMatrix(.Rows - 1, 2) = lng过敏ID
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetInoculate(str接种日期 As String, str接种名称 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置接种情况
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str接种名称 Then
                        .TextMatrix(i, j - 1) = str接种日期
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str接种日期
                .TextMatrix(.Rows - 1, j + 1) = str接种名称
                Exit Sub
            End If
        Next
        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        
    End With
End Sub
Private Sub SetLinkInfo(str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, str其他关系 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置联系人相关信息
    '编制:56599
    '日期:2012-12-12 09:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    With vsLinkMan
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str姓名 And .TextMatrix(i, 2) = str身份证号 Then
                    .TextMatrix(i, 1) = str关系: .TextMatrix(i, 3) = str电话
                    If i = 1 Then
                        txt联系人身份证号.Text = str身份证号
                        txt联系人姓名.Text = str姓名
                        Call cbo.SeekIndex(cbo联系人关系, str关系, , True)
                        If cbo联系人关系.ListIndex = -1 And str关系 <> "" Then
                            cbo联系人关系.ListIndex = 8: txt其他关系.Text = str关系
                        ElseIf cbo联系人关系.ListIndex = 8 Then
                            txt其他关系.Text = str其他关系
                        End If
                        txt联系人电话.Text = str电话
                    End If
                    Exit Sub
                End If
            Next
        End If
        
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str姓名
        If cbo.FindIndex(cbo联系人关系, str关系, True) = -1 And str关系 <> "" Then
            .TextMatrix(.Rows - 1, 1) = "其他": .TextMatrix(.Rows - 1, 4) = str关系
        Else
            .TextMatrix(.Rows - 1, 1) = str关系
            .TextMatrix(.Rows - 1, 4) = str其他关系
        End If
        .TextMatrix(.Rows - 1, 3) = str电话
        .TextMatrix(.Rows - 1, 2) = str身份证号
        If .Rows - 1 = 1 Then
            txt联系人身份证号.Text = str身份证号
            txt联系人姓名.Text = str姓名
            Call cbo.SeekIndex(cbo联系人关系, str关系, , True)
            If cbo联系人关系.ListIndex = -1 And str关系 <> "" Then
                cbo联系人关系.ListIndex = 8: txt其他关系.Text = str关系
            ElseIf cbo联系人关系.ListIndex = 8 Then
                txt其他关系.Text = str其他关系
            End If
            txt联系人电话.Text = str电话
        End If
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetOtherInfo(str信息名 As String, str信息值 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置其他情况
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str信息名 Then
                        .TextMatrix(i, j + 1) = str信息值
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str信息名
                .TextMatrix(.Rows - 1, j + 1) = str信息值
                Exit Sub
            End If
        Next
        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        
    End With
End Sub
Private Sub Load健康卡相关信息(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人健康卡信息
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-12 14:55:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs过敏药物 As Recordset
    Dim rs免疫记录 As Recordset
    Dim rsABO血型 As Recordset
    Dim rsRH As Recordset
    Dim rs医学警示 As Recordset
    Dim rs其他医学警示 As Recordset
    Dim rs病人信息 As Recordset
    Dim rs联系人 As Recordset
    Dim rs其他信息 As Recordset
    Dim str医学警示 As String
    Dim str联系人姓名 As String
    Dim str联系人关系 As String
    Dim str联系人电话 As String
    Dim str联系人身份证号 As String
    Dim str附加信息 As String
    Dim lng联系人数量 As Long
    Dim i As Long
    On Error GoTo ErrHandl:
    '读取照片
    ReadPatPricture lng病人ID
    
    If mEditType = Cr_绑定卡 Or mEditType = Cr_发卡 Then
        '获取过敏药物
        strSQL = "" & _
        "   Select 病人ID,过敏药物ID,过敏药物,过敏反应 From 病人过敏药物 Where 病人ID=[1]"
        Set rs过敏药物 = zlDatabase.OpenSQLRecord(strSQL, "病人过敏药物", lng病人ID)
        While rs过敏药物.EOF = False
            SetDrugAllergy Nvl(rs过敏药物!过敏药物), Nvl(rs过敏药物!过敏反应), Nvl(rs过敏药物!过敏药物ID, 0)
            rs过敏药物.MoveNext
        Wend
        '获取免疫记录
        strSQL = "" & _
        "   Select 病人ID,接种时间,接种名称 From 病人免疫记录 Where 病人ID=[1]"
        Set rs免疫记录 = zlDatabase.OpenSQLRecord(strSQL, "病人免疫记录", lng病人ID)
        While rs免疫记录.EOF = False
            SetInoculate Nvl(rs免疫记录!接种时间), Nvl(rs免疫记录!接种名称)
            rs免疫记录.MoveNext
        Wend
        '82072:李南春,2015/1/23,血型和RH取就诊ID 为null的记录
        '血型
        strSQL = "" & _
        "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='血型' And 就诊ID Is NULL "
        Set rsABO血型 = zlDatabase.OpenSQLRecord(strSQL, "ABO血型", lng病人ID)
        While rsABO血型.EOF = False
            For i = 0 To cboBloodType.ListCount - 1
                '76314,李南春，2014-08-06，病人信息正确获取
                If zlstr.NeedName(cboBloodType.List(i), ".") = zlstr.NeedName(Nvl(rsABO血型!信息值)) Then cboBloodType.ListIndex = i
            Next
            rsABO血型.MoveNext
        Wend
        'RH
        strSQL = "" & _
        "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='RH' And 就诊ID Is NULL "
        Set rsRH = zlDatabase.OpenSQLRecord(strSQL, "RH", lng病人ID)
        While rsRH.EOF = False
            For i = 0 To cboBH.ListCount - 1
                If cboBH.List(i) = Nvl(rsRH!信息值) Then cboBH.ListIndex = i
            Next
            rsRH.MoveNext
        Wend
        '医学警示
        strSQL = "" & _
        "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='医学警示' "
        Set rs医学警示 = zlDatabase.OpenSQLRecord(strSQL, "医学警示", lng病人ID)
        While rs医学警示.EOF = False
            str医学警示 = str医学警示 & ";" & Nvl(rs医学警示!信息值)
            rs医学警示.MoveNext
        Wend
        If str医学警示 <> "" Then str医学警示 = Mid(str医学警示, 2)
        txtMedicalWarning.Text = str医学警示
        '其他医学警示
        strSQL = "" & _
        "  Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='其他医学警示' "
        Set rs其他医学警示 = zlDatabase.OpenSQLRecord(strSQL, "其他医学警示", lng病人ID)
        While rs其他医学警示.EOF = False
            txtOtherWaring.Text = Nvl(rs其他医学警示!信息值)
            rs其他医学警示.MoveNext
        Wend
        '联系人相关信息
        '取病人信息表中的联系人信息
        '84313,李南春,2015/4/27,联系人关系以及其他关系
        strSQL = "" & _
        "   Select  A.联系人姓名,A.联系人关系,A.联系人电话,A.联系人身份证号,B.信息值 as 附加信息 From 病人信息 A,病人信息从表 B " & _
        "   Where   A.病人ID=B.病人ID(+) And A.病人ID=[1] And B.信息名(+)='联系人附加信息' And Not A.联系人姓名 is Null"
        Set rs病人信息 = zlDatabase.OpenSQLRecord(strSQL, "病人信息联系人信息", lng病人ID)
            If rs病人信息.EOF = False Then
            txt联系人身份证号.Text = Nvl(rs病人信息!联系人身份证号)
            txt联系人姓名.Text = Nvl(rs病人信息!联系人姓名)
            Call cbo.SeekIndex(cbo联系人关系, Nvl(rs病人信息!联系人关系), , True)
            If cbo联系人关系.ListIndex = -1 And Not IsNull(rs病人信息!联系人关系) Then
                cbo联系人关系.ListIndex = 8
                txt其他关系.Text = rs病人信息!联系人关系
            ElseIf cbo联系人关系.ListIndex = 8 Then
                txt其他关系.Text = Nvl(rs病人信息!附加信息)
            End If
            txt联系人电话.Text = Nvl(rs病人信息!联系人电话)
            
            SetLinkInfo Nvl(rs病人信息!联系人姓名), Nvl(rs病人信息!联系人关系), Nvl(rs病人信息!联系人电话), Nvl(rs病人信息!联系人身份证号), txt其他关系.Text
        End If
        '取病人信息从表中的联系人信息
        strSQL = "" & _
        "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 like '联系人%' order by 信息名 Asc"
        Set rs联系人 = zlDatabase.OpenSQLRecord(strSQL, "联系人相关信息", lng病人ID)
        If rs联系人.EOF = False Then
            rs联系人.Filter = "信息名 like '联系人姓名%'"
            lng联系人数量 = rs联系人.RecordCount
            rs联系人.Filter = ""
            For i = 2 To lng联系人数量 + 1
                While rs联系人.EOF = False
                    Select Case Nvl(rs联系人!信息名)
                        Case "联系人姓名" & i
                            str联系人姓名 = Nvl(rs联系人!信息值)
                        Case "联系人关系" & i
                            str联系人关系 = Nvl(rs联系人!信息值)
                        Case "联系人电话" & i
                            str联系人电话 = Nvl(rs联系人!信息值)
                        Case "联系人身份证号" & i
                            str联系人身份证号 = Nvl(rs联系人!信息值)
                        Case "联系人附加信息" & i
                            str附加信息 = Nvl(rs联系人!信息值)
                    End Select
                    rs联系人.MoveNext
                Wend
                SetLinkInfo str联系人姓名, str联系人关系, str联系人电话, str联系人身份证号, str附加信息
                rs联系人.MoveFirst
            Next
        End If
        '其他信息
        strSQL = "" & _
        "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 Not in ('血型','ABO','RH','医学警示','其他医学警示') And 信息名 Not like '联系人%'"
        Set rs其他信息 = zlDatabase.OpenSQLRecord(strSQL, "联系人其他信息", lng病人ID)
        '问题号:115886,焦博,2017/11/24,挂号提取该病人信息时，程序报错
        While rs其他信息.EOF = False
            If Nvl(rs其他信息!信息名) <> "" Then
                SetOtherInfo Nvl(rs其他信息!信息名), Nvl(rs其他信息!信息值)
            End If
            rs其他信息.MoveNext
        Wend
        '医疗卡属性
        Set mdic医疗卡属性 = Nothing
    End If
    
    Exit Sub
ErrHandl:
     If ErrCenter() = 1 Then Resume
End Sub

Private Sub Add健康卡相关信息(ByVal lng病人ID As Long, ByRef colPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:健康卡数据处理
    '入参:
    '编制:56599
    '日期:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim varKey As Variant
    '过敏药物
    With vsDrug
        If .Rows > 1 Then
            '清除该病人所有记录
            strSQL = " Zl_病人过敏药物_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSQL
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人过敏药物_Update("
                    '病人ID_In 病人过敏药物.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '过敏药物ID_In 病人过敏药物.过敏药物ID%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 2) = "", "", .TextMatrix(i, 2)) & "',"
                    '过敏药物_In  病人过敏药物.过敏药物%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '过敏反应_In 病人过敏反应.过敏反应%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '接种信息
    With vsInoculate
        If .Rows > 1 Then
            '清除该病人所有记录
            strSQL = " Zl_病人免疫记录_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSQL
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '其他信息
    'ABO血型
    '病人信息从表
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'血型',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & zlstr.NeedName(cboBloodType.Text, ".") & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'RH',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '医学警示
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & txtMedicalWarning.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '其他医学警示
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'其他医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & txtOtherWaring.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    '其他关系
    If txt联系人姓名.Text <> "" And txt其他关系.Visible Then
        strSQL = "Zl_病人信息从表_Update("
        '病人ID_In 病人信息从表.病人Id%Type
        strSQL = strSQL & "" & lng病人ID & ","
        '信息名_In 病人信息从表.信息名%Type
        strSQL = strSQL & "'联系人附加信息',"
        '信息值_In 病人信息从表.信息值%Type
        strSQL = strSQL & "'" & txt其他关系.Text & "',"
        '就诊Id_In 病人信息从表.就诊Id%Type
        strSQL = strSQL & "'')"
        zlAddArray colPro, strSQL
    End If
    
    '联系人相关信息
    With vsLinkMan
        If .Rows > 1 And .TextMatrix(1, 0) <> "" And Not mrsInfo Is Nothing Then
            SaveModifyPati
        End If
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '联系人姓名不能为空
                    For j = 0 To .Cols - 1
                        strSQL = "Zl_病人信息从表_Update("
                        '病人ID_In 病人信息从表.病人Id%Type
                        strSQL = strSQL & "" & lng病人ID & ","
                        '信息名_In 病人信息从表.信息名%Type
                        strSQL = strSQL & "'联系人" & .TextMatrix(0, j) & i & "',"
                        '信息值_In 病人信息从表.信息值%Type
                        strSQL = strSQL & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '就诊Id_In 病人信息从表.就诊Id%Type
                        strSQL = strSQL & "'')"
                        
                        zlAddArray colPro, strSQL
                    Next
                End If
            Next
        End If
    End With
    '其他信息
     With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 0) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 2) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
     End With
     '医疗卡属性
     If Not mdic医疗卡属性 Is Nothing Then
        For Each varKey In mdic医疗卡属性.Keys
            strSQL = "Zl_病人医疗卡属性_Update("
            strSQL = strSQL & lng病人ID & ","
            strSQL = strSQL & mlngCardTypeID & ","
            strSQL = strSQL & "'" & Trim(txt卡号.Text) & "',"
            strSQL = strSQL & "'" & varKey & "',"
            strSQL = strSQL & "'" & mdic医疗卡属性(varKey) & "')"
            zlAddArray colPro, strSQL
        Next
     End If
End Sub
Private Sub DeletePatPicture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-14 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo Errhand:
    strSQL = strSQL & "Zl_病人照片_Delete("
    strSQL = strSQL & lng病人ID & ")"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SavePatPicture(lng病人ID As Long, strFile As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New Recordset
    
        If strFile = "" Then Exit Sub
        If Sys.SaveLob(glngSys, 27, lng病人ID, strFile, 0) = False Then
            ShowMsgbox "保存照片有误,请确认文件是否被删除!"
            Exit Sub
        End If
End Sub

Private Function ReadPatPricture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-13 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTmp As String
    Dim rsData As Recordset
    
    '67776:刘尔旋,2013-11-20,提取无照片的病人信息，照片没有清除
    Set imgPatient.Picture = Nothing

    strTmp = Sys.ReadLob(glngSys, 27, lng病人ID)
    mstr采集图片 = strTmp
    imgPatient.Picture = LoadPicture(strTmp)
    If strTmp <> "" Then Kill strTmp
End Function

Private Function Get制卡XML(lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取制卡XML串(用于传给制卡对象已进行制卡操作)
    '入参:
    '编制:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    
    strXML = strXML & "<卡号>" & Trim(txt卡号.Text) & "</卡号>"  'Varchar2 20
    strXML = strXML & "<姓名>" & Trim(txtPatient.Text) & "</姓名>"  'Varchar2 100
    strXML = strXML & "<性别>" & zlstr.NeedName(cbo性别) & "</性别>"  'Varchar2 4
    strXML = strXML & "<年龄>" & txt年龄.Text & cbo年龄单位.Text & "</年龄>"  'Varchar2 10
    strXML = strXML & "<出生日期>" & Format(txt出生日期.Text & " " & txt出生时间.Text, "yyyy-mm-dd hh24:mi:ss") & "</出生日期>" 'Varchar2 20 yyyy-mm-dd hh24:mi:ss
    strXML = strXML & "<出生地点>" & Trim(txt出生地点.Text) & "</出生地点>"  'Varchar2 50
    strXML = strXML & "<身份证号>" & Trim(txt身份证号.Text) & "</身份证号>"  'Varchar2 18
    strXML = strXML & "<其他证件>" & Trim(txt其他证件.Text) & "</其他证件>" 'Varchar2 20
    strXML = strXML & "<职业>" & zlstr.NeedName(cbo职业.Text, mstrCboSplit) & "</职业>" 'Varchar2 80
    strXML = strXML & "<民族>" & zlstr.NeedName(cbo民族.Text) & "</民族>" 'Varchar2 20
    strXML = strXML & "<国籍>" & zlstr.NeedName(cbo国籍.Text) & "</国籍>" 'Varchar2 30
    strXML = strXML & "<学历>" & zlstr.NeedName(cbo学历.Text) & "</学历>" 'Varchar2 10
    strXML = strXML & "<婚姻状况>" & zlstr.NeedName(cbo婚姻状况.Text) & "</婚姻状况>" 'Varchar2 4
    strXML = strXML & "<区域>" & zlstr.NeedName(txt区域.Text) & "</区域>" 'Varchar2 30
    strXML = strXML & "<家庭地址>" & IIf(mblnStructAdress, Trim(padd家庭地址.value), Trim(txt家庭地址.Text)) & "</家庭地址>" 'Varchar2 50
    strXML = strXML & "<家庭电话>" & Trim(txt家庭电话.Text) & "</家庭电话>" 'Varchar2 20
    strXML = strXML & "<手机号>" & Trim(txt手机.Text) & "</手机号>" 'Varchar2 20
    strXML = strXML & "<户口邮编>" & Trim(txt户口地址邮编.Text) & "</户口邮编>" 'Varchar2 6
    strXML = strXML & "<监护人>" & "" & "</监护人>" 'Varchar2 64
    strXML = strXML & "<联系人姓名>" & Trim(txt联系人姓名.Text) & "</联系人姓名>" 'Varchar2 64
    strXML = strXML & "<联系人关系>" & zlstr.NeedName(cbo联系人关系.Text) & "</联系人关系>" 'Varchar2 30
    strXML = strXML & "<联系人附加信息>" & Trim(txt其他关系.Text) & "</联系人附加信息>" 'Varchar2 30
    strXML = strXML & "<联系人地址>" & Trim(txt联系人地址.Text) & "</联系人地址>" 'Varchar2 50
    strXML = strXML & "<联系人电话>" & Trim(txt联系人电话.Text) & "</联系人电话>" 'Varchar2 20
    strXML = strXML & "<工作单位>" & Trim(txt工作单位.Text) & "</工作单位>" 'Varchar2 100
    strXML = strXML & "<单位电话>" & Trim(txt单位电话.Text) & "</单位电话>" 'Varchar2 20
    strXML = strXML & "<单位邮编>" & Trim(txt单位邮编.Text) & "</单位邮编>" 'Varchar2 6
    strXML = strXML & "<单位开户行>" & Trim(txt单位开户行.Text) & "</单位开户行>" 'Varchar2 50
    strXML = strXML & "<单位帐号>" & Trim(txt单位帐户.Text) & "</单位帐号>" 'Varchar2 20
    strXML = strXML & "<病人ID>" & lng病人ID & "</病人ID>" 'Varchar2 18
    strXML = strXML & "<ABO血型>" & cboBloodType.Text & "</ABO血型>"  'Varchar2 10
    strXML = strXML & "<RH>" & cboBH.Text & "</RH>"  'Varchar2 10
    '医学警示
    strXML = strXML & "<哮喘标志>" & Get医学警示("哮喘标志") & "</哮喘标志>" 'Varchar2 2
    strXML = strXML & "<心脏病标志>" & Get医学警示("心脏病标志") & "</心脏病标志>" 'Varchar2 2
    strXML = strXML & "<心脑血管病标志>" & Get医学警示("心脑血管病标志") & "</心脑血管病标志>" 'Varchar2 2
    strXML = strXML & "<癫痫病标志>" & Get医学警示("癫痫病标志") & "</癫痫病标志>" 'Varchar2 2
    strXML = strXML & "<凝血紊乱标志>" & Get医学警示("凝血紊乱标志") & "</凝血紊乱标志>" 'Varchar2 2
    strXML = strXML & "<糖尿病标志>" & Get医学警示("糖尿病标志") & "</糖尿病标志>" 'Varchar2 2
    strXML = strXML & "<青光眼标志>" & Get医学警示("青光眼标志") & "</青光眼标志>" 'Varchar2 2
    strXML = strXML & "<透析标志>" & Get医学警示("透析标志") & "</透析标志>" 'Varchar2 2
    strXML = strXML & "<器官移植标志>" & Get医学警示("器官移植标志") & "</器官移植标志>" 'Varchar2 2
    strXML = strXML & "<器官缺失标志>" & Get医学警示("器官缺失标志") & "</器官缺失标志>" 'Varchar2 2
    strXML = strXML & "<可装卸义肢标志>" & Get医学警示("可装卸义肢标志") & "</可装卸义肢标志>" 'Varchar2 2
    strXML = strXML & "<心脏起搏器标志>" & Get医学警示("心脏起搏器标志") & "</心脏起搏器标志>" 'Varchar2 2
    '其他医学警示
    strXML = strXML & "<其他医学警示>" & Trim(txtOtherWaring.Text) & "</其他医学警示>" 'Varchar2 100
    '联系人相关信息
    strXML = strXML & GetLinkXML
    '健康档案编号
    strXML = strXML & "<健康档案编号>" & GetOtherInfo("健康档案编号") & "</健康档案编号>" 'Varchar2 18
    '新农合证号
    strXML = strXML & "<新农合证号>" & GetOtherInfo("新农合证号") & "</新农合证号>" 'Varchar2 18
    '其他证件
    strXML = strXML & Get其他证件
    '其他信息
    strXML = strXML & Get其他信息
    '过敏情况
    strXML = strXML & Get过敏药物
    '免疫记录
    strXML = strXML & Get免疫记录
    '医疗卡属性
    strXML = strXML & Get医疗卡属性
    
    Get制卡XML = strXML
End Function
Private Function Get医疗卡属性() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡属性XML
    '入参:
    '编制:56599
    '日期:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varKey As Variant
    Dim strXML As String
    strXML = "<医疗卡属性>"
    For Each varKey In mdic医疗卡属性
        strXML = strXML & "<信息名>" & varKey & "</信息名>"
        strXML = strXML & "<信息值>" & mdic医疗卡属性.Item(varKey) & "</信息值>"
    Next
    strXML = strXML & "</医疗卡属性>"
End Function
Private Function Get免疫记录() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取免疫记录XML
    '入参:
    '编制:56599
    '日期:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim i As Long
    
    strXML = "<免疫记录>"
    With vsInoculate
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                strXML = strXML & "<疫苗名称>" & .TextMatrix(i, 1) & "</疫苗名称>"
                strXML = strXML & "<接种时间>" & .TextMatrix(i, 0) & "</接种时间>"
            End If
            If .TextMatrix(i, 3) <> "" Then
                strXML = strXML & "<疫苗名称>" & .TextMatrix(i, 3) & "</疫苗名称>"
                strXML = strXML & "<接种时间>" & .TextMatrix(i, 2) & "</接种时间>"
            End If
        Next
    End With
    strXML = strXML & "</免疫记录>"
End Function
Private Function Get过敏药物() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取过敏药物
    '入参:
    '编制:56599
    '日期:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim i As Long
    
    strXML = "<过敏情况>"
    With vsDrug
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                strXML = strXML & "<药物名称>" & .TextMatrix(i, 0) & "</药物名称>"
                strXML = strXML & "<药物反应>" & .TextMatrix(i, 1) & "</药物反应>"
            End If
        Next
    End With
    strXML = strXML & "</过敏情况>"
    
    Get过敏药物 = strXML
End Function
Private Function Get其他信息() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取其他信息
    '入参:
    '编制:56599
    '日期:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim strSQL As String
    Dim rs证件类型 As Recordset
    Dim str证件类型 As String
    Dim i As Long
    On Error GoTo Errhand
    strSQL = "Select 名称 From 证件类型"
    Set rs证件类型 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs证件类型.EOF Then Get其他信息 = "": Exit Function
    While rs证件类型.EOF = False
        str证件类型 = str证件类型 & ";" & Nvl(rs证件类型!名称)
        rs证件类型.MoveNext
    Wend
    str证件类型 = str证件类型 & ";"
    strXML = "<其他信息>"
    With vsOtherInfo
        For i = 1 To .Rows - 1
            If InStr(str证件类型, ";" & .TextMatrix(i, 0) & ";") < 0 Then
                strXML = strXML & "<信息名>" & .TextMatrix(i, 0) & "</信息名>"
                strXML = strXML & "<信息值>" & .TextMatrix(i, 1) & "</信息值>"
            End If
            If InStr(str证件类型, ";" & .TextMatrix(i, 2) & ";") < 0 Then
                strXML = strXML & "<信息名>" & .TextMatrix(i, 2) & "</信息名>"
                strXML = strXML & "<信息值>" & .TextMatrix(i, 3) & "</信息值>"
            End If
        Next
    End With
    strXML = strXML & "</其他信息>"
    Get其他信息 = strXML
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function Get其他证件() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取其他证号
    '入参:
    '编制:56599
    '日期:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim strSQL As String
    Dim rs证件类型 As Recordset
    Dim str证件类型 As String
    Dim i As Long
    On Error GoTo Errhand
    strSQL = "Select 名称 From 证件类型"
    Set rs证件类型 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs证件类型.EOF Then Get其他证件 = "": Exit Function
    While rs证件类型.EOF = False
        str证件类型 = str证件类型 & ";" & Nvl(rs证件类型!名称)
        rs证件类型.MoveNext
    Wend
    str证件类型 = str证件类型 & ";"
    strXML = "<其他证件>"
    With vsOtherInfo
        For i = 1 To .Rows - 1
            If InStr(str证件类型, ";" & .TextMatrix(i, 0) & ";") > 0 Then
                strXML = strXML & "<信息名>" & .TextMatrix(i, 0) & "</信息名>"
                strXML = strXML & "<信息值>" & .TextMatrix(i, 1) & "</信息值>"
            End If
            If InStr(str证件类型, ";" & .TextMatrix(i, 2) & ";") > 0 Then
                strXML = strXML & "<信息名>" & .TextMatrix(i, 2) & "</信息名>"
                strXML = strXML & "<信息值>" & .TextMatrix(i, 3) & "</信息值>"
            End If
        Next
    End With
    strXML = strXML & "</其他证件>"
    Get其他证件 = strXML
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Function Get医学警示(str标志 As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医学警示
    '入参:str标志 - 需要查找的标志
    '编制:56599
    '日期:2012-12-18 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Get医学警示 = IIf(InStr(";" & txtMedicalWarning.Text & ";", Replace(str标志, "标志", "")) > 0, 1, 0)
End Function
Private Function GetLinkXML() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取联系人信息XML字符串
    '入参:
    '编制:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim i As Long

    strXML = "<联系信息>"
    With vsLinkMan
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then '联系人姓名不允许为空
                strXML = strXML & "<姓名>" & .TextMatrix(i, 0) & "</姓名>"
                strXML = strXML & "<关系>" & .TextMatrix(i, 1) & "</关系>"
                strXML = strXML & "<电话>" & .TextMatrix(i, 3) & "</电话>"
                strXML = strXML & "<身份证号>" & .TextMatrix(i, 2) & "</身份证号>"
            End If
        Next
    End With
    strXML = strXML & "</联系信息>"
    GetLinkXML = strXML
End Function
Private Function GetOtherInfo(str信息名 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定节点获取其他信息中指定的内容
    '入参:
    '编制:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strFind As String
    With vsOtherInfo
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = str信息名 Then
                strFind = .TextMatrix(i, 1)
            End If
            If .TextMatrix(i, 2) = str信息名 Then
                strFind = .TextMatrix(i, 3)
            End If
        Next
    End With
    GetOtherInfo = strFind
End Function

Private Function WriteCard(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写卡
    '入参:lng病人ID - 病人ID
    '编制:王吉
    '问题:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    If mobjReadCard Is Nothing Then
       If Not SetBrushCardObject() Then Exit Function
    End If
    If mobjReadCard Is Nothing Then Exit Function
    '84196:李南春,2015/4/22，接口不支持写卡的信息提示
    On Error Resume Next
    WriteCard = mobjReadCard.zlBandCardArfter(Me, mlngModule, mlngCardTypeID, lng病人ID, strExpend)
    'vb实时错误438对象不支持该属性或方法
    If Err = 438 Then
        MsgBox mCardType.str卡名称 & "接口不支持写卡", vbInformation, gstrSysName
        Err.Clear
    End If
    If Err = 0 Then Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function Check发卡性质(lng病人ID As Long, lng卡类别ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发卡时检查是否限制病人的发卡张数
    '入参:lng病人ID - 病人ID;lng卡类别ID  - 医疗卡的类别ID
    '编制:王吉
    '问题:57326
    '日期:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
        If Not (mEditType = Cr_绑定卡 Or mEditType = Cr_发卡) Or chkCancel = 1 Then Check发卡性质 = True: Exit Function
        strSQL = "Select Count(1) as 存在 From 病人医疗卡信息 Where 状态=0 And 病人ID=[1] And 卡类别ID=[2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng卡类别ID)
        If Val(Nvl(rsTemp!存在)) <= 0 Then Check发卡性质 = True: Exit Function
        Select Case mCardType.lng发卡性质
            Case 0 '不限制
                Check发卡性质 = True
            Case 1 '同一个病人只允许发一张卡
                If InStr(mstrPrivs, ";补卡;") > 0 Then
                    If MsgBox("该病人已经发(绑定)过" & mCardType.str卡名称 & ",不能再进行发(绑定)卡操作,但可以进行补卡操作,是否补卡?", vbQuestion + vbYesNo) = vbYes Then
                        Check发卡性质 = True
                        mEditType = Cr_补卡
                    End If
                Else
                    MsgBox "该病人已经发(绑定)过" & mCardType.str卡名称 & ",不能再进行发(绑定)卡操作!", vbInformation + vbOKOnly
                    Check发卡性质 = False
                End If
            Case 2 '同一个病人允许发多张卡,但需要提醒
               Check发卡性质 = MsgBox("该病人已经发(绑定)过" & mCardType.str卡名称 & ",是否要进行发(绑定)卡操作?", vbQuestion + vbYesNo) = vbYes
        End Select
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

'72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
Private Sub IDKindPay_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String
    
    If IsCardType(IDKindPay, "IC卡号") Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt卡号.Text = mobjICCard.Read_Card()
            If txt卡号.Text <> "" Then
                mblnICCard = True
                Call CheckFreeCard(txt卡号.Text)
                Call txt卡号_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If
    
    lng卡类别ID = IDKindPay.GetCurCard.接口序号
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
    'Call InitInterFacel(Me, mlngModule, lng卡类别ID, False, mobjCardObject)
    strExpand = lng卡类别ID
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txt卡号.Text = strOutCardNo
    If txt卡号.Text <> "" Then
        Call CheckFreeCard(txt卡号.Text)
        Call txt卡号_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub IDKindPay_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If wndTaskPanel.Groups.count = 0 Or IDKindPay.Enabled = False Then Exit Sub
    wndTaskPanel.Groups.Item(wndTaskPanel.Groups.count).Caption = objCard.名称
    mlngCardTypeID = objCard.接口序号
    '重新初始化卡类别和对应卡费
    Call InitCardType: Call LoadCardFee
    txt卡号.MaxLength = mCardType.lng卡号长度
    '85565,李南春,2015/7/10:读卡性质
    txt卡号.Locked = Not (objCard.是否刷卡 Or objCard.是否扫描)
    Call SetCardPayOrBound

     '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    Set gobjSquare.objCurCard = objCard
    
    mlng医疗卡长度 = objCard.卡号长度
    txt卡号.PasswordChar = IIf(IDKindPay.ShowPassText, "*", "")
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txt卡号.Text <> "" Then txt卡号.Text = ""
    If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txt卡号.IMEMode = 0
End Sub

Private Sub IDKindPay_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    
    If IDKindPay.Enabled = False Then Exit Sub
    If txt卡号.Visible = False Then Exit Sub
    '只能读选择类别的卡
    If mCardType.lng卡类别ID <> objCard.接口序号 Then Exit Sub
    
    txt卡号.Text = objPatiInfor.卡号
'    Call CheckAvailableCard(objPatiInfor)
    If txt卡号.Text <> "" Then
        Call ReLoadCardFee(True)
        Call CheckFreeCard(txt卡号.Text)
        Call FromKindLoadPati(objPatiInfor)
        Call zlQueryEMPIPatiInfo
    End If
    '76505,冉俊明,2014-8-11,发卡时读卡,界面焦点丢失
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

'72541,冉俊明,2014-5-7,收费处的发卡只能发放就诊卡的问题
Private Sub tbPageDo_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If chkCancel.Visible = True Then chkCancel.value = 0
    Select Case Item.Index
    Case 0
        mEditType = Cr_发卡
    Case 1
        mEditType = Cr_绑定卡
    End Select
    
    txt卡号.Text = ""
    txtPass.Text = ""
    txtAudi.Text = ""
    
    Call SetCardView
End Sub

Private Sub SetCardPayOrBound()
    '-------------------------------------------------------------------------------------
    '功能：在发卡与绑定卡之间切换时，设置当前操作类型
    '编制：冉俊明
    '日期：2014-5-7
    '问题号：72541
    '说明：
    '-------------------------------------------------------------------------------------
    Dim i As Integer
    Dim blnPay As Boolean, blnBound As Boolean
    Dim objItem As TabControlItem
    
    If mblnFromCardMgr Then mblnAddPage = False: tbPageDo.RemoveAll: Exit Sub    '发卡管理模块进入直接是默认操作
    '是否可发卡
    blnPay = zlstr.IsHavePrivs(mstrPrivs, "发卡") And (mCardType.bln自制卡 Or (Not mCardType.bln自制卡 And mCardType.bln是否发卡))
    '是否可绑定卡
    blnBound = zlstr.IsHavePrivs(mstrPrivs, "绑定卡") And (Not mCardType.bln自制卡 Or (mCardType.bln自制卡 And mCardType.bln是否重复使用))
    
    If blnPay And blnBound Then '当前卡类别可发卡，也可绑定卡
        If tbPageDo.ItemCount <> 0 Then tbPageDo.RemoveAll
        Set objItem = tbPageDo.InsertItem(0, "发卡", fraCard.hWnd, 0): objItem.Tag = Cr_发卡
        Set objItem = tbPageDo.InsertItem(1, "绑定卡", fraCard.hWnd, 0): objItem.Tag = Cr_绑定卡
        If mEditType = Cr_绑定卡 Then
            tbPageDo(1).Selected = True
        Else
            tbPageDo(1).Selected = True: tbPageDo(0).Selected = True
        End If
        mblnAddPage = True
    ElseIf blnPay And Not blnBound Then '当前卡类别仅可发卡
        mEditType = Cr_发卡
        mblnAddPage = False: tbPageDo.RemoveAll
    ElseIf Not blnPay And blnBound Then
        mEditType = Cr_绑定卡
        mblnAddPage = False: tbPageDo.RemoveAll
    End If
    Call SetCardView
End Sub

Private Sub SetCardView()
    '-------------------------------------------------------------------------------------
    '功能：在发卡与绑定卡之间切换时，调整界面显示
    '编制：冉俊明
    '日期：2014-5-7
    '问题号：72541
    '说明：
    '-------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    Dim sngTaskHeight As Single, sngWinHeight As Single
    
    '调整显示信息
    cmdCreateCard.Visible = (mEditType = Cr_发卡 Or mEditType = Cr_绑定卡) And InStr(1, mstrPrivs, ";制卡;") > 0 And mCardType.bln是否制卡

    blnVisible = mEditType = Cr_补卡 Or mEditType = Cr_发卡 Or mEditType = Cr_退卡 Or chkCancel.value = 1
    lbl卡费.Visible = blnVisible: txt卡费.Visible = blnVisible
    blnVisible = (blnVisible Or (mbln病历费 And (mEditType = Cr_绑定卡 Or mEditType = Cr_换卡))) = gSystemPara.bln免挂号模式 = False
    
    chk记帐.Visible = blnVisible
    lbl支付方式.Visible = blnVisible: cbo支付方式.Visible = blnVisible: txt合计.Visible = blnVisible
    '重新布局当前界面控件
    Call SetCtrlMove
End Sub

Private Function FromKindLoadPati(ByVal objPatiInfor As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载IDKind返回病人信息,读取病人信息
    '编制:冉俊明
    '日期:2014-05-08
    '问题号：72541,75848
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '问题号:56599
    Dim str过敏药物 As String, str过敏反应 As String '问题号:56599
    Dim str接种日期 As String, str接种名称 As String '问题号:56599
    Dim strABO血型 As String '问题号:56599
    Dim str信息名 As String, str信息值 As String '问题号:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '问题号:56599
    Dim str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, str地址 As String '问题号:56599
    Dim str其他关系 As String, strBirth As String
    On Error GoTo errHandle
    If Not (mEditType = Cr_绑定卡 Or mEditType = Cr_发卡) Then Exit Function
    If objPatiInfor Is Nothing Then Exit Function
    
    With objPatiInfor
        If .姓名 = "" Then Exit Function '如果姓名为空，则认为没有提取出数据
        Call ClearData
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
        txt卡号.Text = .卡号
        '    姓名    Varchar2    64
        txtPatient.Text = .姓名
        '    性别    Varchar2    4
        If .性别 <> "" Then
            Call zlControl.CboLocate(cbo性别, .性别)
            If cbo性别.ListIndex = -1 Then
                cbo性别.AddItem .性别
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
        End If
        '    年龄    Varchar2    10
        If .年龄 <> "" Then
            Call LoadOldData(.年龄, txt年龄, cbo年龄单位)
        End If
        '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
        txt出生日期.Text = Format(IIf(.出生日期 = "", "____-__-__", .出生日期), "YYYY-MM-DD")
        mblnNotChange = True
        If .出生日期 <> "" Then
             txt年龄.Text = ReCalcOld(CDate(txt出生日期.Text), cbo年龄单位)      '修改的时候,根据出生日期重算年龄
             If CDate(txt出生日期.Text) - CDate(.出生日期) <> 0 Then txt出生时间.Text = Format(.出生日期, "HH:MM")
         Else
            '103807:李南春，2016/12/20，年龄反算精确到小时
            If Not mobjPubPatient Is Nothing Then
                If mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), strBirth) Then
                    txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                    txt出生时间.Text = Format(strBirth, "hh:mm")
                End If
            End If
         End If
         mblnNotChange = False
        '    出生地点    Varchar2    50
        txt出生地点.Text = .出生地址
        '    身份证号    VARCHAR2    18
        txt身份证号.Text = .身份证号
        '    其他证件    Varchar2    20
        txt其他证件.Text = .其他证件
        '    职业    Varchar2    80
        Call cbo.SeekIndex(cbo职业, .职业)
        If cbo职业.ListIndex = -1 Then
            cbo职业.AddItem .职业, 0
            cbo职业.ListIndex = cbo职业.NewIndex
        End If
        '    民族    Varchar2    20
        Call cbo.SeekIndex(cbo民族, .民族, , True)
        If cbo民族.ListIndex = -1 And .民族 <> "" Then
            cbo民族.AddItem .民族, 0
            cbo民族.ListIndex = cbo民族.NewIndex
        End If
        '    国籍    Varchar2    30
        Call cbo.SeekIndex(cbo国籍, .国籍, , True)
        If cbo国籍.ListIndex = -1 And .国籍 <> "" Then
            cbo国籍.AddItem .国籍, 0
            cbo国籍.ListIndex = cbo国籍.NewIndex
        End If
        '    学历    Varchar2    10
        Call cbo.SeekIndex(cbo学历, .学历, , True)
        If cbo学历.ListIndex = -1 And .学历 <> "" Then
            cbo学历.AddItem .学历, 0
            cbo学历.ListIndex = cbo学历.NewIndex
        End If
        '    婚姻状况    Varchar2    4
        Call cbo.SeekIndex(cbo婚姻状况, .婚姻状况, , True)
        If cbo婚姻状况.ListIndex = -1 And .婚姻状况 <> "" Then
            cbo婚姻状况.AddItem .婚姻状况, 0
            cbo婚姻状况.ListIndex = cbo婚姻状况.NewIndex
        End If
        '    区域    Varchar2    30
        txt区域.Text = .区域
        '    家庭地址    Varchar2    50
        txt家庭地址.Text = .家庭地址
        Call zlReadAddrInfo(padd家庭地址, .病人ID, 0, 3, .家庭地址)
        '    户口地址    Varchar2    50
        txt户口地址.Text = .户口地址
        Call zlReadAddrInfo(padd户口地址, .病人ID, 0, 4, .户口地址)
        '    家庭电话    Varchar2    20
        txt家庭电话.Text = .家庭电话
        '    家庭地址邮编    Varchar2    6
        txt家庭邮编.Text = .家庭邮编
        '    监护人  Varchar2    64
'        txt监护人.Text = .监护人
        '    联系人姓名  Varchar2    64
        txt联系人姓名.Text = .联系人
        '84313,李南春,2015/4/27,联系人关系以及其他关系
        '    联系人关系  Varchar2    30
        Call cbo.SeekIndex(cbo联系人关系, .联系人关系, , True)
        If cbo联系人关系.ListIndex = -1 And .联系人关系 <> "" Then
            cbo联系人关系.ListIndex = 8
            txt其他关系.Text = .联系人关系
        ElseIf cbo联系人关系.ListIndex = 8 Then
            str其他关系 = Get其他关系(Val(.病人ID))
            txt其他关系.Text = str其他关系
        End If
        '    联系人地址  Varchar2    50
        txt联系人地址.Text = .联系人地址
        '    联系人电话  Varchar2    20
        txt联系人电话.Text = .联系人电话
        '    工作单位    Varchar2    100
        txt工作单位.Text = .工作单位
        lbl工作单位.Tag = ""
        '    单位电话    Varchar2    20
        txt单位电话.Text = .工作单位地址
        '    单位邮编    Varchar2    6
        txt单位邮编.Text = .工作单位邮编
        '    单位开户行  Varchar2    50
        txt单位开户行.Text = .工作单位开户行
        '    单位帐号    Varchar2    20
        txt单位帐户.Text = .工作单位开户行帐户
        '    手机号      Varchar2    20
        txt手机.Text = .手机号
    End With
    FromKindLoadPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetValidKindStr(ByVal lngCardTypeID As Long) As Boolean
    '----------------------------------------------
    '功能：构造有效医疗卡IDKind传入字符串,并判断传入医疗卡类别是否在有效医疗卡中
    '参数：判断是否存在的医疗卡类别ID
    '返回：
    '   1:查找医疗卡类别存在
    '   2:查找医疗卡类别在IDKind控件“属性设置”中未启用
    '   0:查找医疗卡类别在参数设置中未启用
    '编制：冉俊明
    '时间：2014-5-16
    '问题：72541
    '说明：
    '
    '----------------------------------------------
    Dim objCard As Card, i As Long, blnDelete As Boolean
    Dim objCards As Cards
    
    On Error GoTo errHandle
    If Not IDKindPay.Cards Is Nothing Then
        Set objCards = IDKindPay.Cards
        For Each objCard In objCards
            blnDelete = False
            With objCard
                If mEditType = Cr_补卡 Then
                    '只有自制卡才能补卡
                    If Not (zlstr.IsHavePrivs(mstrPrivs, "补卡") And .自制卡) Then blnDelete = True
                Else
                    If Not zlstr.IsHavePrivs(mstrPrivs, "发卡") And .自制卡 Then blnDelete = True   '无发卡权限不能发卡
                    If Not zlstr.IsHavePrivs(mstrPrivs, "绑定卡") And .自制卡 = False Then blnDelete = True '无绑定卡权限不能绑定卡
                End If
                If mblnFromCardMgr And .接口序号 <> lngCardTypeID Then blnDelete = True '发卡界面进入只能对当前卡进行操作
                If .接口序号 = 0 Then blnDelete = True '删除默认发卡类别
                '移除
                If blnDelete Then
                    If .接口序号 = 0 Then
                        objCards.Remove "M" & .名称
                    Else
                        objCards.Remove "K" & .接口序号
                    End If
                Else
                    If .接口序号 = lngCardTypeID Then GetValidKindStr = True
                End If
            End With
        Next

    End If
    Set IDKindPay.Cards = objCards
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载身份证图像
    '编制:刘兴洪
    '日期:2014-06-30 16:20:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    imgPatient.Picture = objStdPic
    mlng图像操作 = 4
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function CreateObjectPlugIn() As Boolean
    '功能:创建渠道附加信息插件
    '返回:创建成功,返回True,否则返回False
    '问题号:73935
    '编制:冉俊明
    '日期:2014-07-3
    mblnPlugin = False
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        Call mobjPlugIn.Initialize(gcnOracle, glngSys, mlngModule)
        mblnPlugin = Err = 0
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
    End If
    CreateObjectPlugIn = True
End Function

Private Function InitTaskPanelOther() As Boolean
    '功能:加载附加信息页面
    '返回:
    '问题号:73935
    '编制:冉俊明
    '日期:2014-07-3
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    
    Err = 0: On Error GoTo Errhand
    If Not mobjPlugIn Is Nothing Then
        If mlngPlugInHwnd <> 0 Then
            With wndTaskPanelOther
                Call .SetGroupInnerMargins(0, 0, 0, 0)
                Call .SetGroupOuterMargins(-1, -24, -1, -1)
                
                Set tkpGroup = .Groups.Add(1, "附加信息")
                tkpGroup.CaptionVisible = False
                tkpGroup.Expandable = False
                tkpGroup.Expanded = True
                
                Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
                Call HideFormCaption(mlngPlugInHwnd, False) '隐藏窗体边框
                Item.Handle = mlngPlugInHwnd
                
                .HotTrackStyle = xtpTaskPanelHighlightItem
                .Reposition
                .DrawFocusRect = True
            End With
        End If
    End If
    InitTaskPanelOther = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Sub HideFormCaption(ByVal lngHwnd As Long, Optional ByVal blnBorder As Boolean = True)
'功能：隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(lngHwnd, vRect)
    lngStyle = GetWindowLong(lngHwnd, GWL_STYLE)

    If blnBorder Then
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
    Else
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    End If
    SetWindowLong lngHwnd, GWL_STYLE, lngStyle
    SetWindowPos lngHwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建zlPublicPatient部件
    '返回:创建成功,返回True,否则返回False
    '编制:冉俊明
    '日期:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "病人信息公共部件（zlPublicPatient）创建失败！", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "病人信息公共部件（zlPublicPatient）初始化失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CreatePublicPatient = True
End Function

Private Sub SetCmdCtrlMove()
    '78408:李南春,2014/10/9,过敏药物选择方式
    With vsDrug
        If .Row >= 1 And .Col = 0 And .Visible = True And .Enabled = True Then
            cmdSelDrug.Left = .CellLeft + .CellWidth - cmdSelDrug.Width
            cmdSelDrug.Top = .CellTop + 15
            cmdSelDrug.Visible = True
        Else
            cmdSelDrug.Visible = False
        End If
    End With
End Sub

Private Sub InitControl()
    '功能:根据参数mParaData.strControl设置输入框的属性，并统计必须输入项目。
    '80503:李南春,2015/1/23,输入项目参数控制
    Dim objCtl As Control, Arr() As String, SubArr() As String
    Dim i As Integer
    
    mstr必输项目 = ""
    Arr() = Split(mParaData.strControl & "|", "|")
    For Each objCtl In Controls
        For i = LBound(Arr) To UBound(Arr)
            SubArr() = Split(Arr(i) & ",", ",")
            If SubArr(0) = objCtl.Tag And (UCase(TypeName(objCtl)) = UCase("TextBox") Or UCase(TypeName(objCtl)) = UCase("ComboBox") _
                                            Or UCase(TypeName(objCtl)) = UCase("CommandButton")) And UBound(SubArr) = 4 Then
                If SubArr(1) = 1 Then
                    objCtl.Enabled = False
                    objCtl.BackColor = &H8000000F
                Else
                    If SubArr(2) = 1 Then mstr必输项目 = mstr必输项目 & SubArr(0) & ","
                    objCtl.TabStop = IIf(SubArr(3) = 1, True, False)
                End If
                Exit For
            End If
        Next
    Next
    '验证医保号与医保号同步处理
    txt验证医保号.Enabled = txt医保号.Enabled
    txt验证医保号.BackColor = txt医保号.BackColor
    If InStr(mstr必输项目, "医保号") > 0 Then mstr必输项目 = mstr必输项目 & "验证医保号" & ","
    txt验证医保号.TabStop = txt医保号.TabStop
End Sub

Private Function Check必须输入项(ByVal objCtl As Control) As Boolean
    '功能：检查是否是必须输入对象。
    '入参：objCtl-检查的对象
    '80503:李南春,2015/1/23,输入项目参数控制
    If InStr("," & mstr必输项目 & ",", "," & objCtl.Tag & ",") > 0 And objCtl.Text = "" Then
        MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "必须输入,请检查", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Check必须输入项 = True
End Function

Private Sub Led欢迎信息()
    Dim strInfo As String, lngPatient As Long
    'LED初始化
    If gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModule, gcnOracle
        End If
        strInfo = Trim(txtPatient.Text)
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!性别 & " " & mrsInfo!年龄: lngPatient = Val(Nvl(mrsInfo!病人ID))
        End If
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub

Private Function Get其他关系(ByVal lng病人ID As Long) As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "Select 信息值  From 病人信息从表 Where 病人ID=[1] And 信息名='联系人附加信息'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If Not rsTemp.EOF Then Get其他关系 = Nvl(rsTemp!信息值)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Clear健康档案()
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:清除界面信息
'入参:
'编制:李南春
'日期:2015/4/30 10:38:36
'问题:84424
'---------------------------------------------------------------------------------------------------------------------------------------------
    '血型
    If cboBloodType.ListCount > 0 Then cboBloodType.ListIndex = -1
    'RH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
    '医学警示
    txtMedicalWarning.Text = ""
    '其他医学警示
    txtOtherWaring.Text = ""
    '联系人信息
    With vsLinkMan
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
    End With
    '接种情况
    With vsInoculate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '其他信息
    With vsOtherInfo
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '过敏反应
    With vsDrug
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        If .Cols > 2 Then .TextMatrix(1, 2) = ""
    End With
    '证件信息
    With vsCertificate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
End Sub

Private Function LinkManValid() As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:清除界面信息
'入参:
'编制:李南春
'日期:2015/4/30 10:38:36
'问题:84672
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    '长度检查以及联系人检查
    If CheckTextLength("其他关系", txt其他关系) = False Then Exit Function
    If txt联系人姓名.Text = "" And (txt联系人电话.Text <> "" Or txt联系人身份证号.Text <> "" Or cbo联系人关系.Text <> "") Then
        If MsgBox("没有输入联系人姓名，联系人信息不会保存，是否继续？", vbYesNo + vbInformation, gstrSysName) = vbNo Then
            Exit Function
        Else
            txt联系人电话.Text = "": txt联系人身份证号.Text = ""
            cbo联系人关系.ListIndex = -1: txt其他关系.Text = "": txt其他关系.Visible = False
        End If
    End If
    With vsLinkMan
        If .Rows >= 3 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) = "" And (.TextMatrix(i, 1) <> "" Or .TextMatrix(i, 2) <> "" Or .TextMatrix(i, 3) <> "") Then
                    If MsgBox("联系人列表第" & i & "行没有输入联系人姓名，此行联系人信息不会保存，是否继续？", vbYesNo + vbInformation, gstrSysName) = vbNo Then
                        Exit Function
                    Else
                        .TextMatrix(i, 1) = "": .TextMatrix(i, 2) = "": .TextMatrix(i, 3) = "": .TextMatrix(i, 4) = ""
                    End If
                End If
            Next
        End If
    End With
    LinkManValid = True
End Function

Private Sub InitCertificate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:90875
    '日期:2015/12/17 16:59:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand
    Dim strSQL As String, rsTemp As ADODB.Recordset, str关系 As String, i As Integer
    With vsCertificate
    '初始化列表属性
        vsCertificate.Editable = flexEDKbdMouse
    '设置列头
        SetColumHeader vsCertificate, C_CertificateHeader
    '设置列信息
        strSQL = "Select 名称,缺省标志 from 证件类型  Where  名称 Not Like '其他%' and 名称 Not Like '%身份证'" & vbNewLine & _
                " And Not 名称 in (Select 名称 from  医疗卡类别 Where Nvl(是否证件,0)=0 or Nvl(是否启用,0)=0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTemp.RecordCount = 0 Then .Editable = flexEDNone: Exit Sub
        Do While Not rsTemp.EOF
            str关系 = str关系 & "|" & Nvl(rsTemp!名称)
            rsTemp.MoveNext
        Loop
        str关系 = Mid(str关系, 2)
        If str关系 <> "" Then .ColComboList(0) = str关系: .ColComboList(2) = str关系
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsCertificate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long
    If Row < 1 Or Col < 0 Then Exit Sub
    '问题号:90875
    With vsCertificate
        If Col = 1 Or Col = 3 Then '证件号码不能超过30
           If Len(.TextMatrix(Row, Col)) > 30 Then
                MsgBox "证件输入字符超出最大字符数30,多出的字符将被自动截除！", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 30)
           End If
           If Col = 3 And .Rows - 1 = Row And .TextMatrix(Row, Col) <> "" Then
                .Rows = .Rows + 1
            End If
        ElseIf Col = 0 Or Col = 2 Then '检查是否选择了重复的证件类型
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If (lngRow <> Row Or lngCol <> Col) And .TextMatrix(lngRow, lngCol) = .TextMatrix(Row, Col) And .TextMatrix(Row, Col) <> "" Then
                        MsgBox .TextMatrix(lngRow, lngCol) & "已存在，不能重复选择。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        .Select Row, Col
                        Exit Sub
                    End If
                Next
            Next
        End If
    End With
End Sub
Private Sub vsCertificate_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:90875
    If KeyCode = 27 And vsCertificate.Rows = 2 Then
        If vsCertificate.TextMatrix(1, 3) <> "" Then
            vsCertificate.TextMatrix(1, 2) = "": vsCertificate.TextMatrix(1, 3) = ""
        Else
            vsCertificate.TextMatrix(1, 0) = "": vsCertificate.TextMatrix(1, 1) = ""
        End If
    End If
    If KeyCode = 27 And vsCertificate.Rows > 2 Then 'Esc
        If vsCertificate.TextMatrix(vsCertificate.Rows - 1, 2) <> "" Or vsCertificate.TextMatrix(vsCertificate.Rows - 1, 3) <> "" Then
            vsCertificate.TextMatrix(vsCertificate.Rows - 1, 2) = "": vsCertificate.TextMatrix(vsCertificate.Rows - 1, 3) = ""
        Else
            vsCertificate.Rows = vsCertificate.Rows - 1
        End If
    End If
End Sub

Private Sub vsCertificate_KeyPress(KeyAscii As Integer)
    '78408:李南春,2014/10/9,光标跳转
    If KeyAscii = 13 Then
        If vsCertificate.Col = 3 And vsCertificate.Rows - 1 = vsCertificate.Row Then
            zlCommFun.PressKey vbKeyTab
        ElseIf vsCertificate.Col = 3 Then
            vsCertificate.Col = 0: vsCertificate.Row = vsCertificate.Row + 1
            zlCommFun.PressKey vbKeyReturn
        Else
            zlCommFun.PressKey vbKeyRight
        End If
    End If
End Sub

Private Sub LoadCertificate(ByVal lng病人ID As Long)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人的证件信息到界面
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    
    On Error GoTo Errhand
    strSQL = "Select  A.名称,A.ID,B.卡号 from 医疗卡类别 A, 病人医疗卡信息 B " & _
            "Where A.ID= B.卡类别ID And A.是否启用=1 And A.是否证件=1 And B.状态=0  And  B.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsCertificate
        .Clear 1
        .Rows = 2
        lngRow = 1: lngCol = 0
        While Not rsTemp.EOF
            .TextMatrix(lngRow, lngCol) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, lngCol + 1) = Nvl(rsTemp!卡号)
            lngCol = lngCol + 2
            If lngCol > 2 Then .Rows = .Rows + 1: lngRow = lngRow + 1: lngCol = 0
            rsTemp.MoveNext
        Wend
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub AddCardTypeSQL(ByVal intOper As Integer, ByVal lng卡类别ID As Long, ByVal strCode As String, ByVal str全名 As String, ByVal str短名 As String, _
                           ByVal lng卡号长度 As Long, ByRef colPro As Collection)
    Dim strSQL As String

    ' Zl_医疗卡类别_Update
    strSQL = "Zl_医疗卡类别_Update("
    '  Id_In           In 医疗卡类别.ID%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  编码_In         In 医疗卡类别.编码%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '  名称_In         In 医疗卡类别.名称%Type,
    strSQL = strSQL & "'" & str全名 & "',"
    '  短名_In         In 医疗卡类别.短名%Type,
    strSQL = strSQL & "'" & str短名 & "',"
    '  前缀文本_In     In 医疗卡类别.前缀文本%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  卡号长度_In     In 医疗卡类别.卡号长度%Type,
    strSQL = strSQL & "" & lng卡号长度 & ","
    '  缺省标志_In     In 医疗卡类别.缺省标志%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否固定_In     In 医疗卡类别.是否固定%Type,
    strSQL = strSQL & "1,"
    '  是否严格控制_In In 医疗卡类别.是否严格控制%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否自制_In     In 医疗卡类别.是否自制%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否存在帐户_In In 医疗卡类别.是否存在帐户%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否全退_In     In 医疗卡类别.是否全退%Type,
    strSQL = strSQL & "0,"
    '  部件_In         In 医疗卡类别.部件%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  备注_In         In 医疗卡类别.备注%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  特定项目_In     In 医疗卡类别.特定项目%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '    收费细目id_In   In 收费项目目录.ID%Type,
    strSQL = strSQL & "" & "0" & ","
    '  结算方式_In     In 医疗卡类别.结算方式%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  是否启用_In     In 医疗卡类别.是否启用%Type,
    strSQL = strSQL & "1,"
    '  卡号密文_In     In 医疗卡类别.卡号密文%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  是否重复使用_In In 医疗卡类别.是否重复使用%Type,
    strSQL = strSQL & "" & 1 & ","
    '密码长度_In     In 医疗卡类别.密码长度%Type,
    strSQL = strSQL & "" & 10 & ","
    '密码长度限制_In In 医疗卡类别.密码长度限制%Type,
    strSQL = strSQL & "" & 0 & ","
    '密码规则_In     In 医疗卡类别.密码规则%Type,
    strSQL = strSQL & "" & 0 & ","
    strSQL = strSQL & "" & 1 & ","
    '  操作方式_In     In Integer := 0
    strSQL = strSQL & "" & intOper & ","
    '是否模糊查找_In     In 医疗卡类别.是否模糊查找%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '问题号:51072
    '密码输入限制_In     In 医疗卡类别.密码输入限制%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '是否缺省密码_In     In 医疗卡类别.是否缺省密码%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '问题号:56508
    '是否制卡_In
    strSQL = strSQL & "" & 0 & ","
    '是否发卡_In
    strSQL = strSQL & "" & 0 & ","
    '是否写卡_In
    strSQL = strSQL & "" & 0 & ","
    '问题号:57697
    '险类_In
    strSQL = strSQL & "" & 0 & ","
    '问题号:57326
    strSQL = strSQL & "" & 1 & ","
    '77872,李南春,2014/12/3:是否支持转帐及代扣
    '是否转帐及代扣_In  In 医疗卡类别.是否转帐及代扣%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '读卡性质_In       In 医疗卡类别.读卡性质%Type := '1000',
    strSQL = strSQL & "" & "1000" & ","
    '键盘控制方式_In   In 医疗卡类别.键盘控制方式%Type := 0,
    strSQL = strSQL & "" & 0 & ","
    '90875:李南春,2015/12/16,增加医疗卡证件类型
    '是否证件_In  In 医疗卡类别.是否证件%Type:=0
    strSQL = strSQL & "" & 1 & ")"
    
    zlAddArray colPro, strSQL
End Sub

Private Sub AddCertificate(ByVal lng病人ID As Long, ByRef colPro As Collection, ByVal dtCurdate As Date)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:建立证件卡类信息，如果是第一次建立卡类别
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, rsPatiCard As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    Dim lngID As Long, strCode As String
    
    On Error GoTo Errhand
    '绑定卡前要判断卡类别是否存在
    strSQL = "Select B.ID,B.编码,B.卡号长度,B.名称,A.卡号,A.病人ID,Decode(A.卡号 ,NULL,1,0) as 标识 from 病人医疗卡信息 A,医疗卡类别 B " & _
            " Where A.卡类别ID(+)=B.ID And B.是否证件=1 And A.状态(+)=0 And A.病人ID(+)=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    Set rsPatiCard = zlDatabase.CopyNewRec(rsTemp)
    With vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    lngID = 0: strCode = ""
                    rsTemp.Filter = "名称='" & .TextMatrix(lngRow, lngCol) & "'"
                    If rsTemp.RecordCount = 0 Then
                        lngID = zlDatabase.GetNextId("医疗卡类别")
                        If mstrFirstCode = "" Then
                            strCode = zlDatabase.GetMax("医疗卡类别", "编码", 4)
                            mstrFirstCode = strCode
                        Else
                            strCode = CStr(Val(mstrFirstCode) + 1)
                            strCode = Format(strCode, String(4, "0"))
                            mstrFirstCode = strCode
                        End If
                        Call AddCardTypeSQL(0, lngID, strCode, .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), colPro)
                    ElseIf Len(.TextMatrix(lngRow, lngCol + 1)) > Val(Nvl(rsTemp!卡号长度)) Then
                        Call AddCardTypeSQL(1, Val(Nvl(rsTemp!id)), Nvl(rsTemp!编码), .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), colPro)
                    End If

                    '进行证件卡绑定
                    rsPatiCard.Filter = "名称='" & .TextMatrix(lngRow, lngCol) & "' And 卡号='" & .TextMatrix(lngRow, lngCol + 1) & "'"
                    If rsPatiCard.RecordCount = 0 Then
                        'Zl_医疗卡变动_Insert
                         strSQL = "Zl_医疗卡变动_Insert("
                        '      变动类型_In   Number,
                        '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
                        strSQL = strSQL & "" & 11 & ","
                        '      病人id_In     住院费用记录.病人id%Type,
                        strSQL = strSQL & "" & lng病人ID & ","
                        '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
                        strSQL = strSQL & "" & IIf(lngID = 0, rsTemp!id, lngID) & ","
                        '      原卡号_In     病人医疗卡信息.卡号%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      医疗卡号_In   病人医疗卡信息.卡号%Type,
                        strSQL = strSQL & "'" & .TextMatrix(lngRow, lngCol + 1) & "',"
                        '      变动原因_In   病人医疗卡变动.变动原因%Type,
                        '      --变动原因_In:如果密码调整，变动原因为密码.加密的
                        strSQL = strSQL & "'" & "证件卡绑定" & "',"
                        '      密码_In       病人信息.卡验证码%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        '      变动时间_In   住院费用记录.登记时间%Type,
                        strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                        '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
                        strSQL = strSQL & "'" & "" & "',"
                        '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
                        strSQL = strSQL & "NULL)"
                    
                        zlAddArray colPro, strSQL
                    Else
                        rsPatiCard!标识 = 1
                        rsPatiCard.Update
                    End If
                End If
            Next
        Next
    End With
    mstrFirstCode = ""
    
    '卡号列表中没有证件号，要解除绑定
    rsPatiCard.Filter = "标识=0"
    If rsPatiCard.RecordCount > 0 Then
        rsPatiCard.MoveFirst
        Do While Not rsPatiCard.EOF
            'Zl_医疗卡变动_Insert
             strSQL = "Zl_医疗卡变动_Insert("
            '      变动类型_In   Number,
            '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
            strSQL = strSQL & "" & 14 & ","
            '      病人id_In     住院费用记录.病人id%Type,
            strSQL = strSQL & "" & lng病人ID & ","
            '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
            strSQL = strSQL & "" & rsPatiCard!id & ","
            '      原卡号_In     病人医疗卡信息.卡号%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      医疗卡号_In   病人医疗卡信息.卡号%Type,
            strSQL = strSQL & "'" & rsPatiCard!卡号 & "',"
            '      变动原因_In   病人医疗卡变动.变动原因%Type,
            '      --变动原因_In:如果密码调整，变动原因为密码.加密的
            strSQL = strSQL & "'" & "证件卡取消绑定" & "',"
            '      密码_In       病人信息.卡验证码%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '      变动时间_In   住院费用记录.登记时间%Type,
            strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
            strSQL = strSQL & "'" & "" & "',"
            '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
            strSQL = strSQL & "NULL)"
        
            zlAddArray colPro, strSQL
            rsPatiCard.MoveNext
        Loop
    End If
    rsPatiCard.Close
    Exit Sub
Errhand:
    rsPatiCard.Close
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function IsCertificateCard(ByVal lng病人ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:证件卡类检查
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, str证件类型 As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCardName As String
    
    On Error GoTo Errhand
    With vsCertificate
        '检查输入是否完整
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    MsgBox "请选择卡号" & .TextMatrix(lngRow, lngCol + 1) & "的证件类型", vbInformation, gstrSysName
                    .Select lngRow, lngCol
                    Exit Function
                End If
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    strSQL = "Select 1 from 病人医疗卡信息 A,医疗卡类别 B " & _
                            "Where A.卡类别ID=B.ID And B.名称=[1] And B.是否证件=1 And A.卡号=[2] And  A.病人ID<>[3]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, lngCol), Trim(.TextMatrix(lngRow, lngCol + 1)), lng病人ID)
                    If rsTmp.RecordCount > 0 Then
                        MsgBox .TextMatrix(lngRow, lngCol) & ":" & .TextMatrix(lngRow, lngCol + 1) & "正在被使用,请检查!", vbInformation, gstrSysName
                        .Select lngRow, lngCol
                        Exit Function
                    End If
                    str证件类型 = str证件类型 & ",'" & .TextMatrix(lngRow, lngCol) & "'"
                End If
            Next
        Next
        
        '检查证件类型是否与非证件的医疗卡类别重复，重复则不保存信息
        str证件类型 = Mid(str证件类型, 2)
        If str证件类型 = "" Then IsCertificateCard = True: Exit Function
        strSQL = "Select 名称 From 医疗卡类别 where 名称 in (" & str证件类型 & ") And Nvl(是否证件,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strCardName = strCardName & "," & Nvl(rsTmp!名称)
            Loop
            
            strCardName = Mid(strCardName, 2)
            MsgBox "医疗卡类别【" & strCardName & "】名称重复,不能继续添加。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    IsCertificateCard = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsCheck医疗卡() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医疗卡的输入是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-06 17:44:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCard As String, strICCard As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo Errhand
    If chk病历费.value And Not (mEditType = Cr_退卡 Or chkCancel.value = 1) Then
        If Not mFeeType.rs病历费 Is Nothing Then
            If mFeeType.rs病历费!是否变价 = 1 Then
                If mFeeType.rs病历费!现价 <> 0 And Abs(CCur(Val(txt病历费.Text))) > Abs(mFeeType.rs病历费!现价) Then
                    MsgBox "病历费绝对值不能大于最高限价：" & Format(Abs(mFeeType.rs病历费!现价), "0.00"), vbExclamation, gstrSysName
                    If txt病历费.Enabled And txt病历费.Visible Then txt病历费.SetFocus:  Exit Function
                End If
                
                If mFeeType.rs病历费!原价 <> 0 And Abs(CCur(Val(txt病历费.Text))) < Abs(mFeeType.rs病历费!原价) Then
                    MsgBox "病历费绝对值不能小于最低限价：" & Format(Abs(mFeeType.rs病历费!原价), "0.00"), vbExclamation, gstrSysName
                    If txt病历费.Enabled And txt病历费.Visible Then txt病历费.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    
    '检查本次充值金额
    '114422:李南春：2017/11/3,先判断结算方式
    If cbo支付方式.Enabled And cbo支付方式.Visible Then
        If cbo支付方式.ListIndex = -1 Then
            MsgBox "就诊卡场合没有可用的结算方式，请先到结算方式管理中设置！", vbExclamation, gstrSysName
            If cbo支付方式.Enabled And cbo支付方式.Visible Then cbo支付方式.SetFocus: Exit Function
        ElseIf IDKindPayMode.IDKind = 2 And cbo支付方式.ItemData(cbo支付方式.ListIndex) <> 1 And Val(txt余额.Text) < 0 Then
            MsgBox "充预交金额不能为负数，请再次确认缴款金额！", vbExclamation, gstrSysName
            If txt合计.Enabled And txt合计.Visible Then txt合计.SetFocus: Exit Function
        End If
    End If
    
    '不是绑定卡、发卡、补卡、换卡则退出检查
    If Not (mEditType = Cr_绑定卡 Or mEditType = Cr_发卡 Or mEditType = Cr_补卡 Or mEditType = Cr_换卡) Or chkCancel.value = 1 Then IsCheck医疗卡 = True: Exit Function
    strCard = UCase(txt卡号.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    '1.变价金额检查
    If (mEditType = Cr_发卡 Or mEditType = Cr_补卡) And (mCardType.bln自制卡 = True Or mCardType.bln是否发卡) Then
        If Not mCardType.rs医疗卡费 Is Nothing Then
            If mCardType.rs医疗卡费!是否变价 = 1 Then
                '70595:刘尔旋,2014-03-04,卡费未输入时报错的情况
                If mCardType.rs医疗卡费!现价 <> 0 And Abs(CCur(Val(txt卡费.Text))) > Abs(mCardType.rs医疗卡费!现价) Then
                    MsgBox mCardType.str卡名称 & "的卡费绝对值不能大于最高限价：" & Format(Abs(mCardType.rs医疗卡费!现价), "0.00"), vbExclamation, gstrSysName
                    If txt卡费.Enabled And txt卡费.Visible Then txt卡费.SetFocus:  Exit Function
                End If
                
                If mCardType.rs医疗卡费!原价 <> 0 And Abs(CCur(Val(txt卡费.Text))) < Abs(mCardType.rs医疗卡费!原价) Then
                    MsgBox mCardType.str卡名称 & "的卡费绝对值不能小于最低限价：" & Format(Abs(mCardType.rs医疗卡费!原价), "0.00"), vbExclamation, gstrSysName
                    If txt卡费.Enabled And txt卡费.Visible Then txt卡费.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    
    '104238:李南春，2017/2/15，检查卡号是否满足发卡控制限制
    If txt卡号.Text <> "" And Len(txt卡号.Text) <> mCardType.lng卡号长度 And Not mCardType.bln严格控制 Then
        Select Case mCardType.byt发卡控制
            Case 0
                MsgBox "输入的卡号小于" & mCardType.str卡名称 & "设定的卡号长度，请重新输入！", vbExclamation, gstrSysName
                If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                    Exit Function
            Case 2
                If MsgBox("输入的卡号小于" & mCardType.str卡名称 & "设定的卡号长度，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                    Exit Function
                End If
        End Select
    End If
    
    '108779:李南春,2017/5/8,如果密码为空，只检查密码输入控制
    If txt卡号.Text <> "" And txtPass.Text <> "" And txtPass.Visible Then
        Select Case mCardType.int密码长度限制
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCardType.int密码长度 Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & mCardType.int密码长度 & "位", vbOKOnly + vbInformation
                txtPass.Text = "": txtAudi.Text = ""
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCardType.int密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(mCardType.int密码长度限制) & "位以上.", vbOKOnly + vbInformation
                txtPass.Text = "": txtAudi.Text = ""
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        End Select
    End If
    
    If txtPass.Text <> txtAudi.Text And txt卡号.Text <> "" Then
        MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus: Exit Function
    End If

    If mEditType = Cr_绑定卡 Or mEditType = Cr_发卡 Or mEditType = Cr_补卡 Or mEditType = Cr_换卡 Then
        If Trim(txt卡号.Text) = "" Then
            MsgBox "请刷卡或输入就诊卡号！", vbExclamation, gstrSysName
            If txt卡号.Enabled And txt卡号.Enabled Then txt卡号.SetFocus
            Exit Function
        End If
    End If
    
    If txt卡号.Text <> "" And (mEditType = Cr_发卡 Or mEditType = Cr_补卡 Or mEditType = Cr_换卡) Then
        '保存前检查就诊卡是否有，是否在范围内
        If CheckBILL(txt卡号.Text) = False Then Exit Function
    End If
    
    If txt卡号.Text <> "" And (mEditType = Cr_发卡 Or mEditType = Cr_补卡 Or mEditType = Cr_换卡) Then
        '保存前检查就诊卡是否有效
        strSQL = "Select 1 From 病人医疗卡信息 Where 卡类别ID=[1] And 卡号=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, txt卡号.Text)
        If rsTmp.RecordCount > 0 Then
            MsgBox "该卡号已被使用，请重新输入新的卡号！", vbExclamation, gstrSysName
            If txt卡号.Enabled And txt卡号.Enabled Then txt卡号.SetFocus
            Exit Function
        End If
        '不允许重复使用的医疗卡，还需要检查以前是否发过卡
        If Not mCardType.bln是否重复使用 Then
            strSQL = "Select 1 From 病人医疗卡变动 Where 卡类别id = [1] And 卡号 = [2] And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, txt卡号.Text)
            If rsTmp.RecordCount > 0 Then
                MsgBox mCardType.str卡名称 & "不支持重复使用，请重新输入未被使用过的卡号！", vbExclamation, gstrSysName
                If txt卡号.Enabled And txt卡号.Enabled Then txt卡号.SetFocus
                Exit Function
            End If
        End If
    End If
    
    IsCheck医疗卡 = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub AddDepositSQL(ByVal lng病人ID As Long, ByVal dtCurdate As Date, _
    ByRef cllPro As Collection, ByRef lng结帐ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加预交款的SQL
    '编制:刘兴洪
    '日期:2011-07-26 18:26:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, strSQL As String, i As Integer
    Dim dblMoney As Double, str结算方式 As String
     
    '病人预交款记录
    str结算方式 = mcolPayMode(cbo支付方式.ListIndex + 1)(6)
    If str结算方式 = "" Then str结算方式 = zlstr.NeedName(cbo支付方式.Text)
    If Not cbo支付方式.Enabled Then str结算方式 = ""
        
    mstrPrePayNo = zlDatabase.GetNextNo(11)
    mlng预交ID = zlDatabase.GetNextId("病人预交记录")
    mlng预交病人ID = lng病人ID
    mdat预交时间 = dtCurdate
    dblMoney = StrToNum(txt余额.Text)
    'Zl_病人预交记录_Insert
    strSQL = "Zl_病人预交记录_Insert("
    '  Id_In         病人预交记录.ID%Type,
    strSQL = strSQL & "" & mlng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSQL = strSQL & "'" & mstrPrePayNo & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strSQL = strSQL & "" & IIf(mblnPrepayPrint, "'" & mstrPrepayInvioce & "'", "Null") & ","
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strSQL = strSQL & "NULL,"
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "NULL,"
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  摘要_In       病人预交记录.摘要%Type,
    strSQL = strSQL & "'医疗卡:" & mCurPayMoney.strNO & "',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(mlng领用ID = 0, "NULL", mlng领用ID) & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSQL = strSQL & "" & 1 & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPayMoney.lng医疗卡类别ID = 0 Or mCurPayMoney.bln消费卡, "NULL", mCurPayMoney.lng医疗卡类别ID) & ","
   '  结算卡序号_in 病人预交记录.结算卡序号%type:=NULL,
    strSQL = strSQL & "" & IIf(mCurPayMoney.lng医疗卡类别ID = 0 Or Not mCurPayMoney.bln消费卡, "NULL", mCurPayMoney.lng医疗卡类别ID) & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPayMoney.str刷卡卡号 = "", "NULL", "'" & mCurPayMoney.str刷卡卡号 & "'") & ","
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  合作单位_In   病人预交记录.合作单位%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  收款时间_In   病人预交记录.收款时间%Type := Null
    '108001:李南春，2017/5/8，格式化预交时间为24小时制
    strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '   操作类型_In Integer:=0 :0-正常缴预交;1-存为划价单
    strSQL = strSQL & "0 )"
    zlAddArray cllPro, strSQL
End Sub

Private Function CheckDepositFactValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取预交发票号
    '返回:正常获取,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-30 11:14:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lng领用ID As Long, strInvoice As String
    Dim blnInput As Boolean, blnValid As Boolean
    
    On Error GoTo errHandle
    mlng领用ID = 0
    
    mstrPrepayInvioce = "": mblnPrepayPrint = False
    '不存在冲预交
    If Not (Val(txt余额.Text) > 0 And IDKindPayMode.IDKind = 2) Then CheckDepositFactValied = True: Exit Function

    mFactProperty = zl_GetInvoicePreperty(mlngModule, 2, 1)
    
    Select Case mFactProperty.intInvoicePrint
    Case 0 '不打印
        CheckDepositFactValied = True: Exit Function
    Case 1 '自动打印
        mblnPrepayPrint = True
    Case 2 '选择是否打印
        If MsgBox("是否打印预交票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) <> vbYes Then CheckDepositFactValied = True: Exit Function
        mblnPrepayPrint = True
    End Select
    
    If mblnBill预交 = False Then
        '有可能是第一次使用
        Do
            blnInput = False
            '非严格控制时直接从本地读取
            strInvoice = UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngModule, ""))
            
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("没有找到已用的预交票据的最大票据号码，无法确定将要使用的开始票据号。" & _
                                vbCrLf & "请输入将要使用的预交票据的开始票据号码：", gstrSysName, _
                                "", Me.Left + 3000, Me.Top + 3000))
                blnInput = True
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                strInvoice = UCase(InputBox("请确认使用的预交票据的开始票据号码：", gstrSysName, _
                                strInvoice, Me.Left + 3000, Me.Top + 3000))
                blnInput = True
            End If
                
            '用户取消输入,允许打印
            If strInvoice = "" Then
                If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnValid = True
            Else
                '检查输入有效性
                If blnInput Then
                    If zlCommFun.ActualLen(strInvoice) <> mbyt预交 Then
                        MsgBox "输入预交的票据号码长度应该为 " & mbyt预交 & " 位！", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
        mstrPrepayInvioce = strInvoice
        CheckDepositFactValied = True: Exit Function
    End If
    
    Do
        '根据票据领用读取
        blnInput = False
        mlng领用ID = CheckUsedBill(2, IIf(mlng领用ID > 0, mlng领用ID, mFactProperty.lngShareUseID), "", 1)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的门诊预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Exit Function
                Case -2
                    MsgBox "本地的共用预交票据的门诊预交票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                    Exit Function
                    strInvoice = ""
            End Select
        End If
        strInvoice = GetNextBill(mlng领用ID)

        If strInvoice = "" Then
            '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
            strInvoice = UCase(InputBox("无法根据票据领用情况获取将要使用预交票据的开始票据号，" & _
                            vbCrLf & "请你输入将要使用的票据号码：", gstrSysName, _
                            "", Me.Left + 1500, Me.Top + 1500))
            blnInput = True
        Else
            strInvoice = UCase(InputBox("请确认使用使用预交票据的票据号码：", gstrSysName, _
                            strInvoice, Me.Left + 1500, Me.Top + 1500))
            blnInput = True
        End If
        
        '用户取消输入,不打印
        If strInvoice = "" Then Exit Function
        
        '检查输入有效性
        If blnInput Then
            mlng领用ID = GetInvoiceGroupID(2, 1, mlng领用ID, mFactProperty.lngShareUseID, strInvoice, 1)
            If lng领用ID < 0 Then
                MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
            Else
                blnValid = True
            End If
        Else
            blnValid = True
        End If
    Loop While Not blnValid
    mstrPrepayInvioce = strInvoice
    CheckDepositFactValied = True: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckChargeFactValied() As Boolean
    Dim strMsg As String
    
    strMsg = "你是否要打印" & IIf(mEditType <> Cr_绑定卡 And gbln收费发票, "发卡票据", "发卡凭条") & "?"
    mPrint.blnPrint = False
    Select Case mPrint.bytPrintType
     Case 0 '不打印
     Case 2 '选择是否打印
         If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
             mPrint.blnPrint = True
         End If
     Case Else
          mPrint.blnPrint = True
    End Select
    
    If mEditType = Cr_绑定卡 Then CheckChargeFactValied = True: Exit Function
    
    
    '收费票据号码检查
    If gbln收费发票 And mPrint.blnPrint Then
        If Trim(txtFact.Text) = "" Then
            MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
            txtFact.SetFocus: Exit Function
        End If
        If gblnBill发卡 Then
            mPrint.lng领用ID = CheckUsedBill(1, IIf(mPrint.lng领用ID > 0, mPrint.lng领用ID, glngShareUseID), txtFact.Text, IIf(gblnStartFactUseType, mPrint.strUseType, ""))
            If mPrint.lng领用ID <= 0 Then
                Select Case mPrint.lng领用ID
                Case 0    '操作失败
                Case -1
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "票据号码不在当前有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    txtFact.SetFocus
                End Select
                Exit Function
            End If
        Else
            If Len(txtFact.Text) <> gbyt收费 And txtFact.Text <> "" Then
                MsgBox "票据号码长度应该为 " & gbyt收费 & " 位！", vbInformation, gstrSysName
                txtFact.SetFocus: Exit Function
            End If
        End If
    Else
        mPrint.lng领用ID = 0
    End If
    CheckChargeFactValied = True
End Function

Private Sub setFact()
    If Not mblnBill预交 And mstrPrepayInvioce <> "" Then
        zlDatabase.SetPara "当前预交票据号", mstrPrepayInvioce, glngSys, mlngModule
    End If
End Sub

Private Function zlSaveEMPIPatiInfo(ByVal blnNewPati As Boolean, ByVal lngPatiID As Long, ByVal lngClinicID As Long, ByRef strErrMsg As String) As Boolean
    '功能:上传病人信息到EMPI平台,如果平台信息保存失败，连同HIS数据一起回退
    '参数: In-lngPatiID 病人ID,lngClinicID 挂号ID
    '      Out-strErrMsg 错误信息，函数返回失败有效
    '返回:True-EMPI平台保存信息成功,False-保存失败
    '编制:李南春
    '说明:101170
    Dim lngRet As Long
    
    If mblnPlugin = False Then zlSaveEMPIPatiInfo = True: Exit Function
    If mobjPlugIn Is Nothing Then zlSaveEMPIPatiInfo = True: Exit Function
    If mEditType <> Cr_发卡 And mEditType <> Cr_绑定卡 And mEditType <> Cr_调整病人信息 Or chkCancel.value = 1 Then zlSaveEMPIPatiInfo = True: Exit Function
    If mEditType <> Cr_调整病人信息 And Not blnNewPati Then zlSaveEMPIPatiInfo = True: Exit Function
    On Error GoTo Errhand
    If mrsEMPIOut Is Nothing Then
        'EMPI没有病人信息，需要新建
        On Error Resume Next
        lngRet = mobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    Else
        '判断平台回传的信息是否发生改变
        With mrsEMPIOut
            If Not txt门诊号.Locked And ExistClinicNO(Nvl(!门诊号), lngPatiID) = False Then
                If txt门诊号.Text <> Nvl(!门诊号) Then GoTo EMPIModify
            End If
            If txt医保号.Text <> Nvl(!医保号) Then GoTo EMPIModify
            If txt身份证号.Text <> Nvl(!身份证号) Then GoTo EMPIModify
            If InStr(mstrPrivsPubPatient, ";基本信息调整;") > 0 Or blnNewPati Then
                If txtPatient.Text <> Nvl(!姓名) Then GoTo EMPIModify
                If cbo性别.ListIndex <> cbo.FindIndex(cbo性别, Nvl(!性别), True) Then GoTo EMPIModify
                If Format(txt出生日期.Text, "YYYY-MM-DD") <> Format(Nvl(!出生日期), "YYYY-MM-DD") Then GoTo EMPIModify
                If Format(txt出生时间.Text, "HH:MM") <> Format(Nvl(!出生日期), "HH:MM") Then GoTo EMPIModify
            End If
            If txt出生地点.Text <> Nvl(!出生地点) Then GoTo EMPIModify
            If cbo国籍.ListIndex <> cbo.FindIndex(cbo国籍, Nvl(!国籍), True) Then GoTo EMPIModify
            If cbo民族.ListIndex <> cbo.FindIndex(cbo民族, Nvl(!民族), True) Then GoTo EMPIModify
            If cbo职业.ListIndex <> cbo.FindIndex(cbo职业, Nvl(!职业)) Then GoTo EMPIModify
            If txt工作单位.Text <> Nvl(!工作单位) Then GoTo EMPIModify
            If txt家庭电话.Text <> Nvl(!家庭电话) Then GoTo EMPIModify
            If txt联系人电话.Text <> Nvl(!联系人电话) Then GoTo EMPIModify
            If txt单位电话.Text <> Nvl(!单位电话) Then GoTo EMPIModify
            If txt家庭地址.Text <> Nvl(!家庭地址) Then GoTo EMPIModify
            If txt家庭邮编.Text <> Nvl(!家庭地址邮编) Then GoTo EMPIModify
            If txt户口地址.Text <> Nvl(!户口地址) Then GoTo EMPIModify
            If txt户口地址邮编.Text <> Nvl(!户口地址邮编) Then GoTo EMPIModify
            If txt单位邮编.Text <> Nvl(!单位邮编) Then GoTo EMPIModify
            If txt联系人姓名.Text <> Nvl(!联系人姓名) Then GoTo EMPIModify
            If cbo联系人关系.ListIndex <> cbo.FindIndex(cbo联系人关系, Nvl(!联系人关系), True) Then GoTo EMPIModify
        End With
    End If
    zlSaveEMPIPatiInfo = True
    Exit Function
EMPIModify:
    On Error Resume Next
    lngRet = mobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, 0, lngClinicID, strErrMsg)
    Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
    If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
    Err.Clear: On Error GoTo Errhand
    zlSaveEMPIPatiInfo = True
    Exit Function
Errhand:
    strErrMsg = Err.Description
    Call zlPlugInErrH(Err, "zlSaveEMPIPatiInfo")
    Call SaveErrLog
End Function

Private Function CheckMobile(strMobile As String) As Boolean
    '检查是否与其他病人的手机号重复
    Dim strSQL As String, lng病人ID As Long
    Dim rsTmp As ADODB.Recordset
    On Error GoTo Errhand
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State = 0 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    strSQL = "Select 1 From 病人信息 Where 手机号=[1] And 病人ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人发卡管理", strMobile, lng病人ID)
    If rsTmp.RecordCount > 0 Then
        If MsgBox("输入的手机号与其他病人重复，是否确定录入？", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Function
    End If
    CheckMobile = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RefreshFact(Optional ByVal blnNew As Boolean, Optional ByVal strFact As String) As Boolean
'参数：blnNew=是否新单保存时调用,这时对于非严格控制的票据是保存当前号
    If Not gbln收费发票 Then mPrint.lng领用ID = 0: Exit Function
    If gblnBill发卡 Then
        mPrint.lng领用ID = CheckUsedBill(1, IIf(mPrint.lng领用ID > 0, mPrint.lng领用ID, glngShareUseID), , IIf(gblnStartFactUseType, mPrint.strUseType, ""))
        If mPrint.lng领用ID <= 0 Then
            Select Case mPrint.lng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            txtFact.Text = "": Exit Function
        Else
            '严格：取下一个号码
            txtFact.Text = GetNextBill(mPrint.lng领用ID)
        End If
    Else
        '松散：取下一个号码
        If Not blnNew Then
            strFact = zlDatabase.GetPara("当前收费票据号", glngSys, 1121)
            txtFact.Text = zlstr.Increase(strFact)
        Else
            zlDatabase.SetPara "当前收费票据号", strFact, glngSys, 1121
            txtFact.Text = zlstr.Increase(strFact)
        End If
    End If
    RefreshFact = True
End Function

Private Sub ReInitPatiInvoice(Optional ByVal lng病人ID_In As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng病人ID As Long
    Dim lng领用ID As Long, strUseType As String
    lng领用ID = mPrint.lng领用ID: mPrint.lng领用ID = 0
    If gbln收费发票 = False Then Exit Sub '使用门诊收据
    If mPrint.bytPrintType = 0 Then Exit Sub '票据允许打印
    If mEditType = Cr_查询 Or mEditType = Cr_挂失 Or mEditType = Cr_绑定卡 Or mEditType = Cr_调整病人信息 Or mEditType = Cr_退卡 Or chkCancel.value = 1 Then Exit Sub '支持使用票据
    
    lng病人ID = lng病人ID_In
    mPrint.lng领用ID = lng领用ID
    If lng病人ID_In = 0 Then
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                lng病人ID = mrsInfo!病人ID
            End If
        End If
    End If
    strUseType = mPrint.strUseType
    mPrint.strUseType = "": mPrint.lng领用ID = 0
    mPrint.strUseType = zl_GetInvoiceUserType(lng病人ID, 0, 0)
    '切换了票据类型
    If mPrint.strUseType <> strUseType And gblnStartFactUseType Then mPrint.lng领用ID = 0

    Call RefreshFact
End Sub
