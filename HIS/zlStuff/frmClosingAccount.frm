VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmClosingAccount 
   Caption         =   "卫材结存管理"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10965
   Icon            =   "frmClosingAccount.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picIni 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1800
      Picture         =   "frmClosingAccount.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picShowDetail 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   4200
      ScaleHeight     =   4695
      ScaleWidth      =   9015
      TabIndex        =   8
      Top             =   1920
      Width           =   9015
      Begin VB.CommandButton cmd药品 
         Height          =   300
         Left            =   3600
         Picture         =   "frmClosingAccount.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   270
      End
      Begin VB.ComboBox cbo单位 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   60
         Width           =   1395
      End
      Begin VB.TextBox txt明细卫材 
         Height          =   300
         Left            =   600
         TabIndex        =   9
         Top             =   60
         Width           =   3000
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMistake 
         Height          =   765
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   3600
         Visible         =   0   'False
         Width           =   2895
         _cx             =   5106
         _cy             =   1349
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
         BackColorSel    =   16764622
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   275
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClosingAccount.frx":13CE
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfStuff 
         Height          =   765
         Left            =   0
         TabIndex        =   13
         Top             =   2040
         Width           =   4455
         _cx             =   7858
         _cy             =   1349
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
         BackColorSel    =   16764622
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   20
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   275
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClosingAccount.frx":14DB
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   765
         Left            =   0
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   5295
         _cx             =   9340
         _cy             =   1349
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
         BackColorSel    =   16764622
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   275
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClosingAccount.frx":1809
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfMistake 
         Height          =   765
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   3000
         Width           =   2895
         _cx             =   5106
         _cy             =   1349
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
         BackColorSel    =   16764622
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   275
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClosingAccount.frx":18EF
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
         VirtualData     =   0   'False
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
      Begin VB.Label lbl明细卫材 
         AutoSize        =   -1  'True
         Caption         =   "卫材"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lbl单位 
         AutoSize        =   -1  'True
         Caption         =   "单位"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   5040
         TabIndex        =   16
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   3
      Top             =   480
      Width           =   3015
      Begin VB.PictureBox picList 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   2055
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
         Begin VSFlex8Ctl.VSFlexGrid vsfList 
            Height          =   1005
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   4815
            _cx             =   8493
            _cy             =   1773
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
            BackColorSel    =   16764622
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   275
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmClosingAccount.frx":1986
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
            VirtualData     =   0   'False
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
            Begin VB.Image Image1 
               Height          =   15
               Left            =   1080
               Top             =   240
               Width           =   135
            End
         End
      End
      Begin VB.ComboBox cbo库房 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   600
         TabIndex        =   4
         Text            =   "cbo库房"
         Top             =   60
         Width           =   1935
      End
      Begin VB.Label lbl库房 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "库房"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
      Begin VB.Frame fraLine 
         Height          =   2085
         Left            =   120
         TabIndex        =   1
         Top             =   -120
         Width           =   45
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   7560
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14261
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
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmClosingAccount.frx":1B8B
      Left            =   1200
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmClosingAccount.frx":1B9F
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   240
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmClosingAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'默认的窗体大小
Private Const mcstlngWinNormalWidth As Long = 12000
Private Const mcstlngWinNormalHeight As Long = 8000

Private mrsAccount As ADODB.Recordset         '用于缓存结存记录
Private mrsDetail As ADODB.Recordset
Private mrsMistake As ADODB.Recordset

Private mint结存方式 As Integer         '结存方式 -1-手工结存 >=0自动结存
Private mint结存时点 As Integer         '结存时点
Private mlng结存ID As Long
Private mstr当前日期 As String          '当前系统日期


Private mblnStart As Boolean

Private Const mconPane_Dept_Condition = 1                     '条件栏

''编辑菜单
'卫材结存
Private Const mconMenu_Edit_CA_AddIniAccount = 3300            '初始结存
Private Const mconMenu_Edit_CA_AddNewAccount = 3301             '新增结存记录
Private Const mconMenu_Edit_CA_VerifyAccount = 3302             '审核结存记录
Private Const mconMenu_Edit_CA_CancelAccount = 3303             '取消结存记录
Private Const mconMenu_Edit_CA_VerifyMistake = 3304             '审核结存误差
Private Const mconMenu_Edit_CA_DeleteAccount = 3305             '删除结存记录

Private Const mconMenu_CA_Refresh = 7001                        '刷新

'结存明细列表
Private Const mconTab_CA_Detail = 0                 '结存明细列表
Private Const mconTab_CA_Stuff = 1                '卫材明细列表
Private Const mconTab_CA_Mistake = 2                '结存误差列表
Private Const mstrCaption As String = "卫材结存管理"

Private Type Type_Privs
    bln所有库房 As Boolean
    bln初始化 As Boolean
    bln审核 As Boolean
End Type
Private mPrives As Type_Privs

Private mstrPrivs As String


Private Sub cbo单位_Click()
    Dim intIndex As Integer
    
    If mblnStart = False Then Exit Sub
    With vsfList
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("审核人")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    With cbo单位
        If Val(.Tag) <> .ListIndex Then
            .Tag = .ListIndex
            If tbcDetail.Selected.Index = mconTab_CA_Detail Then
                LoadInOutList intIndex, mlng结存ID
            ElseIf tbcDetail.Selected.Index = mconTab_CA_Stuff Then
                LoadDetailList intIndex, mlng结存ID
            ElseIf tbcDetail.Selected.Index = mconTab_CA_Mistake Then
                LoadMistakeList intIndex, mlng结存ID
            End If
        End If
    End With
End Sub


Private Sub cbo库房_Click()
    Dim lng库房ID As Long
    Dim Str期间 As String
    Dim strSQL As String
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    
    If Val(cbo库房.Tag) = cbo库房.ListIndex Then Exit Sub '选择的还是上次库房不需刷新
    cbo库房.Tag = cbo库房.ListIndex
    
    If mblnStart = True Then
        Call CheckClosAccount
    End If
    
    If mblnStart = False Then Exit Sub
    
    Call GetAccountRecord
    Call RefreshList
    
End Sub

Private Sub cbo库房_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo库房.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cbo库房.ListIndex >= 0 Then
        If Val(cbo库房.Tag) = cbo库房.ListIndex Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cbo库房, Trim(cbo库房.Text), "V,W") = False Then
        cbo库房.ListIndex = Val(cbo库房.Tag) '未找到，还原上次库房
        Exit Sub
    End If
    If cbo库房.ListIndex >= 0 Then
        cbo库房.Tag = cbo库房.ListIndex
    End If

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intIndex As Integer
    
    With vsfList
        If .TextMatrix(.Row, .ColIndex("审核人")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    Select Case Control.Id
        ''''打印
        Case conMenu_File_PrintSet
            '打印设置
            zlPrintSet
        Case conMenu_File_Preview
            '打印预览
            subPrint intIndex, tbcDetail.Selected.Index, 2
        Case conMenu_File_Print
            '打印
            subPrint intIndex, tbcDetail.Selected.Index, 1
        Case conMenu_File_Excel
            '输出到Excel
            subPrint intIndex, tbcDetail.Selected.Index, 3

        ''''功能
        Case mconMenu_Edit_CA_VerifyAccount
            '审核结存
            Call MediAccountProcess_VerifyAccount
        Case mconMenu_Edit_CA_AddIniAccount
            '初始结存/初始化
            Call MediAccountProcess_AddIniAccount(0)
        Case mconMenu_Edit_CA_AddNewAccount
            '结存
            Call MediAccountProcess_AddIniAccount(1)
        Case mconMenu_Edit_CA_DeleteAccount
            '删除结存
            Call MediAccountProcess_DeleteAccount
        Case mconMenu_Edit_CA_CancelAccount
            '取消结存
            Call MediAccountProcess_CancleAccount
        ''''查看
        Case mconMenu_CA_Refresh '刷新
            Call GetAccountRecord
            Call RefreshList
        ''''帮助
        Case conMenu_Help_Help                         '帮助
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
        Case conMenu_Help_Web                          'WEB上的中联
        Case conMenu_Help_Web_Home                     '中联主页
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum                    '中联论坛
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail                     '发送反馈
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About                        '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)

        Case conMenu_File_Exit
            '退出
            Unload Me
        Case Else
            If Control.Id > 401 And Control.Id < 499 Then
                '执行自定义报表
                Call BillPrint_Custom(Control)
            End If
    End Select
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '打印自定义报表

    '默认参数：卫材=材料id，药房=药房id，结存ID=结存ID
    Dim strName As String
    Dim intType As Integer
    Dim lng结存id As Long
    Dim lng库房ID As Long

    strName = Split(Control.Parameter, ",")(1)

    If strName = "ZL" & glngSys \ 100 & "_INSIDE_1332" Then
        Call ReportOpen(gcnOracle, glngSys, strName, Me)
    Else
        If vsfList.Row <> 0 Then
            lng结存id = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("结存ID")))
            lng库房ID = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("库房ID")))
        End If

        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
            "结存id=" & lng结存id, _
            "库房id=" & lng库房ID, _
            "卫材=" & IIf(Val(txt明细卫材.Tag) = 0, "", Val(txt明细卫材.Tag)))
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
      
    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub



Private Sub Cmd药品_Click()
    Dim intIndex As Integer
    
    With vsfList
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("审核人")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    GetSelect ""

    DoEvents

    If tbcDetail.Selected.Index = mconTab_CA_Detail Then
        LoadInOutList intIndex, mlng结存ID
    ElseIf tbcDetail.Selected.Index = mconTab_CA_Stuff Then
        LoadDetailList intIndex, mlng结存ID
    ElseIf tbcDetail.Selected.Index = mconTab_CA_Mistake Then
        LoadMistakeList intIndex, mlng结存ID
    End If
End Sub


Private Sub GetSelect(ByVal strInput As String)
    Dim vRect As RECT
    Dim strReturn As String
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    
    vRect = zlControl.GetControlRect(txt明细卫材.hwnd)
    sngX = vRect.Left + picShowDetail.Left
    sngY = vRect.Top + picShowDetail.Top
    
    strReturn = SelectInput(strInput, sngX, sngY, sngH)
    
    If strReturn = "" Then Exit Sub
            
    txt明细卫材.Tag = Val(Split(strReturn, ";")(0))
    txt明细卫材.Text = Split(strReturn, ";")(1)
End Sub


Private Function SelectInput(ByVal strKey As String, ByVal sngX As Single, ByVal sngY As Single, ByVal sngH As Single) As String
    Dim strFindString As String
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    
    strKey = UCase(Trim(strKey))
    
    If strKey <> "" Then
        strFindString = " And (B.编码 Like [1] OR C.名称 Like [2] OR C.简码 LIKE [2])"
        
        If IsNumeric(strKey) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
             strFindString = " And (B.编码 Like [1] Or C.简码 Like [2] And C.码类=3)"
        ElseIf zlStr.IsCharAlpha(strKey) Then         '01,11.输入全是字母时只匹配简码
             strFindString = " And C.简码 Like [2] "
        ElseIf zlStr.IsCharChinese(strKey) Then
            strFindString = " And B.名称 Like [2] "
        End If
    End If
    
    gstrSQL = "Select Rownum As ID, 材料id, 卫材名称, 规格, 产地 as 生产商,散装单位, 包装单位, 包装系数 " & _
        " From (Select Distinct A.材料id, B.编码, '['||B.编码||']'|| B.名称 As 卫材名称, B.规格, B.产地," & _
        "         B.计算单位 As 散装单位, A.包装单位, A.换算系数 包装系数 " & _
        "       From 材料特性 A, " & _
        "      (Select B.ID, B.编码, B.名称, B.规格,B.产地,B.计算单位 From 收费项目目录 B, 收费项目别名 C " & _
        "       Where (B.站点 = [4] Or B.站点 is Null) And B.ID = C.收费细目id And B.类别 In ('4') " & strFindString & _
        ") B, 收费项目别名 C "
    
    If Val(cbo库房.ItemData(cbo库房.ListIndex)) > 0 Then
        gstrSQL = gstrSQL & ", 收费执行科室 D "
    End If
    
    gstrSQL = gstrSQL & " Where A.材料id = B.ID And A.材料id = C.收费细目id(+) And C.性质(+) = 3 "
    
    If Val(cbo库房.ItemData(cbo库房.ListIndex)) > 0 Then
        gstrSQL = gstrSQL & " And A.材料ID = D.收费细目ID And 执行科室ID = [3] "
    End If

    gstrSQL = gstrSQL & " Order By B.编码)"
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "药品选择器", False, "", "选择药品", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                    strKey & "%", "%" & strKey & "%", _
                    Val(cbo库房.ItemData(cbo库房.ListIndex)), gstrNodeNo)
    
    If blnCancel = True Then Exit Function
    
    If rsTemp Is Nothing Then
        SelectInput = ""
    Else
        SelectInput = rsTemp!材料ID & ";" & rsTemp!卫材名称 & ";" & rsTemp!包装单位 & "," & rsTemp!包装系数 & "|" & rsTemp!散装单位 & "," & "1"
    End If
End Function


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picMain.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Me.Width = mcstlngWinNormalWidth
    Me.Height = mcstlngWinNormalHeight
    
    mint结存方式 = Val(zlDatabase.GetPara(281, 100, , 0))
    mint结存时点 = Val(zlDatabase.GetPara(281, 100, , 0))
    mstr当前日期 = Format(sys.Currentdate, "yyyy-mm-dd")

    Call GetPrivs
    
    Call initGrid   '初始化列表 如合并信息
    Call InitDetailRec '构建结存明细记录集
    If GetStockName = False Then
        Unload Me
        Exit Sub
    End If
    Call IniDrugUnit '为单位下拉列表填充值
    
    Call InitComandBars
    Call InitPanes
    Call InitTabControl
  
    '    添加自定义报表
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    RestoreWinState Me, App.ProductName, mstrCaption
    
    
    mblnStart = True
    Call CheckClosAccount

    '载入数据
    Call GetAccountRecord
    Call RefreshList
    
    If mint结存时点 = 0 Then
        Me.Caption = "卫材结存管理(每月最后一天结存)"
    Else
        Me.Caption = "卫材结存管理(每月" & mint结存时点 & "日结存)"
    End If
End Sub


Private Sub initGrid()
    Const cstRowHeight = 300
    
    With vsfList
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .RowHeightMin = cstRowHeight
    End With


    With vsfDetail
        .Rows = 1
        .RowHeightMin = cstRowHeight
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With

    With vsfStuff
        .Rows = 2
        .RowHeightMin = cstRowHeight
        
        .Cell(flexcpFontBold, 0, 0, 1, .Cols - 1) = True
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
    End With
    
    With vsfMistake(0)
        .RowHeightMin = cstRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
    End With
    
    With vsfMistake(1)
        .RowHeightMin = cstRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
    End With
End Sub


Private Sub GetAccountRecord()
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrHandle

    gstrSQL = "Select A.ID, Nvl(A.库房id, 0) As 库房id, B.名称 As 库房名称, A.期初日期, A.期末日期, A.填制人" & _
              "   , A.填制日期, 审核人, 审核日期,取消人,取消日期, Nvl(A.上次结存id, 0) As 上次结存id,a.期间,a.性质 " & _
              "From 材料结存记录 A, 部门表 B " & _
              "Where A.库房id = B.ID(+) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "取结存记录")

    Call InitAccountRec

    With mrsAccount
        Do While Not rsTmp.EOF
            .AddNew
            !Id = rsTmp!Id
            !库房id = rsTmp!库房id
            !库房名称 = NVL(rsTmp!库房名称, "全院")
            !期初日期 = IIf(IsNull(rsTmp!期初日期), "", Format(rsTmp!期初日期, "YYYY-MM-DD HH:MM:SS"))
            !期末日期 = IIf(IsNull(rsTmp!期末日期), "", Format(rsTmp!期末日期, "YYYY-MM-DD HH:MM:SS"))
            !填制人 = NVL(rsTmp!填制人, "")
            !填制日期 = IIf(IsNull(rsTmp!填制日期), "", Format(rsTmp!填制日期, "YYYY-MM-DD HH:MM:SS"))
            !审核人 = NVL(rsTmp!审核人, "")
            !审核日期 = IIf(IsNull(rsTmp!审核日期), "", Format(rsTmp!审核日期, "YYYY-MM-DD HH:MM:SS"))
            !取消人 = NVL(rsTmp!取消人, "")
            !取消日期 = IIf(IsNull(rsTmp!取消日期), "", Format(rsTmp!取消日期, "YYYY-MM-DD HH:MM:SS"))
            !上次结存id = rsTmp!上次结存id
            !期间 = NVL(rsTmp!期间, "")
            !性质 = Val(NVL(rsTmp!性质, "1"))
            .Update

            rsTmp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshList()
    '刷新结存记录列表,为结存记录表中填充值
    Dim strFilter As String
    Dim Str期间 As String
    Dim strSQL As String
    
    Str期间 = Format(sys.Currentdate, "yyyyMM")

    mrsAccount.Filter = "库房id=" & Val(cbo库房.ItemData(cbo库房.ListIndex))
    mrsAccount.Sort = "审核日期 Desc"
    
    With vsfList
        .Redraw = flexRDNone
        .Rows = 1
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        
        Do While Not mrsAccount.EOF
            .Rows = .Rows + 1

            .TextMatrix(.Rows - 1, .ColIndex("结存ID")) = mrsAccount!Id
            .TextMatrix(.Rows - 1, .ColIndex("上次结存ID")) = mrsAccount!上次结存id
            .TextMatrix(.Rows - 1, .ColIndex("库房ID")) = NVL(mrsAccount!库房id, 0)

            .TextMatrix(.Rows - 1, .ColIndex("期初日期")) = Format(mrsAccount!期初日期, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("期末日期")) = Format(mrsAccount!期末日期, "YYYY-MM-DD HH:MM:SS")

            .TextMatrix(.Rows - 1, .ColIndex("填制人")) = mrsAccount!填制人
            .TextMatrix(.Rows - 1, .ColIndex("填制日期")) = Format(mrsAccount!填制日期, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("审核人")) = mrsAccount!审核人
            .TextMatrix(.Rows - 1, .ColIndex("审核日期")) = Format(mrsAccount!审核日期, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("取消人")) = mrsAccount!取消人
            .TextMatrix(.Rows - 1, .ColIndex("取消日期")) = Format(mrsAccount!取消日期, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("期间")) = mrsAccount!期间
            .TextMatrix(.Rows - 1, .ColIndex("性质")) = mrsAccount!性质
            
            If mrsAccount!性质 = 0 Then
                .Cell(flexcpPicture, .Rows - 1, .ColIndex("性质"), .Rows - 1, .ColIndex("性质")) = picIni.Picture
            End If
            
            '期初数据蓝色标识
            If mrsAccount!性质 = 0 Then
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
            End If
            
            '未审核数据用红色标识
            If Format(mrsAccount!审核日期, "YYYY-MM-DD HH:MM:SS") = "" Then
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
            End If
            
            mrsAccount.MoveNext
        Loop
                  
        If .Rows > 1 Then
            .Row = 1
        End If

        .Redraw = flexRDDirect
    End With
    
    stbThis.Panels(2).Text = ""
    If vsfList.Rows = 1 Then
        stbThis.Panels(2).Text = "[" & cbo库房.Text & "]" & "无期初结存数据，请通过盘点等方式确保当前库房数据正确。" & vbCrLf & "按结存可以手工产生初结存数据或在每月固定日期自动产生结存数据！"
    End If
End Sub


Private Function IsInString(ByVal strTarget As String, ByVal strOrigin As String, Optional strSplit As String = "") As Boolean
    '某个字符串是否包含另一个字符串
    'strTarget：目标字符串
    'strOrigin：原字符串
    'strSplit：分隔符（不为空时为精确匹配）
    '在strTarget中是否包含strOrigin
    
    IsInString = InStrB(strSplit & strTarget & strSplit, strSplit & strOrigin & strSplit) > 0
End Function



Private Sub InitComandBars()
    '初始化菜单：加载全部菜单，工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
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
    Me.cbsMain.Icons = Me.imgPublic.Icons
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsMain.ActiveMenuBar.Title = "菜单"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.Id = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.Id = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_AddIniAccount, "初始化(&N)")
        cbrControlMain.Visible = mPrives.bln审核
        '
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_AddNewAccount, "结存(A)")
        cbrControlMain.Enabled = IIf(mint结存方式 = -1, True, False)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_DeleteAccount, "删除(D)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_VerifyAccount, "审核(&V)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_CancelAccount, "取消(C)")
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.Id = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_CA_Refresh, "刷新(&R)")
    End With

    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.Id = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的中联")
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
        cbrControlMain.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_CA_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
  
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_AddIniAccount, "初始化")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_AddNewAccount, "结存")
        cbrControlMain.Enabled = IIf(mint结存方式 = -1, True, False)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_DeleteAccount, "删除")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_VerifyAccount, "审核")
        cbrControlMain.Visible = mPrives.bln审核
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_CancelAccount, "取消")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_CA_Refresh, "刷新")
        cbrControlMain.BeginGroup = True
         
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        cbrControlMain.BeginGroup = True
    End With
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitPanes()
    '初始化分栏控件
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane
    Dim objPaneList As Pane
    Dim objPaneDetail As Pane
    
    Set objPaneCon = Me.dkpMain.CreatePane(mconPane_Dept_Condition, 250, 100, DockLeftOf, Nothing)
    objPaneCon.Title = "结存明细"
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
'    objPaneCon.MaxTrackSize.SetSize 290, 500
End Sub


Private Sub InitTabControl()
    '初始化分页控件
    
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mconTab_CA_Detail, "结存明细", Me.picShowDetail.hwnd, 0).Tag = "结存明细_"
        .InsertItem(mconTab_CA_Stuff, "卫材明细", Me.picShowDetail.hwnd, 0).Tag = "卫材明细_"
        .InsertItem(mconTab_CA_Mistake, "误差明细", Me.picShowDetail.hwnd, 0).Tag = "误差明细_"
        
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < mcstlngWinNormalWidth Then Me.Width = mcstlngWinNormalWidth
    If Me.Height < mcstlngWinNormalHeight Then Me.Height = mcstlngWinNormalHeight
End Sub



Private Sub Form_Unload(Cancel As Integer)
    mblnStart = False
    
    SaveWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLine
'        .Top = 0
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLine.Left + 50
        .Width = picDetail.Width - fraLine.Width
        .Height = picDetail.Height - 50
    End With
End Sub



Private Sub picList_Resize()
    On Error Resume Next
    
    With vsfList
        .Move 0, 0, picList.Width, picList.Height
    End With
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    
    With cbo库房
        .Width = picMain.Width - .Left - 60
    End With

    With picList
        .Top = cbo库房.Top + cbo库房.Height + 120
        .Left = 0
        .Width = picMain.Width
        .Height = picMain.Height - .Top
    End With
End Sub


Private Sub picShowDetail_Resize()
    On Error Resume Next
    
    With vsfDetail
        .Top = txt明细卫材.Top + txt明细卫材.Height + 120
        .Left = 0
        .Width = picShowDetail.Width
        .Height = picShowDetail.Height - .Top
    End With
    
    
    With vsfStuff
        .Top = vsfDetail.Top
        .Left = 0
        .Width = vsfDetail.Width
        .Height = vsfDetail.Height
    End With
    
    With vsfMistake(0)
        .Top = vsfDetail.Top
        .Left = 0
        .Width = vsfDetail.Width
        .Height = vsfDetail.Height
    End With
    
    With vsfMistake(1)
        .Top = vsfDetail.Top
        .Left = 0
        .Width = vsfDetail.Width
        .Height = vsfDetail.Height
    End With
End Sub


Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intIndex As Integer
    
    vsfDetail.Visible = False
    vsfStuff.Visible = False
    vsfMistake(0).Visible = False
    vsfMistake(1).Visible = False
    
    With vsfList
        If .TextMatrix(.Row, .ColIndex("审核人")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    If Item.Index = mconTab_CA_Detail Then
        vsfDetail.Visible = True
        LoadInOutList intIndex, mlng结存ID
    ElseIf Item.Index = mconTab_CA_Stuff Then
        vsfStuff.Visible = True
        LoadDetailList intIndex, mlng结存ID
    ElseIf Item.Index = mconTab_CA_Mistake Then
         vsfMistake(1).Visible = True
        LoadMistakeList intIndex, mlng结存ID
    End If
End Sub


Private Function GetStockName() As Boolean
    '取当前操作员允许操作的库房
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle

    gstrSQL = "SELECT DISTINCT a.id, a.编码 || '-' || a.名称 as 名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = [2] Or a.站点 is Null) And c.工作性质 = b.名称 " _
            & "  AND Instr('WV',b.编码,1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
            & IIf(mPrives.bln所有库房 = True, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])") _
            & "Order by a.编码 || '-' || a.名称 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取操作员允许操作的库房", UserInfo.Id, gstrNodeNo)
    
    If rsTmp.EOF Then
        MsgBox "当前操作员不属于任何库房，不能进行结存操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With rsTmp
        cbo库房.Clear
        
        Do While Not .EOF
            cbo库房.AddItem !名称
            cbo库房.ItemData(cbo库房.NewIndex) = !Id
          
            .MoveNext
        Loop
        
        cbo库房.ListIndex = 0
    End With
    
    GetStockName = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub IniDrugUnit()
    '卫材使用的单位
    With cbo单位
        .Clear
        .AddItem "散装单位"
        .AddItem "包装单位"
        .ListIndex = 0
    End With
End Sub




Private Sub txt明细卫材_GotFocus()
    zlControl.TxtSelAll txt明细卫材
End Sub

Private Sub txt明细卫材_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub

    txt明细卫材_Validate True
End Sub

Private Sub txt明细卫材_KeyPress(KeyAscii As Integer)
    If InStr(" ';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt明细卫材_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(txt明细卫材.Text) = "" Then
        txt明细卫材.Tag = 0
    End If
End Sub

Private Sub txt明细卫材_Validate(Cancel As Boolean)
    Dim intIndex As Integer
    
    With vsfList
        If .TextMatrix(.Row, .ColIndex("审核人")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    With txt明细卫材
        If Trim(.Text) = "" Then
            .Tag = 0
        Else
            GetSelect .Text
        End If

        DoEvents

        If tbcDetail.Selected.Index = mconTab_CA_Detail Then
            LoadInOutList intIndex, mlng结存ID
        ElseIf tbcDetail.Selected.Index = mconTab_CA_Stuff Then
            LoadDetailList intIndex, mlng结存ID
        ElseIf tbcDetail.Selected.Index = mconTab_CA_Mistake Then
            LoadMistakeList intIndex, mlng结存ID
        End If
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim i As Integer
    Dim intIndex As Integer
    
    With vsfList
        If Val(.Tag) = .Row Then Exit Sub

        .Tag = .Row

        If .Row <= vsfList.FixedRows - 1 Then
            ClearDetailList
            ClearStuffList
            ClearMistakeList
            Exit Sub
        End If

        If Val(.TextMatrix(.Row, .ColIndex("结存ID"))) = 0 Then
            ClearDetailList
            ClearStuffList
            ClearMistakeList
            Exit Sub
        End If
        
        With vsfList
            If .TextMatrix(.Row, .ColIndex("审核人")) = "" Then
                intIndex = 0
            Else
                intIndex = 1
              End If
        End With
        
        Call LoadData(intIndex)
        
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        .Redraw = flexRDDirect
    End With
 End Sub


Private Sub ClearDetailList()
    vsfDetail.Rows = 1
    vsfDetail.Rows = 2
End Sub

Private Sub ClearStuffList()
    vsfStuff.Rows = 2
    vsfStuff.Rows = 3
End Sub

Private Sub ClearMistakeList()
    vsfMistake(0).Rows = 1
    vsfMistake(0).Rows = 2

    vsfMistake(1).Rows = 1
    vsfMistake(1).Rows = 2
End Sub

Private Sub LoadData(ByVal intType As Integer)
    '提取数据
    Dim lng上次结存ID As Long
    Dim str期初日期 As String
    Dim str期末日期 As String
    Dim rsTemp As ADODB.Recordset

    With vsfList
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("结存ID"))) = 0 Then Exit Sub

        mlng结存ID = Val(.TextMatrix(.Row, .ColIndex("结存ID")))
        lng上次结存ID = Val(.TextMatrix(.Row, .ColIndex("上次结存ID")))
        gstrSQL = "Select 期初日期, 期末日期 From 材料结存记录 Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "日期查询", mlng结存ID)
        
        If rsTemp.RecordCount <= 0 Then '对应结存记录已被删除
            MsgBox "该记录已经被删除！", vbInformation, gstrSysName
            GetAccountRecord
            RefreshList
            Exit Sub
        End If
        
        str期初日期 = IIf(IsNull(rsTemp!期初日期), "", rsTemp!期初日期)
        str期末日期 = IIf(IsNull(rsTemp!期末日期), "", rsTemp!期末日期)
    End With

    Call AviShow(Me)

    On Error GoTo ErrHandle

    GetDetailRecord intType, mlng结存ID, str期初日期, str期末日期

    vsfDetail.Visible = False
    vsfStuff.Visible = False
    vsfMistake(0).Visible = False
    vsfMistake(1).Visible = False

    If tbcDetail.Selected.Index = mconTab_CA_Detail Then
        LoadInOutList intType, mlng结存ID
    ElseIf tbcDetail.Selected.Index = mconTab_CA_Stuff Then
        LoadDetailList intType, mlng结存ID
    ElseIf tbcDetail.Selected.Index = mconTab_CA_Mistake Then
        LoadMistakeList intType, mlng结存ID
    End If

    Call AviShow(Me, False)

    Exit Sub
ErrHandle:
    Call AviShow(Me, False)
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadInOutList(ByVal intType As Integer, ByVal lng结存id As Long)
     '卫材入出汇总
    Dim lng库房ID As Long
    Dim lng材料ID As Long
    Dim blnShowSubType As Boolean
    Dim strFilter As String
    Dim strOrder As String
    
    Dim str业务大类 As String
    Dim str业务分类 As String
    
    Dim dbl数量 As Double
    Dim Dbl金额 As Double
    Dim dbl差价 As Double
    
    Dim intRow As Integer
    
    
    Dim strTmp As String
    Dim str单位 As String
    Dim dbl包装 As String
    
    Dim intShowNumberDigit As Integer          '数量小数位数
    Dim intShowMoneyDigit As Integer           '金额小数位数
    Dim intUnit As Integer  '1-散装;2-包装
    
    Call ClearDetailList
    
    '取价格，数量，金额的显示精度intUnit= 1-散装
    If cbo单位.ListIndex = 0 Then
        intUnit = 1
    ElseIf cbo单位.ListIndex = 1 Then
        intUnit = 2
    End If
            
    intShowNumberDigit = newGetDigit(0, 3, intUnit)
    intShowMoneyDigit = newGetDigit(0, 4)
    
    lng库房ID = Val(cbo库房.ItemData(cbo库房.ListIndex))
    lng材料ID = Val(txt明细卫材.Tag)
    
    strFilter = "类型=" & intType & " And 结存ID=" & lng结存id
    If lng库房ID > 0 Then strFilter = strFilter & " And 库房ID=" & lng库房ID
    If lng材料ID > 0 Then strFilter = strFilter & " And 材料ID=" & lng材料ID
    
    strOrder = "业务大类,业务分类"
    
    vsfDetail.Visible = True
    
    
    With vsfDetail
        mrsDetail.Filter = strFilter
        mrsDetail.Sort = strOrder
        If mrsDetail.RecordCount = 0 Then
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        .Rows = 1
        
        Do While Not mrsDetail.EOF
            If cbo单位.ListIndex = 0 Then
                str单位 = mrsDetail!散装单位
                dbl包装 = mrsDetail!散装包装
            ElseIf cbo单位.ListIndex = 1 Then
                str单位 = mrsDetail!包装单位
                dbl包装 = mrsDetail!包装包装
            End If
            
            If str业务大类 & str业务分类 <> mrsDetail!业务大类 & mrsDetail!业务分类 Then
                .Rows = .Rows + 1
                intRow = .Rows - 1
                
                str业务大类 = mrsDetail!业务大类
                str业务分类 = mrsDetail!业务分类
            End If
            
            .TextMatrix(intRow, .ColIndex("业务大类")) = mrsDetail!业务大类
            .TextMatrix(intRow, .ColIndex("业务分类")) = mrsDetail!业务分类
            
            
            If mrsDetail!数量 <> 0 Then .TextMatrix(intRow, .ColIndex("数量")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("数量"))) + mrsDetail!数量 / dbl包装, intShowNumberDigit, , True)
            If mrsDetail!金额 <> 0 Then .TextMatrix(intRow, .ColIndex("售价金额")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("售价金额"))) + mrsDetail!金额, intShowMoneyDigit, , True)
            If mrsDetail!差价 <> 0 Then .TextMatrix(intRow, .ColIndex("差价")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("差价"))) + mrsDetail!差价, intShowMoneyDigit, , True)
            .TextMatrix(intRow, .ColIndex("成本金额")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("售价金额"))) - Val(.TextMatrix(intRow, .ColIndex("差价"))), intShowMoneyDigit, , True)
            
            
            mrsDetail.MoveNext
        Loop
        
        '指定卫材对数量处理
        If lng材料ID > 0 Then .TextMatrix(0, .ColIndex("数量")) = "数量(" & str单位 & ")"
        .ColHidden(.ColIndex("数量")) = Not (lng材料ID > 0)
        If .ColHidden(.ColIndex("数量")) Then
            .ColWidth(.ColIndex("数量")) = 0
        Else
            .ColWidth(.ColIndex("数量")) = 1500
        End If

        '出库标记为红色；入库标记为蓝色
        For intRow = 2 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("业务大类")) = "3-出库" Then
                .Cell(flexcpForeColor, intRow, .ColIndex("数量"), intRow, .ColIndex("成本金额")) = vbRed
            ElseIf .TextMatrix(intRow, .ColIndex("业务大类")) = "2-入库" Then
                .Cell(flexcpForeColor, intRow, .ColIndex("数量"), intRow, .ColIndex("成本金额")) = vbBlue
            End If
        Next

        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub LoadDetailList(ByVal intType As Integer, ByVal lng结存id As Long)
    '卫材入出明细
    Dim lng库房ID As Long
    Dim lng材料ID As Long
    Dim blnShowSubType As Boolean
    Dim strFilter As String
    Dim strOrder As String
    
    Dim str业务大类 As String
    Dim str业务分类 As String
    
    Dim str卫材名称 As String
    
    Dim dbl数量 As Double
    Dim Dbl金额 As Double
    Dim dbl差价 As Double
    
    Dim lngRow As Long
    
    Dim strTmp As String
    Dim str单位 As String
    Dim dbl包装 As String
    
    Dim intShowNumberDigit As Integer          '数量小数位数
    Dim intShowMoneyDigit As Integer           '金额小数位数
    Dim intUnit As Integer  '1-散装;2-包装;
    
    ClearStuffList
    
    vsfStuff.Visible = True
    
    '取价格，数量，金额的显示精度
    If cbo单位.ListIndex = 0 Then
        intUnit = 1
    ElseIf cbo单位.ListIndex = 1 Then
        intUnit = 2
    End If
            
    intShowNumberDigit = newGetDigit(0, 3, intUnit)
    intShowMoneyDigit = newGetDigit(0, 4)
    
    lng库房ID = Val(cbo库房.ItemData(cbo库房.ListIndex))
    lng材料ID = Val(txt明细卫材.Tag)
    
    strFilter = "类型=" & intType & " And 结存ID=" & lng结存id
    If lng库房ID > 0 Then strFilter = strFilter & " And 库房ID=" & lng库房ID
    If lng材料ID > 0 Then strFilter = strFilter & " And 材料ID=" & lng材料ID
    
    strOrder = "卫材名称,业务大类,业务分类"
    
    With vsfStuff
        mrsDetail.Filter = strFilter
        mrsDetail.Sort = strOrder
        If mrsDetail.RecordCount = 0 Then
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        .Rows = 2
        
        Do While Not mrsDetail.EOF
            If cbo单位.ListIndex = 0 Then
                str单位 = mrsDetail!散装单位
                dbl包装 = mrsDetail!散装包装
            ElseIf cbo单位.ListIndex = 1 Then
                str单位 = mrsDetail!包装单位
                dbl包装 = mrsDetail!包装包装
            End If
            
            If lng库房ID = 0 And (mrsDetail!业务分类 = "卫材库房入库" Or mrsDetail!业务分类 = "卫材库房出库") Then
                '统计全院时，不计算内部流通（移库）
            Else
                If str卫材名称 <> mrsDetail!卫材名称 Then
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                   
                    str卫材名称 = mrsDetail!卫材名称
                End If
                
                .TextMatrix(lngRow, .ColIndex("卫材名称")) = mrsDetail!卫材名称
                .TextMatrix(lngRow, .ColIndex("商品名")) = NVL(mrsDetail!商品名)
                .TextMatrix(lngRow, .ColIndex("规格")) = mrsDetail!规格
                .TextMatrix(lngRow, .ColIndex("单位")) = str单位
                
                If mrsDetail!业务大类 = "1-期初" Then
                    .TextMatrix(lngRow, .ColIndex("期初数量")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("期初数量"))) + mrsDetail!数量 / dbl包装, intShowNumberDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("期初金额")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("期初金额"))) + mrsDetail!金额, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("期初差价")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("期初差价"))) + mrsDetail!差价, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("期初成本")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("期初金额"))) - Val(.TextMatrix(lngRow, .ColIndex("期初差价"))), intShowMoneyDigit, , True)
                ElseIf mrsDetail!业务大类 = "2-入库" Then
                    .TextMatrix(lngRow, .ColIndex("入库数量")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("入库数量"))) + mrsDetail!数量 / dbl包装, intShowNumberDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("入库金额")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("入库金额"))) + mrsDetail!金额, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("入库差价")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("入库差价"))) + mrsDetail!差价, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("入库成本")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("入库金额"))) - Val(.TextMatrix(lngRow, .ColIndex("入库差价"))), intShowMoneyDigit, , True)
                ElseIf mrsDetail!业务大类 = "3-出库" Then
                    .TextMatrix(lngRow, .ColIndex("出库数量")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("出库数量"))) + mrsDetail!数量 / dbl包装, intShowNumberDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("出库金额")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("出库金额"))) + mrsDetail!金额, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("出库差价")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("出库差价"))) + mrsDetail!差价, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("出库成本")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("出库金额"))) - Val(.TextMatrix(lngRow, .ColIndex("出库差价"))), intShowMoneyDigit, , True)
                ElseIf mrsDetail!业务大类 = "4-期末" Then
                    .TextMatrix(lngRow, .ColIndex("期末数量")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("期末数量"))) + mrsDetail!数量 / dbl包装, intShowNumberDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("期末金额")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("期末金额"))) + mrsDetail!金额, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("期末差价")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("期末差价"))) + mrsDetail!差价, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("期末成本")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("期末金额"))) - Val(.TextMatrix(lngRow, .ColIndex("期末差价"))), intShowMoneyDigit, , True)
                End If
            End If
            
            mrsDetail.MoveNext
        Loop
                
        '出库标记为红色；入库标记为蓝色
        .Cell(flexcpForeColor, 2, .ColIndex("入库数量"), .Rows - 1, .ColIndex("入库成本")) = vbBlue
        .Cell(flexcpForeColor, 2, .ColIndex("出库数量"), .Rows - 1, .ColIndex("出库成本")) = vbRed
       
        '合计粗体表示
'        .Cell(flexcpFontBold, 2, .ColIndex("期末数量"), .rows - 1, .ColIndex("期末成本")) = True
            
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub LoadMistakeList(ByVal intType As Integer, ByVal lng结存id As Long)
    Dim lng库房ID As Long
    Dim lng材料ID As Long
    Dim strFilter As String
    Dim strOrder As String
    
    Dim str库房 As String
    
    Dim dbl数量差 As Double
    Dim dbl金额差 As Double
    Dim dbl差价差 As Double
    
    Dim intRow As Integer
    Dim strUnit As String
    
    Dim intShowNumberDigit As Integer          '数量小数位数
    Dim intShowMoneyDigit As Integer           '金额小数位数
    Dim intUnit As Integer  '1-售价;2-门诊;3-住院;4-药库;
    
    On Error GoTo ErrHandle

    Call ClearMistakeList
    
    '取价格，数量，金额的显示精度
    If cbo单位.ListIndex = 0 Then
        intUnit = 1
    ElseIf cbo单位.ListIndex = 1 Then
        intUnit = 2
    End If
            
    intShowNumberDigit = newGetDigit(0, 3, intUnit)
    intShowMoneyDigit = newGetDigit(0, 4)
    
    GetMistakeRecord intType, lng结存id
    
    lng库房ID = Val(cbo库房.ItemData(cbo库房.ListIndex))
    lng材料ID = Val(txt明细卫材.Tag)
    
    strFilter = "类型=" & intType & " And 结存ID=" & lng结存id
    If lng库房ID > 0 Then strFilter = strFilter & " And 库房ID=" & lng库房ID
    If lng材料ID > 0 Then strFilter = strFilter & " And 材料ID=" & lng材料ID
    
    strOrder = "卫材名称,批次"
    If lng库房ID > 0 Then strOrder = "库房名称"
    
    If lng库房ID > 0 Then
        vsfMistake(1).Visible = True
        vsfMistake(0).Visible = False
    Else
        vsfMistake(1).Visible = False
        vsfMistake(0).Visible = True
    End If
    
    mrsMistake.Filter = strFilter
    mrsMistake.Sort = strOrder
    If mrsMistake.RecordCount = 0 Then Exit Sub
            
    If lng库房ID > 0 Then
        With vsfMistake(1)
            .Redraw = flexRDNone
            
            .Rows = 1
        
            Do While Not mrsMistake.EOF
                .Rows = .Rows + 1
                intRow = .Rows - 1
                
                .TextMatrix(intRow, .ColIndex("卫材")) = mrsMistake!卫材名称
                .TextMatrix(intRow, .ColIndex("商品名")) = NVL(mrsMistake!商品名, "")
                .TextMatrix(intRow, .ColIndex("规格")) = mrsMistake!规格
                .TextMatrix(intRow, .ColIndex("批次")) = mrsMistake!批次
                .TextMatrix(intRow, .ColIndex("金额差")) = zlStr.FormatEx(mrsMistake!金额差, intShowMoneyDigit, , True)
                .TextMatrix(intRow, .ColIndex("差价差")) = zlStr.FormatEx(mrsMistake!差价差, intShowMoneyDigit, , True)
                
                Select Case intUnit
                Case 2  '"包装单位"
                    .TextMatrix(intRow, .ColIndex("单位")) = mrsMistake!包装单位
                    .TextMatrix(intRow, .ColIndex("数量差")) = zlStr.FormatEx(mrsMistake!数量差 / mrsMistake!包装包装, intShowNumberDigit, , True)
                
                Case Else
                    .TextMatrix(intRow, .ColIndex("单位")) = mrsMistake!计算单位
                    .TextMatrix(intRow, .ColIndex("数量差")) = zlStr.FormatEx(mrsMistake!数量差, intShowNumberDigit, , True)
                End Select
                
                mrsMistake.MoveNext
            Loop
            
            .Redraw = flexRDDirect
        End With
    Else
        With vsfMistake(0)
            .Redraw = flexRDNone
            
            .Rows = 1
            
            Do While Not mrsMistake.EOF
                If str库房 <> mrsMistake!库房名称 Then
                    .Rows = .Rows + 1
                    intRow = .Rows - 1
                    
                    str库房 = mrsMistake!库房名称
                End If
                
                .TextMatrix(intRow, .ColIndex("库房")) = mrsMistake!库房名称
                
                Select Case intUnit
                Case 2  '"包装单位"
                    .TextMatrix(intRow, .ColIndex("数量差")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("数量差"))) + mrsMistake!数量差 / mrsMistake!包装包装, intShowNumberDigit, , True)
                Case Else
                  .TextMatrix(intRow, .ColIndex("数量差")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("数量差"))) + mrsMistake!数量差, intShowNumberDigit, , True)
                End Select
                
                .TextMatrix(intRow, .ColIndex("金额差")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("金额差"))) + mrsMistake!金额差, intShowMoneyDigit, , True)
                .TextMatrix(intRow, .ColIndex("差价差")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("差价差"))) + mrsMistake!差价差, intShowMoneyDigit, , True)
                
                mrsMistake.MoveNext
            Loop
            
            .Redraw = flexRDDirect
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'
Private Sub GetMistakeRecord(ByVal intType As Integer, ByVal lng结存id As Long)
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    '判断是否已保存该结存误差记录
    mrsMistake.Filter = "类型=" & intType & " And 结存ID=" & lng结存id
    If mrsMistake.RecordCount > 0 Then Exit Sub

    mrsMistake.Filter = ""

    '没找到时从数据库读取
    '[' || B.编码 || ']' As 编码, B.名称 As 通用名, E.名称 As 商品名
    gstrSQL = "Select Distinct A.结存id, A.库房id, A.材料id, Nvl(A.批次, 0) 批次, A.数量差, A.金额差, A.差价差, " & _
        " '[' || F.编码 || ']' As 编码, F.名称 As 通用名, E.名称 As 商品名, F.规格, D.名称 As 库房名称, F.计算单位, " & _
        " B.包装单位, B.换算系数 包装包装 " & _
        " From 材料结存误差 A, 材料特性 B, 收费项目目录 F, 收费项目别名 E, 部门表 D " & _
        " Where A.材料id = B.材料id And B.材料id = F.ID And A.库房id = D.ID And B.材料id = E.收费细目id(+) And " & _
        " E.性质(+) = 3 And A.结存id = [1] "

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "提取结存明细记录", lng结存id)

    '更新结存明细数据集
    With rsTmp
        Do While Not .EOF
            mrsMistake.AddNew
            mrsMistake!类型 = intType
            mrsMistake!结存ID = lng结存id
            mrsMistake!库房id = !库房id
            mrsMistake!材料ID = !材料ID
            mrsMistake!批次 = !批次
            mrsMistake!库房名称 = !库房名称
            
            mrsMistake!卫材名称 = !编码 & !通用名

            mrsMistake!商品名 = NVL(!商品名, "")
            mrsMistake!规格 = NVL(!规格, "")
            mrsMistake!数量差 = NVL(!数量差, 0)
            mrsMistake!金额差 = NVL(!金额差, 0)
            mrsMistake!差价差 = NVL(!差价差, 0)
            mrsMistake!计算单位 = !计算单位
            mrsMistake!包装单位 = !包装单位
            mrsMistake!包装包装 = !包装包装
            mrsMistake.Update

            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub





Private Sub GetDetailRecord(ByVal intType As Integer, ByVal lng结存id As Long, ByVal str期初日期 As String, ByVal str期末日期 As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSqlUnit As String
    Dim strSqlGroup As String

    On Error GoTo ErrHandle
    '判断是否已记录该结存明细
    mrsDetail.Filter = "类型=" & intType & " And 结存ID=" & lng结存id
    If mrsDetail.RecordCount > 0 Then Exit Sub

    mrsDetail.Filter = ""

    ''''没找到时从数据库读取
    gstrSQL = "Select Distinct  A.*, E.名称 As 商品名 From ("

    '取上期结存的期末数据作为本期的期初数据
    gstrSQL = gstrSQL & "Select A.库房id, A.业务大类, A.业务分类, '[' || B.编码 || ']' As 编码, B.名称 As 通用名, B.规格, A.材料id, Sum(A.数量) As 数量, Sum(A.金额) As 金额, Sum(A.差价) As 差价,C.包装单位,C.换算系数 包装包装,B.计算单位 as 散装单位,1 as 散装包装 " & _
        " From (Select A.库房id, '1-期初' As 业务大类, '' As 业务分类, A.材料id As 材料id, Sum(A.期初数量) As 数量, Sum(A.期初金额) As 金额, Sum(A.期初差价) As 差价 " & _
        "       From 材料结存明细 A Where 结存id = [1] " & _
        "       Group By A.库房id, A.材料id "

    '取期间发生额
    '注意用单据类型或库房的工作性质来确保只统计药品数据
    If str期初日期 <> "" Then
        gstrSQL = gstrSQL & _
        "     Union All " & _
        "     Select * From " & _
        "       (Select A.库房id, Decode(B.系数, 1, '2-入库', '3-出库') As 业务大类, B.名称 As 业务分类, A.药品id As 药品id, Sum(Nvl(A.实际数量, 0) * Nvl(A.付数, 1)) As 数量, Sum(Nvl(A.零售金额, 0)) As 金额, Sum(Nvl(A.差价, 0)) As 差价 " & _
        "       From 药品收发记录 A, 药品入出类别 B " & _
        "       Where A.入出类别id = B.ID And A.单据 In (15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 26) And " & _
        "       A.审核日期 Between [2] And [3] " & _
        "       Group By A.库房id, A.药品id, B.名称, Decode(B.系数, 1, '2-入库', '3-出库')) X " & _
        "     where X.数量 <> 0 Or X.金额 <> 0 Or X.差价 <> 0"
    End If

    '取本期期末数据
    gstrSQL = gstrSQL & _
        "       Union All " & _
        "       Select A.库房id, '4-期末' As 业务大类, '' As 业务分类, A.材料id, Sum(A.期末数量) As 数量, Sum(A.期末金额) As 金额, Sum(A.期末差价) As 差价 " & _
        "       From 材料结存明细 A " & _
        "       Where 结存id = [1] " & _
        "       Group By A.库房id, A.材料id) A, 收费项目目录 B, 材料特性 C " & _
        " Where A.材料id = B.Id And A.材料ID = C.材料ID " & _
        " Group By A.业务大类, A.业务分类, A.库房id, '[' || B.编码 || ']' , B.名称, B.规格, A.材料id,C.包装单位,C.换算系数,B.计算单位 "

    gstrSQL = gstrSQL & ") A, 收费项目别名 E " & _
        " Where A.材料id = E.收费细目id(+) And E.性质(+) = 3 " & _
        " Order By A.业务大类, A.业务分类, A.库房id, A.编码, A.通用名, E.名称, A.规格, A.材料id"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "提取结存明细记录", lng结存id, CDate(IIf(str期初日期 = "", "1990-01-01", str期初日期)), CDate(str期末日期))

    '更新结存明细数据集
    With rsTmp
        Do While Not .EOF
            mrsDetail.AddNew
            mrsDetail!类型 = intType
            mrsDetail!结存ID = lng结存id
            mrsDetail!业务大类 = NVL(!业务大类, "")
            mrsDetail!业务分类 = NVL(!业务分类, "")
            mrsDetail!库房id = !库房id
            
            mrsDetail!卫材名称 = !编码 & !通用名

            mrsDetail!商品名 = NVL(!商品名, "")
            mrsDetail!规格 = NVL(!规格, "")
            mrsDetail!材料ID = !材料ID
            mrsDetail!数量 = NVL(!数量, 0)
            mrsDetail!金额 = NVL(!金额, 0)
            mrsDetail!差价 = NVL(!差价, 0)
            mrsDetail!散装单位 = !散装单位
            mrsDetail!散装包装 = !散装包装
            mrsDetail!包装单位 = !包装单位
            mrsDetail!包装包装 = !包装包装
            mrsDetail.Update

            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub InitDetailRec()
    '构建结存明细记录集
    Set mrsDetail = New ADODB.Recordset
    With mrsDetail
        If .State = 1 Then .Close
        
        .Fields.Append "类型", adDouble, 1, adFldIsNullable
        .Fields.Append "结存ID", adDouble, 18, adFldIsNullable
        .Fields.Append "业务大类", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "业务分类", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "库房ID", adDouble, 18, adFldIsNullable
        .Fields.Append "卫材名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .Fields.Append "差价", adDouble, 18, adFldIsNullable
        .Fields.Append "散装单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "散装包装", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "包装单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "包装包装", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '误差记录集
    Set mrsMistake = New ADODB.Recordset
    With mrsMistake
        If .State = 1 Then .Close
        
        .Fields.Append "类型", adDouble, 1, adFldIsNullable
        .Fields.Append "结存ID", adDouble, 18, adFldIsNullable
        .Fields.Append "库房ID", adDouble, 18, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "库房名称", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "卫材名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量差", adDouble, 18, adFldIsNullable
        .Fields.Append "金额差", adDouble, 18, adFldIsNullable
        .Fields.Append "差价差", adDouble, 18, adFldIsNullable
        .Fields.Append "计算单位", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "散装单位", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "包装单位", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "散装包装", adDouble, 10, adFldIsNullable
        .Fields.Append "包装包装", adDouble, 10, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub


Private Sub CheckClosAccount()
    '当有数据时则说明已经初始化了，程序控制只能初始化一次
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim cbrControlAdd As CommandBarControl
    Dim cbrMenuAdd As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    '初始化
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_AddIniAccount, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_AddIniAccount, , True)
    '结存 未初始化也不能做结存功能，只有做了初始化后才能做结存
    Set cbrControlAdd = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_AddNewAccount, , True)
    Set cbrMenuAdd = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_AddNewAccount, , True)
    
    gstrSQL = "select 1 from 材料结存记录 where 库房id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否初始化", cbo库房.ItemData(cbo库房.ListIndex))
    
    If mPrives.bln初始化 = True Then
        If rsTemp.RecordCount > 0 Then
            cbrMenu.Enabled = False
            cbrControl.Enabled = False
            If mint结存方式 = -1 Then
                cbrControlAdd.Enabled = True
                cbrMenuAdd.Enabled = True
            End If
        Else
            cbrMenu.Enabled = True
            cbrControl.Enabled = True
            cbrControlAdd.Enabled = False
            cbrMenuAdd.Enabled = False
        End If
    Else
        cbrMenu.Visible = False
        cbrControl.Visible = False
               
        If rsTemp.RecordCount > 0 Then
            If mint结存方式 = -1 Then
                cbrControlAdd.Enabled = True
                cbrMenuAdd.Enabled = True
            End If
        Else
            cbrControlAdd.Enabled = False
            cbrMenuAdd.Enabled = False
        End If
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub GetPrivs()
    '权限
    mstrPrivs = GetPrivFunc(glngSys, 1729)
    
    With mPrives
        .bln所有库房 = IsInString(mstrPrivs, "所有库房", ";")
        .bln初始化 = IsInString(mstrPrivs, "初始结存", ";")
        .bln审核 = IsInString(mstrPrivs, "审核", ";")
    End With

End Sub


Private Sub InitAccountRec()
    '结存记录记录集
    Set mrsAccount = New ADODB.Recordset
    With mrsAccount
        If .State = 1 Then .Close
        
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "库房ID", adDouble, 18, adFldIsNullable
        .Fields.Append "库房名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "期初日期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "期末日期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "填制人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "填制日期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "审核人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "审核日期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "取消人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "取消日期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "上次结存ID", adDouble, 10, adFldIsNullable
        .Fields.Append "期间", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "性质", adLongVarChar, 2, adFldIsNullable
                
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub subPrint(ByVal intListIndex As Integer, ByVal intDetailindex As Integer, bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    Dim str期初日期 As String
    Dim str期末日期 As String

    With vsfList
        If .Row = 0 Then Exit Sub

        str期初日期 = "期初日期：" & IIf(.TextMatrix(.Row, .ColIndex("期初日期")) = "", "(初始结存)", .TextMatrix(.Row, .ColIndex("期初日期")))
        str期末日期 = "期末日期：" & .TextMatrix(.Row, .ColIndex("期末日期"))
    End With

    str期初日期 = Format(str期初日期, "yyyy-mm-dd")
    str期末日期 = Format(str期末日期, "yyyy-mm-dd")

    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True

    If intDetailindex = mconTab_CA_Detail Then
        objPrint.Title.Text = "卫材结存汇总"
    ElseIf intDetailindex = mconTab_CA_Stuff Then
        objPrint.Title.Text = "卫材结存明细"
    ElseIf intDetailindex = mconTab_CA_Mistake Then
        objPrint.Title.Text = "卫材结存误差"
    End If

    objRow.Add str期初日期 & "   " & str期末日期
    objRow.Add "库房：" & cbo库房.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow

    If intDetailindex = mconTab_CA_Detail Then
        Set objPrint.Body = vsfDetail
    ElseIf intDetailindex = mconTab_CA_Stuff Then
        Set objPrint.Body = vsfStuff
    ElseIf intDetailindex = mconTab_CA_Mistake Then
        If vsfMistake(0).Visible Then
            Set objPrint.Body = vsfMistake(0)
        Else
            Set objPrint.Body = vsfMistake(1)
        End If
    End If

    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
              zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub


Private Sub MediAccountProcess_AddIniAccount(ByVal int结存方式 As Integer)
    '新增初初始化
    'int结存方式 0-初始化结存 1-结存
    Dim lng库房ID As Long
    Dim rsData As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    On Error GoTo ErrHandle
    
    '只有初始化才提示，结存不提示
    If int结存方式 = 0 Then
        If MsgBox("提示：初始化将以当前库存数据作为初始结存数据，请先通过盘点确保当前库存数据正确。" & vbCrLf & vbCrLf & "是否现在进行初始化？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    gstrSQL = "Select 1 From 药品收发记录 Where 单据 = 22 And 库房id = [1] And 审核日期 Is Null And Rownum = 1 "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "IsAccountTime", Val(cbo库房.ItemData(cbo库房.ListIndex)))
    
    If Not rsData.EOF Then
        MsgBox "[" & cbo库房.Text & "]" & "还有盘点单据未审核，请审核后再进行本次" & IIf(int结存方式 = 1, "结存！", "初始化！"), vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrSQL = "select 1 记录 from 材料结存记录 where 库房id=[1] and 审核日期 is null and rownum<=1"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "未审核结存", Val(cbo库房.ItemData(cbo库房.ListIndex)))
    
    If Not rsData.EOF Then
        MsgBox "[" & cbo库房.Text & "]" & "还有结存单据未审核，请审核后再进行本次" & IIf(int结存方式 = 1, "结存！", "初始化！"), vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call AviShow(Me)
    
    lng库房ID = Val(cbo库房.ItemData(cbo库房.ListIndex))
    
    gstrSQL = "Zl_材料结存记录_Insert("
    'lng库房ID
    gstrSQL = gstrSQL & IIf(lng库房ID = 0, "Null", lng库房ID)
    '填制人
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    '性质
    gstrSQL = gstrSQL & "," & int结存方式
    gstrSQL = gstrSQL & ")"

    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "结存")

    Call GetAccountRecord
    Call RefreshList
            
    If mblnStart = True Then
        Call CheckClosAccount
    End If

    Call AviShow(Me, False)

    Exit Sub
ErrHandle:
    Call AviShow(Me, False)
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub MediAccountProcess_VerifyAccount()
    '审核结存
    Dim lng结存id As Long
    
    On Error GoTo ErrHandle
    
    With vsfList
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("审核日期")) <> "" Then Exit Sub
        
        lng结存id = Val(.TextMatrix(.Row, .ColIndex("结存ID")))
    End With
    
    If lng结存id = 0 Then Exit Sub

    gstrSQL = "Zl_材料结存记录_Verify("
    '结存ID
    gstrSQL = gstrSQL & lng结存id
    '审核人
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    gstrSQL = gstrSQL & ")"

    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "审核结存记录")

    Call GetAccountRecord
    Call RefreshList
    
    MsgBox "结存审核完毕，请查看。", vbInformation, gstrSysName
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub MediAccountProcess_CancleAccount()
    '取消结存单据，只能从最大开始取消，中途单据不能取消
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    With vsfList
        If .TextMatrix(.Row, .ColIndex("审核日期")) <> "" And .TextMatrix(.Row, .ColIndex("取消日期")) = "" Then
            gstrSQL = "Select Max(期末日期) as 日期 From 材料结存记录 Where 库房id = [1] And 审核日期 Is Not Null And 取消人 Is Null"

            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结存最大查询", Val(.TextMatrix(.Row, .ColIndex("库房id"))))
            
            If rsTemp.RecordCount > 0 Then
                If rsTemp!日期 = CDate(.TextMatrix(.Row, .ColIndex("期末日期"))) Then
                    gstrSQL = "Zl_材料结存记录_Cancel("
                    '结存id
                    gstrSQL = gstrSQL & .TextMatrix(.Row, .ColIndex("结存id"))
                    '取消人
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "结存取消")
                    
                    Call GetAccountRecord
                    Call RefreshList
                Else
                    MsgBox "请从该库房最近一次结存记录取消，最近一次结存记录期末日期为：(" & Format(rsTemp!日期, "YYYY-MM-DD HH:MM:SS") & ")", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub MediAccountProcess_DeleteAccount()

    On Error GoTo ErrHandle
    
    '删除结存单据
    With vsfList
        If .TextMatrix(.Row, .ColIndex("审核日期")) = "" Then
            gstrSQL = "Zl_材料结存记录_Delete("
            '结存id
            gstrSQL = gstrSQL & .TextMatrix(.Row, .ColIndex("结存id")) & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "结存删除")
            
            Call GetAccountRecord
            Call RefreshList
            
            If mblnStart = True Then
                Call CheckClosAccount
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfList_RowColChange()
    '菜单状态
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
            
    With vsfList
        '移动第一栏的标记到当前行！
        .Cell(flexcpText, 0, 0, .Rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
        
        If mPrives.bln审核 Then
            '审核菜单
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_VerifyAccount, , True)
            Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_VerifyAccount, , True)
    
            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = (.TextMatrix(.Row, .ColIndex("审核日期")) = "")
            If Not cbrControl Is Nothing Then cbrControl.Enabled = (.TextMatrix(.Row, .ColIndex("审核日期")) = "")
        End If
        '删除菜单
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_DeleteAccount, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_DeleteAccount, , True)

        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = (.TextMatrix(.Row, .ColIndex("审核日期")) = "")
        If Not cbrControl Is Nothing Then cbrControl.Enabled = (.TextMatrix(.Row, .ColIndex("审核日期")) = "")
        '取消菜单
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_CancelAccount, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_CancelAccount, , True)

        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = (.TextMatrix(.Row, .ColIndex("审核日期")) <> "" And .TextMatrix(.Row, .ColIndex("取消日期")) = "")
        If Not cbrControl Is Nothing Then cbrControl.Enabled = (.TextMatrix(.Row, .ColIndex("审核日期")) <> "" And .TextMatrix(.Row, .ColIndex("取消日期")) = "")
        If Val(.TextMatrix(.Row, .ColIndex("性质"))) = 0 Then '初始化单据不能取消
            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
            If Not cbrControl Is Nothing Then cbrControl.Enabled = False
        End If
    End With
End Sub

'取卫材金额、价格和数量的小数位数
Private Function newGetDigit(ByVal int性质 As Integer, ByVal int内容 As Integer, Optional ByVal int单位 As Integer) As Integer
    'int性质：0-计算精度;
    'int内容：1-成本价;2-零售价;3-数量;4-金额
    'int单位：如果是取金额位数，可以不输入该参数
    '         卫材单位:1-散装;2-包装
    '返回：最小2，最大为数据库最大小数位数
    
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    If int内容 = 4 Then '取金额 单位=5的才是金额
        int单位 = 5
    End If
    
    gstrSQL = "Select Nvl(精度, 0) 精度 From 药品卫材精度 Where 性质 = [1] And 类别 = 2 And 内容 = [2] And 单位 = [3] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取药品" & Choose(int内容, "成本价", "零售价", "数量") & "小数位数", int性质, int内容, int单位)
    
    If rsTmp.RecordCount > 0 Then
        newGetDigit = rsTmp!精度
    End If
    
    If newGetDigit = 0 Then
        '如果没有设置精度，则取数据库允许的最大位数
        newGetDigit = Choose(int内容, gtype_UserDrugDigits.Digit_成本价, gtype_UserDrugDigits.Digit_零售价, gtype_UserDrugDigits.Digit_数量, gtype_UserDrugDigits.Digit_金额)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    newGetDigit = Choose(int内容, gtype_UserDrugDigits.Digit_成本价, gtype_UserDrugDigits.Digit_零售价, gtype_UserDrugDigits.Digit_数量, gtype_UserDrugDigits.Digit_金额)
End Function



