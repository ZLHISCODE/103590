VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProcUpgrade 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "变动过程升级管理"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12360
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmProcUpgrade.frx":0000
   ScaleHeight     =   7845
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid vsfModule 
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   7335
      _cx             =   12938
      _cy             =   4471
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483636
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
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
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vsfSp 
      Height          =   1935
      Left            =   1560
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   4935
      _cx             =   8705
      _cy             =   3413
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483636
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
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
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.PictureBox pctBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   12360
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2640
      Width           =   12360
      Begin VB.CommandButton cmdManual 
         Caption         =   "过程调整(&U)"
         Height          =   345
         Left            =   7320
         TabIndex        =   28
         Top             =   443
         Width           =   1215
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "导出脚本(&I)"
         Height          =   345
         Left            =   8640
         TabIndex        =   27
         Top             =   443
         Width           =   1215
      End
      Begin VB.Frame fra2 
         Height          =   30
         Index           =   1
         Left            =   2640
         TabIndex        =   25
         Top             =   120
         Width           =   9615
      End
      Begin VB.Frame fra1 
         Height          =   30
         Index           =   1
         Left            =   0
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         ForeColor       =   &H80000000&
         Height          =   270
         Left            =   1800
         TabIndex        =   18
         Text            =   "输过程名称或修改人后按回车进行定位"
         Top             =   480
         Width           =   3735
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAlter 
         Height          =   2415
         Left            =   3960
         TabIndex        =   22
         Top             =   840
         Width           =   7000
         _cx             =   12347
         _cy             =   4260
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
         ForeColorFixed  =   -2147483636
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   200
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
      Begin VSFlex8Ctl.VSFlexGrid vsfProc 
         Height          =   2415
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   3735
         _cx             =   6588
         _cy             =   4260
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483636
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   200
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   5000
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
         ExplorerBar     =   1
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
      Begin VB.Label lblCheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "用户变动过程升级处理"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   26
         Top             =   0
         Width           =   1800
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "查找"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1320
         TabIndex        =   21
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lblAlter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "升级后会被修改的用户变动过程"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4320
         TabIndex        =   20
         Top             =   525
         Width           =   2520
      End
      Begin VB.Label lblProc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "用户变动过程"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   525
         Width           =   1080
      End
   End
   Begin VB.Frame fra1 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "检查(&A)"
      Height          =   350
      Left            =   10200
      TabIndex        =   3
      Top             =   1970
      Width           =   1455
   End
   Begin VB.Frame fra2 
      Height          =   30
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   9975
   End
   Begin VB.Label lblSp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "，请确保当前系统安装目录下特殊SP脚本没有遗漏."
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   2520
      TabIndex        =   31
      Top             =   2295
      Width           =   4050
   End
   Begin VB.Label lblSp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "特殊SP脚本"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   1
      Left            =   1540
      TabIndex        =   30
      Top             =   2295
      Width           =   900
   End
   Begin VB.Label lblSp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前系统执行过"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   2295
      Width           =   1260
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   15
   End
   Begin VB.Label lblWarn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "根据本机安装、升级脚本中的标准产品过程与数据库中的过程进行对比,找出哪些是修改了的过程，以及是否在升级后被修改"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   9810
   End
   Begin VB.Label lblVisable 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "选择…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label lblCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "用户变动过程检查"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label lblCurrent 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "当前版本系统安装目录："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6840
      TabIndex        =   9
      Top             =   2040
      Width           =   1980
   End
   Begin VB.Label lblTarget 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "目标版本系统安装目录："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2760
      TabIndex        =   8
      Top             =   2040
      Width           =   1980
   End
   Begin VB.Label lblCurPath 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   8880
      TabIndex        =   7
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label lblTargetPath 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C:\AppSoft"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4800
      TabIndex        =   6
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label lblCurCmd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "选择…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   9240
      TabIndex        =   5
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label lblTargetCmd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "选择…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   6000
      TabIndex        =   4
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   $"frmProcUpgrade.frx":803A
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   10980
   End
   Begin VB.Image Img 
      Height          =   555
      Left            =   240
      Picture         =   "frmProcUpgrade.frx":8114
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "变动过程升级管理"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmProcUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmCollect As frmProcCollect
Attribute mfrmCollect.VB_VarHelpID = -1
Private mblnChanged As Boolean
Private Enum txtColor
    黑色 = &H80000012
    灰色 = &H80000010
End Enum

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用

End Function

Private Sub cmdCheck_Click()
    Dim strSys As String, i As Integer
    Dim strMsg As String, intNum As Integer
    Dim strInitFile As String, strCurInitPath As String

    strCurInitPath = lblCurPath.Caption
    With vsfModule
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                '获取当前系统的配置文件
                strInitFile = lblTargetPath.Caption & "\" & Decode(.TextMatrix(i, .ColIndex("编号")) \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                    6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                    23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                    25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\应用脚本\ZLSETUP.INI"
                                                                    
                '传入参数格式为: "系统号,系统名称,当前版本,目标版本,目录",多个系统之间用分号间隔
                If strSys = "" Then
                    strSys = .TextMatrix(i, .ColIndex("编号")) & "," & .TextMatrix(i, .ColIndex("系统名称")) & "," & _
                                .TextMatrix(i, .ColIndex("当前版本号")) & "," & .TextMatrix(i, .ColIndex("目标版本号")) & "," & strInitFile
                Else
                    strSys = strSys & ";" & .TextMatrix(i, .ColIndex("编号")) & "," & .TextMatrix(i, .ColIndex("系统名称")) & "," & _
                                                    .TextMatrix(i, .ColIndex("当前版本号")) & "," & .TextMatrix(i, .ColIndex("目标版本号")) & "," & strInitFile
                End If
                intNum = intNum + 1
            End If
        Next
    End With
    
    If intNum = 0 Then
        strMsg = "没有选中任何系统，请重新选择"
        MsgBox strMsg, vbApplicationModal, gstrSysName
        Exit Sub
    End If
    
    strMsg = "共选择" & intNum & "个应用系统，为保证检查结果的正确性，请确保脚本文件的完整性。" & vbNewLine & _
                    "检查执行完成后，将对上次检查结果进行调整，同时删除上次调整的过程代码，你确定要继续吗？"
                    
    If MsgBox(strMsg, vbYesNo, "检查确认") = vbNo Then Exit Sub
    
    vsfAlter.Rows = 1
    vsfProc.Rows = 1
    Set mfrmCollect = New frmProcCollect
    mfrmCollect.ShowMe strSys, strCurInitPath
    
End Sub

Private Sub cmdExport_Click()
    '脚本导出
    Dim strPath As String, i As Long
    Dim blnExp As Boolean, lngNum As Long
    Dim strProc As String, strName As String
    
    On Error GoTo errH

    With vsfAlter
        If .Rows = 1 Then
            MsgBox "本次升级没有变动过程被修改，无需导出。", , "提示"
            Exit Sub
        End If
        
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("修正状态")) = "已调整" Then
                blnExp = True
            ElseIf .TextMatrix(i, .ColIndex("修正状态")) = "待调整" Then
                lngNum = lngNum + 1
            End If
        Next
        
        If Not blnExp Then
            MsgBox "请先对在升级中被修改的变动过程进行检查调整后再进行导出。", , "提示"
            Exit Sub
        Else
            If lngNum > 0 Then
                MsgBox "目前系统下有" & lngNum & "个过程未经过人工调整处理，该部分过程不会通过脚本导出。" & _
                            vbNewLine & "为避免过程遗漏，请在人工调整后对该部分过程再执行脚本导出", , "提示"
            End If
            
            strPath = OpenFolder(Me, "请选择导出脚本目录")
            If strPath = "" Then Exit Sub
            If Right(strPath, 1) <> "\" Then
                strPath = strPath & "\"
            End If
            strPath = strPath & "ProcExport.Sql"
            
            For i = 1 To .Rows - 1
                If i = 1 Then
                    gobjFile.CreateTextFile strPath
                End If
                
                If .TextMatrix(i, .ColIndex("修正状态")) = "已调整" Then
                    strName = .TextMatrix(i, .ColIndex("过程名称"))
                    strProc = GetPorcTxtByName(strName, 3)
                    
                    '需要转出的过程数量低于20,就不执行转出
                    If .Rows - lngNum > 20 Then
                        ShowFlash "正在将过程" & strName & "导出至脚本"
                    End If
                    
                    Do While Right(strProc, 2) = vbNewLine
                        strProc = Left(strProc, Len(strProc) - 2)
                    Loop
                    
                    gobjFile.OpenTextFile(strPath, ForAppending).Write strProc & vbNewLine & "/" & vbNewLine '导出脚本
                    gcnOracle.Execute "Update zlProcedure Set 状态 = 4 Where 名称 =  '" & strName & "'" '修改状态
                    .TextMatrix(i, .ColIndex("修正状态")) = "已导出"
                End If
            Next
            gobjFile.OpenTextFile(strPath).Close
            ShowFlash ""
            MsgBox "过程导出成功。" & vbNewLine & "已经将过程保存至" & strPath, , "提示"
        End If
    End With
    Exit Sub
errH:
    ShowFlash ""
    If 0 = 1 Then
        Resume
    End If
    MsgBox "导出脚本出现错误。" & vbNewLine & Err.Description, , gstrSysName
End Sub

Private Sub cmdManual_Click()
    Dim arrIds() As String, lngIdx As Long
    Dim i As Long
    
    With vsfAlter
        If .Row = 0 Then
            MsgBox "所选过程在本次升级中不会被修改，无需修正。", , gstrSysName
            Exit Sub
        End If
        
        '因为要连续操作,所以把需要调整的过程ID都传到子窗体
        lngIdx = .Row - 1
        ReDim arrIds(.Rows - 2)
        
        For i = 1 To .Rows - 1
            arrIds(i - 1) = .RowData(i) & ":" & .TextMatrix(i, .ColIndex("过程名称"))
        Next
        
    End With
    
    If frmProcDiff.ShowMe(arrIds, lngIdx) Then
        LoadProc
    End If
End Sub


Private Sub Form_Activate()
    Call LoadSystems
    Call LoadSpVer
    Call LoadProc
End Sub


Private Sub Form_Load()
    Dim strCol As String

    '表格初始化
    strCol = " ,400,1;编号,1000,1;系统名称,2000,1;当前版本号,1800,1;目标版本号,1800,1"
    Call InitTable(vsfModule, strCol)
    vsfModule.FixedCols = 1
    vsfModule.ColDataType(0) = flexDTBoolean
    vsfModule.Cell(flexcpChecked, 0, 0) = flexUnchecked
    vsfModule.Cell(flexcpForeColor, 0, 0, 0, vsfModule.Cols - 1) = &H80000008
    
    strCol = " ,350,1;系统,2000,1;过程名称,2000,1;升级前最新脚本,2000,1;修改人,2000,1;修改时间,2000,1;修改说明,2000,1"
    Call InitTable(vsfProc, strCol)
    vsfProc.FixedCols = 1
    vsfProc.Rows = 1
    vsfProc.Cell(flexcpForeColor, 0, 0, 0, vsfProc.Cols - 1) = &H80000008
    
    strCol = " ,390,1;过程名称,3000,1;升级后最新脚本,2400,1;修正状态,500,1"
    Call InitTable(vsfAlter, strCol)
    vsfAlter.FixedCols = 1
    vsfAlter.Rows = 1
    vsfAlter.Cell(flexcpForeColor, 0, 0, 0, vsfAlter.Cols - 1) = &H80000008
    
    strCol = "编号,600,1;系统,2000,1;特殊SP版本,2000,1"
    Call InitTable(vsfSp, strCol)
End Sub

Private Sub ResizeLable()
    On Error Resume Next
    
    
    lblSystem.Width = IIf(lblSystem.Width > 4000, 4000, lblSystem.Width)
    lblSystem.Left = lblWarn.Left
    lblVisable.Left = lblSystem.Left + lblSystem.Width + 60
    lblTarget.Left = lblVisable.Left + lblVisable.Width + 240
    lblTargetPath.Left = lblTarget.Left + lblTarget.Width + 60
    lblTargetCmd.Left = lblTargetPath.Left + lblTargetPath.Width + 60
    lblCurrent.Left = lblTargetCmd.Left + lblTargetCmd.Width + 240
    lblCurPath.Left = lblCurrent.Left + lblCurrent.Width + 60
    lblCurCmd.Left = lblCurPath.Left + lblCurPath.Width + 60
    If lblCurCmd.Visible Then
        cmdCheck.Left = lblCurCmd.Left + lblCurCmd.Width + 240
    Else
        cmdCheck.Left = lblTargetCmd.Left + lblTargetCmd.Width + 240
    End If
    
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    fra2(0).Width = Me.ScaleWidth - fra2(0).Left
    fra2(1).Width = Me.ScaleWidth - fra2(1).Left
    
    ResizeLable

    vsfModule.Left = lblWarn.Left
    
    pctBottom.Width = Me.ScaleWidth
    pctBottom.Height = Me.ScaleHeight - pctBottom.Top
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mfrmCollect = Nothing
End Sub

Private Sub lblSp_Click(Index As Integer)
    If Index <> 1 Then Exit Sub
    With vsfSp
        .Visible = Not .Visible
        If .Visible Then .SetFocus  '可见就获取焦点
    End With
End Sub

Private Sub lblVisable_Click()
    Dim i As Long, intNum  As Integer
    
    With vsfModule
        .Visible = Not .Visible
        
        If .Visible Then .SetFocus  '可见就获取焦点
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                intNum = intNum + 1
            End If
        Next
        lblSystem.Caption = "共有" & .Rows - 1 & "个系统，已选" & intNum & "个系统"
        ResizeLable
    End With
End Sub

Private Sub pctBottom_Resize()
    On Error Resume Next


    
    vsfProc.Width = pctBottom.Width - vsfProc.Left - vsfAlter.Width - 500
    vsfAlter.Left = vsfProc.Width + vsfProc.Left + 360
    vsfProc.Height = pctBottom.ScaleHeight - vsfProc.Top - 200
    vsfAlter.Height = vsfProc.Height
    lblProc.Left = vsfProc.Left
    lblAlter.Left = vsfAlter.Left
    txtFind.Left = vsfProc.Left + vsfProc.Width - txtFind.Width
    lblFind.Left = txtFind.Left - lblFind.Width - 60

    cmdExport.Left = vsfAlter.Width + vsfAlter.Left - cmdExport.Width
    cmdManual.Left = cmdExport.Left - cmdManual.Width - 40
End Sub


Private Sub LoadProc()
    '加载数据库中保存的变动过程
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    ShowFlash "正在加载变动过程..."
    strSQL = "Select a.Id, a.系统编号, a.名称, a.类型, a.状态, a.所有者, a.修改人员, To_Char(a.修改时间, 'yyyy-mm-dd hh24:mi') 修改时间, a.上次修改人员," & vbNewLine & _
                "       To_Char(a.上次修改时间, 'yyyy-mm-dd hh24:mi') 上次修改时间, a.升级前版本, a.升级后版本, a.性质, a.说明, c.名称 系统" & vbNewLine & _
                "From (Select Distinct a.Id, a.系统编号, a.名称, a.类型, a.状态, a.所有者, a.修改人员, a.修改时间, a.上次修改人员, a.上次修改时间, a.升级前版本, a.升级后版本, b.性质," & vbNewLine & _
                "                       a.说明" & vbNewLine & _
                "       From Zlprocedure A, Zlproceduretext B" & vbNewLine & _
                "       Where 类型 = 1 And a.Id = b.过程id And (b.性质 = 1 Or b.性质 = 4)) A, zlSystems C" & vbNewLine & _
                "Where a.系统编号 = c.编号" & vbNewLine & _
                "Order By a.系统编号, a.名称"
    Set rsTmp = OpenSQLRecord(strSQL, "加载变动过程")
    
    '加载变动过程
    With vsfProc
        rsTmp.Filter = "性质 = 1"
        If rsTmp.RecordCount = 0 Then Exit Sub
        rsTmp.MoveFirst
        
        .Rows = 1
        .Rows = rsTmp.RecordCount + 1
        .MergeCells = flexMergeRestrictRows
        .MergeCol(.ColIndex("系统")) = True
        
        .Redraw = flexRDNone
        i = 1
        Do While Not rsTmp.EOF
            .TextMatrix(i, 0) = i
            .TextMatrix(i, .ColIndex("系统")) = rsTmp!系统 & ""
            .TextMatrix(i, .ColIndex("过程名称")) = rsTmp!名称 & ""
            .TextMatrix(i, .ColIndex("修改人")) = rsTmp!修改人员 & ""
            .TextMatrix(i, .ColIndex("修改时间")) = rsTmp!修改时间 & ""
            .TextMatrix(i, .ColIndex("修改说明")) = rsTmp!说明 & ""
            .TextMatrix(i, .ColIndex("升级前最新脚本")) = rsTmp!升级前版本 & ""
            .RowData(i) = rsTmp!Id & ""
            i = i + 1
            rsTmp.MoveNext
        Loop
        .AutoResize = True
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
        
    End With
    
    '加载修改的变动过程
    rsTmp.Filter = "性质 = 4"
    If rsTmp.RecordCount = 0 Then Exit Sub
    rsTmp.MoveFirst
    
    With vsfAlter
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTmp.RecordCount + 1
        
        i = 1
        Do While Not rsTmp.EOF
            .TextMatrix(i, 0) = i
            .TextMatrix(i, .ColIndex("过程名称")) = rsTmp!名称 & ""
            .TextMatrix(i, .ColIndex("升级后最新脚本")) = rsTmp!升级后版本 & ""
            .TextMatrix(i, .ColIndex("修正状态")) = Decode(rsTmp!状态, 1, "待调整", 2, "调整中", 3, "已调整", 4, "已导出") & ""
            If rsTmp!状态 = 1 Then
                .Cell(flexcpForeColor, i, .ColIndex("修正状态")) = 兰色
            Else
                .Cell(flexcpForeColor, i, .ColIndex("修正状态")) = 黑色
            End If
            .RowData(i) = rsTmp!Id & ""
            
            rsTmp.MoveNext
            i = i + 1
        Loop
        .Redraw = flexRDDirect
    End With
    
    ShowFlash ""
    Exit Sub
errH:
    ShowFlash ""
    If 0 = 1 Then
        Resume
    End If
    MsgBox Err.Description, , gstrSysName
End Sub

Private Sub LoadSystems()
    '加载安装的系统
    Dim strSQL As String, rsSys As New ADODB.Recordset
    Dim i As Long, strTmp As String
    
    '首先获取系统编号等信息
    strSQL = "Select 编号 系统编号, 名称 系统名称, 版本号 系统版本号, 所有者 系统所有者, 正常安装 From Zlsystems where Upper(所有者)=[1] Order by Nvl(共享号,0),编号"
    Set rsSys = OpenSQLRecord(strSQL, "读取安装系统", gstrUserName)
    
    If rsSys.RecordCount = 0 Then
        MsgBox "请使用系统所有者登录。", , gstrSysName
        Exit Sub
    Else
        With vsfModule
            i = .FixedRows
            .Rows = .FixedRows
            .Rows = rsSys.RecordCount + .FixedRows
            Do While Not rsSys.EOF
                .TextMatrix(i, .ColIndex("编号")) = rsSys!系统编号 & ""
                .TextMatrix(i, .ColIndex("系统名称")) = rsSys!系统名称 & ""
                .TextMatrix(i, .ColIndex("当前版本号")) = rsSys!系统版本号 & ""
                .TextMatrix(i, .ColIndex("目标版本号")) = ""
                rsSys.MoveNext
                i = i + 1
            Loop
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = flexAlignCenterCenter
        End With
    End If
    
    LoadUpdateSystem lblTargetPath.Caption
End Sub

Private Sub LoadUpdateSystem(ByVal strPath As String)
    '获取各个系统的升级目标版本
    Dim i As Long, strInitFile As String
    Dim strTarget As String, blnStep As Boolean
    Dim intNum As Integer
    
    With vsfModule
        For i = 1 To .Rows - 1
            strInitFile = strPath & "\" & Decode(.TextMatrix(i, .ColIndex("编号")) \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                                6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                                23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                                25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\应用脚本\ZLSETUP.INI"
            If gobjFile.FileExists(strInitFile) Then
                If GetUpgradeFiles(Nothing, Val(.TextMatrix(i, .ColIndex("编号"))), .TextMatrix(i, .ColIndex("当前版本号")), strInitFile, , , , strTarget, , True, False) Is Nothing Then
                    .Cell(flexcpText, 1, .ColIndex("目标版本号"), .Rows - 1, .ColIndex("目标版本号")) = ""
                    .Cell(flexcpChecked, i, 0) = flexUnchecked
                    Exit Sub
                End If
                .TextMatrix(i, .ColIndex("目标版本号")) = strTarget
                
                If strTarget <> "" Then
                    intNum = intNum + 1
                    .Cell(flexcpChecked, i, 0) = flexChecked
                    
                    '检查是否跨版本升级
                    If .TextMatrix(i, .ColIndex("当前版本号")) <> "" And GetPrimaryVer(.TextMatrix(i, .ColIndex("当前版本号"))) <> GetPrimaryVer(strTarget) Then
                        blnStep = True
                    End If
                Else
                    .TextMatrix(i, .ColIndex("目标版本号")) = ""
                    .Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            Else
                .TextMatrix(i, .ColIndex("目标版本号")) = ""
                .Cell(flexcpChecked, i, 0) = flexUnchecked
            End If
        Next
        
        lblSystem.Caption = "共有" & .Rows - 1 & "个系统，已选" & intNum & "个系统"
        lblCurrent.Visible = blnStep
        lblCurPath.Visible = blnStep
        lblCurCmd.Visible = blnStep
        ResizeLable
    End With
End Sub

Private Sub LoadSpVer()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    
    strSQL = "Select a.系统, b.名称,a.目标版本 特殊SP版本" & vbNewLine & _
                "From zlUpGrade A, zlSystems B" & vbNewLine & _
                "Where a.结果版本 Like '%.%.%.%' And a.系统 = b.编号 And" & vbNewLine & _
                "      Substr(a.结果版本, 1, Instr(a.结果版本, '.', 1, 2) - 1) = Substr(b.版本号, 1, Instr(b.版本号, '.', 1, 2) - 1)" & vbNewLine & _
                "Order By a.系统"
    
    Set rsTmp = OpenSQLRecord(strSQL, "获取特殊SP版本")
    
    If rsTmp.RecordCount = 0 Then
        lblSp(0).Visible = False
        lblSp(1).Visible = False
        lblSp(2).Visible = False
    Else
        lblSp(0).Visible = True
        lblSp(1).Visible = True
        lblSp(2).Visible = True
        
        With vsfSp
            .Rows = 1: .Rows = rsTmp.RecordCount + 1
            i = 1
            Do While Not rsTmp.EOF
                .TextMatrix(i, .ColIndex("编号")) = rsTmp!系统
                .TextMatrix(i, .ColIndex("系统")) = rsTmp!名称
                .TextMatrix(i, .ColIndex("特殊SP版本")) = rsTmp!特殊SP版本
                i = i + 1
                rsTmp.MoveNext
            Loop
        End With

    End If
    
End Sub

Private Sub mfrmCollect_ReturnChangedProc(ByVal rsTmp As ADODB.Recordset, ByVal intType As Integer)
    '接收到记录集后加载至表格中
    'intType: 1表示变动过程 2表示升级脚本中的变动过程
    Dim i As Long
    
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    If intType = 1 Then
        
        With vsfProc
            .Redraw = flexRDNone
            .MergeCells = flexMergeRestrictRows
            .MergeCol(.ColIndex("系统")) = True
            i = .Rows
            .Rows = rsTmp.RecordCount + .Rows
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                .TextMatrix(i, .ColIndex("系统")) = rsTmp!P_System & ""
                .TextMatrix(i, 0) = i
                .TextMatrix(i, .ColIndex("过程名称")) = rsTmp!P_Name & ""
                .TextMatrix(i, .ColIndex("升级前最新脚本")) = rsTmp!P_Ver & ""
                rsTmp.MoveNext
                i = i + 1
            Loop
            .AutoResize = True
            .AutoSize 0, .Cols - 1
            .Redraw = flexRDDirect
        End With
    Else
        With vsfAlter
            .Redraw = flexRDNone
            i = .Rows
            .Rows = .Rows + rsTmp.RecordCount
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                .TextMatrix(i, 0) = i
                .TextMatrix(i, .ColIndex("过程名称")) = rsTmp!P_Name & ""
                .TextMatrix(i, .ColIndex("升级后最新脚本")) = rsTmp!P_Ver & ""
                .TextMatrix(i, .ColIndex("修正状态")) = "待调整"
                rsTmp.MoveNext
                i = i + 1
            Loop
            .AutoResize = True
            .AutoSize 0, .Cols - 1
            .Redraw = flexRDDirect
        End With
    End If
    
End Sub

Private Sub lblCurCmd_Click()
    Dim strPath As String
    
    strPath = OpenFolder(Me, "选择当前版本系统安装目录")
    If strPath = "" Then Exit Sub
    
    lblCurPath.Caption = strPath
    lblCurCmd.Left = lblCurPath.Left + lblCurPath.Width + 60
    
    LoadUpdateSystem strPath
End Sub

Private Sub lblTargetCmd_Click()
    Dim strPath As String
    
    strPath = OpenFolder(Me, "选择目标版本系统安装目录")
    If strPath = "" Then Exit Sub
    
    lblTargetPath.Caption = strPath
    lblTargetCmd.Left = lblTargetPath.Left + lblTargetPath.Width + 60
    
    LoadUpdateSystem strPath
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = "输过程名称或修改人后按回车进行定位" Then
        txtFind.Text = ""
        txtFind.ForeColor = 黑色
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If txtFind.Text = "" Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    
    GetRowPos vsfProc, txtFind.Text, "过程名称,修改人"
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "输过程名称或修改人后按回车进行定位"
        txtFind.ForeColor = 灰色
    End If
End Sub

Private Sub vsfAlter_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '联动选中
    Dim strProc As String, i As Long
    
    On Error Resume Next

    With vsfAlter
        If .Rows = 1 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        
        .Cell(flexcpForeColor, OldRow, 0) = Color.深灰色
        .Cell(flexcpFontBold, OldRow, 0) = False
        .Cell(flexcpFontBold, NewRow, 0) = True
        .Cell(flexcpForeColor, NewRow, 0) = Color.兰色
        
        If mblnChanged Then '防止重复调用换行事件
            mblnChanged = False
            Exit Sub
        End If
        mblnChanged = True
        strProc = .TextMatrix(NewRow, .ColIndex("过程名称"))

    End With
    
    With vsfProc
        If .Rows = 1 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        For i = 1 To .Rows - 1
            If strProc = .TextMatrix(i, .ColIndex("过程名称")) Then
                .Select i, 0
                .TopRow = i - (vsfAlter.Row - vsfAlter.TopRow)
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub vsfSp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If vsfSp.Visible = True Then vsfSp.Visible = False
    End If
End Sub
Private Sub vsfModule_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If vsfModule.Visible = True Then vsfModule.Visible = False
    End If
End Sub

Private Sub vsfModule_LostFocus()
    If vsfModule.Visible Then
        lblVisable_Click
    End If
End Sub

Private Sub vsfProc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '联动选中
    Dim strProc As String, i As Long
    
    On Error Resume Next
    With vsfProc
    
        If .Rows = 1 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        
        .Cell(flexcpForeColor, OldRow, 0) = Color.深灰色
        .Cell(flexcpFontBold, OldRow, 0) = False
        .Cell(flexcpFontBold, NewRow, 0) = True
        .Cell(flexcpForeColor, NewRow, 0) = Color.兰色
    
        If mblnChanged Then
            mblnChanged = False
            Exit Sub
        End If
    
        mblnChanged = True
        strProc = .TextMatrix(NewRow, .ColIndex("过程名称"))
    End With
    
    With vsfAlter
        If .Rows = 1 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        
        For i = 1 To .Rows - 1
            If strProc = .TextMatrix(i, .ColIndex("过程名称")) Then
                .Select i, 0
                .TopRow = i - (vsfProc.Row - vsfProc.TopRow)
                Exit Sub
            End If
        Next
        
        If i = .Rows - 1 And strProc <> .TextMatrix(i, .ColIndex("过程名称")) Then
            .Select 0, 0
        End If
    End With
    
End Sub

Private Sub vsfModule_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then
        Cancel = True
    End If
    
    '没有升级脚本,不能选中
    With vsfModule
        If .TextMatrix(Row, .ColIndex("当前版本号")) = "" Or .TextMatrix(Row, .ColIndex("目标版本号")) = "" Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfModule_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsfModule
        If .Redraw = flexRDNone Then Exit Sub
        If .Rows = 1 Then Exit Sub
        
        '全选
        If Row = 0 Then
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    If .TextMatrix(i, .ColIndex("当前版本号")) <> "" And .TextMatrix(i, .ColIndex("目标版本号")) <> "" Then
                        .Cell(flexcpChecked, i, 0) = flexChecked
                    End If
                Else
                    .Cell(flexcpChecked, 1, 0, .Rows - 1, 0) = flexUnchecked
                End If
            Next
        End If
    End With
End Sub


