VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "表格式病历编辑器"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11775
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMain"
   ScaleHeight     =   7440
   ScaleWidth      =   11775
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picAtt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8970
      Left            =   420
      ScaleHeight     =   8970
      ScaleWidth      =   3165
      TabIndex        =   13
      Top             =   1350
      Visible         =   0   'False
      Width           =   3165
      Begin VB.CheckBox cmdAvg 
         Caption         =   "均值"
         Height          =   330
         Left            =   600
         MouseIcon       =   "frmMain.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5665
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CheckBox cmdSum 
         Caption         =   "合计"
         Height          =   330
         Left            =   105
         MouseIcon       =   "frmMain.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5665
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Text            =   "frmMain.frx":1A5E
         Top             =   6105
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox txtSum 
         Appearance      =   0  'Flat
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
         Height          =   3570
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "合计应用"
         Height          =   350
         Left            =   1875
         TabIndex        =   32
         Top             =   5655
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   1
         Left            =   165
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "frmMain.frx":1B69
         Top             =   6210
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "frmMain.frx":1C82
         Top             =   6315
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   3
         Left            =   345
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "frmMain.frx":1CD1
         Top             =   6390
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   4
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "frmMain.frx":1D52
         Top             =   6510
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   5
         Left            =   615
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   27
         Text            =   "frmMain.frx":1DB9
         Top             =   6615
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   6
         Left            =   750
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmMain.frx":1E10
         Top             =   6705
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   7
         Left            =   855
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmMain.frx":1E69
         Top             =   6825
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   8
         Left            =   1005
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "frmMain.frx":1EA8
         Top             =   6960
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.Frame fraType 
         Caption         =   "单元格类型"
         Height          =   1980
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   2775
         Begin VB.CheckBox chkType 
            Caption         =   "列控签名"
            Height          =   350
            Index           =   8
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1230
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "行控签名"
            Height          =   350
            Index           =   7
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "签名位"
            Height          =   350
            Index           =   6
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   570
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "报告图"
            Height          =   350
            Index           =   5
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "参考图"
            Height          =   350
            Index           =   4
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1560
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "混合编辑"
            Height          =   350
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1230
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "单要素"
            Height          =   350
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "多行文本"
            Height          =   350
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   570
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "固定文本"
            Height          =   350
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Shape shpTxtSum 
         BorderColor     =   &H00E09060&
         Height          =   255
         Left            =   3300
         Top             =   4035
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.Timer timeTmp 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3165
      Tag             =   "用于改变行高列宽后处理缓存记录"
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar Processing 
      Height          =   270
      Left            =   1875
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox picHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   345
      ScaleHeight     =   1800
      ScaleWidth      =   2625
      TabIndex        =   5
      Top             =   465
      Width           =   2625
      Begin VSFlex8Ctl.VSFlexGrid vsHistory 
         Height          =   1275
         Left            =   165
         TabIndex        =   6
         Top             =   105
         Width           =   2070
         _cx             =   3651
         _cy             =   2249
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
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
         AutoSizeMode    =   1
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
   Begin VB.PictureBox picMainBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5385
      Left            =   3780
      ScaleHeight     =   5385
      ScaleWidth      =   8415
      TabIndex        =   2
      Top             =   180
      Width           =   8415
      Begin zlTableEPR.Document Doc 
         Height          =   885
         Left            =   2775
         TabIndex        =   12
         Top             =   4260
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1561
         Border          =   0   'False
      End
      Begin VB.PictureBox PicDy 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   885
         Index           =   0
         Left            =   1425
         ScaleHeight     =   885
         ScaleWidth      =   1170
         TabIndex        =   10
         Top             =   4260
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.PictureBox picRulerV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3825
         Left            =   0
         ScaleHeight     =   3825
         ScaleWidth      =   225
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Y向标尺"
         Top             =   225
         Width           =   225
      End
      Begin VB.PictureBox picRulerH 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   5340
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "X向标尺"
         Top             =   0
         Width           =   5340
      End
      Begin TTF160Ctl.F1Book F1Main 
         Height          =   3825
         Left            =   285
         TabIndex        =   0
         Top             =   240
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   6747
         _0              =   $"frmMain.frx":1EE7
         _1              =   $"frmMain.frx":22F0
         _2              =   $"frmMain.frx":26F9
         _3              =   $"frmMain.frx":2B03
         _4              =   $"frmMain.frx":2F0C
         _count          =   5
         _ver            =   2
      End
      Begin zlTableEPR.ElementEdit elEdit 
         Height          =   2055
         Left            =   5415
         TabIndex        =   9
         Top             =   570
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   3625
      End
      Begin zlTableEPR.PictureEdit PicEdit 
         Height          =   1860
         Left            =   5460
         TabIndex        =   11
         Top             =   630
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   3281
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7080
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":320C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   "msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
            Text            =   "姓名:最大姓名 住院号:12345678 床号:12345 年龄:12岁 性别:未知 医保号:1234567890123"
            TextSave        =   "姓名:最大姓名 住院号:12345678 床号:12345 年龄:12岁 性别:未知 医保号:1234567890123"
            Key             =   "PatInfo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   794
            MinWidth        =   88
            Text            =   "Ins"
            TextSave        =   "Ins"
            Key             =   "Insert"
            Object.ToolTipText     =   "插入键是否开启"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
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
   Begin zlTableEPR.ColorPicker ColorForeColor 
      Height          =   2190
      Left            =   2205
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   -45
      Top             =   1755
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A9E
            Key             =   "HIGHLIGHT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C13
            Key             =   "FORECOLOR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D60
            Key             =   "FILLCOLOR"
         EndProperty
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
      Bindings        =   "frmMain.frx":3ED4
      Left            =   480
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type SelRange        'F1Book选择区域,起始行、列,终止行、列
    lsRow As Long
    lsCol As Long
    leRow As Long
    leCol As Long
End Type
Private Const conPane_SentenceList = 500
Private Const conPane_Attribute = 501
Private Const conPane_History = 502
Private Const conPane_Content = 503
Private Const conPane_PacsPic = 504
Public Document As cTableEPR, editType As Integer, EditMode As Integer
Private DocOld As Collection
'窗口
Private WithEvents mfrmSentence As frmSentenceList
Attribute mfrmSentence.VB_VarHelpID = -1
Private WithEvents mfrmMainError As frmMainMsg
Attribute mfrmMainError.VB_VarHelpID = -1
Private WithEvents mfrmPacsPic As frmPACSImg
Attribute mfrmPacsPic.VB_VarHelpID = -1
Private WithEvents mfrmEPRModelSaveAs As frmEPRModelSaveAs
Attribute mfrmEPRModelSaveAs.VB_VarHelpID = -1
Private mfrmTipInfo As New frmTipInfo

'窗口变量
Private SelCell As New cTabCell
Private Undo As New cTabUndos
Private mReadOnly As Byte, mstrSex As String '性别
Private mfrmParent As Object, mstrPrivs As String, mstrModelPrivate As String, mblnCanPrint As Boolean, mblnMoved As Boolean
Private mblnInit As Boolean '初始化过程中
Private mblnChangeRC As Boolean '改变行高列宽
Private mblnClickZ As Boolean '在第0列或第0行鼠标按下
Private mblnShowAtt As Boolean, mblnAdd As Boolean '显示属性中、追加
Private mblnEditing As Boolean '文本或固定文本处于编辑状态
Private mbFunType As Byte '点击过函数按键，正在指定函数数据源 =0 Sum =1 Avg

Private Sub AddUndo(TmpCell As cTabCell)
    If TmpCell.Key = "" Then Exit Sub
    Undo.Add Undo.Count & "_" & TmpCell.Key
    With Undo(Undo.Count)
        .Key = TmpCell.Key
        .CT = TmpCell.对象类型
        .CTxt = TmpCell.内容文本
        .Ekey = TmpCell.ElementKey
        .Tkey = TmpCell.TextKey
        .PKey = TmpCell.PictureKey
        If Len(.PKey) <> 0 Then
            Set .OrigPic = Document.Pictures("K" & .PKey).OrigPic
        End If
        .PmKey = TmpCell.PicMarkKey
    End With
End Sub

Public Sub ShowMe(ByVal frmParent As Object, DocTab As cTableEPR, ByVal strModelPrivate As String, ByVal blnMoved As Boolean, Optional blnCanPrint As Boolean = True, Optional ByVal intStyle As Integer)
'## 参数：  frmParent       :父窗体
'##         Doc             :外部程序创建的cTableEPR类,包含文档中各对象类
'##         strModelPrivate:调用模块拥有权限
'##         blnMoved        :当前病历是否被转移
'##         blnCanPrint     :是否允许预览、打印
'################################################################################################################
Dim bfrmMode As Byte
    '设置窗体显示状态
    mblnInit = False: mblnChangeRC = False: mblnClickZ = False: mblnShowAtt = False: mblnAdd = False: mblnEditing = False: mbFunType = 0
    On Error GoTo errHand
    Set mfrmParent = frmParent
    mstrModelPrivate = strModelPrivate
    mstrPrivs = GetPrivFunc(glngSys, 1070)
    mblnCanPrint = blnCanPrint
    mblnMoved = blnMoved
    mblnInit = True
    Set Document = DocTab              '对像赋值
    editType = Document.ET: EditMode = Document.EM
    
    Call OpenDoc(False)                   '跟据编辑模式打开文档或新增文档
    Call DockPaneState              '左边列表状态
    Call RefreshPatiInfo            '刷新病人信息栏
    If editType = TabET_单病历审核 Then Call RereshHistory             '刷新历史版本
    If Document.EPRFileInfo.种类 = Tab诊疗报告 Then zlRefreshPacsPic
    Me.Caption = IIf(mReadOnly = 2, "查阅历史版本----", "") & Document.EPRFileInfo.名称
    mblnInit = False: zlCommFun.StopFlash
    If Not Me.Visible Then
       'bfrmMode = 0
        If editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑 Then intStyle = 1
       'If mblnMoved = 2 Then bfrmMode = 1
        Me.Show intStyle, mfrmParent           '显示编辑器,查看模式以模态窗口显示
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
    Unload Me
    Call SaveErrLog
End Sub
Private Function SaveDoc(Optional blnSign As Boolean, Optional blnExit As Boolean) As Boolean
'功能：保存数据
'参数：blnSign表示签名保存
Dim arrSQL As Variant, i As Integer, blnTran As Boolean, SignCellKey As String
'1 对行控签名,列控签名,单签名 进行处理
'2 对行高,列宽保存,及其它数据校验
'2 对修改情况下对象变化,审核时,保存前对每个单元格内容验证,变动过的先生成终止版SQL,然后ID清空,开始版=终止版,终止版=0
'3 调用Document.Save函数最得保存数据SQL,失败直接退出
'4 对开启事务执行SQL,同时显示进度条
'5 再次刷新界面

    On Error GoTo errHand
1    mblnInit = True
2    arrSQL = Array()
3    If blnSign Then
4        If frmSign.Visible Then Exit Function
5    End If
    
    
6    Processing.Value = 0: Processing.Visible = True: Processing.Max = 200
    

7    stbThis.Panels("msg").Text = "开始检查数据完整性--------"
8    If Doc.Visible Or mblnEditing Then F1Main_GotFocus '如果处于编辑状态，需要将内容更新下来
    
9    If Not ValiCellDate(Not mblnAdd) Then Processing.Visible = False: GoTo lOut
10    Processing.Value = Processing.Value + 10
    
11    stbThis.Panels("msg").Text = "开始检查变更数据--------"
12    If editType = TabET_单病历审核 Then Call CompareChange(arrSQL)
13    Processing.Value = Processing.Value + 10
    
14    If blnSign Then
15        If Not AddSign(arrSQL, SignCellKey) Then Processing.Visible = False: GoTo lOut '签名
16    End If
    
17    zlCommFun.ShowFlash "正在保存数据，请稍等！", Me
18    stbThis.Panels("msg").Text = "开始生成数据保存SQL--------"
19    If Not Document.SaveDoc(arrSQL) Then
20        Processing.Visible = False
21        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "警告：" & vbCrLf & "      保存出错,数据将不会被记录", True, 0
22        stbThis.Panels("msg").Text = "保存出错,数据将不会被记录"
23        GoTo lOut
24    End If
25    Processing.Value = Processing.Value + 10

26    Err.Clear
27    gcnOracle.BeginTrans '--------------------------写入数据
28    stbThis.Panels("msg").Text = "开始提交数据--------"
29    Processing.Max = Processing.Value + UBound(arrSQL) + 1: blnTran = True
30    For i = 0 To UBound(arrSQL)
31        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "写入数据")
32        Processing.Value = Processing.Value + 1
33    Next
34    stbThis.Panels("msg").Text = "保存完毕"
35    Call gcnOracle.CommitTrans: blnTran = False: Processing.Visible = False
    
36    If (editType = TabET_单病历编辑 Or editType = TabET_单病历审核) And Not blnExit Then
37        If EditMode = TabEm_新增 Then Document.EM = TabEm_修改: EditMode = TabEm_修改
38        Call Document.EPRPatiRecInfo.GetPatiRecordInfo(Document.EPRPatiRecInfo.ID, mblnMoved) '重读电子病历记录
39    End If
    
40    SaveDoc = True: Call RelateFeedback(True)

lOut:   On Error Resume Next
41        mfrmParent.RefreshList
42        Call mfrmParent.Event_Saved(Document.EPRPatiRecInfo.ID) '诊疗单据需要，因为可能是非模态方式调用，不能用事件方式
43        Err.Clear
44        mblnInit = False
45        zlCommFun.StopFlash
46        Exit Function
errHand:
    Call MsgBox("SaveDoc错误行：" & Erl(), vbInformation, gstrSysName)
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False: zlCommFun.StopFlash
    If blnTran Then gcnOracle.RollbackTrans
    
    Call SaveErrLog
    mReadOnly = 2 '变为查看模式，不允许再次保存，需要退出以重整数据
    stbThis.Panels("msg").Text = "保存出错,数据将不会被记录"
    If blnSign Then
        F1Main.TextRC(Document.Cells(SignCellKey).Row, Document.Cells(SignCellKey).Col) = "[签名位]"
    End If
    Processing.Visible = False
    MsgBox "保存出错,数据将不会被记录！", vbExclamation, gstrSysName
End Function
Private Sub OpenDoc(Optional ByVal blnNew As Boolean)
'功能：读取文件结构，文件内容，刷新文档界面
'说明:blnNew 表示界面上点新建
    If blnNew Then mblnInit = True
    
    If blnNew Then
        Document.EM = TabEm_新增
        ClearPicture
    End If
    If blnNew And editType = TabET_病历文件定义 Then   '文件定义时点新建
        Document.InitEmptyStructure         '初始化一个空文档
    Else
        If Document.ReadFileStructure Then   '读取文件结构
            Document.ReadFileContent mblnMoved  '读取文件内容
        Else
            Document.InitEmptyStructure         '初始化一个空文档
        End If
    End If
    mReadOnly = Document.mReadOnly  '在OpenDoc过程中可能改变0-正常,1-签名后点修改,2-主界面打开查阅或查阅历次签名版本
    If editType = TabET_单病历审核 Then Set DocOld = New Collection  '审核打开时先保存原始记录,用于在保存时进行对比
    Call RefreshF1Main: mblnInit = False                        '填充表格内容
End Sub
Private Sub DockPaneState()
Dim PaneHistory As Pane, PaneSentenceList As Pane, PaneAttribute As Pane, PanePacsPic As Pane, PaneContent As Pane

    On Error GoTo errHand
    
    Set PaneHistory = dkpMain.FindPane(conPane_History)
    Set PaneSentenceList = dkpMain.FindPane(conPane_SentenceList)
    Set PaneAttribute = dkpMain.FindPane(conPane_Attribute)
    Set PanePacsPic = dkpMain.FindPane(conPane_PacsPic)
    Set PaneContent = dkpMain.FindPane(conPane_Content)
    
    If Not PaneSentenceList Is Nothing Then
        dkpMain_AttachPane PaneSentenceList
    End If
    If Not PaneAttribute Is Nothing Then
        dkpMain_AttachPane PaneAttribute
    End If
    If Not PaneHistory Is Nothing Then
        dkpMain_AttachPane PaneHistory
    End If
    If Not PanePacsPic Is Nothing Then
        dkpMain_AttachPane PanePacsPic
    End If
    
    Select Case editType
        Case TabET_病历文件定义, TabET_全文示范编辑
            If Not PaneHistory Is Nothing Then
                PaneHistory.Close
                PanePacsPic.Close
            End If

            dkpMain.ShowPane conPane_Attribute
        Case TabET_单病历编辑, TabET_单病历审核
            If Not PaneAttribute Is Nothing Then
                PaneAttribute.Close
            End If
            dkpMain.ShowPane conPane_SentenceList
            
            If Document.EPRFileInfo.种类 <> Tab诊疗报告 Then
                If Not PanePacsPic Is Nothing Then PanePacsPic.Close
            End If
            
            If mReadOnly = 2 Then
                PaneSentenceList.Close
                PanePacsPic.Close
                PaneHistory.Close
            End If
            
            If editType = TabET_单病历编辑 Then PaneHistory.Close
    End Select
    
    If Not PaneContent Is Nothing Then PaneContent.Selected = True
    PostMessage Processing.hWnd, PBM_SETBARCOLOR, 0, &H80FF80
    PostMessage Processing.hWnd, PBM_SETBKCOLOR, 0, vbWhite
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RefreshPatiInfo()
    If editType <> TabET_单病历编辑 And editType <> TabET_单病历审核 Then
        stbThis.Panels("PatInfo").Text = "进度栏"
        If stbThis.Panels("PatInfo").Width > 3800 Then
            stbThis.Panels("PatInfo").Width = stbThis.Panels("PatInfo").Width / 3
        End If
        Exit Sub
    End If

    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    If Document.EPRPatiRecInfo.婴儿 <> 0 Then
        gstrSQL = "Select '姓名:' ||  nvl(B.婴儿姓名,A.姓名 || '之婴' || B.序号) || Decode([2], 2, '  母亲住院号:' || A.住院号 || '  母亲床号:' || A.当前床号, '  母亲门诊号:' || A.门诊号) ||" & vbNewLine & _
                "        '  年龄:' || To_Char(B.出生时间, 'YYYY-MM-DD HH24:MI:SS') || '  性别:' || Nvl(B.婴儿性别,'未知') || '  医保号' || A.医保号 As 信息, Nvl(B.婴儿性别,'未知') 性别" & vbNewLine & _
                "From 病人信息 A, 病人新生儿记录 B" & vbNewLine & _
                "Where A.病人id = [1] And A.病人id = B.病人id And B.主页id = [3] And B.序号 = [4]"
    Else
        gstrSQL = "Select '姓名:' || 姓名 || Decode([2], 2, '  住院号:' || 住院号 || '  床号:' || 当前床号, '  门诊号:' || 门诊号) || '  性别:' || 性别 || '  年龄:' || 年龄 ||" & vbNewLine & _
                "        '  医保号' || 医保号 As 信息, 性别" & vbNewLine & _
                "From 病人信息" & vbNewLine & _
                "Where 病人id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.病人来源, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.婴儿)
    If rsTemp.EOF Then
        stbThis.Panels("PatInfo").Text = "病人信息栏": mstrSex = "未知"
    Else
        stbThis.Panels("PatInfo").Text = rsTemp!信息
    End If
    

    If Me.Document.EPRPatiRecInfo.医嘱id = 0 Then
        Select Case Me.Document.EPRPatiRecInfo.病人来源
        Case TabPF_门诊
            gstrSQL = "Select r.急诊 From 病人挂号记录 r Where r.Id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.Document.EPRPatiRecInfo.主页ID)
            If Not rsTemp.EOF Then
                stbThis.Panels("PatInfo").Text = stbThis.Panels("PatInfo").Text & "  病况:" & IIf(NVL(rsTemp!急诊, 0) = 1, "急", "")
            End If
        Case TabPF_住院
            gstrSQL = "Select 入院病况, 出院日期, 出院方式 From 病案主页 Where 病人id = [1] And Nvl(主页id, 0) = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID)
            If Not rsTemp.EOF Then
                If IsNull(rsTemp!出院日期) Then
                    stbThis.Panels("PatInfo").Text = stbThis.Panels("PatInfo").Text & "  病况:" & NVL(rsTemp!入院病况)
                Else
                    stbThis.Panels("PatInfo").Text = stbThis.Panels("PatInfo").Text & "  病况:" & NVL(rsTemp!出院方式) & "(出院)"
                End If
            End If
        End Select
    Else
        gstrSQL = "Select 紧急标志 From 病人医嘱记录 Where Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.Document.EPRPatiRecInfo.医嘱id)
        If Not rsTemp.EOF Then
            stbThis.Panels("PatInfo").Text = stbThis.Panels("PatInfo").Text & "  病况:" & IIf(NVL(rsTemp!紧急标志, 0) = 1, "急", "")
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RereshHistory()
Dim rsTemp As ADODB.Recordset
On Error GoTo errHand
    With vsHistory
        .Clear: .Rows = 2: .Cols = 4
        .ColWidth(0) = 1000: .ColWidth(1) = 2400: .ColWidth(2) = 1000: .ColWidth(3) = 600
        .TextMatrix(0, 0) = "签名人": .TextMatrix(0, 1) = "签名时间": .TextMatrix(0, 2) = "签名级别": .TextMatrix(0, 3) = "版本"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    If mReadOnly = 2 Then Exit Sub '查历史版本时不再显示历史版本
    
    gstrSQL = "Select 要素表示, 内容文本, 对象属性, 终止版" & vbNewLine & _
                "From 电子病历内容" & vbNewLine & _
                "Where 文件id = [1] And 对象类型 In (6, 7, 8) And Nvl(终止版, 0)>0 and Nvl(终止版, 0)<=[2]" & vbNewLine & _
                "Order By 终止版"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Document.EPRPatiRecInfo.ID, Document.EPRPatiRecInfo.最后版本)
    If rsTemp.EOF Then
        On Error Resume Next
        dkpMain.FindPane(conPane_History).Close
        Exit Sub
    End If
    
    With vsHistory
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .RowHeight(rsTemp.AbsolutePosition) = 800
            .TextMatrix(rsTemp.AbsolutePosition, 0) = NVL(rsTemp!内容文本)
            .TextMatrix(rsTemp.AbsolutePosition, 1) = Split(Split(rsTemp!对象属性, "|")(1), ";")(4)
            .TextMatrix(rsTemp.AbsolutePosition, 2) = Decode(Document.EPRPatiRecInfo.病历种类, 4, Decode(rsTemp!要素表示, 3, "护士长", "护士"), Decode(rsTemp!要素表示, 3, "主任医师", 2, "主治医师", "经治医师"))
            .TextMatrix(rsTemp.AbsolutePosition, 3) = CInt(rsTemp!终止版)
            rsTemp.MoveNext
        Loop
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 2) = flexAlignLeftCenter
        .Cell(flexcpFontSize, 1, 0, .Rows - 1, .Cols - 1) = 12
    End With

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub MainCommandbarDefine()
'## 菜单初始化
    Dim cbpPopup As CommandBarPopup                     '临时对象
    Dim subPopup As CommandBarPopup                     '子菜单
    Dim objControl As CommandBarControl                 '工具栏控件
    Dim objCustControl As CommandBarControlCustom       '自定义控件'
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsMain
        .VisualTheme = xtpThemeOffice2003
        .StatusBar.Visible = False
        .ActiveMenuBar.Title = "菜单栏"
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        .Icons = frmPublicIcon.imgPublic.Icons
        .EnableCustomization (False)
        .Options.IconsWithShadow = True '放在VisualTheme后有效
        .Options.ToolBarAccelTips = True
        .Options.ShowExpandButtonAlways = False '显示扩展按钮
        .Options.UseDisabledIcons = True
        .Options.AlwaysShowFullMenus = False '是否显示所有菜单
    End With
    
'------------------------------------------------文件-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "文件(&F)"): cbpPopup.ID = ID_File_Menu
    With cbpPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_CLEAR, "清空(&C)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "保存(&S)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE_QUIT, "保存退出(&Q)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVEASEPRDEMO, "另存为范文(&M)")
        Set objControl = .Add(xtpControlButton, ID_EDIT_SAVEASPHRASE, "另存为词句(&D)")
        Set objControl = .Add(xtpControlButton, ID_FILE_EXPORTTOXML, "导出为XML文件(&E)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORTFROMXML, "从XML文件导入(&I)")

        Set objControl = .Add(xtpControlButton, ID_FILE_PAGESETUP, "页面设置(&U)..."): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINTPREVIEW, "打印预览(&V)")
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "打印(&P)...")
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "退出(&X)")
    End With
    
'------------------------------------------------编辑-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "编辑(&E)"): cbpPopup.ID = ID_Edit_Menu
    With cbpPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "撤销(&U)"): objControl.ToolTipText = "撤销以单元格为最小单位的内容变化"
        Set objControl = .Add(xtpControlButton, ID_EDIT_REDO, "重做(&R)")
        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "剪切(&X)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
        Set objControl = .Add(xtpControlButton, ID_EDIT_PASTE, "粘贴(&V)")
        Set objControl = .Add(xtpControlButton, ID_EDIT_DELETE, "删除(&D)")
        
        Set subPopup = .Add(xtpControlPopup, 0, "签名与修订(&S)"): subPopup.ID = ID_SIGN: subPopup.BeginGroup = True
        Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_SIGN_QUIT, "签名(&S)")
        Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_UNTREAD, "回退(&C)")
    End With

'------------------------------------------------插入-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "插入(&I)"): cbpPopup.ID = ID_Insert_Menu
    With cbpPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_INSERT_DATETIME, "日期和时间(&D)...")
        Set objControl = .Add(xtpControlButton, ID_INSERT_DATE, "插入日期")
        Set objControl = .Add(xtpControlButton, ID_INSERT_TIME, "插入时间")
        Set objControl = .Add(xtpControlButton, ID_INSERT_SPECIALCHAR, "特殊符号(&S)...")
        Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "图片(&P)")
        Set objControl = .Add(xtpControlButton, ID_INSERT_ELEMENT, "要素(&E)")
        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORT, "历史文件(&H)...")
        Set objControl = .Add(xtpControlButton, ID_INSERT_EPRDEMO, "导入范文(&F)...")
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTINHERITROW, "插入继承行(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTINHERITCOL, "插入继承列(&C)")
    End With

'------------------------------------------------表格-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "表格(&T)"): cbpPopup.ID = ID_TABLE_INSERTTABLE
    With cbpPopup.CommandBar.Controls
        Set subPopup = .Add(xtpControlPopup, 0, "行(&R)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_FORMATROWHEIGHT, "行高(&R)...")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_SAMEROWHEIGHT, "相同行高(&S)")
            
        Set subPopup = .Add(xtpControlPopup, 0, "列(&C)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_FORMATCOLWIDTH, "列宽(&C)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_SAMECOLWIDTH, "相同列宽(&S)")
        
        Set subPopup = .Add(xtpControlPopup, 0, "插入(&I)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTCOLLEFT, "列(在左侧)(&L)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTCOLRIGHT, "列(在右侧)(&T")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTROWUP, "行(在上方)(&B)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTROWDOWN, "行(在下方)(&A)")

        
        Set subPopup = .Add(xtpControlPopup, 0, "删除(&D)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETECOL, "列(&C)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETEROW, "行(&R)")
        
        Set subPopup = .Add(xtpControlPopup, 0, "格式(&F)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_FONT, "字体(&F)...")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_BOLD, "粗体(&B)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_ITALIC, "斜体(&I)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_UNDERLINE, "下划线")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_PROTECT, "保护(&P)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_MERGE, "合并(&M)")
            
        Set subPopup = .Add(xtpControlSplitButtonPopup, ID_FORMAT_FORECOLOR, "字体颜色")
            Set objCustControl = subPopup.CommandBar.Controls.Add(xtpControlCustom, 0, ""): objCustControl.Handle = ColorForeColor.hWnd
            
        Set subPopup = .Add(xtpControlPopup, ID_TABLE_CELLALIGNMENT, "对齐方式(&A)")
            subPopup.CommandBar.SetTearOffPopup "对齐方式", ID_TABLE_CELLALIGNMENT, 100
            subPopup.CommandBar.SetPopupToolBar True
            subPopup.BeginGroup = True
            subPopup.CommandBar.Width = 70
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "靠上左对齐(&1)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "靠上居中(&2)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "靠上右对齐(&3)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "中部左对齐(&4)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "中部居中(&5)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "中部右对齐(&6)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "靠下左对齐(&7)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "靠下居中(&8)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "靠下右对齐(&9)"
        
        Set objControl = .Add(xtpControlButton, ID_TABLE_BORDERSTYLE, "边框样式(&B)")
    End With

'------------------------------------------------帮助-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "帮助(&H)")
    With cbpPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "帮助主题(&H)")
        Set cbpPopup = .Add(xtpControlPopup, 0, "&Web上的" & gstrProductName)
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_HELP_ONLINE, gstrProductName & "主页(&H)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_HELP_WEBFORUM, gstrProductName & "论坛(&F)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_HELP_CONTACT, "发送反馈(&M)")
            If App.LogMode = 0 Then Call cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DesignTest, "测试按扭")
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "关于(&A)...")
    End With

    '## 工具栏初始化
    Call MainToolbarDefine
    '菜单隐藏与Key绑定
    Call MainbarHidBinding
End Sub
Private Sub MainToolbarDefine()
Dim Bar常用 As CommandBar                           '常用工具栏
Dim Bar格式 As CommandBar                           '格式工具栏
Dim Bar表格 As CommandBar                            '表格工具栏
Dim Bar签名 As CommandBar                           '签名、修订与诊断
Dim Combo As CommandBarComboBox                     '工具栏下拉框控件
Dim cbpPopup As CommandBarPopup                     '临时对象
Dim objControl As CommandBarControl                 '工具栏控件
Dim objCustControl As CommandBarControlCustom       '自定义控件'

    Set Bar常用 = cbsMain.Add("常用", xtpBarTop): Bar常用.BarID = ID_Com_Bar
    With Bar常用.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_CLEAR, "清空")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "保存")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE_QUIT, "保存退出")
        
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "打印"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINTPREVIEW, "打印预览")

        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "剪切"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPY, "复制")
        Set objControl = .Add(xtpControlButton, ID_EDIT_PASTE, "粘贴")


        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "撤销")
        Set objControl = .Add(xtpControlButton, ID_EDIT_REDO, "重做")

        Set objControl = .Add(xtpControlButton, ID_INSERT_DATETIME, "插入日期与时间"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_INSERT_DATE, "插入日期")
        Set objControl = .Add(xtpControlButton, ID_INSERT_TIME, "插入时间")
        Set objControl = .Add(xtpControlButton, ID_INSERT_SPECIALCHAR, "插入特殊符号")

        Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "插入图形"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_INSERT_ELEMENT, "诊治要素")

        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORT, "历史文件"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_INSERT_EPRDEMO, "导入范文")
    End With
    
    Set Bar签名 = cbsMain.Add("签名", xtpBarTop): Bar签名.BarID = ID_Sign_Bar
    With Bar签名.Controls
        Set objControl = .Add(xtpControlButton, ID_SIGN_QUIT, "签名"): objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_UNTREAD, "回退"): objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "帮助(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "退出(&Q)"): objControl.Style = xtpButtonIconAndCaption
    End With

    Set Bar格式 = cbsMain.Add("格式", xtpBarTop): Bar格式.BarID = ID_Format_Bar
    With Bar格式.Controls
        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTNAME, "字体", -1, False): Combo.BeginGroup = True
        Dim FontsCol As New Collection, i As Long
        Set FontsCol = GetAllFonts
        For i = 1 To FontsCol.Count
            Combo.AddItem FontsCol.Item(i)
            If FontsCol.Item(i) = "宋体" Then Combo.ListIndex = i
        Next
        Combo.Width = 90: Combo.DropDownWidth = 250: Combo.DropDownListStyle = True: Combo.flags = xtpFlagRightAlign
        
        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTSIZE, "字体尺寸", -1, False)
        '字号列表
        Combo.AddItem "初号", 1: Combo.AddItem "小初", 2: Combo.AddItem "一号", 3: Combo.AddItem "小一", 4
        Combo.AddItem "二号", 5: Combo.AddItem "小二", 6: Combo.AddItem "三号", 7: Combo.AddItem "小三", 8
        Combo.AddItem "四号", 9: Combo.AddItem "小四", 10: Combo.AddItem "五号", 11: Combo.AddItem "小五", 12
        Combo.AddItem "六号", 13: Combo.AddItem "小六", 14: Combo.AddItem "七号", 15: Combo.AddItem "八号", 16
        Combo.AddItem 5, 17:    Combo.AddItem 5.5, 18:      Combo.AddItem 6.5, 19:  Combo.AddItem 7.5, 20
        Combo.AddItem 8, 21:    Combo.AddItem 9, 22:        Combo.AddItem 10, 23:   Combo.AddItem 10.5, 24
        Combo.AddItem 11, 25:   Combo.AddItem 12, 26:       Combo.AddItem 14, 27:   Combo.AddItem 16, 28
        Combo.AddItem 18, 29:   Combo.AddItem 20, 30:       Combo.AddItem 22, 31:   Combo.AddItem 24, 32
        Combo.AddItem 26, 33:   Combo.AddItem 28, 34:       Combo.AddItem 36, 35:   Combo.AddItem 48, 36
        Combo.AddItem 72, 37
        Combo.ListIndex = 12: Combo.Width = 50: Combo.DropDownWidth = 80: Combo.DropDownListStyle = True

        Set objControl = .Add(xtpControlButton, ID_FORMAT_BOLD, "粗体"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FORMAT_ITALIC, "斜体")
        Set objControl = .Add(xtpControlButton, ID_FORMAT_UNDERLINE, "下划线")
        Set objControl = .Add(xtpControlButton, ID_FORMAT_PROTECT, "保护")
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_FORMAT_FORECOLOR, "字体颜色")
            Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
                objCustControl.Handle = ColorForeColor.hWnd
    End With

    Set Bar表格 = cbsMain.Add("表格", xtpBarTop): Bar表格.BarID = ID_Table_Bar
    With Bar表格.Controls
        Set objControl = .Add(xtpControlButton, ID_TABLE_MERGE, "合并单元格")
        
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_TABLE_CELLALIGNMENT, "对齐方式")
        cbpPopup.CommandBar.SetTearOffPopup "单元格对齐方式", ID_TABLE_CELLALIGNMENT, 100
        cbpPopup.CommandBar.SetPopupToolBar True
        cbpPopup.CommandBar.Width = 70
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "靠上左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "靠上居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "靠上右对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "中部左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "中部居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "中部右对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "靠下左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "靠下居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "靠下右对齐"
        
        Set objControl = .Add(xtpControlButton, ID_TABLE_BORDERSTYLE, "边框样式(&B)")
        Set objControl = .Add(xtpControlButton, ID_TABLE_FORMATROWHEIGHT, "行高")
        Set objControl = .Add(xtpControlButton, ID_TABLE_SAMEROWHEIGHT, "相同行高")
        Set objControl = .Add(xtpControlButton, ID_TABLE_FORMATCOLWIDTH, "列宽")
        Set objControl = .Add(xtpControlButton, ID_TABLE_SAMECOLWIDTH, "相同列宽")
        
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTCOLLEFT, "在左侧插入列"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTCOLRIGHT, "在右侧插入列")
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTROWUP, "在上方插入行")
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTROWDOWN, "在下方插入行")
        Set objControl = .Add(xtpControlButton, ID_TABLE_DELETECOL, "删除列")
        Set objControl = .Add(xtpControlButton, ID_TABLE_DELETEROW, "删除行")
    End With

    '工具栏位置调整
    If Screen.Width / Screen.TwipsPerPixelX > 1024 Then
        DockingRightOf cbsMain, Bar表格, Bar签名
        DockingRightOf cbsMain, Bar格式, Bar表格
        DockingRightOf cbsMain, Bar常用, Bar格式
    Else
        DockingRightOf cbsMain, Bar常用, Bar签名
        DockingRightOf cbsMain, Bar格式, Bar表格
    End If
    Bar常用.EnableDocking xtpFlagHideWrap
    Bar签名.EnableDocking xtpFlagHideWrap
    Bar格式.EnableDocking xtpFlagHideWrap
    Bar表格.EnableDocking xtpFlagHideWrap
End Sub
Private Sub MainbarHidBinding()

    'Ctrl热键
    cbsMain.KeyBindings.Add FCONTROL, Asc("S"), ID_FILE_SAVE                '保存
    cbsMain.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT               '打印
    cbsMain.KeyBindings.Add FCONTROL, Asc("Z"), ID_EDIT_UNDO                '撤消
    cbsMain.KeyBindings.Add FCONTROL, Asc("Y"), ID_EDIT_REDO                '重做
    cbsMain.KeyBindings.Add FCONTROL, Asc("X"), ID_EDIT_CUT                 '剪切
    cbsMain.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY                '复制
    cbsMain.KeyBindings.Add FCONTROL, Asc("V"), ID_EDIT_PASTE               '粘贴
    cbsMain.KeyBindings.Add FCONTROL, Asc("N"), ID_FILE_CLEAR               '清空 新建
    cbsMain.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT                '退出
    cbsMain.KeyBindings.Add FCONTROL, Asc("D"), ID_EDIT_SAVEASPHRASE        '存为词句
    cbsMain.KeyBindings.Add FCONTROL, Asc("M"), ID_FILE_SAVEASEPRDEMO       '存为范文
    cbsMain.KeyBindings.Add FCONTROL, Asc("E"), ID_FILE_EXPORTTOXML                 '导出XML
    cbsMain.KeyBindings.Add FCONTROL, Asc("R"), ID_FILE_IMPORTFROMXML               '导入XML
    
    '设计时需要快捷键
    cbsMain.KeyBindings.Add FCONTROL, Asc("B"), ID_FORMAT_BOLD  ' "粗体"
    cbsMain.KeyBindings.Add FCONTROL, Asc("I"), ID_FORMAT_ITALIC ' "斜体")
    cbsMain.KeyBindings.Add FCONTROL, Asc("U"), ID_FORMAT_UNDERLINE ' "下划线")
    cbsMain.KeyBindings.Add FCONTROL, Asc("T"), ID_FORMAT_PROTECT ' "保护")
    cbsMain.KeyBindings.Add FCONTROL, Asc("J"), ID_TABLE_MERGE ' "合并单元格")
    
    'Ctrl+Shift热键
    cbsMain.KeyBindings.Add FCONTROL Or FSHIFT, Asc("S"), ID_FILE_SAVE_QUIT                   '保存退出
    cbsMain.KeyBindings.Add FCONTROL Or FSHIFT, Asc("P"), ID_FILE_PAGESETUP                  '页面设置
    'F热键
    cbsMain.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT                           '帮助
    cbsMain.KeyBindings.Add 0, VK_F2, ID_FILE_PRINTPREVIEW                      '打印预览
    cbsMain.KeyBindings.Add 0, VK_F4, ID_INSERT_DATETIME                        '插入长时间
    cbsMain.KeyBindings.Add FCONTROL, VK_F4, ID_INSERT_DATE                     '插入日期
    cbsMain.KeyBindings.Add FALT, VK_F4, ID_INSERT_TIME                         '插入时间
    cbsMain.KeyBindings.Add FSHIFT, VK_F4, ID_INSERT_SPECIALCHAR                '插入特殊字符
    cbsMain.KeyBindings.Add 0, VK_F6, ID_FILE_IMPORT                            '历史文件
    cbsMain.KeyBindings.Add 0, VK_F5, ID_INSERT_PICTURE                         '插入图片
    cbsMain.KeyBindings.Add 0, VK_F7, ID_INSERT_ELEMENT                         '插入要素
    cbsMain.KeyBindings.Add 0, VK_F9, ID_INSERT_EPRDEMO                         '插入范文
    cbsMain.KeyBindings.Add 0, VK_F11, ID_SIGN_QUIT                             '签名
    cbsMain.KeyBindings.Add FCONTROL, VK_F11, ID_UNTREAD                               '回退
    cbsMain.KeyBindings.Add FCONTROL Or FSHIFT, VK_F1, ID_DesignTest
     
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Error GoTo errHand
    If Control.Visible = False Or Control.Enabled = False Or mblnInit Then Exit Sub
    
    Select Case Control.ID
        Case ID_FILE_CLEAR '清空(&C)")
            If MsgBox("确实要使用新文档覆盖正编辑的文档吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then Call OpenDoc(True)
        Case ID_FILE_SAVE '保存(&S)")
            If SaveDoc Then
                Call Me.ShowMe(mfrmParent, Document, mstrModelPrivate, mblnMoved, mblnCanPrint)
            End If
        Case ID_FILE_SAVE_QUIT '保存退出(&Q)")
            If SaveDoc(, True) Then Unload Me
        Case ID_FILE_SAVEASEPRDEMO '另存为范文(&D)...")
            Call SaveAsDemo
        Case ID_FILE_EXPORTTOXML '导出为XML文件(&E)")
            Call ExportXml
        Case ID_FILE_IMPORTFROMXML '从XML文件导入(&I)")
            Call ImportXml
        Case ID_FILE_PAGESETUP '页面设置(&U)...")
            Call PageSetUp
        Case ID_FILE_PRINTPREVIEW '打印预览(&V)")
            Call PrintDoc(True)
        Case ID_FILE_PRINT '打印(&P)...")
            Call PrintDoc(False)
        Case ID_FILE_EXIT '退出(&X)")
            Unload Me
        Case ID_EDIT_UNDO '撤销(&U)")
            Call ExeUndo
        Case ID_EDIT_REDO '重做(&R)")
        Case ID_EDIT_CUT '剪切(&X)")
            Call ContentMove("Cut")
        Case ID_EDIT_COPY '复制(&C)")
            Call ContentMove("Copy")
        Case ID_EDIT_PASTE '粘贴(&V)")
            Call ContentMove("Paste")
        Case ID_EDIT_DELETE '删除(&D)")
            Call F1Main_KeyDown(vbKeyDelete, 0)
        Case ID_SIGN_QUIT '签名(&S)")
            If SaveDoc(True, True) Then Unload Me
        Case ID_UNTREAD '回退(&C)")
            Call RollBack
        Case ID_VIEW_HEADFOOT '页眉页脚(&H)")
        Case ID_INSERT_DATETIME '日期和时间(&D)...")
            Call InsertOtherText(SelCell.Key, "日期时间")
        Case ID_INSERT_DATE '日期
            Call InsertOtherText(SelCell.Key, "日期")
        Case ID_INSERT_TIME '时间
            Call InsertOtherText(SelCell.Key, "时间")
        Case ID_INSERT_SPECIALCHAR '特殊符号(&S)...")
            Call InsertOtherText(SelCell.Key, "特殊符号")
        Case ID_INSERT_PICTURE '插入图片(&P)")
            Call InsertPicture(SelCell.Key, SelCell.PictureKey)
        Case ID_INSERT_ELEMENT '插入要素(&E)")
            Call InsertElement(SelCell.Key)
        Case ID_FILE_IMPORT '历史文件(&H)...")
        Case ID_INSERT_EPRDEMO '导入范文(&F)...")
            Call ImportDemo
        Case ID_TABLE_FORMATROWHEIGHT '调整行高(&R)...")
            Call SetRowCol("行高")
        Case ID_TABLE_SAMEROWHEIGHT '相同行高(&S)")
            Call SetRowCol("相同行高")
        Case ID_TABLE_FORMATCOLWIDTH '调整列宽(&C)")
            Call SetRowCol("列宽")
        Case ID_TABLE_SAMECOLWIDTH '相同列宽(&S)")
            Call SetRowCol("相同列宽")
        Case ID_TABLE_FORMATCELL '调整单元格属性(&E)...")
        Case ID_TABLE_INSERTCOLLEFT '插入列(在左侧)(&L)")
            Call InsertRowCol("InsertLeftCol")
        Case ID_TABLE_INSERTCOLRIGHT '插入列(在右侧)(&T")
            Call InsertRowCol("InsertRightCol")
        Case ID_TABLE_INSERTROWUP '插入行(在上方)(&A)")
            Call InsertRowCol("InsertUpRow")
        Case ID_TABLE_INSERTROWDOWN '插入行(在下方)(&B)")
            Call InsertRowCol("InsertDnRow")
        Case ID_TABLE_INSERTINHERITROW '插入插入继承行(&R)")
            Call InsertInherit("Row")
        Case ID_TABLE_INSERTINHERITCOL '插入插入继承列(&C)")
            Call InsertInherit("Col")
        Case ID_TABLE_DELETECOL '删除列(&C)")
            Call DeleteRowCol("Col")
        Case ID_TABLE_DELETEROW '删除行(&R)")
            Call DeleteRowCol("Row")
        Case ID_FORMAT_FONT '字体(&F)...")
            Call SetCellFont
        Case ID_FORMAT_FONTSIZE '字号
            Call SetCellFormat("字号", Control.Text)
        Case ID_FORMAT_FONTNAME '字体名
            Call SetCellFormat("字体名称", Control.Text)
        Case ID_FORMAT_BOLD '粗体(&B)")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("粗体", Control.Checked)
        Case ID_FORMAT_ITALIC '斜体(&I)")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("斜体", Control.Checked)
        Case ID_FORMAT_UNDERLINE '下划线")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("下划线", Control.Checked)
        Case ID_FORMAT_PROTECT '保护(&P)")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("保护", Control.Checked)
        Case ID_TABLE_MERGE '合并(&M)")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("合并", Control.Checked)
        Case ID_FORMAT_FORECOLOR ' "字体颜色")
            Call ColorForeColor_pOK(False)
        Case ID_TABLE_CELLALIGNMENT1 '靠上左对齐"
            Call SetCellFormat("靠上左对齐", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT2 '靠上居中"
            Call SetCellFormat("靠上居中", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT3 '靠上右对齐"
            Call SetCellFormat("靠上右对齐", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT4 '中部左对齐"
            Call SetCellFormat("中部左对齐", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT5 '中部居中"
            Call SetCellFormat("中部居中", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT6 '中部右对齐"
            Call SetCellFormat("中部右对齐", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT7 '靠下左对齐"
            Call SetCellFormat("靠下左对齐", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT8 '靠下居中"
            Call SetCellFormat("靠下居中", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT9 '靠下右对齐"
            Call SetCellFormat("靠下右对齐", Control.Checked)
        Case ID_TABLE_BORDERSTYLE '边框样式
            Call SetCellBorder
        Case ID_HELP_CONTENT '帮助主题(&H)")
            ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
        Case ID_HELP_ONLINE ' gstrProductName & "主页(&H)")
            Call zlHomePage(Me.hWnd)
        Case ID_HELP_WEBFORUM ' gstrProductName & "论坛(&F)")
            Call zlWebForum(Me.hWnd)
        Case ID_HELP_CONTACT '发送反馈(&M)")
            Call zlMailTo(Me.hWnd)
        Case ID_HELP_ABOUT '关于(&A)...")
            ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
        Case ID_DesignTest '设计环境测试按扭
            Call DesignTest
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    On Error Resume Next
    Processing.Width = stbThis.Panels("PatInfo").Width
    Processing.Left = stbThis.Panels("PatInfo").Left
    Processing.Top = stbThis.Top + 60
    Err.Clear
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mblnInit = True Then Exit Sub

    Select Case Control.ID
        Case ID_FILE_CLEAR '清空(&C)")
            Control.Enabled = editType <> TabET_单病历审核
        Case ID_FILE_SAVE '保存(&S)")
            Control.Enabled = mReadOnly = 0
        Case ID_FILE_SAVE_QUIT '保存退出(&Q)")
            Control.Enabled = mReadOnly = 0
        Case ID_FILE_SAVEASEPRDEMO '另存为范文(&D)...")
            Control.Visible = editType <> TabET_病历文件定义 And mReadOnly <> 2
        Case ID_EDIT_SAVEASPHRASE '存为词句
            Control.Visible = False
        Case ID_FILE_EXPORTTOXML '导出为XML文件(&E)")
        Case ID_FILE_IMPORTFROMXML '从XML文件导入(&I)")
            Control.Enabled = editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑 Or (editType = TabET_单病历编辑 And EditMode = TabEm_新增)
        Case ID_FILE_PAGESETUP '页面设置(&U)...")
            Control.Enabled = mblnCanPrint And mReadOnly <> 2
        Case ID_FILE_PRINTPREVIEW '打印预览(&V)")
            Control.Enabled = mblnCanPrint
        Case ID_FILE_PRINT '打印(&P)...")
            Control.Enabled = mblnCanPrint
        Case ID_FILE_EXIT '退出(&X)")
        Case ID_EDIT_UNDO '撤销(&U)")
            Control.Enabled = Undo.Count > 0 And mReadOnly <> 2 And (Not Doc.Visible) And (Not elEdit.Visible) And (Not PicEdit.Visible)
            If Control.Enabled Then
                Control.ToolTipText = "撤销 " & Undo(Undo.Count).Row & "行 " & Undo(Undo.Count).Col & "列 " & _
                        Decode(Undo(Undo.Count).CT, cprCTFixtext, "文本", cprCTText, "文本", cprCTElement, "要素", cprCTTextElement, "混合编辑", cprCTPicture, "参考图", cprCTReportPic, "报告图") & "变化"
            Else
                Control.ToolTipText = "撤销以单元格为最小单位的内容变化"
            End If
        Case ID_EDIT_REDO '重做(&R)")
            Control.Visible = False
        Case ID_EDIT_CUT '剪切(&X)")
            Control.Enabled = (SelCell.对象类型 = cprCTText Or SelCell.对象类型 = cprCTTextElement Or SelCell.对象类型 = cprCTFixtext)
        Case ID_EDIT_COPY '复制(&C)")
            Control.Enabled = (SelCell.对象类型 = cprCTText Or SelCell.对象类型 = cprCTTextElement Or SelCell.对象类型 = cprCTFixtext)
        Case ID_EDIT_PASTE '粘贴(&V)")
            Control.Enabled = (SelCell.对象类型 = cprCTText Or SelCell.对象类型 = cprCTTextElement Or SelCell.对象类型 = cprCTFixtext)
        Case ID_EDIT_DELETE '删除(&D)")
            Control.Enabled = (SelCell.对象类型 = cprCTFixtext Or SelCell.对象类型 = cprCTText Or (SelCell.对象类型 = cprCTTextElement And Doc.Visible) Or SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic)
        Case ID_SIGN_QUIT '签名(&S)")
            Control.Visible = Not (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then Control.Enabled = mReadOnly = 0
        Case ID_UNTREAD '回退(&C)")
            Control.Visible = Not (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then Control.Enabled = mReadOnly < 2 '0-正常 1-签名后点修改 2-编辑界面点打开查阅
        Case ID_INSERT_DATETIME '日期和时间(&D)...")
            Control.Enabled = (SelCell.对象类型 = cprCTText Or (SelCell.对象类型 = cprCTTextElement And Doc.Visible))
        Case ID_INSERT_DATE '日期
            Control.Enabled = (SelCell.对象类型 = cprCTText Or (SelCell.对象类型 = cprCTTextElement And Doc.Visible))
        Case ID_INSERT_TIME '时间
            Control.Enabled = (SelCell.对象类型 = cprCTText Or (SelCell.对象类型 = cprCTTextElement And Doc.Visible))
        Case ID_INSERT_SPECIALCHAR '特殊符号(&S)...")
            Control.Enabled = (SelCell.对象类型 = cprCTFixtext Or SelCell.对象类型 = cprCTText Or (SelCell.对象类型 = cprCTTextElement And Doc.Visible))
        Case ID_INSERT_PICTURE '图片(&P)")
            Control.Enabled = SelCell.对象类型 = cprCTPicture
            If Control.Enabled Then Control.Enabled = editType <> TabET_单病历审核
        Case ID_INSERT_ELEMENT '要素(&E)")
            Control.Enabled = (SelCell.对象类型 = cprCTElement And (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)) Or (SelCell.对象类型 = cprCTTextElement And Doc.Visible)
        Case ID_FILE_IMPORT '历史文件(&H)...")
             Control.Visible = False '暂时不开，有时间再处理
'            Control.Visible = editType = TabET_全文示范编辑 Or (editType = TabET_单病历编辑 And EditMode = TabEm_新增)
        Case ID_INSERT_EPRDEMO '导入范文(&F)...")
            Control.Visible = editType = TabET_全文示范编辑 Or (editType = TabET_单病历编辑 And EditMode = TabEm_新增)
        Case ID_EDIT_SAVEASPHRASE '存为词句
            Control.Visible = False
        Case ID_TABLE_FORMATROWHEIGHT, ID_TABLE_SAMEROWHEIGHT, ID_TABLE_FORMATCOLWIDTH, ID_TABLE_SAMECOLWIDTH '行高,相同行高 列宽 相同列宽
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
            End If
        Case ID_TABLE_INSERTCOLLEFT, ID_TABLE_INSERTCOLRIGHT, ID_TABLE_INSERTROWUP, ID_TABLE_INSERTROWDOWN '新增列
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
            End If
        Case ID_TABLE_DELETECOL, ID_TABLE_DELETEROW
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)  '删除行列
                If Control.Enabled Then Control.Enabled = Not mblnClickZ
            End If
        Case ID_TABLE_INSERTINHERITROW '插入继承行
            Control.Visible = editType = TabET_单病历审核 And mReadOnly <> 2
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
            End If
        Case ID_TABLE_INSERTINHERITCOL '插入继承列(&R)")
            Control.Visible = False '暂时不开，有时间再处理
        Case ID_FORMAT_FONT '字体(&F)...")
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
            End If
        Case ID_FORMAT_FONTNAME '字体名称
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
            End If
            If Control.Visible And Control.Enabled Then
                Control.Text = SelCell.FontName
            End If
        Case ID_FORMAT_FONTSIZE '字号
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
            End If
            If Control.Visible And Control.Enabled Then
                Control.Text = GetFontSizeChinese(SelCell.FontSize)
            End If
        Case ID_FORMAT_BOLD '粗体(&B)")
            If (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑) Then
                Control.Visible = True
                If Control.Visible Then
                    Control.Enabled = Not (SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic)
                    If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                    If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
                End If
                If Control.Visible And Control.Enabled Then
                    Control.Checked = SelCell.FontBold
                End If
            Else
                Control.Visible = False
                If Not Control.Parent Is Nothing Then Control.Parent.Visible = False
            End If
        Case ID_FORMAT_ITALIC '斜体(&I)")
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
            End If
            If Control.Visible And Control.Enabled Then
                Control.Checked = SelCell.FontItalic
            End If
        Case ID_FORMAT_UNDERLINE '下划线")
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
            End If
            If Control.Visible And Control.Enabled Then
                Control.Checked = SelCell.FontUnderline
            End If
        Case ID_FORMAT_PROTECT '保护(&P)")
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
            End If
            If Control.Visible And Control.Enabled Then
                Control.Checked = SelCell.保留对象
            End If
        Case ID_TABLE_MERGE '合并(&M)")
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
                Control.Checked = SelCell.Merge
            End If
        Case ID_FORMAT_FORECOLOR '字体颜色
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
            End If
        Case ID_TABLE_CELLALIGNMENT '对齐方式
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                If (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignLeft) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT1
                ElseIf (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignCenter) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT2
                ElseIf (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignRight) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT3
                ElseIf (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignLeft) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT4
                ElseIf (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignCenter) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT5
                ElseIf (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignRight) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT6
                ElseIf (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignLeft) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT7
                ElseIf (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignCenter) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT8
                ElseIf (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignRight) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT9
                End If
            End If
            Control.Enabled = True
            If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
            If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
        Case ID_TABLE_CELLALIGNMENT1 '靠上左对齐"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignLeft)
            End If
        Case ID_TABLE_CELLALIGNMENT2 '靠上居中"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignCenter)
            End If
        Case ID_TABLE_CELLALIGNMENT3 '靠上右对齐"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignRight)
            End If
        Case ID_TABLE_CELLALIGNMENT4 '中部左对齐"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignLeft)
            End If
        Case ID_TABLE_CELLALIGNMENT5 '中部居中"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignCenter)
            End If
        Case ID_TABLE_CELLALIGNMENT6 '中部右对齐"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignRight)
            End If
        Case ID_TABLE_CELLALIGNMENT7 '靠下左对齐"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignLeft)
            End If
        Case ID_TABLE_CELLALIGNMENT8 '靠下居中"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignCenter)
            End If
        Case ID_TABLE_CELLALIGNMENT9 '靠下右对齐"
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignRight)
            End If
        Case ID_TABLE_BORDERSTYLE '边框样式
            If (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑) Then
                Control.Visible = True
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '不是通过点击第0行/列选择,通过点击第0行/列难以控制
            Else
                Control.Visible = False
                If Not Control.Parent Is Nothing Then Control.Parent.Visible = False
            End If
        Case ID_HELP_CONTENT '帮助主题(&H)")
        Case ID_HELP_ONLINE ' gstrProductName & "主页(&H)")
        Case ID_HELP_WEBFORUM ' gstrProductName & "论坛(&F)")
        Case ID_HELP_CONTACT '发送反馈(&M)")
        Case ID_HELP_ABOUT '关于(&A)...")
        Case ID_TABLE_INSERTTABLE      '表格菜单
            Control.Visible = (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
        Case ID_SIGN        '格式和表格制作工具条
            Control.Visible = Not (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑)
    End Select
End Sub
Private Sub MainDockPaneDefine()
    '初始界面布局
    Dim PaneSentence As Pane
    Dim PaneAttribute As Pane
    Dim PaneHistory As Pane
    Dim PaneContent As Pane
    Dim PanePacsPic As Pane
    
    With Me.dkpMain
        .SetCommandBars cbsMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
    
    Set PaneSentence = dkpMain.CreatePane(conPane_SentenceList, 200, 0, DockLeftOf, Nothing)
    PaneSentence.Title = "词句示范"
    PaneSentence.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set PaneAttribute = dkpMain.CreatePane(conPane_Attribute, 200, 0, DockBottomOf, PaneSentence)
    PaneAttribute.Title = "属性"
    PaneAttribute.Options = PaneNoCloseable Or PaneNoFloatable
    dkpMain.AttachPane PaneAttribute, PaneSentence
    
    Set PaneHistory = dkpMain.CreatePane(conPane_History, 200, 0, DockBottomOf, PaneSentence)
    PaneHistory.Title = "历史版本"
    PaneHistory.Options = PaneNoCloseable Or PaneNoFloatable
    dkpMain.AttachPane PaneHistory, PaneSentence
    
    Set PanePacsPic = dkpMain.CreatePane(conPane_PacsPic, 200, 0, DockBottomOf, PaneSentence)
    PanePacsPic.Title = "PACS报告图"
    PanePacsPic.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption
    dkpMain.AttachPane PanePacsPic, PaneSentence
    
    PaneSentence.MaxTrackSize.Width = 200:  PaneAttribute.MaxTrackSize.Width = 200
    PaneHistory.MaxTrackSize.Width = 200:   PanePacsPic.MaxTrackSize.Width = 200

    Set PaneContent = dkpMain.CreatePane(conPane_Content, 1080, 0, DockRightOf, Nothing)
    PaneContent.Title = "表格全文"
    PaneContent.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable
End Sub

Private Sub chkType_Click(Index As Integer)
Dim i As Integer, strAttribute As String, blnReturn As Boolean, IntOld As Integer
    If mblnShowAtt Or mblnInit Then Exit Sub
    For i = 0 To chkType.UBound '记下原先类型
        If chkType(i).Value Then IntOld = i: Exit For
    Next
    Call F1Main_GotFocus
    Call SetCellAttribute(Index, strAttribute, blnReturn)
    If Not blnReturn Then
        Index = IntOld
        MsgBox strAttribute, vbInformation, gstrSysName
        strAttribute = ""
    End If
    DoEvents
    Call ShowAttr(Index, strAttribute)
End Sub
Private Sub ShowAttr(ByVal IntType As Integer, ByVal strAttribute As String)
'显示单无格属性及说明
Dim i As Integer
    mblnShowAtt = True
    For i = 0 To chkType.UBound '设置类型
        If i = IntType Then
            chkType(i).Value = vbChecked: Text1(i).Visible = True
        Else
            chkType(i).Value = vbUnchecked: Text1(i).Visible = False
        End If
    Next

    If (IntType = 0 Or IntType = 1) Then
        cmdApply.Visible = True: txtSum.Visible = True: txtSum.Locked = False: shpTxtSum.Visible = True
'        cmdSum.Visible = True: cmdAvg.Visible = True
        If InStr(strAttribute, ";") > 0 Then '合计单元格
            txtSum.Text = strAttribute
            txtSum.Locked = False
        ElseIf InStr(strAttribute, ",") > 0 Then '源单元格,不能嵌套合计
            txtSum.Text = strAttribute
            txtSum.Locked = True
        Else                                '无合计属性单元格
            txtSum.Text = ""
            txtSum.Locked = False
        End If
    Else
        cmdApply.Visible = False: txtSum.Visible = False: shpTxtSum.Visible = False
'        cmdSum.Visible = True: cmdAvg.Visible = True
    End If
    mblnShowAtt = False
End Sub
Private Sub chkType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If chkType(Index).Enabled And (editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑) Then
        If Not SelCell Is Nothing Then
            If SelCell.对象类型 = cprCTPicture Or SelCell.对象类型 = cprCTReportPic Then '以便从有图且被选中的情况下转换成其它类型
                chkType(Index).SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmdApply_Click()
Dim i As Integer, strAttribute As String, blnReturn As Boolean
    If txtSum.Text <> "" Then
        If UBound(Split(txtSum.Text, ";")) < 1 Then '确保必须有一个以上单元格组成
            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      格式不正确，格式组成在下方有说明描述" & vbCrLf & "      请检查！", True, 1
            Exit Sub
        End If
        
        For i = 0 To UBound(Split(txtSum.Text, ";"))
            If UBound(Split(Split(txtSum.Text, ";")(i), ",")) <> 1 Then '确保组成合计的单元格有效
                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      格式不正确，格式组成在下方有说明描述" & vbCrLf & "      请检查！", True, 1
                Exit Sub
            End If
        Next
    End If
    strAttribute = Trim(txtSum.Text) '可以是空,表示取消单元格的合计属性
    If Not SetSumAtt(strAttribute) Then
        MsgBox strAttribute, vbInformation, gstrSysName
        txtSum.SelStart = 0: txtSum.SelLength = Len(txtSum): txtSum.SetFocus
    Else
        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提示：" & vbCrLf & "      合计属性设定成功", True, 0
    End If
End Sub

Private Sub cmdSum_Click()
    If cmdSum.Value = vbChecked Then
        F1Main.MousePointer = vbCustom
        F1Main.MouseIcon = cmdSum.MouseIcon
        mbFunType = 0
        mfrmTipInfo.ShowTipInfo txtSum.hWnd, "提示：" & vbCrLf & "      请指定当前单元格由哪些单元格合计组成。", True, 0
    End If
End Sub

Private Sub ColorForeColor_pOK(ByVal ControlSelf As Boolean)
    Call SetCellFormat("字体颜色", ColorForeColor.Color)
    Call SetColorIcon(ID_FORMAT_FORECOLOR, ColorForeColor.Color)
    If ControlSelf Then SendKeys "{ESCAPE}"
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Item.ID
    Case conPane_SentenceList
        Item.Handle = mfrmSentence.hWnd
    Case conPane_Attribute
        Item.Handle = picAtt.hWnd
    Case conPane_History
        Item.Handle = picHistory.hWnd
    Case conPane_PacsPic
        Item.Handle = mfrmPacsPic.hWnd
    Case conPane_Content
        Item.Handle = picMainBack.hWnd
    End Select
    Err.Clear
End Sub
Private Sub InitForms()
    Set mfrmSentence = New frmSentenceList: mfrmSentence.mstrPrivs = mstrPrivs
    Set mfrmMainError = New frmMainMsg
    Set mfrmPacsPic = New frmPACSImg
End Sub
Private Sub RefreshF1Main()
Dim lngRow As Long, lngCol As Long, lngCell As Long, vCell As F1CellFormat, lngCount As Long, strShow As String
    On Error GoTo errHand
    With F1Main
        .DeleteRange .MinRow, .MinCol, .MaxRow, .MaxCol, F1ShiftRows
        .ShowTabs = F1TabsOff
        .AllowMoveRange = False '移动选中区域
        .AllowFillRange = False '拖动范围赋值,无事件不可控制
        .AllowInCellEditing = False '单元格编辑
        .AllowEditHeaders = False '编辑列头
        .AllowDesigner = False  '允许设计
        .AllowDelete = False '提示是英文的，最好不要允许而自已通过KeyDown控制
        .ShowLockedCellsError = False '对锁定单元格进行编辑时的消息提示
        .ScrollToLastRC = False '允许滚动到最后一个单元格
        .ColWidthUnits = F1ColWidthUnitsTwips '列宽计算单位为堤
        .ShowSelections = F1On   '当焦点不在控件上时，直接单击单元格即选中
        .DefaultFontName = "宋体"
        .DefaultFontSize = 9
        .MaxCol = Me.Document.Cells.Cols
        .MaxRow = Me.Document.Cells.Rows
        
        If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
            .ShowColHeading = False '显示固定列
            .ShowRowHeading = False '显示固定行
            .AllowResize = True    '拖动、自动换行时 改变行高列宽
        Else
            .ShowColHeading = True
            .ShowRowHeading = True
            .HdrHeight = 300
            .HdrWidth = 300
            .AllowResize = True
        End If
        
        F1Main.SetSelection 1, 1, F1Main.MaxRow, F1Main.MaxCol
        F1Main.SetAlignment F1HAlignJustify, True, F1VAlignTop, 0
        '定行高列宽
        For lngRow = 1 To .MaxRow
            .RowHeight(lngRow) = Me.Document.Cells.Cell(lngRow, 1).Height
        Next
        For lngCol = 1 To .MaxCol
            .ColWidthTwips(lngCol) = Me.Document.Cells.Cell(1, lngCol).Width
            .ColText(lngCol) = lngCol '列头显示数字
        Next
        
        lngCount = Me.Document.Cells.Count
        If Me.Visible Then Processing.Max = lngCount
        For lngCell = 1 To lngCount
            If Me.Visible Then Processing.Value = lngCell: Processing.Visible = True
            
            lngRow = Me.Document.Cells(lngCell).Row: lngCol = Me.Document.Cells(lngCell).Col
            With Me.Document.Cells.Cell(lngRow, lngCol)
                '指定区域
                If .Merge And InStr(.MergeRange, ";") > 0 Then 'MergeRange数据格式 (左上方)行,列;(右下方)行,列
                    F1Main.SetSelection Split(Split(.MergeRange, ";")(0), ",")(0), Split(Split(.MergeRange, ";")(0), ",")(1), Split(Split(.MergeRange, ";")(1), ",")(0), Split(Split(.MergeRange, ";")(1), ",")(1)
                Else
                    F1Main.SetSelection lngRow, lngCol, lngRow, lngCol
                End If
                Set vCell = F1Main.CreateNewCellFormat
                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then '只有合并单元格首个或非合并单元格才刷新
'                    vCell.ProtectionLocked = .保留对象  '是否锁定,保护区域,行控,列控,签名、保存时写入Database
                    vCell.MergeCells = .Merge
                    vCell.WordWrap = True
                    vCell.FontName = .FontName          '字体>宋体</字体>
                    vCell.FontSize = .FontSize          '<字号>9</字号>
                    vCell.FontBold = .FontBold          '<粗体>False</粗体>
                    vCell.FontItalic = .FontItalic        '<斜体>False</斜体>
                    vCell.FontUnderline = .FontUnderline     '<下划线>False</下划线>
                    vCell.FontStrikeout = .FontStrikeout    '<删除线>False</删除线>
                    vCell.FontColor = .FontColor         '<字体颜色>vbblack</字体颜色>
                    vCell.AlignHorizontal = .HAlignment       '<横向对齐>F1HAlignCenter</横向对齐>
                    vCell.AlignVertical = .VAlignment       '<纵向对齐>F1VAlignCenter</纵向对齐>

                    Select Case .对象类型
                        Case cprCTFixtext    '0-固定文本(不可编辑)
                            F1Main.TextRC(lngRow, lngCol) = .内容文本
                        Case cprCTText '1-文本型(可编辑多行文本)
                            F1Main.TextRC(lngRow, lngCol) = .内容文本
                            If editType = TabET_单病历审核 Then DocOld.Add .内容文本, .Key
                        Case cprCTElement    '2-单要素
                            If editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑 Then
                                If .ElementKey <> "" Then
                                    If Document.Elements("K" & .ElementKey).输入形态 = 1 And Document.Elements("K" & .ElementKey).要素类型 <> 2 Then '输入形态=展开
                                        F1Main.TextRC(lngRow, lngCol) = Document.Elements("K" & .ElementKey).内容文本
                                    Else
                                        F1Main.TextRC(lngRow, lngCol) = "[" & Document.Elements("K" & .ElementKey).要素名称 & "]" & Document.Elements("K" & .ElementKey).要素单位
                                    End If
                                End If
                            Else
                                strShow = ""
                                If .内容文本 = "" Then
                                    If Document.Elements("K" & .ElementKey).替换域 = 1 Then '自动替换要素
                                        strShow = GetReplaceEleValue(Document.Elements("K" & .ElementKey).要素名称, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.病人来源, Document.EPRPatiRecInfo.医嘱id, Document.EPRPatiRecInfo.婴儿)
                                        If strShow = "" And Not Document.Elements("K" & .ElementKey).自动转文本 Then '没取到值，是否自动转换成文本(空)
                                            strShow = "[" & Document.Elements("K" & .ElementKey).要素名称 & "]" & Document.Elements("K" & .ElementKey).要素单位
                                        Else
                                            Document.Elements("K" & .ElementKey).内容文本 = strShow
                                            .内容文本 = strShow & Document.Elements("K" & .ElementKey).要素单位
                                            strShow = .内容文本
                                        End If
                                    Else
                                        If Document.Elements("K" & .ElementKey).输入形态 = 1 And Document.Elements("K" & .ElementKey).要素类型 <> 2 Then '输入形态=展开
                                            .内容文本 = Document.Elements("K" & .ElementKey).内容文本 & Document.Elements("K" & .ElementKey).要素单位
                                            strShow = .内容文本
                                        Else
                                            strShow = "[" & Document.Elements("K" & .ElementKey).要素名称 & "]" & Document.Elements("K" & .ElementKey).要素单位
                                        End If
                                    End If
                                    F1Main.TextRC(lngRow, lngCol) = strShow
                                Else
                                    F1Main.TextRC(lngRow, lngCol) = .内容文本
                                End If
                            End If
                            If editType = TabET_单病历审核 Then DocOld.Add .内容文本, .Key
                        Case cprCTTextElement '3-文本与多要素混合编辑
                            GetTextELement .Key     '跟据Text Element填写F1Main中的单元格及类的内容文本
                            If editType = TabET_单病历审核 Then DocOld.Add .内容文本, .Key
                        Case cprCTReportPic, cprCTPicture    '5-报告图
                            If Me.Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                Call PaintPictureOnTable(.Key)
                                If editType = TabET_单病历审核 Then DocOld.Add Document.Pictures("K" & .PictureKey).OrigPic.Handle, .Key
                            Else
                                If editType = TabET_单病历审核 Then DocOld.Add 0, .Key
                            End If
                            F1Main.TextRC(lngRow, lngCol) = IIf(.对象类型 = cprCTPicture, "参考图", "报告图")
                        Case cprCTSign         '6-签名'签名在设计时仅为占位,无实际信息；没有签名时终止版=0；普通签名后审核时不显示，以便再次签名；行控/列控签名后审核时要显示
                            strShow = ""
                            If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
                                Select Case mReadOnly 'mReadOnly 0-正常,1-签名后点修改,2-主界面打开查阅或查阅历次签名版本
                                    Case 0
                                        If .终止版 <> 0 Then
                                            With Document.Signs("K" & .SignKey)
                                                strShow = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
                                                strShow = strShow & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
                                            End With
                                        Else
                                            strShow = "[签名位]"
                                        End If
                                    Case 1, 2
                                        If .终止版 = 0 Then
                                            strShow = "[签名位]"
                                        Else
                                            With Document.Signs("K" & .SignKey)
                                                strShow = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
                                                strShow = strShow & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
                                            End With
                                        End If
                                End Select
                            Else
                                strShow = "[签名位]"
                            End If
                            F1Main.TextRC(lngRow, lngCol) = strShow '前置文字 & 姓名 & 显示手签 & 显示时间<>""(format(签名时间,显示时间)
                        Case cprCTRowSign, cprCTColSign '7-行控签名 '8-列控签名
                            strShow = ""
                            If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
                                If .终止版 <> 0 Then
                                    With Document.Signs("K" & .SignKey)
                                        strShow = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
                                        strShow = strShow & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
                                    End With
                                Else
                                    strShow = "[签名位]"
                                End If
                            Else
                                strShow = "[签名位]"
                            End If
                            F1Main.TextRC(lngRow, lngCol) = strShow '前置文字 & 姓名 & 显示手签 & 显示时间<>""(format(签名时间,显示时间)
                    End Select
                    F1Main.SetCellFormat vCell
                    Call F1Main.SetBorder(-1, .CellLineLeft, .CellLineRight, .CellLineTop, .CellLineBottom, 0, -1, .CellLineLeftColor, .CellLineRightColor, .CellLineTopColor, .CellLineBottomColor)
                End If
            End With
        Next
        Processing.Visible = False
    End With
    
'    F1Main.SetSelection F1Main.MaxRow, F1Main.MaxCol, F1Main.MaxRow, F1Main.MaxCol
    F1Main.SetSelection 1, 1, 1, 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Doc_BeforeKeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
    If KeyCode = vbKeyReturn And Shift = 0 Then '直接回车表示退出编辑
        KeyCode = 0
        F1Main_GotFocus
        F1Main.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyReturn And Shift = 4 Then 'ALT+回车表示换行
'        SendKeys "^~"
    End If
    If Shift <> 0 Then Exit Sub
    
    Select Case KeyCode
    Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
        vbKeyEscape, vbKeyTab, vbKeyDelete, vbKeyBack, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
        vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
        vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12
        Exit Sub
    End Select

    With Doc
        If .SelLength > 0 Then .Range(Doc.Selection.EndPos, .Selection.EndPos).Selected
        i = .Selection.StartPos
        If i = 0 Then
            .Range(i, i).Font.Protected = False
            .Range(i, i).Font.Hidden = False
        ElseIf .Range(i - 1, i).Font.Hidden And _
            .Range(i, i + 1).Font.Hidden = False And _
            .Range(i, i + 1).Font.Protected = False Then
            'A问题：（隐藏文本）|普通文本
            .Range(i, i).Font.Protected = False
            .Range(i, i).Font.Hidden = False
        ElseIf .Range(i - 1, i).Font.Protected And .Range(i - 1, i).Font.Hidden = False And .Range(i, i + 1).Font.Hidden And .Range(i, i + 3).Text = "EE(" Then
            'B问题1：（隐藏关键字）[要素]|（隐藏关键字）（隐藏关键字）[要素]（隐藏关键字）
            If KeyCode = vbKeySpace Then KeyCode = 0
            .Range(i + 16, i + 16).Text = " "
            .Range(i + 16, i + 17).Font.Protected = False
            .Range(i + 16, i + 17).Font.Hidden = False
            .Range(i + 17, i + 17).Selected
        ElseIf .Range(i - 1, i).Font.Hidden And .Range(i, i + 1).Font.Protected And .Range(i, i + 1).Font.Hidden = False And (i - 16 <> 0) Then
            'B问题2：（隐藏关键字）[要素]（隐藏关键字）（隐藏关键字）|[要素]（隐藏关键字）
            If KeyCode = vbKeySpace Then KeyCode = 0
            .Range(i - 16, i - 16).Text = " "
            .Range(i - 16, i - 15).Font.Protected = False
            .Range(i - 16, i - 15).Font.Hidden = False
            .Range(i - 15, i - 15).Selected
        ElseIf i - 16 = 0 And .Range(i - 1, i).Font.Hidden And .Range(i - 1, i).Font.Protected And .Range(i, i + 1).Font.Hidden = False Then
            '问题2：0（隐藏关键字）|[要素]（隐藏关键字）
            .Range(i - 16, i - 16).Font.Protected = False
            .Range(i - 16, i - 16).Font.Hidden = False
            .Range(i - 16, i - 16).Selected
        End If
    End With
End Sub



Private Sub Doc_Change()
    With Doc
        .Range(0, Len(.Text)).Font.Name = SelCell.FontName
        .Range(0, Len(.Text)).Font.Size = SelCell.FontSize
        .Range(0, Len(.Text)).Font.Italic = SelCell.FontItalic
        .Range(0, Len(.Text)).Font.Bold = SelCell.FontBold
        .Range(0, Len(.Text)).Font.ForeColor = SelCell.FontColor
        .Range(0, Len(.Text)).Font.Underline = IIf(SelCell.FontUnderline, cprHair, cprNone)
        .Range(0, Len(.Text)).Font.Strikethrough = SelCell.FontStrikeout
    End With
End Sub


Private Sub Doc_DblClick()
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
Dim pt As POINTAPI, lHheight As Long, lHwidth As Long
    pt.x = 0: pt.y = 0
    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '固定行高度
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '固定列宽度
    bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys And sType = "E" Then
        If Document.Elements("K" & lKey).输入形态 = 1 Then Exit Sub
        ShowElInDoc lSE, lES, lKey
    End If
End Sub
Private Sub ShowElInDoc(ByVal lSE As Long, ByVal lES As Long, ByVal lKey As Long)
'在RichEditor中 显示要素编辑器
Dim pt As POINTAPI, lHheight As Long, lHwidth As Long
    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '固定行高度
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '固定列宽度
    
    Doc.Range(lSE, lES).Selected
    ClientToScreen Doc.hWnd, pt
    Dim lLeft As Long, lTOp As Long
    '获取起始位置坐标
    Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp
    elEdit.SetElement Document.Elements("K" & lKey), 0, editType
    elEdit.Move F1Main.Left + Doc.Left + lLeft, F1Main.Top + Doc.Top + lTOp
    If elEdit.Top + elEdit.Height > F1Main.Top + F1Main.Height Then
        elEdit.Top = F1Main.Top + Doc.Top + lTOp - elEdit.Height - 300 - Screen.TwipsPerPixelY * 2
    End If
    If elEdit.Left + elEdit.Width > F1Main.Left + F1Main.Width Then
        elEdit.Left = F1Main.Left + Doc.Left + lLeft - elEdit.Width - Screen.TwipsPerPixelX * 2
    End If

    elEdit.Visible = True: elEdit.ZOrder 0: elEdit.SetFocus
End Sub

Private Sub Doc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
Dim i As Long
    
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Or KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        i = Doc.Selection.StartPos: If i = 0 Then i = 1
        bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded)  '处在关键字对之间
        If bInKeys Then
            Select Case KeyCode
                Case vbKeyDelete, vbKeyBack
                    If Doc.Range(i - 1, i + 3).Text Like ")?S(" And Doc.Range(i - 1, i + 3).Font.Hidden = True Then
                        '（隐藏文本）（保护文本）（隐藏文本）|（隐藏文本）（保护文本）（隐藏文本）
                        If KeyCode = vbKeyBack Then
                            i = i - 16
                        Else
                            i = i + 16
                        End If
                        bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
                    ElseIf Doc.Range(i - 17, i - 13).Text Like ")?S(" And Doc.Range(i - 17, i + 13).Font.Hidden = True Then
                            '（隐藏文本）（保护文本）（隐藏文本）（隐藏文本）|（保护文本）（隐藏文本）
                        If KeyCode = vbKeyBack Then
                            i = i - 32
                            bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
                        End If
                    ElseIf Doc.Range(i + 15, i + 19).Text Like ")?S(" And Doc.Range(i + 15, i + 19).Font.Hidden = True Then
                        '（隐藏文本）（保护文本）|（隐藏文本）（隐藏文本）（保护文本）（隐藏文本）
                        If KeyCode = vbKeyDelete Then
                            i = i + 32
                            bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '在前面个要素末位按Del,应删后面个要素
                        End If
                    ElseIf Doc.Range(i - 17, i - 16).Font.Hidden = False And Doc.Range(i - 17, i - 16).Font.Protected = False Then
                        bInKeys = False
                    End If
                    If bInKeys Then
                        KeyCode = 0
                        If editType <> TabET_单病历审核 Or Document.Elements("K" & lKey).ID = 0 Then
                            Document.Elements("K" & lKey).DeleteFromEditor Doc: Exit Sub  '删除要素
                        End If
                    End If
                Case vbKeySpace, vbKeyReturn
                    KeyCode = 0
                    Call ShowElInDoc(lSE, lES, lKey): Exit Sub
            End Select
        Else
            With Doc
                If .Range(i - 1, i).Font.Hidden And KeyCode = vbKeyBack Then  '在关键字后按Back
                    If i <= 1 Then
                        i = i + 15
                    Else
                        i = i - 16
                    End If
                    bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '处在关键字对之前
                    If bInKeys Then
                        KeyCode = 0
                        If editType <> TabET_单病历审核 Or Document.Elements("K" & lKey).ID = 0 Then
                            Document.Elements("K" & lKey).DeleteFromEditor Doc: Exit Sub '删除要素
                        End If
                    End If
                ElseIf .Range(i, i + 1).Font.Hidden And KeyCode = vbKeyDelete Then '在关键字前按DEL
                    i = i + 16: KeyCode = 0
                    bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '处在关键字对之前
                    If editType <> TabET_单病历审核 Or Document.Elements("K" & lKey).ID = 0 Then Document.Elements("K" & lKey).DeleteFromEditor Doc: Exit Sub '删除要素
                End If
            End With
        End If
    End If
End Sub

Private Sub Doc_KeyPress(KeyAscii As Integer)
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '处在关键字对之间
If bInKeys Then KeyAscii = 0
End Sub


Private Sub Doc_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyZ Then
        Doc.Undo
    ElseIf Shift = 2 And KeyCode = vbKeyY Then
        Doc.Redo
    End If
End Sub

Private Sub Doc_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bFinded As Boolean, bNeeded As Boolean

    bFinded = IsBetweenKeys(Doc, Doc.Selection.StartPos + 1, "E", lSS, lSE, lES, lEE, lKey, bNeeded)
    If bFinded Then '点击的是元素内部，则表示选择某个选项
        If Document.Elements("K" & lKey).输入形态 = 1 Then '展开形式的要素录入     '○●□■
            Dim strTmp As String, p As Long, P1 As Long, P2 As Long, blnForce As Boolean, lMax As Long
            With Doc
                .Freeze
                .ForceEdit = True
                strTmp = .Range(lSE, lES).Text
                p = .Selection.StartPos
                If Document.Elements("K" & lKey).要素表示 = 2 Then
                    P1 = .Selection.StartPos - lSE + 1
                    P1 = InStrRev(strTmp, "○", P1)
                    P2 = .Selection.StartPos - lSE + 1
                    P2 = InStrRev(strTmp, "●", P2)
                    If P1 > P2 And P1 > 0 Then
                        '单选
                        strTmp = Replace(strTmp, "●", "○")
                        Mid(strTmp, P1, 1) = "●"
                        .Range(lSE, lES).Text = strTmp
                    ElseIf P2 > P1 And P2 > 0 Then
                        strTmp = Replace(strTmp, "●", "○")
                        Mid(strTmp, P2, 1) = "○"
                        .Range(lSE, lES).Text = strTmp
                    End If
                    Document.Elements("K" & lKey).内容文本 = strTmp
                ElseIf Document.Elements("K" & lKey).要素表示 = 3 Then
                    P1 = .Selection.StartPos - lSE + 1
                    P1 = InStrRev(strTmp, "□", P1)
                    P2 = .Selection.StartPos - lSE + 1
                    P2 = InStrRev(strTmp, "■", P2)
                    If P1 > P2 And P1 > 0 Then
                        Mid(strTmp, P1, 1) = "■"
                        .Range(lSE, lES).Text = strTmp
                    ElseIf P2 > P1 And P2 > 0 Then
                        Mid(strTmp, P2, 1) = "□"
                        .Range(lSE, lES).Text = strTmp
                    End If
                    Document.Elements("K" & lKey).内容文本 = strTmp
                End If
                Me.Document.Elements("K" & lKey).内容文本 = strTmp
                .Range(p, p).Selected
                .UnFreeze
            End With
        Else
            
        End If
    End If
End Sub

Private Sub Doc_RequestRightMenu(ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim objPopup As CommandBar
Dim objControl As CommandBarControl

    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "剪切")
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPY, "复制")
        Set objControl = .Add(xtpControlButton, ID_EDIT_PASTE, "粘贴")
    End With
    objPopup.ShowPopup
End Sub

Private Sub elEdit_LostFocus()
    elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0: elEdit.Tag = ""
End Sub

Private Sub elEdit_pCancel()
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    If Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTElement Then '单要素取消编辑
        If F1Main.Visible And F1Main.Enabled Then
            F1Main.SetFocus
        End If
    ElseIf Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTText Then '当前为文本，插入日期时间，特殊符号
    Else
    
    End If
End Sub

Private Sub elEdit_pChange()
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    On Error GoTo errHand
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    If Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTElement Then '单要素变更编辑
        AddUndo Document.Cells.Cell(lsRow, lsCol)
        Document.Cells.Cell(lsRow, lsCol).内容文本 = elEdit.Element.内容文本
        F1Main.TextRC(lsRow, lsCol) = IIf(elEdit.Element.内容文本 <> "", elEdit.Element.内容文本, "[" & elEdit.Element.要素名称 & "]") & elEdit.Element.要素单位
        If elEdit.Visible Then
            elEdit.SetFocus
        End If
    ElseIf Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTText Then '当前为文本，插入日期时间，特殊符号
        '变更时不处理
    Else
    
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub elEdit_pOk()
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, strShow As String
    On Error GoTo errHand
    If UBound(Split(elEdit.Element.区域, "|")) > 0 Then
        lsRow = Split(elEdit.Element.区域, "|")(0): lsCol = Split(elEdit.Element.区域, "|")(1)
    Else
        Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    End If

    If Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTElement Then '单要素确定编辑
        AddUndo Document.Cells.Cell(lsRow, lsCol)
        With elEdit.Element
            If .替换域 = 1 Then
                If Trim(.内容文本) = "" Then
                    If .自动转文本 Then
                        strShow = " " & elEdit.Element.要素单位
                    Else
                        strShow = "[" & elEdit.Element.要素名称 & "]" & elEdit.Element.要素单位
                    End If
                Else
                    strShow = .内容文本
                End If
            Else
                strShow = IIf(elEdit.Element.内容文本 <> "", elEdit.Element.内容文本, "[" & elEdit.Element.要素名称 & "]") & elEdit.Element.要素单位
            End If
        End With
        Document.Cells.Cell(lsRow, lsCol).内容文本 = Trim(elEdit.Element.内容文本)
        F1Main.TextRC(lsRow, lsCol) = strShow
        elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0
        If F1Main.Enabled Then F1Main.SetFocus
    ElseIf Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTText Then '当前为文本，插入日期时间，特殊符号
        AddUndo Document.Cells.Cell(lsRow, lsCol)
        strShow = Document.Cells.Cell(lsRow, lsCol).内容文本 & elEdit.Element.内容文本
        Document.Cells.Cell(lsRow, lsCol).内容文本 = strShow
        F1Main.TextRC(lsRow, lsCol) = strShow
        elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0
    ElseIf Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTTextElement Then
        With elEdit.Element
            If .替换域 = 1 Then
                If Trim(.内容文本) = "" Then
                    If .自动转文本 Then
                        strShow = " " & elEdit.Element.要素单位
                    Else
                        strShow = "[" & elEdit.Element.要素名称 & "]" & elEdit.Element.要素单位
                    End If
                Else
                    strShow = .内容文本
                End If
            Else
                strShow = IIf(elEdit.Element.内容文本 <> "", elEdit.Element.内容文本, "[" & elEdit.Element.要素名称 & "]") & elEdit.Element.要素单位
            End If
        End With
        
        Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
        bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '处在关键字对之间
        If bInKeys Then
            If InStr(Document.Cells.Cell(lsRow, lsCol).ElementKey, lKey) = 0 Then '在当前单元格的要素Key串中找不到该要素Key表明是工具条插入的时间，将以文本方式加入，否则为双击要素弹出
                Doc.Range(lEE, lEE).Selected
                Doc.Range(lEE, lEE).Font.Protected = False
                Doc.Range(lEE, lEE).Font.Hidden = False
                Doc.Range(lEE, lEE).Text = strShow
                Doc.Range(lEE + Len(strShow), lEE + Len(strShow)).Selected
            Else
                Doc.Range(lSE, lES).Text = strShow
            End If
        Else
            Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Selected
            Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Font.Protected = False
            Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Font.Hidden = False
            Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Text = strShow
            Doc.Range(Doc.Selection.StartPos + Len(strShow), Doc.Selection.StartPos + Len(strShow)).Selected
        End If
        elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub elEdit_TitleMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage elEdit.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub F1Main_DblClick(ByVal nRow As Long, ByVal nCol As Long)
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    If nRow = 0 Or nCol = 0 Then Exit Sub
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Call EnterEdit(lsRow, lsCol, leRow, leCol, 0, True)
End Sub

Private Sub F1Main_EndEdit(EditString As String, Cancel As Integer)
    mblnEditing = False
    EditString = ToVarchar(EditString, 4000)
    '功能:跟据输入内容变更类存储
    If SelCell.Key = "" Then Exit Sub
    Call AddUndo(SelCell)
    With SelCell
        If .对象类型 = cprCTFixtext Or .对象类型 = cprCTText Then
        .内容文本 = EditString
            If InStr(.对象属性, ",") > 0 And InStr(.对象属性, ";") = 0 Then '合计单元格的源单元格
                Dim lsumRow As Long, lsumCol As Long
                lsumRow = Split(.对象属性, ",")(0): lsumCol = Split(.对象属性, ",")(1) '合计单元格的行列
                Call CalcSumRange(lsumRow, lsumCol)
            End If
        End If
    End With
End Sub

Private Sub F1Main_GotFocus()

    If PicEdit.Visible Then PicEdit.Visible = False: PicEdit.Top = 0: PicEdit.Left = 0: PicEdit.Tag = ""
    If mblnEditing Then F1Main.EndEdit
    If elEdit.Visible Then elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0: elEdit.Tag = ""
    If Doc.Visible Then GetFromDoc Doc.Tag, True: Doc.Text = "": Doc.ForceEdit = False: Doc.Visible = False: Doc.Top = 0: Doc.Left = 0: Doc.Title = "": Doc.Tag = ""
End Sub

Private Sub F1Main_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyDelete Then Exit Sub
    
    On Error Resume Next
    '删除图片或文本
    If F1Main.SelectionCount > 1 Then Exit Sub
    Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    
    With Document.Cells.Cell(lsRow, lsCol)
        If Not (.对象类型 = cprCTFixtext Or .对象类型 = cprCTText Or .对象类型 = cprCTElement Or .对象类型 = cprCTPicture Or .对象类型 = cprCTReportPic) Then Exit Sub '非(固定，普通文本，图片)不可以表格上直接删除内容
        If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
            If Not AllowEdit(lsRow, lsCol) Then Exit Sub          '不允许编辑直接退出
        End If
        
        AddUndo Document.Cells(.Key)
        Select Case .对象类型
            Case cprCTFixtext, cprCTText
                F1Main.TextRC(lsRow, lsCol) = ""
                .内容文本 = ""
                If InStr(.对象属性, ",") > 0 And InStr(.对象属性, ";") = 0 Then '合计单元格的源单元格
                    Dim lsumRow As Long, lsumCol As Long
                    lsumRow = Split(.对象属性, ",")(0): lsumCol = Split(.对象属性, ",")(1) '合计单元格的行列
                    Call CalcSumRange(lsumRow, lsumCol)
                End If
            Case cprCTElement
                If Document.Elements("K" & .ElementKey).输入形态 = 0 And Document.Elements("K" & .ElementKey).要素类型 = 2 Then
                    .内容文本 = ""
                    Document.Elements("K" & .ElementKey).内容文本 = ""
                    F1Main.TextRC(lsRow, lsCol) = "[" & Document.Elements("K" & .ElementKey).要素名称 & "]" & Document.Elements("K" & .ElementKey).要素单位
                End If
            Case cprCTPicture, cprCTReportPic
                If editType = TabET_单病历审核 Then Exit Sub
                If .PictureKey <> "" Then
                    Document.Pictures("K" & .PictureKey).OrigPic = New StdPicture
                    .PicMarkKey = ""
                    If ChkControl(PicDy(.Index)) Then Unload PicDy(.Index)
                End If
        End Select
    End With
    Err.Clear

End Sub

Private Sub F1Main_KeyPress(KeyAscii As Integer)
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, blnEndEdit As Boolean
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii = 3 Then KeyAscii = 0: Exit Sub 'Ctrl+C
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub 'Ctrl+v
    If KeyAscii = 24 Then KeyAscii = 0: Exit Sub 'Ctrl+X
    
    If F1Main.SelectionCount > 1 Then Exit Sub  '不允许批量赋值
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Call EnterEdit(lsRow, lsCol, leRow, leCol, KeyAscii)
End Sub

Private Sub F1Main_LostFocus()
    F1Main.EndEdit
End Sub

Private Sub F1Main_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngRow As Long, lngCol As Long
    F1Main.TwipsToRC x, y, lngRow, lngCol
    If lngRow = 0 Or lngCol = 0 Then
        mblnClickZ = True
        mblnChangeRC = True
    Else
        mblnClickZ = False
    End If
'    If lngRow > 0 And lngCol > 0 And Shift = 0 Then
'        Call F1Main.SetSelection(lngRow, lngCol, lngRow, lngCol)
'    End If
End Sub

Private Sub F1Main_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'缺省鼠标太难看，但在行/列头上需要显示Resize图标，动态改变
Dim lngRow As Long, lngCol As Long, vRect As F1Rect
    On Error Resume Next
    F1Main.TwipsToRC x, y, lngRow, lngCol
    If lngRow = 0 Then
        If lngCol = 0 Then F1Main.MousePointer = F1Arrow: Err.Clear: Exit Sub
        Set vRect = F1Main.RangeToTwipsEx(1, lngCol, 1, lngCol)
        If x < vRect.Left + 20 Or x > vRect.Right - 20 Then
            F1Main.MousePointer = F1SizeWE
        Else
            F1Main.MousePointer = F1Arrow
        End If
    ElseIf lngCol = 0 Then
        If lngRow = 0 Then F1Main.MousePointer = F1Arrow: Err.Clear: Exit Sub
        Set vRect = F1Main.RangeToTwipsEx(lngRow, 1, lngRow, 1)
        If y < vRect.Top + 20 Or y > vRect.Bottom - 20 Then
            F1Main.MousePointer = F1SizeNS
        Else
            F1Main.MousePointer = F1Arrow
        End If
    Else
        F1Main.MousePointer = F1Arrow
    End If
    Err.Clear
End Sub

Private Sub F1Main_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnChangeRC Then timeTmp.Enabled = True
    If mblnInit Then Exit Sub
Dim lngRow As Long, lngCol As Long
    F1Main.TwipsToRC x, y, lngRow, lngCol

    If lngRow = 0 Or lngCol = 0 Then Exit Sub
    With Document.Cells.Cell(lngRow, lngCol)
        '非文本类在定义、范文时，单击是选中，在编辑时单击进入编辑状态
        If (Not (.对象类型 = cprCTFixtext Or .对象类型 = cprCTText)) And (editType = TabET_单病历编辑 Or editType = TabET_单病历审核) Then
            Call F1Main_DblClick(lngRow, lngCol)
        End If
    End With
End Sub

Private Sub F1Main_SelChange()
Dim vCell As F1CellFormat, lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long
Dim lngRow As Long, lngCol As Long, lngWidth As Long, lngHeight As Long
    On Error GoTo errHand
    If mblnEditing Then mblnEditing = False
    If Not mblnInit Then
        Call F1Main.GetSelection(F1Main.SelectionCount - 1, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        If lngEndCol > F1Main.MaxCol Then lngEndCol = F1Main.MaxCol: F1Main.SetSelection lngStarRow, lngStarCol, lngEndRow, lngEndCol '通过选择第0行选择本行所有列
        If lngEndRow > F1Main.MaxRow Then lngEndRow = F1Main.MaxRow: F1Main.SetSelection lngStarRow, lngStarCol, lngEndRow, lngEndCol '通过选择第0列选择本列所有行
        
        Set vCell = F1Main.GetCellFormat
        Set SelCell = Document.Cells.Cell(lngStarRow, lngStarCol)
        Call SetColorIcon(ID_FORMAT_FORECOLOR, SelCell.FontColor)
        
        If F1Main.SelectionCount > 1 Then
            stbThis.Panels("msg").Text = "间隔选取多单元格"
            Call ShowAttr(-1, "")
        Else
            stbThis.Panels("msg").Text = lngStarRow & "行 " & lngStarCol & "列--" & lngEndRow & "行 " & lngEndCol & "列"
            '刷新左边属性
            If vCell.MergeCells Then '选中的是合并单元格
                ShowAttr Me.Document.Cells.Cell(lngStarRow, lngStarCol).对象类型, Me.Document.Cells.Cell(lngStarRow, lngStarCol).对象属性
                stbThis.Panels("msg").Text = stbThis.Panels("msg").Text & " 类型:" & Me.Document.Cells.Cell(lngStarRow, lngStarCol).CellTypeName
                For lngRow = lngStarRow To lngEndRow
                    lngHeight = lngHeight + F1Main.RowHeight(lngRow)
                Next
                For lngCol = lngStarCol To lngEndCol
                    lngWidth = lngWidth + F1Main.ColWidth(lngCol)
                Next
                If editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑 Then
                    stbThis.Panels("msg").Text = stbThis.Panels("msg").Text & " 高度(毫米):" & Round(Me.ScaleY(lngHeight, vbTwips, vbMillimeters), 2) & " 宽度(毫米):" & Round(Me.ScaleX(lngWidth, vbTwips, vbMillimeters), 2)
                End If
            Else
                If lngStarRow <> lngEndRow Or lngStarCol <> lngEndCol Then '选中的多个单元格
                    Call ShowAttr(-1, "")
                Else
                    ShowAttr Me.Document.Cells.Cell(lngStarRow, lngStarCol).对象类型, Me.Document.Cells.Cell(lngStarRow, lngStarCol).对象属性
                    stbThis.Panels("msg").Text = stbThis.Panels("msg").Text & " 类型:" & Me.Document.Cells.Cell(lngStarRow, lngStarCol).CellTypeName
                    lngHeight = lngHeight + F1Main.RowHeight(lngStarRow): lngWidth = lngWidth + F1Main.ColWidth(lngStarCol)
                    If editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑 Then
                        stbThis.Panels("msg").Text = stbThis.Panels("msg").Text & " 高度(毫米):" & Round(Me.ScaleY(lngHeight, vbTwips, vbMillimeters), 2) & " 宽度(毫米):" & Round(Me.ScaleX(lngWidth, vbTwips, vbMillimeters), 2)
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub F1Main_StartEdit(EditString As String, Cancel As Integer)
    mblnEditing = True
End Sub

Private Sub F1Main_TopLeftChanged()
Dim i As Integer, strCellKey As String
    On Error GoTo errHand
    If mblnInit Then Exit Sub
    If elEdit.Visible Then Exit Sub
    If Doc.Visible Then Exit Sub
    If PicEdit.Visible Then Exit Sub
    For i = 1 To PicDy.UBound
        If ChkControl(PicDy(i)) Then
            If PicDy(i).Picture.Handle <> 0 Then
                strCellKey = Split(PicDy(i).Tag, "|")(1)
                Call PaintPictureOnTable(strCellKey)
            End If
        End If
    Next
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
'功能：如果主窗体以模态方式打开，然后再将弹出其它模态窗体，主窗体中以DockPane方式加入的窗体会被莫名的设成不可用，以下代码保证子窗体处于可用状态
Dim lngStyle As Long
    On Error Resume Next
    If Not mfrmSentence Is Nothing Then
        If Not mfrmSentence.Enabled Then
            lngStyle = GetWindowLong(mfrmSentence.hWnd, GWL_STYLE)
            SetWindowLong mfrmSentence.hWnd, GWL_STYLE, lngStyle And Not WS_DISABLED
        End If
    End If
    If Not mfrmPacsPic Is Nothing Then
        If Not mfrmPacsPic.Enabled Then
            lngStyle = GetWindowLong(mfrmPacsPic.hWnd, GWL_STYLE)
            SetWindowLong mfrmPacsPic.hWnd, GWL_STYLE, lngStyle And Not WS_DISABLED
        End If
    End If
    Err.Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Select Case KeyCode
            Case vbKeyA
                ContentMove "All"
        End Select
    End If
End Sub

Private Sub Form_Load()
    Call InitForms
    Call MainCommandbarDefine
    Call MainDockPaneDefine
    Call mfrmSentence.zlSubRefClass(Document.EPRFileInfo.种类, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.医嘱id, Me)
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim IntReturn As Integer
    Err.Clear
    On Error Resume Next
1    If stbThis.Panels("msg").Text <> "保存完毕" And mReadOnly = 0 Then
2        '签名，保存退出会提前保存；菜单退出，关闭窗体，控制键退出不会，所以需要提示.
3        IntReturn = MsgBox("是否保存后再退出？" & vbCrLf & vbCrLf & "    [是]保存退出，[否]直接退出,[取消]不退出。", vbQuestion + vbYesNoCancel + vbDefaultButton1, gstrSysName)
4        If IntReturn = vbYes Then
5            Call SaveDoc(, True)
6        ElseIf IntReturn = vbCancel Then
7            Cancel = True
8            Exit Sub
9        End If
10    End If
    
11    If Document.EPRPatiRecInfo.医嘱id <> 0 Then Call Document.frmEditorClosed(Document.EPRPatiRecInfo.医嘱id)
12    Call SaveWinState(Me, App.ProductName)
13    If Not mfrmPacsPic Is Nothing Then Unload mfrmPacsPic
14    Set mfrmPacsPic = Nothing
15    If Not mfrmSentence Is Nothing Then Unload mfrmSentence
16    Set mfrmSentence = Nothing
17    If Not mfrmMainError Is Nothing Then Unload mfrmMainError
18    Set mfrmMainError = Nothing
19    If Not mfrmEPRModelSaveAs Is Nothing Then Unload mfrmEPRModelSaveAs
20    Set mfrmEPRModelSaveAs = Nothing
21    If Not mfrmTipInfo Is Nothing Then Unload mfrmTipInfo
22    Set mfrmTipInfo = Nothing
23    Set SelCell = Nothing
24    Set Document = Nothing
25    Set DocOld = Nothing
26    Set Undo = Nothing
27    Unload frmPublicIcon
28    Unload frmPicTypeset
      Err.Clear
      Exit Sub
End Sub

Private Sub mfrmEPRModelSaveAs_SaveModels(lngDemoId As Long, blnOK As Boolean)
Dim boldem As Byte, boldet As Byte
Dim arrSQL As Variant, i As Integer, blnBegin As Boolean

    On Error GoTo errHand
    arrSQL = Array()
    boldem = EditMode
    boldet = editType
    
    Document.EM = TabEm_修改
    Document.ET = TabET_全文示范编辑
    Document.EPRDemoInfo.GetDemoInfo lngDemoId
    Call Document.SaveDoc(arrSQL)
    
    blnBegin = True
    gcnOracle.BeginTrans '--------------------------写入数据e
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "写入数据")
    Next
    blnOK = True
    gcnOracle.CommitTrans
    
    Document.EM = boldem
    Document.ET = boldet
    blnBegin = False
    Exit Sub
errHand:
    If blnBegin Then gcnOracle.RollbackTrans
    Document.EM = boldem
    Document.ET = boldet
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmMainError_Location(ByVal strRange As String, ByVal elKey As Long)
Dim lRow As Long, lCol As Long
    On Error GoTo errHand
    If strRange = "" Or InStr(strRange, "|") = 0 Then
        Call F1Main.SetSelection(1, 1, 1, 1)
        F1Main.SetFocus
        Exit Sub
    End If
    
    lRow = Split(strRange, "|")(0): lCol = Split(strRange, "|")(1)
    If lRow <> 0 And lCol <> 0 Then '单要素直接定位
        F1Main.SetSelection lRow, lCol, lRow, lCol
        If F1Main.Visible And F1Main.Enabled Then F1Main.SetFocus
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmPacsPic_InsertPicture(pic As StdPicture)
Dim lngKey As Long, l As Long
    On Error GoTo errHand
    
    If SelCell.对象类型 <> cprCTReportPic Then Exit Sub '所选单元格为非报告图
    If Document.Pictures("K" & SelCell.PictureKey).OrigPic.Handle <> 0 Then
        '计算图片区域长宽
        Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, vR As F1Rect, lWidth As Long, lHeight As Long
        If SelCell.Merge Then
            lsRow = Split(Split(SelCell.MergeRange, ";")(0), ",")(0): leRow = Split(Split(SelCell.MergeRange, ";")(1), ",")(0)
            lsCol = Split(Split(SelCell.MergeRange, ";")(0), ",")(1): leCol = Split(Split(SelCell.MergeRange, ";")(1), ",")(1)
            Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
            lWidth = vR.Width: lHeight = vR.Height
        Else
            lWidth = SelCell.Width: lHeight = SelCell.Height
        End If
        
        frmPicTypeset.ShowTypeset Me, SelCell.Key, Document.EPRPatiRecInfo.医嘱id, lWidth, lHeight, _
            Document.Pictures("K" & SelCell.PictureKey).OrigPic, pic, Document.EPRFileInfo.lngModule
    Else
        '无图情况
        lngKey = Document.Pictures.Add
        Set Document.Pictures("K" & lngKey).OrigPic = pic  '加载图片
        Document.Cells(SelCell.Key).PictureKey = lngKey
        Call PaintPictureOnTable(SelCell.Key)    '重绘图片和标记
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmSentence_RowDblClick(ByVal lngWordId As Long)
Dim rsTemp As ADODB.Recordset, strText As String, lngStart As Long, lngLen As Long, lKey As Long
    
    On Error GoTo errHand
    If SelCell Is Nothing Then Exit Sub
    
    gstrSQL = "Select * From 病历词句组成 Where 词句id = [1] Order By 排列次序"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngWordId)
    If rsTemp.EOF Then Exit Sub
    
    Select Case SelCell.对象类型
        Case cprCTText                    '文本型，直接将要素转换成文本
            strText = ""
            With rsTemp
                Do While Not .EOF
                    Select Case !内容性质
                    Case 0 '自由文字
                        strText = strText & IIf(IsNull(!内容文本), " ", !内容文本)
                    Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                        strText = strText & IIf(IsNull(!内容文本), "{" & !要素名称 & "}" & !要素单位, "{" & !内容文本 & "}")
                    End Select
                    .MoveNext
                Loop
            End With
            If Trim(strText) = "" Then Exit Sub
            strText = SelCell.内容文本 & strText
            SelCell.内容文本 = strText
            F1Main.TextRC(SelCell.Row, SelCell.Col) = strText
        Case cprCTTextElement           '混合编辑型
            If Not Doc.Visible Then Exit Sub '当前处于非编辑状态,则不作处理
            Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, leKey As Long, bInKeys As Boolean, bNeeded As Boolean
            bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, leKey, bNeeded)
            If bInKeys Then Doc.Range(lEE, lEE).Selected
            
            Do Until rsTemp.EOF
                Select Case rsTemp!内容性质
                    Case 0              '自由文本
                        If NVL(rsTemp!内容文本) <> "" Then
                            lngStart = Doc.Selection.StartPos
                            strText = rsTemp!内容文本
                            lngLen = Len(strText)
                            Doc.Range(lngStart, lngStart).Text = strText
                            Doc.Range(lngStart, lngStart + lngLen).Font.Protected = False
                            Doc.Range(lngStart, lngStart + lngLen).Font.Hidden = False
                            Doc.Range(lngStart + lngLen, lngStart + lngLen).Selected
                        End If
                    Case 1, 2           '要素
                        lngStart = Doc.Selection.StartPos
                        lKey = Me.Document.Elements.Add
                        Me.Document.Elements("K" & lKey).ID = 0
                        Me.Document.Elements("K" & lKey).内容文本 = NVL(rsTemp!内容文本)
                        Me.Document.Elements("K" & lKey).要素名称 = NVL(rsTemp!要素名称)
                        Me.Document.Elements("K" & lKey).诊治要素ID = NVL(rsTemp!诊治要素ID, 0)
                        Me.Document.Elements("K" & lKey).替换域 = NVL(rsTemp!替换域, 0)
                        Me.Document.Elements("K" & lKey).要素类型 = NVL(rsTemp!要素类型, 0)
                        Me.Document.Elements("K" & lKey).要素长度 = NVL(rsTemp!要素长度, 0)
                        Me.Document.Elements("K" & lKey).要素小数 = NVL(rsTemp!要素小数, 0)
                        Me.Document.Elements("K" & lKey).要素单位 = NVL(rsTemp!要素单位)
                        Me.Document.Elements("K" & lKey).要素表示 = NVL(rsTemp!要素表示, 0)
                        Me.Document.Elements("K" & lKey).要素值域 = NVL(rsTemp!要素值域)
                        Me.Document.Elements("K" & lKey).输入形态 = NVL(rsTemp!输入形态, 0)
                        Me.Document.Elements("K" & lKey).对象属性 = NVL(rsTemp!对象属性, "||")
                        Me.Document.Elements("K" & lKey).区域 = SelCell.Row & "|" & SelCell.Col
                        If Me.Document.Elements("K" & lKey).替换域 = 1 And (Me.Document.ET = TabET_单病历编辑 Or Me.Document.ET = TabET_单病历审核) Then
                            Me.Document.Elements("K" & lKey).内容文本 = GetReplaceEleValue(Me.Document.Elements("K" & lKey).要素名称, _
                                Me.Document.EPRPatiRecInfo.病人ID, _
                                Me.Document.EPRPatiRecInfo.主页ID, _
                                Me.Document.EPRPatiRecInfo.病人来源, _
                                Me.Document.EPRPatiRecInfo.医嘱id, Me.Document.EPRPatiRecInfo.婴儿)
                        End If
                        Me.Document.Elements("K" & lKey).InsertIntoEditor Doc, editType, lngStart '将要素插入当前位置，光标定位到要素末端
                End Select
                rsTemp.MoveNext
            Loop
            If Doc.Enabled And Doc.Visible Then Doc.SetFocus
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub picAtt_Resize()
Dim i As Integer
    On Error Resume Next '&H00FDD6C6&
    fraType.Top = 80: fraType.Left = 120: fraType.BackColor = &HFDD6C6
    txtSum.Left = 80: txtSum.Top = fraType.Height + 80: txtSum.Width = picAtt.Width - 160
    shpTxtSum.Move txtSum.Left - Screen.TwipsPerPixelX, txtSum.Top - Screen.TwipsPerPixelY, txtSum.Width + Screen.TwipsPerPixelX * 2, txtSum.Height + Screen.TwipsPerPixelY * 2
    cmdApply.Move txtSum.Left + txtSum.Width - cmdApply.Width, txtSum.Top + txtSum.Height + 80
    cmdSum.Move txtSum.Left, txtSum.Top + txtSum.Height + 80
    cmdAvg.Move cmdSum.Left + cmdSum.Width - 20, txtSum.Top + txtSum.Height + 80
    For i = 0 To Text1.UBound
        Text1(i).BackColor = &HFDD6C6
        Text1(i).Move 0, cmdApply.Top + cmdApply.Height + 80
        Text1(i).Width = picAtt.Width
        Text1(i).Height = picAtt.Height - Text1(i).Top
    Next
    Err.Clear
End Sub

Private Sub PicDy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, strKey As String, strRange As String
    On Error GoTo errHand
    If mblnClickZ Then mblnClickZ = False
    If editType = TabET_单病历审核 Then Exit Sub '审核时不能编辑
    If Button = vbLeftButton Then
        If F1Main.Enabled And F1Main.Visible Then
            strRange = Split(PicDy(Index).Tag, "|")(0)
            If InStr(strRange, ";") > 0 Then
                lsRow = Split(Split(strRange, ";")(0), ",")(0): lsCol = Split(Split(strRange, ";")(0), ",")(1)
                leRow = Split(Split(strRange, ";")(1), ",")(0): leCol = Split(Split(strRange, ";")(1), ",")(1)
            Else
                lsRow = Split(strRange, ",")(0): lsCol = Split(strRange, ",")(1)
                leRow = Split(strRange, ",")(0): leCol = Split(strRange, ",")(1)
            End If
            Call F1Main.SetSelection(lsRow, lsCol, leRow, leCol)
            strKey = Split(PicDy(Index).Tag, "|")(1)
            EditPicture strKey
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PicEdit_LostFocus()
Dim strKey As String
    strKey = PicEdit.mselKey
    PicEdit.Visible = False: PicEdit.Top = 0: PicEdit.Left = 0: PicEdit.Tag = ""
    If strKey <> "" Then PaintPictureOnTable strKey
    If F1Main.Visible And F1Main.Enabled Then F1Main.SetFocus
End Sub

Private Sub picHistory_Resize()
    vsHistory.Width = picHistory.Width
    vsHistory.Height = picHistory.Height
    vsHistory.Top = 0: vsHistory.Left = 0
End Sub

Private Sub picMainBack_GotFocus()
    If F1Main.Visible And F1Main.Enabled Then
        F1Main.SetFocus
    End If
End Sub

Private Sub picMainBack_Resize()
On Error Resume Next
    picRulerH.Top = 0: picRulerH.Left = picRulerV.Width: picRulerH.Width = picMainBack.Width
    picRulerV.Top = picRulerH.Height: picRulerV.Left = 0: picRulerV.Height = picMainBack.Height
'    F1Main.Width = picMainBack.Width - picRulerV.Width: F1Main.Height = picMainBack.Height - picRulerH.Height
'    F1Main.Top = picRulerH.Height: F1Main.Left = picRulerV.Width
    picRulerH.Visible = False: picRulerV.Visible = False
    F1Main.Width = picMainBack.Width: F1Main.Height = picMainBack.Height
    F1Main.Top = 0: F1Main.Left = 0
    picMainBack.BackColor = RGB(255, 255, 255)
Err.Clear
End Sub

Private Sub SaveAsDemo()
    On Error GoTo errHand
    
    If Doc.Visible Or mblnEditing Then F1Main_GotFocus '如果处于编辑状态，需要将内容更新下来
    
    If mfrmEPRModelSaveAs Is Nothing Then
        Set mfrmEPRModelSaveAs = New frmEPRModelSaveAs
    End If
    
    If editType = TabET_全文示范编辑 Then
        mfrmEPRModelSaveAs.ShowMe 1, Document.EPRDemoInfo.ID
    Else
        mfrmEPRModelSaveAs.ShowMe 2, Document.EPRPatiRecInfo.ID
    End If
    
    Unload mfrmEPRModelSaveAs
    Set mfrmEPRModelSaveAs = Nothing
    
    mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提示：" & vbCrLf & "      成功保存成范文！", True, 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ImportDemo()
Dim lngModId As Long, lngFileID As Long, mfrmImportEPRDemo As New frmImportEPRDemo
    On Error GoTo errHand
    lngModId = frmImportEPRDemo.ShowMe(Me)
    If lngModId = 0 Then Exit Sub
    
    Call ClearPicture
    If editType = TabET_全文示范编辑 Then lngFileID = Document.EPRDemoInfo.ID '本身就是范文编辑先记下自身范文ID
    Document.EPRDemoInfo.GetDemoInfo lngModId                                 '跟据指定导入的范文ID读取相关信息
    Document.EM = TabEm_修改: Document.ET = TabET_全文示范编辑    '指定当前模式为范文模式
    If Not Document.ReadFileStructure Then Exit Sub                           '读取文件结构
    Document.ReadFileContent mblnMoved                                        '读取文件内容
    Document.EM = EditMode: Document.ET = editType                '恢复之前的模式
    If lngFileID <> 0 Then Document.EPRDemoInfo.GetDemoInfo lngFileID         '恢复范文编辑模式自身相关信息
    mblnInit = True
    RefreshF1Main                                                             '刷新界面
    mblnInit = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ExportXml()
Dim strFile As String
    On Error GoTo errHand
    strFile = GetSaveFile(Me.hWnd, Document.EPRFileInfo.名称 & ".xml", "XML文档" & Chr(0) & "*.xml" & Chr(0), "导出病历文件")
    
    If strFile = "" Then Exit Sub
    If gobjFSO.FileExists(strFile) Then
        If MsgBox("是否覆盖当前已存在的文件?", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call ValiCellDate
    If Not Document.BuildXmlFile(strFile, True) Then Exit Sub
    If gobjFSO.FileExists(strFile) Then
        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提示：" & vbCrLf & "      成功导出文件 <" & strFile & ">", True, 0
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ImportXml()
Dim strFile As String
'过滤类型
    On Error GoTo errHand
    strFile = GetOpenFile(Me.hWnd, "*.xml", "XML文件" & Chr(0) & "*.xml" & Chr(0), "导入病历文件")
    
    If strFile = "" Then Exit Sub
    If MsgBox("确实要使用导入文档覆盖正编辑的文档吗?", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call ClearPicture
    If Not Document.AnalyseFileStrucuture(strFile, True) Then Exit Sub
    mblnInit = True:  RefreshF1Main: mblnInit = False
    Document.EM = TabEm_新增
    mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提示：" & vbCrLf & "      成功从文件<" & strFile & ">导入数据", True, 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False
End Sub
Private Sub CompareChange(arrSQL As Variant)
'功能：比较变化的单元格，在某些情况下终止上一版，新生成一版
'调用：审核编辑时保存，再次修改保存，审核签名
Dim l As Long, strKey As String, lCount As Long, lastVar As Long, blnChange As Boolean, lngTmp As Long
    lCount = Document.Cells.Count: lastVar = Document.EPRPatiRecInfo.最后版本 + 1
    For l = 1 To lCount
        blnChange = False
        With Document.Cells(l)
        If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) And .ID <> 0 Then 'ID=0表示是追加的行记录
        Select Case .对象类型
            Case cprCTText
                If .内容文本 <> DocOld(.Key) Then '原记录与新记录内容不同
                    If .开始版 <> lastVar Then '签名后审核修改或签名后再签名(审核修改后再修改不做处理)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Endvar(" & .ID & "," & lastVar & ")"
                        .ID = 0
                        .开始版 = lastVar: .终止版 = 0
                        .对象序号 = Document.mMaxNo + 1: Document.mMaxNo = .对象序号
                    End If
                End If
            Case cprCTElement
                If .内容文本 <> DocOld(.Key) Then '原记录与新记录内容不同
                    If Document.Elements("K" & .ElementKey).开始版 <> lastVar Then   '签名后审核修改或签名后再签名(审核修改后再修改不做处理)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Endvar(" & .ID & "," & lastVar & ")"
                        .ID = 0
                        Document.Elements("K" & .ElementKey).开始版 = lastVar:  .开始版 = lastVar
                        Document.Elements("K" & .ElementKey).终止版 = 0:        .终止版 = 0
                        .对象序号 = Document.mMaxNo + 1: Document.mMaxNo = .对象序号
                    End If
                End If
            Case cprCTTextElement '混合型单元格 内容中的文本和要素随父对象一起终止,然后直接新增
                If .内容文本 <> DocOld(.Key) Then '原记录与新记录内容不同
                    If .开始版 <> lastVar Then '签名后审核修改或签名后再签名(审核修改后再修改不做处理)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Endvar(" & .ID & "," & lastVar & ")"
                        .ID = 0
                        .开始版 = lastVar: .终止版 = 0
                        .对象序号 = Document.mMaxNo + 1: Document.mMaxNo = .对象序号
                        
                        '对文本及要素开始版及终止版进行处理
                        For lngTmp = 1 To UBound(Split(.TextKey, "|")) '对本单元格中所有文本进行保存
                            With Document.Texts("K" & Split(.TextKey, "|")(lngTmp))
                                .ID = 0: .开始版 = lastVar: .终止版 = 0
                            End With
                        Next
                        For lngTmp = 1 To UBound(Split(.ElementKey, "|"))
                            With Document.Elements("K" & Split(.ElementKey, "|")(lngTmp))
                                .ID = 0: .开始版 = lastVar: .终止版 = 0
                            End With
                        Next
                    End If
                End If
                
            Case cprCTPicture, cprCTReportPic '图片不做比较
        End Select
        End If
        End With
    Next
End Sub
Private Sub PageSetUp()
Dim mfrmPageSetup As New frmPageSetup
    On Error GoTo errHand
    
    If mfrmPageSetup.ShowMe(Me, Document) = False Then Exit Sub
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub PrintDoc(ByVal blnPreview As Boolean)
'是否是预揽blnPreview
    On Error GoTo errHand
    If Doc.Visible Or mblnEditing Then F1Main_GotFocus '如果处于编辑状态，需要将内容更新下来
    If PicEdit.Visible Then Call PicEdit_LostFocus '如果图片处于编辑状态，需要先重绘出图片
    
    Call Document.PrintDoc(Me, blnPreview)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub SetCellFormat(ByVal strFormat As String, ByVal vData As Variant)
'功能:设置单元格格式
'参数:strFormat 格式名,vData 参数值
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long
Dim vCell As F1CellFormat, l As Long
    On Error GoTo errHand
    mblnInit = True
    '先处理类存储因为界面操作需要用类属性还原
    For l = 0 To F1Main.SelectionCount - 1 '间隔选择
        Call F1Main.GetSelection(l, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        For lngRow = lngStarRow To lngEndRow
            For lngCol = lngStarCol To lngEndCol
                With Document.Cells.Cell(lngRow, lngCol)
                    Select Case strFormat
                        Case "字体名称"
                            .FontName = vData
                        Case "字号"
                            .FontSize = GetFontSizeNumber(CStr(vData))
                        Case "粗体"
                            .FontBold = vData
                        Case "斜体"
                            .FontItalic = vData
                        Case "下划线"
                            .FontUnderline = vData
                        Case "删除线"
                            .FontStrikeout = vData
                        Case "保护"
                            .保留对象 = vData
                        Case "合并"
                            .Merge = vData
                            If lngStarRow = lngEndRow And lngStarCol = lngEndCol Then
                                .Merge = False
                            Else
                                Call ClearChildMember(.Key)
                                If .Merge Then '合并
                                    If lngRow = lngStarRow And lngCol = lngStarCol Then
                                        .MergeRange = lngRow & "," & lngCol & ";" & lngEndRow & "," & lngEndCol
                                    Else
                                        .MergeRange = lngRow & "," & lngCol
                                    End If
                                Else            '取消合并
                                    .MergeRange = lngRow & "," & lngCol
                                    .CellLineTop = F1BorderThin: .CellLineBottom = F1BorderThin
                                    .CellLineLeft = F1BorderThin: .CellLineRight = F1BorderThin
                                    .对象序号 = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo + 1
                                End If
                            End If
                        Case "字体颜色"
                            .FontColor = vData
                        Case "靠上左对齐"
                            .VAlignment = VALignTop: .HAlignment = HALignLeft
                        Case "靠上居中"
                            .VAlignment = VALignTop: .HAlignment = HAlignCenter
                        Case "靠上右对齐"
                            .VAlignment = VALignTop: .HAlignment = HALignRight
                        Case "中部左对齐"
                            .VAlignment = VAlignCenter: .HAlignment = HALignLeft
                        Case "中部居中"
                            .VAlignment = VAlignCenter: .HAlignment = HAlignCenter
                        Case "中部右对齐"
                            .VAlignment = VAlignCenter: .HAlignment = HALignRight
                        Case "靠下左对齐"
                            .VAlignment = VALignBottom: .HAlignment = HALignLeft
                        Case "靠下居中"
                            .VAlignment = VALignBottom: .HAlignment = HAlignCenter
                        Case "靠下右对齐"
                            .VAlignment = VALignBottom: .HAlignment = HALignRight
                    End Select
                End With
            Next
        Next
    Next

    If strFormat = "合并" Then                  '合并单独处理
'        For l = 0 To F1Main.SelectionCount - 1 '间隔选择
'            Call F1Main.GetSelection(l, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
'            If lngStarRow = lngEndRow And lngStarCol = lngEndCol Then
'                Set vCell = F1Main.GetCellFormat  '
'                vCell.MergeCells = False
'            Else
'                Set vCell = F1Main.GetCellFormat  '
'                vCell.MergeCells = vData
'                vCell.BorderStyle(F1TopBorder) = F1BorderThin: vCell.BorderStyle(F1BottomBorder) = F1BorderThin
'                vCell.BorderStyle(F1LeftBorder) = F1BorderThin: vCell.BorderStyle(F1RightBorder) = F1BorderThin
'                For lngRow = lngStarRow To lngEndRow
'                    For lngCol = lngStarCol To lngEndCol
'                        F1Main.TextRC(lngRow, lngCol) = ""
'                    Next
'                Next
'            End If
'        Next
        Call RefreshF1Main
        F1Main.SetSelection lngStarRow, lngStarCol, lngStarRow, lngStarCol
    Else
        Dim SelR() As SelRange
        ReDim SelR(F1Main.SelectionCount - 1) As SelRange
        For l = 0 To F1Main.SelectionCount - 1 '间隔选择        '将选择区域起止行列复制备用,因为接下来会改变Selection
            Call F1Main.GetSelection(l, SelR(l).lsRow, SelR(l).lsCol, SelR(l).leRow, SelR(l).leCol) '第N次间隔选择的起始行列
        Next
        
        For l = 0 To UBound(SelR)
            For lngRow = SelR(l).lsRow To SelR(l).leRow
                For lngCol = SelR(l).lsCol To SelR(l).leCol
                With Document.Cells.Cell(lngRow, lngCol)
                    If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then '只有合并单元格首个和非合并单元格有效
                        Call F1Main.SetSelection(lngRow, lngCol, lngRow, lngCol) '选择该区域
                        Set vCell = F1Main.GetCellFormat  '处理显示
                        Select Case strFormat
                            Case "字体名称"
                                vCell.FontName = vData
                            Case "字号"
                                vCell.FontSize = GetFontSizeNumber(CStr(vData))
                            Case "粗体"
                                vCell.FontBold = vData
                            Case "斜体"
                                vCell.FontItalic = vData
                            Case "下划线"
                                vCell.FontUnderline = vData
                            Case "删除线"
                                vCell.FontStrikeout = vData
                            Case "保护"
                                vCell.ProtectionLocked = vData
                            Case "字体颜色"
                                vCell.FontColor = vData
                            Case "靠上左对齐"
                                vCell.AlignVertical = F1VAlignTop: vCell.AlignHorizontal = F1HAlignLeft
                            Case "靠上居中"
                                vCell.AlignVertical = F1VAlignTop: vCell.AlignHorizontal = F1HAlignCenter
                            Case "靠上右对齐"
                                vCell.AlignVertical = F1VAlignTop: vCell.AlignHorizontal = F1HAlignRight
                            Case "中部左对齐"
                                vCell.AlignVertical = F1VAlignCenter: vCell.AlignHorizontal = F1HAlignLeft
                            Case "中部居中"
                                vCell.AlignVertical = F1VAlignCenter: vCell.AlignHorizontal = F1HAlignCenter
                            Case "中部右对齐"
                                vCell.AlignVertical = F1VAlignCenter: vCell.AlignHorizontal = F1HAlignRight
                            Case "靠下左对齐"
                                vCell.AlignVertical = F1VAlignBottom: vCell.AlignHorizontal = F1HAlignLeft
                            Case "靠下居中"
                                vCell.AlignVertical = F1VAlignBottom: vCell.AlignHorizontal = F1HAlignCenter
                            Case "靠下右对齐"
                                vCell.AlignVertical = F1VAlignBottom: vCell.AlignHorizontal = F1HAlignRight
                        End Select
                        F1Main.SetCellFormat vCell
                    End If
                End With
                Next
            Next
        Next
        For l = 0 To UBound(SelR)
            Call F1Main.AddSelection(SelR(l).lsRow, SelR(l).lsCol, SelR(l).leRow, SelR(l).leCol)
        Next
    End If
    
    mblnInit = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub SetCellBorder()
'功能:设置单元格格式
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long
Dim lOutBorder As Integer, lTopBorder As Integer, lBottomBorder As Integer, lLeftBorder As Integer, lRightBorder As Integer, lShade As Integer, lOutColor As Long, lTopColor As Long, lBottomColor As Long, lLeftColor As Long, lRightColor As Long
Dim vCell As F1CellFormat, l As Long, mfrmBorder As New frmBorder
    On Error GoTo errHand
    mblnInit = True '不触SelChange
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
    With Document.Cells.Cell(lngStarRow, lngStarCol)
        lLeftBorder = .CellLineLeft: lRightBorder = .CellLineRight: lTopBorder = .CellLineTop: lBottomBorder = .CellLineBottom
        lLeftColor = .CellLineLeftColor: lRightColor = .CellLineRightColor: lTopColor = .CellLineTopColor: lBottomColor = .CellLineBottomColor
    End With
    lOutBorder = -1: lOutColor = -1
    
    If Not mfrmBorder.ShowMe(lOutBorder, lLeftBorder, lRightBorder, lTopBorder, lBottomBorder, lShade, lOutColor, lLeftColor, lRightColor, lTopColor, lBottomColor, Me) Then mblnInit = False: Exit Sub
    
'   界面处理比较繁琐，因合并单元格等原因多变性，无法处理，直接处理类存储后，刷新

    '处理类存储当前单格
    On Error Resume Next
    For l = 0 To F1Main.SelectionCount - 1 '间隔选择
        Call F1Main.GetSelection(l, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        For lngRow = lngStarRow To lngEndRow  '起始行-终止行
            For lngCol = lngStarCol To lngEndCol '起始列-终止列
                With Document.Cells.Cell(lngRow, lngCol)
                    If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then '有效单元格
                        .CellLineTop = lTopBorder: .CellLineBottom = lBottomBorder: .CellLineLeft = lLeftBorder: .CellLineRight = lRightBorder
                        .CellLineTopColor = lTopColor: .CellLineBottomColor = lBottomColor: .CellLineLeftColor = lLeftColor: .CellLineRightColor = lRightColor
                        If lngRow > 1 Then
                        With Document.Cells.Cell(lngRow - 1, lngCol)    '被设定单元格的上方单元格下边线
                            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                .CellLineBottom = lTopBorder
                                .CellLineBottomColor = lTopColor
                            End If
                        End With
                        End If
                        
                        If lngRow < F1Main.MaxRow Then
                        With Document.Cells.Cell(lngRow + 1, lngCol)    '被设定单元格下方单元格的上边线
                            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                .CellLineTop = lBottomBorder
                                .CellLineTopColor = lBottomColor
                            End If
                        End With
                        End If
                        
                        If lngCol > 1 Then
                        With Document.Cells.Cell(lngRow, lngCol - 1)    '被设定单元格左方单元格的右边线
                            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                .CellLineRight = lLeftBorder
                                .CellLineRightColor = lLeftColor
                            End If
                        End With
                        End If
                        
                        If lngCol < F1Main.MaxCol Then
                        With Document.Cells.Cell(lngRow, lngCol + 1)    '被设定单元格右方单元格的左边线
                            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                .CellLineLeft = lRightBorder
                                .CellLineLeftColor = lRightColor
                            End If
                        End With
                        End If
                    Else                                                   '合并单元格的非首个单元格为无效单元格
                        If lngCol <> lngEndCol Then
                            .CellLineTop = 0: .CellLineBottom = 0: .CellLineLeft = 0: .CellLineRight = 0
                            .CellLineTopColor = 0: .CellLineBottomColor = 0: .CellLineLeftColor = 0: .CellLineRightColor = 0
                        Else    '合并单元格后面
                            If lngCol < F1Main.MaxCol Then
                            With Document.Cells.Cell(lngRow, lngCol + 1)    '被设定单元格右方单元格的左边线
                                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                    .CellLineLeft = lRightBorder
                                    .CellLineLeftColor = lRightColor
                                End If
                            End With
                            End If
                        End If
                    End If
                End With
            Next
        Next
    Next
    RefreshF1Main
    Err.Clear
    mblnInit = False
    F1Main.SetSelection lngStarRow, lngStarCol, lngStarRow, lngStarCol
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False
End Sub
Private Function SetSumAtt(ByRef vData As String) As Boolean
',当前单元格由哪些单元格合计得来
'vdata为Sum时格式为 行,列;行,列----,其它为空,
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long, strTmp As String, l As Long
    On Error GoTo errHand
    
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
    lngRow = lngStarRow: lngCol = lngStarCol
        
    If vData <> "" Then '=""表示取消合计属性
        For l = 0 To UBound(Split(vData, ";"))
            strTmp = Split(vData, ";")(l) '单元格行列=0或超过表格最大行列或=当前单元格行列视为无效
            If Split(strTmp, ",")(0) > F1Main.MaxRow Or Split(strTmp, ",")(1) > F1Main.MaxCol Then
                vData = "合计单元格的来源单元格 " & Split(strTmp, ",")(0) & "行 " & Split(strTmp, ",")(1) & "列 超出表格范围，请检查！": Exit Function
            End If
            If (Split(strTmp, ",")(0) = lngRow And Split(strTmp, ",")(1) = lngCol) Then
                vData = "合计单元格的来源单元格不能是自身 " & Split(strTmp, ",")(0) & "行 " & Split(strTmp, ",")(1) & "列，请检查！": Exit Function
            End If
                 
            If Not (Document.Cells.Cell(Split(strTmp, ",")(0), Split(strTmp, ",")(1)).对象类型 = cprCTFixtext Or Document.Cells.Cell(Split(strTmp, ",")(0), Split(strTmp, ",")(1)).对象类型 = cprCTText) Then
                vData = "合计单元格的来源单元格 " & Split(strTmp, ",")(0) & "行 " & Split(strTmp, ",")(1) & "列 不是固定文本/多行文本单元格，请检查！": Exit Function
            End If
        Next
    End If
    
    
    With Document.Cells.Cell(lngRow, lngCol)
        strTmp = "" '先取消原有合计属性
        If .对象属性 <> "" And (.对象类型 = cprCTFixtext Or .对象类型 = cprCTText) Then '目标单元格
            For l = 0 To UBound(Split(.对象属性, ";"))
                strTmp = Split(.对象属性, ";")(l)
                If UBound(Split(strTmp, ",")) > 0 Then '确保对象属性正确性,只有固定文本和多行文本具有合计属性
                    With Document.Cells.Cell(Split(strTmp, ",")(0), Split(strTmp, ",")(1)) '源单元格
                        If .对象类型 = cprCTFixtext Or .对象类型 = cprCTText Then
                            .对象属性 = ""
                        End If
                    End With
                End If
            Next
            .对象属性 = ""
        End If
        
        If vData <> "" Then '设定合计属性
            For l = 0 To UBound(Split(vData, ";"))
                strTmp = Split(vData, ";")(l)
                Document.Cells.Cell(Split(strTmp, ",")(0), Split(strTmp, ",")(1)).对象属性 = lngRow & "," & lngCol
            Next
        End If
        .对象属性 = vData
    End With
    SetSumAtt = True: vData = ""
    CalcSumRange lngRow, lngCol
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetCellAttribute(ByVal strType As String, ByRef vData As String, blnReturn As Boolean)
'功能:设置单元格格式
'参数:strType 类型0,1,2,3,4,5,6,7,8 分别表示固定TXT,多行TXT,单要素,混合编辑,参考图,报告图,行控签名,列控签名,签名位
'    vData 传入参数,失败时返回消息;blnReturn设定是否成功
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long
Dim lmsRow As Long, lmsCol As Long, lmeRow As Long, lmeCol As Long, lCellCount As Long
Dim lngTmp As Long, vR As F1Rect, strTmp As String, lS As Long, l As Long, j As Long
    
    
    For lS = 0 To F1Main.SelectionCount - 1
        Call F1Main.GetSelection(lS, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        If strType = cprCTColSign Or strType = cprCTRowSign Or strType = cprCTSign Then '对行、列控、签名位的判断
            lCellCount = Document.Cells.Count
            For l = 1 To lCellCount
                With Document.Cells(l)
                    Select Case strType '行、列控、签名不能在同一文档中出现
                        Case cprCTSign
                            If .对象类型 = cprCTColSign Or .对象类型 = cprCTRowSign Then
                                vData = "行控、列控、普通签名不能在同一文档内出现！": Exit Sub
                            End If
                        Case cprCTColSign
                            If .对象类型 = cprCTSign Or .对象类型 = cprCTRowSign Then
                                vData = "行控、列控、普通签名不能在同一文档内出现！": Exit Sub
                            End If
                        Case cprCTRowSign
                            If .对象类型 = cprCTSign Or .对象类型 = cprCTColSign Then
                                vData = "行控、列控、普通签名不能在同一文档内出现！": Exit Sub
                            End If
                    End Select
                    
                    If .Merge And InStr(.MergeRange, ";") > 0 Then '对合并的区域进行判断
                        lmsRow = Val(Split(Split(.MergeRange, ";")(0), ",")(0)): lmsCol = Val(Split(Split(.MergeRange, ";")(0), ",")(1))
                        lmeRow = Val(Split(Split(.MergeRange, ";")(1), ",")(0)): lmeCol = Val(Split(Split(.MergeRange, ";")(1), ",")(1))
                        If strType = cprCTColSign Then
                            If lmsCol <> lmeCol And .Row <> 1 Then '有跨列合并的情况且不在第一行,因为第一行可能是标题
                            For j = lngStarCol To lngEndCol '选中的列中出现被合并情况
                                If j >= lmsCol And j <= lmeCol Then
                                    vData = "列控签名所在列不能有跨列合并单元格！": Exit Sub
                                End If
                            Next
                            End If
                        ElseIf strType = cprCTRowSign Then
                            If lmsRow <> lmeRow Then '有跨行合并的情况
                            For j = lngStarRow To lngEndRow '选中的行中出现被合并的情况
                                If j >= lmsRow And j <= lmeRow Then
                                    vData = "行控签名所在行不能有跨行合并单元格！": Exit Sub
                                End If
                            Next
                            End If
                        End If
                    End If
                End With
            Next
        End If
        
        For lngRow = lngStarRow To lngEndRow
            For lngCol = lngStarCol To lngEndCol
                With Document.Cells.Cell(lngRow, lngCol)
                    '先删除原有对象类型 在集合中的记录,可能间隔选择多个单元格,只处理 变更前与变更后类型发生变化的
                    If .对象类型 <> strType Then
                        Call ClearChildMember(.Key)
                        .对象类型 = strType: .对象属性 = "": .内容行次 = 0: .内容文本 = ""
                        Select Case .对象类型
                            Case cprCTFixtext          '0-固定文本(不可编辑)
                                .保留对象 = True
                                F1Main.TextRC(lngRow, lngCol) = .内容文本
                            Case cprCTText            '1-文本型(可编辑多行文本)
                                .保留对象 = False
                                F1Main.TextRC(lngRow, lngCol) = .内容文本
                            Case cprCTElement          '2-单要素
                                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                    .ElementKey = Me.Document.Elements.Add
                                    Me.Document.Elements("K" & .ElementKey).ID = .ID
                                    Me.Document.Elements("K" & .ElementKey).父ID = 0
                                    Me.Document.Elements("K" & .ElementKey).区域 = lngRow & "|" & lngCol
                                    InsertElement .Key '插入要素
                                End If
                                .保留对象 = False
                            Case cprCTTextElement       '3-文本与多要素混合编辑\
                                .保留对象 = False: .ElementKey = "": .TextKey = ""
                                F1Main.TextRC(lngRow, lngCol) = .内容文本
                            Case cprCTPicture, cprCTReportPic          '4-参考图
                                .保留对象 = False
                                .内容文本 = IIf(.对象类型 = cprCTPicture, "参考图", "报告图")
                                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                    Set vR = F1Main.RangeToTwipsEx(lngStarRow, lngStarCol, lngEndRow, lngEndCol)
                                    .PictureKey = Me.Document.Pictures.Add
                                    Me.Document.Pictures("K" & .PictureKey).PicID = .ID
                                    Me.Document.Pictures("K" & .PictureKey).DesWidth = vR.Width
                                    Me.Document.Pictures("K" & .PictureKey).DesHeight = vR.Height
                                End If
                                F1Main.TextRC(lngRow, lngCol) = .内容文本
                            Case cprCTSign, cprCTRowSign, cprCTColSign          '6-签名
                                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                    .SignKey = Me.Document.Signs.Add
                                End If
                                .内容文本 = "[签名位]"
                                .保留对象 = False
                                F1Main.TextRC(lngRow, lngCol) = .内容文本
                        End Select
                    End If
                End With
            Next
        Next
    Next
    blnReturn = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub SetRowCol(ByVal strType As String)
'功能:调整行列
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long
Dim l As Long, lngHeight As Long, lngWidth As Long, mfrmSetRowCol As New frmSetRowCol
    On Error GoTo errHand
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol) '取得最后一个选择区域起止行列
    Select Case strType
        Case "行高"
            lngHeight = F1Main.RowHeight(lngStarRow)
            Call mfrmSetRowCol.SetRowCol(Me, strType, lngHeight)
        Case "列宽"
            lngWidth = F1Main.ColWidthTwips(lngStarCol)
            Call mfrmSetRowCol.SetRowCol(Me, strType, lngWidth)
        Case "相同行高"
            lngHeight = F1Main.RowHeight(lngStarRow)
        Case "相同列宽"
            lngWidth = F1Main.ColWidthTwips(lngStarCol)
    End Select
    If lngHeight = -1 Or lngWidth = -1 Then Exit Sub
    
    For l = 0 To F1Main.SelectionCount - 1
        Call F1Main.GetSelection(l, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        If lngHeight <> 0 Then '调整行高
            For lngRow = lngStarRow To lngEndRow
                F1Main.RowHeight(lngRow) = lngHeight
            Next
        End If
        If lngWidth <> 0 Then '调整列宽
            For lngCol = lngStarCol To lngEndCol
                F1Main.ColWidthTwips(lngCol) = lngWidth
            Next
        End If
    Next
    timeTmp.Enabled = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function InsertRowCol(ByVal strType As String) As Boolean
'功能:插入行列
'说明：插入前判断合并单元格，插入后改变类存储，新增类存储，改变展现样式
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long
Dim lmsRow As Long, lmsCol As Long, lmeRow As Long, lmeCol As Long '合并区域的起止行列
Dim lCRow As Long, lCCol As Long, lNRow As Long, lNCol As Long
Dim vCell As F1CellFormat, l As Long, IntInsertType As Integer, j As Integer, strKey As String, lCellCount As Long
Dim lMaxR As Long, lMaxC As Long, lR As Long, lC As Long, TmpCell As cTabCell
    On Error GoTo errHand
    mblnInit = True
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol) '取得首个选择区域起止行列
    If lngEndRow > F1Main.MaxRow Then lngEndRow = F1Main.MaxRow
    If lngEndCol > F1Main.MaxCol Then lngEndCol = F1Main.MaxCol
    lCellCount = Document.Cells.Count: lMaxR = Document.Cells.Rows: lMaxC = Document.Cells.Cols
    '确定在哪行，哪列前插入
    Select Case strType
        Case "InsertLeftCol" '在左插入新列
            lNRow = lngStarRow: lNCol = lngStarCol: lCRow = lngStarRow: lCCol = lngStarCol: IntInsertType = F1ShiftCols
        Case "InsertRightCol" '在右插入新列
            lNRow = lngEndRow: lNCol = lngEndCol + 1: lCRow = lngEndRow: lCCol = lngEndCol: IntInsertType = F1ShiftCols
        Case "InsertUpRow"  '在上插入新行
            lNRow = lngStarRow: lNCol = lngStarCol: lCRow = lngStarRow: lCCol = lngStarCol: IntInsertType = F1ShiftRows
        Case "InsertDnRow"  '在下插入新行
            lNRow = lngEndRow + 1: lNCol = lngEndCol: lCRow = lngEndRow: lCCol = lngEndCol: IntInsertType = F1ShiftRows
    End Select
    
    If Not (lNRow > F1Main.MaxRow Or lNCol > F1Main.MaxCol) Then '在最后追加不需要判断
    For l = 1 To lCellCount
        With Document.Cells(l)
            If .Merge And InStr(.MergeRange, ";") > 0 Then '对合并的区域进行判断
                lmsRow = Val(Split(Split(.MergeRange, ";")(0), ",")(0)): lmsCol = Val(Split(Split(.MergeRange, ";")(0), ",")(1))
                lmeRow = Val(Split(Split(.MergeRange, ";")(1), ",")(0)): lmeCol = Val(Split(Split(.MergeRange, ";")(1), ",")(1))
                
                If IntInsertType = F1ShiftCols Then '插入列
                    If lmsCol <> lmeCol Then '有跨列合并的情况
                        If strType = "InsertLeftCol" Then '在左方插入列，本列不能出现有与前一列合并的情况
                            If lCCol > lmsCol And lCCol <= lmeCol Then
                                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      插入新列，当前单元格所在列不能有跨列合并单元格" & vbCrLf & "      请检查！", True, 1
'                                stbThis.Panels("msg").Text = "插入新列时，当前选中的单元格所在列不能有跨列合并单元格！"
                                mblnInit = False: Exit Function
                            End If
                        Else                              '在右方插入列，本列不能出现有与后一列合并的情况
                            If lCCol >= lmsCol And lCCol < lmeCol Then
                                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      插入新列，当前单元格所在列不能有跨列合并单元格" & vbCrLf & "      请检查！", True, 1
'                                stbThis.Panels("msg").Text = "插入新列时，当前选中的单元格所在列不能有跨列合并单元格！"
                                mblnInit = False: Exit Function
                            End If
                        End If
                    End If
                Else                    '插入行
                    If lmsRow <> lmeRow Then '有跨行合并的情况
                        If strType = "InsertUpRow" Then '在上方插入行，本行不能出现有与前一行合并的情况
                            If lCRow > lmsRow And lCRow <= lmeRow Then
                                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      插入新行，当前单元格所在行不能有跨行合并单元格" & vbCrLf & "      请检查！", True, 1
'                                stbThis.Panels("msg").Text = "插入新行时，当前选中的单元格所在行不能有跨行合并单元格！"
                                mblnInit = False: Exit Function
                            End If
                        Else                            '在下方插入行，本行不能出现有与后一行合并的情况
                            If lCRow >= lmsRow And lCRow < lmeRow Then
                                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      插入新行，当前单元格所在行不能有跨行合并单元格" & vbCrLf & "      请检查！", True, 1
                                'stbThis.Panels("msg").Text = "插入新行时，当前选中的单元格所在行不能有跨行合并单元格！"
                                mblnInit = False: Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next
    End If
    
    '插入行，列
    Call F1Main.InsertRange(lNRow, lNCol, lNRow, lNCol, IntInsertType)
    If IntInsertType = F1ShiftCols Then '增加列
        F1Main.MaxCol = F1Main.MaxCol + 1
        For l = 1 To F1Main.MaxCol
            F1Main.ColText(l) = l '列头显示数字
        Next
    Else                                '增加行
        F1Main.MaxRow = F1Main.MaxRow + 1
    End If
    
    '变更类存储
    Select Case strType
        Case "InsertLeftCol", "InsertRightCol" '插入新列
            For lC = lMaxC To 1 Step -1
                If lC >= lNCol Then '新列之后的列,先删除再新增
                    For lR = lMaxR To 1 Step -1
                        Set TmpCell = Document.Cells("K" & lR & "_" & lC)
                        With Document
                            .Cells.Remove ("K" & lR & "_" & lC) '先删除
                            .Cells.Add lR, lC + 1               '再新增以确保集合中关键字唯一
                            With .Cells("K" & lR & "_" & lC + 1)
                                .ID = TmpCell.ID
                                .文件ID = TmpCell.文件ID
                                .对象序号 = TmpCell.对象序号
                                .对象类型 = TmpCell.对象类型
                                .保留对象 = TmpCell.保留对象
                                .对象属性 = TmpCell.对象属性
                                .内容行次 = TmpCell.内容行次
                                .内容文本 = TmpCell.内容文本
                                .开始版 = TmpCell.开始版
                                .终止版 = TmpCell.终止版
                                .Row = TmpCell.Row
                                .Col = TmpCell.Col + 1
                                .Width = TmpCell.Width
                                .Height = TmpCell.Height
                                .FontName = TmpCell.FontName
                                .FontSize = TmpCell.FontSize
                                .FontBold = TmpCell.FontBold
                                .FontItalic = TmpCell.FontItalic
                                .FontUnderline = TmpCell.FontUnderline
                                .FontStrikeout = TmpCell.FontStrikeout
                                .FontColor = TmpCell.FontColor
                                .HAlignment = TmpCell.HAlignment
                                .VAlignment = TmpCell.VAlignment
                                .CellLineTop = TmpCell.CellLineTop
                                .CellLineBottom = TmpCell.CellLineBottom
                                .CellLineLeft = TmpCell.CellLineLeft
                                .CellLineRight = TmpCell.CellLineRight
                                .CellLineTopColor = TmpCell.CellLineTopColor
                                .CellLineBottomColor = TmpCell.CellLineBottomColor
                                .CellLineLeftColor = TmpCell.CellLineLeftColor
                                .CellLineRightColor = TmpCell.CellLineRightColor
                                .Merge = TmpCell.Merge
                                .MergeRange = TmpCell.MergeRange
                                .TextKey = TmpCell.TextKey
                                .ElementKey = TmpCell.ElementKey
                                .PictureKey = TmpCell.PictureKey
                                .SignKey = TmpCell.SignKey
                                .PicMarkKey = TmpCell.PicMarkKey
                                If .Merge And InStr(.MergeRange, ";") > 0 Then '合并单元格首个,行不变,列+1
                                    .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) & "," & Split(Split(.MergeRange, ";")(0), ",")(1) + 1 & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) & "," & Split(Split(.MergeRange, ";")(1), ",")(1) + 1
                                Else
                                    .MergeRange = .Row & "," & .Col
                                End If
                                For j = 0 To UBound(Split(.ElementKey, "|")) '改变元素中区域的保存
                                    If Len(Split(.ElementKey, "|")(j)) > 0 Then
                                        Document.Elements("K" & Split(.ElementKey, "|")(j)).区域 = .Row & "|" & .Col
                                    End If
                                Next
                                For j = 0 To UBound(Split(.TextKey, "|")) '改变文字中区域的保存
                                    If Len(Split(.TextKey, "|")(j)) > 0 Then
                                        Document.Texts("K" & Split(.TextKey, "|")(j)).区域 = .Row & "|" & .Col
                                    End If
                                Next
                                If .PictureKey <> "" Then
                                    If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                        If PicEdit.Visible Then PicEdit.Visible = False
                                        Unload PicDy(TmpCell.Index)
                                    End If
                                End If
                            End With
                        End With
                    Next
                End If
            Next
            '新增类单元格,行为所在行序，列为新增列序
            For l = 1 To F1Main.MaxRow
                strKey = Document.Cells.Add(l, lNCol)
                With Document.Cells(strKey)
                    .Height = F1Main.RowHeight(l)
                    .Width = F1Main.ColWidthTwips(Decode(lNCol, F1Main.MaxCol, lNCol - 1, F1Main.MinCol, lNCol + 1, lNCol - 1))
                    .对象序号 = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo + 1
                    If lNCol = F1Main.MaxCol Then '追加继承列
                        .Merge = Document.Cells.Cell(l, lNCol - 1).Merge
                        If InStr(Document.Cells.Cell(l, lNCol - 1).MergeRange, ";") > 0 Then '有效合并单元格
                            .MergeRange = Document.Cells.Cell(l, lNCol - 1).MergeRange '追加的列，只可能出现跨行合并的情况，此时新列继承上列的合并属性,列+1行不变
                            If Val(Split(Split(.MergeRange, ";")(0), ",")(1)) <> Val(Split(Split(.MergeRange, ";")(1), ",")(1)) Then '有跨列合并的情况不处理合并
                                .MergeRange = l & "," & lNCol
                                .Merge = False
                            Else
                                .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) & "," & Split(Split(.MergeRange, ";")(0), ",")(1) + 1 & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) & "," & Split(Split(.MergeRange, ";")(1), ",")(1) + 1
                                .Merge = True
                            End If
                        Else '非合并单元格和合并的无效单元格
                            .Merge = False
                            .MergeRange = l & "," & lNCol
                        End If
                    Else                         '普通增加列不对合并进行处理
                        .MergeRange = l & "," & lNCol
                    End If
                End With
            Next
            F1Main.ColWidthTwips(lNCol) = F1Main.ColWidthTwips(Decode(lNCol, F1Main.MaxCol, lNCol - 1, F1Main.MinCol, lNCol + 1, lNCol - 1))
            Document.Cells.Cols = Document.Cells.Cols + 1
        Case "InsertUpRow", "InsertDnRow" '插入新行
            For lR = lMaxR To 1 Step -1
                If lR >= lNRow Then '新行之后的行,先删除再新增
                    For lC = lMaxC To 1 Step -1
                        Set TmpCell = Document.Cells("K" & lR & "_" & lC)
                        With Document
                            .Cells.Remove ("K" & lR & "_" & lC) '先删除
                            .Cells.Add lR + 1, lC             '再新增以确保集合中关键字唯一
                            With .Cells("K" & lR + 1 & "_" & lC)
                                .ID = TmpCell.ID
                                .文件ID = TmpCell.文件ID
                                .对象序号 = TmpCell.对象序号
                                .对象类型 = TmpCell.对象类型
                                .保留对象 = TmpCell.保留对象
                                .对象属性 = TmpCell.对象属性
                                .内容行次 = TmpCell.内容行次
                                .内容文本 = TmpCell.内容文本
                                .开始版 = TmpCell.开始版
                                .终止版 = TmpCell.终止版
                                .Row = TmpCell.Row + 1
                                .Col = TmpCell.Col
                                .Width = TmpCell.Width
                                .Height = TmpCell.Height
                                .FontName = TmpCell.FontName
                                .FontSize = TmpCell.FontSize
                                .FontBold = TmpCell.FontBold
                                .FontItalic = TmpCell.FontItalic
                                .FontUnderline = TmpCell.FontUnderline
                                .FontStrikeout = TmpCell.FontStrikeout
                                .FontColor = TmpCell.FontColor
                                .HAlignment = TmpCell.HAlignment
                                .VAlignment = TmpCell.VAlignment
                                .CellLineTop = TmpCell.CellLineTop
                                .CellLineBottom = TmpCell.CellLineBottom
                                .CellLineLeft = TmpCell.CellLineLeft
                                .CellLineRight = TmpCell.CellLineRight
                                .CellLineTopColor = TmpCell.CellLineTopColor
                                .CellLineBottomColor = TmpCell.CellLineBottomColor
                                .CellLineLeftColor = TmpCell.CellLineLeftColor
                                .CellLineRightColor = TmpCell.CellLineRightColor
                                .Merge = TmpCell.Merge
                                .MergeRange = TmpCell.MergeRange
                                .TextKey = TmpCell.TextKey
                                .ElementKey = TmpCell.ElementKey
                                .PictureKey = TmpCell.PictureKey
                                .SignKey = TmpCell.SignKey
                                .PicMarkKey = TmpCell.PicMarkKey
                                If .Merge And InStr(.MergeRange, ";") > 0 Then '合并单元格首个,列不变,行+1
                                    .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) + 1 & "," & Split(Split(.MergeRange, ";")(0), ",")(1) & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) + 1 & "," & Split(Split(.MergeRange, ";")(1), ",")(1)
                                Else
                                    .MergeRange = .Row & "," & .Col
                                End If
                                For j = 0 To UBound(Split(.ElementKey, "|")) '改变元素中区域的保存
                                    If Len(Split(.ElementKey, "|")(j)) > 0 Then
                                        Document.Elements("K" & Split(.ElementKey, "|")(j)).区域 = .Row & "|" & .Col
                                    End If
                                Next
                                For j = 0 To UBound(Split(.TextKey, "|")) '改变文字中区域的保存
                                    If Len(Split(.TextKey, "|")(j)) > 0 Then
                                        Document.Texts("K" & Split(.TextKey, "|")(j)).区域 = .Row & "|" & .Col
                                    End If
                                Next
                                If .PictureKey <> "" Then
                                    If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                        If PicEdit.Visible Then PicEdit.Visible = False
                                        Unload PicDy(TmpCell.Index)
                                    End If
                                End If
                            End With
                        End With
                    Next
                End If
            Next
            '新增类单元格,行为新增行序，列为所在列序
            For l = 1 To F1Main.MaxCol
                strKey = Document.Cells.Add(lNRow, l)
                With Document.Cells(strKey)
                    .Width = F1Main.ColWidthTwips(l)
                    .Height = F1Main.RowHeight(Decode(lNRow, F1Main.MaxRow, lNRow - 1, F1Main.MinRow, lNRow + 1, lNRow - 1))
                    .对象序号 = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo + 1
                    If lNRow = F1Main.MaxRow Then '追加继承列
                        .Merge = Document.Cells.Cell(lNRow - 1, l).Merge
                        If InStr(Document.Cells.Cell(lNRow - 1, l).MergeRange, ";") > 0 Then '有效合并单元格
                            .MergeRange = Document.Cells.Cell(lNRow - 1, l).MergeRange '追加的行，只可能出现跨列合并的情况，此时新行继承上行的合并属性,行+1列不变
                            If Val(Split(Split(.MergeRange, ";")(0), ",")(0)) <> Val(Split(Split(.MergeRange, ";")(1), ",")(0)) Then '有跨行合并的情况不处理合并
                                .Merge = False
                                .MergeRange = lNRow & "," & l
                            Else
                                .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) + 1 & "," & Split(Split(.MergeRange, ";")(0), ",")(1) & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) + 1 & "," & Split(Split(.MergeRange, ";")(1), ",")(1)
                                .Merge = True
                            End If
                        Else '非合并单元格和合并的无效单元格
                            .Merge = False
                            .MergeRange = lNRow & "," & l
                        End If
                    Else                         '普通增加列不对合并进行处理
                        .MergeRange = lNRow & "," & l
                    End If
                End With
            Next
            F1Main.RowHeight(lNRow) = F1Main.RowHeight(Decode(lNRow, F1Main.MaxRow, lNRow - 1, F1Main.MinRow, lNRow + 1, lNRow - 1))
            Document.Cells.Rows = Document.Cells.Rows + 1
    End Select
    If editType <> TabET_单病历审核 Then Call RefreshF1Main
    mblnInit = False
    Call F1Main.SetSelection(lngStarRow, lngStarCol, lngEndRow, lngEndCol)
    InsertRowCol = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False
End Function
Private Sub DeleteRowCol(ByVal strType As String)
'功能：删除行或列
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long
Dim lmsRow As Long, lmsCol As Long, lmeRow As Long, lmeCol As Long
Dim l As Long, j As Integer, strDelKey As String, lCellCount As Long, lngDel As Long, strChangeKey As String
Dim lMaxR As Long, lMaxC As Long, lR As Long, lC As Long, TmpCell As cTabCell
    On Error GoTo errHand
    
    mblnInit = True
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol) '取得第一个选择区域起止行列
    lCellCount = Document.Cells.Count: lMaxR = Document.Cells.Rows: lMaxC = Document.Cells.Cols
    
    If strType = "Col" Then '删除的行列中不能出现合并单元格
        lngDel = (lngEndCol - lngStarCol) + 1
    Else
        lngDel = (lngEndRow - lngStarRow) + 1
    End If
    
    If (lngDel = F1Main.MaxRow And strType = "Row") Or (lngDel = F1Main.MaxCol And strType = "Col") Then
        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "警告：" & vbCrLf & "      至少保留一行或一列" & vbCrLf & "      请检查！", True, 2
        mblnInit = False
        Exit Sub
    End If
        
    For l = 1 To lCellCount
        With Document.Cells(l)
            If .Merge And InStr(.MergeRange, ";") > 0 Then '对合并的区域进行判断
                lmsRow = Val(Split(Split(.MergeRange, ";")(0), ",")(0)): lmsCol = Val(Split(Split(.MergeRange, ";")(0), ",")(1))
                lmeRow = Val(Split(Split(.MergeRange, ";")(1), ",")(0)): lmeCol = Val(Split(Split(.MergeRange, ";")(1), ",")(1))
                If strType = "Col" Then '删除列
                    If lmsCol <> lmeCol Then '有跨列合并的情况
                    For j = lngStarCol To lngEndCol '选中的列中出现被合并情况
                        If j >= lmsCol And j <= lmeCol Then
                            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      删除列时，当前单元格所在列不能有跨列合并单元格" & vbCrLf & "      请检查！", True, 1
                            mblnInit = False: Exit Sub
                        End If
                    Next
                    End If
                Else                    '删除行
                    If lmsRow <> lmeRow Then '有跨行合并的情况
                    For j = lngStarRow To lngEndRow '选中的行中出现被合并的情况
                        If j >= lmsRow And j <= lmeRow Then
                            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      删除行时，当前单元格所在行不能有跨行合并单元格" & vbCrLf & "      请检查！", True, 1
                            mblnInit = False: Exit Sub
                        End If
                    Next
                    End If
                End If
            End If
        End With
    Next
    
    '删除行，列
    Call F1Main.DeleteRange(lngStarRow, lngStarCol, lngEndRow, lngEndCol, Decode(strType, "Row", F1ShiftRows, "Col", F1ShiftCols))
    If strType = "Col" Then '减少列
        F1Main.MaxCol = F1Main.MaxCol - lngDel
        For l = 1 To F1Main.MaxCol
            F1Main.ColText(l) = l '列头显示数字
        Next
    Else                    '减少行
        F1Main.MaxRow = F1Main.MaxRow - lngDel
    End If
    
    '不能在循环中删除集合中的类成员，因为会改变集合数量和索引顺序，先记下Key值
    For l = 1 To lCellCount
        With Document.Cells(l)
            If strType = "Row" Then
                If .Row >= lngStarRow And .Row <= lngEndRow Then strDelKey = strDelKey & "|" & .Key
            Else
                If .Col >= lngStarCol And .Col <= lngEndCol Then strDelKey = strDelKey & "|" & .Key
            End If
        End With
    Next
    '要删除的类成员,同时删除类成员的下属成员
    For l = 1 To UBound(Split(strDelKey, "|"))
        Call ClearChildMember(Split(strDelKey, "|")(l))
        Call Document.Cells.Remove(Split(strDelKey, "|")(l))
    Next
    
    If strType = "Row" Then
        For lR = 1 To lMaxR
            If lR > lngEndRow Then
                For lC = 1 To lMaxC
                    Set TmpCell = Document.Cells("K" & lR & "_" & lC)
                    With Document
                        .Cells.Remove TmpCell.Key
                        .Cells.Add lR - lngDel, lC
                        With .Cells("K" & lR - lngDel & "_" & lC)
                            .ID = TmpCell.ID
                            .文件ID = TmpCell.文件ID
                            .对象序号 = TmpCell.对象序号
                            .对象类型 = TmpCell.对象类型
                            .保留对象 = TmpCell.保留对象
                            .对象属性 = TmpCell.对象属性
                            .内容行次 = TmpCell.内容行次
                            .内容文本 = TmpCell.内容文本
                            .开始版 = TmpCell.开始版
                            .终止版 = TmpCell.终止版
                            .Row = lR - lngDel
                            .Col = lC
                            .Width = TmpCell.Width
                            .Height = TmpCell.Height
                            .FontName = TmpCell.FontName
                            .FontSize = TmpCell.FontSize
                            .FontBold = TmpCell.FontBold
                            .FontItalic = TmpCell.FontItalic
                            .FontUnderline = TmpCell.FontUnderline
                            .FontStrikeout = TmpCell.FontStrikeout
                            .FontColor = TmpCell.FontColor
                            .HAlignment = TmpCell.HAlignment
                            .VAlignment = TmpCell.VAlignment
                            .CellLineTop = TmpCell.CellLineTop
                            .CellLineBottom = TmpCell.CellLineBottom
                            .CellLineLeft = TmpCell.CellLineLeft
                            .CellLineRight = TmpCell.CellLineRight
                            .CellLineTopColor = TmpCell.CellLineTopColor
                            .CellLineBottomColor = TmpCell.CellLineBottomColor
                            .CellLineLeftColor = TmpCell.CellLineLeftColor
                            .CellLineRightColor = TmpCell.CellLineRightColor
                            .Merge = TmpCell.Merge
                            .MergeRange = TmpCell.MergeRange
                            .TextKey = TmpCell.TextKey
                            .ElementKey = TmpCell.ElementKey
                            .PictureKey = TmpCell.PictureKey
                            .SignKey = TmpCell.SignKey
                            .PicMarkKey = TmpCell.PicMarkKey
                            If .Merge And InStr(.MergeRange, ";") > 0 Then '合并单元格首个
                                .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) - lngDel & "," & Split(Split(.MergeRange, ";")(0), ",")(1) & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) - lngDel & "," & Split(Split(.MergeRange, ";")(1), ",")(1)
                            Else
                                .MergeRange = .Row & "," & .Col
                            End If
                            For j = 0 To UBound(Split(.ElementKey, "|")) '改变元素中区域的保存
                                If Len(Split(.ElementKey, "|")(j)) > 0 Then
                                    Document.Elements("K" & Split(.ElementKey, "|")(j)).区域 = .Row & "|" & .Col
                                End If
                            Next
                            For j = 0 To UBound(Split(.TextKey, "|")) '改变文字中区域的保存
                                If Len(Split(.TextKey, "|")(j)) > 0 Then
                                    Document.Texts("K" & Split(.TextKey, "|")(j)).区域 = .Row & "|" & .Col
                                End If
                            Next
                            If .PictureKey <> "" Then
                                If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                    If PicEdit.Visible Then PicEdit.Visible = False
                                    Unload PicDy(TmpCell.Index)
                                End If
                            End If
                        End With
                    End With
                Next
            End If
        Next
    Else
        For lC = 1 To lMaxC      '将删除行、列之后的行列单元格先删除再新增，以确保集合中的关键字与ROW/COL对应
            If lC > lngEndCol Then
                For lR = 1 To lMaxR
                    Set TmpCell = Document.Cells("K" & lR & "_" & lC)
                    With Document
                        .Cells.Remove TmpCell.Key
                        .Cells.Add lR, lC - lngDel
                        With .Cells("K" & lR & "_" & lC - lngDel)
                            .ID = TmpCell.ID
                            .文件ID = TmpCell.文件ID
                            .对象序号 = TmpCell.对象序号
                            .对象类型 = TmpCell.对象类型
                            .保留对象 = TmpCell.保留对象
                            .对象属性 = TmpCell.对象属性
                            .内容行次 = TmpCell.内容行次
                            .内容文本 = TmpCell.内容文本
                            .开始版 = TmpCell.开始版
                            .终止版 = TmpCell.终止版
                            .Row = lR
                            .Col = lC - lngDel
                            .Width = TmpCell.Width
                            .Height = TmpCell.Height
                            .FontName = TmpCell.FontName
                            .FontSize = TmpCell.FontSize
                            .FontBold = TmpCell.FontBold
                            .FontItalic = TmpCell.FontItalic
                            .FontUnderline = TmpCell.FontUnderline
                            .FontStrikeout = TmpCell.FontStrikeout
                            .FontColor = TmpCell.FontColor
                            .HAlignment = TmpCell.HAlignment
                            .VAlignment = TmpCell.VAlignment
                            .CellLineTop = TmpCell.CellLineTop
                            .CellLineBottom = TmpCell.CellLineBottom
                            .CellLineLeft = TmpCell.CellLineLeft
                            .CellLineRight = TmpCell.CellLineRight
                            .CellLineTopColor = TmpCell.CellLineTopColor
                            .CellLineBottomColor = TmpCell.CellLineBottomColor
                            .CellLineLeftColor = TmpCell.CellLineLeftColor
                            .CellLineRightColor = TmpCell.CellLineRightColor
                            .Merge = TmpCell.Merge
                            .MergeRange = TmpCell.MergeRange
                            .TextKey = TmpCell.TextKey
                            .ElementKey = TmpCell.ElementKey
                            .PictureKey = TmpCell.PictureKey
                            .SignKey = TmpCell.SignKey
                            .PicMarkKey = TmpCell.PicMarkKey
                            If .Merge And InStr(.MergeRange, ";") > 0 Then '合并单元格首个
                                .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) & "," & Split(Split(.MergeRange, ";")(0), ",")(1) - lngDel & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) & "," & Split(Split(.MergeRange, ";")(1), ",")(1) - lngDel
                            Else
                                .MergeRange = .Row & "," & .Col
                            End If
                            For j = 0 To UBound(Split(.ElementKey, "|")) '改变元素中区域的保存
                                If Len(Split(.ElementKey, "|")(j)) > 0 Then
                                    Document.Elements("K" & Split(.ElementKey, "|")(j)).区域 = .Row & "|" & .Col
                                End If
                            Next
                            For j = 0 To UBound(Split(.TextKey, "|")) '改变文字中区域的保存
                                If Len(Split(.TextKey, "|")(j)) > 0 Then
                                    Document.Texts("K" & Split(.TextKey, "|")(j)).区域 = .Row & "|" & .Col
                                End If
                            Next
                            If .PictureKey <> "" Then
                                If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                    If PicEdit.Visible Then PicEdit.Visible = False
                                    Unload PicDy(TmpCell.Index)
                                End If
                            End If
                        End With
                    End With
                Next
            End If
        Next
    End If
    
    '变更类集合中行列总数
    If strType = "Row" Then
        Document.Cells.Rows = Document.Cells.Rows - lngDel
        lngStarRow = lngStarRow - lngDel: lngEndRow = lngEndRow - lngDel
    Else
        Document.Cells.Cols = Document.Cells.Cols - lngDel
        lngStarCol = lngStarCol - lngDel: lngEndCol = lngEndCol - lngDel
    End If
    
    If lngStarRow < 1 Then lngStarRow = 1: If lngEndRow < 1 Then lngEndRow = 1
    If lngStarCol < 1 Then lngStarCol = 1: If lngEndCol < 1 Then lngEndCol = 1
    Call RefreshF1Main
    mblnInit = False
    Call F1Main.SetSelection(lngStarRow, lngStarCol, lngEndRow, lngEndCol)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertInherit(ByVal strType As String)
'功能：插入继承行列
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long, cbrObj As CommandBarControl
Dim vCell As F1CellFormat, l As Long, j As Integer, strDelKey As String, lngCellCount As Long, lngDel As Long, strChangeKey As String
    On Error GoTo errHand
    mblnInit = True
    Call F1Main.SetSelection(F1Main.MaxRow, F1Main.MaxCol, F1Main.MaxRow, F1Main.MaxCol)
    If strType = "Row" Then
        If Not InsertRowCol("InsertDnRow") Then Exit Sub    '先插入行，最后行列下方或右边
        mblnInit = True
        For l = 1 To F1Main.MaxCol '内容复制,先复制单元格(包括下级成员的Key),再根据下级成员Key新生下级成员对象
            Call Document.Cells.Cell(F1Main.MaxRow - 1, l).Clone(Document.Cells.Cell(F1Main.MaxRow, l))
            With Document.Cells.Cell(F1Main.MaxRow, l)
                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                    Call CloneChildMember(.Key)
                End If
            End With
'           在InsertRowCol中已处理过 Document.Cells.Cell(F1Main.MaxRow, l).对象序号 = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo '追加行是在编辑时，对象序号应取最大值
        Next
    Else
        If Not InsertRowCol("InsertRightCol") Then Exit Sub '先插入列，最后行列下方或右边
        mblnInit = True
        For l = 1 To F1Main.MaxRow '内容复制,先复制单元格(包括下级成员的Key),再根据下级成员Key新生下级成员对象
            Call Document.Cells.Cell(l, F1Main.MaxCol - 1).Clone(Document.Cells.Cell(l, F1Main.MaxCol))
            With Document.Cells.Cell(l, F1Main.MaxCol)
                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                    Call CloneChildMember(.Key)
                End If
            End With
'           在InsertRowCol中已处理过 Document.Cells.Cell(l, F1Main.MaxCol).对象序号 = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo
        Next
    End If
    mblnInit = False
    Set cbrObj = cbsMain.FindControl(, ID_FILE_SAVE)
    mblnAdd = True
    If cbrObj.Enabled And cbrObj.Visible Then cbsMain_Execute cbrObj
    mblnAdd = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False
End Sub
Private Sub CloneChildMember(ByVal strCellKey As String)
'功能:复制指定单元格的下级成员，strCellKey表示成员所在单元格的Key
Dim i As Integer, lngKey As Long, strNewKey As String, vCell As F1CellFormat, strShow As String
    With Document.Cells(strCellKey)
        For i = 1 To UBound(Split(.TextKey, "|"))
            lngKey = Document.Texts.Add: strNewKey = strNewKey & "|" & lngKey
            Call Document.Texts("K" & Split(.TextKey, "|")(i)).Clone(Document.Texts("K" & lngKey))
            Document.Texts("K" & lngKey).父ID = .ID
            Document.Texts("K" & lngKey).开始版 = 1
            Document.Texts("K" & lngKey).终止版 = 0
            Document.Texts("K" & lngKey).区域 = .Row & "|" & .Col
        Next
        .TextKey = strNewKey: strNewKey = ""
        
        If .ElementKey <> "" Then
            For i = 0 To UBound(Split(.ElementKey, "|"))
                If Split(.ElementKey, "|")(i) = "" Then '可能是混合区域,混合区域从1开始
                    lngKey = Document.Elements.Add: strNewKey = strNewKey & "|" & lngKey
                    Call Document.Elements("K" & Split(.ElementKey, "|")(i)).Clone(Document.Elements("K" & lngKey))
                    Document.Elements("K" & lngKey).父ID = .ID
                    Document.Elements("K" & lngKey).区域 = .Row & "|" & .Col
                    Document.Elements("K" & lngKey).开始版 = 1
                    Document.Elements("K" & lngKey).终止版 = 0
                End If
            Next
        End If
        .ElementKey = strNewKey: strNewKey = ""
        
        For i = 1 To UBound(Split(.PicMarkKey, "|"))
            lngKey = Document.PicMarks.Add: strNewKey = strNewKey & "|" & lngKey
            Call Document.PicMarks("K" & Split(.PicMarkKey, "|")(i)).Clone(Document.PicMarks("K" & lngKey))
            Document.PicMarks("K" & lngKey).父ID = .ID
            Document.PicMarks("K" & lngKey).开始版 = 1
            Document.PicMarks("K" & lngKey).终止版 = 0
        Next
        .PicMarkKey = strNewKey: strNewKey = ""
        
        If .PictureKey <> "" Then
            lngKey = Document.Pictures.Add: strNewKey = lngKey
            Call Document.Pictures("K" & .PictureKey).Clone(Document.Pictures("K" & lngKey))
            Document.Pictures("K" & lngKey).PicID = .ID
        End If
        .PictureKey = strNewKey: strNewKey = ""
        
        If .SignKey <> "" Then
            lngKey = Document.Signs.Add: strNewKey = lngKey
            Call Document.Signs("K" & .SignKey).Clone(Document.Signs("K" & lngKey))
        End If
        .SignKey = strNewKey: strNewKey = ""
        
        '成员内容或图片
        Select Case .对象类型
            Case cprCTPicture, cprCTReportPic
                strShow = IIf(.对象类型 = cprCTReportPic, "报告图", "参考图")
                .内容文本 = strShow
                PaintPictureOnTable strCellKey
            Case cprCTSign, cprCTColSign, cprCTRowSign
                strShow = "[签名位]"
            Case Else
                strShow = .内容文本
                DocOld.Add .内容文本, .Key
        End Select
        .开始版 = 1: .终止版 = 0
        F1Main.TextRC(.Row, .Col) = strShow
        Call F1Main.SetSelection(.Row, .Col, .Row, .Col)
        Set vCell = F1Main.GetCellFormat
        vCell.FontName = .FontName: vCell.FontSize = .FontSize: vCell.FontBold = .FontBold: vCell.FontColor = .FontColor
        vCell.FontItalic = .FontItalic: vCell.FontUnderline = .FontUnderline: vCell.FontStrikeout = .FontStrikeout
        vCell.BorderStyle(F1TopBorder) = .CellLineTop: vCell.BorderStyle(F1BottomBorder) = .CellLineBottom
        vCell.BorderStyle(F1LeftBorder) = .CellLineLeft: vCell.BorderStyle(F1RightBorder) = .CellLineRight
        vCell.BorderColor(F1TopBorder) = .CellLineTopColor: vCell.BorderColor(F1BottomBorder) = .CellLineBottomColor
        vCell.BorderColor(F1LeftBorder) = .CellLineLeftColor: vCell.BorderColor(F1RightBorder) = .CellLineRightColor
        Call F1Main.SetCellFormat(vCell)
    End With
End Sub

Private Sub ClearChildMember(ByVal strCellKey As String)
'功能:清除指定单元的下级成员,不删除成员类只处理Key串，因为会改变索引顺序
Dim j As Long
    With Document.Cells(strCellKey)
        .TextKey = ""
        .ElementKey = ""
        .PicMarkKey = ""
        
        If .PictureKey <> "" Then
            If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                If PicEdit.Visible Then PicEdit.Visible = False
                Unload PicDy(.Index)
            End If
            .PictureKey = ""
        End If
        
        .SignKey = ""
        
        .对象类型 = 0: .对象属性 = "": .内容行次 = 0: .内容文本 = ""
    End With
End Sub
Public Sub PaintPictureOnTable(ByVal strCellKey As String)
'功能:在指定单元格绘图
Dim objTmp As Object, vR As F1Rect, i As Integer, lHheight As Long, lHwidth As Long, lpLeft As Long, lpTop As Long '图片框,区域,固定列高度,固定行宽度,图片框XY坐标
Dim lsRow As Long, leRow As Long, lsCol As Long, leCol As Long '区域起止行列
Dim lsPosX As Long, lsPosY As Long, lpHeight As Long, lpWidth As Long '图片源剪切XY坐标,图片高宽

    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '固定行高度
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '固定列宽度

    With Document.Cells(strCellKey)
        If .PictureKey = "" Then Exit Sub
        If Document.Pictures("K" & .PictureKey).OrigPic.Handle = 0 Then Exit Sub
        
        '确定图片框所在区域
        If .Merge Then  'MergeRange数据格式 (左上方)行,列;(右下方)行,列
            lsRow = Split(Split(.MergeRange, ";")(0), ",")(0): leRow = Split(Split(.MergeRange, ";")(1), ",")(0)
            lsCol = Split(Split(.MergeRange, ";")(0), ",")(1): leCol = Split(Split(.MergeRange, ";")(1), ",")(1)
        Else
            lsRow = .Row: leRow = .Row: lsCol = .Col: leCol = .Col
        End If
        Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
        '确定图片框大小及位置及裁剪坐标
        If vR.Right - lHwidth <= 0 Or vR.Bottom - lHheight <= 0 Then '不在可显示区域
            If ChkControl(PicDy(.Index)) Then
                PicDy(.Index).Visible = False
            End If
            Exit Sub
        ElseIf vR.Left >= 0 And vR.Top >= 0 Then '区域处在表格中间
            lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + vR.Top: lpWidth = vR.Width: lpHeight = vR.Height: lsPosX = 0: lsPosY = 0
        ElseIf vR.Left >= 0 And vR.Top < 0 Then '区域上方部份隐藏(滚动引起)
            lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + lHheight: lpWidth = vR.Width: lpHeight = vR.Height + vR.Top - lHheight: lsPosX = 0: lsPosY = vR.Height - lpHeight
        ElseIf vR.Left < 0 And vR.Top >= 0 Then '区域左方部份隐藏(滚动引起)
            lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + vR.Top: lpWidth = vR.Width + vR.Left - lHwidth: lpHeight = vR.Height: lsPosX = vR.Width - lpWidth: lsPosY = 0
        ElseIf vR.Left < 0 And vR.Top < 0 Then '区域上方左方都隐藏(滚动引起)
            lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + lHheight: lpWidth = vR.Width + vR.Left - lHwidth: lpHeight = vR.Height + vR.Top - lHheight: lsPosX = vR.Width - lpWidth: lsPosY = vR.Height - lpHeight
        Else                                    '不在可显示区域
            If ChkControl(PicDy(.Index)) Then
                PicDy(.Index).Visible = False
            End If
            Exit Sub
        End If
        
        '动态加载图片框数组
        If Not ChkControl(PicDy(.Index)) Then
            Load PicDy(.Index)
        End If
        Set objTmp = PicDy(.Index): objTmp.Cls
        objTmp.Tag = .MergeRange & "|" & strCellKey: objTmp.ToolTipText = IIf(.对象类型 = cprCTReportPic, "报告图", "参考图")
        objTmp.AutoRedraw = True: objTmp.BorderStyle = 0: Set objTmp.Container = picMainBack
        
        '先晳定图片大小并绘出标记
        LockWindowUpdate Me.hWnd
        objTmp.Width = vR.Width - Screen.TwipsPerPixelX * 2: objTmp.Height = vR.Height - Screen.TwipsPerPixelY * 2
        Set objTmp.Picture = Document.Pictures("K" & .PictureKey).OrigPic
        objTmp.PaintPicture objTmp.Picture, 0, 0, objTmp.Width, objTmp.Height
        If .PicMarkKey <> "" Then '有标记图先绘标记
            For i = 1 To UBound(Split(.PicMarkKey, "|"))
                ShowPicMark objTmp, Me.Document.PicMarks("K" & Split(.PicMarkKey, "|")(i))
            Next
        End If
        Set objTmp.Picture = objTmp.Image
        '最后根据实际显示大小及坐标重绘
        objTmp.Move lpLeft + Screen.TwipsPerPixelX * 2, lpTop + Screen.TwipsPerPixelY * 2, lpWidth - Screen.TwipsPerPixelX * 2, lpHeight - Screen.TwipsPerPixelY * 2
        objTmp.PaintPicture objTmp.Picture, 0, 0, objTmp.Width, objTmp.Height, lsPosX, lsPosY
        objTmp.Visible = True: objTmp.ZOrder
        LockWindowUpdate 0
    End With
End Sub
Private Sub SetCellFont()
Dim tmpFont As New StdFont, tmpColor As Long
'通过字体窗口设置字体
    On Error GoTo errHand
    tmpFont.Name = SelCell.FontName: tmpFont.Size = SelCell.FontSize: tmpFont.Bold = SelCell.FontBold
    tmpFont.Italic = SelCell.FontItalic: tmpFont.Underline = SelCell.FontUnderline: tmpFont.Strikethrough = SelCell.FontStrikeout
    tmpColor = SelCell.FontColor
    If SetFont(Me.hWnd, Me.hdc, tmpFont, tmpColor) Then
        Call SetCellFormat("字号", tmpFont.Size)
        Call SetCellFormat("字体名称", tmpFont.Name)
        Call SetCellFormat("粗体", tmpFont.Bold)
        Call SetCellFormat("斜体", tmpFont.Italic)
        Call SetCellFormat("下划线", tmpFont.Underline)
        Call SetCellFormat("删除线", tmpFont.Strikethrough)
        Call SetCellFormat("字体颜色", tmpColor)
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub DesignTest()
'功能：此过程用于在设计和调试阶段查看内存变量及相关数据
Dim picTmp As New StdPicture
    On Error GoTo errHand
       Debug.Print SelCell.内容文本
       
       If App.LogMode = 0 Then MsgBox "哈哈": Stop
       

       
       
       
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function AllowEdit(nRow As Long, nCol As Long) As Boolean
'功能：判断是否处于保护状态,禁止编辑并给出提示
    On Error GoTo errHand
    Select Case mReadOnly
        Case 1
            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      当前编辑状态仅能用于回退" & vbCrLf & "      不能保存和签名！", True, 1
            Exit Function
        Case 2
            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      当前处于查看状态" & vbCrLf & "      不能编辑！", True, 1
            Exit Function
    End Select
    
    If Document.Cells.Cell(nRow, nCol).保留对象 = True Or Document.Cells.Cell(nRow, nCol).对象类型 = cprCTFixtext Then
        Call Beep
        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      所选区域处于保护的控制区域" & vbCrLf & "      不能编辑！", True, 1
        Exit Function
    End If
    
    Dim lRows As Long, lCols As Long, l As Long
    lRows = Document.Cells.Rows '检查所在列是否有列控签名已签名
    For l = 1 To lRows
        With Document.Cells.Cell(l, nCol)
            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                If .对象类型 = cprCTColSign And .终止版 <> 0 Then
                    Call Beep
                    mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      所选区域处于已签名的控制区域" & vbCrLf & "      不能编辑！", True, 1
                    Exit Function
                End If
            End If
        End With
    Next
    
    lCols = Document.Cells.Cols '检查所在行是否有行控签名已签名
    For l = 1 To lCols
        With Document.Cells.Cell(nRow, l)
            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                If .对象类型 = cprCTRowSign And .终止版 <> 0 Then
                    Call Beep
                    mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提醒：" & vbCrLf & "      所选区域处于已签名的控制区域" & vbCrLf & "      不能编辑！", True, 1
                    Exit Function
                End If
            End If
        End With
    Next
    
    AllowEdit = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub EnterEdit(ByVal lsRow As Long, ByVal lsCol As Long, ByVal leRow As Long, ByVal leCol As Long, Optional KeyAscii As Integer, Optional DbClick As Boolean)
'功能：进入编辑状态
Dim vR As F1Rect
    On Error GoTo errHand
    If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
        If Not AllowEdit(lsRow, lsCol) Then F1Main.AllowInCellEditing = False: Exit Sub
    End If
    
    If Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTText Or Document.Cells.Cell(lsRow, lsCol).对象类型 = cprCTFixtext Then
        F1Main.AllowInCellEditing = True
    Else
        F1Main.AllowInCellEditing = False
    End If
    
    With Document.Cells.Cell(lsRow, lsCol)
        If .Merge Then
            Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
        Else
            Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, lsRow, lsCol)
        End If
        Select Case .对象类型
            Case cprCTText, cprCTFixtext
                If DbClick Then
                    Call F1Main.StartEdit(False, True, False)
                Else
                    Call F1Main.StartEdit(True, True, False)
                End If
            Case cprCTElement
                If Document.Elements("K" & .ElementKey).要素名称 = "" Then
                    If (editType = TabET_病历文件定义 Or TabET_全文示范编辑) Then
                        Call InsertElement(.Key)
                    End If
                Else
                    If (editType = TabET_病历文件定义 Or TabET_全文示范编辑) And Not DbClick And KeyAscii = 0 Then
                        Call InsertElement(.Key)
                    Else
                        Call EditElement(.Key, KeyAscii)
                    End If
                End If
            Case cprCTTextElement
                Call PopDoc(.Key, KeyAscii)
            Case cprCTPicture
                If KeyAscii <> 0 Then KeyAscii = 0: Exit Sub '键盘输入对图片无效
                If .PictureKey = "" Then
                    Call InsertPicture(.Key)
                Else '进入图片编辑
                    If Document.Pictures("K" & .PictureKey).OrigPic = 0 Then
                        Call InsertPicture(.Key, .PictureKey)
                    End If
                End If
            Case cprCTReportPic
                If KeyAscii <> 0 Then KeyAscii = 0: Exit Sub '键盘输入对图片无效
                If Not dkpMain.FindPane(conPane_PacsPic) Is Nothing Then
                    If Not dkpMain.FindPane(conPane_PacsPic).Closed Then dkpMain.ShowPane conPane_PacsPic
                End If
                Call F1Main.SetFocus
            Case cprCTSign, cprCTColSign, cprCTRowSign
                If KeyAscii <> 0 And KeyAscii <> vbKeySpace Then KeyAscii = 0: Exit Sub '键盘输入对图片无效
                If Not (editType = TabET_病历文件定义 Or TabET_全文示范编辑) And mReadOnly = 0 Then
                    If SaveDoc(True, True) Then Unload Me
                End If
        End Select
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertPicture(ByVal strCellKey As String, Optional ByVal strPicKey As String)
'功能：插入图片,当strPicKey<>""时表示换掉当前图片(由工具条调用)
'此时已经确定选中的只有一个CELL并且是图片类型
Dim tmpPic As StdPicture, lngKey As Long, vR As F1Rect, l As Long, ary As Variant
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    On Error GoTo errHand
    
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
    If frmInsertPicture.ShowMe(Me, vR.Width, vR.Height, tmpPic) Then
        AddUndo Document.Cells(strCellKey)
        If strPicKey = "" Then
            lngKey = Document.Pictures.Add
        Else
            lngKey = Val(strPicKey) '换图时Key值不变
            If ChkControl(PicDy(Document.Cells(strCellKey).Index)) Then Unload PicDy(Document.Cells(strCellKey).Index)
        End If
        Set Document.Pictures("K" & lngKey).OrigPic = tmpPic '加载图片
        Document.Pictures("K" & lngKey).DesHeight = vR.Height
        Document.Pictures("K" & lngKey).DesWidth = vR.Width
        Document.Cells(strCellKey).PictureKey = lngKey
        Call PaintPictureOnTable(strCellKey) '重绘图片和标记
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertOtherText(ByVal strCellKey As String, ByVal strType As String)
'功能：插入日期，时间，特殊符号，目标单元格可能是Text和混合编辑区域
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, vR As F1Rect, lHheight As Long, lHwidth As Long, elTmpKey As Long
    On Error GoTo errHand
    If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
        If Not AllowEdit(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) Then Exit Sub
    End If
    If strType = "特殊符号" Then
        Dim strTmp As String
        strTmp = frmInsSymbol.ShowMe(Decode(mstrSex, "男", 1, "女", 2, 0))
        If strTmp <> "" Then
            With Document.Cells(strCellKey) '只有文本和混合区域可以插入
                AddUndo Document.Cells(strCellKey)
                If .对象类型 = cprCTText Or .对象类型 = cprCTFixtext Then
                    strTmp = .内容文本 & strTmp
                    .内容文本 = strTmp
                    F1Main.TextRC(.Row, .Col) = strTmp
                Else
                    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
                    bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '处在关键字对之间
                    If bInKeys Then
                        Doc.Range(lEE, lEE).Selected
                        Doc.Range(lEE, lEE).Font.Protected = False
                        Doc.Range(lEE, lEE).Font.Hidden = False
                        Doc.Range(lEE, lEE).Text = IIf(Mid(strTmp, 1, 1) <> "，", "，" & strTmp, strTmp)
                        Doc.Range(lEE + Len(strTmp), lEE + Len(strTmp)).Selected
                    Else
                        If Doc.Range(Doc.Selection.StartPos - 1, Doc.Selection.StartPos).Font.Hidden Then strTmp = "，" & strTmp '如果在关键之后
                        Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Selected
                        Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Font.Protected = False
                        Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Font.Hidden = False
                        Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Text = strTmp
                        Doc.Range(Doc.Selection.StartPos + Len(strTmp), Doc.Selection.StartPos + Len(strTmp)).Selected
                    End If
                End If
            End With
        End If
    Else
        elTmpKey = Document.Elements.Add '初始化一个空临时要素
        With Document.Elements("K" & elTmpKey)
               .内容文本 = ""
               .内容行次 = 0
               .诊治要素ID = 0
               .替换域 = 0
               .要素名称 = strType
               .要素类型 = 2
               .要素长度 = Decode(strType, "日期时间", 19, "日期", 10, "时间", 8, 19)
               .要素小数 = 0
               .要素单位 = ""
               .要素表示 = 0
               .输入形态 = 0
               .要素值域 = "0;0"
               .保留对象 = False
               .自动转文本 = True
               .必填 = 0
        End With
        If Doc.Visible Then
            Call ShowElInDoc(Doc.Selection.StartPos, Doc.Selection.StartPos, elTmpKey)
        Else
            Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
            If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '固定行高度
            If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '固定列宽度
            Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
            
            elEdit.Move F1Main.Left + IIf(vR.Left < 0, lHwidth, vR.Left) + Screen.TwipsPerPixelX * 2, F1Main.Top + vR.Bottom + Screen.TwipsPerPixelX * 2: elEdit.Tag = strCellKey
            elEdit.SetElement Document.Elements("K" & elTmpKey), 0, editType
            
            If elEdit.Top + elEdit.Height > F1Main.Top + F1Main.Height Then
                elEdit.Top = vR.Top - elEdit.Height - Screen.TwipsPerPixelY * 2
            End If
            
            If elEdit.Left + elEdit.Width > F1Main.Left + F1Main.Width Then
                elEdit.Left = vR.Left - elEdit.Width - Screen.TwipsPerPixelX * 2
            End If
            elEdit.Visible = True: elEdit.ZOrder 0: elEdit.SetFocus
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub EditPicture(ByVal strKey As String)
'图片编辑,strKey图片所在单元的KEY
Dim tmpPic As StdPicture, lngKey As Long, vR As F1Rect
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    On Error GoTo errHand
    If PicEdit.Visible Then PicEdit_LostFocus  '在两个图之间切换时，先丢失焦点以便重绘原图
    With Document.Cells(strKey)
        If .保留对象 = True Then Exit Sub
        If .PictureKey = "" Then Exit Sub
        If Document.Pictures("K" & .PictureKey).OrigPic.Handle = 0 Then Exit Sub
    '确定图片框所在区域
        If .Merge Then  'MergeRange数据格式 (左上方)行,列;(右下方)行,列
            lsRow = Split(Split(.MergeRange, ";")(0), ",")(0): leRow = Split(Split(.MergeRange, ";")(1), ",")(0)
            lsCol = Split(Split(.MergeRange, ";")(0), ",")(1): leCol = Split(Split(.MergeRange, ";")(1), ",")(1)
        Else
            lsRow = .Row: leRow = .Row: lsCol = .Col: leCol = .Col
        End If
    End With
    AddUndo Document.Cells(strKey)
    Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
    If vR.Top < -15 Or vR.Left < -15 Then Exit Sub '部份超过可视区域,禁止编辑
    PicEdit.Move F1Main.Left + vR.Left + Screen.TwipsPerPixelX * 2, F1Main.Top + vR.Top + Screen.TwipsPerPixelY * 2, vR.Width - Screen.TwipsPerPixelX * 2, vR.Height - Screen.TwipsPerPixelY * 2
    PicEdit.Tag = IIf(Document.Cells(strKey).对象类型 = cprCTPicture, "参考图", "报告图")
    Call PicEdit.EditPic(Document, cbsMain, strKey)
    PicEdit.Visible = True: PicEdit.SetFocus: PicEdit.ZOrder 0
    '弹出菜单
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertElement(ByVal strCellKey As String)
'功能： 插入要素，当要素名称为空时插入新要素，非空时修改要素
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, mfrmInsElement As New frmInsElement, strShow As String
    On Error GoTo errHand
    With Document.Cells(strCellKey)
        If .对象类型 = cprCTElement Then
             strShow = F1Main.TextRC(.Row, .Col)
            If mfrmInsElement.ShowMe(Me, Document.Elements("K" & .ElementKey), True, True) Then
                If Document.Elements("K" & .ElementKey).输入形态 = 1 And Document.Elements("K" & .ElementKey).要素类型 <> 2 Then
                    strShow = Document.Elements("K" & .ElementKey).内容文本 & Document.Elements("K" & .ElementKey).要素单位
                    .内容文本 = Document.Elements("K" & .ElementKey).内容文本
                Else
                    strShow = "[" & Document.Elements("K" & .ElementKey).要素名称 & "]" & Document.Elements("K" & .ElementKey).要素单位
                End If
            End If
            F1Main.TextRC(.Row, .Col) = strShow
        Else
            Doc.Tag = strCellKey
            Call InsertElementInRich(strCellKey)
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertElementInRich(ByVal strCellKey As String)
'功能：在混合编辑区中插入要素,当前处在要素关键字中间为修改要素,当DOC不可见即未处于编辑状态时不会进入本函数
Dim lngCp As Long, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, loldKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean, mfrmInsElement As New frmInsElement
Dim blnAutoSpace As Boolean
    On Error GoTo errHand
    lngCp = Doc.Selection.StartPos
    bBeteenKeys = IsBetweenAnyKeys(Doc, lngCp + 1, sKeyType, lKSS, lKSE, lKES, lKEE, loldKey, bNeeded)
    If bBeteenKeys And Doc.SelLength = 0 Then
    '在要素前后插入要素，需要调整位置
        With Doc
            If .Range(lngCp - 1, lngCp).Font.Protected And .Range(lngCp - 1, lngCp).Font.Hidden = False And .Range(lngCp, lngCp + 1).Font.Hidden And .Range(lngCp, lngCp + 3).Text = "EE(" And .Range(lngCp + 16, lngCp + 17).Font.Hidden = False Then
            'B问题1：（隐藏关键字）[要素]|（隐藏关键字）普通文本
                Call .Range(lngCp + 16, lngCp + 16).Selected
                bBeteenKeys = False
            ElseIf .Range(lngCp - 1, lngCp).Font.Protected And .Range(lngCp - 1, lngCp).Font.Hidden = False And .Range(lngCp, lngCp + 1).Font.Hidden And .Range(lngCp, lngCp + 3).Text = "EE(" And .Range(lngCp + 16, lngCp + 19).Text = "ES(" Then
            'B问题1：（隐藏关键字）[要素]|（隐藏关键字）（隐藏关键字）[要素]（隐藏关键字）
                .Range(lngCp + 16, lngCp + 16).Text = " "
                .Range(lngCp + 16, lngCp + 17).Font.Protected = False
                .Range(lngCp + 16, lngCp + 17).Font.Hidden = False
                Call .Range(lngCp + 17, lngCp + 17).Selected
                lngCp = lngCp + 17
                blnAutoSpace = True
                bBeteenKeys = False
            ElseIf .Range(lngCp - 1, lngCp).Font.Hidden And .Range(lngCp, lngCp + 1).Font.Protected And .Range(lngCp, lngCp + 1).Font.Hidden = False And .Range(lngCp - 16, lngCp - 13).Text = "ES(" And .Range(lngCp - 17, lngCp - 16).Font.Hidden = False Then
            'B问题2：普通文本（隐藏关键字）|[要素]（隐藏关键字）
                Call .Range(lngCp - 16, lngCp - 16).Selected
                bBeteenKeys = False
            ElseIf .Range(lngCp - 1, lngCp).Font.Hidden And .Range(lngCp, lngCp + 1).Font.Protected And .Range(lngCp, lngCp + 1).Font.Hidden = False And .Range(lngCp - 16, lngCp - 13).Text = "ES(" And .Range(lngCp - 32, lngCp - 29).Text = "EE(" Then
            'B问题2：（隐藏关键字）[要素]（隐藏关键字）（隐藏关键字）|[要素]（隐藏关键字）
                .Range(lngCp - 16, lngCp - 16).Text = " "
                lngCp = lngCp + 1
                .Range(lngCp - 17, lngCp - 16).Font.Protected = False
                .Range(lngCp - 17, lngCp - 16).Font.Hidden = False
                Call .Range(lngCp - 16, lngCp - 16).Selected
                lngCp = lngCp - 16
                blnAutoSpace = True
                bBeteenKeys = False
            ElseIf .Range(lngCp - 1, lngCp).Font.Hidden And .Range(lngCp, lngCp + 1).Font.Protected And .Range(lngCp, lngCp + 1).Font.Hidden = False And .Range(lngCp - 16, lngCp + 13).Text = "ES(" And lngCp - 16 = 0 Then
                Call .Range(0, 0).Selected
                bBeteenKeys = False
            End If
        End With
    End If
    With Document
        If bBeteenKeys Then '修改要素
            If sKeyType = "E" Then
                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提示：" & vbCrLf & "      当前光标所在位置只能修改要素" & vbCrLf & "      若需在两项要素之间插入要素请先输入空格！", True, 0
                If mfrmInsElement.ShowMe(Me, Document.Elements("K" & loldKey), True, True, False, True) Then
                    If .Elements("K" & loldKey).替换域 = 1 And (editType = TabET_单病历编辑 Or editType = TabET_单病历审核) Then '编辑时插入要素
                        .Elements("K" & loldKey).内容文本 = GetReplaceEleValue(.Elements("K" & loldKey).要素名称, .EPRPatiRecInfo.病人ID, .EPRPatiRecInfo.主页ID, .EPRPatiRecInfo.病人来源, .EPRPatiRecInfo.医嘱id, .EPRPatiRecInfo.婴儿)
                    End If
                    Call .Elements("K" & loldKey).Refresh(Doc)
                End If
            End If
        Else '新增要素
            Dim NewElement As New cTabElement, lnewKey As Long
            If mfrmInsElement.ShowMe(Me, NewElement, True, True, False, True) Then
                lnewKey = .Elements.Add                                          '新增要素
                Call NewElement.Clone(.Elements("K" & lnewKey))                  '将新要素内容取出
                If .Elements("K" & lnewKey).替换域 = 1 And (editType = TabET_单病历编辑 Or editType = TabET_单病历审核) Then '编辑时插入要素
                    .Elements("K" & lnewKey).内容文本 = GetReplaceEleValue(.Elements("K" & lnewKey).要素名称, .EPRPatiRecInfo.病人ID, .EPRPatiRecInfo.主页ID, .EPRPatiRecInfo.病人来源, .EPRPatiRecInfo.医嘱id, .EPRPatiRecInfo.婴儿)
                End If
                .Elements("K" & lnewKey).区域 = .Cells(strCellKey).Row & "|" & .Cells(strCellKey).Col
                Call .Elements("K" & lnewKey).InsertIntoEditor(Doc, editType)  '刷新显示
                If blnAutoSpace Then '在要素之间插入要素，先自动追加空格，插入要素后删除
                    If FindKey(Doc, "E", lnewKey, lKSS, lKSE, lKES, lKEE, bNeeded) Then
                        If Doc.Range(lKSS - 1, lKSS).Text = " " Then
                            Doc.Range(lKSS - 1, lKSS).Text = ""
                        End If
                    End If
                End If
                Call GetFromDoc(strCellKey, False)
                If Doc.Enabled And Doc.Visible Then Doc.SetFocus
            End If
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub GetTextELement(ByVal strCellKey As String)
'功能：跟据Text Element填写F1Main中的单元格及类的内容文本
Dim i As Long, lCount As Long, strTmp As String, ltCount As Long, leCount As Long, cleTmp As cTabElement, strAEl As String
    On Error GoTo errHand
    With Document.Cells(strCellKey)
        ltCount = UBound(Split(.TextKey, "|")): If ltCount < 0 Then ltCount = 0
        leCount = UBound(Split(.ElementKey, "|")): If leCount < 0 Then leCount = 0
        lCount = ltCount + leCount
        For i = 1 To lCount
            Set cleTmp = .clElement(Document.Elements, i)
            If cleTmp Is Nothing Then '该次序为文本
                strTmp = strTmp & ToVarchar(.clText(Document.Texts, i).内容文本, 4000)
            Else
                With Document.Elements("K" & cleTmp.Key)
                    If .替换域 = 1 And (editType = TabET_单病历编辑 Or editType = TabET_单病历审核) Then
                        If Trim(.内容文本) = "" Then
                            strAEl = GetReplaceEleValue(.要素名称, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.病人来源, Document.EPRPatiRecInfo.医嘱id, Me.Document.EPRPatiRecInfo.婴儿)
                            .内容文本 = strAEl
                            If strAEl = "" Then
                                If .自动转文本 Then
                                    strTmp = strTmp & " " & .要素单位
                                Else
                                    strTmp = strTmp & "[" & .要素名称 & "]" & .要素单位
                                End If
                            Else
                                strTmp = strTmp & strAEl
                            End If
                        Else
                            strTmp = strTmp & .内容文本 & .要素单位
                        End If
                    Else
                        If .输入形态 = 0 Then
                            strTmp = strTmp & IIf(Trim(.内容文本) = "", "[" & .要素名称 & "]", .内容文本) & .要素单位
                        Else
                            strTmp = strTmp & .内容文本 & .要素单位
                        End If
                    End If
                End With
            End If
        Next
        .内容文本 = strTmp
        F1Main.TextRC(.Row, .Col) = strTmp
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub GetFromDoc(ByVal strCellKey As String, ByVal blnRefreshCell As Boolean)
'从DOC中取出TEXT,ELEMENT,重组.textkey;.elementkey,之后刷新F1单元格及类
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bsInKeys As Boolean, sKeyType As String, bNeeded As Boolean
Dim lngEnd As Long, p As Long, strText As String, strTmp As String, lNo As Long
Dim txtKeys As String, elKeys As String, ltKey As Long, leKey As Long
    
    On Error GoTo errHand
    If Not Doc.Visible Then Exit Sub
    
    AddUndo Document.Cells(strCellKey)
    strText = Doc.Text:             lngEnd = Len(Doc.Text)
    p = 0:                          lNo = 1
    Do While p < lngEnd
        
        bsInKeys = FindNextAnyKey(Doc, p, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bsInKeys Then '找到要素
            '先处理要素关键字之前的TXT
            If p <> lKSS Then   '表明要素之前有文本
                strTmp = Mid(strText, p + 1, lKSS - p)
Process:        If strTmp <> "" Then
                    ltKey = Document.Texts.Add
                    Document.Texts("K" & ltKey).内容文本 = ToVarchar(strTmp, 4000)
                    Document.Texts("K" & ltKey).内容行次 = lNo
                    txtKeys = txtKeys & "|" & ltKey
                    lNo = lNo + 1
                End If
            End If
            If p > lngEnd Then Exit Do '可能是最后的文本调用
            
            '再处理要素
            If lKey = 0 Then
                leKey = Document.Elements.Add
            Else
                leKey = lKey
            End If
            Document.Elements("K" & leKey).内容行次 = lNo
            elKeys = elKeys & "|" & leKey
            p = lKEE:          lNo = lNo + 1
        Else            '文本
            '再也找不到下一个要素表明其后没有要素
            strTmp = Mid(strText, p + 1)
            p = lngEnd + 1
            GoTo Process
        End If
    Loop
    
    Document.Cells(strCellKey).TextKey = txtKeys '取得新的文本Key串
    Document.Cells(strCellKey).ElementKey = elKeys '取得新的要素Key串
    If blnRefreshCell Then GetTextELement strCellKey
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshToDoc(ByVal strCellKey As String)
'功能：跟据当前单元格刷新Rich编辑器内容
Dim i As Long, lCount As Long, strTmp As String, ltCount As Long, leCount As Long, cleTmp As cTabElement
    On Error GoTo errHand
    With Document.Cells(strCellKey)
        ltCount = UBound(Split(.TextKey, "|")): If ltCount < 0 Then ltCount = 0
        leCount = UBound(Split(.ElementKey, "|")): If leCount < 0 Then leCount = 0
        lCount = ltCount + leCount
        For i = 1 To lCount
            Set cleTmp = .clElement(Document.Elements, i)
            If cleTmp Is Nothing Then '该次序为文本
                .clText(Document.Texts, i).InsertIntoEditor Doc
            Else
                cleTmp.InsertIntoEditor Doc, editType
            End If
        Next
        Doc.ForceEdit = True
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub EditElement(ByVal strKey As String, ByVal KeyAscii As Integer)
'功能：编辑要素
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, vR As F1Rect, lHheight As Long, lHwidth As Long
    On Error GoTo errHand
    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '固定行高度
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '固定列宽度
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
    elEdit.Tag = strKey
    With Document.Elements("K" & Document.Cells(strKey).ElementKey)
        If KeyAscii = vbKeySpace Then KeyAscii = 0
        If (Not (.要素表示 = 2 Or .要素表示 = 3)) And KeyAscii <> 0 Then Exit Sub  '除多\单选,数值型以外，其它类型只允许空格和双击进入编辑状态
        If (.要素表示 = 2 Or .要素表示 = 3) And InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 0 Then Exit Sub '多、单选要素允许空格和双击和数值
        'If .要素类型 = 0 And .要素表示 = 0 And InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 0 Then Exit Sub '数值型要素允许空格和双击和数值
        
        If .输入形态 = 1 Then
            If vR.Top < -15 Or vR.Left < -15 Then Exit Sub '部份超过可视区域,禁止编辑
            elEdit.Top = F1Main.Top + vR.Top + Screen.TwipsPerPixelX * 2: elEdit.Left = F1Main.Left + vR.Left + Screen.TwipsPerPixelX * 2
            elEdit.Width = vR.Width - Screen.TwipsPerPixelX * 2: elEdit.Height = vR.Height - Screen.TwipsPerPixelY * 2
        Else
            elEdit.Top = F1Main.Top + vR.Bottom + Screen.TwipsPerPixelX * 2: elEdit.Left = F1Main.Left + IIf(vR.Left < 0, lHwidth, vR.Left) + Screen.TwipsPerPixelX * 2
        End If
    End With
    elEdit.SetElement Document.Elements("K" & Document.Cells(strKey).ElementKey), KeyAscii, editType
    
    If Document.Elements("K" & Document.Cells(strKey).ElementKey).输入形态 = 0 Then
        If elEdit.Top + elEdit.Height > F1Main.Top + F1Main.Height Then
            elEdit.Top = vR.Top - elEdit.Height - Screen.TwipsPerPixelY * 2
        End If
        
        If elEdit.Left + elEdit.Width > F1Main.Left + F1Main.Width Then
            elEdit.Left = vR.Left - elEdit.Width - Screen.TwipsPerPixelX * 2
        End If
    End If

    elEdit.Visible = True: elEdit.ZOrder 0: elEdit.SetFocus
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub PopDoc(ByVal strCellKey As String, ByVal KeyAscii As Integer, Optional ByVal blnNew As Boolean = True)
'功能：显示Rich控件,blnNew=true 表示初始化控件并加载内容,=false表示已有内容，只作显示
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, vR As F1Rect, lHheight As Long, lHwidth As Long
Dim lpLeft As Long, lpTop As Long, lrHeight As Long, lrWidth As Long 'XY坐标,高宽
    On Error GoTo errHand
    Doc.Title = "编辑" '表示进入编辑状态
    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '固定行高度
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '固定列宽度
    
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
    If vR.Right - lHwidth <= 0 Or vR.Bottom - lHheight <= 0 Then '不在可显示区域
        Doc.Title = "": Call F1Main_GotFocus: Exit Sub
    ElseIf vR.Left >= 0 And vR.Top >= 0 Then '区域处在表格中间
        lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + vR.Top: lrWidth = vR.Width: lrHeight = vR.Height
    ElseIf vR.Left >= 0 And vR.Top < 0 Then '区域上方部份隐藏(滚动引起)
        lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + lHheight: lrWidth = vR.Width: lrHeight = vR.Height + vR.Top - lHheight
    ElseIf vR.Left < 0 And vR.Top >= 0 Then '区域左方部份隐藏(滚动引起)
        lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + vR.Top: lrWidth = vR.Width + vR.Left - lHwidth: lrHeight = vR.Height
    ElseIf vR.Left < 0 And vR.Top < 0 Then '区域上方左方都隐藏(滚动引起)
        lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + lHheight: lrWidth = vR.Width + vR.Left - lHwidth: lrHeight = vR.Height + vR.Top - lHheight
    Else '意外,未知
        Doc.Title = "": Call F1Main_GotFocus: Exit Sub
    End If
    '控件定位
    Doc.Move lpLeft + Screen.TwipsPerPixelX * 2, lpTop + Screen.TwipsPerPixelY * 2, lrWidth - Screen.TwipsPerPixelX * 2, lrHeight - Screen.TwipsPerPixelY * 2
    Doc.Tag = strCellKey
    If blnNew Then
        Doc.NewDoc
        RefreshToDoc strCellKey
    End If
    '控件焦点
    Doc.Visible = True: Doc.ZOrder 0: Doc.SetFocus: Doc.ForceEdit = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ClearPicture()
'功能：新建空文档前清空已的的图片,在编辑的图片，文本，要素控件隐蔽
Dim l As Long, lCount As Long
    On Error Resume Next
    lCount = Document.Cells.Count
    For l = 1 To lCount
        With Document.Cells.Item(l)
            If .PictureKey <> "" Then
                ClearChildMember .Key
            End If
        End With
    Next
    If PicEdit.Visible Then PicEdit.Visible = False
    If elEdit.Visible Then elEdit.Visible = False
    If Doc.Visible Then Doc.Visible = False
End Sub
Private Sub CalcSumRange(ByVal nRow As Long, ByVal nCol As Long)
'功能:计算指定单元格由哪些单元格合计得来
Dim SumRange As String, subRange As String, SumVal As Double, l As Long

    With Document.Cells.Cell(nRow, nCol) '合计单元格
        SumRange = .对象属性 '合计单元格的源单元格串
        If UBound(Split(SumRange, ";")) > 0 Then
            For l = 0 To UBound(Split(SumRange, ";"))
                subRange = Split(SumRange, ";")(l)
                SumVal = SumVal + Val(Document.Cells.Cell(Split(subRange, ",")(0), Split(subRange, ",")(1)).内容文本)
            Next
            .内容文本 = Format(SumVal, "0.00")
            F1Main.TextRC(nRow, nCol) = Format(SumVal, "0.00")
        End If
    End With
End Sub

Private Function ValiCellDate(Optional DataVerify As Boolean = True) As Boolean
'功能：1 将 不能通过事件  将数据保存到类中的数据 集中保存,拖动改变的行高、列宽
'      2 保存前对数据进行较验 , 目前只较验要素在病历定义时
Dim l As Long, lCount As Long, lngWidth As Long, lngHeight As Long, blnChangeRC As Boolean
    If timeTmp.Enabled Then timeTmp.Enabled = False: blnChangeRC = True
    On Error GoTo errHand
    lCount = Document.Cells.Count
    For l = 1 To lCount
        With Document.Cells(l)
            If (editType = TabET_病历文件定义 Or TabET_全文示范编辑) Then
                On Error Resume Next
                .Width = F1Main.ColWidthTwips(.Col)
                .Height = F1Main.RowHeight(.Row)
                Err.Clear
            End If
            
            If DataVerify Then
                On Error GoTo errHand
                Select Case .对象类型
                    Case cprCTElement
                        If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                            If .ElementKey <> "" Then
                                If Document.Elements("K" & .ElementKey).要素名称 = "" Then
                                    MsgBox .Row & "行" & .Col & "列 " & "为要素单元格,但未指定具体要素！", vbInformation, gstrSysName
                                    Call F1Main.SetSelection(.Row, .Col, .Row, .Col)
                                    Exit Function
                                End If
                            Else
                                MsgBox .Row & "行" & .Col & "列 " & "为要素单元格,但未指定具体要素！", vbInformation, gstrSysName
                                Call F1Main.SetSelection(.Row, .Col, .Row, .Col)
                                Exit Function
                            End If
                        End If
                    Case Else
                End Select
            End If
        End With
    Next
    
    If DataVerify Then
        If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
            If Not mfrmMainError.ShowNotice(Me) Then Exit Function
            Me.Refresh
        End If
    End If
    
    If blnChangeRC Then F1Main_SelChange
    
    ValiCellDate = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub ContentMove(strType As String)
'功能:全选,剪切,复制,粘贴,目前只支持文本,和混合编辑区域
Dim strCellKey As String, strTmp As String
    On Error GoTo errHand
    If SelCell Is Nothing Then Exit Sub
    strCellKey = SelCell.Key
    If strCellKey = "" Then Exit Sub
    If Not (Document.Cells(strCellKey).对象类型 = cprCTFixtext Or Document.Cells(strCellKey).对象类型 = cprCTText Or Document.Cells(strCellKey).对象类型 = cprCTTextElement) Then Exit Sub
    If Doc.Visible Then
'        If UCase(strType) <> "PASTE" And Doc.Selection.StartPos = Doc.Selection.EndPos Then Exit Sub
        Dim sType As String, lsSS As Long, lsSE As Long, lsES As Long, lsEE As Long, leKey As Long, bsInKeys As Boolean, bNeeded As Boolean
        bsInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lsSS, lsSE, lsES, lsEE, leKey, bNeeded)
        Dim leSS As Long, leSE As Long, leES As Long, leEE As Long, beInKeys As Boolean
        beInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.EndPos + 1, sType, leSS, leSE, leES, leEE, leKey, bNeeded)
    End If
    
    Select Case UCase(strType)
        Case "ALL"
            Select Case Document.Cells(strCellKey).对象类型
                Case cprCTText, cprCTFixtext
                    
                Case cprCTTextElement
                    If Doc.Visible Then
                        Call Doc.SelectAll
                    End If
            End Select
        Case "CUT"
            If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
                If Not AllowEdit(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) Then Exit Sub '可能受行列控限制
            End If
            Select Case Document.Cells(strCellKey).对象类型
                Case cprCTText, cprCTFixtext
                    If mblnEditing Then
                        SendKeys "^X"
                    Else
                        Call Clipboard.SetText(Document.Cells(strCellKey).内容文本)
                        Document.Cells(strCellKey).内容文本 = ""
                        F1Main.TextRC(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) = ""
                    End If
                Case cprCTTextElement
                    If Doc.Visible Then
                        If Doc.Selection.StartPos = Doc.Selection.EndPos Then Exit Sub
                        If bsInKeys And beInKeys Then '起始位和终止位都处在关键字之间
                            strTmp = Doc.Range(lsSS, leEE).Text
                            Doc.Range(lsSS, leEE).Text = ""
                        ElseIf bsInKeys Then          '起始位在关键字之间
                            strTmp = Doc.Range(lsSS, Doc.Selection.EndPos).Text
                            Doc.Range(lsSS, Doc.Selection.EndPos).Text = ""
                        ElseIf beInKeys Then          '终止位在关键字之间
                            strTmp = Doc.Range(Doc.Selection.StartPos, leEE).Text
                            Doc.Range(Doc.Selection.StartPos, leEE).Text = ""
                        Else
                            strTmp = Doc.Selection.Text
                            Doc.Range(Doc.Selection.StartPos, Doc.Selection.EndPos).Text = ""
                        End If
                        If strTmp = "" Then Exit Sub
                        strTmp = Doc.GetCleanTxt(strTmp)
                        Clipboard.SetText strTmp
                    Else
                        strTmp = Document.Cells(strCellKey).内容文本
                        If strTmp = "" Then Exit Sub
                        F1Main.TextRC(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) = ""
                        Document.Cells(strCellKey).内容文本 = "": Document.Cells(strCellKey).TextKey = "": Document.Cells(strCellKey).ElementKey = ""
                        Clipboard.SetText strTmp
                    End If
            End Select
        Case "COPY"
            Select Case Document.Cells(strCellKey).对象类型
                Case cprCTText, cprCTFixtext
                    If mblnEditing Then
                        SendKeys "^C"
                    Else
                        Call Clipboard.SetText(Document.Cells(strCellKey).内容文本)
                    End If
                Case cprCTTextElement
                    If Doc.Visible Then
                        If Doc.Selection.StartPos = Doc.Selection.EndPos Then Exit Sub
                        If bsInKeys And beInKeys Then '起始位和终止位都处在关键字之间
                            strTmp = Doc.Range(lsSS, leEE).Text
                        ElseIf bsInKeys Then          '起始位在关键字之间
                            strTmp = Doc.Range(lsSS, Doc.Selection.EndPos).Text
                        ElseIf beInKeys Then          '终止位在关键字之间
                            strTmp = Doc.Range(Doc.Selection.StartPos, leEE).Text
                        Else
                            strTmp = Doc.Selection.Text
                        End If
                        If strTmp = "" Then Exit Sub
                        strTmp = Doc.GetCleanTxt(strTmp)
                        Clipboard.SetText strTmp
                    Else
                        Call Clipboard.SetText(Document.Cells(strCellKey).内容文本)
                    End If
            End Select
        Case "PASTE"
            If editType = TabET_单病历编辑 Or editType = TabET_单病历审核 Then
                If Not AllowEdit(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) Then Exit Sub '可能受行列控限制
            End If
            Select Case Document.Cells(strCellKey).对象类型
                Case cprCTText, cprCTFixtext
                    If mblnEditing Then
                        SendKeys "^V"
                    Else
                        Document.Cells(strCellKey).内容文本 = Clipboard.GetText()
                        F1Main.TextRC(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) = Document.Cells(strCellKey).内容文本
                    End If
                Case cprCTTextElement
                    If Doc.Visible Then
                        strTmp = Clipboard.GetText
                        If bsInKeys And beInKeys Then '起始位和终止位都处在关键字之间
                            Doc.Range(leEE, leEE).Selected
                        ElseIf bsInKeys Then          '起始位在关键字之间
                            Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Selected
                        ElseIf beInKeys Then          '终止位在关键字之间
                            Doc.Range(leEE, leEE).Selected
                        Else
                            Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Selected
                        End If
                        If strTmp = "" Or strTmp = "GetText" Then Exit Sub
                        Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Font.Hidden = False
                        Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Font.Protected = False
                        Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Text = strTmp
                        Doc.Range(0, Len(Doc.Text)).Font.Name = SelCell.FontName '对文字样式赋值
                        Doc.Range(0, Len(Doc.Text)).Font.Size = SelCell.FontSize
                        Doc.Range(0, Len(Doc.Text)).Font.Bold = SelCell.FontBold
                        Doc.Range(0, Len(Doc.Text)).Font.Italic = SelCell.FontItalic
                        Doc.Range(0, Len(Doc.Text)).Font.Underline = SelCell.FontUnderline
                        Doc.Range(0, Len(Doc.Text)).Font.ForeColor = SelCell.FontColor
                        Doc.Range(0, Len(Doc.Text)).Font.Strikethrough = SelCell.FontStrikeout
                    End If
            End Select
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function AddSign(arrSQL As Variant, SignCellKey As String) As Boolean
Dim SignKey As String, SignTxt As String, oSign As cTabSign
    On Error GoTo errHand
1    SignKey = "": SignCellKey = ""
2    If InStr("6,7,8", SelCell.对象类型) > 0 Then '确定是否是签名位
3        SignKey = SelCell.SignKey
4        SignCellKey = SelCell.Key
5    Else
6        Dim l As Integer, lCount As Long
7        lCount = Document.Cells.Count
8        For l = 1 To lCount
9            If IIf(Document.Cells(l).Merge, InStr(Document.Cells(l).MergeRange, ";") > 0, True) Then
10                If InStr("6,7,8", Document.Cells(l).对象类型) > 0 Then
11                    If SignKey = "" Then
12                        SignKey = Document.Cells(l).SignKey
13                        SignCellKey = Document.Cells(l).Key
14                    Else
15                        SignKey = "": Exit For '当存在多个签名单元格时给于提示
16                    End If
17                End If
18            End If
19        Next
20    End If
    
21    If SignKey = "" Then
22        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提示：" & vbCrLf & "      请先选定需要签名的单元格！" & vbCrLf, True, 0
23        Exit Function
24    End If
25    If InStr("7,8", Document.Cells(SignCellKey).对象类型) > 0 And Document.Cells(SignCellKey).终止版 <> 0 Then
26        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "提示：" & vbCrLf & "      该单元格处于保护状态，不能签名！" & vbCrLf & "      请检查！", True, 0
27        Exit Function
28    End If
    
29    Set oSign = frmSign.ShowMe(SignKey, Me)  '签名窗体对签名元素赋值
30    If Not oSign Is Nothing Then
31        Set Document.Signs("K" & SignKey) = oSign
32    Else
33        Exit Function
34    End If

35    If Document.Cells(SignCellKey).终止版 = 0 Then
36        Document.Cells(SignCellKey).开始版 = 1
37        Document.Cells(SignCellKey).终止版 = IIf(editType = TabET_单病历审核, Document.EPRPatiRecInfo.最后版本 + 1, 1)
38    Else
39        Document.Cells(SignCellKey).ID = 0                          '同一签名位多次签名
40        Document.Cells(SignCellKey).对象序号 = Document.mMaxNo + 1
41        Document.Cells(SignCellKey).开始版 = Document.EPRPatiRecInfo.最后版本 + 1
42        Document.Cells(SignCellKey).终止版 = Document.Cells(SignCellKey).开始版
43        Document.mMaxNo = Document.mMaxNo + 1
44    End If
45    With Document.Signs("K" & SignKey)
46        SignTxt = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
47        SignTxt = SignTxt & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
48    End With
49    Dim lSignRow As Long, lSignCol As Long
50    lSignRow = Document.Cells(SignCellKey).Row: lSignCol = Document.Cells(SignCellKey).Col
51    F1Main.TextRC(lSignRow, lSignCol) = SignTxt
52    AddSign = True
53    Exit Function
errHand:
    Call MsgBox("AddSign错误行:" & Erl(), vbInformation, gstrSysName)
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function RollBack() As Boolean
Dim mfrmUntread As New frmUntread
    On Error GoTo errHand
1    If mfrmUntread.ShowMe(Me, mstrPrivs) Then
2        On Error Resume Next
3        Call mfrmParent.RefreshList
4        Call mfrmParent.Event_Saved(Document.EPRPatiRecInfo.ID) '诊疗单据需要，因为可能是非模态方式调用，不能用事件方式
5        Err.Clear
        
6        On Error GoTo errHand
7        Document.EPRPatiRecInfo.GetPatiRecordInfo Document.EPRPatiRecInfo.ID, mblnMoved '重新读取最后版
8        Call Me.ShowMe(mfrmParent, Document, mstrModelPrivate, mblnMoved, mblnCanPrint)
9    End If
    Exit Function
errHand:
    Call MsgBox("RollBack错误行:" & Erl(), vbInformation, gstrSysName)
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub timeTmp_Timer()
    ValiCellDate False
End Sub
Private Sub txtSum_KeyPress(KeyAscii As Integer)
    If InStr("1234567890,;" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsHistory_DblClick()
    On Error GoTo errHand
    If Val(vsHistory.TextMatrix(vsHistory.Row, vsHistory.Cols - 1)) = 0 Then Exit Sub
    Me.Enabled = False
    zlCommFun.ShowFlash "请稍等－－－－正在打开文件", Me
    Dim DocTmp As New cTableEPR
    DocTmp.InitOpenEPR mfrmParent, TabEm_修改, TabET_单病历审核, Document.EPRPatiRecInfo.ID, False, 2, Document.EPRPatiRecInfo.病人来源, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.婴儿, UserInfo.部门ID, Document.EPRPatiRecInfo.医嘱id, mstrModelPrivate, mblnMoved, mblnCanPrint, gbytEsign
    DocTmp.EPRPatiRecInfo.最后版本 = Val(vsHistory.TextMatrix(vsHistory.Row, vsHistory.Cols - 1))
    DocTmp.frmEditor.ShowMe mfrmParent, DocTmp, mstrModelPrivate, mblnMoved, mblnCanPrint
    Me.Enabled = True: zlCommFun.StopFlash
    Exit Sub
errHand:
    Me.Enabled = True: zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Sub zlRefreshPacsPic()
    mfrmPacsPic.zlRefresh Document.EPRPatiRecInfo.医嘱id, Document.EPRFileInfo.lngModule
End Sub
Private Sub ExeUndo()
'执行撤销操作
'1 对类对像赋值 2 对表格显示内容刷新显示或图片刷新显示 3 删除Undo类中最后一个成员
Dim strShow As String, lRow As Long, lCol As Long
    On Error GoTo errHand
    If Undo.Count < 1 Then Exit Sub
    With Undo(Undo.Count)
        Select Case .CT
            Case cprCTFixtext, cprCTText
                Document.Cells(.Key).内容文本 = .CTxt
                F1Main.TextRC(.Row, .Col) = .CTxt
                
                If InStr(Document.Cells(.Key).对象属性, ",") > 0 And InStr(Document.Cells(.Key).对象属性, ";") = 0 Then '合计单元格的源单元格
                    lRow = Split(Document.Cells(.Key).对象属性, ",")(0): lCol = Split(Document.Cells(.Key).对象属性, ",")(1) '合计单元格的行列
                    Call CalcSumRange(lRow, lCol)
                End If
            Case cprCTElement
                Document.Cells(.Key).ElementKey = .Ekey: lRow = .Row: lCol = .Col: strShow = .CTxt
                With Document.Cells(.Key)
                    .内容文本 = strShow: strShow = ""
                    If .内容文本 <> "" Then
                        strShow = .内容文本
                    Else
                        If editType = TabET_病历文件定义 Or editType = TabET_全文示范编辑 Then
                            If .ElementKey <> "" Then
                                If Document.Elements("K" & .ElementKey).输入形态 = 1 Then
                                    strShow = Document.Elements("K" & .ElementKey).内容文本
                                Else
                                    strShow = "[" & Document.Elements("K" & .ElementKey).要素名称 & "]" & Document.Elements("K" & .ElementKey).要素单位
                                End If
                            Else
                                strShow = ""
                            End If
                        Else
                            If Document.Elements("K" & .ElementKey).替换域 = 1 Then '自动替换要素
                                If Document.Elements("K" & .ElementKey).自动转文本 Then '没取到值，是否自动转换成文本(空)
                                    strShow = ""
                                Else
                                    strShow = "[" & Document.Elements("K" & .ElementKey).要素名称 & "]" & Document.Elements("K" & .ElementKey).要素单位
                                End If
                            Else
                                If Document.Elements("K" & .ElementKey).输入形态 = 1 Then '输入形态=展开
                                    strShow = Document.Elements("K" & .ElementKey).内容文本 & Document.Elements("K" & .ElementKey).要素单位
                                Else
                                    strShow = "[" & Document.Elements("K" & .ElementKey).要素名称 & "]" & Document.Elements("K" & .ElementKey).要素单位
                                End If
                            End If
                        End If
                    End If
                    F1Main.TextRC(lRow, lCol) = strShow
                End With
            Case cprCTTextElement
                Document.Cells(.Key).ElementKey = .Ekey
                Document.Cells(.Key).TextKey = .Tkey
                GetTextELement .Key '显示混合区域
            Case cprCTPicture, cprCTReportPic
                Document.Cells(.Key).PictureKey = .PKey
                If Len(.PKey) <> 0 Then
                    Set Document.Pictures("K" & .PKey).OrigPic = .OrigPic
                End If
                Document.Cells(.Key).PicMarkKey = .PmKey
                PaintPictureOnTable .Key
        End Select
    End With
    Undo.Remove (Undo.Count)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'   用途：  动态更新工具栏“颜色”图标。
'################################################################################################################
Private Sub SetColorIcon(ID As Long, Color As OLE_COLOR)
    Dim ctlPictureBox As VB.PictureBox
    Set ctlPictureBox = Controls.Add("VB.PictureBox", "ctlPictureBox1")
    Dim ListImage As ListImage
    Set ListImage = imgColor.ListImages("FORECOLOR")

    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = imgColor.MaskColor

    ctlPictureBox.Picture = ListImage.ExtractIcon

    If Color = vbWhite Then Color = RGB(254, 254, 254)
    ctlPictureBox.Line (1, ctlPictureBox.Height * 0.6)-(ctlPictureBox.Width, ctlPictureBox.Height), Color, BF
    ctlPictureBox.Refresh

    'Replace icon
    imgColor.ListImages.Remove imgColor.ListImages("FORECOLOR").Index
    imgColor.ListImages.Add 1, "FORECOLOR", ctlPictureBox.Image

    'OK Now replace Tag property
    imgColor.ListImages(1).Tag = ID

    cbsMain.AddImageList imgColor
    cbsMain.RecalcLayout

    Me.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
End Sub

Private Function RelateFeedback(ByVal isRelated As Boolean) As Boolean
'功能：传染病报告卡，关联阳性结果反馈单，或者取消关联
'参数：isRelated  true-关联；false-取消关联
    Dim strSQL As String
    Dim rsDisease As ADODB.Recordset
    Dim strIDs As String
    Dim arrayID() As String
    Dim i As Long
    Dim objDisease As Object
  
On Error GoTo errHand
    If Me.Document.EPRPatiRecInfo.病历种类 <> Tab诊断文书 Then Exit Function
  
    If isRelated Then   '关联
        If Me.Document.EPRPatiRecInfo.病人来源 = TabPF_门诊 Then
            strSQL = "select rowNum as NO,a.ID,c.名称 as 科室, a.登记时间 from  疾病阳性记录 A ,病人挂号记录 B ,部门表 C where A.文件ID is NULL  and A.挂号单 = B.NO and A.病人ID = B.病人ID and A.记录状态 <> 3 and A.登记科室ID = C.ID  and A.病人ID = [1] and B.ID = [2]"
        ElseIf Me.Document.EPRPatiRecInfo.病人来源 = TabPF_住院 Then
            strSQL = "select rowNum as NO,a.ID ,c.名称 as 科室,a.登记时间 from  疾病阳性记录 A ,部门表 C  where A.文件ID is NULL  and A.记录状态 <> 3  and A.登记科室ID = C.ID and A.病人ID = [1] and A.主页ID = [2] "
        End If
        Set rsDisease = zlDatabase.OpenSQLRecord(strSQL, "查询该报告对应的阳性结果反馈单", Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID)
        If rsDisease.RecordCount = 1 Then
            strSQL = "Zl_疾病阳性检测记录_Update(2," & rsDisease!ID & "," & Me.Document.EPRPatiRecInfo.ID & ",NULL,NULL,NULL,NULL)"
            Call zlDatabase.ExecuteProcedure(strSQL, "关联疾病报告和阳性结果反馈单")
        ElseIf rsDisease.RecordCount > 1 Then
            Set objDisease = CreateObject("zl9Disease.clsDisease")
            If objDisease Is Nothing Then Exit Function
            If objDisease.GetFeedbackList().ShowMe(Me, rsDisease, strIDs) Then
                If strIDs <> "" Then
                    arrayID = Split(strIDs, ",")
                    For i = LBound(arrayID) To UBound(arrayID)
                        If Val(arrayID(i)) <> 0 Then
                            strSQL = "Zl_疾病阳性检测记录_Update(2," & arrayID(i) & "," & Me.Document.EPRPatiRecInfo.ID & ",NULL,NULL,NULL,NULL)"
                            Call zlDatabase.ExecuteProcedure(strSQL, "关联疾病报告和阳性结果反馈单")
                        End If
                    Next
                End If
            End If
        End If
    Else '取消关联
        strSQL = "Zl_疾病阳性检测记录_Update(3, NULL " & "," & Me.Document.EPRPatiRecInfo.ID & ",NULL,NULL,NULL,NULL)"
        Call zlDatabase.ExecuteProcedure(strSQL, "取消疾病报告和阳性结果反馈单的关联")
    End If
    
    RelateFeedback = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

