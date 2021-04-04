VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPathItemEditOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "项目设置"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   Icon            =   "frmPathItemEditOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraImportRef 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      TabIndex        =   33
      Top             =   4680
      Width           =   12015
      Begin RichTextLib.RichTextBox rtfImportRef 
         Height          =   1600
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   2831
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmPathItemEditOut.frx":058A
      End
      Begin VB.Label lblImportRef 
         AutoSize        =   -1  'True
         Caption         =   "未成功导入的医嘱内容"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   60
         Width           =   1800
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12015
      TabIndex        =   5
      Top             =   0
      Width           =   12015
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径项目"
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
         Left            =   1095
         TabIndex        =   7
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "    设置门诊临床路径表中一个阶段中的项目信息，包括对应的医嘱、病历等，并可以指定项目的可选执行结果。"
         Height          =   180
         Left            =   1095
         TabIndex        =   6
         Top             =   480
         Width           =   9000
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathItemEditOut.frx":0627
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   12840
         Y1              =   825
         Y2              =   825
      End
   End
   Begin VB.Frame fraContent 
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   12255
      Begin VB.Frame fraERPType 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         TabIndex        =   36
         Top             =   383
         Width           =   1575
         Begin VB.OptionButton optEPRType 
            Caption         =   "新版"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   38
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optEPRType 
            Caption         =   "老版"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame fraSend 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   750
         Width           =   11775
         Begin VB.OptionButton optExecute 
            Caption         =   "无须生成(用于路径表中的说明文字等)"
            Height          =   180
            Index           =   0
            Left            =   3600
            TabIndex        =   27
            Top             =   30
            Width           =   3420
         End
         Begin VB.OptionButton optExecute 
            Caption         =   "必须生成"
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   26
            Top             =   30
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optExecute 
            Caption         =   "必要时生成"
            Height          =   180
            Index           =   3
            Left            =   2145
            TabIndex        =   25
            Top             =   30
            Width           =   1200
         End
         Begin VB.Label lblSendKind 
            Caption         =   "生成方式"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   -7
            Width           =   855
         End
      End
      Begin VB.Frame fraKind 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   960
         TabIndex        =   20
         Top             =   413
         Width           =   2895
         Begin VB.OptionButton optType 
            Caption         =   "其他类"
            Height          =   180
            Index           =   2
            Left            =   1905
            TabIndex        =   23
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton optType 
            Caption         =   "病历类"
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   22
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton optType 
            Caption         =   "医嘱类"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   960
         MaxLength       =   1000
         TabIndex        =   16
         Top             =   30
         Width           =   9375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11400
         MouseIcon       =   "frmPathItemEditOut.frx":2169
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblEPR 
         Caption         =   "病历版本"
         Height          =   180
         Left            =   4440
         TabIndex        =   39
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目名称"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblIcon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目图标"
         Height          =   180
         Left            =   10605
         TabIndex        =   18
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目类型"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12015
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8310
      Width           =   12015
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10680
         TabIndex        =   30
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9480
         TabIndex        =   29
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   0
         X2              =   12720
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraExecute 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   0
      TabIndex        =   11
      Top             =   6720
      Width           =   12015
      Begin VSFlex8Ctl.VSFlexGrid vsResult 
         Height          =   1395
         Left            =   5205
         TabIndex        =   4
         Top             =   120
         Width           =   6615
         _cx             =   11668
         _cy             =   2461
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathItemEditOut.frx":22BB
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         Begin MSComctlLib.ImageList imgNature 
            Left            =   540
            Top             =   390
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   6
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":2339
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":28D3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":2E6D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":3407
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":39A1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":3F3B
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   12840
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblExePrompt 
         Caption         =   $"frmPathItemEditOut.frx":44D5
         Height          =   1455
         Left            =   960
         TabIndex        =   31
         Top             =   120
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmPathItemEditOut.frx":4577
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblResult 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行结果"
         Height          =   180
         Left            =   4380
         TabIndex        =   12
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   555
      Top             =   435
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraAdvice 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   12015
      Begin zlCISPath.UCAdviceList UcAdvice 
         Height          =   2415
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   4128
      End
      Begin VB.OptionButton optSend 
         Caption         =   "选择使用"
         Height          =   255
         Index           =   1
         Left            =   10695
         TabIndex        =   3
         Top             =   83
         Width           =   1095
      End
      Begin VB.OptionButton optSend 
         Caption         =   "全部使用"
         Height          =   255
         Index           =   0
         Left            =   9510
         TabIndex        =   2
         Top             =   83
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdvice 
         Caption         =   "医嘱编辑(&E)"
         Height          =   350
         Left            =   90
         TabIndex        =   1
         Top             =   60
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   4
         X1              =   0
         X2              =   12885
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   5
         X1              =   0
         X2              =   12885
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生成路径项目时，以下医嘱"
         Height          =   180
         Left            =   7320
         TabIndex        =   14
         Top             =   120
         Width           =   2160
      End
   End
   Begin VB.Frame fraEPR 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   12015
      Begin VSFlex8Ctl.VSFlexGrid vsEPR 
         Height          =   4155
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   11760
         _cx             =   20743
         _cy             =   7329
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathItemEditOut.frx":4D75
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
   End
   Begin XtremeCommandBars.CommandBars cbsIcon 
      Left            =   90
      Top             =   660
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathItemEditOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CheckDataValid(PathItem As TYPE_PATH_ITEM, Cancel As Boolean)

Private mrsResult       As ADODB.Recordset          '执行结果集
Private mrsNature       As ADODB.Recordset          '结果性质集
Private mrsAdvice       As ADODB.Recordset
Private mvPreItem       As TYPE_PATH_ITEM
Private mvBakItem       As TYPE_PATH_ITEM
Private mvItem          As TYPE_PATH_ITEM
Private mlng路径ID      As Long
Private mlngItemID      As Long
Private mblnAdjust      As Boolean                  '是否微调模式
Private mblnReadOnly    As Boolean                  '是否只读查看模式
Private mblnUseExecute  As Boolean                  '是否启用执行环节
Private mblnReturn      As Boolean
Private mblnChange      As Boolean
Private mblnOK          As Boolean
Private mstrPrivs       As String                   '模块权限

Private Enum CONST_COL_执行结果
    col执行图标 = 0
    col执行结果 = 1
    col结果性质 = 2
    col缺省结果 = 3
End Enum

Public Sub ShowView(frmParent As Object, ByVal lngItemID As Long)
'功能：查看项目
    mlngItemID = lngItemID
    mblnReadOnly = True

    Me.Show 1, frmParent
End Sub

Public Function ShowEdit(frmParent As Object, rsAdvice As ADODB.Recordset, vItem As TYPE_PATH_ITEM, vPreItem As TYPE_PATH_ITEM, ByVal blnAdjust As Boolean, Optional ByVal lng路径ID As Long, Optional ByVal strPrivs As String) As Boolean
'功能：设置当前选择项目的详细内容
'参数：rsAdvice=(入/出)已经定义的当前路径表中的医嘱记录全集
'      vItem=(入/出)主要是修改时当前项目的内容
'      mvPreItem = (入)前一个时间阶段中相同的项目，用于设置时参考
'      blnAdjust = 是否进行微调模式
'      lng路径ID = 设计路径表时编辑的路径ID
    Set mrsAdvice = rsAdvice
    mvItem = vItem
    mvBakItem = vItem
    mvPreItem = vPreItem
    mblnAdjust = blnAdjust
    mlng路径ID = lng路径ID
    mstrPrivs = strPrivs

    Me.Show 1, frmParent

    If mblnOK Then
        vItem = mvItem
    End If
    ShowEdit = mblnOK
End Function

Private Sub cbsIcon_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID = -1 Then
        picIcon.Cls
        mvItem.图标ID = 0
    Else
        mvItem.图标ID = Control.ID
        Call DrawPicture(GetPathIcon(mvItem.图标ID))
    End If
    mblnChange = True
End Sub

Private Sub DrawPicture(objPic As StdPicture)
    Dim X As Long, Y As Long, W As Long, H As Long

    W = picIcon.ScaleX(objPic.Width, vbHimetric, vbTwips)
    H = picIcon.ScaleY(objPic.Height, vbHimetric, vbTwips)

    X = (picIcon.ScaleWidth - W) / 2
    Y = (picIcon.ScaleHeight - H) / 2

    picIcon.PaintPicture objPic, X, Y, W, H
End Sub

Private Sub cmdAdvice_Click()
'功能：编辑项目所对应的医嘱
    Dim rsScheme As ADODB.Recordset
    Dim strFilter As String, lng序号 As Long
    Dim colAdviceID As New Collection
    Dim str医嘱IDs As String, lng医嘱ID As Long
    Dim strItem As String, i As Long
    Dim strSql As String, rsTmp As Recordset
    Dim str使用科室 As String
    Dim strSelectedID As String
    Dim strSelectedIDAlt As String
    Dim blnUpdate As Boolean

    If mvItem.医嘱IDs <> "" Then
        If optSend(1).Value Then    '界面上勾选后还没有更新"mrsAdvice!是否缺省"，保存时才更新
            strSelectedID = "," & UCAdvice.GetAdviceIDSelected & ","
            strSelectedIDAlt = "," & UCAdvice.GetAdviceIDSelected(1) & ","
        End If

        Call InitSchemeRecordset(rsScheme)

        strFilter = "": str医嘱IDs = ""
        If mvItem.Edit = 2 And mvItem.待审核医嘱IDs <> "" Then
           '当前待审核医嘱还未保存
            str医嘱IDs = mvItem.待审核医嘱IDs
        Else
            str医嘱IDs = mvItem.医嘱IDs
        End If
        For i = 0 To UBound(Split(str医嘱IDs, ","))
            strFilter = strFilter & " Or ID=" & Split(str医嘱IDs, ",")(i)
        Next
        mrsAdvice.Filter = Mid(strFilter, 5)
        Do While Not mrsAdvice.EOF
            rsScheme.AddNew
            rsScheme!序号 = mrsAdvice!ID
            rsScheme!相关序号 = mrsAdvice!相关id
            rsScheme!期效 = mrsAdvice!期效
            rsScheme!诊疗项目ID = mrsAdvice!诊疗项目ID
            rsScheme!收费细目ID = mrsAdvice!收费细目ID
            rsScheme!医嘱内容 = mrsAdvice!医嘱内容
            rsScheme!单次用量 = mrsAdvice!单次用量
            rsScheme!总给予量 = mrsAdvice!总给予量
            rsScheme!医生嘱托 = mrsAdvice!医生嘱托
            rsScheme!执行频次 = mrsAdvice!执行频次
            rsScheme!频率次数 = mrsAdvice!频率次数
            rsScheme!频率间隔 = mrsAdvice!频率间隔
            rsScheme!间隔单位 = mrsAdvice!间隔单位
            rsScheme!时间方案 = mrsAdvice!时间方案
            rsScheme!执行科室ID = mrsAdvice!执行科室ID
            rsScheme!执行性质 = mrsAdvice!执行性质
            rsScheme!标本部位 = mrsAdvice!标本部位
            rsScheme!检查方法 = mrsAdvice!检查方法
            rsScheme!是否缺省 = IIf(InStr(strSelectedID, "," & mrsAdvice!ID & ",") > 0, 1, 0)
            rsScheme!是否备选 = IIf(InStr(strSelectedIDAlt, "," & mrsAdvice!ID & ",") > 0, 1, 0)
            rsScheme!配方ID = mrsAdvice!配方ID
            rsScheme!组合项目ID = mrsAdvice!组合项目ID
            rsScheme!执行标记 = mrsAdvice!执行标记
            rsScheme.Update
            mrsAdvice.MoveNext
        Loop
        mrsAdvice.Filter = ""
    End If

    On Error GoTo errH
    strSql = "Select 科室ID From 门诊路径科室 Where 路径ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
    Do Until rsTmp.EOF
        str使用科室 = str使用科室 & "," & rsTmp!科室ID
        rsTmp.MoveNext
    Loop
    On Error GoTo 0

    Set rsScheme = gobjKernel.ShowSchemeEdit(Me, 1, rsScheme, False, optSend(1).Value, Mid(str使用科室, 2), 1)
    '路径医嘱调整
    If Not rsScheme Is Nothing Then
        'blnUpdate=True 将审核且未停用的路径医嘱变动记录保存到【门诊路径医嘱变动】中，待审核后再更新表单路径医嘱。
        blnUpdate = InStr(mstrPrivs, "门诊路径医嘱调整") = 0 And mvItem.原医嘱IDs <> ""
        '先产生新的医嘱ID
        str医嘱IDs = ""
        Do While Not rsScheme.EOF
            lng医嘱ID = zlDatabase.GetNextId("门诊路径医嘱内容")
            colAdviceID.Add lng医嘱ID, "_" & rsScheme!序号
            str医嘱IDs = str医嘱IDs & "," & lng医嘱ID
            rsScheme.MoveNext
        Loop

        If Not blnUpdate Then
            str医嘱IDs = Mid(str医嘱IDs, 2)
            mvItem.医嘱IDs = str医嘱IDs
        Else
            str医嘱IDs = Mid(str医嘱IDs, 2)
            mvItem.待审核医嘱IDs = str医嘱IDs
        End If
        '加入新的医嘱
        rsScheme.MoveFirst: lng序号 = 1
        Do While Not rsScheme.EOF
            lng医嘱ID = colAdviceID("_" & rsScheme!序号)
            mrsAdvice.AddNew

            mrsAdvice!ID = lng医嘱ID
            If Not IsNull(rsScheme!相关序号) Then
                mrsAdvice!相关id = colAdviceID("_" & rsScheme!相关序号)
            End If
            mrsAdvice!序号 = lng序号
            mrsAdvice!期效 = rsScheme!期效
            mrsAdvice!诊疗项目ID = rsScheme!诊疗项目ID
            mrsAdvice!收费细目ID = rsScheme!收费细目ID
            If IsNull(rsScheme!诊疗项目ID) Then
                mrsAdvice!医嘱内容 = rsScheme!医嘱内容 '自由录入医嘱才保存
            End If
            mrsAdvice!单次用量 = rsScheme!单次用量
            mrsAdvice!总给予量 = rsScheme!总给予量
            mrsAdvice!医生嘱托 = rsScheme!医生嘱托
            mrsAdvice!执行频次 = rsScheme!执行频次
            mrsAdvice!频率次数 = rsScheme!频率次数
            mrsAdvice!频率间隔 = rsScheme!频率间隔
            mrsAdvice!间隔单位 = rsScheme!间隔单位
            mrsAdvice!时间方案 = rsScheme!时间方案
            mrsAdvice!执行科室ID = rsScheme!执行科室ID
            mrsAdvice!执行性质 = rsScheme!执行性质
            mrsAdvice!标本部位 = rsScheme!标本部位
            mrsAdvice!检查方法 = rsScheme!检查方法
            mrsAdvice!是否缺省 = rsScheme!是否缺省
            mrsAdvice!是否备选 = rsScheme!是否备选
            mrsAdvice!配方ID = rsScheme!配方ID
            mrsAdvice!组合项目ID = rsScheme!组合项目ID
            mrsAdvice!执行标记 = rsScheme!执行标记
            mrsAdvice!待审核 = IIf(blnUpdate, 1, 0)
            mrsAdvice!项目ID = IIf(blnUpdate, mvItem.ID, 0)
            mrsAdvice.Update

            lng序号 = lng序号 + 1
            rsScheme.MoveNext
        Loop
        '刷新显示
        Call ShowAdvice(str医嘱IDs)

        '缺省项目内容
        If txtItem.Text = "" Then
            txtItem.Text = UCAdvice.GetAdviceTitle
        End If
        mblnChange = True
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean
    Dim strFilter As String, strSelectedID As String
    Dim i As Long
    Dim strSelectedAltID As String
    Dim blnIsAllSelect As Boolean
    Dim strTmp As String

    If mblnReadOnly Then
        mblnOK = True: Unload Me: Exit Sub
    End If

    '数据检查
    If Trim(txtItem.Text) = "" Then
        MsgBox "请输入路径项目的内容。", vbInformation, gstrSysName
        txtItem.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txtItem.Text) > txtItem.MaxLength Then
        MsgBox "项目内容中最多允许 " & txtItem.MaxLength \ 2 & " 个汉字或者 " & txtItem.MaxLength & "。", vbInformation, gstrSysName
        txtItem.SetFocus: Exit Sub
    End If

    '检查医嘱
    If optType(0).Value Then
        If mvItem.医嘱IDs = "" Then
            MsgBox "没有定义当前项目所对应的医嘱内容。", vbInformation, gstrSysName
            If cmdAdvice.Enabled Then cmdAdvice.SetFocus
            Exit Sub
        End If
        strFilter = ""
        For i = 0 To UBound(Split(mvItem.医嘱IDs, ","))
            strFilter = strFilter & " Or ID=" & Split(mvItem.医嘱IDs, ",")(i)
        Next
        strSelectedID = "," & UCAdvice.GetAdviceIDSelected & ","
        strSelectedAltID = "," & UCAdvice.GetAdviceIDSelected(1, blnIsAllSelect) & ","

        mrsAdvice.Filter = Mid(strFilter, 5)
        strFilter = ""
        If blnIsAllSelect Then
            '至少要有一个不是备选
            MsgBox "一个路径项目至少有一个不是备选医嘱。", vbInformation, Me.Caption
            Exit Sub
        End If
        Do While Not mrsAdvice.EOF
            If InStr(strSelectedID, "," & mrsAdvice!ID & ",") > 0 Then
                mrsAdvice!是否缺省 = 1
            Else
                mrsAdvice!是否缺省 = 0
            End If
            If InStr(strSelectedAltID, "," & mrsAdvice!ID & ",") > 0 Then
                mrsAdvice!是否备选 = 1
            Else
                mrsAdvice!是否备选 = 0
            End If
            mrsAdvice.Update
            mrsAdvice.MoveNext
        Loop
        mrsAdvice.Filter = ""
    Else
        mvItem.医嘱IDs = ""
    End If

    '检查病历
    If optType(1).Value Then
        With vsEPR
            strFilter = ""
            mvItem.病历IDs = "": mvItem.新版病历IDs = "": mvItem.病历详情 = ""
            For i = .FixedRows To .Rows - 1
                If .RowData(i) <> "" Then
                    strTmp = .RowData(i) '格式：(NEW/OLD)|ID
                    If Split(strTmp, "|")(0) = "OLD" Then
                        mvItem.病历IDs = mvItem.病历IDs & "," & Split(strTmp, "|")(1)
                    Else
                        mvItem.新版病历IDs = mvItem.新版病历IDs & "," & Split(strTmp, "|")(1)
                    End If
                    '病历详情:文件ID,原型ID,名称,序号
                    mvItem.病历详情 = mvItem.病历详情 & ";" & IIf(Split(strTmp, "|")(0) = "OLD", Split(strTmp, "|")(1) & ",", "," & Split(strTmp, "|")(1)) & "," & Trim(.TextMatrix(i, 0)) & "," & i
                    If InStr(strFilter & ",", "," & .TextMatrix(i, 0) & ",") = 0 Then
                        strFilter = strFilter & "," & .TextMatrix(i, 0)
                    Else
                        MsgBox "指定了重复的病历文件""" & .TextMatrix(i, 0) & """。", vbInformation, gstrSysName
                        .Row = i: Call .ShowCell(.Row, .Col)
                        .SetFocus: Exit Sub
                    End If
                End If
            Next
            mvItem.病历IDs = Mid(mvItem.病历IDs, 2)
            mvItem.病历详情 = Mid(mvItem.病历详情, 2)
            mvItem.新版病历IDs = Mid(mvItem.新版病历IDs, 2)
            If mvItem.病历IDs = "" And mvItem.新版病历IDs = "" Then
                MsgBox "请指定项目所对应的病历文件。", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
        End With
    Else
        mvItem.病历IDs = "": mvItem.新版病历IDs = ""
    End If

    '检查结果
    If Not optExecute(0).Value And fraExecute.Visible Then
        With vsResult
            strFilter = ""
            mvItem.项目结果 = ""
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col执行结果)) <> "" Then
                    If InStr(strFilter & ",", "," & .TextMatrix(i, col执行结果) & ",") = 0 _
                        And InStr(strFilter, "," & .TextMatrix(i, col执行结果) & "|") = 0 Then
                        strFilter = strFilter & "," & .TextMatrix(i, col执行结果)
                        If .TextMatrix(i, col结果性质) <> "" Then
                            mrsNature.Filter = "名称='" & .TextMatrix(i, col结果性质) & "'"
                            strFilter = strFilter & "|" & mrsNature!编码
                        End If
                    Else
                        MsgBox "指定了重复的执行结果""" & .TextMatrix(i, col执行结果) & """。", vbInformation, gstrSysName
                        .Row = i: Call .ShowCell(.Row, .Col)
                        .SetFocus: Exit Sub
                    End If

                    '缺省结果
                    If Val(.TextMatrix(i, col缺省结果)) <> 0 Then
                        mvItem.项目结果 = .TextMatrix(i, col执行结果)
                    End If
                End If
            Next
            strFilter = Mid(strFilter, 2)
            If strFilter = "" Then
                MsgBox "请指定项目所对应的执行结果。", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            If mvItem.项目结果 = "" Then
                MsgBox "请指定项目所对应的缺省结果。", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            mvItem.项目结果 = strFilter & vbTab & mvItem.项目结果
        End With
    Else
        mvItem.项目结果 = ""
    End If

    '其他数据收集
    If mvItem.医嘱IDs <> "" Then
        mvItem.内容要求 = IIf(optSend(0).Value, 0, 1) '0-全部生成，1-选择生成
    Else
        mvItem.内容要求 = 0
    End If
    mvItem.项目内容 = txtItem.Text
    For i = 0 To optExecute.UBound
        If i <> 2 Then
            If optExecute(i).Value Then mvItem.执行方式 = i: Exit For
        End If
    Next
    RaiseEvent CheckDataValid(mvItem, blnCancel)
    If blnCancel Then Exit Sub

    If mblnChange Then
        mvItem.导入结果 = 1
    End If

    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If TypeName(ActiveControl) <> "VSFlexGrid" Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strDefault As String, strSql As String
    Dim arrResult As Variant, i As Long
    Dim objControl As Object
    Dim strTmp As String

    On Error GoTo errH

    mblnOK = False
    fraEPR.BackColor = Me.BackColor
    fraAdvice.BackColor = Me.BackColor
    fraExecute.BackColor = Me.BackColor

    mblnUseExecute = Val(zlDatabase.GetPara("是否启用路径执行环节", glngSys, P门诊路径应用, 1))
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsIcon.VisualTheme = xtpThemeOffice2003
    cbsIcon.ActiveMenuBar.Visible = False
    With Me.cbsIcon.Options
        .ToolBarAccelTips = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With

    '只读查看模式时，获取一些数据
    If mblnReadOnly And mlngItemID <> 0 Then
        strSql = "Select 项目内容,执行方式,项目结果,图标ID,内容要求 From 门诊路径项目 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngItemID)

        '项目基本信息
        mvItem.ID = mlngItemID
        mvItem.项目内容 = rsTmp!项目内容
        mvItem.项目结果 = NVL(rsTmp!项目结果)
        mvItem.执行方式 = NVL(rsTmp!执行方式, 0)
        mvItem.图标ID = NVL(rsTmp!图标ID, 0)
        mvItem.内容要求 = Val("" & rsTmp!内容要求)

        '关联医嘱信息
        strSql = "Select 医嘱内容ID From 门诊路径医嘱 Where 路径项目ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngItemID)
        strSql = ""
        Do While Not rsTmp.EOF
            strSql = strSql & "," & rsTmp!医嘱内容ID
            rsTmp.MoveNext
        Loop
        mvItem.医嘱IDs = Mid(strSql, 2)

        '关联病历信息
        strSql = "Select 文件ID,原型ID From 门诊路径病历 Where 项目ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngItemID)
        strSql = ""
        Do While Not rsTmp.EOF
            If rsTmp!文件ID & "" <> "" Then
                strSql = strSql & "," & rsTmp!文件ID
            Else
                strTmp = strTmp & "," & rsTmp!原型ID
            End If
            rsTmp.MoveNext
        Loop
        mvItem.病历IDs = Mid(strSql, 2)
        mvItem.新版病历IDs = Mid(strTmp, 2)
        '医嘱记录集
        If mvItem.医嘱IDs <> "" Then
            strSql = " Select Distinct A.ID,A.相关ID,A.序号,A.期效,A.诊疗项目ID,A.收费细目ID," & _
                     " A.医嘱内容,A.单次用量,A.总给予量,A.标本部位,A.检查方法,A.医生嘱托," & _
                     " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.执行性质,A.执行标记,A.执行科室ID,A.时间方案,A.是否缺省,A.是否备选,A.配方ID,A.组合项目ID" & _
                     " From 门诊路径医嘱内容 A,门诊路径医嘱 B" & _
                     " Where A.ID=B.医嘱内容ID And B.路径项目ID=[1]" & _
                     " Order by A.序号,A.ID"
            Set mrsAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngItemID)
        End If
    End If

    '读取结果性质集
    strSql = "Select 编码,名称 From 路径结果性质 Order by 编码"
    Set mrsNature = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsNature, strSql, Me.Caption)
    strSql = ""
    Do While Not mrsNature.EOF
        strSql = strSql & "|" & mrsNature!编码 & "-" & mrsNature!名称
        mrsNature.MoveNext
    Loop
    vsResult.ColData(col结果性质) = Mid(strSql, 2)

    '读取可用结果集
    strSql = "Select A.编码,A.名称,Nvl(基本,0) as 基本,B.名称 as 性质" & _
        " From 路径常见结果 A,路径结果性质 B" & _
        " Where A.末级=1 And Nvl(A.性质,0)=B.编码(+)" & _
        " Order by A.编码"
    Set mrsResult = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsResult, strSql, Me.Caption)

    '编辑数据时的一些处理
    If mvItem.ID <> 0 Then
        txtItem.Text = mvItem.项目内容
        If mvItem.图标ID <> 0 Then
            Call DrawPicture(GetPathIcon(mvItem.图标ID))
        End If
        '----
        If mvItem.医嘱IDs <> "" Then
            '显示医嘱
            optType(0).Value = True
            Call ShowAdvice(mvItem.医嘱IDs)
            
            optSend(0).Value = (mvItem.内容要求 = 0)
            optSend(1).Value = Not optSend(0).Value

            Call UCAdvice.Set选择列的可见性(optSend(0).Value)
        ElseIf mvItem.病历IDs <> "" Or mvItem.新版病历IDs <> "" Then
            '显示病历
            optType(1).Value = True
            If mvItem.Edit = 0 Then
                If mvItem.新版病历IDs = "" Then '老版
                    strSql = "Select /*+ Rule*/ ID as 文件ID,名称,1 as 版本 From 病历文件列表" & _
                        " Where ID IN(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
                        " Order by 编号"
                ElseIf mvItem.病历IDs <> "" Then '新版+老版
                    strSql = "Select A.文件ID,A.原型ID,Nvl(a.名称, b.名称) as 名称,decode(文件ID,NULL,2,1) as 版本 From 临床路径病历 A, 病历文件列表 B Where a.项目id = [3] And a.文件id = b.Id(+)" & vbNewLine & _
                        "order by a.序号"
                Else '新版
                    strSql = "Select T.原型ID,T.名称,2 as 版本 From 门诊路径病历 T Where t.项目id = [3] And t.文件id Is Null And t.原型ID IN(Select Column_Value From Table(Cast(f_STR2list([2]) As zlTools.t_Strlist)))" & _
                        " Order by 序号"
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mvItem.病历IDs, mvItem.新版病历IDs, mvItem.ID)
            Else
                Set rsTmp = FuncGetEMRInfo(mvItem.病历详情)
            End If
            With vsEPR
                .Rows = .FixedRows
                Do While Not rsTmp.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsTmp!名称
                    .Cell(flexcpData, .Rows - 1, 0) = .TextMatrix(.Rows - 1, 0)
                    If rsTmp!版本 & "" = "1" Then
                        .RowData(.Rows - 1) = "OLD|" & rsTmp!文件ID
                    Else
                        .RowData(.Rows - 1) = "NEW|" & rsTmp!原型ID
                    End If
                    rsTmp.MoveNext
                Loop
                If Not mblnReadOnly Then .Rows = .Rows + 1 '保持一空行用于输入
            End With
        ElseIf mvItem.导入结果 <> 1 Then
            optType(0).Value = True
        Else
            optType(2).Value = True
        End If
        optType(0).Enabled = True
        optType(1).Enabled = True
        '----
        If mvItem.执行方式 = 2 Then mvItem.执行方式 = 3
        optExecute(mvItem.执行方式).Value = True
        '----
        With vsResult
            .Rows = .FixedRows
            If mvItem.项目结果 <> "" Then
                If UBound(Split(mvItem.项目结果, vbTab)) >= 1 Then
                    strDefault = Split(mvItem.项目结果, vbTab)(1)
                End If
                arrResult = Split(Split(mvItem.项目结果, vbTab)(0), ",")
                For i = 0 To UBound(arrResult)
                    .Rows = .Rows + 1

                    .TextMatrix(.Rows - 1, col执行结果) = Split(arrResult(i), "|")(0)
                    .Cell(flexcpData, .Rows - 1, col执行结果) = .TextMatrix(.Rows - 1, col执行结果)

                    '处理结果性质
                    If UBound(Split(arrResult(i), "|")) > 0 Then
                        Set .Cell(flexcpPicture, .Rows - 1, col执行图标) = imgNature.ListImages(Val(Split(arrResult(i), "|")(1))).Picture
                        mrsNature.Filter = "编码=" & Val(Split(arrResult(i), "|")(1))
                        .TextMatrix(.Rows - 1, col结果性质) = mrsNature!名称
                    End If

                    If Split(arrResult(i), "|")(0) = strDefault Then
                        .TextMatrix(.Rows - 1, col缺省结果) = 1
                    End If
                Next
            End If
            If Not mblnReadOnly And Not mblnAdjust Then .Rows = .Rows + 1 '保持一空行用于输入
        End With
    Else
        '新增时读取基本的执行结果
        mvItem.导入结果 = 1
        mrsResult.Filter = "基本=1"
        If Not mrsResult.EOF Then
            vsResult.Rows = vsResult.FixedRows + 1
            Call SetResultInput(vsResult.FixedRows, mrsResult)
        End If
        If optType(0).Value = True Then Call UCAdvice.Set选择列的可见性(optSend(0).Value)
    End If

    If Not mblnReadOnly Then
        vsEPR.Row = 0: vsEPR.Row = 1: vsEPR.Col = 0
        If Not mblnAdjust Then
            vsResult.Row = 0: vsResult.Row = 1: vsResult.Col = col执行结果
        End If
    End If

    '只读查看时的一些界面处理
    If mblnReadOnly Then
        cmdCancel.Visible = False
        cmdOK.Left = cmdCancel.Left

        cmdAdvice.Visible = False

        vsEPR.Editable = flexEDNone
        vsResult.Editable = flexEDNone

        For Each objControl In Me.Controls
            If TypeName(objControl) = "TextBox" Then
                objControl.Locked = True
            ElseIf TypeName(objControl) = "OptionButton" Then
                objControl.Enabled = False
            End If
        Next
    ElseIf mblnAdjust Then
        txtItem.BackColor = Me.BackColor
        txtItem.TabStop = False

        vsResult.Editable = flexEDNone
        vsResult.BackColor = Me.BackColor
        vsResult.BackColorBkg = Me.BackColor
        vsResult.TabStop = False

        For Each objControl In Me.Controls
            If TypeName(objControl) = "TextBox" Then
                objControl.Locked = True
            ElseIf TypeName(objControl) = "OptionButton" Then
                objControl.Enabled = False
            End If
        Next
    End If
    '导入参考
    rtfImportRef.Text = mvItem.导入参考

    Call SetFormFace

    mblnChange = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK And mvItem.ID <> 0 And mblnChange Then
        If MsgBox("该路径项目的信息已被更改，确实要放弃更改退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If

    If Not mrsResult Is Nothing Then
        If mrsResult.State = 1 Then mrsResult.Close
        Set mrsResult = Nothing
    End If
    If Not mrsNature Is Nothing Then
        If mrsNature.State = 1 Then mrsNature.Close
        Set mrsNature = Nothing
    End If

    mlngItemID = 0
    mblnAdjust = False
    mblnReadOnly = False
End Sub

Private Sub optSend_Click(Index As Integer)
    Call UCAdvice.Set选择列的可见性(optSend(0).Value)

    If Visible Then mblnChange = True
End Sub

Private Sub optType_Click(Index As Integer)
    Call SetFormFace

    If Visible Then
        If Index = 0 Then
            If cmdAdvice.Enabled Then cmdAdvice.SetFocus
            Call optSend_Click(0)
        ElseIf Index = 1 Then
            vsEPR.SetFocus
        End If
    End If
End Sub

Private Sub optExecute_Click(Index As Integer)
    Call SetFormFace

    If Visible Then mblnChange = True
End Sub

Private Sub picBottom_GotFocus()
    If cmdOK.Enabled And cmdOK.Visible Then
        cmdOK.SetFocus
    ElseIf cmdCancel.Enabled And cmdCancel.Visible Then
        cmdCancel.SetFocus
    End If
End Sub

Private Sub SetFormFace()
'功能：根据内容属性设置界面的可见内容和尺寸
    Dim lngTop As Long, lngHeight As Long

    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    fraAdvice.Enabled = optType(0).Value: fraAdvice.Visible = fraAdvice.Enabled
    fraEPR.Enabled = optType(1).Value: fraEPR.Visible = fraEPR.Enabled
    fraExecute.Enabled = Not optExecute(0).Value And mblnUseExecute: fraExecute.Visible = fraExecute.Enabled
    If optType(1).Value Then
        If gobjEmr Is Nothing Then
            fraERPType.Visible = False: lblEPR.Visible = False
            optEPRType(0).Value = True
        Else
            fraERPType.Visible = True: lblEPR.Visible = True
            optEPRType(1).Value = True
        End If
    Else
        fraERPType.Visible = False: lblEPR.Visible = False
    End If

    fraImportRef.Enabled = mvItem.导入结果 <> 1
    fraImportRef.Visible = fraImportRef.Enabled And fraAdvice.Enabled
    '当在Load事件中调用该过程时，设置fraImportRef.Visible=True这条语句无效，其值始终保持False
    If fraImportRef.Enabled And fraAdvice.Enabled Then
        fraImportRef.BackColor = fraAdvice.BackColor
        fraImportRef.Top = IIf(fraExecute.Enabled, fraExecute.Top, picBottom.Top) - 2000
        fraImportRef.Height = 2000
        rtfImportRef.Top = lblImportRef.Top + lblImportRef.Height + 30
        rtfImportRef.Height = fraImportRef.Height - rtfImportRef.Top
        fraAdvice.Height = fraImportRef.Top - fraAdvice.Top
    Else
        fraAdvice.Height = Me.Height - fraAdvice.Top - IIf(fraExecute.Enabled, fraExecute.Height, 0) - picBottom.Height - 450
        fraEPR.Height = fraAdvice.Height
    End If

    UCAdvice.Height = fraAdvice.Height - cmdAdvice.Height - 60
    vsEPR.Height = fraEPR.Height - 60
End Sub

Private Sub picIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long

    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vPoint As POINTAPI

    On Error GoTo errH

    If mblnReadOnly Or mblnAdjust Then Exit Sub

    If img16.ListImages.count = 0 Then
        strSql = "Select ID,Nvl(性质,0) as 性质 From 临床路径图标 Order by 性质 Desc,ID"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
        Do While Not rsTmp.EOF
            img16.ListImages.Add , "_" & IIf(rsTmp!性质 = 1, 1, -1) * rsTmp!ID, GetPathIcon(rsTmp!ID)
            img16.ListImages(img16.ListImages.count).Tag = CStr(rsTmp!ID) '要CStr
            rsTmp.MoveNext
        Loop
        cbsIcon.AddImageList img16
    End If

    Set objPopup = cbsIcon.Add("Popup", xtpBarPopup)
    objPopup.SetPopupToolBar True
    objPopup.Width = 260
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, -1, "清除项目图标")
        objControl.Flags = xtpFlagControlStretched
        For i = 1 To img16.ListImages.count
            Set objControl = .Add(xtpControlButton, img16.ListImages(i).Tag, "")
            If i = 1 Then
                objControl.BeginGroup = True
            ElseIf Val(Mid(img16.ListImages(i).Key, 2)) < 0 Then
                If Val(Mid(img16.ListImages(i - 1).Key, 2)) > 0 Then
                    objControl.BeginGroup = True
                End If
            End If
        Next
    End With

    vPoint.X = (fraContent.Left + lblIcon.Left - lblIcon.Width - 120) / Screen.TwipsPerPixelX
    vPoint.Y = (fraContent.Top + picIcon.Height) / Screen.TwipsPerPixelY
    ClientToScreen Me.Hwnd, vPoint
    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txtItem_GotFocus()
    Call zlControl.TxtSelAll(txtItem)
End Sub

Private Sub InitSchemeRecordset(rsScheme As ADODB.Recordset)
    Set rsScheme = New ADODB.Recordset
    rsScheme.Fields.Append "是否备选", adSmallInt
    rsScheme.Fields.Append "是否缺省", adSmallInt
    rsScheme.Fields.Append "序号", adBigInt
    rsScheme.Fields.Append "相关序号", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "期效", adSmallInt
    rsScheme.Fields.Append "诊疗项目ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "医嘱内容", adVarChar, 1000, adFldIsNullable
    rsScheme.Fields.Append "天数", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "单次用量", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "总给予量", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "医生嘱托", adVarChar, 1000, adFldIsNullable
    rsScheme.Fields.Append "执行频次", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "频率次数", adSmallInt, , adFldIsNullable
    rsScheme.Fields.Append "频率间隔", adSmallInt, , adFldIsNullable
    rsScheme.Fields.Append "间隔单位", adVarChar, 10, adFldIsNullable
    rsScheme.Fields.Append "时间方案", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "执行性质", adSmallInt
    rsScheme.Fields.Append "标本部位", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "检查方法", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "配方ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "组合项目ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "执行标记", adSingle, , adFldIsNullable

    rsScheme.CursorLocation = adUseClient
    rsScheme.LockType = adLockOptimistic
    rsScheme.CursorType = adOpenStatic
    rsScheme.Open
End Sub

Private Function ShowAdvice(ByVal str医嘱IDs As String) As Boolean
'功能：显示路径项目对应的医嘱内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim str中药 As String, str煎法 As String
    Dim str麻醉 As String, str标本 As String
    Dim strFilter As String
    Dim i As Long, j As Long

    If str医嘱IDs = "" Then
        Call UCAdvice.ShowAdvice(0, "", 0, 0, mblnReadOnly)
        ShowAdvice = True: Exit Function
    End If

    '生成动态SQL
    For i = 0 To UBound(Split(str医嘱IDs, ","))
        strFilter = strFilter & " Or ID=" & Split(str医嘱IDs, ",")(i)
    Next
    With mrsAdvice
        strSql = ""
        .Filter = Mid(strFilter, 5)
        Do While Not .EOF
            strSql = strSql & " Union ALL Select "
            For i = 0 To .Fields.count - 1
                If Not IsNull(.Fields(i).Value) Then
                    If Rec.IsType(.Fields(i).Type, adVarChar) Then
                        strSql = strSql & "'" & Replace(Replace(.Fields(i).Value, "[", "("), "]", ")") & "'"
                    Else
                        strSql = strSql & .Fields(i).Value '没有日期型
                    End If
                Else
                    If Rec.IsType(.Fields(i).Type, adBigInt) Or Rec.IsType(.Fields(i).Type, adSmallInt) Or Rec.IsType(.Fields(i).Type, adSingle) Then
                        strSql = strSql & "-Null"
                    Else
                        strSql = strSql & "Null"
                    End If
                End If
                strSql = strSql & " As " & .Fields(i).Name & ","
            Next
            strSql = Left(strSql, Len(strSql) - 1) & " From Dual"
            .MoveNext
        Loop
        .Filter = ""
        strSql = Mid(strSql, 12)
    End With
    If strSql = "" Then
        Call UCAdvice.ShowAdvice(0, "", 0, 0, mblnReadOnly)
    Else
        Call UCAdvice.ShowAdvice(0, strSql, 0, 0, mblnReadOnly)
    End If
    ShowAdvice = True
End Function

Private Sub UcAdvice_DataChange()
    If Visible Then mblnChange = True
End Sub

Private Sub vsEPR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsEPR_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsEPR_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsEPR
        If NewCol = 0 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsEPR_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI

    With vsEPR
        If optEPRType(0).Value Then  '表示新版电子病历未安装或安装不正常（走老版流程）
            strSql = "Select A.ID,Decode(A.种类,1,'门诊病历',5,'疾病证明报告',6,'知情文件') as 种类," & _
                " A.编号,A.名称,A.说明 From 病历文件列表 A" & _
                " Where A.种类 IN(1,5,6) And Nvl(A.保留,0) IN(0,1,2) And A.通用 IN(1,2)" & _
                " Order by A.种类,A.编号"
        Else
            '新版流程
            If gobjEmr Is Nothing Then Exit Sub

            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                Set gobjEmr = Nothing
                MsgBox "电子病历服务器不在线或导航台登录时未能成功连接电子病历服务器!", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            Else
                'gobjEmr.GetAntetypeList(byref strParameter as string) as Adodb.RecordSet
                '记录集包含字段：分类编号，分类名称，分组名称，ID，编号，名称，说明
                On Error Resume Next
                Set rsTmp = gobjEmr.GetAntetypeList("")
                Err.Clear: On Error GoTo 0
                If rsTmp Is Nothing Then Exit Sub
                strSql = Rec.ToSQL(rsTmp)
            End If
        End If
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "病历文件", False, "", "", False, False, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有病历文件数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call SetEPRInput(Row, rsTmp)
            Call EPREnterNextCell(True)
        End If
    End With
End Sub

Private Sub vsEPR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long

    If mblnReadOnly Then Exit Sub

    With vsEPR
        If KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 0) <> "" Then
                If MsgBox("确实要清除该行病历文件吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsEPR_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsEPR_KeyPress(KeyAscii As Integer)
    With vsEPR
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EPREnterNextCell
        ElseIf .Col = 0 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsEPR_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsEPR_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsEPR_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsEPR.EditSelStart = 0
    vsEPR.EditSelLength = zlCommFun.ActualLen(vsEPR.EditText)
End Sub

Private Sub vsEPR_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim i As Long
    Dim strFilter As String, strTag As String

    With vsEPR
        If Col = 0 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call EPREnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call EPREnterNextCell
            Else
                strInput = UCase(.EditText)
                If optEPRType(0).Value Then
                    strSql = "Select A.ID,Decode(A.种类,1,'门诊病历',4,'护理病历',5,'疾病证明报告',6,'知情文件') as 种类," & _
                             " A.编号,A.名称,A.说明 From 病历文件列表 A" & _
                             " Where A.种类 IN(1,4,5,6) And Nvl(A.保留,0) IN(0,1,2)" & _
                             " And A.通用 IN(1,2) And (A.编号 Like [1] Or A.名称 Like [2] Or zlSpellCode(A.名称) Like [2])" & _
                             " Order by A.种类,A.编号"
                Else
                    '新版流程
                    If gobjEmr Is Nothing Then Exit Sub
                    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                        Set gobjEmr = Nothing
                        MsgBox "电子病历服务器不在线或导航台登录时未能成功连接电子病历服务器!", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        'gobjEmr.GetAntetypeList(byref strParameter as string) as Adodb.RecordSet
                        '记录集包含字段：分类编号，分类名称，分组名称，ID，编号，名称，说明
                        On Error Resume Next
                        Set rsTmp = gobjEmr.GetAntetypeList("")
                        Err.Clear: On Error GoTo 0
                        If rsTmp Is Nothing Then Exit Sub
                        strSql = Rec.ToSQL(rsTmp)
                        strSql = "select A.* from (" & strSql & ") A where A.编号 Like [1] Or A.名称 Like [2] Or zlSpellCode(A.名称) Like [2]  order by 分类编号,编号"
                    End If

                End If
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "病历文件", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有找到匹配的病历文件。", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call SetEPRInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call EPREnterNextCell(True)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub SetEPRInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理病历文件的输入
    Dim strItem As String, i As Long
    Dim strTmp As String

    With vsEPR
        For i = 1 To rsInput.RecordCount
            If i > 1 Then
                .AddItem "", lngRow + 1
                lngRow = lngRow + 1
            End If
            If optEPRType(0).Value Then
                strTmp = "OLD" '旧版
            Else
                strTmp = "NEW" '新版
            End If
            .RowData(lngRow) = strTmp & "|" & rsInput!ID   '新版ID是32位字符串
            .TextMatrix(lngRow, 0) = rsInput!名称
            .Cell(flexcpData, lngRow, 0) = .TextMatrix(lngRow, 0)

            strItem = strItem & "、" & rsInput!名称

            rsInput.MoveNext
        Next

        '缺省项目内容
        If txtItem.Text = "" Then txtItem.Text = "书写" & Mid(strItem, 2)

        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If

        mblnChange = True
    End With
End Sub

Private Sub EPREnterNextCell(Optional ByVal blnNewRow As Boolean)
    Dim i As Long, j As Long

    With vsEPR
        If blnNewRow Then
            .Row = .Rows - 1: .Col = 0
            .ShowCell .Row, .Col
        Else
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long

    With vsResult
        If Col = col结果性质 Then
            .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.TextMatrix(Row, Col))
            If .TextMatrix(Row, Col) <> "" Then
                mrsNature.Filter = "名称='" & .TextMatrix(Row, Col) & "'"
                Set .Cell(flexcpPicture, .Row, col执行图标) = imgNature.ListImages(Val(mrsNature!编码)).Picture
            Else
                Set .Cell(flexcpPicture, .Row, col执行图标) = Nothing
            End If
        ElseIf Col = col缺省结果 Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                For i = .FixedRows To .Rows - 1
                    If i <> Row Then .TextMatrix(i, col缺省结果) = 0
                Next
            End If
        End If
    End With
End Sub

Private Sub vsResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsResult
        If Not ResultCellEditable(NewRow, NewCol) Then
            .FocusRect = flexFocusLight
            .ComboList = ""
        ElseIf NewCol = col执行结果 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        ElseIf NewCol = col结果性质 Then
            .ComboList = .ColData(NewCol)
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsResult_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI

    With vsResult
        '如果用子查询，则数据树形顺序不对，需要特别排序
        strSql = " Select A.编码 As ID,A.上级 As 上级id,A.编码,A.名称,A.简码,Nvl(A.末级,0) As 末级,B.名称 As 性质" & _
                 " From 路径常见结果 A,路径结果性质 B Where Nvl(A.性质,0)=B.编码(+)" & _
                 " Start With A.上级 Is Null Connect By Prior A.编码=A.上级"
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 2, "常见结果", True, "", "", False, True, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有常见结果数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call SetResultInput(Row, rsTmp)
            Call ResultEnterNextCell
        End If
    End With
End Sub

Private Sub vsResult_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long

    If Col = col结果性质 Then
        With vsResult
            If .TextMatrix(Row, Col) <> "" Then
                For i = 0 To .ComboCount - 1
                    If zlCommFun.GetNeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                        .ComboIndex = i: Exit For
                    End If
                Next
            End If
        End With
    End If
End Sub

Private Sub vsResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long

    If mblnReadOnly Or mblnAdjust Then Exit Sub

    With vsResult
        If KeyCode = vbKeyDelete Then
            If .Col = col结果性质 Then
                .TextMatrix(.Row, .Col) = ""
                Set .Cell(flexcpPicture, .Row, col执行图标) = Nothing
            ElseIf .TextMatrix(.Row, col执行结果) <> "" Then
                If MsgBox("确实要清除该行结果吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsResult_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsResult_KeyPress(KeyAscii As Integer)
    With vsResult
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call ResultEnterNextCell
        ElseIf KeyAscii = Asc(",") Then
            KeyAscii = 0
        ElseIf .Col = col执行结果 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsResult_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsResult_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    If KeyAscii = Asc(",") Then KeyAscii = 0
End Sub

Private Sub vsResult_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsResult.EditSelStart = 0
    vsResult.EditSelLength = zlCommFun.ActualLen(vsResult.EditText)
End Sub

Private Sub vsResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ResultCellEditable(Row, Col) Then Cancel = True
End Sub

Private Function ResultCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    ResultCellEditable = True

    With vsResult
        If lngCol = col执行图标 Then
            ResultCellEditable = False
        ElseIf lngCol > col执行结果 And .TextMatrix(lngRow, col执行结果) = "" Then
            ResultCellEditable = False
        ElseIf lngCol = col结果性质 And .TextMatrix(lngRow, col执行结果) <> "" Then
            '字典中的结果性质不允许更改,手工输入的才允许
            If .TextMatrix(lngRow, col结果性质) <> "" Then
                mrsResult.Filter = "名称='" & .TextMatrix(lngRow, col执行结果) & "'"
                If Not mrsResult.EOF Then
                    If NVL(mrsResult!性质) = .TextMatrix(lngRow, col结果性质) Then
                        ResultCellEditable = False
                    End If
                End If
            End If
        End If
    End With
End Function

Private Sub vsResult_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI

    With vsResult
        If Col = col执行结果 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call ResultEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ResultEnterNextCell
            Else
                strInput = UCase(.EditText)
                strSql = "Select A.编码 as ID,A.编码,A.名称,A.简码,B.名称 as 性质" & _
                    " From 路径常见结果 A,路径结果性质 B" & _
                    " Where Nvl(A.性质,0)=B.编码(+) And A.末级=1" & _
                    " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])" & _
                    " Order by A.编码"
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "常见结果", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%")
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    Call SetResultInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call ResultEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col结果性质 Then
            If mblnReturn Then Call ResultEnterNextCell
            mblnReturn = False
        End If
    End With
End Sub

Private Sub SetResultInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理项目结果的输入
    Dim i As Long

    With vsResult
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                End If

                .TextMatrix(lngRow, col执行结果) = rsInput!名称

                '处理结果性质
                If Not IsNull(rsInput!性质) Then
                    mrsNature.Filter = "名称='" & rsInput!性质 & "'"
                    Set .Cell(flexcpPicture, lngRow, col执行图标) = imgNature.ListImages(Val(mrsNature!编码)).Picture
                    .TextMatrix(lngRow, col结果性质) = rsInput!性质
                End If

                If i = 1 And lngRow = .FixedRows Then
                    .TextMatrix(lngRow, col缺省结果) = 1
                    Call vsResult_AfterEdit(lngRow, col缺省结果)
                End If

                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col执行结果) = .EditText
        End If
        .Cell(flexcpData, lngRow, col执行结果) = .TextMatrix(lngRow, col执行结果)

        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If

        mblnChange = True
    End With
End Sub

Private Sub ResultEnterNextCell()
    With vsResult
        If .Col + 1 <= .Cols - 1 Then
            .Col = .Col + 1
        ElseIf .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1: .Col = col执行结果
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub
