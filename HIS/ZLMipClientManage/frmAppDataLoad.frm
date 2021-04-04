VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppDataLoad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "消息数据安装"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7965
   Icon            =   "frmAppDataLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgList 
      Left            =   1500
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataLoad.frx":6852
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataLoad.frx":6DEC
            Key             =   "已完成"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataLoad.frx":7386
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataLoad.frx":7920
            Key             =   "待执行"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataLoad.frx":7EBA
            Key             =   "全清"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步(&N)"
      Height          =   345
      Left            =   6765
      TabIndex        =   6
      Top             =   4725
      Width           =   1100
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   45
      ScaleHeight     =   840
      ScaleWidth      =   7995
      TabIndex        =   4
      Top             =   0
      Width           =   7995
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "消息数据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   1260
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   7170
         Picture         =   "frmAppDataLoad.frx":8454
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   345
      Left            =   270
      TabIndex        =   3
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一步(&P)"
      Height          =   345
      Left            =   5610
      TabIndex        =   2
      Top             =   4725
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   870
      Width           =   8100
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   8100
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   3
      Left            =   -30
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   21
      Top             =   900
      Width           =   7950
      Begin VB.Frame Frame3 
         Height          =   2505
         Left            =   825
         TabIndex        =   22
         Top             =   570
         Width           =   6840
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   5
            Left            =   1320
            TabIndex        =   26
            Top             =   1184
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1320
            TabIndex        =   24
            Top             =   360
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   1320
            TabIndex        =   25
            Top             =   772
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   3
            Left            =   1320
            TabIndex        =   27
            Top             =   1596
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   28
            Top             =   2010
            Width           =   5070
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "端口号"
            Height          =   180
            Index           =   11
            Left            =   180
            TabIndex        =   33
            Top             =   1237
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "数据库地址"
            Height          =   180
            Index           =   4
            Left            =   180
            TabIndex        =   31
            Top             =   405
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "数据库实例"
            Height          =   180
            Index           =   7
            Left            =   180
            TabIndex        =   30
            Top             =   821
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "数据库所有者"
            Height          =   180
            Index           =   8
            Left            =   180
            TabIndex        =   29
            Top             =   1653
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "所有者密码"
            Height          =   180
            Index           =   9
            Left            =   180
            TabIndex        =   23
            Top             =   2070
            Width           =   900
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "安装智能电子病历的消息数据(版本要求：10.34.10以上)"
         Height          =   180
         Index           =   3
         Left            =   870
         TabIndex        =   32
         Top             =   255
         Width           =   4500
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   0
         Left            =   165
         Picture         =   "frmAppDataLoad.frx":B8D6
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   5
      Left            =   15
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   15
      Top             =   900
      Width           =   7950
      Begin VB.CommandButton cmdSetup 
         Caption         =   "安装(&S)"
         Height          =   345
         Left            =   960
         TabIndex        =   17
         Top             =   3195
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   225
         Left            =   2130
         TabIndex        =   16
         Top             =   3345
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStep 
         Height          =   2490
         Left            =   975
         TabIndex        =   18
         Top             =   600
         Width           =   6840
         _cx             =   2088840993
         _cy             =   2088833320
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "正在安装.."
         Height          =   180
         Index           =   12
         Left            =   2145
         TabIndex        =   34
         Top             =   3150
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   3
         Left            =   195
         Picture         =   "frmAppDataLoad.frx":D258
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "点击“安装”即开始安装已勾选的消息数据"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   20
         Top             =   165
         Width           =   3420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   180
         Index           =   6
         Left            =   7395
         TabIndex        =   19
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   1
      Left            =   45
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   10
      Top             =   915
      Width           =   7950
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1155
         TabIndex        =   12
         Top             =   795
         Width           =   6330
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   0
         Left            =   7500
         Picture         =   "frmAppDataLoad.frx":EBDA
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   795
         Width           =   315
      End
      Begin MSComDlg.CommonDialog cdl 
         Left            =   7410
         Top             =   165
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统名"
         Height          =   180
         Index           =   10
         Left            =   1170
         TabIndex        =   35
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择消息集成平台客户端配置文件"
         Height          =   180
         Index           =   0
         Left            =   1170
         TabIndex        =   14
         Top             =   270
         Width           =   2700
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "版本号"
         Height          =   180
         Index           =   2
         Left            =   1185
         TabIndex        =   13
         Top             =   1305
         Width           =   540
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   2
         Left            =   225
         Picture         =   "frmAppDataLoad.frx":1542C
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   2
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   7
      Top             =   900
      Width           =   7950
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2730
         Left            =   975
         TabIndex        =   8
         Top             =   615
         Width           =   6840
         _cx             =   2088840993
         _cy             =   2088833743
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请勾选需要安装哪些系统的消息数据"
         Height          =   180
         Index           =   5
         Left            =   975
         TabIndex        =   9
         Top             =   225
         Width           =   2880
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   1
         Left            =   165
         Picture         =   "frmAppDataLoad.frx":16DAE
         Top             =   165
         Width           =   480
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   4
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   36
      Top             =   825
      Width           =   7950
      Begin VB.Frame Frame4 
         Height          =   2430
         Left            =   750
         TabIndex        =   37
         Top             =   720
         Width           =   6840
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   6
            Left            =   1320
            TabIndex        =   38
            Top             =   360
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   7
            Left            =   1320
            TabIndex        =   39
            Top             =   765
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   41
            Top             =   1575
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   9
            Left            =   1320
            TabIndex        =   40
            Top             =   1170
            Width           =   5070
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "服务器地址"
            Height          =   180
            Index           =   13
            Left            =   210
            TabIndex        =   45
            Top             =   405
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "服务器端口"
            Height          =   180
            Index           =   14
            Left            =   210
            TabIndex        =   44
            Top             =   810
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "连接密码"
            Height          =   180
            Index           =   15
            Left            =   405
            TabIndex        =   43
            Top             =   1635
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "连接用户"
            Height          =   180
            Index           =   16
            Left            =   405
            TabIndex        =   42
            Top             =   1215
            Width           =   720
         End
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   4
         Left            =   135
         Picture         =   "frmAppDataLoad.frx":18730
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入服务器地址、端口、用户、密码，以创建消息服务平台服务器连接。"
         Height          =   180
         Index           =   17
         Left            =   765
         TabIndex        =   46
         Top             =   375
         Width           =   5940
      End
   End
End
Attribute VB_Name = "frmAppDataLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mobjFso As New FileSystemObject
Private mclsOracle As clsDataOracle
Private mblnStep(1 To 2) As Boolean
Private mstrManageVersion As String
Private mstrVersion As String
Private mintPage As Integer
Private mclsVsf As zlVSFlexGrid.clsVsf
Private mclsVsfStep As zlVSFlexGrid.clsVsf
Private mclsVsfUser As zlVSFlexGrid.clsVsf
Private mbytMode As Byte
Private mcolSigns As New Collection
Private mblnSpecialEMR As Boolean
Private mstrEmrOra As String
Private mcnOracle As ADODB.Connection
Private WithEvents mclsMipClientManage As clsMipClientManage
Attribute mclsMipClientManage.VB_VarHelpID = -1
Private mfrmErrorInfo As frmErrorInfo
Private mblnImportDB As Boolean

Private WithEvents mobjScript As clsOracleScript
Attribute mobjScript.VB_VarHelpID = -1

Public Function ShowDialog() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    mblnOK = False
    
    Set mclsOracle = New clsDataOracle
    
    
    For intLoop = 1 To picPage.UBound
        picPage(intLoop).Left = 0
        picPage(intLoop).Top = 915
        picPage(intLoop).Width = 7950
        picPage(intLoop).Height = 3645
    Next
    
    Call InitGrid
    
    mbytMode = 1
    mintPage = 1
    Call ShowPage(mintPage)
    
    Me.Show 1
    ShowDialog = mblnOK
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTBoolean, "", "[选择]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("标识", 0, flexAlignLeftCenter, flexDTString, , "trigger_type", True)
        Call .AppendColumn("是否安装", 0, flexAlignLeftCenter, flexDTString, , "trigger_type", True)
        Call .AppendColumn("系统号", 0, flexAlignLeftCenter, flexDTString, , "trigger_type", True)
        Call .AppendColumn("系统名称", 3000, flexAlignLeftCenter, flexDTString, , "item_code", True)
        Call .AppendColumn("所 有 者", 900, flexAlignLeftCenter, flexDTString, , "item_title", True)
        Call .AppendColumn("系统版本", 1080, flexAlignLeftCenter, flexDTString, , "trigger_type", True)
        Call .AppendColumn("最低版本", 1080, flexAlignLeftCenter, flexDTString, , "trigger_type", True)
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(vsf.ColIndex("选择"), True, vbVsfEditCheck)

        vsf.Cell(flexcpPicture, 0, .ColIndex("选择")) = imgList.ListImages("全选").Picture
        .AppendRows = True
        
    End With
    '------------------------------------------------------------------------------------------------------------------
            
    Set mclsVsfStep = New zlVSFlexGrid.clsVsf
    With mclsVsfStep
        Call .Initialize(Me.Controls, vsfStep, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("step", 1500, flexAlignLeftCenter, flexDTString, , "item_note", True)
        vsfStep.RowHidden(0) = True
    End With
    
    InitGrid = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ShowPage(ByVal intPage As Integer)
    Dim intLoop As Integer
    
    For intLoop = 1 To picPage.UBound
        picPage(intLoop).Visible = False
    Next
    
    picPage(intPage).Visible = True
        
    cmdNext.Enabled = (intPage < picPage.UBound)
    cmdPrevious.Enabled = (intPage > 1)
    
End Sub

Private Function OpenDialog(ByRef objDlg As Object, ByVal strDialogTitle As String, ByVal strFilter As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim strTmp As String
            
    With objDlg
        .DialogTitle = strDialogTitle
        .Filter = strFilter
    
        On Error Resume Next
    
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowOpen
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            OpenDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    MsgBox "不能打开文件(" & strTmp & "),该文件可能正在使用或文件不存在!"
End Function

Private Function SaveDialog(ByRef objDlg As Object, ByVal strDialogTitle As String, ByVal strFilter As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim strTmp As String
            
    With objDlg
        .DialogTitle = strDialogTitle
        .Filter = strFilter
    
        On Error Resume Next
    
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowSave
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            SaveDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    MsgBox "不能保存为文件(" & strTmp & "),该文件可能正在使用或文件已经存在!"
End Function

Private Sub cmd_Click(Index As Integer)
    Dim strFile As String
    
    Select Case Index
    Case 0
        strFile = OpenDialog(cdl, "请选择配置文件", "配置文件(*.ini)|*.ini")
        
        If strFile <> "" Then
            txt(0).Text = strFile
            mblnStep(1) = CheckSetupFile(strFile)
        End If
    Case 1
        strFile = SaveDialog(cdl, "请选择日志文件", "日志文件(*.log)|*.log")
        
        If strFile <> "" Then
            txt(1).Text = strFile
        End If
    End Select
    
End Sub


Private Function CheckPassword(ByVal strUser As String, ByVal strPassword As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    CheckPassword = mclsOracle.OraDataOpen(gstrServerName, strUser, strPassword, True)
End Function

Private Function CheckSetupFile(ByVal strFile As String) As Boolean
    '******************************************************************************************************************
    '功能：检查解释安装配置文件的正确性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strIniPath As String
    Dim strTemp As String
    Dim objText As TextStream
    Dim strManageVersion As String
    Dim intLoop As Integer
    Dim aryTemp As Variant
    Dim aryItem As Variant
    Dim aryFlag As Variant
    Dim strSysName As String
    Dim intRows As Integer
    Dim rsData As ADODB.Recordset
    
    strIniPath = Mid(strFile, 1, Len(strFile) - 11)
    
    '相关文件匹配性检查
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    If Dir(strIniPath & "zlMipClientStruct.SQL") = "" Then strTemp = strTemp & vbCr & "结构文件" & strIniPath & "zlMipClientStruct.SQL"
    If Dir(strIniPath & "zlMipClientData.SQL") = "" Then strTemp = strTemp & vbCr & "数据文件" & strIniPath & "zlMipClientData.SQL"
    
    If strTemp <> "" Then
        MsgBox "以下安装的相关文件丢失，不能继续，包括：" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '安装配置文件解释
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error Resume Next
    Set objText = mobjFso.OpenTextFile(strFile)
    


    mstrVersion = ""
    mstrManageVersion = ""
    
    strTemp = Trim(objText.ReadLine)
    
    If Left(strTemp, 5) = "[组件名]" Then
        strSysName = Trim(Mid(strTemp, 6))
    Else
        Err.Raise 10
    End If
    
    strTemp = Trim(objText.ReadLine)
    
    If Left(strTemp, 5) = "[版本号]" Then
        mstrVersion = Trim(Mid(strTemp, 6))
    Else
        Err.Raise 10
    End If
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[数据组]" Then
        strTemp = Trim(Mid(strTemp, 7))
'
'        lst.Clear
'        aryTemp = Split(strTemp, "|")
'        For intLoop = 0 To UBound(aryTemp)
'            aryItem = Split(aryTemp(intLoop), "=")
'            lst.AddItem aryItem(0)
'            lst.ItemData(lst.NewIndex) = aryItem(1)
'        Next
        With vsf
            .Rows = 1
            aryTemp = Split(strTemp, "|")
            For intLoop = 0 To UBound(aryTemp)
                aryItem = Split(aryTemp(intLoop), "=")
                aryFlag = Split(aryItem(1), ",")
                '首先根据配置文件中的编号判断是否已安装了该系统
                Set rsData = CheckSysInfo(aryFlag(1))
                If Not rsData Is Nothing Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("id")) = aryFlag(2)
                    .TextMatrix(.Rows - 1, .ColIndex("标识")) = aryFlag(0)
                    .TextMatrix(.Rows - 1, .ColIndex("系统名称")) = aryItem(0)
                    .TextMatrix(.Rows - 1, .ColIndex("是否安装")) = gclsBusiness.CheckSetuped(aryFlag(0))
                    If aryFlag(0) <> "EMR" Then
                        .TextMatrix(.Rows - 1, .ColIndex("系统号")) = rsData("编号").Value
                        .TextMatrix(.Rows - 1, .ColIndex("所 有 者")) = rsData("所有者").Value
                        .TextMatrix(.Rows - 1, .ColIndex("系统版本")) = rsData("版本号").Value
                    Else
                        .TextMatrix(.Rows - 1, .ColIndex("系统号")) = "-"
                        .TextMatrix(.Rows - 1, .ColIndex("所 有 者")) = "-"
                        .TextMatrix(.Rows - 1, .ColIndex("系统版本")) = "-"
                    End If
                    .TextMatrix(.Rows - 1, .ColIndex("最低版本")) = aryFlag(3)
                    .TextMatrix(.Rows - 1, .ColIndex("选择")) = 1
                    
                    If Val(.TextMatrix(.Rows - 1, .ColIndex("是否安装"))) = 1 Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = 8421504
                    End If
                End If
                
                If aryFlag(0) = "EMR" And aryFlag(2) = 2 And Val(.TextMatrix(.Rows - 1, .ColIndex("是否安装"))) = 0 Then
                    lbl(3).Caption = "安装新版病历的消息数据(版本要求：" & aryFlag(3) & "以上)"
                    lbl(3).Tag = aryFlag(3)
                End If
            Next
            mclsVsf.AppendRows = True
        End With
    Else
        Err.Raise 10
    End If
    
    lbl(2).Caption = "版本号：" & mstrVersion
    lbl(10).Caption = "系统名：" & strSysName
        
    objText.Close
    
    
    CheckSetupFile = True
End Function

Private Function CheckSysInfo(ByVal lngCode As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：根据系统号查询系统信息
    '参数：lngCode 系统号
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "Select 编号,名称,所有者,版本号 From zlSystems Where 编号=[1]"
    Set rsData = zlDataBase.OpenSQLRecord(strSQL, "系统信息", lngCode)
    If rsData.BOF = False Then
        Set CheckSysInfo = rsData
    Else
        Set CheckSysInfo = Nothing
    End If
    
End Function

Private Function VersionValid(ByVal strSysVersion As String, ByVal strFileVersion As String) As Boolean
    Dim dblSysVersion As Double
    Dim dblFileVersion As Double
    
    dblSysVersion = GetVerDouble(strSysVersion)
    dblFileVersion = GetVerDouble(strFileVersion)
    If dblSysVersion < dblFileVersion Then
        VersionValid = False
        Exit Function
    End If
    
    VersionValid = True

End Function

Private Sub SelectedAll()
    '******************************************************************************************************************
    '功能：表格全选和全清
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intRow As Integer
    
    With vsf
        Select Case mbytMode
        Case 1
            '现状态为全选，将变为全清
            For intRow = 1 To .Rows - 1
                If Val(.TextMatrix(intRow, .ColIndex("是否安装"))) = 0 Then
                    .TextMatrix(intRow, .ColIndex("选择")) = 0
                End If
            Next
            .Cell(flexcpPicture, 0, .ColIndex("选择")) = imgList.ListImages("全清").Picture
            mbytMode = 2
        Case 2
            '现状态为全清，将变为全选
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, .ColIndex("选择")) = 1
            Next
            .Cell(flexcpPicture, 0, .ColIndex("选择")) = imgList.ListImages("全选").Picture
            mbytMode = 1
        End Select
    End With
    
End Sub

Private Function CheckEMRConn()
    Dim strUserName As String
    Dim strServerIP As String
    Dim strPassword As String
    Dim strSID As String
    Dim strPort As String
    Dim strNote As String
    On Error GoTo InputError
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txt(3).Text)
    strPassword = Trim(txt(4).Text)
    strServerIP = Trim(txt(1).Text)
    strSID = Trim(txt(2).Text)
    strPort = Trim(txt(5).Text)
    
    '有效字符串效验
    If Len(strServerIP) = 0 Then
        strNote = "数据库地址(IP)"
        txt(1).SetFocus
    End If
    
    If Len(strSID) = 0 Then
        strNote = strNote & vbCrLf & "数据库实例"
        txt(2).SetFocus
    End If
    
    If Len(Trim(strPort)) = 0 Then
        strNote = strNote & vbCrLf & "端口号"
        txt(5).SetFocus
    End If
    
    If Len(strUserName) = 0 Then
        strNote = strNote & vbCrLf & "所有者"
        txt(3).SetFocus
    End If
    
    If strNote <> "" Then
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt(3).SetFocus
            strNote = "用户名错误"
            GoTo InputError
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            txt(4).SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    
    If OraDataOpen(strServerIP, strSID, strUserName, strPassword, strPort) Then
'        mstrUserName = strUserName
'        mstrUserPwd = strPassword
'        mstrServerIP = strServerIP
'        mstrSID = strSID
'        mstrPort = strPort
        CheckEMRConn = True
        Exit Function
    Else
        CheckEMRConn = False
    End If
    Exit Function
InputError:
    If strNote <> "" Then
        MsgBox "以下信息必须输入:" & vbCrLf & strNote, vbExclamation + vbOKOnly, "提示信息"
    End If
    Exit Function
End Function

Private Function OraDataOpen(ByVal strServerIP As String, ByVal strSID As String, ByVal strUserName As String, ByVal strUserPwd As String, ByVal strPort As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    Dim strServer As String
    
    Set mcnOracle = New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    DoEvents
    strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strServerIP & ")(PORT = " & strPort & ")))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
    With mcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If zlComLib.ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Public Function GetVerDouble(ByVal varVer As Variant) As Double
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '功能：根据版本字符串，产生数字化的版本
    '参数：varVer   版本字符串，如9.5.0
    Dim varArray As Variant
    
    varVer = IIf(IsNull(varVer), "", varVer)
    varArray = Split(varVer, ".")
    
    If UBound(varArray) < 2 Then Exit Function
    
    GetVerDouble = Val(varArray(0)) * 10 ^ 8 + Val(varArray(1)) * 10 ^ 4 + Val(varArray(2))
End Function

Public Function SetupMipClient(ByVal strInstallFile As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strPath As String
    Dim intLoop As Integer
    Dim strSQL As String
    Dim intPercent As Integer
    Dim intCount As Integer
    Dim intRow As Integer
    Dim intFlag As Integer
    Dim rsErr As ADODB.Recordset
    
    On Error GoTo errHand
    
    strPath = Left(strInstallFile, Len(strInstallFile) - Len("zlSetup.ini"))
    
'    '安装结构
'    '------------------------------------------------------------------------------------------------------------------
    Set mobjScript = New clsOracleScript

    lbl(12).Visible = True
    lbl(6).Visible = True
    pgb.Visible = True
    '安装业务数据
    intCount = intCount + 1
    pgb.Value = 0
    With vsf
        For intRow = 1 To .Rows - 1
            If Abs(.TextMatrix(intRow, .ColIndex("选择"))) = 1 And Val(.TextMatrix(intRow, .ColIndex("是否安装"))) = 0 Then
                If (vsf.TextMatrix(intRow, vsf.ColIndex("标识")) <> "EMR") Or (vsf.TextMatrix(intRow, vsf.ColIndex("标识")) = "EMR" And mblnSpecialEMR = True) Then
                    If Dir(strPath & .TextMatrix(intRow, .ColIndex("标识")) & "\zlMipClientData" & ".SQL") = "" Then
                        MsgBox "zlMipClientData_" & .TextMatrix(intRow, .ColIndex("标识")) & ".SQL文件不存在!"
                    Else
                        If mobjScript.OpenScriptFile(strPath & .TextMatrix(intRow, .ColIndex("标识")) & "\zlMipClientData" & ".SQL") Then
                            lbl(12).Caption = "正在安装" & .TextMatrix(intRow, .ColIndex("系统名称")) & "系统数据脚本..."
                            
                            vsfStep.Cell(flexcpPicture, intCount, vsfStep.ColIndex("图标")) = imgList.ListImages("执行中").Picture
                            For intLoop = 1 To mobjScript.SQLCount
                                Call mobjScript.ExecuteSQL(gclsMsgOracle.DatabaseConnection, mobjScript.SQL(intLoop))
                                intPercent = 100 * intLoop / mobjScript.SQLCount
                                If pgb.Value <> intPercent Then pgb.Value = intPercent
                                lbl(6).Caption = intPercent & "%"
                            Next
                            
                            '插入安装数据
                            strSQL = "Insert Into zlmip_data_setup(data_code,data_title,data_owner,data_system,data_source,data_db,setup_time) " & _
                                    "Select '" & .TextMatrix(intRow, .ColIndex("标识")) & "','" & .TextMatrix(intRow, .ColIndex("系统名称")) & _
                                    "','" & IIf(.TextMatrix(intRow, .ColIndex("id")) = 1, .TextMatrix(intRow, .ColIndex("所 有 者")), UCase(txt(3).Text)) & "'," & Val(.TextMatrix(intRow, .ColIndex("系统号"))) & _
                                    ",'" & IIf(.TextMatrix(intRow, .ColIndex("id")) = 1, "", mstrEmrOra) & "','" & .TextMatrix(intRow, .ColIndex("id")) & "',to_date('" & Format(Now, "YYYY-MM-DD HH:mm:SS") & "','YYYY-MM-DD HH24:MI:SS') From Dual"
                            gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
                            
                            vsfStep.Cell(flexcpPicture, intCount, vsfStep.ColIndex("图标")) = imgList.ListImages("已完成").Picture
                            intCount = intCount + 1
                        End If
                    End If
                End If
            End If
        Next
    End With
    intFlag = intCount
0:
    With vsf
        intCount = intFlag
        '初始化错误记录集
        Set rsErr = Nothing
        If rsErr Is Nothing Then
            Set rsErr = New ADODB.Recordset
            rsErr.Fields.Append "序号", adBSTR
            rsErr.Fields.Append "内容", adBSTR
            rsErr.Open
        End If
        For intRow = 1 To .Rows - 1
            If Abs(.TextMatrix(intRow, .ColIndex("选择"))) = 1 And Val(.TextMatrix(intRow, .ColIndex("是否安装"))) = 0 Then
                If (vsf.TextMatrix(intRow, vsf.ColIndex("标识")) <> "EMR") Or (vsf.TextMatrix(intRow, vsf.ColIndex("标识")) = "EMR" And mblnSpecialEMR = True) Then
                    If Dir(strPath & .TextMatrix(intRow, .ColIndex("标识")) & "\zlMipServerData" & ".db") <> "" Then
                        lbl(12).Caption = "正在向服务器导入" & .TextMatrix(intRow, .ColIndex("标识")) & "消息..."
                        vsfStep.Cell(flexcpPicture, intCount, vsfStep.ColIndex("图标")) = imgList.ListImages("执行中").Picture
                        Call mclsMipClientManage.CommunicateProxyInstall(strPath & .TextMatrix(intRow, .ColIndex("标识")) & "\zlMipServerData" & ".db", rsErr)
                        vsfStep.Cell(flexcpPicture, intCount, vsfStep.ColIndex("图标")) = imgList.ListImages("已完成").Picture
                        intCount = intCount + 1
                    End If
                End If
            End If
        Next
        
    End With
    
    If Not (rsErr Is Nothing) Then
        If rsErr.RecordCount > 0 Then
            If mfrmErrorInfo Is Nothing Then
                Set mfrmErrorInfo = New frmErrorInfo
            End If
            
            If mfrmErrorInfo.ShowError(Me, rsErr) = False Then
                GoTo 0
            End If
        End If
    End If
    If mblnImportDB Then
        Call mclsMipClientManage.CommunicateProxyLogout
    End If
    Set mclsMipClientManage = Nothing
    lbl(12).Caption = "数据安装完成!"
    
    SetupMipClient = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If MsgBox("出现下列错误，是否继续？" & vbCrLf & "    " & Err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
End Function

Private Sub cmdNext_Click()
    Dim intRow As Integer
    Dim blnSelected As Boolean
    Dim strSQL As String
    Dim rsData As New ADODB.Recordset
    
    Select Case mintPage
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        
        If txt(0).Text = "" Then
            ShowSimpleMsg "必须选择消息集成平台客户端安装配置文件！"
            Exit Sub
        End If
                
        If Dir(txt(0).Text) = "" Then
            ShowSimpleMsg "选择消息集成平台客户端安装配置文件不存在或者无效！"
            Exit Sub
        End If
        
        mintPage = mintPage + 1
        Call ShowPage(mintPage)
        
        '初始下一页
        
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        
'        '检查ZLTOOLS密码有效性
'        If CheckPassword("ZLTOOLS", txt(2).Text) = False Then
'            Exit Sub
'        End If
         '检查版本正确性
        With vsf
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, .ColIndex("标识")) <> "EMR" Then
                    If VersionValid(.TextMatrix(intRow, .ColIndex("系统版本")), .TextMatrix(intRow, .ColIndex("最低版本"))) = False And Abs(.TextMatrix(intRow, .ColIndex("选择"))) = 1 Then
                        MsgBox "【" & .TextMatrix(intRow, .ColIndex("系统名称")) & "】的系统版本不能低于要求的最低版本!", vbInformation + vbOKOnly, "提示信息"
                        Exit Sub
                    End If
                End If
            Next
        End With
        '判断是否选择了病历业务数据
        mblnSpecialEMR = False
        With vsf
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, .ColIndex("id")) = 2 And .TextMatrix(intRow, .ColIndex("标识")) = "EMR" And Abs(.TextMatrix(intRow, .ColIndex("选择"))) = 1 And Val(.TextMatrix(intRow, .ColIndex("是否安装"))) = 0 Then
                    mblnSpecialEMR = True
                    Exit For
                End If
            Next
        End With
        
        '判断是否没有选择任何一个未安装的系统
        With vsf
            For intRow = 1 To .Rows - 1
                If Abs(.TextMatrix(intRow, .ColIndex("选择"))) = 1 And Val(.TextMatrix(intRow, .ColIndex("是否安装"))) = 0 Then
                    blnSelected = True
                    Exit For
                End If
            Next
        End With
        If blnSelected = False Then
            MsgBox "当前未选择任何可被安装的系统!", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        If mblnSpecialEMR = True Then
            mintPage = mintPage + 1
        Else
            mintPage = mintPage + 2
            InitVsfSetup
        End If
        Call ShowPage(mintPage)
               
        
    '------------------------------------------------------------------------------------------------------------------
    Case 3
         '验证连接
        If CheckEMRConn = False Then
            Exit Sub
        Else
            '获取版本号
            strSQL = "Select value From sys_config Where Title='版本号'"
            If rsData.State = adStateOpen Then rsData.Close
            rsData.Open strSQL, mcnOracle
            
            If Err.Number <> 0 Then
                MsgBox "连接的服务器没有检测到新版电子病历，请确认是否安装。", vbInformation + vbOKOnly, gstrSysName
                Err.Clear
                Exit Sub
            End If
            
            '匹配版本号
            If VersionValid(rsData("value").Value, lbl(3).Tag) = False Then
                MsgBox "系统版本不能低于要求的最低版本!", vbInformation + vbOKOnly, "提示信息"
                Exit Sub
            End If
            
            '保存连接字符串
            mstrEmrOra = "<root>" & vbNewLine & _
                        "<ip>" & txt(1).Text & "</ip>" & vbNewLine & _
                        "<db_instance>" & txt(2).Text & "</db_instance>" & vbNewLine & _
                        "<db_owner>" & txt(3).Text & "</db_owner>" & vbNewLine & _
                        "<port>" & txt(5).Text & "</port>" & vbNewLine & _
                        "</root>"
        End If

        InitVsfSetup
        mintPage = mintPage + 1
        Call ShowPage(mintPage)
    Case 4
        '验证连接，是否能连接消息服务器
        Set mclsMipClientManage = Nothing
        If mclsMipClientManage Is Nothing Then
            Set mclsMipClientManage = New clsMipClientManage
        End If
        
        If mclsMipClientManage.CommunicateProxyLogin(txt(6).Text, txt(7).Text, txt(9).Text, txt(8).Text) = False Then
            mblnImportDB = False
            Exit Sub
        Else
            mblnImportDB = True
        End If
        
        mintPage = mintPage + 1
        Call ShowPage(mintPage)
    End Select
    
    
End Sub

Private Sub InitVsfSetup()
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim strPath As String
    '初始化安装步骤
    
    strPath = Left(txt(0).Text, Len(txt(0).Text) - Len("zlSetup.ini"))
    With vsfStep
        .Rows = 1
        intLoop = 0
        For intRow = 1 To vsf.Rows - 1
            If Abs(vsf.TextMatrix(intRow, vsf.ColIndex("选择"))) = 1 Then
                If ((vsf.TextMatrix(intRow, vsf.ColIndex("标识")) <> "EMR") Or (vsf.TextMatrix(intRow, vsf.ColIndex("标识")) = "EMR" And mblnSpecialEMR = True)) And Val(vsf.TextMatrix(intRow, vsf.ColIndex("是否安装"))) = 0 Then
                    .Rows = .Rows + 1
                    .TextMatrix(intLoop + 1, .ColIndex("step")) = "装载" & vsf.TextMatrix(intRow, vsf.ColIndex("系统名称")) & "消息数据"
                    .Cell(flexcpPicture, intLoop + 1, .ColIndex("图标"), intLoop + 1, .ColIndex("图标")) = imgList.ListImages("待执行").Picture
                    intLoop = intLoop + 1
                End If
            End If
        Next
        
        For intRow = 1 To vsf.Rows - 1
            If Abs(vsf.TextMatrix(intRow, vsf.ColIndex("选择"))) = 1 Then
                If ((vsf.TextMatrix(intRow, vsf.ColIndex("标识")) <> "EMR") Or (vsf.TextMatrix(intRow, vsf.ColIndex("标识")) = "EMR" And mblnSpecialEMR = True)) And Val(vsf.TextMatrix(intRow, vsf.ColIndex("是否安装"))) = 0 Then
                    If Dir(strPath & "\" & vsf.TextMatrix(intRow, vsf.ColIndex("标识")) & "\zlMipServerData.db") <> "" Then
                        .Rows = .Rows + 1
                        .TextMatrix(intLoop + 1, .ColIndex("step")) = "导入" & vsf.TextMatrix(intRow, vsf.ColIndex("系统名称")) & "消息数据"
                        .Cell(flexcpPicture, intLoop + 1, .ColIndex("图标"), intLoop + 1, .ColIndex("图标")) = imgList.ListImages("待执行").Picture
                        intLoop = intLoop + 1
                    End If
                End If
            End If
        Next
    End With
End Sub


Private Sub cmdPrevious_Click()

    Select Case mintPage
    '------------------------------------------------------------------------------------------------------------------
    Case 2, 3, 5
        
        mintPage = mintPage - 1
        Call ShowPage(mintPage)
    Case 4
        '判断是否选择了病历数据
        If mblnSpecialEMR = True Then
            mintPage = mintPage - 1
        Else
            mintPage = mintPage - 2
        End If
        Call ShowPage(mintPage)
    End Select
    
End Sub

Private Sub cmdSetup_Click()
    
    If MsgBox("确定需要安装消息集成平台客户端吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
     '安装脚本
    If SetupMipClient(txt(0).Text) Then
        MsgBox "消息数据安装成功!", vbInformation + vbOKOnly, "提示信息"
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strValue As String
    Dim strArr() As String
    
    '读注册表
    strValue = GetSetting("ZLSOFT", "公共全局\消息平台客户端", "EMR连接", "")
    If strValue <> "" Then
        strArr = Split(strValue, ";")
        txt(1).Text = strArr(0)
        txt(2).Text = strArr(1)
        txt(5).Text = strArr(2)
        txt(3).Text = strArr(3)
    End If
    
    strValue = GetSetting("ZLSOFT", "公共全局\消息平台客户端", "消息服务器连接", "")
    If strValue <> "" Then
        strArr = Split(strValue, ";")
        txt(6).Text = strArr(0)
        txt(7).Text = strArr(1)
        txt(9).Text = strArr(2)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Not (mclsOracle Is Nothing) Then
'        Set mclsOracle = Nothing
'    End If
'
'    Dim frmThis As Form
'
'    On Error Resume Next
'
'    '关闭本部件窗体
'    For Each frmThis In Forms
'        If frmThis.Caption <> Me.Caption Then
'            Unload frmThis
'        End If
'    Next
'
    Unload Me
    
    If Not (mcolSigns Is Nothing) Then
        Set mcolSigns = Nothing
    End If
End Sub

Private Sub mclsMipClientManage_AfterCommunicateChange(ByVal strTitle As String, ByVal strNumber As String)
    lbl(12).Caption = strTitle
    lbl(6).Caption = strNumber & "%"
    pgb.Value = strNumber
End Sub

Private Sub mobjScript_AfterAnalyseLine(ByVal Line As Long, ByVal Lines As Long)
    Dim intPercent As Integer
    
'    If pgb.Visible = False Then pgb.Visible = True
'    If lbl(4).Visible = False Then
'        lbl(4).Visible = True
'        lbl(4).Caption = "正在分析脚本文件...."
'    End If
'
'    intPercent = 100 * Line / Lines
'    If pgb.Value <> intPercent Then pgb.Value = intPercent
'
End Sub

Private Sub mobjScript_BeforeAnalyseLine(ByVal Line As Long, ByVal Lines As Long)
'    If pgb.Visible = False Then pgb.Visible = True
'    If lbl(4).Visible = False Then
'        lbl(4).Visible = True
'        lbl(4).Caption = "正在分析脚本文件...."
'    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
End Sub

Private Sub vsf_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
    
    '判断当前选中行是否已安装
    If Val(vsf.TextMatrix(Row, vsf.ColIndex("是否安装"))) = 1 Then
        Cancel = True
    End If
End Sub

Private Sub vsf_Click()
    If vsf.MouseRow = 0 And vsf.Col = vsf.ColIndex("选择") Then
        Call SelectedAll
    End If
End Sub
