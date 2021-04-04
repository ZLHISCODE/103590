VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicDelBalance 
   Caption         =   "病人退费结算"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicDelBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10365
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "返回(&X)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8364
      TabIndex        =   27
      Top             =   900
      Width           =   1704
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   36
      ScaleHeight     =   3090
      ScaleWidth      =   7995
      TabIndex        =   16
      Top             =   0
      Width           =   7995
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1260
         Left            =   45
         ScaleHeight     =   1230
         ScaleWidth      =   3060
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1572
         Width           =   3090
         Begin VB.Label lbl退费合计 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   612
            Left            =   2028
            TabIndex        =   22
            Top             =   552
            Width           =   1008
         End
         Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption2 
            Height          =   420
            Left            =   15
            TabIndex        =   21
            Top             =   30
            Width           =   3045
            _Version        =   589884
            _ExtentX        =   5371
            _ExtentY        =   741
            _StockProps     =   6
            Caption         =   "退费合计"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   15.76
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   2736
         Left            =   3204
         ScaleHeight     =   2700
         ScaleWidth      =   4710
         TabIndex        =   18
         Top             =   90
         Width           =   4740
         Begin VB.ComboBox cbo支付方式 
            BackColor       =   &H8000000F&
            ForeColor       =   &H8000000D&
            Height          =   408
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   204
            Width           =   1245
         End
         Begin VB.TextBox txt结算号码 
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   1368
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1332
            Width           =   3225
         End
         Begin VB.TextBox txt缴款 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   468
            IMEMode         =   3  'DISABLE
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   2
            Top             =   183
            Width           =   1920
         End
         Begin VB.TextBox txt摘要 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1368
            MaxLength       =   50
            TabIndex        =   8
            Top             =   1992
            Width           =   3210
         End
         Begin VB.TextBox txt找补 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   1368
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   744
            Width           =   3225
         End
         Begin VB.Label lbl结算号码 
            AutoSize        =   -1  'True
            Caption         =   "结算号码"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   108
            TabIndex        =   5
            Top             =   1416
            Width           =   1260
         End
         Begin VB.Label lbl找补 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "找　补"
            Height          =   312
            Left            =   372
            TabIndex        =   3
            Top             =   828
            Width           =   996
         End
         Begin VB.Label lbl摘要 
            AutoSize        =   -1  'True
            Caption         =   "摘  要"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   396
            TabIndex        =   7
            Top             =   2064
            Width           =   960
         End
         Begin VB.Label lblPayType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退  款"
            Height          =   312
            Left            =   384
            TabIndex        =   0
            Top             =   240
            Width           =   984
         End
      End
      Begin VB.PictureBox picTotal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1308
         Left            =   48
         ScaleHeight     =   1275
         ScaleWidth      =   3060
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   90
         Width           =   3090
         Begin XtremeSuiteControls.ShortcutCaption stcCurDelTitle 
            Height          =   450
            Left            =   15
            TabIndex        =   19
            Top             =   30
            Width           =   3045
            _Version        =   589884
            _ExtentX        =   5371
            _ExtentY        =   794
            _StockProps     =   6
            Caption         =   "当前应退"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   15.76
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label lbl未退金额 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   612
            Left            =   2016
            TabIndex        =   9
            Top             =   588
            Width           =   1008
         End
      End
   End
   Begin VB.Frame fraSplitLeft 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   8100
      TabIndex        =   13
      Top             =   -84
      Width           =   30
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   6120
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3572
            MinWidth        =   882
            Picture         =   "frmClinicDelBalance.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   2
            Object.Tag             =   "用于收费预交余额显示"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   1
            Object.Tag             =   "用于收费三方卡余额的显示"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmClinicDelBalance.frx":115E
            Key             =   "Calc"
            Object.ToolTipText     =   "计算器:ALT+?"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1693
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1693
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBlance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2868
      Left            =   84
      ScaleHeight     =   2835
      ScaleWidth      =   10140
      TabIndex        =   15
      Top             =   3096
      Width           =   10176
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8730
         TabIndex        =   23
         Top             =   75
         Width           =   1080
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBlance 
         Height          =   2295
         Left            =   15
         TabIndex        =   12
         Top             =   495
         Width           =   9930
         _cx             =   17515
         _cy             =   4048
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmClinicDelBalance.frx":1838
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
      Begin VB.Label lbl已退合计 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "已付合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4305
         TabIndex        =   11
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label lblDeledInfor 
         Caption         =   "本次已退情况"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   98
         Width           =   2145
      End
   End
   Begin VB.PictureBox pic误差 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   8232
      ScaleHeight     =   1140
      ScaleWidth      =   2040
      TabIndex        =   24
      Top             =   1656
      Width           =   2040
      Begin VB.Label lbl误差额 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0111"
         Height          =   285
         Left            =   135
         TabIndex        =   26
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label lbl误差 
         Caption         =   "本次误差"
         Height          =   315
         Left            =   105
         TabIndex        =   25
         Top             =   90
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8376
      TabIndex        =   28
      Top             =   210
      Width           =   1716
   End
End
Attribute VB_Name = "frmClinicDelBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------
'程序入口相关变量
Public Enum gChargeDelType
    EM_FUN_退费 = 0
    EM_FUN_重退 = 1
End Enum
Private mobjDelBalance As clsCliniDelBalance
Private mbytFunc As gChargeDelType  '0-收费;1-作废
Private mfrmMain As frmClinicDelAndView
Private mcllDelPro As Collection
Private mlngModule As Long, mstrPrivs As String
Private mcllForceDelToCash As Collection '强制退现信息：Array(操作员,卡类别名称)
'------------------------------------------------------------------------------------------
Private mrsBalance As ADODB.Recordset '当前结算数据
Private mstr退支票 As String
Private mblnSingleBalance As Boolean  '除医保结算方式以外，是否只使用了一种结算方式
    
Private mobjPayCards As Cards
Private mblnNotClick  As Boolean '不触发点击事件
Private mblnOK As Boolean
Private mblnUnloaded  As Boolean
Private mblnLoad As Boolean
Private mlngR  As Long
'------------------------------------------------------------------------------------------
'局部变量
Private mblnFirst As Boolean
Private mblnUnLoad As Boolean '是否Unload窗体
Private mbln已报价 As Boolean
Private mcur个帐余额 As Currency
Private mlngPre支付方式 As Long
'----------------------------------------------------------------------------------------------
'医保相关
'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    不提醒缴款金额不足 As Boolean    '27536
    门诊连续收费 As Boolean
    分币处理 As Boolean
End Type
Private mInsurePara As TYPE_MedicarePAR

Private Type TY_BrushCard    '刷卡类型
    str卡号 As String
    str密码 As String
    str交易流水号 As String    '交易流水号
    str交易说明  As String     '交易信息
    str扩展信息 As String    '交易的扩展信息
    dbl帐户余额 As Double
End Type
Private mCurBrushCard As TY_BrushCard   '当前的刷卡信息
Private Type TY_ChargeMoney
    dbl退费合计 As Double
    dbl已算误差 As Double
    dbl本次应收 As Double
    dbl本次医保退费 As Double
    dbl已退合计 As Double
    dbl本次退预交  As Double
    dbl当前未退 As Double
    dbl预交余额 As Double
    dbl费用余额 As Double
    dbl可用预交 As Double
    dbl应缴累计 As Double
    dbl本次误差费 As Double
End Type
Private mCurCarge As TY_ChargeMoney
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsOneCard As ADODB.Recordset
Private mrsUsedCards As ADODB.Recordset

Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:是否缓存了回车键,可能存在在收费界面刷卡中本身包含了回车,因此需要判断
Private mrsClassMoney As ADODB.Recordset
Private mcllSquareBalance As Collection '消费卡退费结算信息
Private mcllSquareChargeBalance As Collection '消费卡收费结算信息
Private mcllCurSquareBalance As Collection '当前消费卡刷卡信息
Private mblnNotChange As Boolean
Private mstrTittle As String
Private mblnTurnFee As Boolean

Public Function zlDelCharge(ByVal frmMain As Object, _
    ByVal bytFunc As gChargeDelType, _
    ByVal lngModule As Long, ByVal strPrivs As String, objDelBalance As clsCliniDelBalance, _
    ByVal cllDelPro As Collection, Optional ByVal strDefault结算方式 As String = "", _
    Optional ByVal cllForceDelToCash As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口:表示进入退费结算窗口
    '入参:frmMain-调用的主窗体
    '       bytFunc-0- 退费;1-退异常退费
    '       lngModule -模块号
    '       strPrivs-权限串
    '       objDelBalance-退费相关结算信息
    '       cllDelPro-退费前需要执行的SQL
    '       strDefault结算方式-缺省的结算方式
    '       cllForceDelToCash 强制退现信息：Array(操作员,卡类别名称)
    '出参:
    '返回:完成收费,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-12 09:59:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mobjDelBalance = objDelBalance: Set mcllDelPro = cllDelPro
    mblnOK = False
    mblnUnLoad = False: mblnUnloaded = False
    mblnTurnFee = IsTurnFee(mobjDelBalance.AllNos)
    mstrPrivs = strPrivs: mlngModule = lngModule
    Set mfrmMain = frmMain
    mbytFunc = bytFunc
    mblnOK = False
    If cllForceDelToCash Is Nothing Then Set cllForceDelToCash = New Collection
    Set mcllForceDelToCash = cllForceDelToCash
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    Set objDelBalance = mobjDelBalance
    zlDelCharge = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsTurnFee(ByVal strNos As String) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select 1" & vbNewLine & _
            " From 费用审核记录 A, 门诊费用记录 B" & vbNewLine & _
            " Where a.费用id = b.Id And b.No In (Select Column_Value From Table(f_Str2list([1])))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查是否为门诊转住院费用", strNos)
    If rsTmp.EOF Then
        IsTurnFee = False
    Else
        IsTurnFee = True
    End If
End Function

Private Sub initInsure()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2011-08-21 18:55:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDelBalance.intInsure = 0 Then Exit Sub
    mInsurePara.门诊连续收费 = gclsInsure.GetCapability(support门诊连续收费, mobjDelBalance.病人ID, mobjDelBalance.intInsure)
    '刘兴洪:27536 20100119
    mInsurePara.不提醒缴款金额不足 = gclsInsure.GetCapability(support不提醒缴款金额不足, mobjDelBalance.病人ID, mobjDelBalance.intInsure)
    mInsurePara.分币处理 = gclsInsure.GetCapability(support分币处理, mobjDelBalance.病人ID, mobjDelBalance.intInsure)
End Sub

Private Sub InitBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算数据
    '编制:刘兴洪
    '日期:2012-02-05 16:02:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearBanalce
    With mCurCarge
        .dbl退费合计 = mobjDelBalance.退费合计
        .dbl已算误差 = 0
        .dbl本次医保退费 = 0
        .dbl已退合计 = 0
        .dbl当前未退 = .dbl退费合计 - .dbl本次医保退费
        .dbl本次退预交 = 0
        .dbl本次误差费 = 0
    End With
    
    '加载原样退
    Call Load原样退
End Sub

Public Function IsSingleBalance(ByVal lng原结帐ID As Long) As Boolean
    '判断单据第一次结帐除医保结算方式外是否只使用了一种结算方式
    '入参：
    '   lng原结帐ID - 原结帐ID，医保部分退的是上一次重收的结帐ID
    Dim rsBalance As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    '1.第一次结帐的结帐ID
    strSQL = "Select Distinct m.结帐id" & vbNewLine & _
            " From 门诊费用记录 M, 门诊费用记录 N" & vbNewLine & _
            " Where Mod(m.记录性质, 10) = Mod(n.记录性质, 10) And m.No = n.No" & vbNewLine & _
            "       And m.记录性质 = 1 And m.记录状态 In (1, 3) And n.结帐id = [1]"
    '2.第一次结算的结算信息
    strSQL = "With 原始结帐id As(" & strSQL & ")" & vbNewLine & _
            " Select Decode(a.记录性质, 11, '冲预交', a.结算方式) As 结算方式, a.冲预交, c.性质" & vbNewLine & _
            " From 病人预交记录 A, 原始结帐id B, 结算方式 C" & vbNewLine & _
            " Where a.结帐id = b.结帐id And a.记录性质 In (11, 3) And a.结算方式 = c.名称(+) And c.应收款 <> 1 And c.应付款 <> 1"
    '3.除医保结算方式以外的结算方式
    '结算方式.性质：1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算(老版),8-结算卡结算(新版),9-误差费
    strSQL = "Select Distinct 结算方式 From (" & strSQL & ") Where 性质 Not In (3, 4, 9)"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, "获取非医保的结算方式", lng原结帐ID)
    IsSingleBalance = (rsBalance.RecordCount <= 1)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Load原样退()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载原样退的结算方式
    '编制:刘兴洪
    '日期:2014-07-31 14:16:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsBalance As ADODB.Recordset
    Dim strTemp As String, lngCurRecord As Long
    Dim i As Long, strCardNo As String, lng卡类别ID As Long
    Dim j As Long, ingCount As Integer
    Dim blnFind As Boolean
    Dim bln普通结算 As Boolean
    Dim varTemp As Variant, lng消费卡ID As Long
    Dim objCard As Card, dblMoney As Double
    
    On Error GoTo errHandle
    '原样退的话,先将退费的金额显示出来
    '注意：如果一卡通或消费卡不能退现，同时该卡未启用则按非单种结算方式处理（104555）
    Set rsBalance = mobjDelBalance.rsBalance
    If rsBalance Is Nothing Then Exit Sub
    If rsBalance.State <> 1 Then Exit Sub
    
    mblnSingleBalance = IsSingleBalance(mobjDelBalance.原结帐ID)
    
    If mblnSingleBalance Then
        rsBalance.Filter = "类型<>2 And 结算性质 <> 9 And 退费=0 And 结算方式<>'" & mstr退支票 & "'"
        rsBalance.Sort = "ID Asc"
        If rsBalance.RecordCount > 0 Then
            rsBalance.MoveFirst
            If RoundEx(mCurCarge.dbl当前未退, 6) <> 0 Then
                '未退金额为零缺省其它结算方式没有意义
                If Val(Nvl(rsBalance!类型)) = 1 Then
                    mobjDelBalance.缺省结算方式 = "退预存款"
                Else
                    mobjDelBalance.缺省结算方式 = Trim(Nvl(rsBalance!结算方式))
                End If
            End If
            '3-一卡通 不能退现且不是全退的不能切换退款方式，不能退现全退的已在进入结算界面前结算
            '5-消费卡，不能退现的不能切换退款方式
            If Val(Nvl(rsBalance!类型)) = 3 Or Val(Nvl(rsBalance!类型)) = 5 Then
                If Val(Nvl(rsBalance!类型)) = 3 Then
                    Set objCard = GetPayCard(Val(Nvl(rsBalance!卡类别ID)), False)
                Else
                    Set objCard = GetPayCard(Val(Nvl(rsBalance!结算卡序号)), True)
                End If
                If Not objCard Is Nothing Then
                    If Val(Nvl(rsBalance!类型)) = 5 Then
                        cbo支付方式.Enabled = (Val(Nvl(rsBalance!是否退现)) = 1)
                    End If
                    '如果不退现，金额还未退完，则不能编辑退款方式
                    dblMoney = GetOldBalanceMoney(Val(Nvl(rsBalance!类型)), objCard)
                    If RoundEx(dblMoney, 6) = 0 Then cbo支付方式.Enabled = True
                End If
            End If
        End If
    Else
        '77873,冉俊明,2014-9-15
        If mobjDelBalance.部分退费 Then
            '加载必须全退的消费卡和一卡通部分
            '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
            '     是否密文,是否全退,是否退现,冲预交
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
            rsBalance.Filter = "(是否退现=0 AND 类型=3) or (是否退现=0 AND 类型=5) OR 是否全退=1"
            rsBalance.Sort = "是否退现 asc,是否全退 desc"
            If rsBalance.RecordCount = 0 Then
                rsBalance.Filter = 0
                Exit Sub
            End If
        Else
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            rsBalance.Filter = "类型=0 and 结算性质=1"
            If Not rsBalance.EOF Then
                 mobjDelBalance.缺省结算方式 = Trim(Nvl(rsBalance!结算方式))
            End If
            rsBalance.Filter = "类型<>2 and 类型<>4 "
            rsBalance.Sort = "类型 desc,结算性质 desc"
        End If
    
        If rsBalance.RecordCount <> 0 Then rsBalance.MoveFirst
        With rsBalance
            bln普通结算 = False
            lngCurRecord = 1
            mobjDelBalance.缺省结算方式 = ""
            Do While Not .EOF
                strTemp = ""
                '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                Select Case Val(Nvl(!类型))
                Case 0 '普通结算
                    If InStr(gTy_Module_Para.str缺省退现, Trim(Nvl(!结算方式))) = 0 Then
                        strTemp = Trim(Nvl(!结算方式))
                    End If
                    bln普通结算 = True
                Case 1 '预交款
                    strTemp = "退预存款"
                    mCurCarge.dbl本次退预交 = RoundEx(mCurCarge.dbl本次退预交 + Val(Nvl(!冲预交)), 6)
                Case 2 '医保,不处理
                    '医保已经在退费前处理
                Case 3 '一卡通
                    Set objCard = GetPayCard(Val(Nvl(!卡类别ID)), False)
                    If objCard Is Nothing Then
                        strTemp = "" '卡未启用已在前面判断
                    ElseIf Not (objCard.是否退现 And objCard.是否缺省退现) Then
                        strTemp = Trim(Nvl(!结算方式))
                    End If
                Case 4 '一卡通(老)
                Case 5  '消费卡
                    strTemp = Trim(Nvl(!结算方式))
                End Select
                
                If Val(Nvl(!类型)) = 0 And Val(Nvl(!结算性质)) = 1 Then strTemp = "" '现金不加
                If Val(Nvl(rsBalance!结算性质)) = 9 Then strTemp = ""   '误差费不加入
                If strTemp <> "" Then
                    With vsBlance
                        i = 1
                        If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
                             blnFind = False
                            If Nvl(rsBalance!类型) = 1 Then
                                For j = 1 To .Rows - 1
                                    If Val(.TextMatrix(j, .ColIndex("类型"))) = 1 Then
                                        blnFind = True
                                        i = j: Exit For
                                    End If
                                Next
                            ElseIf Nvl(rsBalance!类型) = 5 Then
                                For j = 1 To .Rows - 1
                                    If strTemp = Trim(.TextMatrix(j, .ColIndex("支付方式"))) _
                                        And Val(Nvl(rsBalance!消费卡ID)) = Val(.TextMatrix(j, .ColIndex("消费卡ID"))) Then
                                        blnFind = True
                                        i = j: Exit For
                                    End If
                                Next
                            Else
                                For j = 1 To .Rows - 1
                                    If strTemp = Trim(.TextMatrix(j, .ColIndex("支付方式"))) Then
                                        blnFind = True
                                        i = j: Exit For
                                    End If
                                Next
                            
                            End If
                            If Not blnFind Then
                                .Rows = .Rows + 1
                                .RowPosition(.Rows - 1) = 1
                            End If
                        End If
                        
                        If Not (Val(.TextMatrix(i, .ColIndex("结算状态"))) = 1 And blnFind) Then  '是否已结算:1-已结算;0-未结算
                            strCardNo = Nvl(rsBalance!卡号)
                            If Nvl(rsBalance!类型) = 5 Then
                                lng卡类别ID = Val(Nvl(rsBalance!结算卡序号))
                                If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
                                'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文,剩余未退金额)
                                mcllSquareBalance.Add Array(lng卡类别ID, Val(Nvl(rsBalance!消费卡ID)), _
                                 0, strCardNo, "", "", Val(Nvl(rsBalance!是否密文)), Format(Val(Nvl(rsBalance!冲预交)), "0.00"))
                            Else
                                lng卡类别ID = Val(Nvl(rsBalance!卡类别ID))
                            End If
                            .RowData(i) = Nvl(rsBalance!类型)
                            .TextMatrix(i, .ColIndex("类型")) = Val(Nvl(rsBalance!类型))
                            .TextMatrix(i, .ColIndex("结算性质")) = Val(Nvl(rsBalance!结算性质))
                            If Nvl(rsBalance!类型) = 5 Then
                                .TextMatrix(i, .ColIndex("删除标志")) = IIf(Val(Nvl(rsBalance!是否退现)) = 1, 0, 1) '是否允许编辑:1-禁止编辑;0-不禁止编辑
                            Else
                                .TextMatrix(i, .ColIndex("删除标志")) = 0 '是否允许编辑:1-禁止编辑;0-不禁止编辑
                            End If
                            .TextMatrix(i, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
                            .TextMatrix(i, .ColIndex("卡类别ID")) = lng卡类别ID
                            .TextMatrix(i, .ColIndex("消费卡ID")) = Val(Nvl(rsBalance!消费卡ID))
                            .TextMatrix(i, .ColIndex("支付方式")) = strTemp
                            ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                            .Cell(flexcpData, i, .ColIndex("支付方式")) = lng卡类别ID & "|" & IIf(Val(Nvl(rsBalance!类型)) = 5, 1, 0) & "|" & Val(Nvl(rsBalance!自制卡)) & "|" & Val(Nvl(rsBalance!是否全退)) & "|" & Val(Nvl(rsBalance!是否退现)) & "|" & Nvl(rsBalance!卡类别名称)
                            .TextMatrix(i, .ColIndex("支付金额")) = FormatEx(-1 * Val(.Cell(flexcpData, i, .ColIndex("支付金额"))) + Val(Nvl(rsBalance!冲预交)), 6, , , 2)
                            .Cell(flexcpData, i, .ColIndex("支付金额")) = FormatEx(Val(.Cell(flexcpData, i, .ColIndex("支付金额"))) + -1 * Val(Nvl(rsBalance!冲预交)), 6)
                            If Nvl(rsBalance!类型) <> 1 Then '预交款不显示结算号码、摘要、卡号、交易流水号、交易说明
                                .TextMatrix(i, .ColIndex("结算号码")) = Nvl(rsBalance!结算号码)
                                .TextMatrix(i, .ColIndex("备注")) = Nvl(rsBalance!摘要)
                                .TextMatrix(i, .ColIndex("交易流水号")) = Nvl(rsBalance!交易流水号)
                                .TextMatrix(i, .ColIndex("交易说明")) = Nvl(rsBalance!交易说明)
                                .TextMatrix(i, .ColIndex("卡号")) = IIf(Val(Nvl(rsBalance!是否密文)) = 1, String(Len(strCardNo), "*"), strCardNo)
                                .TextMatrix(i, .ColIndex("是否退现")) = Val(Nvl(rsBalance!是否退现))
                                .TextMatrix(i, .ColIndex("是否全退")) = Val(Nvl(rsBalance!是否全退))
                                .TextMatrix(i, .ColIndex("是否转帐及代扣")) = Val(Nvl(rsBalance!是否转帐及代扣))
                                .TextMatrix(i, .ColIndex("卡类别名称")) = Nvl(rsBalance!卡类别名称)
                                .Cell(flexcpData, i, .ColIndex("卡号")) = Nvl(rsBalance!卡号)
                            End If
                            mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 + -1 * Val(Nvl(rsBalance!冲预交)), 6)
                        End If
                    End With
                End If
                .MoveNext
            Loop
            mCurCarge.dbl当前未退 = RoundEx(mCurCarge.dbl退费合计 - mCurCarge.dbl已退合计, 6)
            
            '77873,冉俊明,2014-9-15
            '85597,部分退费后，将剩余部分全退时支付方式默认金额不正确
            '86248,部分退费时退为支票，第二次将剩余部分全退时不应该默认为收支票
            With vsBlance
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                        If .Cell(flexcpData, i, .ColIndex("支付金额")) >= 0 Then '收款都不默认显示
                            mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 - .Cell(flexcpData, i, .ColIndex("支付金额")), 6)
                            mCurCarge.dbl当前未退 = RoundEx(mCurCarge.dbl当前未退 + .Cell(flexcpData, i, .ColIndex("支付金额")), 6)
                            .TextMatrix(i, .ColIndex("支付金额")) = 0
                            .Cell(flexcpData, i, .ColIndex("支付金额")) = 0
                        End If
                    End If
                Next
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                        If mCurCarge.dbl当前未退 > 0 Then
                            '93114,全退可退现的可转帐的一卡通，则缺省为当前退款金额
                            If Val(.TextMatrix(i, .ColIndex("是否全退"))) = 1 _
                                And (Val(.TextMatrix(i, .ColIndex("类型"))) <> 3 _
                                    Or (Val(.TextMatrix(i, .ColIndex("类型"))) = 3 _
                                        And (Val(.TextMatrix(i, .ColIndex("是否退现"))) = 0 _
                                            Or Val(.TextMatrix(i, .ColIndex("是否转帐及代扣"))) = 0))) Then
                                Exit For
                            End If
                            If mCurCarge.dbl当前未退 > -1 * .Cell(flexcpData, i, .ColIndex("支付金额")) Then
                                mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 - .Cell(flexcpData, i, .ColIndex("支付金额")), 6)
                                mCurCarge.dbl当前未退 = RoundEx(mCurCarge.dbl当前未退 + .Cell(flexcpData, i, .ColIndex("支付金额")), 6)
                                .TextMatrix(i, .ColIndex("支付金额")) = 0
                                .Cell(flexcpData, i, .ColIndex("支付金额")) = 0
                            Else
                                If .TextMatrix(i, .ColIndex("支付金额")) <> 0 Then
                                    '特殊处理：收费时保存的结算金额最多两位小数,因此此处退款金额也要先四舍五入处理
                                    '如当前未退30.105，收款时支付金额为30.11，如果不先进行四舍五入处理，就会认为收款时是30.105
                                    mCurCarge.dbl已算误差 = RoundEx(mCurCarge.dbl退费合计 - Format(mCurCarge.dbl退费合计, "0.00"), 6)
                                    
                                    .TextMatrix(i, .ColIndex("支付金额")) = Format(.TextMatrix(i, .ColIndex("支付金额")) - (mCurCarge.dbl当前未退 - mCurCarge.dbl已算误差), "0.00")
                                    .Cell(flexcpData, i, .ColIndex("支付金额")) = Format(.Cell(flexcpData, i, .ColIndex("支付金额")) + (mCurCarge.dbl当前未退 - mCurCarge.dbl已算误差), "0.00")
                                    mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 + (mCurCarge.dbl当前未退 - mCurCarge.dbl已算误差), 6)
                                    mCurCarge.dbl当前未退 = mCurCarge.dbl已算误差
                                    Exit For
                                End If
                            End If
                        Else
                            '93114,支持转帐且可退现的一卡通，缺省为所有的可转帐金额
                            '排除金额为零的，金额为零表示要移除的
    '                        If Val(.TextMatrix(i, .ColIndex("类型"))) = 3 And Val(.TextMatrix(i, .ColIndex("是否转帐及代扣"))) = 1 _
    '                            And Val(.TextMatrix(i, .ColIndex("是否退现"))) = 1 And .TextMatrix(i, .ColIndex("支付金额")) <> 0 Then
    '                            .TextMatrix(i, .ColIndex("支付金额")) = Format(.TextMatrix(i, .ColIndex("支付金额")) - mCurCarge.dbl当前未退, "0.00")
    '                            .Cell(flexcpData, i, .ColIndex("支付金额")) = .Cell(flexcpData, i, .ColIndex("支付金额")) + mCurCarge.dbl当前未退
    '                            mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 + mCurCarge.dbl当前未退, 6)
    '                            mCurCarge.dbl当前未退 = 0
    '                        End If
                        End If
                    End If
                Next
                
                i = 1
                Do While True
                    If Val(.TextMatrix(i, .ColIndex("支付金额"))) = 0 Then
                        '移除金额为零的支付类别数据未退
                        lng卡类别ID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                        lng消费卡ID = Val(.TextMatrix(i, .ColIndex("消费卡ID")))
                        Call ClearReMoveSquareBalance(lng卡类别ID, lng消费卡ID)
                        If .Rows <= 2 Then
                            .Rows = 2
                            .Clear 1
                            .RowData(1) = ""
                            .Cell(flexcpData, 1, 0, .Rows - 1, .COLS - 1) = ""
                            Exit Do
                        Else
                            .RemoveItem i
                        End If
                    Else
                        i = i + 1
                    End If
                    If i > .Rows - 1 Then Exit Do
                Loop
            End With
        End With
    End If
    mobjDelBalance.rsBalance.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub ClearBanalce()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除结算数据
    '编制:刘兴洪
    '日期:2012-02-05 16:02:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurCarge
        .dbl退费合计 = 0
        .dbl已算误差 = 0
        .dbl本次医保退费 = 0
        .dbl已退合计 = 0
        .dbl本次应收 = 0
        .dbl当前未退 = 0
        .dbl本次退预交 = 0
        .dbl本次误差费 = 0
    End With
    With vsBlance
        .Clear 1: .Rows = 2
    End With
End Sub
Private Sub LoadData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '编制:刘兴洪
    '日期:2011-08-20 19:49:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, bln消费卡 As Boolean, lng卡类别ID As Long
    Dim strCardNo As String
    Dim blnYb As Boolean
    Dim dbl退预交款 As Double
    
    On Error GoTo errHandle
    
    Call ClearBanalce
    If mobjDelBalance.SaveBilled = False Then Exit Sub
    
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    'bytType-查找类型:0-根据结帐ID查找;1-根据结算序号查找
    Set mrsBalance = zlFromIDGetChargeBalance(1, mobjDelBalance.结算序号, False, True)
    mrsBalance.Filter = 0
    mrsBalance.Sort = "类型,结算方式"
    If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    With mrsBalance
        i = 1: blnYb = False
        dbl退预交款 = 0
        Do While Not .EOF
            Select Case Nvl(!类型)
            Case 1 '预交款
                mCurCarge.dbl本次退预交 = RoundEx(mCurCarge.dbl本次退预交 + Val(Nvl(!冲预交)), 6)
                mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 + Val(Nvl(!冲预交)), 6)
                dbl退预交款 = RoundEx(dbl退预交款 + Val(Nvl(!冲预交)), 6)
            Case 2, 3, 5 '医保,一卡通,消费卡
                If Nvl(!类型) = 2 Then
                    mCurCarge.dbl本次医保退费 = RoundEx(mCurCarge.dbl本次医保退费 + Nvl(!冲预交, 0), 6)
                    blnYb = True
                End If
'                If Val(Nvl(mrsBalance!校对标志, 0)) = 2 Then
                    With vsBlance
                        If .TextMatrix(i, .ColIndex("支付方式")) <> "" Then
                            .Rows = .Rows + 1
                            i = i + 1
                        End If
                        .RowData(i) = Nvl(mrsBalance!类型)
                        If Nvl(mrsBalance!类型) = 5 Then
                            lng卡类别ID = Val(Nvl(mrsBalance!结算卡序号))
                        Else
                            lng卡类别ID = Val(Nvl(mrsBalance!卡类别ID))
                        End If
                        
                        strCardNo = Nvl(mrsBalance!卡号)
                        If Nvl(mrsBalance!类型) = 5 Then
                            If Val(Nvl(mrsBalance!冲预交)) <= 0 Then
                                If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
                                'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
                                mcllSquareBalance.Add Array(lng卡类别ID, Val(Nvl(mrsBalance!消费卡ID)), _
                                Format(Val(Nvl(mrsBalance!冲预交)), "0.00"), strCardNo, "", "", Val(Nvl(mrsBalance!是否密文)), Format(Val(Nvl(mrsBalance!冲预交)), "0.00"))
                            Else
                                If mcllSquareChargeBalance Is Nothing Then Set mcllSquareChargeBalance = New Collection
                                'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
                                mcllSquareChargeBalance.Add Array(lng卡类别ID, Val(Nvl(mrsBalance!消费卡ID)), _
                                Format(Val(Nvl(mrsBalance!冲预交)), "0.00"), strCardNo, "", "", Val(Nvl(mrsBalance!是否密文)), Format(Val(Nvl(mrsBalance!冲预交)), "0.00"))
                            End If
                        End If
                        .TextMatrix(i, .ColIndex("类型")) = Val(Nvl(mrsBalance!类型))
                        .TextMatrix(i, .ColIndex("结算性质")) = Val(Nvl(mrsBalance!结算性质))
                        .TextMatrix(i, .ColIndex("删除标志")) = 1  '是否允许编辑:1-禁止编辑;0-不禁止编辑
                        .TextMatrix(i, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
                        .TextMatrix(i, .ColIndex("卡类别ID")) = lng卡类别ID
                        .TextMatrix(i, .ColIndex("消费卡ID")) = Val(Nvl(mrsBalance!消费卡ID))

                        .TextMatrix(i, .ColIndex("支付方式")) = Nvl(mrsBalance!结算方式)
                        
                        ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                        .Cell(flexcpData, i, .ColIndex("支付方式")) = lng卡类别ID & "|" & IIf(Val(Nvl(mrsBalance!类型)) = 5, 1, 0) & "|" & Val(Nvl(mrsBalance!自制卡)) & "|" & Val(Nvl(mrsBalance!是否全退)) & "|" & Val(Nvl(mrsBalance!是否退现)) & "|" & Nvl(mrsBalance!卡类别名称)
                        .TextMatrix(i, .ColIndex("支付金额")) = Format(-1 * Val(Nvl(mrsBalance!冲预交)), "0.00")
                        .Cell(flexcpData, i, .ColIndex("支付金额")) = Format(Val(Nvl(mrsBalance!冲预交)), "0.00")
                        .TextMatrix(i, .ColIndex("结算号码")) = Nvl(mrsBalance!结算号码)
                        .TextMatrix(i, .ColIndex("备注")) = Nvl(mrsBalance!摘要)
                        .TextMatrix(i, .ColIndex("交易流水号")) = Nvl(mrsBalance!交易流水号)
                        .TextMatrix(i, .ColIndex("交易说明")) = Nvl(mrsBalance!交易说明)
                        .TextMatrix(i, .ColIndex("卡号")) = IIf(Val(Nvl(mrsBalance!是否密文)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .TextMatrix(i, .ColIndex("是否退现")) = Val(Nvl(mrsBalance!是否退现))
                        .TextMatrix(i, .ColIndex("是否全退")) = Val(Nvl(mrsBalance!是否全退))
                        .TextMatrix(i, .ColIndex("卡类别名称")) = Nvl(mrsBalance!卡类别名称)
  
                        
                        .Cell(flexcpData, i, .ColIndex("卡号")) = Nvl(mrsBalance!卡号)
                        .Cell(flexcpBackColor, i, 0, i, .COLS - 1) = Me.BackColor
                        mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 + Val(Nvl(mrsBalance!冲预交)), 6)
                    End With
'                End If
            Case Else '0-普通结算
                With vsBlance
                   If .TextMatrix(i, .ColIndex("支付方式")) <> "" And Nvl(mrsBalance!结算方式) <> "" Then
                       .Rows = .Rows + 1
                       i = i + 1
                   End If
                   If Trim(Nvl(mrsBalance!结算方式)) <> "" Then
                        .RowData(i) = Nvl(mrsBalance!类型)
                        
                        .TextMatrix(i, .ColIndex("类型")) = Val(Nvl(mrsBalance!类型))
                        .TextMatrix(i, .ColIndex("结算性质")) = Val(Nvl(mrsBalance!结算性质))
                        .TextMatrix(i, .ColIndex("删除标志")) = 1  '是否允许编辑:1-禁止编辑;0-不禁止编辑
                        .TextMatrix(i, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
                        .TextMatrix(i, .ColIndex("卡类别ID")) = 0
                        .TextMatrix(i, .ColIndex("消费卡ID")) = 0

                        
                        .TextMatrix(i, .ColIndex("支付方式")) = Nvl(mrsBalance!结算方式)
                        .TextMatrix(i, .ColIndex("支付金额")) = Format(-1 * Val(Nvl(mrsBalance!冲预交)), "0.00")
                        .Cell(flexcpData, i, .ColIndex("支付金额")) = Format(Val(Nvl(mrsBalance!冲预交)), "0.00")
                        .TextMatrix(i, .ColIndex("结算号码")) = Nvl(mrsBalance!结算号码)
                        .TextMatrix(i, .ColIndex("备注")) = Nvl(mrsBalance!摘要)
                        .TextMatrix(i, .ColIndex("交易流水号")) = Nvl(mrsBalance!交易流水号)
                        .TextMatrix(i, .ColIndex("交易说明")) = Nvl(mrsBalance!交易说明)
                        .TextMatrix(i, .ColIndex("卡号")) = IIf(Val(Nvl(mrsBalance!是否密文)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .Cell(flexcpData, i, .ColIndex("卡号")) = Nvl(mrsBalance!卡号)
                        .Cell(flexcpBackColor, i, 0, i, .COLS - 1) = Me.BackColor
                        mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 + Val(Nvl(mrsBalance!冲预交)), 6)
                    End If
                End With
            End Select
            .MoveNext
        Loop
    End With
    '先计算出退费合计
    gstrSQL = "" & _
    "   Select B.NO,B.结帐ID, Nvl(Sum(Nvl(B.应收金额, 0)), 0)  As 本次应收合计, " & _
    "       Nvl(Sum(Nvl(B.实收金额, 0)), 0)  As 本次实收合计 " & _
    "   From 门诊费用记录 B " & _
    "    Where B.结帐ID=[1] Or B.结帐ID=[2]" & _
    "    Group by B.NO,B.结帐ID"
   Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.冲销ID, mobjDelBalance.结帐ID)
   With mCurCarge
         .dbl退费合计 = 0:
         .dbl本次应收 = 0
        Do While Not rsTemp.EOF
            .dbl退费合计 = RoundEx(.dbl退费合计 + Val(Nvl(rsTemp!本次实收合计)), 6)
            .dbl本次应收 = RoundEx(.dbl本次应收 + Val(Nvl(rsTemp!本次应收合计)), 6)
            rsTemp.MoveNext
        Loop
        .dbl当前未退 = RoundEx(.dbl退费合计 - .dbl已退合计, 6)
    End With
    Call Load原样退
                   
    If dbl退预交款 <> 0 Then
        With vsBlance
            If .Rows = 2 Then .Row = 1
            If .Row < 0 Then .Row = 1
            i = .Row
            If Trim(.TextMatrix(.Row, .ColIndex("支付方式"))) <> "" Then
                .Rows = .Rows + 1
                i = .Rows - 1
            End If
            .RowData(i) = 1
            .TextMatrix(i, .ColIndex("删除标志")) = 1   ' 是否允许编辑:1-禁止编辑;0-不禁止编辑
            .TextMatrix(i, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
            .TextMatrix(i, .ColIndex("支付方式")) = "退预存款"
            .TextMatrix(i, .ColIndex("支付金额")) = Format(-1 * mCurCarge.dbl本次退预交, "0.00")
            .Cell(flexcpData, i, .ColIndex("支付金额")) = Format(mCurCarge.dbl本次退预交, "0.00")
            .TextMatrix(i, .ColIndex("类型")) = 1
            
            .Cell(flexcpBackColor, i, 0, i, .COLS - 1) = Me.BackColor
        End With
    End If
   vsBlance_AfterRowColChange 0, 0, vsBlance.Row, vsBlance.Col
   Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Init退费方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, objCard As Card, objCards As Cards
    Dim lngKey As Long
    
    Set mobjPayCards = New Cards
    Set objCards = New Cards
    
    Set rsTemp = mobjDelBalance.rs结算方式
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare Is Nothing Then
    ' zlGetCards(ByVal BytType As Byte)
        '入参:bytType-  0-所有医疗卡;
    '                        1-启用的医疗卡,
    '                        2-所有存在三方账户的三方卡
    '                        3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            blnFind = False
            For i = 1 To objCards.Count
                If objCards(i).结算方式 = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                If Not (Val(Nvl(rsTemp!性质)) = 3 Or Val(Nvl(rsTemp!性质)) = 4 _
                    Or Val(Nvl(rsTemp!性质)) = 7 Or Val(Nvl(rsTemp!性质)) = 8 _
                    Or Val(Nvl(rsTemp!应付款)) = 1) Then
                    
                    '不加入医保的结算方式或退支票的
                     Set objCard = New Card
                     objCard.短名 = Mid(Nvl(!名称), 1, 1)
                     objCard.接口编码 = Nvl(!编码)
                     objCard.接口程序名 = ""
                     objCard.接口序号 = -1 * lngKey
                     objCard.结算方式 = Nvl(!名称)
                     objCard.名称 = Nvl(!名称)
                     objCard.启用 = True
                     objCard.缺省标志 = Val(Nvl(rsTemp!缺省)) = 1
                     objCard.支付启用 = True
                     objCard.结算性质 = Val(!性质)
                     If objCard.结算性质 = 7 And objCard.接口序号 <= 0 Then   '一卡通未启用时,不加入
                        mrsOneCard.Filter = "结算方式='" & objCard.结算方式 & "'"
                        If Not mrsOneCard.EOF Then
                            mobjPayCards.Add objCard, "K" & lngKey
                            lngKey = lngKey + 1
                        End If
                     Else
                        mobjPayCards.Add objCard, "K" & lngKey
                        lngKey = lngKey + 1
                     End If
              End If
            End If
            .MoveNext
        Loop
    End With
    
    '加三方卡
    For i = 1 To objCards.Count
        rsTemp.Filter = "名称='" & objCards(i).结算方式 & "'" '结算方式要设置了"费用"应用场合才能使用
        If Not rsTemp.EOF Then
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If mobjPayCards.Count = 0 Then
        MsgBox "没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    
    '加制加入预交金额
     Set objCard = New Card
     objCard.短名 = "预"
     objCard.接口编码 = ""
     objCard.接口程序名 = ""
     objCard.接口序号 = -1 * lngKey
     objCard.结算方式 = "预交款"
     objCard.名称 = "预交款"
     objCard.启用 = True
     objCard.缺省标志 = False
     objCard.支付启用 = True
     objCard.结算性质 = "-99"
     mobjPayCards.Add objCard, "K" & lngKey
End Sub

Private Sub StartAndStop预存款()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前的退款金额，动态加载预交的支付方式
    '编制:刘兴洪
    '日期:2014-07-08 15:21:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsBalance As ADODB.Recordset
    Dim objCard As Card
    Dim blnStart As Boolean, i As Long, dblMoney As Double
    
    Set rsBalance = mobjDelBalance.rsBalance
    '退预存款
    '114528,当前未退只有误差金额时应该是退款，不应该出现收款
    If RoundEx(mCurCarge.dbl当前未退 - mCurCarge.dbl本次误差费, 6) <= 0 Then
        '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        rsBalance.Filter = "类型=1"
        If rsBalance.RecordCount > 0 Then
            dblMoney = 0
            Do While Not rsBalance.EOF
                dblMoney = dblMoney + Nvl(rsBalance!冲预交)
                rsBalance.MoveNext
            Loop
            dblMoney = RoundEx(dblMoney, 6)
            If RoundEx(dblMoney, 6) <> 0 Then
                For i = 1 To mobjPayCards.Count
                   Set objCard = mobjPayCards(i)
                   If objCard.结算性质 = -99 Then
                      objCard.结算方式 = "退预存款"
                      objCard.名称 = "退预存款"
                      objCard.支付启用 = True
                   End If
                Next
            End If
        End If
    Else '冲预存款
        For i = 1 To mobjPayCards.Count
           Set objCard = mobjPayCards(i)
           If objCard.结算性质 = -99 Then
              objCard.结算方式 = "冲预存款"
              objCard.名称 = "冲预存款"
              objCard.支付启用 = True
           End If
        Next
    End If
    
    blnStart = True
    With vsBlance
        For i = 1 To .Rows - 1
             If Val(.TextMatrix(i, .ColIndex("类型"))) = 1 Then
                blnStart = False: Exit For
             End If
        Next
    End With
    If Not blnStart Then
        For i = 1 To mobjPayCards.Count
           Set objCard = mobjPayCards(i)
           If objCard.结算性质 = -99 Then objCard.支付启用 = False
        Next
    End If
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化控件
    '编制:刘兴洪
    '日期:2011-06-13 14:09:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl金额 As Double, rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    With vsBlance
        .Cell(flexcpFontBold, 1, 0, 1, .COLS - 1) = True
        .Clear: .Rows = 2: i = 0: .COLS = 18
        .TextMatrix(0, i) = "卡类别ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "消费卡ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算性质": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "类型": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "支付方式": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "支付金额": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "结算号码": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "备注": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "卡号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易流水号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易说明": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "删除标志": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否退现": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否全退": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否转帐及代扣": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "卡类别名称": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否验证": .ColWidth(i) = 0: i = i + 1 '用于判断预交款是否已验证
        
        For i = 0 To .COLS - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            Select Case .ColKey(i)
            Case "结算性质", "类型", "删除标志", "是否退现", "是否全退", "是否转帐及代扣", "卡类别名称", "结算状态", "是否验证"
                .ColHidden(i) = True
            Case "支付金额"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
    End With
    With mCurCarge
        .dbl本次退预交 = 0
        .dbl退费合计 = 0
        .dbl已算误差 = 0
        .dbl本次医保退费 = 0
        .dbl已退合计 = 0
        .dbl本次应收 = 0
        .dbl当前未退 = 0
        .dbl费用余额 = 0
        .dbl可用预交 = 0
        .dbl预交余额 = 0
    End With
    
    mstr退支票 = ""
    If mobjDelBalance.rs结算方式 Is Nothing Then
        Set mobjDelBalance.rs结算方式 = Get结算方式("收费")
    ElseIf mobjDelBalance.rs结算方式.State <> 1 Then
        Set mobjDelBalance.rs结算方式 = Get结算方式("收费")
    End If
    mobjDelBalance.rs结算方式.Filter = "应付款=1"
    If Not mobjDelBalance.rs结算方式.EOF Then
         mstr退支票 = Nvl(mobjDelBalance.rs结算方式!名称)
    End If
    mobjDelBalance.rs结算方式.Filter = 0
    Call initInsure
    Call Init退费方式
    
    If mbytFunc = EM_FUN_退费 And mcllDelPro.Count <> 0 Then
        '单据未保存时,清除
        Call InitBalanceData
    Else
        Call LoadData
    End If
    Call SetDeleteVisible '进入结算界面时删除按钮应该根据情况显示
    
    Call Load退费方式: Call LoadPatiInfor
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetOldBalanceMoney(ByVal int类型 As Integer, ByVal objCard As Card) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据类型，确定原结算方式的金额
    '入参:int类型-类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    '返回:返回原结算金额
    '编制:刘兴洪
    '日期:2014-07-08 15:49:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Integer, blnFindByList As Boolean
    
    On Error GoTo errHandle
    With mobjDelBalance
        If .rsBalance Is Nothing Then Exit Function
        If .rsBalance.State <> 1 Then Exit Function
        .rsBalance.Filter = ""
        
        '93114，退费时使用转帐方式
        If CheckThreeSwapCanTransfer(objCard, mobjDelBalance.原结帐ID) Then
            '计算可转帐金额
            Do While Not .rsBalance.EOF
                Select Case Val(Nvl(.rsBalance!类型))
                Case 0, 1, 4 '普通结算,预交款,老一卡通
                    dblMoney = dblMoney + Val(Nvl(.rsBalance!冲预交))
                Case 3, 5 '一卡通,消费卡
                    If Val(Nvl(.rsBalance!是否退现)) = 1 Then
                        dblMoney = dblMoney + Val(Nvl(.rsBalance!冲预交))
                    End If
                End Select
                .rsBalance.MoveNext
            Loop
            
            '减去已退款非医保的金额
            For i = 1 To vsBlance.Rows - 1
                Select Case Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("类型")))
                Case 0, 1, 4
                    dblMoney = dblMoney - Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("支付金额")))
                Case 3, 5
                    If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("类型"))) = 3 And Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("卡类别ID"))) = objCard.接口序号 Then
                        '对列表中的进行结算
                        blnFindByList = True
                    Else
                        If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("是否退现"))) = 1 Then
                            dblMoney = dblMoney - Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("支付金额")))
                        End If
                    End If
                End Select
            Next
            
            If blnFindByList Then
                If dblMoney > -1 * mCurCarge.dbl退费合计 Then dblMoney = -1 * mCurCarge.dbl退费合计
            Else
                If dblMoney > -1 * mCurCarge.dbl当前未退 Then dblMoney = -1 * mCurCarge.dbl当前未退
            End If
            If dblMoney < 0 Then dblMoney = 0
            GetOldBalanceMoney = RoundEx(dblMoney, 6)
            Exit Function
        End If
       
        '77338,冉俊明,2014-9-1,没有正确获取预交款金额
        If objCard.接口序号 > 0 Then
            If objCard.消费卡 = False Then '一卡通
                .rsBalance.Filter = "类型=" & int类型 & " And 卡类别ID=" & objCard.接口序号
            Else '消费卡
                .rsBalance.Filter = "类型=" & int类型 & " And 结算卡序号=" & objCard.接口序号
            End If
        ElseIf objCard.结算性质 = 2 And objCard.结算方式 Like "*卡" Then '87532
            .rsBalance.Filter = "类型=" & int类型 & " And 结算方式='" & objCard.结算方式 & "'"
        Else
            .rsBalance.Filter = "类型=" & int类型
        End If
        If .rsBalance.EOF Then
            .rsBalance.Filter = 0
            Exit Function
        End If
        .rsBalance.MoveFirst
        Do While Not .rsBalance.EOF
            dblMoney = dblMoney + Val(Nvl(.rsBalance!冲预交))
            .rsBalance.MoveNext
        Loop
        GetOldBalanceMoney = RoundEx(dblMoney, 6)
        .rsBalance.Filter = 0
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Sub SetControlProperty(Optional ByVal blnLoadDefault As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件属性
    '入参:blnLoadDefault-是否加载缺省值
    '编制:刘兴洪
    '日期:2014-07-10 17:49:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngTop As Long, sngSplitHeight As Single, dbl现金 As Double
    Dim bln分币 As Boolean, dblMoney As Double, dblTemp As Double
    Dim bln退款 As Boolean '主要是医保相关结算大于了单据收费
    Dim blnVisible As Boolean, blnEnabled As Boolean
    Dim objCard As Card, intIndex As Integer
    Dim blnDel As Boolean
    
    blnDel = mCurCarge.dbl当前未退 <= 0

    If GetCurCard(objCard) = False Then
        Set objCard = New Card
    End If
    
    sngSplitHeight = 80
    lbl已退合计.Caption = "已付合计:" & Format(Abs(mCurCarge.dbl已退合计), "###0.00;-###0.00;0.00;0.00;")
    
    If objCard.结算性质 = 1 Then
        If RoundEx(mCurCarge.dbl已算误差, 6) = RoundEx(mCurCarge.dbl当前未退, 6) Then
            dbl现金 = 0
        Else
            dblMoney = mCurCarge.dbl当前未退
            If mobjDelBalance.intInsure > 0 Then
                If mInsurePara.分币处理 Then
                    bln分币 = True
                    dbl现金 = CentMoney(CCur(dblMoney))
                Else
                    dbl现金 = Format(dblMoney, "0.00")
                End If
            Else
                bln分币 = True
                dbl现金 = RoundEx(CentMoney(CCur(dblMoney)), 6)
            End If
        End If
        lbl未退金额.Caption = Format(Abs(dbl现金), "0.00")
    Else
        lbl未退金额.Caption = Format(Abs(mCurCarge.dbl当前未退), "0.00")
    End If
    If blnDel Then
        stcCurDelTitle.Caption = "当前应退"
        lbl未退金额.ForeColor = vbRed
        lblPayType.Caption = "退  款"
        lbl找补.Caption = "收  零"
    Else
        stcCurDelTitle.Caption = "当前应收"
        lbl未退金额.ForeColor = vbBlue
        lblPayType.Caption = "收  款"
        lbl找补.Caption = "找  补"
    End If
    
    '其他非医保结算和一卡通和老版一卡通
    '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
    '77353,冉俊明,2014-9-1,退费时收款,若病人存在预交款,允许使用预交款进行缴款,选择该性质的结算方式时「结算号码」,「摘要」不允许输入
    blnEnabled = InStr(",1,3,4,5,6,-99,", "," & objCard.结算性质 & ",") = 0
    txt结算号码.Enabled = blnEnabled
    txt摘要.Enabled = blnEnabled
                
    '缺省金额的设置
    If blnLoadDefault Then
        '77324,冉俊明,2014-9-1,对于三方账户允许退现时,应该禁止录入退款金额,只能按照收取的金额默认
        txt缴款.Locked = False
        If objCard.接口序号 > 0 Then          '三方结算和消费卡
            '不能超过已退金额
            If mCurCarge.dbl当前未退 <= 0 Then
                 '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                dblTemp = GetOldBalanceMoney(IIf(objCard.消费卡, 5, 3), objCard)
                If objCard.是否全退 Then
                    txt缴款.Text = FormatEx(dblTemp, 6, , , 2)
                    txt缴款.Locked = True
                Else
                    If dblTemp >= Abs(mCurCarge.dbl当前未退) Then
                        txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
                    Else
                        txt缴款.Text = FormatEx(dblTemp, 6, , , 2)
                    End If
                    txt缴款.Locked = objCard.是否退现 = False
                End If
            Else
                txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
            End If
        ElseIf objCard.结算性质 = 7 And objCard.接口序号 <= 0 Then '老一卡通
             '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If mCurCarge.dbl当前未退 <= 0 Then
                dblTemp = GetOldBalanceMoney(4, objCard)
                If dblTemp >= Abs(mCurCarge.dbl当前未退) Then
                    txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
                Else
                    txt缴款.Text = FormatEx(dblTemp, 6, , , 2)
                End If
'                txt缴款.Locked = True
            Else
                txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
            End If
        ElseIf objCard.结算性质 = -99 Then  '冲预交
             '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If mCurCarge.dbl当前未退 <= 0 Then
                dblTemp = GetOldBalanceMoney(1, objCard)
                If dblTemp >= Abs(mCurCarge.dbl当前未退) Then
                    txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
                Else
                    txt缴款.Text = FormatEx(dblTemp, 6, , , 2)
                End If
            Else
                txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
            End If
        ElseIf objCard.结算性质 = 1 Then    '现金处理
            If gTy_Module_Para.bln现金退款缺省方式 Then
                txt缴款.Text = Format(Abs(dbl现金), "0.00")
            Else
                txt缴款.Text = "0.00"
            End If
        ElseIf objCard.结算性质 = 2 And objCard.结算方式 Like "*卡" Then  '非医保结算方式为"***卡",87532
            If mCurCarge.dbl当前未退 <= 0 Then
                dblTemp = GetOldBalanceMoney(0, objCard)
                If dblTemp >= Abs(mCurCarge.dbl当前未退) Then
                    txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
                Else
                    txt缴款.Text = FormatEx(dblTemp, 6, , , 2)
                End If
            Else
                txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
            End If
        Else
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            txt缴款.Text = Format(Abs(mCurCarge.dbl当前未退), "0.00")
        End If
    End If
    
    '计算找补
    If blnDel Then
        dblTemp = RoundEx(Val(txt缴款.Text) - (-1 * mCurCarge.dbl当前未退 + mCurCarge.dbl本次误差费), 6)
        If dblTemp > 0 Then
            txt找补.Text = Format(dblTemp, "0.00")
            txt找补.ForeColor = vbRed
        Else
            txt找补.Text = ""
        End If
    Else
        dblTemp = Val(txt缴款.Text) - mCurCarge.dbl当前未退
        txt找补.ForeColor = lbl找补.ForeColor
        If dblTemp > 0 Then
            txt找补.Text = Format(dblTemp, "0.00")
        End If
    End If
    Call SetControlColor
End Sub

Private Sub cbo支付方式_Click()
    Dim intIndex As Integer
    Dim objCard As Card, i As Integer
    Dim intSelectIndex As Integer
    
    If mblnFirst Then Exit Sub
    If mblnNotClick Then Exit Sub
    If mlngPre支付方式 = cbo支付方式.ItemData(cbo支付方式.ListIndex) Then Exit Sub
    
    '105432
    If mlngPre支付方式 > 0 And Val(txt缴款.Text) <> 0 Then
        '如果不在收费结算方式中就不用检查，主要针对支持“转帐及代扣”的
        Set objCard = mobjPayCards(mlngPre支付方式)
        mobjDelBalance.rsBalance.Filter = "结算方式='" & objCard.结算方式 & "' And 退费=0"
        
        If Not mobjDelBalance.rsBalance.EOF Then
            mblnNotClick = True
            intSelectIndex = cbo支付方式.ListIndex
            cbo支付方式.ListIndex = cbo.FindIndex(cbo支付方式, mlngPre支付方式)
            If ThreeBalanceCheck(Me, mlngModule, mobjPayCards(mlngPre支付方式), _
                  mcllForceDelToCash, cbo支付方式.Text) = False Then mblnNotClick = False: Exit Sub
            cbo支付方式.ListIndex = intSelectIndex
            mblnNotClick = False
        End If
    End If
    
    mlngPre支付方式 = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    txt缴款.Text = ""
    If cbo支付方式.ListIndex < 0 Then GoTo SetProperty:
    
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    Set objCard = mobjPayCards(intIndex)
    '切换回来后要清除
    If objCard.接口序号 > 0 And objCard.消费卡 = False Then
        For i = 1 To mcllForceDelToCash.Count
            If mcllForceDelToCash(i)(1) = objCard.名称 Then Exit For
        Next
        If i <= mcllForceDelToCash.Count Then mcllForceDelToCash.Remove i
    End If
    
    If objCard.结算性质 = 7 And objCard.接口序号 <= 0 Then '老一卡通
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
'    ElseIf objCard.消费卡 Then
'         If IsExistSquare = True Then
'            If MsgBox("已经存在" & cbo支付方式.Text & ",是否删除?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
'                Call ClearSquareBalance
'                Set mcllSquareBalance = Nothing
'            End If
'         End If
    End If
SetProperty:
     Call SetControlProperty(True)
     If txt缴款.Enabled Then txt缴款.SetFocus
End Sub
Private Function IsExistSquare(ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否存在消费卡结算
    '入参:lngCardTypeID-消费卡编号
    '出参:
    '返回:存在成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-12 11:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, varTemp As Variant
    If mcllSquareBalance Is Nothing Then Exit Function
    For i = 1 To mcllSquareBalance.Count
        varTemp = mcllSquareBalance(i)
        ' array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
        If Val(varTemp(0)) = lngCardTypeID Then
            IsExistSquare = True
            Exit Function
        End If
    Next
End Function

Private Function CheckOneCard(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否正确
    '返回:一卡通验证正确或非一卡通,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-23 17:07:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CurOneCard As Currency, dblMoney As Double, dblTemp As Double
    Dim strTittle As String, strCardNo As String
    
    If objCard.结算性质 <> 7 Then CheckOneCard = True: Exit Function
    
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If mobjICCard Is Nothing Then
        MsgBox "一卡通接口创建失败!", vbOKOnly, gstrSysName
        Exit Function
    End If
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
    
    dblMoney = Val(txt缴款.Text)
    If strTittle = "缴款" Then
        CurOneCard = mobjICCard.GetSpare
        If CurOneCard < dblMoney Then
            MsgBox "卡余额不够支付,请检查!" & vbCrLf & vbCrLf & _
            "   卡 余  额" & Format(CurOneCard, "0.00") & vbCrLf & _
            "   本次支付" & FormatEx(Val(txt缴款.Text), 6), vbInformation, gstrSysName
            Exit Function
        End If
        stbThis.Panels(4).Text = Format(CurOneCard, "0.00")
        stbThis.Panels(4).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(CurOneCard, "0.00")
        CheckOneCard = True
        Exit Function
    End If
    
     '退款检查
    If mobjDelBalance.rsBalance Is Nothing Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjDelBalance.rsBalance.State <> 1 Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    mobjDelBalance.rsBalance.Filter = "类型=4"
    If mobjDelBalance.rsBalance.EOF Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo = "" Then
        MsgBox "一卡通读卡失败,请将IC卡放在读卡器中", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If strCardNo <> Nvl(mobjDelBalance.rsBalance!卡号) Then
        MsgBox "当前卡号与扣款卡号不一致,不能进行退费.", vbInformation, gstrSysName
        Exit Function
    End If
    
    dblTemp = Format(Val(Nvl(mobjDelBalance.rsBalance!冲预交)), "0.00")
    If RoundEx(dblMoney, 6) <> Format(dblTemp, "0.00") Then
        MsgBox "一卡通结算必须全退,请检查!" & vbCrLf & vbCrLf & _
        "   结算金额" & Format(dblTemp, "0.00") & vbCrLf & _
        "   本次支付" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    CheckOneCard = True
End Function

Private Function CheckThreeSwapValied(ByVal objCard As Card, Optional dblDelMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方交易验证
    '入参:objCard-三方卡
    '     dblDelMoney-退款金额
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接品的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl帐户余额 As Double
    Dim strBrushCard As TY_BrushCard, cllSquareBalance As Collection
    Dim strExpand As String, bln退现 As Boolean
    Dim strBalanceIDs As String
    
    On Error GoTo errHandle
    If objCard Is Nothing Then
        If GetCurCard(objCard) = False Then
            MsgBox "当前" & lblPayType.Caption & "方式未选择,请选择!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then CheckThreeSwapValied = True: Exit Function
    
    If mCurCarge.dbl当前未退 <= 0 Or dblDelMoney <> 0 Then
        strTittle = "退款"
    Else
        strTittle = "缴款"
    End If
    
    mCurBrushCard = strBrushCard
    If dblDelMoney = 0 Then
        If Val(txt缴款.Text) = 0 Then
            MsgBox strTittle & "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
             Exit Function
        End If
    End If
    
    If strTittle = "缴款" Then
        If Abs(Val(txt缴款.Text)) > Format(Abs(mCurCarge.dbl当前未退), "0.00") And Val(txt缴款.Text) <> 0 Then
            MsgBox strTittle & "金额不能大于本次未付金额:" & Format(mCurCarge.dbl当前未退, "0.00") & " ！", vbInformation, gstrSysName
            Exit Function
        End If
        Set cllSquareBalance = Nothing
        Set mcllCurSquareBalance = Nothing
        
        '   zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln消费卡 As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByRef dbl金额 As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln退费 As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln退现 As Boolean = False, _
            Optional ByVal bln余额不足禁止 As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal bln转预交 As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXmlIn As String = "") As Boolean
            '       strXmlIn-三方卡调用XML入参,目前格式如下:
            '       <IN>
            '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
            '       </IN>
           '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
        
        dblMoney = Val(txt缴款.Text)
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, _
            objCard.接口序号, objCard.消费卡, _
            mobjDelBalance.姓名, mobjDelBalance.性别, mobjDelBalance.年龄, dblMoney, _
            mCurBrushCard.str卡号, mCurBrushCard.str密码, _
            False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>0</CZLX></IN>") = False Then Exit Function
            '保存前,一些数据检查
            'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal strCardTypeID As Long, ByVal strCardNo As String, _
            ByVal dblMoney As Double, ByVal strNOs As String, _
            Optional ByVal strXMLExpend As String
            'mobjDelBalance.strNOs:单独保存时,没有相关时,可能为空.
            If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, objCard.接口序号, _
                objCard.消费卡, mCurBrushCard.str卡号, dblMoney, mobjDelBalance.CurDelNos, strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
            '入参:frmMain-调用的主窗体
            '        lngModule-模块号
            '        strCardNo-卡号
            '        strExpand-预留，为空,以后扩展
            '出参:dblMoney-返回帐户余额
            If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, objCard.接口序号, _
                  mCurBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡) = False Then Exit Function
        
        stbThis.Panels(4).Text = Format(dbl帐户余额, "0.00")
        stbThis.Panels(4).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
        mCurBrushCard.dbl帐户余额 = RoundEx(dbl帐户余额, 2)
        If dbl帐户余额 <> 0 And dbl帐户余额 < dblMoney Then
            MsgBox objCard.结算方式 & "的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        CheckThreeSwapValied = True
        Exit Function
    End If
    
    '退款检查
    If mobjDelBalance.rsBalance Is Nothing Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjDelBalance.rsBalance.State <> 1 Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '93114，退费时使用转帐方式
    If CheckThreeSwapCanTransfer(objCard, mobjDelBalance.原结帐ID) Then
        If dblDelMoney <> 0 Then
            dblMoney = dblDelMoney
        Else
            dblMoney = Val(txt缴款.Text)
        End If
        dblTemp = GetOldBalanceMoney(3, objCard)
        
        If RoundEx(dblTemp, 6) < RoundEx(dblMoney, 6) Then
            MsgBox "注意:" & vbCrLf & "   输入的退款金额大于了" & objCard.名称 & "的可退金额，请检查！" & vbCrLf & _
                   "   可退金额:" & Format(dblTemp, "###0.00;-###0.00;;") & vbCrLf & _
                   "   当前退款:" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '弹出刷卡界面
        'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByVal dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-三方卡调用XML入参,目前格式如下:
        '       <IN>
        '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        '       </IN>
         If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.接口序号, _
             objCard.消费卡, mobjDelBalance.姓名, mobjDelBalance.性别, _
             mobjDelBalance.年龄, dblMoney, mCurBrushCard.str卡号, mCurBrushCard.str密码, _
             True, True, bln退现, True, Nothing, False, False, "<IN><CZLX>1</CZLX></IN>") = False Then Exit Function
    
        '调用转帐接口
        'zlTransferAccountsCheck 转帐检查接口
        '参数名  参数类型    入/出   备注
        'frmMain Object  In  调用的主窗体
        'lngModule   Long    In  HIS调用模块号
        'lngCardTypeID   Long    In  卡类别ID
        'strCardNo   String  In  卡号
        'dblMoney    Double  In  转帐金额(代扣时为负数)
        'strBalanceIDs   String  In  结帐IDs，多个用逗号分离，表示本次对哪此收费项目进行重新医保补结算
        'strXMLExpend String In   XML串:
        '                            <IN>
        '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务
        '                            </IN>
        '                    Out  XML串:
        '                            <OUT>
        '                               <ERRMSG>错误信息</ERRMSG >
        '                            </OUT>
        '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
        '说明:
        '１. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
        '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
        '构造XML串
        strXMLExpend = "<IN><CZLX>4</CZLX></IN>"
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModule, objCard.接口序号, _
            mCurBrushCard.str卡号, dblMoney, mobjDelBalance.原结帐ID, strXMLExpend) = False Then Exit Function
    Else
        '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        mobjDelBalance.rsBalance.Filter = "类型=3 And 卡类别ID=" & objCard.接口序号
        If mobjDelBalance.rsBalance.EOF Then
            MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.结算方式 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        If dblDelMoney <> 0 Then
            dblMoney = dblDelMoney
        Else
            dblMoney = Val(txt缴款.Text)
        End If
        dblTemp = 0
        With mobjDelBalance.rsBalance
            Do While Not .EOF
                dblTemp = dblTemp + Val(Nvl(!冲预交))
                .MoveNext
            Loop
            mobjDelBalance.rsBalance.MoveFirst
            dblTemp = RoundEx(dblTemp, 6)
        End With
    
        If dblTemp < dblMoney Then
            MsgBox "注意:" & vbCrLf & "   输入的退款金额大于了" & objCard.名称 & "的可退金额，请检查！" & vbCrLf & _
                   "   可退金额:" & Format(dblTemp, "###0.00;-###0.00;;") & vbCrLf & _
                   "   当前退款:" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        If objCard.是否全退 And Not objCard.是否退现 Then
            If dblTemp <> dblMoney Then
                MsgBox "注意:" & vbCrLf & objCard.名称 & "进行退款时，必须全退！" & vbCrLf & _
                "  剩余未退:" & Format(dblTemp, "0.00") & vbCrLf & _
                "  当前金额:" & Format(dblMoney, "0.00"), vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
        
        mCurBrushCard.str卡号 = Nvl(mobjDelBalance.rsBalance!卡号)
        mCurBrushCard.str交易流水号 = Nvl(mobjDelBalance.rsBalance!交易流水号)
        mCurBrushCard.str交易说明 = Nvl(mobjDelBalance.rsBalance!交易说明)
        
        'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
            ByVal strBalanceIDs As String, _
            ByVal dblMoney As Double, ByVal strSwapNo As String, _
            ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:帐户回退交易前的检查
            '入参:frmMain-调用的主窗体
            '       lngModule-调用的模块号
            '       lngCardTypeID-卡类别ID
            '       strCardNo-卡号
            '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
            '       dblMoney-退款金额
            '       strSwapNo-交易流水号(退款时检查)
            '       strSwapMemo-交易说明(退款时传入)
            '       strXMLExpend    XML IN  可选参数(扩展用):
            '        <TFDATA> //退费数据
            '          <YCTF>1</YCTF> //是否异常重退:1-异常重退;0-退费 此节点可能没有
            '          <TFLIST> //退费列表
            '            <NO></NO> // 退费单据
            '            <TFITEM> //退费项
            '              <SerialNum></SerialNum> //序号
            '              …
            '            </TFITEM>
            '          </TFLIST>
            '          ....
            '        </TFDATA >
            '返回:退款合法,返回true,否则返回Flase
        strXMLExpend = mfrmMain.GetDelXMLExpend()
        strBalanceIDs = "3|" & mobjDelBalance.原结帐ID '& IIf(mobjDelBalance.结帐ID = 0, "", "," & mobjDelBalance.结帐ID)
        If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModule, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, _
            strBalanceIDs, dblMoney, mCurBrushCard.str交易流水号, mCurBrushCard.str交易说明, strXMLExpend) = False Then Exit Function
        
        If objCard.是否退款验卡 Then
           '弹出刷卡界面
            'zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln消费卡 As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByVal dbl金额 As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln退费 As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln退现 As Boolean = False, _
            Optional ByVal bln余额不足禁止 As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal bln转预交 As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXmlIn As String = "") As Boolean
            '       strXmlIn-三方卡调用XML入参,目前格式如下:
            '       <IN>
            '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
            '       </IN>
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.接口序号, _
                objCard.消费卡, mobjDelBalance.姓名, mobjDelBalance.性别, _
                mobjDelBalance.年龄, dblMoney, mCurBrushCard.str卡号, mCurBrushCard.str密码, _
                True, True, bln退现, True, Nothing, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
        End If
        
    End If
    CheckThreeSwapValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPrepayMoneyIsValied(ByVal objCard As Card, Optional ByVal intType As Integer, Optional ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查预交数据输入是否合法
    '返回:合法,返回true,否则返回False
    '参数:
    '   intTppe:0-结算方式选择预交款 1-结算列表默认预交款
    '编制:刘兴洪
    '日期:2014-07-08 18:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTittle As String, i As Long, str结算方式 As String
    Dim int性质  As Integer, dblTemp As Double
    
    On Error GoTo errHandle
    If objCard.结算性质 <> -99 Then CheckPrepayMoneyIsValied = True: Exit Function
    
    If intType = 0 Then
        Call txt缴款_LostFocus
        strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退", "冲")
        dblMoney = Val(txt缴款.Text)
        If RoundEx(dblMoney, 6) = 0 Then
            MsgBox "未输入" & strTittle & "预交款金额!", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        strTittle = IIf(dblMoney <= 0, "退", "冲")
    End If

    If strTittle = "冲" Then
        Dim str家属IDs As String
        If zlDatabase.PatiIdentify(Me, glngSys, mobjDelBalance.病人ID, dblMoney, mlngModule, 1, , IIf(-1 * gdbl预存款消费验卡 >= dblMoney, False, True), True, str家属IDs, _
            (gdbl预存款消费验卡 <> 0), (gdbl预存款消费验卡 = 2)) = False Then Exit Function
        mobjDelBalance.家属IDs = str家属IDs
        CheckPrepayMoneyIsValied = True
        Exit Function
    End If
    
    dblTemp = RoundEx(GetOldBalanceMoney(1, objCard), 6)
    If dblMoney > dblTemp Then
        MsgBox "退预交款不能超过收费结算的剩余预交款（" & FormatEx(dblTemp, 6) & "）！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If intType = 0 Then
        With vsBlance
            For i = .Rows - 1 To 1 Step -1
                str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
                ' 0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
                If int性质 = 1 And str结算方式 <> "" Then
                    MsgBox "已经使用了" & str结算方式 & ",不能再次使用预交款" & lblPayType.Caption & "!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        End With
    End If
        
    '退预交款
    If gbyt预存款退费验卡 = 0 Then CheckPrepayMoneyIsValied = True: Exit Function
    If Not zlDatabase.PatiIdentify(Me, glngSys, mobjDelBalance.病人ID, dblMoney, , , , , True, , , (gbyt预存款退费验卡 = 2)) Then Exit Function
    CheckPrepayMoneyIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckCashValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查现金支付方式的一些合法情检查
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intIndex As Integer, dblMoney As Double
    Dim strTittle As String
    
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
   
    On Error GoTo errHandle
    
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    If objCard.结算性质 <> 1 Then CheckCashValied = True: Exit Function
    dblMoney = Val(txt缴款.Text)
    If strTittle = "缴款" Then
        Select Case gTy_Module_Para.byt缴款控制
        Case 1, 3 '1-多病缴款;3单病人缴款累计
            If RoundEx(mCurCarge.dbl当前未退 - mCurCarge.dbl本次误差费, 2) > 0 And RoundEx(dblMoney, 6) = 0 Then
               If MsgBox("注意:" & vbCrLf & "    该病人未输入缴款金额,是否继续收费? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        Case 2  '2-收费时必须要输入缴款金额
            If RoundEx(mCurCarge.dbl当前未退 - mCurCarge.dbl本次误差费, 2) > 0 And RoundEx(dblMoney, 6) = 0 Then
                MsgBox "注意:" & vbCrLf & _
                "    该病人未输入缴款金额,不能进行收费!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
        Case Else   ',0-代表不进行缴款输入和累计控制
            '医保结算缴款检查:要缴而未缴时,以缴款作为结束量不处理,因为是强行输入0跳过缴款的
            If mobjDelBalance.intInsure <> 0 And Not mInsurePara.门诊连续收费 And _
                RoundEx(mCurCarge.dbl当前未退 - mCurCarge.dbl本次误差费, 2) > 0 And RoundEx(dblMoney, 6) = 0 Then
                '刘兴洪:27536 20100119
                If mInsurePara.不提醒缴款金额不足 = False Then
                    MsgBox "提醒你:" & vbCrLf & vbTab & "该医保病人的费用未全部结算，请注意收取病人缴款！", vbInformation, gstrSysName
                End If
            End If
        End Select
        If RoundEx(dblMoney, 6) <> 0 Then
            If Val(txt找补.Text) < 0 Then
                MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        CheckCashValied = True
        Exit Function
    End If
    '退款
'    If dblMoney = 0 Then
'        MsgBox "未输入退款金额！", vbInformation, gstrSysName
'        Exit Function
'    End If
    If dblMoney < Abs(Val(lbl未退金额.Caption)) And RoundEx(dblMoney, 6) <> 0 Then
        MsgBox "输入的退款金额不足！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckCashValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckChequeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查支票支付方式的一些合法情检查
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intIndex As Integer, dblMoney As Double
    Dim strTittle As String
    
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
   
    On Error GoTo errHandle
    
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    If objCard.结算性质 <> 2 Or Not objCard.结算方式 Like "*支票*" Then CheckChequeValied = True: Exit Function
    
    dblMoney = Val(txt缴款.Text)
    
    If strTittle = "缴款" Then
        If RoundEx(dblMoney, 6) = 0 Then
            MsgBox "未输入缴款金额！", vbInformation, gstrSysName
            Exit Function
        End If
        CheckChequeValied = True
        Exit Function
    End If
    '退款
    If RoundEx(dblMoney, 6) = 0 And Not mblnTurnFee Then
        MsgBox "未输入退款金额！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckChequeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckOtherValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查支票支付方式的一些合法情检查
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intIndex As Integer, dblMoney As Double
    Dim strTittle As String, dblTemp As Double
    
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
   
    On Error GoTo errHandle
    
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    
    If objCard.接口序号 > 0 Or objCard.结算方式 Like "*支票*" Or objCard.结算性质 = -99 Or objCard.结算性质 = 1 Then CheckOtherValied = True: Exit Function
    
    dblMoney = Val(txt缴款.Text)
    

    If strTittle = "缴款" Then
        If RoundEx(dblMoney, 6) = 0 Then
            MsgBox "未输入缴款金额！", vbInformation, gstrSysName
            Exit Function
        End If
        If dblMoney > RoundEx(mCurCarge.dbl当前未退, 2) Then
            MsgBox "注意:" & vbCrLf & "    输入的缴款金额大于了未支付的金额,不能继续!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        CheckOtherValied = True
        Exit Function
    End If
    
    '退款
    If RoundEx(dblMoney, 6) = 0 And Not mblnTurnFee Then
        MsgBox "未输入退款金额！", vbInformation, gstrSysName
        Exit Function
    End If
    If dblMoney > RoundEx(Abs(mCurCarge.dbl当前未退), 2) Then
        MsgBox "注意:" & vbCrLf & "    输入的退款金额大于了可退金额,不能继续!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.结算性质 = 2 And objCard.结算方式 Like "*卡" Then '87532
        dblTemp = RoundEx(GetOldBalanceMoney(0, objCard), 6)
        If dblMoney > dblTemp Then
            MsgBox "注意：" & vbCrLf & "   输入的退款金额大于了 " & objCard.结算方式 & " 的可退金额，请检查！" & vbCrLf & _
                   "   可退金额：" & Format(dblTemp, "###0.00;-###0.00;;") & vbCrLf & _
                   "   当前退款：" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    CheckOtherValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
        

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入数据的有效性,数据有效,返回true,否则返回False
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-13 16:30:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strTittle As String, i As Long, str结算方式 As String
    Dim int性质 As Integer, objCard As Card
    
    On Error GoTo errHandle
    If Not CheckTextLength("结算号码", txt结算号码) Then Exit Function
    If Not CheckTextLength("摘要", txt摘要) Then Exit Function
    
    If mbytFunc = EM_FUN_退费 And mcllDelPro.Count > 0 Then
        If mfrmMain.CheckSelectItemCanDel(mobjDelBalance.CurDelNos) = False Then Exit Function
    End If
    
    '并发检查
    If mbytFunc = EM_FUN_重退 Then
        If zlIsCheckExistErrBill(mobjDelBalance.结算序号) = False Then
            MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
            Exit Function
        End If
        If zlCheckOtherSessionDoing(mobjDelBalance.结算序号) Then
            MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If GetCurCard(objCard) = False Then
        MsgBox "当前" & lblPayType.Caption & "方式未选择,请选择!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    '93114
    If CheckThreeSwapCanTransfer(objCard, mobjDelBalance.原结帐ID) = False Or objCard.接口序号 <= 0 Then
        If CheckIsExistCashValied(objCard) = False Then Exit Function
    End If
    
    '检查输入的合法性
    If mCurCarge.dbl当前未退 <= 0 Then
        strTittle = "退款"
    Else
        strTittle = "缴款"
    End If
    
    If Not IsNumeric(txt缴款.Text) And txt缴款.Text <> "" Then
        MsgBox strTittle & "输入了无效数值！", vbInformation, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    End If
    If Val(txt缴款.Text) < 0 Then
        MsgBox strTittle & "不能输入负数！", vbInformation, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    End If
    
    If Abs(Val(txt缴款.Text)) > 999999999 Then
        MsgBox "输入的缴款金额过大,最大不能超过-999999999至999999999!", vbOKOnly, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
        Exit Function
    End If
    
    If txt结算号码.Text <> "" Then
        If zlCommFun.ActualLen(txt结算号码) > 30 Then
            MsgBox "结算号码最多允许输入30个字符或 15个汉字！", vbInformation, gstrSysName
            If txt结算号码.Enabled And txt结算号码.Visible Then txt结算号码.SetFocus
            zlControl.TxtSelAll txt结算号码: Exit Function
        End If
        If InStr(txt结算号码, "'") > 0 Then
            MsgBox "结算号码含有非法字符(单引号)！", vbInformation, gstrSysName
            If txt结算号码.Enabled And txt结算号码.Visible Then txt结算号码.SetFocus
            zlControl.TxtSelAll txt结算号码: Exit Function
        End If
    End If
    If txt摘要.Text <> "" Then
        If zlCommFun.ActualLen(txt摘要) > 50 Then
            MsgBox "摘要最多允许输入50个字符或 25个汉字！", vbInformation, gstrSysName
            If txt摘要.Enabled And txt摘要.Visible Then txt摘要.SetFocus
            zlControl.TxtSelAll txt摘要: Exit Function
        End If
        If InStr(txt摘要, "'") > 0 Then
            MsgBox "摘要含有非法字符(单引号)！", vbInformation, gstrSysName
            If txt摘要.Enabled And txt摘要.Visible Then txt摘要.SetFocus
            zlControl.TxtSelAll txt摘要: Exit Function
        End If
    End If
    With vsBlance
        For i = .Rows - 1 To 1 Step -1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            ' 0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            If objCard.结算方式 = str结算方式 And int性质 <> 1 And (int性质 <> 5 Or (int性质 = 5 And .Cell(flexcpData, i, .ColIndex("支付金额")) < 0)) Then
                '预交款在预交款检查函数中有此检查
                MsgBox objCard.结算方式 & " 已经存在,不能再用" & objCard.结算方式 & "进行" & Replace(lblPayType.Caption, " ", "") & "!", vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Function
            End If
        Next
    End With
        
    If CheckInterfaceNumIsValied(objCard) = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    
    '1.一卡通刷卡
    If CheckOneCard(objCard) = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    
'    '2.三卡交易检查
'    If CheckThreeSwapValied(objCard) = False Then
'        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
'        zlControl.TxtSelAll txt缴款
'        Exit Function
'    End If
    
    '3.消费卡检查
    '退费
    If CheckSquareDelValied(objCard) = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    '收费
    If CheckSquareBalanceValied(objCard) = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    
    '3.检查预交款是否合法
    If CheckPrepayMoneyIsValied(objCard) = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    
    '4.现金方式的检查
    If CheckCashValied = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    '5.检查支票处理相关
    If CheckChequeValied = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    '6.其他收费方式检查
    If CheckOtherValied = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If

    '检查当前单据是否被其他人执行完成,主要是并发原因进行检查
    '防止其他操作员操作:
    '45186
    If mobjDelBalance.结帐ID <> 0 Then
        gstrSQL = "" & _
        "   Select  1  From 病人预交记录 A " & _
        "   Where   A.结帐ID=[1] and nvl(A.校对标志,0)<>0 and Rownum =1 and A.记录状态=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.结帐ID)
        If rsTemp.EOF Then
            '估计是被他人执行,现在需要检查是否被他人执行
            gstrSQL = "Select 记录状态, 操作员姓名,费用状态 From 门诊费用记录 Where 结帐ID=[1] And rownum=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.结帐ID)
            If Not rsTemp.EOF Then
                If Val(Nvl(rsTemp!记录状态)) <> 1 Then
                    MsgBox "该单据已经被其他操作员作废,不能再进行收费!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                If Val(Nvl(rsTemp!费用状态)) <> 1 Then
                    MsgBox "该次收费已经被他人收费,不能再进行收费!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                If Nvl(rsTemp!操作员姓名) <> UserInfo.姓名 Then
                    MsgBox "该单据不是本人收费单,不能收取其他操作员的单据!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo支付方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub ClearReMoveSquareBalance(ByVal lng卡类别ID As Long, Optional ByVal lng消费卡ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移除指定的消费卡结算
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-12 12:10:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, varTemp As Variant
    If mcllSquareBalance Is Nothing Then Exit Sub
    j = 1
    Do While True
        If j > mcllSquareBalance.Count Then Exit Do
        varTemp = mcllSquareBalance(j)
        If Val(varTemp(0)) = lng卡类别ID _
            And (lng消费卡ID = 0 Or (lng消费卡ID <> 0 And Val(varTemp(1)) = lng消费卡ID)) Then
            mcllSquareBalance.Remove j
        Else
            j = j + 1
        End If
    Loop
    If mcllSquareBalance.Count = 0 Then Set mcllSquareBalance = Nothing
End Sub

Private Sub cmdDel_Click()
    Dim int类型 As Integer, dblMoney As Double
    Dim lngCardTypeID As Long, lng消费卡ID As Long
    Dim objCard As Card
    Dim bln强制退现 As Boolean
    Dim str卡类别名称 As String
    
    '删除相关的费用
    With vsBlance
        If .Row < 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("删除标志"))) = 1 Then Exit Sub
        
        '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        int类型 = Val(.TextMatrix(.Row, .ColIndex("类型")))
        lngCardTypeID = Val(.TextMatrix(.Row, .ColIndex("卡类别ID")))
        lng消费卡ID = Val(.TextMatrix(.Row, .ColIndex("消费卡ID")))
        str卡类别名称 = Trim(.TextMatrix(.Row, .ColIndex("卡类别名称")))
        
        If int类型 = 3 And Val(.TextMatrix(.Row, .ColIndex("支付金额"))) <> 0 Then
            '105432
            Set objCard = GetPayCard(lngCardTypeID, False, False)
            If ThreeBalanceCheck(Me, mlngModule, objCard, mcllForceDelToCash, _
                str卡类别名称, bln强制退现) = False Then Exit Sub
        End If
        
        mobjDelBalance.原样退 = False
        If int类型 = 5 Then
            Set objCard = GetPayCard(lngCardTypeID, True, False)
            If objCard Is Nothing Then
                MsgBox "注意:" & vbCrLf & "未找到指定的消费卡,不能删除!", vbInformation, gstrSysName
                Exit Sub
            End If
            If objCard.是否退现 = 0 Then
                MsgBox "注意:" & vbCrLf & "    " & objCard.结算方式 & "不支持退现,不能删除!", vbInformation, gstrSysName
                Exit Sub
            End If
            Call ClearSquareBalance(lngCardTypeID, lng消费卡ID) '清除消费卡结算
            Call ClearReMoveSquareBalance(lngCardTypeID, lng消费卡ID)
        Else
            If lngCardTypeID = 0 Then
                Set objCard = GetPayCard(Trim(.TextMatrix(.Row, .ColIndex("支付方式"))), False, False)
            Else
                Set objCard = GetPayCard(lngCardTypeID, False, False)
            End If
            dblMoney = Val(.Cell(flexcpData, .Row, .ColIndex("支付金额")))
            mCurCarge.dbl当前未退 = RoundEx(mCurCarge.dbl当前未退 + dblMoney, 6)
            mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 - dblMoney, 6)
            If .Rows <= 2 Then
                .Clear 1
                .RowData(1) = ""
                .Cell(flexcpData, 1, 0, 1, .COLS - 1) = ""
            Else
                vsBlance.RemoveItem .Row
            End If
        End If
    End With
    Call Set退费方式(IIf(mCurCarge.dbl当前未退 <= 0, 2, 3), , , bln强制退现)
    Call Load退费方式(bln强制退现)
    Call SetDeleteVisible
    Call SetControlProperty(True)
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    mblnOK = False
    Call ExcuteMainReshData
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim blnUnload As Boolean
   
    '单据界面按了回车符
    If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    '再处理其他
    If isValied = False Then Exit Sub
    If txt缴款.Text <> "0.00" Then
        'LED显示
        Call ShowLedInfor
    End If
    If Not Execute原样退 Then Exit Sub
    '2.三卡交易检查
    '93114，从isValied()中放到这里来是因为如果再退费列表里缺省了三方卡，此时若再选择另外的三方卡，那么刷卡信息将会被覆盖
    If CheckThreeSwapValied(Nothing) = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Sub
    End If
    
    If ExecuteDelete(blnUnload) = False Then Exit Sub
    If blnUnload Then
        '刷新主界面信息
        ExcuteMainReshData
        Unload Me
    End If
End Sub

Private Sub ExcuteMainReshData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行主界面的刷新数据
    '编制:刘兴洪
    '日期:2014-06-17 15:09:44
    '说明:主要是应用医保刷新
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gfrmMain Is Nothing Then Exit Sub
    Call mfrmMain.zlExeBalanceWinRefrshData(mblnOK, mobjDelBalance)
End Sub

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的显示状态
    '编制:刘兴洪
    '日期:2014-07-08 19:12:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTemp As Boolean
    
    If mbytFunc = EM_FUN_退费 Then
        '医保进行结算了的,或非医保的,显示完成收费
        cmdOK.Visible = True
        '医保进行了结算后,不能退出
        cmdExit.Visible = mobjDelBalance.SaveBilled = False
        Exit Sub
     End If
     If mbytFunc = EM_FUN_重退 Then
        cmdExit.Caption = "退出(&E)"
        cmdOK.Visible = True: cmdExit.Visible = True
     End If
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call StartAndStop预存款
    Call cbo支付方式_Click
    Call SetControlProperty
    Call Set退费方式(IIf(mCurCarge.dbl当前未退 <= 0, 2, 3)): Call Load退费方式
    Call SetCtrlVisible
    mblnLoad = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
    Case vbKeyAdd, vbKeyF4
        If gTy_Module_Para.bln使用加减切换 = False And KeyCode = vbKeyAdd Then Exit Sub
        If Me.ActiveControl Is txt缴款 And cbo支付方式.Enabled Then
            i = cbo支付方式.ListIndex
            If i >= cbo支付方式.ListCount - 1 Then
                i = 0
            Else
                i = i + 1
            End If
            cbo支付方式.ListIndex = i
        End If
    Case vbKeySubtract
        If gTy_Module_Para.bln使用加减切换 = False And KeyCode = vbKeySubtract Then Exit Sub
        If Me.ActiveControl Is txt缴款 And cbo支付方式.Enabled Then
            i = cbo支付方式.ListIndex
            If i <= 0 Then
                i = cbo支付方式.ListCount - 1
            Else
                i = i - 1
            End If
            cbo支付方式.ListIndex = i
        End If
     Case vbKeyF12
            If Shift = vbCtrlMask Then
                '强制性LED报价,(合计)
                 Call LedVoiceSpeak
            End If
    Case vbKeyF2
        If cmdOK.Visible And cmdOK.Enabled Then
            cmdOK.SetFocus
            cmdOK_Click
        End If
    Case vbKeyReturn
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    '选检查主界面中是否发送了回车键的
    mblnCacheKeyReturn = (GetAsyncKeyState(VK_RETURN) And &H1) <> 0
    mstrTittle = "病人退费结算"
    
    
    RestoreWinState Me, App.ProductName, mstrTittle
    Call SetWindowsSize
    Set mrsOneCard = GetOneCard
    zlControl.CboSetWidth cbo支付方式.hWnd, cbo支付方式.Width * 2
    mblnFirst = True: mblnLoad = True
    mblnUnLoad = False
    zlControl.PicShowFlat picTotal, -1, , taCenterAlign
    zlControl.PicShowFlat Picture1, -1, , taCenterAlign
    zlControl.PicShowFlat picPay, -1, , taCenterAlign
    Call InitFace
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    'If Me.Width < 10530 Then Me.Width = 10530
    'If Me.Height < 7035 Then Me.Height = 7035
    With picBlance
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - stbThis.Height - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    With mCurCarge
           .dbl本次退预交 = 0
           .dbl退费合计 = 0
           .dbl已算误差 = 0
           .dbl本次医保退费 = 0
           .dbl已退合计 = 0
           .dbl本次应收 = 0
           .dbl当前未退 = 0
           .dbl费用余额 = 0
           .dbl可用预交 = 0
           .dbl预交余额 = 0
    End With
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    Set mrsClassMoney = Nothing
    With mCurBrushCard
        .dbl帐户余额 = 0
        .str交易流水号 = ""
        .str交易说明 = ""
        .str卡号 = ""
        .str扩展信息 = ""
        .str密码 = ""
    End With
    Set mrsUsedCards = Nothing
    SaveWinState Me, App.ProductName, mstrTittle
End Sub

 

 
Private Sub picBlance_Resize()
    Err = 0: On Error Resume Next
    With vsBlance
        .Left = picBlance.ScaleLeft
        .Width = picBlance.ScaleWidth
        .Height = picBlance.ScaleHeight - .Top
    End With
End Sub
 
Private Sub LoadPatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息
    '编制:刘兴洪
    '日期:2011-08-13 10:52:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    stbThis.Panels(2).Text = mobjDelBalance.姓名
    Set rsTemp = GetMoneyInfo(mobjDelBalance.病人ID, 0, False, 1, False, 0, True)
    Dim dbl家属余额 As Double
    With mCurCarge
        .dbl预交余额 = 0
        .dbl费用余额 = 0
        Do While Not rsTemp.EOF
            .dbl预交余额 = .dbl预交余额 + Val(Nvl(rsTemp!预交余额))
            .dbl费用余额 = .dbl费用余额 + Val(Nvl(rsTemp!费用余额))
            If Nvl(rsTemp!家属, 0) = 1 Then
                dbl家属余额 = Val(Nvl(rsTemp!预交余额)) - Val(Nvl(rsTemp!费用余额))
            End If
            rsTemp.MoveNext
        Loop
        .dbl可用预交 = .dbl预交余额 - .dbl费用余额
    End With
    If RoundEx(mCurCarge.dbl可用预交, 6) = 0 And RoundEx(dbl家属余额, 6) = 0 Then
        stbThis.Panels(3).Visible = False
    Else
        stbThis.Panels(3).Visible = True
        stbThis.Panels(3).Text = "预交:" & Format(mCurCarge.dbl可用预交, "0.00") & _
            IIf(dbl家属余额 > 0, "(含家属:" & Format(dbl家属余额, "0.00") & ")", "")
    End If
    
    lbl退费合计.Caption = Format(Abs(mCurCarge.dbl退费合计), "###0.00;-###0.00;0.00;0.00;")
End Sub

Private Sub LedVoiceSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:语音报价
    '编制:刘兴洪
    '日期:2011-08-13 16:38:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'If mCurBrushCard.int性质 <> 1 Then Exit Sub
    If gblnLED = False Then Exit Sub
    
    If mobjDelBalance.intInsure <> 0 Then Exit Sub
'    If mCurCarge.dbl退费合计 = 0 Then Exit Sub
'    If mCurCarge.dbl当前未退 = 0 Then Exit Sub

    If mCurCarge.dbl当前未退 < 0 Then
'        zl9LedVoice.Speak "#21 " & Format(-1 * lbl未退金额.Caption, "0.00")
    Else
        zl9LedVoice.Speak "#21 " & Format(lbl未退金额.Caption, "0.00")
    End If
    mbln已报价 = True
End Sub

 

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
   If Panel.Key = "Calc" Then
        mlngR = FindWindow("SciCalc", "计算器")
        If mlngR <> 0 Then
            BringWindowToTop mlngR
        Else
            On Error Resume Next
            Shell "calc.exe", vbNormalFocus
        End If
  End If
End Sub
Private Function zlGetClassMoney(ByRef lng结帐序号 As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle

'    If Not mrsClassMoney Is Nothing Then
'        Set rsMoney = mrsClassMoney: zlGetClassMoney = True: Exit Function
'    End If
    If lng结帐序号 = 0 Then
        Call mfrmMain.zlGetClassMoney(rsMoney)
        zlGetClassMoney = True: Exit Function
    End If
    '初始化数据结构
    Set mrsClassMoney = New ADODB.Recordset
    mrsClassMoney.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    mrsClassMoney.Fields.Append "金额", adDouble, , adFldIsNullable
    mrsClassMoney.CursorLocation = adUseClient
    mrsClassMoney.LockType = adLockOptimistic
    mrsClassMoney.CursorType = adOpenStatic
    mrsClassMoney.Open
    strSQL = "" & _
    "   Select  A.收费类别,nvl(sum(实收金额) ,0) as 金额   " & _
    "   From 门诊费用记录 A,(Select 结帐ID From 病人预交记录 where 结算序号=[1] ) B " & _
    "   Where A.结帐ID=B.结帐ID " & _
    "   Group by 收费类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐序号)

    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            mrsClassMoney.Find "收费类别='" & Nvl(!收费类别, "无") & "'", , adSearchForward, 1
            If mrsClassMoney.EOF Then mrsClassMoney.AddNew
            mrsClassMoney!收费类别 = Nvl(!收费类别, "无")
            mrsClassMoney!金额 = Val(Nvl(mrsClassMoney!金额)) + Val(Nvl(!金额))
            mrsClassMoney.Update
            .MoveNext
        Loop
    End With
    Set rsMoney = mrsClassMoney
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt缴款_Change()
    Call Show误差金额
    Call SetControlProperty
End Sub

Private Sub txt缴款_GotFocus()
    Dim strTittle As String
    '只以缴款作为收费结束条件时,必须输入缴款或0
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
    Select Case strTittle
    Case "缴款"
        If gTy_Module_Para.byt缴款控制 = 1 _
            Or gTy_Module_Para.byt缴款控制 = 3 _
            Or gTy_Module_Para.byt缴款控制 = 2 Then
            If Val(txt缴款.Text) = 0 And Me.ActiveControl Is txt缴款 Then txt缴款.Text = ""
        End If
    Case "退款"
    End Select
    Call SetControlProperty(True)
    '自动报价或手工报价时由热键激活
    If Not mbln已报价 Then Call LedVoiceSpeak
    zlControl.TxtSelAll txt缴款
End Sub

Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示Led信息
    '编制:刘兴洪
    '日期:2011-08-13 15:25:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer, objCard As Card
    Dim strTittle As String
    If gblnLED = False Then Exit Sub
    If mCurCarge.dbl退费合计 = 0 Then Exit Sub
    
    Call GetCurCard(objCard)
    If mCurCarge.dbl当前未退 <= 0 Then
        strTittle = "退款"
    Else
        strTittle = "缴款"
    End If
    
    Select Case strTittle
    Case "缴款"
        '只有缴现才显示
        If objCard.结算性质 = 1 Then
            zl9LedVoice.DispCharge mCurCarge.dbl当前未退, Val(txt缴款.Text), Val(txt找补.Text)
        Else
            Call zl9LedVoice.DisplayBank( _
                "合计:" & lbl退费合计.Caption & "元,应付:" & lbl未退金额.Caption & "元", _
                "收您:" & txt缴款.Text & "元" & IIf(Val(txt找补.Text) = 0, "", ",找您:" & Val(txt找补.Text) & "元"))
        End If
        zl9LedVoice.Speak "#22 " & Val(txt缴款.Text)
        zl9LedVoice.Speak "#23 " & Val(txt找补.Text)
        zl9LedVoice.Speak "#3"
    Case "退款"
        '只有缴现才显示
        If objCard.结算性质 = 1 Then
            zl9LedVoice.DispCharge mCurCarge.dbl当前未退, -1 * Val(txt缴款.Text), -1 * Val(txt找补.Text)
        Else
            Call zl9LedVoice.DisplayBank( _
                "合计:" & lbl退费合计.Caption & "元,应退:" & lbl未退金额.Caption & "元", _
                "退您:" & txt缴款.Text & "元" & IIf(Val(txt找补.Text) = 0, "", ",收零:" & Val(txt找补.Text) & "元"))
        End If
'        zl9LedVoice.Speak "#22 " & -1 * Val(txt缴款.Text)
'        zl9LedVoice.Speak "#23 " & -1 * Val(txt找补.Text)
    End Select
'    zl9LedVoice.Speak "#3"
End Sub

Private Sub LedDisplayBank(ByVal blnLedAsked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示已结算信息
    '入参:blnLedAsked-是否已报价
    '编制:刘兴洪
    '日期:2011-12-15 13:40:46
    '问题:52117
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double, i As Long
    Dim str医保 As String, str三方交易 As String, str老一卡通 As String, str普通结算 As String
    Dim varPara  As Variant, str结算方式 As String
    Dim strTittle As String
    If Not gblnLED Then Exit Sub
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
    With vsBlance
        '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        For i = 1 To .Rows - 1
            '医保交易
            If .TextMatrix(i, .ColIndex("支付方式")) <> "" Then
                Select Case Val(.TextMatrix(i, .ColIndex("类型")))
                Case 2 '医保
                    str医保 = str医保 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.Cell(flexcpData, i, .ColIndex("支付金额"))), "0.00")
                Case 3 '三方接口交易
                    str三方交易 = str三方交易 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.Cell(flexcpData, i, .ColIndex("支付金额"))), "0.00")
                Case 4   ' 一卡通交易
                    str老一卡通 = str老一卡通 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.Cell(flexcpData, i, .ColIndex("支付金额"))), "0.00")
                Case Else
                    str普通结算 = str普通结算 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.Cell(flexcpData, i, .ColIndex("支付金额"))), "0.00")
                End Select
            End If
        Next
    End With
     
    str结算方式 = ""
    If str医保 <> "" Then str结算方式 = str结算方式 & "||医保结算:||帐户余额:" & Format(mcur个帐余额, "0.00") & str医保
    If str三方交易 <> "" Then str结算方式 = str结算方式 & "||一卡通结算:" & str三方交易
    If str老一卡通 <> "" Then str结算方式 = str结算方式 & "||一卡通结算(老):" & str老一卡通
    If str普通结算 <> "" Then str结算方式 = str结算方式 & "||其他结算:" & str普通结算
    If str结算方式 = "" Then Exit Sub
    str结算方式 = Mid(str结算方式, 3)
    varPara = Split(str结算方式, "||")
    
    '目前最多只能显示10个参数值
    Select Case UBound(varPara)
    Case 0
          zl9LedVoice.DisplayBank varPara(0)
    Case 1
          zl9LedVoice.DisplayBank varPara(0), varPara(1)
    Case 2
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2)
    Case 3
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3)
    Case 4
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4)
    Case 5
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5)
    Case 6
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6)
    Case 7
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7)
    Case 8
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8)
    Case 9
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9)
    Case Else
        str结算方式 = ""
         For i = 10 To UBound(varPara)
            str结算方式 = str结算方式 & ";" & varPara(i)
        Next
        If str结算方式 > "" Then str结算方式 = Mid(str结算方式, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str结算方式
    End Select
    
    If blnLedAsked = False Then
        If strTittle = "退款" Then
'            zl9LedVoice.Speak "#21 " & Format(-1 * Val(lbl未退金额.Caption), "0.00")
        Else
            zl9LedVoice.Speak "#21 " & Format(Val(lbl未退金额.Caption), "0.00")
        End If
    End If
End Sub

Private Function Check缴款() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查缴款金额
    '返回:输入合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 10:30:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intIndex As Integer
    Dim strTittle As String
    
    On Error GoTo errHandle
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If intIndex <= 0 Then Exit Function
    
    Set objCard = mobjPayCards(intIndex)
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
    
    If txt缴款.Text <> "" Then
        If Abs(Val(txt缴款.Text)) > 999999999 Then
            MsgBox "输入的缴款金额过大,最大不能超过-999999999至99999999!", vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If Val(txt缴款.Text) = 0 Then
            If (objCard.接口序号 >= 0 Or objCard.结算性质 <> 1) _
                Or (objCard.结算性质 = 7 And objCard.接口序号 <= 0) Then
                '需要排除三方接口交易
                MsgBox "未输入" & strTittle & "金额,不能用" & objCard.结算方式 & "支付,请检查!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        Check缴款 = True
        Exit Function
    End If
    If CheckCashValied = False Then Exit Function
    Check缴款 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    Dim objCard As Card, strTittle As String
    
    If GetCurCard(objCard) = False Then Exit Sub
    
    zlControl.TxtCheckKeyPress txt缴款, KeyAscii, m金额式
    If KeyAscii <> 13 Then Exit Sub
    If mblnCacheKeyReturn = True Then mblnCacheKeyReturn = False
    KeyAscii = 0
    If Check缴款 = False Then Exit Sub
     
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
    
    
    '只以缴款作为收费结束条件时,必须输入缴款或0
    If gTy_Module_Para.byt缴款控制 = 1 _
        Or gTy_Module_Para.byt缴款控制 = 3 _
        Or gTy_Module_Para.byt缴款控制 = 2 Then
        If txt缴款.Text = "" Then Exit Sub
    End If
    
    If objCard.结算性质 <> 1 Then
        If (objCard.结算方式 Like "*支票*" Or _
            objCard.结算方式 Like "*卡*") And objCard.接口序号 <= 0 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        Call cmdOK_Click
        Call txt缴款_GotFocus
        Exit Sub
    End If
    
    If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
    Select Case strTittle
    Case "缴款"
        If txt缴款.Text <> "0.00" Then
            If Val(txt找补.Text) >= 0 Then
                 Call cmdOK_Click: Exit Sub
            End If
            MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
            txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
            Exit Sub
        End If
    Case "退款"
    End Select
    Call cmdOK_Click
End Sub


Private Sub txt缴款_LostFocus()
    Dim objCard As Card
    Dim dblTemp As Double
    
    If GetCurCard(objCard) = False Then
        Set objCard = New Card
    End If
    If mCurCarge.dbl当前未退 <= 0 Then
        '当前输入金额小于预交款剩余未退金额时，处理为两位小数
        dblTemp = GetOldBalanceMoney(1, objCard)
        If dblTemp > Val(txt缴款.Text) Then
            txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
        Else
            txt缴款.Text = FormatEx(Val(txt缴款.Text), 6, , , 2)
        End If
    Else
        txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
    End If
End Sub

Private Sub txt结算号码_GotFocus()
   zlControl.TxtSelAll txt结算号码
End Sub
Private Sub txt结算号码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt结算号码_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txt结算号码, KeyAscii, m文本式
End Sub

Private Sub txt摘要_GotFocus()
    zlControl.TxtSelAll txt摘要
    zlCommFun.OpenIme True
End Sub
Private Sub txt摘要_LostFocus()
    zlCommFun.OpenIme False
End Sub
Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub
Private Sub txt找补_GotFocus()
    zlControl.TxtSelAll txt找补
End Sub

Private Function ChargeDelOver(ByVal str退费结算 As String, _
    ByVal dbl预存款 As Double, ByRef dbl退支票额 As Double, _
    ByVal cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费完成
    '入参:blnNotCommit-是否没有进行事务提交，完成时再提交事务(原因是对普通病人进行一次提交)
    '编制:刘兴洪
    '日期:2011-08-15 15:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl缴款 As Double, dbl找补 As Double
    Dim cllPro As Collection, objCard As Card
    Dim strSQL As String, i As Long
     
    On Error GoTo errHandle
    
    If GetCurCard(objCard) = False Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    If objCard.结算性质 = 1 Then
        dbl缴款 = Val(txt缴款.Text)
        dbl找补 = Val(txt找补.Text)
    End If
    
    If dbl缴款 = 0 Then
        dbl缴款 = 0: dbl找补 = 0
    End If
    
    '调用之前,先处理数据
    'Zl_门诊退费结算_Modify
    strSQL = "Zl_门诊退费结算_Modify("
    '  操作类型_In   Number,
    '  --操作类型_In:
    '  --   0-原样退
    '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
    '  --   1-普通退费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款
    '  --     ③剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交
    '  --   2.三方卡退费结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退预交_In: 传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退预交_In: 传入零
    '  --   4-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②退预交_In: 传入零
    strSQL = strSQL & "" & 1 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mobjDelBalance.病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & mobjDelBalance.冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & str退费结算 & "',"
    '  冲预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & dbl预存款 & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & GetForceDelToCashNote(mcllForceDelToCash) & "',"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & dbl缴款 & ","
    '  找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & dbl找补 & ","
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "" & mCurCarge.dbl本次误差费 & ","
    '  完成退费_In   Number := 0,
    '0-未完成退费;1-异常完成退费;2-完成退费
    strSQL = strSQL & "2,"
    '77141,冉俊明,2014-8-26,给零费用病人收费/退费后,没有结算信息
    '  原结帐id_In   病人预交记录.结帐id%Type := Null,
    strSQL = strSQL & "null,"
    '  剩余转预交_In Number:=0,
    strSQL = strSQL & "0,"
    '  缺省结算方式_In 结算方式.名称%Type := Null,
    strSQL = strSQL & "'" & Trim(cbo支付方式.Text) & "',"
    '  冲预交病人ids_In Varchar2 := Null
    strSQL = strSQL & "'" & mobjDelBalance.家属IDs & "')"
    zlAddArray cllPro, strSQL
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    mobjDelBalance.缴款 = dbl缴款: mobjDelBalance.找补 = dbl找补
    Set cllBillPro = New Collection
    
    ChargeDelOver = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Show误差金额()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示误差金额
    '编制:刘兴洪
    '日期:2014-07-09 18:44:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl退支票额 As Double
    Dim dbl剩余金额 As Double, dblTemp As Double
    Dim objCard As Card, strTittle As String
    
    On Error GoTo errHandle
    
    If GetCurCard(objCard) = False Then Exit Sub
    
    strTittle = IIf(mCurCarge.dbl当前未退 <= 0, "退款", "缴款")
    
    mCurCarge.dbl本次误差费 = 0
    
    dblMoney = IIf(strTittle = "退款", -1, 1) * Val(txt缴款.Text)
    dbl剩余金额 = RoundEx(mCurCarge.dbl当前未退 - dblMoney, 6)
    
    If RoundEx(mCurCarge.dbl已算误差, 6) = RoundEx(mCurCarge.dbl当前未退, 6) Then
        mCurCarge.dbl本次误差费 = mCurCarge.dbl当前未退
    Else
        If objCard.结算性质 = -99 Then
            mCurCarge.dbl本次误差费 = mCurCarge.dbl退费合计 - mCurCarge.dbl已退合计 - RoundEx(mCurCarge.dbl当前未退, 2)
        ElseIf objCard.结算性质 = 1 Then
            '现金
            dblTemp = IIf(dblMoney = 0, dbl剩余金额, mCurCarge.dbl当前未退): dbl剩余金额 = 0
            If mobjDelBalance.intInsure > 0 Then  '问题:43855
                If mInsurePara.分币处理 Then
                    dblMoney = CentMoney(CCur(dblTemp))
                Else
                    dblMoney = Format(dblTemp, "0.00")
                End If
            Else
                 dblMoney = CentMoney(CCur(dblTemp))
            End If
            mCurCarge.dbl本次误差费 = mCurCarge.dbl退费合计 - mCurCarge.dbl已退合计 - dblMoney
        Else
            mCurCarge.dbl本次误差费 = mCurCarge.dbl退费合计 - mCurCarge.dbl已退合计 - RoundEx(mCurCarge.dbl当前未退, 2)
        End If
    End If
    
    mCurCarge.dbl本次误差费 = RoundEx(mCurCarge.dbl本次误差费, 6)
    pic误差.Visible = mCurCarge.dbl本次误差费 <> 0
    lbl误差额.Caption = FormatEx(mCurCarge.dbl本次误差费, 6, , , 2)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CheckMulitInterfaceNum() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是正同时存在两种以上接口(不含两种)
    '返回:不含两种以上接口的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int性质 As Integer, str结算方式 As String
    Dim varData As Variant, strErrMsg As String
    Dim objCard As Card, strTittle As String
    
    On Error GoTo errHandle
    strErrMsg = ""
    
    If GetCurCard(objCard) = False Then Exit Function
    If objCard.结算性质 = -99 Or objCard.接口序号 <= 0 Then
        CheckMulitInterfaceNum = True: Exit Function
    End If
   '医保算一个接口
   If mobjDelBalance.intInsure <> 0 Then intCount = intCount + 1: strErrMsg = strErrMsg & "医保结算:" & mobjDelBalance.医保结算金额
   With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            'rowdata:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If InStr("34", int性质) > 0 Then
                If int性质 = 4 Then intCount = intCount + 1
                If int性质 = 3 Then '三方接口
                    intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str结算方式 & ":" & .Cell(flexcpData, i, .ColIndex("支付金额"))
                End If
            End If
        Next
    End With
    If intCount > 2 Then
        Call MsgBox("注意:" & vbCrLf & "   本系统目前只支持两种以下接口,现在已经存在如下接口交易:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    CheckMulitInterfaceNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ExecuteDelete(Optional ByRef blnUnload As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存结算数据
    '入参:blnUnload-是否收费完成，退出后，将Unload界面
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-10 09:53:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDel As Boolean, blnHaveMoney As Boolean
    Dim dblMoney As Double, dbl退支票额 As Double
    Dim objCard As Card, strTittle As String, str退费结算 As String
    Dim dbl剩余金额 As Double, dblTemp As Double, dbl预存款 As Double
    Dim j As Long, i As Long, strCardNo As String
    Dim cllBalance As Collection
    
    On Error GoTo errHandle
    blnUnload = False
    If CheckMulitInterfaceNum = False Then Exit Function
    
    If GetCurCard(objCard) = False Then
        MsgBox lblPayType.Caption & "方式未选择!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    blnDel = mCurCarge.dbl当前未退 <= 0
    mobjDelBalance.退费结算 = ""
    
    dblMoney = IIf(blnDel, -1, 1) * Val(txt缴款.Text)
    
    dbl退支票额 = 0
    dbl剩余金额 = mCurCarge.dbl当前未退 - dblMoney - mCurCarge.dbl本次误差费
    
    
    If objCard.结算性质 = -99 Then
        mobjDelBalance.退费结算 = mobjDelBalance.退费结算 & "|" & IIf(blnDel, "退预交款:", "冲预交:") & dblMoney
    ElseIf objCard.结算性质 = 1 Then
        If RoundEx(mCurCarge.dbl已算误差, 6) = RoundEx(mCurCarge.dbl当前未退, 6) Then
            dblMoney = 0
        Else
            dblTemp = IIf(dblMoney = 0, dbl剩余金额, mCurCarge.dbl当前未退): dbl剩余金额 = 0
            If mobjDelBalance.intInsure > 0 Then
                If gclsInsure.GetCapability(support分币处理, , mobjDelBalance.intInsure) Then
                    dblMoney = CentMoney(CCur(dblTemp))
                Else
                    dblMoney = Format(dblTemp, "0.00")
                End If
            Else
                dblMoney = CentMoney(CCur(dblTemp))
            End If
        End If
        
        If Val(txt缴款.Text) <> 0 Then
            mobjDelBalance.退费结算 = mobjDelBalance.退费结算 & "|缴款:" & IIf(blnDel, -1, 1) * Val(txt缴款.Text) & ":1"
            mobjDelBalance.退费结算 = mobjDelBalance.退费结算 & "|找补:" & IIf(blnDel, -1, 1) * Val(txt找补.Text) & ":2"
        End If
        mobjDelBalance.退费结算 = mobjDelBalance.退费结算 & "|" & objCard.结算方式 & ":" & dblMoney
        
    ElseIf objCard.结算方式 Like "*支票*" Then
        mobjDelBalance.退费结算 = mobjDelBalance.退费结算 & "|" & objCard.结算方式 & ":" & dblMoney
        If blnDel = False Then
            '问题:58344
            '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
            If RoundEx(dbl剩余金额, 2) < 0 Then
                If mstr退支票 = "" Then
                    MsgBox "在结算方式中没有设置应付款的结算方式,不能进行退支票处理", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                dbl退支票额 = -1 * Val(txt找补.Text)
                mobjDelBalance.退费结算 = mobjDelBalance.退费结算 & "|" & mstr退支票 & ":" & -1 * dbl退支票额 & ":2"
            End If
        Else
            If RoundEx(dbl剩余金额, 2) > 0 Then
                MsgBox objCard.结算方式 & "必须全退!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        mobjDelBalance.退费结算 = mobjDelBalance.退费结算 & "|" & objCard.结算方式 & ":" & dblMoney
    End If
    Call Show误差金额
    
    If objCard.结算性质 = 1 Then
        '误差不能大于10块钱
        If Abs(mCurCarge.dbl本次误差费) > 1.5 Then
            Call MsgBox("误差过大,请检查是否正确!", vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    If RoundEx(dbl剩余金额, 2) <> 0 Then blnHaveMoney = True
    If blnHaveMoney = False And dblMoney = 0 Then GoTo GoOver:
     
    If blnDel Then
        '退老版一卡通
        If ExecuteOneCardDelInterface(objCard, -1 * dblMoney, mcllDelPro) = False Then Exit Function
        '退三方卡交易
        If ExecuteThreeSwapDelInterface(objCard, -1 * dblMoney, mcllDelPro) = False Then Exit Function
    Else
        '用老版一卡通支付
        If ExecuteOneCardPayInterface(objCard, dblMoney, mcllDelPro) = False Then Exit Function
        '用一卡通支付(三方交易)
        If ExecuteThreeSwapPayInterface(objCard, dblMoney, mcllDelPro) = False Then Exit Function
        '用消费卡支付
        If ExecuteSquarePayInterface(objCard, dblMoney, mcllDelPro) = False Then Exit Function
    End If
    Call SetCtrlVisible
    
GoOver:
    If Not blnHaveMoney Then
         
         dbl预存款 = 0: str退费结算 = Get退费结算(dblMoney, dbl预存款)
        '退消费卡费
        If ExecuteSquareDelInterface(mcllSquareBalance, mcllDelPro) = False Then Exit Function
        Set mcllSquareBalance = Nothing '执行成功，清空集合
        If ChargeDelOver(str退费结算, dbl预存款, dbl退支票额, mcllDelPro) = False Then Exit Function
        mblnOK = True: ExecuteDelete = True: mblnUnloaded = True
        blnUnload = True
        Exit Function
    End If
    If objCard.结算性质 = 1 Then
       '现金
        ExecuteDelete = True: Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle:
    With vsBlance
        If objCard.消费卡 Then
            Call AddSquareBalance(objCard, blnDel)
        Else
            If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
                .Rows = .Rows + 1
                .RowPosition(.Rows - 1) = 1
            End If
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            .RowData(1) = 0
            strCardNo = mCurBrushCard.str卡号
                
            If objCard.结算性质 = -99 Then
                .TextMatrix(1, .ColIndex("支付方式")) = IIf(blnDel, "退预存款", "冲预存款")
                .RowData(1) = 1
                .TextMatrix(1, .ColIndex("删除标志")) = 0  '是否允许编辑:1-禁止编辑;0-不禁止编辑
                .TextMatrix(1, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
                .TextMatrix(1, .ColIndex("是否验证")) = 1
            ElseIf objCard.接口序号 > 0 Then
                .TextMatrix(1, .ColIndex("支付方式")) = objCard.结算方式
                ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                .Cell(flexcpData, 1, .ColIndex("支付方式")) = objCard.接口序号 & "|" & 3 & "|" & objCard.自制卡 & "|" & objCard.是否全退 & "|" & objCard.是否退现 & "|" & objCard.名称
                .RowData(1) = 3
                strCardNo = gobjSquare.objSquareCard.zlGetCardNODencode(mCurBrushCard.str卡号, objCard.接口序号, objCard.消费卡)
                .TextMatrix(1, .ColIndex("删除标志")) = 1  '是否允许编辑:1-禁止编辑;0-不禁止编辑
                .TextMatrix(1, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
                .Cell(flexcpBackColor, 1, 0, 1, .COLS - 1) = Me.BackColor
            ElseIf objCard.结算性质 = 7 And objCard.接口序号 <= 0 Then '老一卡通
                .TextMatrix(1, .ColIndex("支付方式")) = objCard.结算方式
                ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                .Cell(flexcpData, 1, .ColIndex("支付方式")) = objCard.接口序号 & "|" & 3 & "|" & objCard.自制卡 & "|" & objCard.是否全退 & "|" & objCard.是否退现 & "|" & objCard.名称
                .TextMatrix(1, .ColIndex("删除标志")) = 1  '是否允许编辑:1-禁止编辑;0-不禁止编辑
                .TextMatrix(1, .ColIndex("结算状态")) = 1  '是否已结算:1-已结算;0-未结算
                .Cell(flexcpBackColor, 1, 0, 1, .COLS - 1) = Me.BackColor
                .RowData(1) = 4
            Else
                .TextMatrix(1, .ColIndex("支付方式")) = objCard.结算方式
                .TextMatrix(1, .ColIndex("删除标志")) = 0  '是否允许编辑:1-禁止编辑;0-不禁止编辑
                .TextMatrix(1, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
            End If
            .TextMatrix(1, .ColIndex("类型")) = Val(.RowData(1))
            .TextMatrix(1, .ColIndex("结算性质")) = objCard.结算性质
            .TextMatrix(1, .ColIndex("卡类别ID")) = objCard.接口序号
            .TextMatrix(1, .ColIndex("消费卡ID")) = 0
            
            .TextMatrix(1, .ColIndex("支付金额")) = FormatEx(-1 * dblMoney, 6, , , 2)
            .Cell(flexcpData, 1, .ColIndex("支付金额")) = FormatEx(dblMoney, 6)
            .TextMatrix(1, .ColIndex("结算号码")) = IIf(txt结算号码.Visible, Trim(txt结算号码.Text), "")
            .TextMatrix(1, .ColIndex("备注")) = Trim(txt摘要.Text)
            
            If objCard.接口序号 > 0 Then
                .TextMatrix(1, .ColIndex("卡号")) = IIf(objCard.卡号密文规则 <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("卡号")) = mCurBrushCard.str卡号
                .TextMatrix(1, .ColIndex("交易流水号")) = mCurBrushCard.str交易流水号
                .TextMatrix(1, .ColIndex("交易说明")) = mCurBrushCard.str交易说明
                .TextMatrix(1, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
                .TextMatrix(1, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
                .TextMatrix(1, .ColIndex("是否转帐及代扣")) = IIf(objCard.是否转帐及代扣, 1, 0)
                .TextMatrix(1, .ColIndex("卡类别名称")) = objCard.名称
            End If
            
            mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 + dblMoney, 6)
            mCurCarge.dbl当前未退 = RoundEx(mCurCarge.dbl当前未退 - dblMoney, 6)
        End If
        
        '移除当前结算方式
        If Not objCard.消费卡 Or (objCard.消费卡 And blnDel) Then
            Call Set退费方式(IIf(mCurCarge.dbl当前未退 <= 0, 2, 3))
            Call Load退费方式
        Else
            Call SetControlProperty(True)
            txt缴款.Text = ""
        End If
        
        cbo支付方式.Enabled = True '只使用医疗卡或消费卡结算，退费时支付方式先是被禁用了的
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        Call LedDisplayBank(False)
    End With
    Call SetDeleteVisible
    ExecuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txt找补_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt找补_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl找补.Caption <> "找补" Then Exit Sub
    zlCommFun.ShowTipInfo txt找补.hWnd, "", False
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    If NewRow < 0 Then Exit Sub
    Call SetDeleteVisible
End Sub
Private Sub SetDeleteVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置删除控件的visible属性
    '编制:刘兴洪
    '日期:2014-07-10 11:26:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    
    With vsBlance
        If .Row < 0 Then
            blnEdit = False
        Else
             '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
           blnEdit = (Val(.TextMatrix(.Row, .ColIndex("是否退现"))) = 1 And InStr(1, "54", Val(.RowData(.Row))) <> 0) _
                Or (Val(.RowData(.Row)) = 0 And .TextMatrix(.Row, .ColIndex("支付方式")) <> "") _
                Or InStr(1, "13", Val(.RowData(.Row))) > 0
           blnEdit = blnEdit And Val(.TextMatrix(.Row, .ColIndex("删除标志"))) <> 1    '是否允许编辑:1-禁止编辑;0-不禁止编辑
        End If
    End With
    cmdDel.Visible = blnEdit
End Sub

Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体大小
    '编制:刘兴洪
    '日期:2014-07-10 11:27:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If OS.IsDesinMode Then Exit Sub
    '最小窗体尺寸
    With gWinRect
        .MaxW = Me.Width
        .MaxH = Screen.Height * Screen.TwipsPerPixelY
        .MinH = Me.Height
        .MinW = Me.Width
    End With
    glngOld = GetWindowLong(hWnd, GWL_WNDPROC)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SetWindowResizeWndMessage)
End Sub

Private Sub SetControlColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的颜色
    '编制:刘兴洪
    '日期:2014-07-10 11:32:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt缴款.BackColor = IIf(txt缴款.Enabled, &H80000005, Me.BackColor)
    txt找补.BackColor = Me.BackColor
    txt结算号码.BackColor = IIf(txt结算号码.Enabled, &H80000005, Me.BackColor)
    txt摘要.BackColor = IIf(txt摘要.Enabled, &H80000005, Me.BackColor)
End Sub
Public Function Get退费结算(ByVal dblCurDelMoney As Double, _
    ByRef dbl预存款 As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取退费结算数据
    '入参:dblCurDelMoney-当前退费金额
    '出参:dbl预存款-返回本次支付的预款
    '返回:收费用结算方式,格式如下:
    '       结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
    '编制:刘兴洪
    '日期:2014-07-10 11:33:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, i As Integer, int性质 As Integer
    Dim str退费结算 As String, objCard As Card
    Dim dblMoney As Double, blnDel As Double
    
    
    '结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
    '收费完成
    blnDel = IIf(mCurCarge.dbl当前未退 <= 0, True, False)
    str退费结算 = ""
    With vsBlance
        dbl预存款 = 0
        For i = .Rows - 1 To 1 Step -1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            ' 0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            
            If str结算方式 <> "" Then
                Select Case int性质
                Case 0 '普通结算
                  If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                    str退费结算 = str退费结算 & "||" & str结算方式
                    str退费结算 = str退费结算 & "|" & Val(.Cell(flexcpData, i, .ColIndex("支付金额")))
                    str退费结算 = str退费结算 & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("结算号码"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("结算号码"))))
                    str退费结算 = str退费结算 & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("备注"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("备注"))))
                  End If
                Case 1 '预存款
                     dbl预存款 = Val(.Cell(flexcpData, i, .ColIndex("支付金额")))
                End Select
            End If
        Next
        
        If GetCurCard(objCard) = False Then Exit Function
        dblMoney = dblCurDelMoney
        If RoundEx(dblMoney, 6) <> 0 And objCard.接口序号 <= 0 Then
            If objCard.结算性质 <> -99 Then
                str退费结算 = str退费结算 & "||" & objCard.结算方式
                If objCard.结算性质 = 1 Then
                    '现金
                    str退费结算 = str退费结算 & "|" & dblMoney
                    str退费结算 = str退费结算 & "| "
                    str退费结算 = str退费结算 & "| "
                Else
                    str退费结算 = str退费结算 & "|" & dblMoney
                    str退费结算 = str退费结算 & "|" & IIf(Trim(txt结算号码) = "", " ", Trim(txt结算号码))
                    str退费结算 = str退费结算 & "|" & IIf(Trim(txt摘要) = "", " ", Trim(txt摘要))
                End If
            Else
                 dbl预存款 = RoundEx(dbl预存款 + dblMoney, 6)
            End If
        End If
    End With
    If str退费结算 <> "" Then str退费结算 = Mid(str退费结算, 3)
    Get退费结算 = str退费结算
End Function

Private Function GetCurCard(ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前卡
    '出参:objCard-返回当前退款或缴款的卡对象
    '返回:成功,返回卡对象
    '编制:刘兴洪
    '日期:2014-07-09 11:03:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    On Error GoTo errHandle
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    GetCurCard = True
    Exit Function
errHandle:
    Set objCard = New Card
End Function
Private Function ExecuteOneCardPayInterface(ByVal objCard As Card, ByVal dblMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(老版本)
    '入参:lng结算序号-按结算序号进行处理
    '     dblMoney-本次支付金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 10:42:15
    '说明:接口内部进行事务控制
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl余额 As Double, str医院编码 As String
    Dim i As Long, strSQL As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim intCardType As Integer, strSwapNO As String
    '非一卡通支付,直接返回
    If objCard.结算性质 <> 7 Then ExecuteOneCardPayInterface = True: Exit Function

    mrsOneCard.Filter = "结算方式='" & objCard.结算方式 & "'"
    If mrsOneCard.EOF Then
        MsgBox objCard.结算方式 & "未启用,请在『基础参数设置』中设置启用!", vbInformation, gstrSysName
        ExecuteOneCardPayInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    '调用之前,先处理数据
    'Zl_门诊退费结算_Modify
    strSQL = "Zl_门诊退费结算_Modify("
    '  操作类型_In   Number,
    '  --操作类型_In:
    '  --   0-原样退
    '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
    '  --   1-普通退费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款
    '  --     ③剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交
    '  --   2.三方卡退费结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退预交_In: 传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   4-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    strSQL = strSQL & "" & 2 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mobjDelBalance.病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & mobjDelBalance.冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & objCard.结算方式 & "|" & dblMoney & "| | " & "',"
    '  冲预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str卡号 & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str交易说明 & "')"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成退费_In   Number := 0,
    '  原结帐id_In   病人预交记录.结帐id%Type := Null,
    '  剩余转预交_In Number:=0
    zlAddArray cllPro, strSQL
    
    '一卡通结算
    blnTrans = True
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If Not mobjICCard.PaymentSwap(dblMoney, dbl余额, intCardType, Val("" & mrsOneCard!医院编码), mCurBrushCard.str卡号, mCurBrushCard.str交易流水号, mobjDelBalance.结帐ID, mobjDelBalance.病人ID) Then
        gcnOracle.RollbackTrans
        MsgBox objCard.结算方式 & "结算失败!", vbOKOnly, gstrSysName
        Exit Function
    End If
    gstrSQL = "Zl_一卡通结算_Update(" & 0 & ",'" & objCard.结算方式 & "','" & mCurBrushCard.str卡号 & "','" & intCardType & "','" & strSwapNO & "'," & dbl余额 & "," & mobjDelBalance.结算序号 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    blnTrans = False
    ExecuteOneCardPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
 End Function
 Private Function ExecuteOneCardDelInterface(ByVal objCard As Card, ByVal dblDelMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通退费接口(老版)
    '入参:cllBillPro-保存单据的SQL
    '编制:刘兴洪
    '日期:2014-07-10 10:36:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String '医院编码
    Dim i As Long, dblMoney As Double, strNos As String, strSQL As String
    Dim str结算方式 As String
    Dim cllPro As Collection, blnTrans As Boolean
    '非一卡通支付,直接返回
    If objCard.结算性质 <> 7 Then ExecuteOneCardDelInterface = True: Exit Function

    mrsOneCard.Filter = "结算方式='" & objCard.结算方式 & "'"
    If mrsOneCard.EOF Then
        MsgBox objCard.结算方式 & "未启用,请在『基础参数设置』中设置启用!", vbInformation, gstrSysName
        ExecuteOneCardDelInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    On Error GoTo errHandle
    If mobjDelBalance.rsBalance Is Nothing Then Exit Function
    If mobjDelBalance.rsBalance.State <> 1 Then Exit Function
    mobjDelBalance.rsBalance.Filter = "类型=4"
    If mobjDelBalance.rsBalance.RecordCount = 0 Then Exit Function
    With mobjDelBalance.rsBalance
        .MoveFirst
        Do While Not .EOF
            dblMoney = dblMoney + Val(Nvl(mobjDelBalance.rsBalance!冲预交))
            .MoveNext
        Loop
        .MoveFirst
    End With
    dblMoney = RoundEx(dblMoney, 6)
    If RoundEx(dblMoney, 6) = 0 Then Exit Function
    
    If dblDelMoney <> dblMoney Then
        MsgBox objCard.结算方式 & " 必须全退!" & vbCrLf & "原结算金额:" & Format(dblMoney, "0.00") & vbCrLf & " 现退款金额:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    '一卡通(旧):只能使用一种
    With mobjDelBalance.rsBalance
        strCardNo = Nvl(!卡号)
        str结算方式 = Nvl(!结算方式)
        
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = str结算方式 & "|" & -1 * dblMoney
        str结算方式 = str结算方式 & "|" & IIf(Trim(Nvl(!结算号码)) = "", " ", Trim(Nvl(!结算号码)))
        str结算方式 = str结算方式 & "| "
        
        'Zl_门诊退费结算_Modify
        strSQL = "Zl_门诊退费结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & 2 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
         
        strSQL = strSQL & "" & mobjDelBalance.病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mobjDelBalance.冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  退预交_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & Nvl(!交易流水号) & "',"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & Nvl(!交易说明) & "')"
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        '  完成退费_In   Number := 0,
        '  原结帐id_In   病人预交记录.结帐id%Type := Null
    End With
    zlAddArray cllPro, strSQL
    
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    Err = 0: On Error GoTo ErrRoll:
    If Not mobjICCard.ReturnSwap(strCardNo, strHsptCode, strSwap, dblMoney) Then
        gcnOracle.RollbackTrans
        MsgBox "一卡通退费交易调用失败,不能继续退费操作！", vbExclamation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    ExecuteOneCardDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ExecuteThreeSwapPayInterface(objCard As Card, ByVal dblMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:lng结算序号-按结算序号进行处理
    '     dblMoney-本次支付金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long, strSQL As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    '调用之前,先处理数据
    'Zl_门诊退费结算_Modify
    strSQL = "Zl_门诊退费结算_Modify("
    '  操作类型_In   Number,
    '  --操作类型_In:
    '  --   0-原样退
    '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
    '  --   1-普通退费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款
    '  --     ③剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交
    '  --   2.三方卡退费结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退预交_In: 传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   4-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    strSQL = strSQL & "" & 2 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mobjDelBalance.病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & mobjDelBalance.冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & objCard.结算方式 & "|" & dblMoney & "| | " & "',"
    '  冲预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & objCard.接口序号 & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str卡号 & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str交易说明 & "')"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成退费_In   Number := 0,
    '  原结帐id_In   病人预交记录.结帐id%Type := Null,
    '  剩余转预交_In Number:=0
    zlAddArray cllPro, strSQL
    
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
    
    str结帐IDs = mobjDelBalance.冲销ID
    str结帐IDs = str结帐IDs & IIf(mobjDelBalance.结帐ID <> 0, "," & mobjDelBalance.结帐ID, "")
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, _
         str结帐IDs, _
        mobjDelBalance.CurDelNos, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    mCurBrushCard.str交易流水号 = strSwapGlideNO
    mCurBrushCard.str交易说明 = strSwapMemo
    If objCard.消费卡 = False Then
        Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    End If
    Call zlAddThreeSwapSQLToCollection(False, str结帐IDs, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    
    '77156,冉俊明,2014-8-26,普通病人使用银行卡退费后，还可以点击返回按钮导致产生了退费的异常单据
    mobjDelBalance.SaveBilled = True
    ExecuteThreeSwapPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ExecuteThreeSwapDelInterface(ByVal objCard As Card, ByVal dblDelMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:lng结算序号-按结算序号进行处理
    '     dblMoney-本次支付金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblMoney As Double, str结算方式   As String
    Dim strTemp As String, strXMLExpend As String, strSwapExtendInfor As String
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    Err = 0: On Error GoTo Errhand:
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then ExecuteThreeSwapDelInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    '结算方式|结算金额|结算号码|结算摘要||..
    str结算方式 = objCard.结算方式 & "|" & -1 * dblDelMoney
    str结算方式 = str结算方式 & "|" & IIf(Trim(txt结算号码.Text) = "", " ", Trim(txt结算号码.Text))
    str结算方式 = str结算方式 & "|" & IIf(Trim(txt摘要.Text) = "", " ", Trim(txt摘要.Text))

    'Zl_门诊退费结算_Modify
    strSQL = "Zl_门诊退费结算_Modify("
    '  操作类型_In   Number,
    '  --操作类型_In:
    '  --   0-原样退
    '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
    '  --   1-普通退费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款
    '  --     ③剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交
    '  --   2.三方卡退费结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退预交_In: 传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   4-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  -- 误差金额_In:存在误差费时,传入
    '  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费
    '  -- 原结帐ID_IN:原样退时,传入(如果原样退未传入时,则以最后一次结帐为准)
    
    strSQL = strSQL & "" & 2 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mobjDelBalance.病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & mobjDelBalance.冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  退预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & objCard.接口序号 & ","
        '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str卡号 & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str交易说明 & "')"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成退费_In   Number := 0,
    '  原结帐id_In   病人预交记录.结帐id%Type := Null
    zlAddArray cllPro, strSQL
    
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    str结帐IDs = mobjDelBalance.冲销ID & IIf(mobjDelBalance.结帐ID <> 0, "," & mobjDelBalance.结帐ID, "")
    
    '81489,冉俊明,2015-1-22,退费传入冲销ID
    strSwapExtendInfor = "3|" & str结帐IDs: strTemp = strSwapExtendInfor
    
    '93114，退费时使用转帐方式
    If CheckThreeSwapCanTransfer(objCard, mobjDelBalance.原结帐ID) Then
        'zlTransferAccountsMoney
        '参数名  参数类型    入/出   备注
        'frmMain Object  In  调用的主窗体
        'lngModule   Long    In  HIS调用模块号
        'lngCardTypeID   Long    In  卡类别ID
        'strCardNo   String  In  卡号
        'strBalanceID    String  In  结算ID
        'dblMoney    Double  In  转帐金额
        'strSwapGlideNO  String  Out 交易流水号
        'strSwapMemo String  Out 交易说明
        'strSwapExtendInfor  String  In 退费业务时，传入本次退费的冲销ID:
        '                               格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                               收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '                           Out 交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
        'strXMLExpend String In   XML串:
        '                            <IN>
        '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务
        '                            </IN>
        '                    Out  XML串:
        '                            <OUT>
        '                               <ERRMSG>错误信息</ERRMSG >
        '                            </OUT>
        '    Boolean 函数返回    True:调用成功,False:调用失败
        '说明:
        '１. 在医保补充结算时进行的三方转帐时调用。
        '２. 一般来说，成功转帐后，都应该打印相关的结算票据，可以放在此接口进行处理.
        '３. 在转帐成功后，返回交易流水号和相关交易说明；如果存在其他交易信息，可以放在扩展信息中返回.
        '构造XML串
        strXMLExpend = "<IN><CZLX>4</CZLX></IN>"
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModule, objCard.接口序号, mCurBrushCard.str卡号, _
            mobjDelBalance.原结帐ID, dblDelMoney, mCurBrushCard.str交易流水号, mCurBrushCard.str交易说明, strSwapExtendInfor, strXMLExpend) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, mCurBrushCard.str交易流水号, mCurBrushCard.str交易说明, cllUpdate, 2)
    Else
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
        '       strSwapNo-交易流水号(扣款时的交易流水号)
        '       strSwapMemo-交易说明(扣款时的交易说明)
        '       strSwapExtendInfor-出入，本次退费的冲销ID：
        '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '       strSwapExtendInfor-传出，交易的扩展信息
        '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
        If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, _
            "3|" & mobjDelBalance.原结帐ID, dblDelMoney, mCurBrushCard.str交易流水号, mCurBrushCard.str交易说明, strSwapExtendInfor) = False Then gcnOracle.RollbackTrans: Exit Function
        'Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, strCardNO, strSwapNO, strSwapMemo, cllUpdate, 2)
    End If
    
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, str结帐IDs, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    
    '77156,冉俊明,2014-8-26,普通病人使用银行卡退费后，还可以点击返回按钮导致产生了退费的异常单据
    mobjDelBalance.SaveBilled = True
    ExecuteThreeSwapDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Private Sub SumSquareBalance(cllSquareCard As Collection)
'    '按消费卡ID和卡号进行分组求和，获得新的集合
'    Dim cllTemp1 As New Collection, cllTemp2 As New Collection
'    Dim strCards As String, strCard As String
'    Dim varCard As Variant, varTemp As Variant
'    Dim dblSumMoney As Double
'    Dim j As Integer, i As Integer
'
'    On Error GoTo errHandle:
'    If cllSquareCard Is Nothing Then Exit Sub
'    If cllSquareCard.Count = 0 Then Exit Sub
'    '集合不能直接赋值
'    For i = 1 To cllSquareCard.Count
'        cllTemp1.Add cllSquareCard(i)
'        cllTemp2.Add cllSquareCard(i)
'    Next
'
'    Set cllSquareCard = New Collection
'    For i = 1 To cllTemp1.Count
'        'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
'        varCard = cllTemp1(i): dblSumMoney = 0
'        strCard = varCard(0) & "|" & varCard(1) & "|" & varCard(3)
'        If InStr(strCards & "||", "||" & strCard & "||") = 0 Then
'            strCards = strCards & "||" & strCard
'            For j = 1 To cllTemp2.Count
'                varTemp = cllTemp2(j)
'                If strCard = varTemp(0) & "|" & varTemp(1) & "|" & varTemp(3) Then
'                    dblSumMoney = dblSumMoney + Val(varTemp(2))
'                End If
'            Next
'            cllSquareCard.Add Array(varCard(0), varCard(1), RoundEx(dblSumMoney, 6), varCard(3), varCard(4), varCard(5), varCard(6))
'        End If
'    Next
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Sub

Private Function ExecuteSquareDelInterface(ByVal cllSquareCard As Collection, _
    ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡支付
    '入参:lng结算序号-按结算序号进行处理
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     cllSquareCard-本次退的消费卡集(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean, dblDelMoney As Double
    Dim str结帐IDs As String, i As Long, varTemp As Variant
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, dblMoney As Double, str结算方式  As String
    Dim objCard As Card
 
    '无消费卡结算退款,返回true
    If cllSquareCard Is Nothing Then ExecuteSquareDelInterface = True: Exit Function
    If cllSquareCard.Count = 0 Then ExecuteSquareDelInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo Errhand:
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    If mobjDelBalance.rsBalance Is Nothing Then Exit Function
    If mobjDelBalance.rsBalance.State <> 1 Then Exit Function
    
    For i = 1 To cllSquareCard.Count
        'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
        varTemp = cllSquareCard(i): dblDelMoney = Val(varTemp(2))
        'zlGetCard:(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As Card)
        If gobjSquare.objSquareCard.zlGetCard(Val(varTemp(0)), True, objCard) = False Then Exit Function
        
        mobjDelBalance.rsBalance.Filter = "类型=5 And 结算卡序号=" & Val(varTemp(0)) & " And 消费卡ID=" & Val(varTemp(1))
        If mobjDelBalance.rsBalance.RecordCount = 0 Then
            MsgBox "未找到" & objCard.结算方式 & "的原始结算记录，不能进行退款操作！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        With mobjDelBalance.rsBalance
            .MoveFirst
            Do While Not .EOF
                dblMoney = dblMoney + Val(Nvl(mobjDelBalance.rsBalance!冲预交))
                .MoveNext
            Loop
            .MoveFirst
        End With
        dblMoney = RoundEx(dblMoney, 6)
        If RoundEx(dblMoney, 6) = 0 Then
            MsgBox objCard.结算方式 & " 的已经全退退完，不能再退！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        If RoundEx(Val(varTemp(2)), 6) > RoundEx(dblMoney, 6) Then
            MsgBox objCard.结算方式 & " 的退款金额超过了原始结算金额！" & vbCrLf & "原结算金额:" & Format(dblMoney, "0.00") & vbCrLf & " 现退款金额:" & Format(Val(varTemp(2)), "0.00"), vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        
        '卡类别ID|卡号|消费卡ID|消费金额||.
        str结算方式 = str结算方式 & "||" & Val(varTemp(0))
        str结算方式 = str结算方式 & "|" & varTemp(3)
        str结算方式 = str结算方式 & "|" & Val(varTemp(1))
        str结算方式 = str结算方式 & "|" & -1 * dblDelMoney
    Next
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 3)
    
    'Zl_门诊退费结算_Modify
    strSQL = "Zl_门诊退费结算_Modify("
    '  操作类型_In   Number,
    '  --操作类型_In:
    '  --   0-原样退
    '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
    '  --   1-普通退费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款
    '  --     ③剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交
    '  --   2.三方卡退费结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退预交_In: 传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   4-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  -- 误差金额_In:存在误差费时,传入
    '  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费
    '  -- 原结帐ID_IN:原样退时,传入(如果原样退未传入时,则以最后一次结帐为准)
    strSQL = strSQL & "" & 4 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mobjDelBalance.病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & mobjDelBalance.冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  退预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL)"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成退费_In   Number := 0,
    '  原结帐id_In   病人预交记录.结帐id%Type := Null
    zlAddArray cllPro, strSQL
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    Set cllBillPro = New Collection
    
    '77156,冉俊明,2014-8-26,普通病人使用银行卡退费后，还可以点击返回按钮导致产生了退费的异常单据
    mobjDelBalance.SaveBilled = True
    ExecuteSquareDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Execute原样退() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行原样退功能(只处理三方接口)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-31 14:49:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objCard As Card, lng卡类别ID As Long, varTemp As Variant
    Dim dblMoney As Double, int类型 As Integer, varTemp1 As Variant, j As Integer
    Dim strCardTypeIDs As String, cllBalance As New Collection, dblTemp As Double
    
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            ''是否允许编辑:1-禁止编辑;0-不禁止编辑
            lng卡类别ID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
            int类型 = Val(.TextMatrix(i, .ColIndex("类型")))
            varTemp = .TextMatrix(i, .ColIndex("支付方式"))
            If (Val(.TextMatrix(i, .ColIndex("删除标志"))) = 0 Or Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0) Then
                Select Case int类型
                Case 1 '预交款
                    If Val(.TextMatrix(i, .ColIndex("是否验证"))) = 0 Then
                        Set objCard = GetPayCard(Trim(.TextMatrix(i, .ColIndex("支付方式"))), False)
                        dblMoney = RoundEx(Val(.Cell(flexcpData, i, .ColIndex("支付金额"))), 4)
                        If CheckPrepayMoneyIsValied(objCard, 1, dblMoney) = False Then Exit Function
                        .TextMatrix(i, .ColIndex("是否验证")) = 1
                        Call SetDeleteVisible
                    End If
                Case 3 '一卡通
                    '证明是三方接口且允许编辑,因此,需要在完成时,原样退款
                    Set objCard = GetPayCard(lng卡类别ID, False)
                    If objCard Is Nothing Then
                        MsgBox "注意:" & vbCrLf & varTemp & " 不是有效的支付方式，不能进行退款！", vbInformation, gstrSysName
                        '允许退出，否则就卡在这个窗口中了
                        If cmdExit.Visible = False And cmdDel.Visible = False Then cmdExit.Visible = True
                        Exit Function
                    End If
                    '检查合法性
                    dblMoney = RoundEx(-1 * Val(.Cell(flexcpData, i, .ColIndex("支付金额"))), 4)
                    If CheckThreeSwapValied(objCard, dblMoney) = False Then Exit Function
                    '退三方卡交易
                    If ExecuteThreeSwapDelInterface(objCard, dblMoney, mcllDelPro) = False Then Exit Function
                    .TextMatrix(i, .ColIndex("删除标志")) = 1
                    .TextMatrix(i, .ColIndex("结算状态")) = 1
                    .Cell(flexcpBackColor, i, 0, i, .COLS - 1) = Me.BackColor
                    Call SetDeleteVisible
                Case 4 '一卡通(老)
                Case 5 '消费卡
                
                Case Else
                End Select
            End If
        Next
    End With
    
    '消费卡进行退费
    If Not mcllSquareBalance Is Nothing Then
        For i = 1 To mcllSquareBalance.Count
            cllBalance.Add mcllSquareBalance(i)
        Next
        
        strCardTypeIDs = ""
        For i = 1 To cllBalance.Count
            varTemp = cllBalance(i)
            lng卡类别ID = Val(varTemp(0))
            If InStr(1, strCardTypeIDs & ",", "," & lng卡类别ID & ",") = 0 Then
                Set objCard = GetPayCard(lng卡类别ID, True, False)
                If objCard Is Nothing Then
                    MsgBox "注意:" & vbCrLf & " 不是有效的支付方式，不能进行退款！", vbInformation, gstrSysName
                    '允许退出，否则就卡在这个窗口中了
                    If cmdExit.Visible = False And cmdDel.Visible = False Then cmdExit.Visible = True
                    Exit Function
                End If
                dblMoney = 0
                dblTemp = 0
                For j = 1 To cllBalance.Count
                    varTemp1 = cllBalance(j)
                    'array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
                    If lng卡类别ID = Val(varTemp1(0)) Then
                        If UBound(varTemp1) >= 7 Then
                            dblMoney = dblMoney + Val(varTemp1(7))
                        Else
                            dblMoney = dblMoney + Val(varTemp1(2))
                        End If
                        dblTemp = dblTemp + Val(varTemp1(2))
                    End If
                Next
                dblMoney = RoundEx(dblMoney, 6): dblTemp = RoundEx(dblTemp, 6)
                If RoundEx(dblMoney, 6) <> 0 And RoundEx(dblTemp, 6) = 0 Then
                    '有可能不是全部退,以结算列表中金额为准
                    dblMoney = 0
                    For j = 1 To vsBlance.Rows - 1
                        '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                        If Val(vsBlance.TextMatrix(j, vsBlance.ColIndex("类型"))) = 5 _
                            And Val(vsBlance.TextMatrix(j, vsBlance.ColIndex("卡类别ID"))) = lng卡类别ID Then
                            dblMoney = dblMoney + -1 * Val(vsBlance.Cell(flexcpData, j, vsBlance.ColIndex("支付金额")))
                        End If
                    Next
                    dblMoney = RoundEx(dblMoney, 6)
                    If CheckSquareDelValied(objCard, 0, dblMoney) = False Then Exit Function
                
                    dblTemp = 0
                    For j = 1 To mcllSquareBalance.Count
                        varTemp1 = mcllSquareBalance(j)
                        'array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文)
                        If lng卡类别ID = Val(varTemp1(0)) Then
                            dblTemp = dblTemp + Val(varTemp1(2))
                        End If
                    Next
                    dblTemp = RoundEx(dblTemp, 6)
                    If RoundEx(dblTemp, 6) <> RoundEx(dblMoney, 6) Then
                        Set objCard = GetPayCard(lng卡类别ID, True)
                        Call AddSquareBalance(objCard, True)
                        Call MsgBox("注意：" & vbCrLf & objCard.结算方式 & "支付金额与当前刷卡金额不一致，请重新输入退款金额！" & vbCrLf & _
                                    "  原支付金额：" & Format(dblMoney, "0.00") & vbCrLf & _
                                    "  当前刷卡金额：" & Format(dblTemp, "0.00"), vbInformation + vbOKOnly, gstrSysName)
                        mobjDelBalance.原样退 = False
                        Call StartAndStop预存款
                        Call SetDeleteVisible
                        Call SetControlProperty(True)
                        Exit Function
                    End If
                End If
            End If
            strCardTypeIDs = strCardTypeIDs & "," & lng卡类别ID
        Next
      
    End If
    'If ExecuteSquareDelInterface(mcllSquareBalance, mcllDelPro) = False Then Exit Function
    Call SetDeleteVisible
    Execute原样退 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetPayCard(ByVal strCardType As String, ByVal bln消费卡 As Boolean, Optional bln仅启用 As Boolean = True) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡类别ID
    '入参:lngCardTypeID-卡类别ID
    '返回:返回Card对象
    '编制:刘兴洪
    '日期:2014-07-31 15:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim lngCardTypeID As Long
    On Error GoTo errHandle
    If Not IsNumeric(strCardType) Then
        For Each objCard In mobjPayCards
            If objCard.接口序号 <= 0 And objCard.结算方式 = strCardType Then
                Set GetPayCard = objCard
                Exit Function
            End If
        Next
        Exit Function
    End If
    lngCardTypeID = Val(strCardType)
    For Each objCard In mobjPayCards
        If objCard.接口序号 = lngCardTypeID And objCard.消费卡 = bln消费卡 Then
            Set GetPayCard = objCard
            Exit Function
        End If
    Next
    If bln仅启用 = False Then
        If Not gobjSquare.objSquareCard Is Nothing Then
            'zlGetCard:(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As Card)
            If gobjSquare.objSquareCard.zlGetCard(lngCardTypeID, bln消费卡, objCard) = False Then Exit Function
            Set GetPayCard = objCard
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckSquareDelValied(ByVal objCard As Card, Optional ByVal lng消费卡ID As Long, Optional dblDelMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡退费检查
    '入参:objCard-三方卡
    '     dblDelMoney-退款金额
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接品的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double
    Dim cllSquareBalance As Collection, cllBalance As Collection
    Dim dblToTal As Double, dblBrushMoney As Double
    Dim varData As Variant, varTemp As Variant, i As Integer, j As Integer
    Dim strBalances As String, dblRestMoney As Double
    Dim lng消费卡 As Long, str卡号 As String
    
    On Error GoTo errHandle
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then CheckSquareDelValied = True: Exit Function
    '退费
    If Not (mCurCarge.dbl当前未退 <= 0 Or dblDelMoney <> 0) Then CheckSquareDelValied = True: Exit Function
    If dblDelMoney = 0 Then
        If Val(txt缴款.Text) = 0 Then
            MsgBox "未输入退费金额，请检查！", vbInformation + vbOKOnly, gstrSysName
             Exit Function
        Else
            dblDelMoney = Val(txt缴款.Text)
        End If
    End If
     
    '退款检查
    If mobjDelBalance.rsBalance Is Nothing Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjDelBalance.rsBalance.State <> 1 Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    If lng消费卡ID <> 0 Then
        mobjDelBalance.rsBalance.Filter = "类型=5 And 结算卡序号=" & objCard.接口序号 & " And 消费卡ID=" & lng消费卡ID
    Else
        mobjDelBalance.rsBalance.Filter = "类型=5 And 结算卡序号=" & objCard.接口序号
    End If
    
    If mobjDelBalance.rsBalance.EOF Then
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllSquareBalance = New Collection
    Set cllBalance = New Collection
    dblTemp = dblDelMoney: dblToTal = 0
    With mobjDelBalance.rsBalance
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblToTal = dblToTal + Val(Nvl(!冲预交))
            
            lng消费卡 = Val(Nvl(!消费卡ID)): str卡号 = Nvl(!卡号)
            If InStr(strBalances & ",", "," & objCard.接口序号 & "|" & lng消费卡 & "|" & str卡号 & ",") = 0 Then
                '按接口序号、消费卡ID、卡号求剩余未退金额
                dblRestMoney = 0
                j = .AbsolutePosition
                .MoveFirst
                Do While Not .EOF
                    If Val(Nvl(!结算卡序号)) = objCard.接口序号 _
                        And Val(Nvl(!消费卡ID)) = lng消费卡 And Nvl(!卡号) = str卡号 Then
                        dblRestMoney = dblRestMoney + Val(Nvl(!冲预交))
                    End If
                    .MoveNext
                Loop
                .Move j - 1, adBookmarkFirst
                
                '剩余未退金额
                dblRestMoney = RoundEx(dblRestMoney, 6)
                '已刷卡金额
                dblBrushMoney = GetSquareBrushMoney(objCard.接口序号, lng消费卡, str卡号)
                
                If dblRestMoney <> 0 Then
                    'array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
                    cllSquareBalance.Add Array(objCard.接口序号, lng消费卡, dblBrushMoney, str卡号, "", "", 0, dblRestMoney)
                     
                    If dblTemp > dblRestMoney And dblTemp <> 0 Then
                        cllBalance.Add Array(objCard.接口序号, lng消费卡, dblRestMoney, str卡号, "", "", 0)
                        dblTemp = dblTemp - dblRestMoney
                    ElseIf dblTemp <> 0 Then
                        cllBalance.Add Array(objCard.接口序号, lng消费卡, dblTemp, str卡号, "", "", 0)
                        dblTemp = 0
                    End If
                End If
                dblTemp = RoundEx(dblTemp, 6)
                strBalances = strBalances & "," & objCard.接口序号 & "|" & lng消费卡 & "|" & str卡号
            End If
            .MoveNext
        Loop
    End With
    dblToTal = RoundEx(dblToTal, 6)
    
    If RoundEx(dblToTal, 6) < RoundEx(dblDelMoney, 6) Then
        MsgBox "注意:" & vbCrLf & "   输入的退款金额大于了" & objCard.结算方式 & "的可退金额,请检查!" & vbCrLf & _
               "   可退金额:" & Format(dblToTal, "###0.00;-###0.00;;") & vbCrLf & _
               "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If RoundEx(dblToTal, 6) <> RoundEx(dblDelMoney, 6) Then
        If objCard.是否全退 And Not objCard.是否退现 Then
            MsgBox "注意:" & vbCrLf & "   " & objCard.结算方式 & "不支持退现,必须全退,请检查!" & vbCrLf & _
                   "   可退金额:" & Format(dblToTal, "###0.00;-###0.00;;") & vbCrLf & _
                   "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If gbln消费卡退费验卡 Then
       '弹出刷卡界面
        'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByVal dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-三方卡调用XML入参,目前格式如下:
        '       <IN>
        '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        '       </IN>
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.接口序号, _
            objCard.消费卡, mobjDelBalance.姓名, mobjDelBalance.性别, _
            mobjDelBalance.年龄, dblDelMoney, "", "", _
            True, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
        Set cllBalance = cllSquareBalance
    End If
    
    If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
    '从消费卡退费结算信息集合中移除当前卡类别的记录
    j = 1
    Do While True
        If j > mcllSquareBalance.Count Then Exit Do
        varTemp = mcllSquareBalance(j)
        'array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文)
        If objCard.接口序号 = Val(varTemp(0)) Then 'And varData(3) = varTemp(3)
            mcllSquareBalance.Remove j
        Else
           j = j + 1
        End If
    Loop
    
    '将刷卡验证后的当前卡类别的记录添加到消费卡退费结算信息集合中
    dblTemp = 0
    For i = 1 To cllBalance.Count
        varData = cllBalance(i)
        dblTemp = Val(varData(2)) + dblTemp
        mcllSquareBalance.Add varData
    Next
    CheckSquareDelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSquareBrushMoney(ByVal lngCardTypeID As Long, ByVal lng消费卡ID As Long, ByVal strCardNo As String) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡已刷卡金额
    '入参:lngCardTypeId-消费卡接口编号
    '     lng消费卡ID-消费卡ID
    '     strCardNo-卡号
    '出参:
    '返回:返回刷卡金额
    '编制:刘兴洪
    '日期:2014-08-12 11:51:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, varTemp As Variant
    Dim dblMoney As Double
    If mobjDelBalance.原样退 Then Exit Function
    If mcllSquareBalance Is Nothing Then Exit Function
    dblMoney = 0
    'array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文)
    For j = 1 To mcllSquareBalance.Count
        varTemp = mcllSquareBalance(j)
        If Val(varTemp(0)) = lngCardTypeID And _
           ((lng消费卡ID = varTemp(1) And lng消费卡ID <> 0) _
             Or varTemp(3) = strCardNo) Then
             
            dblMoney = dblMoney + Val(varTemp(2))
        End If
    Next
    GetSquareBrushMoney = RoundEx(dblMoney, 6)
End Function

Private Sub ClearSquareBalance(ByVal lngCardTypeID As Long, Optional ByVal lng消费卡ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除消费卡结算
    '编制:刘兴洪
    '日期:2014-08-12 10:39:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBlance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("类型"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("删除标志"))) = 0 _
                And Val(.TextMatrix(j, .ColIndex("卡类别ID"))) = lngCardTypeID _
                And (lng消费卡ID = 0 Or (lng消费卡ID <> 0 And Val(.TextMatrix(j, .ColIndex("消费卡ID"))) = lng消费卡ID)) Then
                dblMoney = Val(.Cell(flexcpData, j, .ColIndex("支付金额")))
                
                mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 - dblMoney, 6)
                mCurCarge.dbl当前未退 = RoundEx(mCurCarge.dbl当前未退 + dblMoney, 6)
                If .Rows <= 2 Then
                    .Rows = 2
                   .Cell(flexcpData, 1, 0, 1, .COLS - 1) = ""
                   .Cell(flexcpText, 1, 0, 1, .COLS - 1) = ""
                   .RowData(1) = ""
                   j = 2
                Else
                    .RemoveItem j
                End If
            Else
                j = j + 1
            End If
        Loop
    End With
End Sub

Private Function CheckIsExistCashValied(objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在退现的三方结算及消费卡
    '入参:
    '出参:
    '返回:存在退现且数据合法的,返回True,否则返回False
    '编制:刘兴洪
    '日期:2014-08-12 18:18:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsBalance As ADODB.Recordset
    Dim lngCardTypeID As Long, strCardTypeIDs As String
    Dim j As Long, blnFind As Boolean, bln消费卡 As Boolean
    Dim int类型 As Integer, lngID As Long, dblTemp As Double
    Dim bln强制退现 As Boolean
    
    On Error GoTo errHandle
    If mblnSingleBalance Then CheckIsExistCashValied = True: Exit Function '不用检查
    If mCurCarge.dbl当前未退 >= 0 Then CheckIsExistCashValied = True: Exit Function '86915
    Set rsBalance = mobjDelBalance.rsBalance
    If rsBalance Is Nothing Then CheckIsExistCashValied = True: Exit Function
    If rsBalance.State <> 1 Then CheckIsExistCashValied = True: Exit Function
    
    rsBalance.Filter = "(类型=3 And 是否退现=0) Or (类型=5 And 是否退现=0)"
    If rsBalance.RecordCount = 0 Then CheckIsExistCashValied = True: Exit Function
    
    With rsBalance
        If .RecordCount <> 0 Then .MoveFirst
        strCardTypeIDs = ""
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(!卡类别ID)): bln消费卡 = False
            If lngCardTypeID = 0 Then lngCardTypeID = Val(Nvl(!结算卡序号)): bln消费卡 = True
            
            bln强制退现 = False
            If lngCardTypeID > 0 Then
                For j = 1 To mcllForceDelToCash.Count
                    If mcllForceDelToCash(j)(1) = Nvl(!卡类别名称) Then bln强制退现 = True: Exit For
                Next
            End If
            
            
            If bln强制退现 = False And lngCardTypeID > 0 And InStr(1, strCardTypeIDs & "||", "||" & lngCardTypeID & "," & IIf(bln消费卡, 1, 0) & "||") = 0 Then
                '查找是否在结算信息中是否存在
                blnFind = False
                For j = 1 To vsBlance.Rows - 1
                    lngID = Val(vsBlance.TextMatrix(j, vsBlance.ColIndex("卡类别ID")))
                     '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                    int类型 = Val(vsBlance.TextMatrix(j, vsBlance.ColIndex("类型")))
                    If bln消费卡 Then
                        If lngID = lngCardTypeID And int类型 = 5 Then
                            blnFind = True: Exit For
                        End If
                    Else
                        If lngID = lngCardTypeID And int类型 = 3 Then
                            blnFind = True: Exit For
                        End If
                    End If
                Next
                If Not objCard Is Nothing Then
                    If objCard.接口序号 = lngCardTypeID And objCard.消费卡 = bln消费卡 Then blnFind = True
                End If
                
                If blnFind = False Then
                    j = .AbsolutePosition
                    dblTemp = 0
                    '检查是否已退完，若退完直接跳过(可能第一次退费已退过)
                    Do While Not .EOF
                        If bln消费卡 Then
                            If Val(Nvl(!类型)) = 5 And Val(Nvl(!结算卡序号)) = lngCardTypeID Then
                                dblTemp = dblTemp + Val(Nvl(!冲预交))
                            End If
                        Else
                            If Val(Nvl(!类型)) = 3 And Val(Nvl(!卡类别ID)) = lngCardTypeID Then
                                dblTemp = dblTemp + Val(Nvl(!冲预交))
                            End If
                        End If
                        .MoveNext
                    Loop
                    dblTemp = RoundEx(dblTemp, 6)
                    .Move j - 1, adBookmarkFirst
                    If dblTemp <> 0 Then
                        MsgBox Nvl(rsBalance!结算方式) & " 不能退现，必须全退！", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                End If
                strCardTypeIDs = strCardTypeIDs & "||" & lngCardTypeID & "," & IIf(bln消费卡, 1, 0)
            End If
            .MoveNext
        Loop
    End With
    CheckIsExistCashValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub AddSquareBalance(ByVal objCard As Card, ByVal blnDel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加消费卡支付方式到结算方式列表
    '编制:刘兴洪
    '日期:2014-08-12 18:18:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As New Collection
    Dim j As Integer, dblMoney As Double, strCardNo As String
    
    With vsBlance
      '先清除原始的消费卡部分,再重新退费
        Call ClearSquareBalance(objCard.接口序号)
        If blnDel Then
            If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
            Set cllBalance = mcllSquareBalance
        Else
            If mcllCurSquareBalance Is Nothing Then Set mcllCurSquareBalance = New Collection
            Set cllBalance = mcllCurSquareBalance
        End If
        
        For j = 1 To cllBalance.Count
            If objCard.接口序号 = Val(cllBalance(j)(0)) Then
                If Not blnDel Then
                    If mcllSquareChargeBalance Is Nothing Then Set mcllSquareChargeBalance = New Collection
                    mcllSquareChargeBalance.Add cllBalance(j)
                End If
                '当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                
                '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                dblMoney = Val(IIf(blnDel, -1, 1) * cllBalance(j)(2))
                .RowData(1) = 5
                .TextMatrix(1, .ColIndex("类型")) = 5
                .TextMatrix(1, .ColIndex("结算性质")) = objCard.结算性质
                .TextMatrix(1, .ColIndex("删除标志")) = IIf(blnDel, 0, 1) '是否允许编辑:1-禁止编辑;0-不禁止编辑
                .TextMatrix(1, .ColIndex("结算状态")) = IIf(blnDel, 0, 1)  '是否已结算:1-已结算;0-未结算
                .TextMatrix(1, .ColIndex("卡类别ID")) = objCard.接口序号
                .TextMatrix(1, .ColIndex("消费卡ID")) = Val(cllBalance(j)(1))
                .TextMatrix(1, .ColIndex("支付方式")) = objCard.结算方式
                 ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                .Cell(flexcpData, 1, .ColIndex("支付方式")) = Val(cllBalance(j)(0)) & "|" & 1 & "|" & IIf(objCard.自制卡, 1, 0) & _
                                                            "|" & IIf(objCard.是否全退, 1, 0) & "|" & IIf(objCard.是否退现, 1, 0) & "|" & objCard.名称
                strCardNo = Trim(cllBalance(j)(3))
                .TextMatrix(1, .ColIndex("卡号")) = IIf(objCard.卡号密文规则 <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("卡号")) = strCardNo
                .TextMatrix(1, .ColIndex("支付金额")) = Format(-1 * dblMoney, "0.00")
                .Cell(flexcpData, 1, .ColIndex("支付金额")) = Format(dblMoney, "0.00")
                .TextMatrix(1, .ColIndex("结算号码")) = ""
                .TextMatrix(1, .ColIndex("备注")) = ""
                .TextMatrix(1, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
                .TextMatrix(1, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
                .TextMatrix(1, .ColIndex("是否转帐及代扣")) = IIf(objCard.是否转帐及代扣, 1, 0)
                .TextMatrix(1, .ColIndex("卡类别名称")) = objCard.名称
                
                mCurCarge.dbl已退合计 = RoundEx(mCurCarge.dbl已退合计 + dblMoney, 6)
                mCurCarge.dbl当前未退 = RoundEx(mCurCarge.dbl当前未退 - dblMoney, 6)
            End If
        Next
    End With
End Sub

Private Function CheckSquareBalanceValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡结算交易检查
    '入参:objCard-三方卡
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接品的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl帐户余额 As Double
    Dim strBrushCard As TY_BrushCard, cllSquareBalance As Collection
    Dim strExpand As String, bln退现 As Boolean
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then CheckSquareBalanceValied = True: Exit Function
    
    On Error GoTo errHandle
    If mCurCarge.dbl当前未退 <= 0 Then CheckSquareBalanceValied = True: Exit Function
    
    If Val(txt缴款) = 0 Then
        MsgBox strTittle & "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    If Abs(Val(txt缴款.Text)) > Format(Abs(mCurCarge.dbl当前未退), "0.00") And Val(txt缴款.Text) <> 0 Then
        MsgBox strTittle & "金额不能大于本次未付金额:" & Format(mCurCarge.dbl当前未退, "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '先检查对应的接口
    If zlGetClassMoney(0, rsMoney) = False Then Exit Function
    
     '构建消费卡的刷卡信息
    Set cllSquareBalance = mcllSquareChargeBalance
    Set mcllCurSquareBalance = Nothing
    
    '   zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln消费卡 As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByRef dbl金额 As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln退费 As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln退现 As Boolean = False, _
            Optional ByVal bln余额不足禁止 As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal bln转预交 As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXmlIn As String = "") As Boolean
            '       strXmlIn-三方卡调用XML入参,目前格式如下:
            '       <IN>
            '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
            '       </IN>
    '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
     
    dblMoney = Val(txt缴款.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, _
            objCard.接口序号, objCard.消费卡, _
            mobjDelBalance.姓名, mobjDelBalance.性别, mobjDelBalance.年龄, dblMoney, _
            mCurBrushCard.str卡号, mCurBrushCard.str密码, _
            False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>0</CZLX></IN>") = False Then Exit Function
        Set mcllCurSquareBalance = cllSquareBalance
        '保存前,一些数据检查
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal strCardTypeID As Long, ByVal strCardNo As String, _
        ByVal dblMoney As Double, ByVal strNOs As String, _
        Optional ByVal strXMLExpend As String
        'mobjDelBalance.strNOs:单独保存时,没有相关时,可能为空.
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, objCard.接口序号, _
            objCard.消费卡, mCurBrushCard.str卡号, dblMoney, mobjDelBalance.CurDelNos, strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
        '入参:frmMain-调用的主窗体
        '        lngModule-模块号
        '        strCardNo-卡号
        '        strExpand-预留，为空,以后扩展
        '出参:dblMoney-返回帐户余额
        If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, objCard.接口序号, _
              mCurBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡) = False Then Exit Function
    
        stbThis.Panels(4).Text = Format(dbl帐户余额, "0.00")
        stbThis.Panels(4).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
        mCurBrushCard.dbl帐户余额 = RoundEx(dbl帐户余额, 2)
        If RoundEx(dbl帐户余额, 6) <> 0 And dbl帐户余额 < dblMoney Then
            MsgBox objCard.结算方式 & "的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        '已经更改了支付金额
        If RoundEx(dblMoney, 6) <> Val(txt缴款.Text) Then
            txt缴款.Text = FormatEx(dblMoney, 6, , , 2)
        End If
        CheckSquareBalanceValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteSquarePayInterface(objCard As Card, ByVal dblMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡支付
    '入参:lng结算序号-按结算序号进行处理
    '     dblMoney-本次支付金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long, strSQL As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim str消费卡结算  As String, j As Long
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '非消费卡支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then ExecuteSquarePayInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    str消费卡结算 = ""  '卡类别ID|卡号|消费卡ID|消费金额||....
    If mcllCurSquareBalance Is Nothing Then Exit Function
    If mcllCurSquareBalance.Count = 0 Then Exit Function
    For j = 1 To mcllCurSquareBalance.Count
        ' array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
        str消费卡结算 = str消费卡结算 & "||" & Val(mcllCurSquareBalance(j)(0))
        str消费卡结算 = str消费卡结算 & "|" & mcllCurSquareBalance(j)(3)
        str消费卡结算 = str消费卡结算 & "|" & Val(mcllCurSquareBalance(j)(1))
        str消费卡结算 = str消费卡结算 & "|" & Val(mcllCurSquareBalance(j)(2))
    Next
    If str消费卡结算 <> "" Then str消费卡结算 = Mid(str消费卡结算, 3)
    
    '调用之前,先处理数据
    'Zl_门诊退费结算_Modify
    strSQL = "Zl_门诊退费结算_Modify("
    '  操作类型_In   Number,
    '  --操作类型_In:
    '  --   0-原样退
    '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
    '  --   1-普通退费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款
    '  --     ③剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交
    '  --   2.三方卡退费结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退预交_In: 传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   4-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    strSQL = strSQL & "" & 4 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mobjDelBalance.病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & mobjDelBalance.冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & str消费卡结算 & "',"
    '  冲预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & objCard.接口序号 & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL)"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成退费_In   Number := 0,
    '  原结帐id_In   病人预交记录.结帐id%Type := Null,
    '  剩余转预交_In Number:=0
    zlAddArray cllPro, strSQL
    
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
    
    str结帐IDs = mobjDelBalance.冲销ID
    str结帐IDs = str结帐IDs & IIf(mobjDelBalance.结帐ID <> 0, "," & mobjDelBalance.结帐ID, "")
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, _
         str结帐IDs, _
        mobjDelBalance.CurDelNos, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    mCurBrushCard.str交易流水号 = strSwapGlideNO
    mCurBrushCard.str交易说明 = strSwapMemo
    If objCard.消费卡 = False Then
        Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    End If
    '扩展交易信息
    Call zlAddThreeSwapSQLToCollection(False, str结帐IDs, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    
    '77156,冉俊明,2014-8-26,普通病人使用银行卡退费后，还可以点击返回按钮导致产生了退费的异常单据
    mobjDelBalance.SaveBilled = True
    ExecuteSquarePayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load退费方式(Optional ByVal bln强制退现 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载退费方式
    '编制:刘兴洪
    '日期:2014-09-01 16:05:44
    '说明:
    '   缺省退款方式规则:
    '       以收款结算时非医保结算为缺省的退款方式 , 如果存在多个, 则按以下规则缺省:
    '       1)三方帐户:存在三方帐户的,缺省该三方帐户,存在多个三方帐户支付,则缺省为第一个三方帐户.
    '       2)收款结算方式中只存在一种非医保结算方式的, 则缺省为该结算方式
    '       3)收款结算方式中存在缺省的结算方式, 则以缺省的为准
    '       4)以现金为准
    '   bln强制退现=true:缺省为现金
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, blnChargeUsed As Boolean
    Dim i As Long, str结算方式 As String, strTemp As String
    Dim blnSetedIndex As Boolean
      
    mblnNotClick = True
    mlngPre支付方式 = 0
    
    Call StartAndStop预存款

    With cbo支付方式
        .Clear
        For i = 1 To mobjPayCards.Count
            Set objCard = mobjPayCards(i)
            If objCard.支付启用 = True And InStr(str结算方式 & "|", "|" & objCard.结算方式 & "|") = 0 Then
                '三方账户的支付方式显示为医疗卡名称，其它显示结算方式
                If objCard.接口序号 > 0 And Not objCard.消费卡 Then
                    .AddItem objCard.名称: str结算方式 = str结算方式 & "|" & objCard.名称
                Else
                    .AddItem objCard.结算方式: str结算方式 = str结算方式 & "|" & objCard.结算方式
                End If
                .ItemData(.NewIndex) = i
            End If
        Next
        '设置缺省值
        For i = .ListCount - 1 To 0 Step -1
            Set objCard = mobjPayCards(.ItemData(i))
            mobjDelBalance.rsBalance.Filter = "结算性质=2"
            strTemp = ""
            If mobjDelBalance.rsBalance.RecordCount > 0 Then
                mobjDelBalance.rsBalance.MoveFirst
                strTemp = Nvl(mobjDelBalance.rsBalance!结算方式)
            End If
            If mCurCarge.dbl当前未退 < 0 Then '退费
                If mblnSingleBalance Then
                    If mobjDelBalance.缺省结算方式 = objCard.结算方式 Then
                        If objCard.接口序号 > 0 Then
                            '三方帐户的,如果不缺省退现,则缺省该三方帐户
                            If Not (objCard.是否退现 And objCard.是否缺省退现) Then .ListIndex = i
                        ElseIf InStr(gTy_Module_Para.str缺省退现, objCard.结算方式) = 0 Then
                            '没有设置缺省退现,则缺省为该结算方式
                            .ListIndex = i
                        End If
                    End If
                Else
                    '三方帐户:存在三方帐户的,如果不缺省退现,则缺省该三方帐户,存在多个三方帐户支付,则缺省为第一个三方帐户.
                    If objCard.接口序号 > 0 Then
                        If Not (objCard.是否退现 And objCard.是否缺省退现) And blnSetedIndex = False Then
                            '93114，缴费时未使用的不缺省
                            Call CheckThreeSwapCanTransfer(objCard, mobjDelBalance.原结帐ID, blnChargeUsed)
                            If blnChargeUsed Then .ListIndex = i: blnSetedIndex = True
                        End If
                    End If
                    '有使用预交款，则缺省预交款
                    If objCard.结算性质 = -99 And .ListIndex < 0 Then .ListIndex = i
                    '收款结算方式中存在一种非医保结算方式的,如果没有设置缺省退现,则缺省为该结算方式
                    If InStr(gTy_Module_Para.str缺省退现, objCard.结算方式) = 0 Then
                        If strTemp = objCard.结算方式 And .ListIndex < 0 Then .ListIndex = i
                    End If
                End If
                '收款结算方式中存在缺省的结算方式,则以缺省的为准
                If objCard.缺省标志 And .ListIndex < 0 Then .ListIndex = i
                '以现金为准
            Else
                If objCard.缺省标志 And .ListIndex < 0 Then .ListIndex = i
                If objCard.结算性质 = 1 And .ListIndex < 0 Then .ListIndex = i
                If mobjDelBalance.缺省结算方式 = objCard.结算方式 Then .ListIndex = i
            End If
        Next
        If gstr结算方式 <> "" And .ListIndex < 0 Then
            For i = 0 To .ListCount - 1
                If .List(i) = gstr结算方式 Then
                    .ListIndex = i: Exit For
                End If
            Next
        End If
        If bln强制退现 Then
            For i = .ListCount - 1 To 0 Step -1
                Set objCard = mobjPayCards(.ItemData(i))
                If objCard.结算性质 = 1 Then .ListIndex = i: Exit For
            Next
        End If
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Call cbo支付方式_Click
End Sub

Private Sub Set退费方式(ByVal bytType As Byte, Optional ByVal objCard As Card, Optional ByVal bln启用 As Boolean, _
    Optional ByVal bln强制退现 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置指定的退费方式是否启用
    '参数:
    '   bytType:1=根据传入卡对象设置结算方式是否可用
    '           2=根据退费结算方式列表和结算数据设置退费结算方式是否可用
    '           3=根据退费结算方式列表设置收费结算方式是否可用
    '编制:刘兴洪
    '日期:2014-09-01 16:14:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Card
    Dim i As Long, j As Long, blnFind As Boolean, dblMoney As Double
    Dim rsTemp As ADODB.Recordset, blnDefault As Boolean
    
    On Error GoTo Errhand
    Select Case bytType
        Case 1
            If objCard Is Nothing Then Exit Sub
            For Each objTemp In mobjPayCards
                If objTemp.接口序号 = objCard.接口序号 And objTemp.消费卡 = objCard.消费卡 Then
                    objTemp.支付启用 = bln启用
                End If
            Next
        Case 2
            Set rsTemp = mobjDelBalance.rsBalance
            For i = 1 To mobjPayCards.Count
                Set objTemp = mobjPayCards(i)
                dblMoney = 0: blnFind = False: blnDefault = False
                '判断结算方式剩余未退金额
                With rsTemp
                    rsTemp.Filter = 0
                    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
                    Do While Not rsTemp.EOF
                        '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                        If Nvl(rsTemp!类型) <> 1 Then '预交款不能通过"结算方式"进行判断
                            If Nvl(rsTemp!结算方式) = objTemp.结算方式 Then dblMoney = dblMoney + Nvl(rsTemp!冲预交)
                        End If
                        rsTemp.MoveNext
                    Loop
                End With
                dblMoney = RoundEx(dblMoney, 6)
                '判断是否在退费结算列表中
                For j = 1 To vsBlance.Rows - 1
                    If vsBlance.TextMatrix(j, vsBlance.ColIndex("支付方式")) = objTemp.结算方式 Then
                        blnFind = True: Exit For
                    End If
                Next
                '结算性质:1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
                If objTemp.结算性质 = 1 Or objTemp.结算性质 = 2 Then blnDefault = True
                If bln强制退现 Then
                    '强制退现时不允许转帐及代扣
                    objTemp.支付启用 = (RoundEx(dblMoney, 6) <> 0 Or blnDefault) And Not blnFind
                Else
                    objTemp.支付启用 = (RoundEx(dblMoney, 6) <> 0 Or blnDefault Or objTemp.是否转帐及代扣) And Not blnFind
                End If
            Next
        Case 3
            For i = 1 To mobjPayCards.Count
                Set objTemp = mobjPayCards(i)
                '判断是否在退费结算列表中，允许刷多次消费卡
                blnFind = False
                For j = 1 To vsBlance.Rows - 1
                    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                    With vsBlance
                        If .TextMatrix(j, .ColIndex("支付方式")) = objTemp.结算方式 And _
                            (Val(.TextMatrix(j, .ColIndex("类型"))) <> 5 Or _
                            (Val(.TextMatrix(j, .ColIndex("类型"))) = 5 And .Cell(flexcpData, j, .ColIndex("支付金额")) < 0)) Then
                            blnFind = True: Exit For
                        End If
                    End With
                Next
                objTemp.支付启用 = Not blnFind
            Next
        Case Else
    
    End Select
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CheckInterfaceNumIsValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查接口数量是否超过2个以上
    '返回:未超过2个数量,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-27 15:23:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, varData As Variant
    Dim strNames As String, i As Long
    
    On Error GoTo errHandle
    
    lngCount = IIf(mobjDelBalance.intInsure <> 0, 1, 0)   '医保算一个数量
'    If objCard.接口序号 <= 0 Or (objCard.消费卡 And objCard.自制卡) Then CheckInterfaceNumIsValied = True: Exit Function
    With vsBlance
        strNames = vbCrLf & IIf(mobjDelBalance.intInsure <> 0, "医保结算", "")
        For i = 1 To .Rows - 1
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If Val(.RowData(i)) = 3 Or Val(.RowData(i)) = 4 Or Val(.RowData(i)) = 5 Then
                ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                varData = Split(.Cell(flexcpData, i, .ColIndex("支付方式")) & "|||||", "|")
                If Val(varData(0)) <> 0 Then
                    If Val(varData(1)) <> 1 Then
                        lngCount = lngCount + 1
                        If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 1 Then
                            strNames = strNames & vbCrLf & varData(5)
                        End If
                    ElseIf Val(varData(2)) = 0 Then
                        '消费卡也是接口的,才算作第三方接口
                        lngCount = lngCount + 1
                        If Val(.TextMatrix(i, .ColIndex("结算状态"))) = 1 Then
                            strNames = strNames & vbCrLf & varData(5)
                        End If
                    End If
                End If
            End If
        Next
    End With
    If lngCount = 2 Then
        If objCard.接口序号 <= 0 Or (objCard.消费卡 And objCard.自制卡) Then CheckInterfaceNumIsValied = True: Exit Function
        MsgBox "  系统暂只支持两种以内的接口，不能再刷卡消费，请检查！" & vbCrLf & "以下为当前已经刷的接口：" & vbCrLf & strNames, vbOKOnly + vbInformation, gstrSysName
        Exit Function
    ElseIf lngCount > 2 Then
        MsgBox "  系统暂只支持两种以内的接口，不能再刷卡消费，请检查！" & vbCrLf & "以下为当前已经刷的接口：" & vbCrLf & strNames, vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    CheckInterfaceNumIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function CheckThreeSwapCanTransfer(ByVal objCard As Card, ByVal lng结帐ID As Long, _
    Optional ByRef blnChargeUsed As Boolean) As Boolean
    '检查三方卡是否可使用转帐方式退款
    '问题号:93114
    '说明：
    '   能使用转帐功能的条件：1.支持转帐及代扣；2.在缴费时未使用或者在缴费时使用了且能退现
    '   在缴费时使用了不能退现的三方卡只能原样退回
    Dim strSQL As String
    
    On Error GoTo errHandle
    blnChargeUsed = False
    If objCard Is Nothing Then Exit Function
    If objCard.接口序号 <= 0 Then Exit Function
    
    If mrsUsedCards Is Nothing Then
        '缓存，防止反复查询数据库
        strSQL = _
            "Select Nvl(a.卡类别id,a.结算卡序号) As 卡类别id" & vbNewLine & _
            " From 病人预交记录 A," & vbNewLine & _
            "      (Select m.结帐id" & vbNewLine & _
            "        From 门诊费用记录 M, 门诊费用记录 N" & vbNewLine & _
            "        Where m.记录性质 = n.记录性质 And m.No = n.No And n.结帐id = [1] And m.记录性质 = 1) B" & vbNewLine & _
            " Where a.结帐id = b.结帐id And a.记录状态 In (1, 3) And (Nvl(a.卡类别id, 0) <> 0 Or Nvl(a.结算卡序号, 0) <> 0)"
        Set mrsUsedCards = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    End If
    mrsUsedCards.Filter = "卡类别id=" & objCard.接口序号
    If mrsUsedCards.EOF Then
        CheckThreeSwapCanTransfer = True
    Else
        blnChargeUsed = True '缴费时使用了
        CheckThreeSwapCanTransfer = objCard.是否退现
    End If
    
    If objCard.是否转帐及代扣 = False Then CheckThreeSwapCanTransfer = False
    '强制退现时，不允许使用转帐及代扣
    If mcllForceDelToCash.Count > 0 Then CheckThreeSwapCanTransfer = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



