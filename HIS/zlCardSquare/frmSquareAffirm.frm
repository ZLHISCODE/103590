VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSquareAffirm 
   Caption         =   "病人消费结算"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSquareAffirm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9930
   StartUpPosition =   1  '所有者中心
   Begin MSCommLib.MSComm mscCom 
      Left            =   7290
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picSum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   135
      ScaleHeight     =   1845
      ScaleWidth      =   3060
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1665
      Width           =   3090
      Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption2 
         Height          =   420
         Left            =   15
         TabIndex        =   20
         Top             =   30
         Width           =   3045
         _Version        =   589884
         _ExtentX        =   5371
         _ExtentY        =   741
         _StockProps     =   6
         Caption         =   "本次消费合计"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label lbl自付合计 
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
         Height          =   615
         Left            =   2025
         TabIndex        =   19
         Top             =   795
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdPara 
      Caption         =   "打印设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8160
      TabIndex        =   17
      Top             =   2610
      Width           =   1680
   End
   Begin VB.PictureBox picFee 
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   75
      ScaleHeight     =   4260
      ScaleWidth      =   11445
      TabIndex        =   15
      Top             =   3630
      Width           =   11445
      Begin VB.Frame fraSplitBottom 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   0
         TabIndex        =   16
         Top             =   -75
         Width           =   11805
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFee 
         Height          =   2505
         Left            =   -15
         TabIndex        =   12
         Top             =   405
         Width           =   9855
         _cx             =   17383
         _cy             =   4419
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         ForeColorSel    =   -2147483634
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSquareAffirm.frx":0442
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "消费明细"
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
         Left            =   -15
         TabIndex        =   11
         Top             =   150
         Width           =   840
      End
   End
   Begin VB.PictureBox picPayMode 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   3300
      ScaleHeight     =   1875
      ScaleWidth      =   4575
      TabIndex        =   14
      Top             =   1650
      Width           =   4575
      Begin VB.TextBox txt冲预交 
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
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   1
         Top             =   210
         Width           =   2760
      End
      Begin VB.TextBox txt金额 
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
         Left            =   1485
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1335
         Width           =   2790
      End
      Begin VB.ComboBox cbo支付方式 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   765
         Width           =   2775
      End
      Begin VB.Label lbl预存款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 预存款"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   255
         Width           =   1110
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   810
         TabIndex        =   4
         Top             =   1410
         Width           =   630
      End
      Begin VB.Label lbl支付方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "支付方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   195
         TabIndex        =   2
         Top             =   870
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   0
      TabIndex        =   13
      Top             =   1500
      Width           =   8025
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
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
      Left            =   8205
      TabIndex        =   7
      Top             =   990
      Width           =   1515
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
      Height          =   3615
      Left            =   8025
      TabIndex        =   8
      Top             =   60
      Width           =   30
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
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
      Left            =   8190
      TabIndex        =   6
      Top             =   375
      Width           =   1500
   End
   Begin VB.Label lbl家属余额 
      AutoSize        =   -1  'True
      Caption         =   "家属余额:3333.22"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4830
      TabIndex        =   26
      Top             =   1110
      Width           =   2580
   End
   Begin VB.Label lbl剩余余额 
      AutoSize        =   -1  'True
      Caption         =   "剩余款额:3333.22"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   25
      Top             =   1110
      Width           =   2580
   End
   Begin VB.Label lbl性别 
      AutoSize        =   -1  'True
      Caption         =   "性别:男"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3045
      TabIndex        =   9
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病　人:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   90
      TabIndex        =   24
      Top             =   210
      Width           =   1110
   End
   Begin VB.Label lblMZH 
      AutoSize        =   -1  'True
      Caption         =   "门诊号:99999"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4830
      TabIndex        =   23
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbl姓名 
      AutoSize        =   -1  'True
      Caption         =   "张三"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1215
      TabIndex        =   22
      Top             =   255
      Width           =   570
   End
   Begin VB.Label lbl费用余额 
      AutoSize        =   -1  'True
      Caption         =   "未结费用:3232.22"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4830
      TabIndex        =   21
      Top             =   690
      Width           =   2580
   End
   Begin VB.Label lbl预交余额 
      AutoSize        =   -1  'True
      Caption         =   "预交余额:3232.22"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   10
      Top             =   690
      Width           =   2580
   End
End
Attribute VB_Name = "frmSquareAffirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'入参变量
Private mbytBillType As Byte '0-不区分收费或记帐单,1-收费记录;2-记帐记录
Private mlngModule As Long, mlngPatiID As Long
Private mstrNos As String, mstr医嘱IDs As String, mstrPrivs As String
Private mstrExpand As String
Private mlngCardTypeID As Long, mbln消费卡 As Boolean
Private mstrPrintNO As String
Private mblnCliniqueRoomPay As Boolean  '诊间支付
Private mobjDrugPacker As Object
Private mblnDrugPacker As Boolean
Private mobjDrugMachine As Object
Private mblnDrugMachine As Boolean
Private mbln使用预交 As Boolean '是否允许使用预交款,104381
'---------------------------------------------------------------------
'模块变量
Private mlng结帐ID As Long, mblnOk As Boolean
Private mcolPayMode As Collection
Private mrsInfo As ADODB.Recordset
Private mblnFirst As Boolean
Private mlng就诊卡长度 As Long
Private mlng三方卡长度 As Long
Private mlng卡类别ID As Long '通过病人刷卡的卡类别ID
Private mstr结帐IDs As String    '结帐ID,用逗号分离,返回的结帐ID情况
'---------------------------------------------------------------------
'模块参数
Private mblnReadCard As Boolean  '正在读取卡号
Private mintFeePrecision  As Integer
Private mstrFeePrecisionFmt  As String
Private mbytFeeMoneyPrecision  As Byte
Private mstrFeeMoneyPrecisionFmt As String
Private mblnSeekName As Boolean    '是否通过姓名进行模糊查找
Private mintNameDays As Integer  '通过姓名模糊查找天数
Private mblnBrushCardPass As Boolean         '刷卡要求输入密码
Private mdbl帐户余额 As Double
Private mstrCardNo As String  '三方刷卡的卡号
Private mstr限制类别 As String
Private Type Ty_Para
        int审核票据格式 As Integer
        int收费票据格式 As Integer
        int审核打印方式 As Integer
        int收费打印方式 As Integer
        int药品单位 As Integer
End Type
Private mintCurType As Integer '1-门诊收费;2-门诊记帐
Private mPara As Ty_Para
Private mbytAssign As Byte '发药窗口动态分配方式(0,1)
'---------------------------------------------------------------------
'接口相关
Private WithEvents mobjIDCard As zlIDCard.clsIDCard  '身份证接口
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard   'IC卡接口
Attribute mobjICCard.VB_VarHelpID = -1
Private mobjCardPay As Object    '三方机构接口或读卡接口
Private mblnPassInputCardNo As Boolean  '是否密文输入卡号
Private mblnDefaultPassInputCardNo As Boolean '缺省刷卡是否密文输入卡号
Private mlng医疗卡长度  As Long
Private mblnPayCardNoPass As Boolean
'----------------------------------------------------------------------------
Private Type TY_ChargeMoney
    dbl本次消费合计 As Double
    dbl本次冲预交  As Double
    dbl当前未付 As Double
    dbl预交余额 As Double
    dbl费用余额 As Double
    dbl可用预交 As Double
End Type
Private mCurCarge As TY_ChargeMoney
'------------------------------------------------------------------------------------------
Private mobjCard As clsCards

'卡支付相关
Private Type TY_PayMoney
    lng医疗卡类别ID As Long
    bln消费卡 As Boolean
    str结算方式 As String
    str名称 As String
    str刷卡卡号 As String
    str刷卡密码 As String
    str交易流水号 As String
    str交易说明 As String
    bln读卡 As Boolean
    bln卡号密文  As Boolean
    int医疗卡长度 As Integer
    bln支票 As Boolean
    bln自制卡 As Boolean
    blnOneCard As Boolean '是否一卡通结算
    int性质 As Integer '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算;<0 表示第三方支付
    strNO As String
    lngID As Long '预交ID
    lng结帐ID As Long
    objCard As clsCard
End Type
Private mCurCardPay As TY_PayMoney '本次卡支付
Private mcllSquareBalance As Collection '消费卡结算


Private mblnOK_Click As Boolean  '点击的是确定:59412
'----------------------------------------------------------------------------
Private mPatiCard As SquareCard '刷卡卡相关
Private mstrPassWord As String
Private mobjPatiCardObject As clsCardObject
Private mrsFeeData As ADODB.Recordset   '记录本次刷卡消费的数据
Private mfrMain As Object
Private mstr家属IDs As String '病人家属ID,79868
Private mdbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证

'药房、窗口控制
Private mlng西药房 As Long '指定的西药房,0为动态分配
Private mlng中药房 As Long '指定的中药房,0为动态分配
Private mlng成药房 As Long '指定的成药房,0为动态分配
Private mlng发料部门 As Long '指定的卫材发料部门,0为动态分配

Private mstr西窗 As String  '指定的西药房发药窗口,空为动态分配
Private mstr中窗 As String '指定的中药房发药窗口,空为动态分配
Private mstr成窗 As String  '指定的成药房发药窗口,空为动态分配
Private mstrPayDrugWins As String '发药窗口字符串，格式：执行部门1;发药窗口1|...

Private Function zlGetFeeData(ByVal lng病人ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次收取的费用数据
    '返回:获取费用数据
    '编制:刘兴洪
    '日期:2011-09-14 20:09:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTableNos As String, strTableIDs As String
    Dim varPara() As Variant, strWhere As String, strSubTable As String
    Dim rsTemp As ADODB.Recordset
    Dim strSfTable As String, strJzTable As String
    
     On Error GoTo errHandle
    If lng病人ID = 0 Then Exit Function
    ReDim Preserve varPara(0 To 1) As Variant
    
    varPara(0) = lng病人ID: varPara(1) = mbytBillType
         
    
    If mstr医嘱IDs <> "" Then
          If zlGetSubTable(0, mstr医嘱IDs, strTableIDs, varPara(), 2) = False Then Exit Function
    End If
    If mstrNos <> "" Then
          If zlGetSubTable(1, mstrNos, strTableNos, varPara(), UBound(varPara) + 1) = False Then Exit Function
    End If
 
    If mstr医嘱IDs <> "" And mstrNos <> "" Then
        strSubTable = " With  医嘱  As (" & strTableIDs & "),单据 as (" & strTableNos & ")"
        'strSQL = strSubTable & strSQL & " And A.NO= C.Column_Value And A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where A.病人ID=[1] and A.医嘱序号=P.ID)"
    ElseIf mstr医嘱IDs <> "" Then
        strSubTable = " With  医嘱  As (" & strTableIDs & ") "
'        strSQL = strSubTable & strSQL & " And  A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where A.病人ID=[1] and A.医嘱序号=P.ID)"
    ElseIf strTableNos <> "" Then
        strSubTable = " With   单据 as (" & strTableNos & ")"
        'strSQL = strSubTable & strSQL & " And A.NO= C.Column_Value  "
    End If
    '110421:李南春,2017/6/23,费用执行时应使用价格父号而不是从属父号
    strSfTable = "": strJzTable = ""
    If mbytBillType <= 1 Then
        strSfTable = "" & _
        "Select /*+ rule */ decode(A.记录性质,1,'收费',2,'记帐',4,'挂号') as 类别,A.记录性质,A.执行部门ID,A.发药窗口,A.病人ID, " & vbNewLine & _
        "       A.NO,nvl(A.价格父号,A.序号) as 序号,B.编码||'-'||B.名称 as 项目,B.规格,nvl(A.付数,1)*A.数次 as 数次, " & vbNewLine & _
        "       B.计算单位,A.收费细目ID,A.标准单价,A.应收金额,A.实收金额,A.收费类别,A.登记时间,a.门诊标志" & vbNewLine & _
        "From 门诊费用记录 A,收费项目目录 B" & IIf(mstrNos <> "", " ,单据 C", "") & vbNewLine & _
        "Where A.收费细目ID=B.ID And A.记录性质=1  And A.病人ID=[1] And  A.记录状态=0 "
        If mstr医嘱IDs <> "" And mstrNos <> "" Then
            '问题:49593
            strSfTable = strSfTable & " And (A.NO= C.Column_Value  or  A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where J.病人ID=[1] and J.医嘱序号=P.ID And J.记录性质=1  ))"
        ElseIf mstr医嘱IDs <> "" Then
            strSfTable = strSfTable & " And  A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where J.病人ID=[1] and J.医嘱序号=P.ID And J.记录性质=1)"
        ElseIf strTableNos <> "" Then
            strSfTable = strSfTable & " And A.NO= C.Column_Value  "
        End If
    End If
    If mbytBillType = 2 Or mbytBillType = 0 Then
        strJzTable = "" & _
        "Select /*+ rule */ decode(A.记录性质,1,'收费',2,'记帐',4,'挂号') as 类别,A.记录性质,A.执行部门ID,A.发药窗口,A.病人ID, " & vbNewLine & _
        "       A.NO,nvl(A.价格父号,A.序号) as 序号,B.编码||'-'||B.名称 as 项目,B.规格,nvl(A.付数,1)*A.数次 as 数次, " & vbNewLine & _
        "       B.计算单位,A.收费细目ID,A.标准单价,A.应收金额,A.实收金额,A.收费类别,A.登记时间,a.门诊标志" & vbNewLine & _
        "From 门诊费用记录 A,收费项目目录 B" & IIf(mstrNos <> "", " ,单据 C", "") & vbNewLine & _
        "Where A.收费细目ID=B.ID And A.记录性质=2  And A.病人ID=[1] And  A.记录状态=0 "
        If mstr医嘱IDs <> "" And mstrNos <> "" Then
            '问题:49593
            'strJzTable = strJzTable & " And A.NO= C.Column_Value And A.医嘱序号=P.ID"
            strJzTable = strJzTable & " And (A.NO= C.Column_Value  or  A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where J.病人ID=[1] and J.医嘱序号=P.ID And J.记录性质=2  ))"
        ElseIf mstr医嘱IDs <> "" Then
            strJzTable = strJzTable & " And   A.NO in (Select Distinct NO From 门诊费用记录 J,医嘱 P Where J.病人ID=[1] and J.医嘱序号=P.ID And J.记录性质=2  ) "
        ElseIf strTableNos <> "" Then
            strJzTable = strJzTable & " And A.NO= C.Column_Value "
        End If
        If strSfTable <> "" Then strJzTable = vbCrLf & " Union all   " & vbCrLf & strJzTable
    End If
    strSQL = strSubTable & vbCrLf & strSfTable & vbCrLf & strJzTable
    strSQL = "  Select /*+ rule */  * From (" & strSQL & ") Order by 记录性质,NO,序号"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "获取病人费用信息", varPara)
    Set zlGetFeeData = rsTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function LoadFeeData(ByVal intTYPE As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费用数据
    ' 参数:intType-1-门诊收费;2-记帐
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-15 14:33:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTableNos As String, strTableIDs As String
    Dim varPara() As Variant, strWhere As String, strSubTable As String
    Dim lngRow As Long
    Dim dblMoney As Double, i As Long
    mintCurType = intTYPE
    If mrsFeeData Is Nothing Then Exit Function
    mrsFeeData.Filter = "记录性质=" & intTYPE
    With vsFee
        .Clear 1
        .Rows = IIf(mrsFeeData.RecordCount = 0, 1, mrsFeeData.RecordCount) + 1
        lngRow = 1
        If mrsFeeData.RecordCount <> 0 Then mrsFeeData.MoveFirst
        Do While Not mrsFeeData.EOF
            .RowData(lngRow) = Val(Nvl(mrsFeeData!序号))
            .TextMatrix(lngRow, .ColIndex("类别")) = Nvl(mrsFeeData!类别)
            .Cell(flexcpData, lngRow, .ColIndex("类别")) = Val(Nvl(mrsFeeData!记录性质))
            .TextMatrix(lngRow, .ColIndex("单据号")) = Nvl(mrsFeeData!NO)
            .Cell(flexcpData, lngRow, .ColIndex("单据号")) = Trim(Nvl(mrsFeeData!收费类别))
            .TextMatrix(lngRow, .ColIndex("项目")) = Nvl(mrsFeeData!项目)
            .TextMatrix(lngRow, .ColIndex("规格")) = Nvl(mrsFeeData!规格)
            .TextMatrix(lngRow, .ColIndex("数次")) = FormatEx(Val(Nvl(mrsFeeData!数次)), 5)
            .TextMatrix(lngRow, .ColIndex("单位")) = Nvl(mrsFeeData!计算单位)
            .TextMatrix(lngRow, .ColIndex("单价")) = FormatEx(Val(Nvl(mrsFeeData!标准单价)), mintFeePrecision)
            .TextMatrix(lngRow, .ColIndex("应收金额")) = FormatEx(Val(Nvl(mrsFeeData!应收金额)), mbytFeeMoneyPrecision)
            .TextMatrix(lngRow, .ColIndex("实收金额")) = FormatEx(Val(Nvl(mrsFeeData!实收金额)), mbytFeeMoneyPrecision)
            .Cell(flexcpData, lngRow, .ColIndex("实收金额")) = Val(Nvl(mrsFeeData!实收金额))
            .TextMatrix(lngRow, .ColIndex("门诊标志")) = Val(Nvl(mrsFeeData!门诊标志))
            dblMoney = dblMoney + Val(Nvl(mrsFeeData!实收金额))
            lngRow = lngRow + 1
            mrsFeeData.MoveNext
        Loop
    End With
    mrsFeeData.Filter = 0
    dblMoney = RoundEx(dblMoney, 2)
    mCurCarge.dbl本次消费合计 = dblMoney
    mCurCarge.dbl当前未付 = dblMoney
    mCurCarge.dbl本次冲预交 = 0
    lbl自付合计.Caption = Format(dblMoney, "####0.00;-###0.00;;")
    lbl自付合计.Tag = dblMoney
    '设置具体的界面属性
    LoadFeeData = True
End Function

Public Function zlSquareAffirm(ByVal frmMain As Object, _
    ByVal lngModule As Long, strPrivs As String, _
    Optional ByVal lngPatiID As Long = 0, _
    Optional ByVal lngCardTypeID As Long = 0, _
    Optional ByVal bln消费卡 As Boolean = False, _
    Optional ByVal blnCliniqueRoomPay As Boolean = False, _
    Optional ByVal bytBillType As Byte, _
    Optional ByVal strNos As String = "", _
    Optional ByVal str医嘱IDs As String = "", _
    Optional ByRef strExpand As String = "", _
    Optional ByRef lng结帐ID As Long = 0, _
    Optional ByVal bln使用预交 As Boolean = True) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 消费确认接口 , 主要是应用于病人在各消费环境进行消费确认
    '入参:frmMain-传入调用对象
    '       lngModule:调用的模块号
    '       strPrivs:权限串
    '       lngPatiID :病人ID,可以不传,在本接口窗体中刷卡!
    '       lngCardTypeID   Long    In  卡类别ID(消费卡为消费接口序号):0为不区分;在确认窗口中处理 目前 , 只有在预交款缴款中使用,传入后,支付方式缺省为该方式.
    '       bln消费卡   Boolean In  缺省为Fase,表示是否消费卡结算
    '       bytBillType:单据类别: 0-不区分收费或记帐单,1-收费记录;2-记帐记录
    '       strNOs:格式为( 单据1,单据2),配合BytBillType单据类型使用.一次只能使用一种性质
    '                   如:  A0001,A002,A003…;
    '       str医嘱IDs:格式为:ID1,ID2,...
    '       strCardNO-主界面中刷的卡号
    '       blnCliniqueRoomPay-诊间支付(诊间支付不弹出刷卡界面),诊间支付时，只针对收费性质
    '       bln使用预交-是否允许使用预交：Ture，允许使用预交款，且存在预交款时缺省使用预交款；False，不允许使用预交款，必须要有启用的三方帐户
    '出参:
    '返回:Boolean 返回    成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-06-15 09:53:37
    '说明:
    '      如果strNos和str医嘱IDs都没传,只是对指定病人的门诊收费划价单收费和门诊记帐划价进行审核.
    '      如果病人ID不传入,则需要在窗体中先进行刷卡找到病人后,再进行消费确认.
    '调用者:
    '    1.  检查;检验;药房等.
    '    2.  其他所有需要进行消费确认的地方都应该调用该接口.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strPayType As String
    On Error GoTo errHandle
    mblnDrugPacker = False:  mblnDrugMachine = False
    mstrExpand = strExpand
    mlngModule = lngModule: mlngPatiID = lngPatiID: mstrPrivs = strPrivs
    mstrNos = strNos: mstr医嘱IDs = str医嘱IDs: mlng结帐ID = 0: mblnOk = False
    mbytBillType = bytBillType: mlngCardTypeID = lngCardTypeID
    mblnCliniqueRoomPay = blnCliniqueRoomPay
    mbln使用预交 = bln使用预交
    '检查诊间支付
    If CliniqueRoomPayValied = False Then Exit Function
    If bln使用预交 = False Then
        strPayType = GetAvailabilityCardType
        If strPayType = "" Then
            MsgBox "注意:" & vbCrLf & "    当前没有可用的支付类别，不能支付！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If mblnCliniqueRoomPay Then  '诊间支付需要设置相关的参数
        Call InitPara
    End If
    Set mrsFeeData = zlGetFeeData(lngPatiID)
    If mrsFeeData Is Nothing Then Exit Function
    If mrsFeeData.State <> 1 Then Exit Function
    If mrsFeeData.RecordCount = 0 Then zlSquareAffirm = True: Exit Function
    
    '95366:李南春,2016/4/19,收取药品费用调用包药机
    Call CreateDrugPacker
    
    Set mfrMain = frmMain
    If mblnCliniqueRoomPay Then
        mblnOk = False
        If ExecuteCliniqueRoomPay = False Then
            Exit Function
        End If
        lng结帐ID = mlng结帐ID
        mblnOk = True
        zlSquareAffirm = mblnOk
        Exit Function
    End If
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    lng结帐ID = mlng结帐ID
    zlSquareAffirm = mblnOk
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除数据
    '编制:刘兴洪
    '日期:2011-06-20 09:29:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mrsInfo = New ADODB.Recordset
    mstr家属IDs = ""
    lbl姓名.Caption = ""
    lbl性别.Caption = "性别:"
    lbl预交余额.Caption = "预交余额:0.00"
    lbl费用余额.Caption = "未结费用:0.00"
    lbl剩余余额.Caption = "剩余款额:0.00"
    lbl家属余额.Caption = "家属余额:0.00"
    lbl家属余额.Visible = False
    lbl自付合计.Caption = "0.00"
    txt冲预交.Text = ""
    txt金额.Text = ""
    vsFee.Clear 1: vsFee.Rows = 2
End Sub
Private Sub cbo支付方式_Click()
    Dim i As Long, lngIndex As Long
    '记帐不处理
    With mCurCardPay
        .lng医疗卡类别ID = 0
        .bln消费卡 = False
        .str结算方式 = ""
        .str名称 = ""
        .str刷卡卡号 = ""
        .str刷卡密码 = ""
        .lngID = 0
        .strNO = ""
        .str名称 = ""
        .bln卡号密文 = False
        .int医疗卡长度 = 0
        .bln读卡 = False
        .bln支票 = False
        .blnOneCard = False
        .bln自制卡 = False
        .int性质 = 0
     End With
    If mintCurType = 2 Then Exit Sub
    With cbo支付方式
        If .ListIndex = -1 Then GoTo SetProperty:
        lngIndex = .ListIndex + 1
        mCurCardPay.int性质 = .ItemData(.ListIndex)
        mCurCardPay.blnOneCard = .ItemData(.ListIndex) = 7
    End With
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not mcolPayMode Is Nothing Then
        With mCurCardPay
            .lng医疗卡类别ID = Val(mcolPayMode(lngIndex)(3))
            .bln消费卡 = Val(mcolPayMode(lngIndex)(5)) = 1
            .str结算方式 = Trim(mcolPayMode(lngIndex)(6))
            .str名称 = Trim(mcolPayMode(lngIndex)(1))
            .bln读卡 = Val(mcolPayMode(lngIndex)(2)) = 0
            .bln自制卡 = Val(mcolPayMode(lngIndex)(8)) = 1
         End With
     Else
            mCurCardPay.str结算方式 = zlstr.NeedName(cbo支付方式.Text)
     End If
    '创建卡对象
    Call CreatePayObject
SetProperty:
End Sub
Private Sub CreatePayObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建支付对象接口
    '编制:刘兴洪
    '日期:2011-06-22 13:15:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng卡类别ID As Long, bln消费卡 As Boolean, int自动读取 As Integer
    Dim strKey As String
    Dim i As Long
    Set mobjCardPay = Nothing:
    Err = 0: On Error Resume Next
    
    If zlGetCardObj(Me, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mobjPatiCardObject) = False Then
        Set mobjPatiCardObject = Nothing
        Set mobjCardPay = Nothing
        Exit Sub
    End If
    Set mobjCardPay = mobjPatiCardObject.CardObject
    If Err <> 0 Then
        MsgBox "未找到" & mCurCardPay.str名称 & "所对应的部件,请检查", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If mobjCardPay Is Nothing Then Exit Sub
End Sub

Private Function GetSelectNOs(ByRef str费用来源 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选择的单据号
    '返回:单据号,单据之间用逗号分离,如:A0001,A0002....
    '编制:刘兴洪
    '日期:2011-06-23 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strNos As String, strNO As String
    
    If mblnCliniqueRoomPay Then
        '诊间支付,需要特别处理
        mrsFeeData.Filter = "记录性质=1"
        mrsFeeData.Sort = "NO"
        strNos = ""
        With mrsFeeData
            Do While Not .EOF
                strNO = Nvl(!NO)
                If strNO <> "" Then
                    If InStr(1, strNos & ",", "," & strNO & ",") = 0 Then
                        strNos = strNos & "," & strNO
                        If InStr(str费用来源, Decode(Val(Nvl(!门诊标志)), 4, 3, 2, 2, 1)) = 0 Then
                            str费用来源 = str费用来源 & "," & Decode(Val(Nvl(!门诊标志)), 4, 3, 2, 2, 1)
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
        If strNos <> "" Then strNos = Mid(strNos, 2)
        GetSelectNOs = strNos
        Exit Function
    End If
    
    With vsFee
        For i = 1 To .Rows - 1
            strNO = Trim(.TextMatrix(i, .ColIndex("单据号")))
            If strNO <> "" Then
                If InStr(1, strNos & ",", "," & strNO & ",") = 0 Then
                    strNos = strNos & "," & strNO
                    If InStr(str费用来源, Decode(Val(.TextMatrix(i, .ColIndex("门诊标志"))), 4, 3, 2, 2, 1)) = 0 Then
                        str费用来源 = str费用来源 & "," & Decode(Val(.TextMatrix(i, .ColIndex("门诊标志"))), 4, 3, 2, 2, 1)
                    End If
                End If
            End If
        Next
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If str费用来源 <> "" Then str费用来源 = Mid(str费用来源, 2)
    GetSelectNOs = strNos
End Function

Private Function GetSelectNOsAndSerialNum(ByRef strNos As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选择的单据号和单据中的序号
    '返回:单据号,单据之间用逗号分离,如:A0001|0;1;2,A0002|1;2;3....
    '编制:刘兴洪
    '日期:2011-06-23 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strNO As String
    Dim str序号 As String, strData As String
    Dim j As Long
    With vsFee
        strData = "": strNos = ""
        For i = 1 To .Rows - 1
            strNO = Trim(.TextMatrix(i, .ColIndex("单据号")))
            If InStr(1, strNos & ",", "," & strNO & ",") = 0 Then
                    str序号 = ""
                    For j = 1 To .Rows - 1
                        If strNO = Trim(.TextMatrix(j, .ColIndex("单据号"))) Then
                            str序号 = str序号 & ";" & .RowData(j)
                        End If
                    Next
                    If str序号 <> "" Then str序号 = Mid(str序号, 2)
                    strNos = strNos & "," & strNO
                    strData = strData & "," & strNO & "|" & str序号
             End If
        Next
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If strData <> "" Then strData = Mid(strData, 2)
    GetSelectNOsAndSerialNum = strData
End Function
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据合法性检查
    '返回:数据合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-22 15:28:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mrsInfo Is Nothing Then
        MsgBox "病人信息不能确定,请检查!", vbInformation + vbOKOnly, gstrSysName
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
        Exit Function
    End If
    
    If mrsInfo.State <> 1 Then
        MsgBox "病人信息不能确定,请检查!!", vbInformation + vbOKOnly, gstrSysName
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
        Exit Function
    End If
    If mintCurType <> 2 Then
        '79621:李南春,2014/11/14,对金额格式化处理
        If RoundEx(Val(txt冲预交.Text) + Val(txt金额), 2) <> RoundEx(Val(lbl自付合计.Tag), 2) Then
            If Val(txt金额) = 0 Then
                MsgBox "病人的预存款余额不足,请充值!", vbInformation + vbOKOnly, gstrSysName
            Else
                MsgBox "本次支付款项合计与本次需要支付的合计不等，请充值!", vbInformation + vbOKOnly, gstrSysName
            End If
            If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
            Exit Function
        End If
    End If
    If (Val(txt冲预交.Text) > 0 Or mintCurType = 2) And Val(lbl预存款.Tag) = 0 Then
        '证明没有验证卡，需要输入密码验证
          If CheckPrepayMoneyIsValied = False Then Exit Function
    End If
    
    isValied = True
End Function
Private Function CheckPrepayMoneyIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查预交数据输入是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-24 10:36:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If BrushcardStrikePrepay = False Then
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交
        Exit Function
    End If
    CheckPrepayMoneyIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub setControlMove()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件属性
    '编制:刘兴洪
    '日期:2011-08-12 10:43:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngTop As Single, sngSplitHeight As Single, bln预交 As Boolean
    Dim sngHeght As Single
    sngSplitHeight = 80
    
    bln预交 = mCurCarge.dbl可用预交 <> 0 Or cbo支付方式.ListCount = 0
    If mintCurType = 1 Then
        ' 收费
        lbl预存款.Visible = bln预交: txt冲预交.Visible = bln预交
        sngHeght = picPayMode.ScaleHeight
        sngHeght = sngHeght - IIf(bln预交, txt冲预交.Height - sngSplitHeight, 0)
        If cbo支付方式.ListCount = 0 Then
            sngTop = (sngHeght + sngSplitHeight) / 2
        Else
            sngHeght = sngHeght - cbo支付方式.Height - sngSplitHeight
            sngHeght = sngHeght - txt金额.Height
            sngTop = sngHeght / IIf(bln预交, 3, 2)
        End If
        If bln预交 Then
            txt冲预交.Top = sngTop: sngTop = txt冲预交.Top + txt冲预交.Height + sngSplitHeight
        End If
        cbo支付方式.Top = sngTop: sngTop = cbo支付方式.Top + cbo支付方式.Height + sngSplitHeight
        txt金额.Top = sngTop
        lbl预存款.Top = txt冲预交.Top + (txt冲预交.Height - lbl预存款.Height) \ 2
        lbl支付方式.Top = cbo支付方式.Top + (cbo支付方式.Height - lbl支付方式.Height) \ 2
        lbl金额.Top = txt金额.Top + (txt金额.Height - lbl金额.Height) \ 2
        Exit Sub
    End If
    '记帐
    sngHeght = picPayMode.ScaleHeight
    sngHeght = sngHeght - txt冲预交.Height
    sngTop = sngHeght / 2
    txt冲预交.Top = sngTop
    lbl预存款.Top = txt冲预交.Top + (txt冲预交.Height - lbl预存款.Height) \ 2
    cbo支付方式.Visible = False: lbl支付方式.Visible = False
    txt金额.Visible = False: lbl金额.Visible = False
End Sub

Private Function BrushcardStrikePrepay() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证刷卡冲预交
    '返回:冲销成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Val(lbl预存款.Tag) = 1 Then BrushcardStrikePrepay = True: Exit Function
    If Val(txt冲预交) = 0 And mintCurType <> 2 Then BrushcardStrikePrepay = True: Exit Function
    If mintCurType <> 2 Then If CheckPrepayValied = False Then Exit Function
     '刷卡确认
    'frmParent As Object, ByVal lngSys As Long, _
    ByVal lng病人ID As Long, ByVal cur金额 As Currency, _
    Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0
    gblnNotCloseWindows = True
    If zlDatabase.PatiIdentify(Me, glngSys, mlngPatiID, Val(txt冲预交), mlngModule, 1, mlngCardTypeID, IIf(-1 * mdbl预存款消费验卡 >= Val(txt冲预交), False, True), True, _
        mstr家属IDs, (mdbl预存款消费验卡 <> 0), (mdbl预存款消费验卡 = 2)) Then
        gblnNotCloseWindows = False
        lbl预存款.Tag = "1"
        txt冲预交.BackColor = Me.BackColor
        txt冲预交.Tag = Val(txt冲预交): txt冲预交.Enabled = False
        Call cbo支付方式_Click
        '59412
        If mblnOK_Click Then BrushcardStrikePrepay = True: Exit Function
        If RoundEx(txt冲预交.Text, 5) = RoundEx(Val(lbl自付合计.Tag), 5) Or mintCurType = 2 Then
            '相等时,保存数据
            If zlExcuteAffirm = False Then
                lbl预存款.Tag = "": txt冲预交.Enabled = True: txt冲预交.BackColor = vbWhite
                txt冲预交.Tag = ""
                If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
                zlControl.TxtSelAll txt冲预交
                Exit Function
            End If
        Else
           If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
            zlControl.TxtSelAll txt金额
        End If
        BrushcardStrikePrepay = True
        Exit Function
    Else
        lbl预存款.Tag = "": txt冲预交.Enabled = True
        txt冲预交.Tag = ""
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交
        gblnNotCloseWindows = False
        Call cbo支付方式_Click
       Exit Function
    End If
    BrushcardStrikePrepay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo支付方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function WriteInforToCard(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String, _
        ByVal objSquareCard As Object, ByVal lngCardTypeID As Long, _
        ByVal strNos As String) As Boolean
    '功能:将门诊信息写入卡中
    '入参：
    '    frmMain - 调用窗体
    '    lngModul - 模块号
    '    strPrivs - 权限串
    '    objSquareCard - 医疗卡对象
    '    strNOs - 单据号，格式：'A0001','A0002','A0003',...或A0001,A0002,A0003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng病人ID As Long, lng结算序号 As Long
    
    Err = 0: On Error GoTo errH:
    '问题:56615
'    If InStr(strPrivs, ";门诊信息写卡;") = 0 Then Exit Function
    
    strSQL = "Select Distinct A.病人ID,B.结算序号" & _
        " From 门诊费用记录 A,病人预交记录 B,Table( f_Str2list([1])) J" & _
        " Where A.结帐ID=B.结帐ID And A.NO=J.Column_Value And  Nvl(A.附加标志,0)<>9 And A.记录性质 = 1 " & _
        "       And A.记录状态 in(1,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取单据结算序号", Replace(strNos, "'", ""))
    If rsTemp.EOF Then Exit Function
    Do While Not rsTemp.EOF
        lng病人ID = Val(Nvl(rsTemp!病人ID))
        lng结算序号 = Val(Nvl(rsTemp!结算序号))
        '调用健康卡写卡接口
        If lng病人ID <> 0 And lng结算序号 <> 0 Then
            Call objSquareCard.zlMzInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng病人ID, lng结算序号)
        End If
        rsTemp.MoveNext
    Loop
    
    WriteInforToCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlExcuteAffirm() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行消费确认
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-14 22:46:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '数据校对
    If isValied = False Then
        Exit Function
    End If
    If SaveData = False Then mstrPrintNO = "": Exit Function
    '打印票据
    Call PrintBill
    
    '银医一卡通写卡，85950
    If mintCurType = 1 Then '门诊划价收费
        Call WriteInforToCard(Me, mlngModule, mstrPrivs, mPatiCard.objSquareCard, 0, mstrPrintNO)
    End If
    Set mPatiCard.objSquareCard = Nothing
    
    If mbytBillType = 0 And mintCurType = 1 Then
        mintCurType = 2
        Call LoadFeeData(2)
        setControlMove
        If vsFee.TextMatrix(1, vsFee.ColIndex("单据号")) = "" Then
            mblnOk = True
            Unload Me
        End If
        Exit Function
    End If
    mblnOk = True: Unload Me
    zlExcuteAffirm = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub PrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印票据
    '编制:刘兴洪
    '日期:2014-01-20 11:01:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean, strFormat As String
    Dim frmMain As Object
    If mblnCliniqueRoomPay Then
        Set frmMain = mfrMain
    Else
        Set frmMain = Me
    End If
    Select Case mbytBillType
    Case 1, 4, 5
        blnPrint = mPara.int收费打印方式 = 1
        If mPara.int收费打印方式 = 2 Then
            If MsgBox("你是否真的要打印清单吗?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrint = True
            End If
        End If
        If blnPrint Then
            strFormat = IIf(mPara.int收费票据格式 = 0, "", "ReportFormat=" & mPara.int收费票据格式)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", frmMain, "NO=" & mstrPrintNO, "药品单位=" & mPara.int药品单位, "PrintEmpty=0", strFormat, 2)
        End If
    Case 2
        blnPrint = mPara.int审核打印方式 = 1
        If mPara.int审核打印方式 = 2 Then
            If MsgBox("你是否真的要打印清单吗?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrint = True
            End If
        End If
        If blnPrint Then
            strFormat = IIf(mPara.int审核票据格式 = 0, "", "ReportFormat=" & mPara.int审核票据格式)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", frmMain, "NO=" & mstrPrintNO, "药品单位=" & mPara.int药品单位, "PrintEmpty=0", strFormat, 2)
        End If
    End Select
End Sub
Private Sub cmdOK_Click()
     mblnOK_Click = True
    Call zlExcuteAffirm
End Sub
Private Function VerifyFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:审核费用
    '返回:审核成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2011-06-23 09:59:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, cllPro As New Collection
    Dim varData As Variant, i As Long, strNO As String, strNos As String
    Dim strNosData As String, varTemp As Variant, str序号 As String
    strNosData = GetSelectNOsAndSerialNum(strNos)
     '记帐的话,要费用报警
    If Not zlAuditingWarn(mstrPrivs, strNos, Val(Nvl(mrsInfo!病人ID))) Then Exit Function
    varData = Split(strNosData, ",")
    For i = 0 To UBound(varData)
        If varData(i) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            strNO = varTemp(0): str序号 = Replace(varTemp(1), ";", ",")
            mstrPrintNO = mstrPrintNO & ",'" & strNO & "'"
            'No_In/操作员编号_In /操作员姓名_In /序号_In/审核时间_In
             strSQL = "zl_门诊记帐记录_Verify('" & strNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & str序号 & "')"
             AddArray cllPro, strSQL
        End If
    Next
    If mstrPrintNO <> "" Then mstrPrintNO = Mid(mstrPrintNO, 2)
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllPro, Me.Caption
    VerifyFee = True
    
    '110319
    If mblnDrugMachine Then
        '门诊格式：1|单据1,处方号1;单据2,处方号2
        Dim strData As String, strReturn As String
        strData = "1|" & "9," & Replace(Replace(strNos, "'", ""), ",", ";9,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
    End If
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    mstrPrintNO = ""
End Function
Private Function SaveCharge() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:划价收费
    '返回:收费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-23 11:38:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, lng结帐ID As Long, lng结帐序号 As Long
    Dim dbl误差费 As Double, dbl结算金额 As Double, dbl限定额 As Double, dblMoney As Double
    Dim dblThreeMoney As Double, dblTemp As Double, dbl冲预交 As Double
    Dim strNos As String, strNO As String, str发生时间 As String, strSQL As String, str结帐IDs As String
    Dim str交易流水号 As String, str交易说明 As String, strSwapExtendInfor As String
    Dim cllPro As New Collection
    Dim int病人来源 As Integer, intIndex As Integer, strTemp As String
    Dim rsTemp As ADODB.Recordset, cllDept As Collection
    Dim str发药窗口 As String
    Dim strReturn As String, strData As String
    Dim str费用来源 As String
 
    Err = 0: On Error GoTo Errhand:
    int病人来源 = IIf(Val(Nvl(mrsInfo!在院)) = 1, 2, 1)
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    strNos = GetSelectNOs(str费用来源)
    mstrPrintNO = "'" & Replace(strNos, ",", "','") & "'"
    strSQL = "" & _
    "   Select   /*+ rule */ NO,Max(付款方式) as 付款方式, " & _
    "               Max(病人科室ID) as 病人科室ID,Max(开单部门ID) as 开单部门Id, " & _
    "               Max(发药窗口) as 发药窗口,max(是否急诊) as 是否急诊," & _
    "               Sum(实收金额) as 金额,sum(case when instr([2],','||A.收费类别||',')>0 then a.实收金额 else 0 end)  as 限定额" & _
    "   From 门诊费用记录 A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) J  " & _
    "  Where  A.No=J.Column_value and 记录状态=0 and A.记录性质=1" & _
    "   Group by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "病人消费结算-获取诊疗信息", strNos, mstr限制类别)
    
    With rsTemp
        Do While Not .EOF
            dbl结算金额 = dbl结算金额 + Val(Nvl(rsTemp!金额))
            dbl限定额 = dbl限定额 + Val(Nvl(rsTemp!限定额))
            .MoveNext
        Loop
    End With
    dblTemp = RoundEx(dbl结算金额, 6)
    dbl结算金额 = RoundEx(dbl结算金额, 2)
    dblMoney = dbl结算金额
    dblThreeMoney = dbl结算金额
    dbl误差费 = dblTemp - dblMoney
    dbl限定额 = RoundEx(dbl限定额, 2)
    
        
    If mblnCliniqueRoomPay = False Then '非诊间支付时，需要检查相关的数据合法性
        '79621:李南春,2014/11/14,对金额格式化处理
        If dbl结算金额 <> RoundEx(Val(lbl自付合计.Tag), 2) Then
            If MsgBox("注意:" & vbCrLf & "    你所选择的划价单据的实收金额已经发生变化,是否重新提取相应单据的费用!", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                Set mrsFeeData = zlGetFeeData(Val(Nvl(mrsInfo!病人ID)))
                Call LoadFeeData(mintCurType): Exit Function
            End If
        End If
        '79621:李南春,2014/11/14,对金额格式化处理
        If RoundEx(Val(txt冲预交.Text) + Val(txt金额.Text), 2) <> dbl结算金额 Then
            MsgBox "注意:" & vbCrLf & "    你输入扣款金额不对(预存款+" & cbo支付方式.Text & "支付不等于本次支付的费用合计,请检查!", vbOKOnly + vbDefaultButton1 + vbInformation
            Exit Function
        End If
        If cbo支付方式.ListIndex >= 0 Then
            intIndex = cbo支付方式.ListIndex + 1
            If Trim(mcolPayMode(intIndex)(6)) = "" Then
                MsgBox "注意:" & vbCrLf & "    " & Trim(mcolPayMode(intIndex)(1)) & "  未设置结算方式,请与系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        ElseIf Val(txt金额.Text) <> 0 Then
            MsgBox "注意:" & vbCrLf & "    未选择支付类别!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        dbl冲预交 = RoundEx(Val(txt冲预交.Text), 2)
        dblMoney = RoundEx(Val(txt金额.Text), 2)
        '79621:李南春,2014/11/14,对金额格式化处理
        If RoundEx(Val(lbl自付合计.Tag) - Val(txt金额.Text) - dbl限定额, 2) > RoundEx(Val(txt冲预交.Text), 2) And Val(txt冲预交.Text) <> 0 Then
            MsgBox "注意:" & vbCrLf & "    扣预存款的额度输入过大,最多只能扣预存款:" & Format(Val(lbl自付合计.Tag) - Val(txt金额.Text) - dbl限定额, "0.00"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If RoundEx(dblMoney, 7) <> RoundEx(Val(txt金额.Text), 7) Then
            MsgBox "注意:" & vbCrLf & "    " & Trim(mcolPayMode(intIndex)(1)) & "  支付合计不正确,请检查!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        '79621:李南春,2014/11/17,对金额格式化处理
        If RoundEx(Val(txt冲预交.Text), 4) <> RoundEx(mCurCarge.dbl本次消费合计, 4) Then
           If BrushCardThreeSwapCheck(strNos, Val(txt金额.Text), str费用来源, lng病人ID) = False Then Exit Function
           dblThreeMoney = Val(txt金额.Text)
        End If
    Else
        If BrushCardThreeSwapCheck(strNos, dblThreeMoney, str费用来源, lng病人ID) = False Then Exit Function
    End If
    
    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    lng结帐序号 = -1 * lng结帐ID
    str发生时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    Set cllDept = New Collection
    With mrsFeeData
        strTemp = ""
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(",5,6,7,", Nvl(!收费类别)) > 0 And _
                InStr(strTemp, "," & Nvl(!收费类别) & "|" & Nvl(!执行部门ID) & ",") = 0 Then
                cllDept.Add Array(Nvl(!收费类别), Val(Nvl(!执行部门ID)), Nvl(!发药窗口))
            End If
            .MoveNext
        Loop
        str发药窗口 = GetPayDrugWindow(lng病人ID, CDate(str发生时间), cllDept)
    End With
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
               
            '---------------------------------------------------------------
            'Zl_病人划价收费_Insert
            strSQL = "Zl_病人划价收费_Insert("
            '  No_In         门诊费用记录.NO%Type,
            strSQL = strSQL & "'" & Nvl(rsTemp!NO) & "',"
            '  病人id_In     门诊费用记录.病人id%Type,
            strSQL = strSQL & "" & ZVal(lng病人ID) & ","
            '  病人来源_In   Number,
            strSQL = strSQL & "" & int病人来源 & ","
            '  付款方式_In   门诊费用记录.付款方式%Type,
            If Nvl(mrsInfo!付款方式编码) <> "" Then
               strSQL = strSQL & "'" & Nvl(mrsInfo!付款方式编码) & "',"
            Else
               strSQL = strSQL & "'" & Nvl(rsTemp!付款方式) & "',"
            End If
            '  姓名_In       门诊费用记录.姓名%Type,
            strSQL = strSQL & "'" & Nvl(mrsInfo!姓名) & "',"
            '  性别_In       门诊费用记录.性别%Type,
            strSQL = strSQL & "'" & Nvl(mrsInfo!性别) & "',"
            '  年龄_In       门诊费用记录.年龄%Type,
            strSQL = strSQL & "'" & Nvl(mrsInfo!年龄) & "',"
            '  病人科室id_In 门诊费用记录.病人科室id%Type,
            strSQL = strSQL & "" & IIf(Val(Nvl(rsTemp!病人科室ID)) = 0, "NULL", Val(Nvl(rsTemp!病人科室ID))) & ","
            '  开单部门id_In 门诊费用记录.开单部门id%Type,
            strSQL = strSQL & "" & IIf(Val(Nvl(rsTemp!开单部门ID)) = 0, "NULL", Val(Nvl(rsTemp!开单部门ID))) & ","
            '  开单人_In     门诊费用记录.开单人%Type,
            strSQL = strSQL & "NULL,"    ' 过程内部处理,保持原来的不变
            '  结帐id_In     门诊费用记录.结帐id%Type,
            strSQL = strSQL & "" & lng结帐ID & ","
            '  发生时间_In   门诊费用记录.发生时间%Type,
            strSQL = strSQL & "to_date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'),"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  发药窗口_In   门诊费用记录.发药窗口%Type := Null,
            strSQL = strSQL & "'" & str发药窗口 & "',"
            '  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
            strSQL = strSQL & "" & Val(Nvl(rsTemp!是否急诊)) & ","
            '  登记时间_In   门诊费用记录.登记时间%Type := Null,
            strSQL = strSQL & "to_date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'))"
            zlAddArray cllPro, strSQL
            .MoveNext
        Loop
    End With
    str结帐IDs = lng结帐ID
    mCurCardPay.lng结帐ID = lng结帐ID
    
    'bytType-1-三方接口支付;2-消费卡支付,0-其他
    If mCurCardPay.bln消费卡 And dblMoney <> 0 Then
        If SetCurBalanceSQL(2, lng病人ID, dblMoney, dbl冲预交, 0, 0, dbl误差费, cllPro) = False Then Exit Function
    ElseIf dbl冲预交 = dblThreeMoney Then
        If SetCurBalanceSQL(0, lng病人ID, 0, dbl冲预交, 0, 0, dbl误差费, cllPro) = False Then Exit Function
    Else
        If SetCurBalanceSQL(1, lng病人ID, dblMoney, dbl冲预交, 0, 0, dbl误差费, cllPro) = False Then Exit Function
    End If
    
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If mblnCliniqueRoomPay = False Then
        If Val(txt金额.Text) = 0 Or mCurCardPay.bln消费卡 Then
            '这个肯定是冲预交或者为消费卡在医院的卡帐户
            gcnOracle.CommitTrans
            mstr结帐IDs = str结帐IDs
            mlng结帐ID = lng结帐序号
            SaveCharge = True
            
            GoTo DoDrugPacker:
            Exit Function
        End If
    End If
    
    ' Public Function zlPaymentMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    '    ByVal strCardNo As String, ByVal strBalanceIDs As String,byval strPrepayNos as string , _
    '    ByVal dblMoney As Double, _
    '    ByRef strSwapGlideNO As String, _
    '    ByRef strSwapMemo As String, _
    '    Optional ByRef strSwapExtendInfor As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:帐户扣款交易
    '    '入参:frmMain-调用的主窗体
    '    '        lngModule-调用模块号
    '    '        strBalanceIDs-结帐ID,多个用逗号分离
    '    '       strCardNo-卡号
    '    '       dblMoney-支付金额
    '    '出参:strSwapGlideNO-交易流水号
    '    '       strSwapMemo-交易说明
    '    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '    '返回:扣款成功,返回true,否则返回Flase
    If mblnCliniqueRoomPay Then
        If mobjCardPay.zlPaymentMoney(mfrMain, mlngModule, mCurCardPay.lng医疗卡类别ID, mCurCardPay.str刷卡卡号, str结帐IDs, "", dblThreeMoney, str交易流水号, str交易说明, strSwapExtendInfor) = False Then
                gcnOracle.RollbackTrans: Exit Function
        End If
    Else
        If mobjCardPay.zlPaymentMoney(Me, mlngModule, mCurCardPay.lng医疗卡类别ID, mCurCardPay.str刷卡卡号, str结帐IDs, "", dblThreeMoney, str交易流水号, str交易说明, strSwapExtendInfor) = False Then
                gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    
    Dim cllUpdate As New Collection, cllOthers As New Collection
    Call zlAddUpdateSwapSQL(False, lng结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, str交易流水号, str交易说明, cllUpdate, 0, 1)
    Call zlAddThreeSwapSQLToCollection(False, lng结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllOthers)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, False, True
    Err = 0: On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllOthers, Me.Caption
    SaveCharge = True
    
DoDrugPacker:
    '95366:李南春,2016/4/19,收取药品费用调用包药机
    If mblnDrugMachine Then
        '新版发药机
        '门诊格式：1|单据1,处方号1;单据2,处方号2;...
        strData = "1|" & "8," & Replace(Replace(strNos, "'", ""), ",", ";8,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
    ElseIf mblnDrugPacker Then
        '格式：单据1,处方号1|单据2,处方号2|...
        strData = "8," & Replace(Replace(strNos, "'", ""), ",", "|8,")
        Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.编号, UserInfo.姓名, 0, strData, strReturn)
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    gcnOracle.CommitTrans   '能保存多少,作多少
    Call ErrCenter
    SaveCharge = True
End Function

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '编制:刘兴洪
    '日期:2011-06-22 16:01:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1-收费记录;2-记帐记录；4-挂号记录;5-卡费;10-缴预交
    
    Select Case mintCurType
    Case 1  '收费划价处理
        If SaveCharge = False Then Exit Function
        SaveData = True:
        '打印相关的票据
    Case 2 '划价记帐审核
        If VerifyFee = False Then Exit Function
        SaveData = True: Exit Function
'    Case 10 '缴预交款
'        If SavePrePayMoney = False Then Exit Function
'        SaveData = True
    End Select
End Function
Private Sub cmdPara_Click()
    If frmSquareAffirmParaSet.SetPara(Me) = False Then Exit Sub
    Call InitFactPara
End Sub
 

Private Sub Form_Activate()
    Dim intTYPE As Integer
    
    If mblnCliniqueRoomPay Then Exit Sub
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If GetPatient() = False Then Unload Me: Exit Sub
    '加载费用
    If mbytBillType = 0 Then
        mrsFeeData.Filter = "记录性质=1"
        If mrsFeeData.RecordCount = 0 Then
            intTYPE = 2
        Else
            intTYPE = 1
        End If
        mbytBillType = intTYPE
        mrsFeeData.Filter = 0
    Else
       intTYPE = mbytBillType
    End If
    Call LoadFeeData(intTYPE)
    If mbln使用预交 Then '不允许使用预交款时不加载预交
        '加载预交
        Call Load预交余额(mrsInfo!病人ID)
    End If
    Call SetCtlEnable
    If mCurCarge.dbl可用预交 = 0 Then
        If cbo支付方式.Enabled And txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        '91315,当前预存款为0或者为负数时，仍可以消费成功
        If txt金额.Visible Then txt金额.Text = FormatEx(Val(lbl自付合计.Tag), mbytFeeMoneyPrecision)
        zlControl.TxtSelAll txt金额
    Else
        '79621:李南春,2014/11/17,对金额格式化处理
        If RoundEx(mCurCarge.dbl可用预交, 2) > RoundEx(Val(lbl自付合计.Tag), 2) Then
            txt冲预交.Text = FormatEx(Val(lbl自付合计.Tag), mbytFeeMoneyPrecision)
        Else
            txt冲预交.Text = FormatEx(mCurCarge.dbl可用预交, mbytFeeMoneyPrecision)
        End If
        If Val(txt冲预交.Text) <> 0 And txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交
    End If
    Call setControlMove
    '78773:李南春,2014-10-29,LED显示一卡通支付信息
    Call ShowLedInfor
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2
        If cmdOK.Enabled = False Then Exit Sub
        Call cmdOK_Click: Exit Sub
    Case vbKeyF4
        If Me.ActiveControl Is txt金额 Then
            If cbo支付方式.Enabled = False Then Exit Sub
            If Me.ActiveControl Is txt金额 And txt金额.Enabled = False Then Exit Sub
            If Shift = vbShiftMask Then
                If cbo支付方式.ListIndex - 1 < 0 Then
                    cbo支付方式.ListIndex = cbo支付方式.ListCount - 1
                Else
                    cbo支付方式.ListIndex = cbo支付方式.ListIndex - 1
                End If
            Else
                If cbo支付方式.ListIndex + 1 > cbo支付方式.ListCount - 1 Then
                    cbo支付方式.ListIndex = 0
                Else
                    cbo支付方式.ListIndex = cbo支付方式.ListIndex + 1
                End If
            End If
        End If
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '78773:李南春,2014-10-29,LED显示一卡通支付信息
    If gblnLED Then zl9LedVoice.DisplayPatient ""
    If Not mobjDrugPacker Is Nothing Then Set mobjDrugPacker = Nothing
    If Not mobjDrugMachine Is Nothing Then Set mobjDrugMachine = Nothing
End Sub

Private Sub picFee_Resize()
    Err = 0: On Error Resume Next
    With picFee
        vsFee.Left = .ScaleLeft
        vsFee.Height = .ScaleHeight - vsFee.Top
        vsFee.Width = .ScaleWidth - vsFee.Left
    End With
End Sub
Private Function Load预交余额(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载预交余额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-21 10:47:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    '79868,将病人家属余额加入病人剩余款
    '获得记录集最多只有两条，一条是病人本人的，一条是病人家属的
    Set rsTemp = GetMoneyInfo(lng病人ID, , , 1, , , True)
    Dim dbl病人余额 As Double, dbl费用余额 As Double, dbl家属余额 As Double
    With mCurCarge
        .dbl预交余额 = 0
        .dbl费用余额 = 0
        Do While Not rsTemp.EOF
            .dbl预交余额 = .dbl预交余额 + Val(Nvl(rsTemp!预交余额))
            .dbl费用余额 = .dbl费用余额 + Val(Nvl(rsTemp!费用余额))
            If Nvl(rsTemp!家属, 0) = 0 Then
                dbl病人余额 = Val(Nvl(rsTemp!预交余额))
                dbl费用余额 = Val(Nvl(rsTemp!费用余额))
            Else
                dbl家属余额 = Val(Nvl(rsTemp!预交余额)) - Val(Nvl(rsTemp!费用余额))
            End If
            rsTemp.MoveNext
        Loop
        .dbl可用预交 = .dbl预交余额 - .dbl费用余额
        If .dbl可用预交 < 0 Then .dbl可用预交 = 0
    End With
    lbl预交余额.Caption = "预交余额:" & Format(dbl病人余额, "###0.00;-###0.00;0.00;0.00")
    lbl预交余额.Tag = mCurCarge.dbl预交余额
    lbl费用余额.Caption = "未结费用:" & Format(dbl费用余额, "###0.00;-###0.00;0.00;0.00")
    lbl费用余额.Tag = mCurCarge.dbl费用余额
    lbl剩余余额.Caption = "剩余款额:" & Format(dbl病人余额 - dbl费用余额, "###0.00;-###0.00;0.00;0.00")
    lbl剩余余额.Tag = mCurCarge.dbl可用预交
    lbl家属余额.Caption = "家属余额:" & Format(dbl家属余额, "###0.00;-###0.00;0.00;0.00")
    lbl家属余额.Visible = dbl家属余额 <> 0
    Load预交余额 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体大小
    '编制:刘兴洪
    '日期:2011-09-15 11:26:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
Private Sub Form_Load()
    mblnFirst = True
    If mblnCliniqueRoomPay Then Exit Sub
    If Not IsDesinMode Then
         Call SetWindowsSize
    End If
    zlControl.CboSetWidth cbo支付方式.hWnd, cbo支付方式.Width * 2
    zlControl.PicShowFlat picSum, -1, , 1: zlControl.PicShowFlat picPayMode, -1, , 1
    '初始化数据
    Call InitFace: Call SetCtlEnable
End Sub
Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2011-06-21 13:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKindStr As String, blnVisible As Boolean
    mblnOK_Click = False
    Set mPatiCard = New SquareCard
    
    '创建对象,增加结算卡的结算
    Err = 0: On Error Resume Next
    Set mPatiCard.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        MsgBox "结算卡部件zl9CardSquare.clsCardSquare创建失败！", vbInformation, gstrSysName
        Err = 0: On Error GoTo 0: Exit Sub
    End If
    If mPatiCard.objSquareCard Is Nothing Then Exit Sub
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If mPatiCard.objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
    
    Call InitPara: Call ClearData: Call Load支付方式
End Sub
Private Sub InitFactPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始发票相关的参数
    '编制:刘兴洪
    '日期:2011-08-11 00:24:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With mPara
        .int收费票据格式 = Val(zlDatabase.GetPara("收费收据格式", glngSys, 1151))
        .int收费打印方式 = Val(zlDatabase.GetPara("收费打印方式", glngSys, 1151))
        .int审核票据格式 = Val(zlDatabase.GetPara("审核收据格式", glngSys, 1151))
        .int审核打印方式 = Val(zlDatabase.GetPara("审核打印方式", glngSys, 1151))
        .int药品单位 = Val(zlDatabase.SetPara("药品单位", glngSys, 1151))
    End With
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数值
    '编制:刘兴洪
    '日期:2011-06-20 16:48:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intStart As Integer
    Dim strValue As String
    
    Call InitFactPara
    '门诊病人消费时需要刷卡验证
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    mdbl预存款消费验卡 = Val(Split(strValue, "|")(0))
    '费用单价保留位数
    mintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    mstrFeePrecisionFmt = "0." & String(mintFeePrecision, "0")
    '费用金额小数点位数
    mbytFeeMoneyPrecision = Val(zlDatabase.GetPara(9, glngSys, , 2))
    mstrFeeMoneyPrecisionFmt = "0." & String(mbytFeeMoneyPrecision, "0")
    mblnSeekName = zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModule) = "1"
    mintNameDays = Val(zlDatabase.GetPara("姓名查找天数", glngSys, mlngModule))
    mbytAssign = Val(zlDatabase.GetPara(19, glngSys, , 0))

    If mblnCliniqueRoomPay Then
        '药房、窗口分配方式
        mstr中窗 = zlDatabase.GetPara(49, glngSys, mlngModule)
        mstr西窗 = zlDatabase.GetPara(50, glngSys, mlngModule)
        mstr成窗 = zlDatabase.GetPara(51, glngSys, mlngModule)
        
        mlng西药房 = Val(zlDatabase.GetPara(18, glngSys, mlngModule))
        mlng成药房 = Val(zlDatabase.GetPara(19, glngSys, mlngModule))
        mlng中药房 = Val(zlDatabase.GetPara(20, glngSys, mlngModule))
        mlng发料部门 = Val(zlDatabase.GetPara(21, glngSys, mlngModule))
    Else
        mstr西窗 = "": mstr中窗 = "": mstr成窗 = ""
        mlng中药房 = 0: mlng西药房 = 0: mlng成药房 = 0: mlng发料部门 = 0
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    fraSplitBottom.Width = Me.ScaleWidth + fraSplitBottom.Left
    picFee.Width = Me.ScaleWidth - picFee.Left * 2
    picFee.Height = Me.ScaleHeight - picFee.Top - 50
End Sub

Private Function GetPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=表示是否就诊卡刷卡
    '出参:
    '返回:病人读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-20 16:04:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    '读取病人信息
    strSQL = "" & _
    "   Select Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID,A.病人类型," & _
    "               A.IC卡号,A.就诊卡号,A.门诊号,A.住院号,A.姓名, A.卡验证码, " & _
    "               A.性别,A.年龄, A.出生日期,A.费别,A.担保额,A.医疗付款方式,M.编码 as 付款方式编码,A.在院," & _
    "               decode(B1.病人性质,NULL,0,1,1,0) as 留观,B1.入院日期,A.险类,C.名称 险类名称" & _
    "   From 病人信息 A,病案主页 B1,保险类别 C ,医疗付款方式 M" & _
    "   Where A.险类 = C.序号(+) And A.医疗付款方式=M.名称(+) " & _
    "               And A.病人ID=B1.病人ID(+) And A.主页ID=B1.主页ID(+) " & _
    "               And A.停用时间 is NULL And A.病人ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, "病人消费结算-获取病人信息", mlngPatiID)
    If mrsInfo.EOF Then GoTo NotFoundPati:
    
    If mblnCliniqueRoomPay = False Then
        lbl姓名.Caption = Nvl(mrsInfo!姓名)
        lbl性别.Caption = "性别:" & Nvl(mrsInfo!性别)
        lblMZH.Caption = "门诊号:" & Nvl(mrsInfo!门诊号)
        '74309:李南春，2014-7-7，病人姓名显示颜色处理
        Call SetPatiColor(lbl姓名, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), &HFF0000, vbRed))
    End If
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Set mrsInfo = New ADODB.Recordset
    Call SaveErrLog
    Exit Function
NotFoundPati:
    MsgBox "病人信息未找到,请检查!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    Set mrsInfo = New ADODB.Recordset
End Function
Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-06-21 11:08:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long
    Dim strPayType As String, varData As Variant, varTemp As Variant, i As Long
    j = 0
    '短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    strPayType = GetAvailabilityCardType: varData = Split(strPayType, ";")
    Set mcolPayMode = New Collection
    With cbo支付方式
        .Clear: j = 0
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                mcolPayMode.Add varTemp, "K" & j
                cbo支付方式.AddItem varTemp(1)
                cbo支付方式.ItemData(cbo支付方式.NewIndex) = Val(varTemp(2))
                j = j + 1
            End If
        Next
    End With
    If cbo支付方式.ListCount > 0 And cbo支付方式.ListIndex < 0 Then cbo支付方式.ListIndex = 0
    
End Sub
Private Function CheckPayIsEnough(Optional blnYesNo As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查预付款是否够支付
    '编制:刘兴洪
    '日期:2011-06-21 11:29:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Val(lbl剩余余额.Tag) < Val(lbl自付合计.Tag) Then
        If blnYesNo Then
            If cbo支付方式.Enabled Then
                '可以用其他支付,所以不提醒
                CheckPayIsEnough = True: Exit Function
            End If
            If MsgBox("注意:" & vbCrLf & "   预存款余额不够支付本次费用,请充值" & vbCrLf & "   本次操作是否继续?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                CheckPayIsEnough = True: Exit Function
            End If
            Exit Function
        Else
            '需要排开其他支付方式,检查是否够用
            '79621:李南春,2014/11/14,对金额格式化处理
            If Val(Val(lbl剩余余额.Tag) + Val(txt金额.Text)) < Val(lbl自付合计.Tag) Then
                Call MsgBox("注意:" & vbCrLf & "   预存款余额不够支付本次费用,请充值!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
                Exit Function
            End If
        End If
    End If
    CheckPayIsEnough = True
End Function
Private Sub SetCtlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enable属性
    '编制:刘兴洪
    '日期:2011-06-21 11:19:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cbo支付方式.Enabled = cbo支付方式.ListCount > 0
    txt金额.Enabled = cbo支付方式.Enabled And cbo支付方式.ListCount > 0
    cbo支付方式.Visible = cbo支付方式.ListCount > 0
    lbl支付方式.Visible = cbo支付方式.ListCount > 0
    txt金额.Visible = cbo支付方式.ListCount > 0
    lbl金额.Visible = cbo支付方式.ListCount > 0
    txt冲预交.Enabled = mCurCarge.dbl可用预交 <> 0
End Sub
Private Sub txt冲预交_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Val(txt冲预交) = 0 Then
        If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        Exit Sub
    End If
    mblnOK_Click = False
    If CheckPrepayMoneyIsValied = False Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txt冲预交_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt冲预交, KeyAscii, m金额式)
End Sub

Private Function CheckPrepayValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查预存款数据是否有效
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-14 22:30:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mblnOk Then Exit Function
    If txt冲预交.Text = "" Then
        txt冲预交.Text = "0.00"
    ElseIf Not IsNumeric(txt冲预交.Text) And txt冲预交.Text <> "" Then
        MsgBox "无效数值！", vbInformation, gstrSysName
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Exit Function
    ElseIf Val(txt冲预交.Text) < 0 Then
        MsgBox "预存款冲款金额不能为负！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Exit Function
    ElseIf Val(txt冲预交.Text) > 0 And mCurCarge.dbl当前未付 < 0 Then
        MsgBox "当前应付金额为负时不能使用预存款！", vbInformation, gstrSysName
        txt冲预交.Text = "0.00"
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交:   Exit Function
    '79621:李南春,2014/11/14,对金额格式化处理
    ElseIf RoundEx(Val(txt冲预交.Text), 2) > RoundEx(mCurCarge.dbl可用预交, 2) Then
        MsgBox "预存款冲款金额不能超过病人的预存余额:" & Format(mCurCarge.dbl可用预交, "0.00") & " ！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Exit Function
    ElseIf RoundEx(Val(txt冲预交.Text), 2) > RoundEx(mCurCarge.dbl当前未付, 2) And Val(txt冲预交.Text) <> 0 Then
        MsgBox "预存款冲款金额不能大于应付金额:" & Format(mCurCarge.dbl当前未付, "0.00") & " ！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Exit Function
    Else
        txt冲预交.Text = Format(Val(txt冲预交.Text), "0.00")
    End If
    CheckPrepayValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub txt冲预交_Validate(Cancel As Boolean)
    If lbl预存款.Tag = "1" Or mlngPatiID = 0 Then Exit Sub
    If Val(txt冲预交.Tag) = Val(txt冲预交.Text) Then Exit Sub
    If CheckPrepayValied = False Then Cancel = True: Exit Sub
End Sub

Private Sub txt金额_GotFocus()
    txt金额.Text = Format(RoundEx(Val(lbl自付合计.Tag), 2) - RoundEx(Val(txt冲预交), 2), "####0.00;-###0.00")
    If txt金额.Text < 0 Then txt金额.Text = ""
    zlControl.TxtSelAll txt金额
End Sub

Private Sub txt金额_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt冲预交, KeyAscii, m金额式)
    picPayMode.Tag = ""
    If KeyAscii <> 13 Then Exit Sub
    mblnOK_Click = False
    If Val(txt金额.Text) = 0 Then txt金额.Text = "0.00"
    If txt金额.Text <> "0.00" Then
        If RoundEx(mCurCarge.dbl本次消费合计 - Val(txt冲预交.Text) - Val(txt金额.Text), 7) <> 0 Then
            MsgBox "交易金额输入错误,请重新输入(" & Format(RoundEx(mCurCarge.dbl本次消费合计 - Val(txt冲预交.Text), 7), "0.00") & ")！", vbInformation, gstrSysName
           If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
           zlControl.TxtSelAll txt金额
            picPayMode.Tag = "1"
        End If
        Call cmdOK_Click
        Exit Sub
    End If
End Sub

Private Sub txt金额_Validate(Cancel As Boolean)
    '79621:李南春,2014/11/14,对金额格式化处理
    If RoundEx(Val(txt冲预交) + Val(txt金额.Text), 2) > RoundEx(Val(lbl自付合计.Tag), 2) Then
        If picPayMode.Tag <> "1" Then MsgBox "输入本次支付的预存款金额与" & cbo支付方式 & "支付的合计大于了本次结算费用合计,不能继续!", vbInformation + vbOKOnly, gstrSysName
        txt金额.Text = Format(RoundEx(Val(lbl自付合计.Tag), 2) - RoundEx(Val(txt冲预交), 2), "####0.00;-###0.00")
        If txt金额.Text < 0 Then txt金额.Text = ""
        If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        zlControl.TxtSelAll txt金额
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub lbl预存款_Change()
    lbl预存款.Tag = ""
End Sub
Private Function IsCheckThreeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查三方交易金额输入是否合法
    '返回:合法成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-15 00:03:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mblnOk Then Exit Function
    If Val(txt金额) = 0 Then
        MsgBox "未输入交易金额,请检查!", vbInformation + vbOKOnly
        If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        zlControl.TxtSelAll txt金额
         Exit Function
    End If
    If Not IsNumeric(txt金额.Text) And txt金额.Text <> "" Then
        MsgBox "无效数值！", vbInformation, gstrSysName
        If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        zlControl.TxtSelAll txt金额: Exit Function
    ElseIf Val(txt金额.Text) < 0 Then
        MsgBox "交易金额不能为负！", vbInformation, gstrSysName
        If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        zlControl.TxtSelAll txt金额: Exit Function
    '79621:李南春,2014/11/14,对金额格式化处理
    ElseIf RoundEx(Val(txt金额.Text), 2) > RoundEx(mCurCarge.dbl当前未付, 2) And Val(txt金额.Text) <> 0 Then
        MsgBox "交易金额不能大于本次未付金额:" & Format(mCurCarge.dbl当前未付, "0.00") & " ！", vbInformation, gstrSysName
        If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        zlControl.TxtSelAll txt金额: Exit Function
    End If
    IsCheckThreeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlGetClassMoney(ByVal strNos As String, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    '实收金额:'问题:50339
    strSQL = "" & _
    "   Select  /*+ rule */  A.收费类别,nvl(sum(实收金额) ,0) as  金额   " & _
    "   From 门诊费用记录 A,Table(f_str2List([1])) B " & _
    "   Where A.NO=B.Column_value and A.记录性质=1 and A.记录状态=0 " & _
    "   Group by A.收费类别"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlBrush消费卡(ByVal dblMoney As Double, ByVal rsClassMoney As ADODB.Recordset, _
    ByVal str费用来源 As String, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:处理消费卡刷卡
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-15 09:54:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As Collection
    Dim frmInput As New frmInputPass
    Set mcllSquareBalance = Nothing
    If mobjPatiCardObject Is Nothing Then
        MsgBox "当前支付类别接口有误,请检查", vbOKOnly, gstrSysName
        Exit Function
    End If
    zlBrush消费卡 = frmInput.zlBrushPay(Me, mlngModule, mobjPatiCardObject, rsClassMoney, _
        mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, Nvl(mrsInfo!姓名), Nvl(mrsInfo!性别), _
        Nvl(mrsInfo!年龄), dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, , True, , False, cllBalance, _
        False, True, str费用来源, lng病人ID)
    Set frmInput = Nothing
    Set mcllSquareBalance = cllBalance

End Function
Private Function BrushCardThreeSwapCheck(ByVal strNos As String, _
    ByVal dblMoney As Double, ByVal str费用来源 As String, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡验证
    '入参:strNos -本次支付的单据号
    '       dblMoney-支付的总金额
    '返回:返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim frmMain As Object
    
    On Error GoTo errHandle
    If mintCurType = 2 Then BrushCardThreeSwapCheck = True: Exit Function
    If mCurCardPay.lng医疗卡类别ID = 0 Then BrushCardThreeSwapCheck = True: Exit Function
    If mblnCliniqueRoomPay = False Then
        If IsCheckThreeValied = False Then Exit Function
        Set frmMain = Me
    Else
        Set frmMain = mfrMain
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
    If mCurCardPay.bln消费卡 And mCurCardPay.bln自制卡 Then
        '问题:50339
        If zlGetClassMoney(strNos, rsMoney) = False Then Exit Function
        '肯定是处理自制卡
        If zlBrush消费卡(dblMoney, rsMoney, str费用来源, lng病人ID) = False Then Exit Function
    Else
        '    zlBrushCard(frmMain As Object, _
        '        ByVal lngModule As Long, _
        '        ByVal lngCardTypeID As Long, _
        '        ByVal strPatiName As String, ByVal strSex As String, _
        '        ByVal strOld As String, ByVal dbl金额 As Double, _
        '        Optional ByRef strCardNo As String, _
        '        Optional ByRef strPassWord As String
        If mobjCardPay.zlBrushCard(frmMain, mlngModule, mCurCardPay.lng医疗卡类别ID, _
         Nvl(mrsInfo!姓名), Nvl(mrsInfo!性别), Nvl(mrsInfo!年龄), dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码) = False Then Exit Function
    End If
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If mobjCardPay.zlPaymentCheck(frmMain, mlngModule, mCurCardPay.lng医疗卡类别ID, _
          mCurCardPay.str刷卡卡号, dblMoney, strNos, strXMLExpend) = False Then Exit Function
    BrushCardThreeSwapCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function interfacePayMoney(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:用第三方接口支付(包含消费卡,银行卡及其他卡)
    '入参:strCardNo-支付的卡号
    '返回:支付成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-22 12:01:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim lng病人ID As Long, strCardPass As String, lng卡类别ID As Long
    Dim dbl金额 As Double
    
    On Error GoTo errHandle
    dbl金额 = RoundEx(Val(txt金额), 2)
    '79621:李南春,2014/11/14,对金额格式化处理
    If RoundEx(Val(txt冲预交) + dbl金额, 2) > RoundEx(Val(lbl自付合计.Tag), 2) Then
        MsgBox "输入本次扣款金额大于了本次结算费用合计,不能继续!", vbInformation + vbOKOnly, gstrSysName
        If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        Exit Function
    End If
    If RoundEx(Val(txt冲预交) + dbl金额, 2) <> RoundEx(Val(lbl自付合计.Tag), 2) Then
        MsgBox "输入本次扣款金额小于了本次结算费用合计,不能继续!", vbInformation + vbOKOnly, gstrSysName
        If txt金额.Enabled And txt金额.Visible Then txt金额.SetFocus
        Exit Function
    End If
    If mrsInfo Is Nothing Then
        MsgBox "请先输入病人!", vbInformation + vbOKOnly, gstrSysName
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "请先输入病人!", vbInformation + vbOKOnly, gstrSysName
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
        Exit Function
    End If
    If Val(lbl预存款.Tag) = 0 And Val(txt冲预交.Text) <> 0 Then
        '未密码验证,需要病人确定输入密码
        MsgBox "使用预交款，必须先刷卡确认消费！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    interfacePayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub setDefaultPrepayMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省预交金额
    '编制:刘兴洪
    '日期:2011-08-13 17:21:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurCarge
         txt冲预交.Text = "0.00"
        If .dbl可用预交 <> 0 Then
            txt冲预交.Text = Format(IIf(.dbl可用预交 > .dbl当前未付, .dbl当前未付, .dbl可用预交), "###0.00;###0.00;0.00;0.00")
        End If
    End With
End Sub
Private Function ExecuteCliniqueRoomPay() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:诊间支付
    '返回:诊间支付成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-01-14 17:28:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intTYPE As Integer, objSquareCard As Object
    On Error GoTo errHandle
    
    '创建对象,增加结算卡的结算
    Err = 0: On Error Resume Next
    Set objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        MsgBox "结算卡部件zl9CardSquare.clsCardSquare创建失败！", vbInformation, gstrSysName
        Err = 0: On Error GoTo 0:      Exit Function
    End If
    If objSquareCard Is Nothing Then Exit Function
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Function
    End If
    
    '获取病人信息
    If GetPatient = False Then Exit Function
     
    '保存数据
    If SaveCharge = False Then mstrPrintNO = "": Exit Function
    Call PrintBill
    
    '银医一卡通写卡，85950
    Call WriteInforToCard(Me, mlngModule, mstrPrivs, objSquareCard, 0, mstrPrintNO)
    Set objSquareCard = Nothing
    
    ExecuteCliniqueRoomPay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CreateLocalTypeObject(ByVal lngCardTypeID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建指定卡类别对象
    '入参:lngCardTypeID-卡类别ID
    '出参:
    '返回:创建成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-01-14 18:19:38
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objCard As clsCard, blnReturn As Boolean
    On Error GoTo errHandle
    '创建指定的对象
    Set mobjCardPay = Nothing
    
    blnReturn = zlGetCardProperty(lngCardTypeID, False, objCard)
    If blnReturn = False Or objCard Is Nothing Then
        MsgBox "注意:" & vbCrLf & _
                      "      在医疗卡类别中，未找到指定的三方帐户所支持的类别， " & vbCrLf & _
                      "可能该类别未启用,请检查［医疗卡类别］", vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.是否存在帐户 = False Then
        MsgBox objCard.名称 & "未设置三方帐户,请检查［医疗卡类别］", vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.结算方式 = "" Then
        MsgBox objCard.名称 & "未设置结算方式,请检查［医疗卡类别］", vbInformation, gstrSysName
        Exit Function
    End If
    If objCard.接口程序名 = "" Then
        MsgBox objCard.名称 & "未设置三方接口所支持的部件,请检查［医疗卡类别］", vbInformation, gstrSysName
        Exit Function
    End If
    With mCurCardPay
       .lng医疗卡类别ID = objCard.接口序号
       .bln消费卡 = objCard.消费卡
       .str结算方式 = objCard.结算方式
       .str名称 = objCard.名称
       .str刷卡卡号 = ""
       .str刷卡密码 = ""
       .lngID = 0
       .strNO = ""
       .bln卡号密文 = False
       .int医疗卡长度 = 0
       .bln读卡 = False
       .bln支票 = False
       .blnOneCard = False
       .bln自制卡 = False
       .int性质 = 0
    End With
    Err = 0: On Error Resume Next
    
    If zlGetCardObj(Me, objCard.接口序号, objCard.消费卡, mobjPatiCardObject, , True) = False Then
        Set mobjPatiCardObject = Nothing
        Set mobjCardPay = Nothing
        Exit Function
    End If
    
    Set mobjCardPay = mobjPatiCardObject.CardObject
    If Err <> 0 Then
        MsgBox "未找到" & mCurCardPay.str名称 & "所对应的部件,请检查", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjCardPay Is Nothing Then Exit Function
    CreateLocalTypeObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function CliniqueRoomPayValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:诊间支付检查
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-01-17 16:36:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mblnCliniqueRoomPay = False Then CliniqueRoomPayValied = True: Exit Function
    If mbytBillType <> 1 Then   '只针对收费单
        MsgBox "注意:" & vbCrLf & "    诊间支付时，不允许针对记帐单据性质的进行支付。", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If mlngCardTypeID = 0 Then
        MsgBox "注意:" & vbCrLf & "    诊间支付时要求指定一个三方帐户支付类别,请与系统管理员联系。", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
 
    '对象创建失败的,不允许支付
    If Not CreateLocalTypeObject(mlngCardTypeID) Then Exit Function
    CliniqueRoomPayValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPayDrugWindow(ByVal lng病人ID As Long, ByVal dt收费时间 As Date, _
    ByVal cllDept As Collection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：分配发药窗口
    '入参:lng病人ID-病人ID
    '     dt收费时间-收费时间
    '     cllDept-具体执行部门:array(收费类别,执行部门ID,发药窗口)
    '返回：发药窗口名称
    '编制：李南春
    '入参:strNO
    '时间：2014-6-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发药窗口 As String, strPayDrugWins As String
    Dim str西窗 As String, str中窗 As String, str成窗 As String
    Dim i As Long, varData As Variant
    
    Err = 0: On Error GoTo Errhand:
    strPayDrugWins = ""
    For i = 1 To cllDept.count
        varData = cllDept(i)
        str发药窗口 = varData(2)
        If str发药窗口 = "" Then
            '判断当前病人是否存在相同执行部门的未发药品，若存在则返回未发药品的发药窗口
            str发药窗口 = Get未发药品发药窗口(lng病人ID, Val(varData(1)))
            If str发药窗口 = "" Then str发药窗口 = GetDrugWindow(Val(varData(1)), Trim(varData(0)))
            If str发药窗口 = "" Then
                str发药窗口 = Get发药窗口(dt收费时间, Val(varData(1)), Trim(varData(0)), str西窗, str成窗, str中窗)
            End If
        End If
        If InStr(1, strPayDrugWins & ";", ";" & Val(varData(1)) & "|") = 0 Then
            strPayDrugWins = strPayDrugWins & ";" & Val(varData(1)) & "|" & str发药窗口
        End If
    Next
    GetPayDrugWindow = strPayDrugWins
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function GetDrugWindow(ByVal lng药房ID As Long, ByVal str类别 As String) As String
'功能：获取缺省的发药窗口,如果参数指定了缺省,则以指定为准,否则,如果是划价单,则以第一药品行的窗口为准,否则以已输入相同药品的窗口为准
'参数：intPage=搜录到的单据编号
'说明：主要用于多单据收费时，不同类别的药品可能动态分配到同一药房，这样他们的窗口也应相同，但强行指定的除外
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim p As Integer, i As Integer, varData As Variant, varTemp As Variant
    Err = 0: On Error GoTo errH:
    GetDrugWindow = GetDefaultWindow(str类别, lng药房ID)
    If GetDrugWindow = "" Then Exit Function
    strSQL = "Select 编码 From 发药窗口 Where 上班否=1 And 药房ID=[1] And 名称=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药房ID, GetDrugWindow)
    If rsTmp.EOF Then GetDrugWindow = ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get未发药品发药窗口(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As String
    '-------------------------------------------------------------------------
    '功能：判断当前病人是否存在相同执行部门的未发药品，若存在则返回未发药品的发药窗口
    '返回：若存在相同执行部门的未发药品，则返回未发药品的发药窗口，否则返回空
    '编制：冉俊明
    '日期：2014-04-09
    '问题：71902
    '说明：
    '   同一个人病人不同时间段多张单据收费，分配同一个发药窗口，方便病人取药
    '-------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select 发药窗口" & vbNewLine & _
            "From 未发药品记录" & vbNewLine & _
            "Where 单据 = 8 And 发药窗口 Is Not Null And 病人id = [1] And 库房id = [2]" & vbNewLine & _
            "Order By 已收费 Desc, 填制日期 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人未发药品发药窗口", lng病人ID, lng执行部门ID)
    
    If Not rsTemp.EOF Then
        Get未发药品发药窗口 = Nvl(rsTemp!发药窗口)
    End If
    rsTemp.Close: Set rsTemp = Nothing
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get发药窗口(ByVal Curdate As Date, ByVal lng药房ID As Long, ByVal str类别 As String, _
    str西窗 As String, str成窗 As String, str中窗 As String) As String
'功能：获取药品对应的发药窗口
'参数：lng药房ID=执行部门ID,curDate=当前时间
'说明：在同一材质类药房的发药窗口内平均分配
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '指定时固定分配(指定是指没有对应药房上班时指定)
    Select Case str类别
        Case "5"
            If str西窗 <> "" Then
                Get发药窗口 = str西窗
            ElseIf mlng西药房 > 0 Then
                Get发药窗口 = GetDefaultWindow(str类别, lng药房ID)
                str西窗 = Get发药窗口
            End If
        Case "6"
            If str成窗 <> "" Then
                Get发药窗口 = str成窗
            ElseIf mlng成药房 > 0 Then
                Get发药窗口 = GetDefaultWindow(str类别, lng药房ID)
                str成窗 = Get发药窗口
            End If
        Case "7"
            If str中窗 <> "" Then
                Get发药窗口 = str中窗
            ElseIf mlng中药房 > 0 Then
                Get发药窗口 = GetDefaultWindow(str类别, lng药房ID)
                str中窗 = Get发药窗口
            End If
    End Select
    
    
    If Get发药窗口 <> "" Then
        strSQL = "Select 编码 From 发药窗口 Where 上班否=1 And 药房ID=[1] And 名称=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng药房ID, Get发药窗口)
        If rsTmp.EOF Then Get发药窗口 = ""
        Exit Function
    End If
    
    '动态分配上班的非专家窗口,98876
    strSQL = "Select Zl_Get发药窗口([1],[2],[3]) As 窗口 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取发药窗口", lng药房ID, mbytAssign, Curdate)
    If Not rsTmp.EOF Then
        Get发药窗口 = Nvl(rsTmp!窗口)
    End If
    
    If Get发药窗口 <> "" Then
        Select Case str类别
            Case "5"
                str西窗 = Get发药窗口
            Case "6"
                str成窗 = Get发药窗口
            Case "7"
                str中窗 = Get发药窗口
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDefaultWindow(ByVal str类别 As String, ByVal lng药房ID As Long) As String
'功能:获取缺省的药房窗口设置
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    
    Select Case str类别
        Case "5"
            If InStr(mstr西窗, ":") > 0 Then '旧数据没有存药房ID
                 strTmp = mstr西窗
            ElseIf mlng西药房 > 0 And mstr西窗 <> "" Then
                strTmp = mlng西药房 & ":" & mstr西窗
            End If
        Case "6"
            If InStr(mstr成窗, ":") > 0 Then
                 strTmp = mstr成窗
            ElseIf mlng成药房 > 0 And mstr成窗 <> "" Then
                 strTmp = mlng成药房 & ":" & mstr成窗
            End If
        Case "7"
            If InStr(mstr中窗, ":") > 0 Then
                 strTmp = mstr中窗
            ElseIf mlng中药房 > 0 And mstr中窗 <> "" Then
                 strTmp = mlng中药房 & ":" & mstr中窗
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str类别
                Case "5"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    GetDefaultWindow = strTmp
End Function
Private Function SetCurBalanceSQL(ByVal bytType As Byte, ByVal lng病人ID As Long, _
    ByVal dblPayMoney As Double, ByVal dbl冲预交 As Double, ByVal dbl缴款 As Double, ByVal dbl找补 As Double, _
    ByVal dbl本次误差费 As Double, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置当前结算的SQL给cllpro过程
    '入参:  bytType-1-三方接口支付;2-消费卡支付;0-其他
    '       dblPayMoney-当前支付金额
    '       dbl冲预交-预交款支付
    '       dbl缴款-投币有效
    '       dbl找补-投币有效
    '       dbl预存款-如果是用预交款的话,传入预交款金额
    '       dbl本次误差费-本次产生的误差费
    '出参:cllPro-执行过程
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-15 15:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL  As String, str收费结算 As String
    Dim dbl预存款 As Double, lngCardTypeID As Long, j As Long
    
    
    On Error GoTo errHandle
    
    ' Zl_门诊收费结算_Modify
    strSQL = "Zl_门诊收费结算_Modify("
    '  ------------------------------------------------------------------------------------------------------------------------------
    '  --功能:收费结算时,修改结算的相关信息
    '  --操作类型_In:
    '  --   0-普通收费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退支票额_In:传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退支票额_In:传入零
    '  --   3-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位
    '  --     ②冲预交_In: 传入零
    '  --     ②退支票额_In:传入零
    '  -- 冲预交_In: 存在冲预交时,传入
    '  -- 误差金额_In:存在误差费时,传入
    '  -- 完成结算_In:1-完成收费;0-未完成收费
    '  ------------------------------------------------------------------------------------------------------------------------------
    ' bytType- 1-三方接口支付;2-消费卡支付,3帐户支付
    Select Case bytType
    Case 1  '1-三方接口支付
        strSQL = strSQL & "1" & ","
        '"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        str收费结算 = mCurCardPay.str结算方式
        str收费结算 = str收费结算 & "|" & dblPayMoney
        str收费结算 = str收费结算 & "|" & " "
        str收费结算 = str收费结算 & "|" & " "
        lngCardTypeID = mCurCardPay.lng医疗卡类别ID
    Case 2 ' 2-消费卡支付
        strSQL = strSQL & "3" & ","
        If mcllSquareBalance Is Nothing Then Exit Function
        If mcllSquareBalance.count = 0 Then Exit Function
        '卡类别ID|卡号|消费卡ID|消费金额||."
        '消费卡ID可以不传,传为0时,以卡号自动查找
        str收费结算 = ""
        For j = 1 To mcllSquareBalance.count
            ' array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
            str收费结算 = str收费结算 & "||" & Val(mcllSquareBalance(j)(0))
            str收费结算 = str收费结算 & "|" & mcllSquareBalance(j)(3)
            str收费结算 = str收费结算 & "|" & Val(mcllSquareBalance(j)(1))
            str收费结算 = str收费结算 & "|" & Val(mcllSquareBalance(j)(2))
        Next
        If str收费结算 <> "" Then str收费结算 = Mid(str收费结算, 3)
        lngCardTypeID = mCurCardPay.lng医疗卡类别ID
    Case Else
        strSQL = strSQL & "0" & ","
    End Select
    '    病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & lng病人ID & ","
    '    结帐id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & mCurCardPay.lng结帐ID & ","
    '    结算方式_In   Varchar2,
        strSQL = strSQL & IIf(str收费结算 = "", "NULL", "'" & str收费结算 & "'") & ","
    '    冲预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & IIf(dbl冲预交 <> 0, dbl冲预交, "NULL") & ","
    '    退支票额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(lngCardTypeID = 0, "NULL", lngCardTypeID) & ","
    '    卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.str刷卡卡号 <> "", "'" & mCurCardPay.str刷卡卡号 & "'", "NULL") & ","
    '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '    交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '    缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & dbl缴款 & ","
    '    找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & dbl找补 & ","
    '    误差金额_In   门诊费用记录.实收金额%Type := Null,
    '    -- 误差金额_In:存在误差费时,传入
    strSQL = strSQL & "" & dbl本次误差费 & ","
    '    完成结算_In Number:=0
    '    -- 完成结算_In:1-完成收费;0-未完成收费
    strSQL = strSQL & "1,"
    '  缺省结算方式_In  结算方式.名称%Type := Null,
    strSQL = strSQL & "NULL,"
    '79868,冉俊明,2015-06-10,使用病人家属预交
    '  冲预交病人ids_In Varchar2:=Null
    strSQL = strSQL & "'" & lng病人ID & "," & mstr家属IDs & "')"
    zlAddArray cllPro, strSQL
    SetCurBalanceSQL = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示病人信息以及消费情况
    '编制:李南春
    '日期:2014-10-29
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strInfo As String, lngPatient As Long
    If gblnLED = False Then Exit Sub
    
    On Error GoTo Errhand
    zl9LedVoice.Reset mscCom
    strInfo = lbl姓名.Caption
    If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!性别 & " " & mrsInfo!年龄: lngPatient = Val("" & mrsInfo!病人ID)
    zl9LedVoice.DisplayPatient strInfo, lngPatient
    '消费总额:本次需要支付的金额，预交余额:病人当前的预交余额
    Call zl9LedVoice.DisplayBank( _
            "消费总额:" & mCurCarge.dbl本次消费合计 & "元" & _
            IIf(mCurCarge.dbl预交余额 = 0, "", ",预交余额:" & mCurCarge.dbl预交余额 & "元"))
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CreateDrugPacker()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建自助发药机(自动化药房)
    '编制:刘兴洪
    '日期:2014-06-05 15:30:47
    '说明:bug-51510
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnDrugPacker = False: mblnDrugMachine = False
    
    '0-不区分收费或记帐单,1-收费记录;2-记帐记录
'    If mbytBillType = 2 Then Exit Sub
    
    If mblnDrugMachine Or mblnDrugPacker Then Exit Sub

    Err = 0: On Error Resume Next
    If Val(zlDatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        '优先新接口
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    If mblnDrugMachine = False Then
        '旧部件
        Err = 0
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err = 0 Then mblnDrugPacker = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        '权限检查
        strPrivs = GetPrivFunc(glngSys, Val("9010-药品自动化设备接口"))
        If InStr(";" & strPrivs & ";", ";基本;") >= 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, objComLib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    ElseIf mblnDrugPacker Then
        mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
    End If
End Sub
