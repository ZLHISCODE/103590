VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmSendCardAndDeposit 
   BorderStyle     =   0  'None
   Caption         =   "预交及发卡"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   15030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.TabStrip tbDeposit 
      Height          =   405
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   714
      Style           =   2
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   882
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "门诊预交(&M)"
            Key             =   "K1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "住院预交(&Z)"
            Key             =   "K2"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fra磁卡 
      Caption         =   "【发卡信息】"
      ForeColor       =   &H00C00000&
      Height          =   1305
      Left            =   45
      TabIndex        =   30
      Top             =   1500
      Width           =   14970
      Begin VB.TextBox txt卡额 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "卡费"
         Top             =   840
         Width           =   1485
      End
      Begin VB.CheckBox chk记帐 
         Caption         =   "记帐"
         Height          =   360
         Left            =   3600
         TabIndex        =   25
         Top             =   840
         Width           =   788
      End
      Begin VB.TextBox txt卡号 
         BackColor       =   &H00EBFFFF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1170
         PasswordChar    =   "*"
         TabIndex        =   18
         Tag             =   "卡号"
         Top             =   405
         Width           =   2625
      End
      Begin VB.TextBox txtPass 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5520
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   20
         Tag             =   "密码"
         Top             =   405
         Width           =   1750
      End
      Begin VB.ComboBox cbo发卡结算 
         Height          =   360
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   840
         Width           =   1750
      End
      Begin VB.TextBox txtAudi 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7860
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   22
         Tag             =   "验证"
         Top             =   405
         Width           =   1750
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   11625
         TabIndex        =   24
         Top             =   405
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483633
         CalendarTitleBackColor=   16744576
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   53608451
         CurrentDate     =   43424
      End
      Begin VB.CheckBox chkEndTime 
         Caption         =   "终止使用时间"
         Height          =   240
         Left            =   9855
         TabIndex        =   23
         Top             =   465
         Width           =   1755
      End
      Begin MSComctlLib.TabStrip tbSendCard 
         Height          =   315
         Left            =   75
         TabIndex        =   17
         Top             =   0
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         Style           =   2
         TabFixedHeight  =   526
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         TabMinWidth     =   882
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "发卡收费(&1)"
               Key             =   "CardFee"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "绑定卡号(&2)"
               Key             =   "CardBind"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   540
         TabIndex        =   35
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lbl卡号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "卡号"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   510
         TabIndex        =   34
         Top             =   450
         Width           =   510
      End
      Begin VB.Label lbl密码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         Height          =   240
         Left            =   4995
         TabIndex        =   19
         Top             =   465
         Width           =   480
      End
      Begin VB.Label lbl验证 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证"
         Height          =   240
         Left            =   7335
         TabIndex        =   21
         Top             =   465
         Width           =   480
      End
      Begin VB.Label lbl卡名称 
         AutoSize        =   -1  'True
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3480
         TabIndex        =   31
         Top             =   15
         Width           =   120
      End
      Begin VB.Label lbl结算方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算方式"
         Height          =   240
         Left            =   4515
         TabIndex        =   26
         Top             =   900
         Width           =   960
      End
   End
   Begin VB.Frame fra预交 
      Caption         =   "【住院预交信息】"
      ForeColor       =   &H00C00000&
      Height          =   1200
      Left            =   30
      TabIndex        =   28
      Top             =   105
      Width           =   14955
      Begin VB.TextBox txt开户行 
         Height          =   360
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   14
         Top             =   735
         Width           =   2805
      End
      Begin VB.TextBox txtFact 
         Height          =   360
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   3
         Top             =   345
         Width           =   1470
      End
      Begin VB.TextBox txt结算号码 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9480
         MaxLength       =   30
         TabIndex        =   9
         Top             =   345
         Width           =   2445
      End
      Begin VB.ComboBox cbo预交结算 
         Height          =   360
         Left            =   6345
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   345
         Width           =   1770
      End
      Begin VB.TextBox txt预交额 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EBFFFF&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   3525
         MaxLength       =   12
         TabIndex        =   5
         Top             =   345
         Width           =   1335
      End
      Begin VB.CheckBox chk单位缴款 
         Caption         =   "单位缴款"
         Height          =   360
         Left            =   13050
         TabIndex        =   10
         Top             =   345
         Width           =   1320
      End
      Begin VB.TextBox txt缴款单位 
         Height          =   360
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   12
         Top             =   735
         Width           =   2745
      End
      Begin VB.TextBox txt帐号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9480
         MaxLength       =   50
         TabIndex        =   16
         Top             =   735
         Width           =   4800
      End
      Begin VB.Label lblAccno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "帐号"
         Height          =   240
         Left            =   8880
         TabIndex        =   15
         Top             =   795
         Width           =   480
      End
      Begin VB.Label lblBank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开户行"
         Height          =   240
         Left            =   4440
         TabIndex        =   13
         Top             =   795
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缴款单位"
         Height          =   240
         Left            =   210
         TabIndex        =   11
         Top             =   795
         Width           =   960
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实际票号"
         Height          =   240
         Left            =   210
         TabIndex        =   2
         Top             =   405
         Width           =   960
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缴款方式"
         Height          =   240
         Left            =   5310
         TabIndex        =   6
         Top             =   405
         Width           =   960
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算号码"
         Height          =   240
         Left            =   8400
         TabIndex        =   8
         Top             =   405
         Width           =   960
      End
      Begin VB.Label lblDepositMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2955
         TabIndex        =   4
         Top             =   405
         Width           =   480
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   825
         TabIndex        =   29
         Top             =   1605
         Width           =   480
      End
      Begin VB.Label lblYBMoney 
         AutoSize        =   -1  'True
         Caption         =   "个人帐户余额:"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3585
         TabIndex        =   1
         Top             =   15
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin zlIDKind.ucQRCodePayButton btQRCodeTemp 
      Height          =   315
      Left            =   14640
      TabIndex        =   32
      Top             =   1305
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
   End
End
Attribute VB_Name = "frmSendCardAndDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************************************************************
'发卡及预交公共窗体
'接口:
'    1.zlInit:初始化接口
'    2.zlRecalcCardFee-新计算卡费用：费别及医疗卡变动时，需要调用
'    3.zlSaveDataBeforCheckIsValid-检查数据的合法性:在保存前需要调用
'    4.zlSaveData-执行数据保存操作
'    5.zlSaveDataAfter-数据保存成功后才调用
'公共事件:
'    1.RequestRefreshPatiInf-请求重新根据XML格式的输出内容，刷新病人信息
'    2.InputOver-输入完成事件(表示最后一项输入完成，以便光标跳转到下一项输入内容
'公共方法
'    1.zlGetSendCard-获取当前的发卡卡对象
'    2.zlSetCardNo-给卡号重新附值
'    3.zlSetUnitInfo-设置缴款单位信息(工作单位，账号等数据完成后，需要设置)
'    4.zlSetInsureInfo:设置医保信息
'    5. zlClearControlInfo-当前当前窗体的所有信息
'    6. zlSetFocus光标移动
'公共属性
'    1.RealName-设置当前病人是否进行了实名认证(实名认证后，需要赋值)
'    2.GetWidth-窗体宽度
'    3.GetWidth-窗体高度
'返回:成功返回true,否则返回False
'编制:刘兴洪
'日期:2019-11-25 14:32:57
'*********************************************************************************************************************************************
'-------------------------------------------------------------------------------------------------
'接口变量
Private mint操作状态 As Integer '0-增加;1-异常重收;2-异常作废
Private WithEvents mbtQRCodePay As ucQRCodePayButton
Attribute mbtQRCodePay.VB_VarHelpID = -1
Private mfrmMain As Object
Private mlngModule As Long
Private mbln门诊预交 As Boolean, mbln住院预交 As Boolean
Private mlngCardTypeID As Long '发卡的卡类别ID
Attribute mlngCardTypeID.VB_VarHelpID = -1
Private mblnView As Boolean
Private mblnAllowSendCard As Boolean, mblnAllowBoundCard As Boolean
Private mblnCancel As Boolean '是否作废
Private mlng异常ID As Long '异常操作
Private mbyt应用场景   As Byte    '1-医疗卡发卡;2-病人信息登记;3-病人入院 登记;4-预约挂号接收
Private mblnShowDepositAndSendCard As Boolean '不管存不存在预交及发卡属性，都应该显示在界面上，主要是配合界面的显示
'------------------------------------------------------------------------------------------------
'内部变量
Private mblnNotClick As Boolean
Private mobjOneCardComLib As clsOneCardComLib
Private mblnInited As Boolean '是否初始化成功的,只有初始化成功的，才允许编辑

Private mobjPubPatient As clsInterFacePatient   '病人公共部件接口
Private mobjService As clsService '公共域部件：药品、卫材及临床及病人域
Private mobjExseSvr As clsExpenceSvr
Private mobjPati As clsPatientInfo
Private mobjThirdSwap As clsThirdSwapCard   '三方交易的相关接口
Private mblnICCard As Boolean
Private mdblRQCodeMoney As Double '扫码付支付金额
Private mbln相同结算 As Boolean '预交和卡费为同一种结算方式

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mstr单位帐号  As String
Private mstr缴款单位   As String
Private mstr单位开户行 As String
Private mstrQRcode As String '当前扫码付的二维码

'-------------------------------------------------------------------------------------------------
'参数变量
Private mbln记账 As Boolean '就诊卡费用以记账方式收取
Private mbytRegValidDays As Integer '挂号有效天数，主要用于实名制体系发卡缺省有效时间
Private mblnNewPatiMustSendCard As Boolean  '建档同时必须发卡

'-------------------------------------------------------------------------------------------------
'卡费相关
Private Type Ty_CardProperty
       objSendCard As Card  '发卡对象
       lng领用ID As Long
       lng共用批次 As Long
       bln变价 As Boolean
       blnOneCard As Boolean '  '是否启用了一卡通接口,此模式下，票号严格管理，票号范围外的发卡或绑定卡不收费
       rs卡费 As ADODB.Recordset
       dbl应收金额 As Double
       dbl实收金额 As Double
End Type
Private mCurSendCard As Ty_CardProperty
Private mblnSendCardLocked As Boolean '是否锁定发卡
Private mrs卡费 As ADODB.Recordset
Private mintPriceGradeStartType As Integer   '启用价格等级类型:'   0-未启用,1-只启用了站点,2-只启用了医疗付款方式,3-站点和医疗款方式都启用了
Private mstrPriceGrade As String, mstrPrePriceGrade As String

Private mobjShowTotalMoneyControl As Object '显示扫码付总额的控件
Private mobjCardFeePayCards As Cards  '卡费支付方式
Private mstrQRCodeTypeIds_CardFee As String '卡费二维码扫码付
Private mobjCardFeeItems As clsBalanceItems '卡费结算信息
Private mblnBoundCarded As Boolean '是否已经绑定卡
Private mrsCardFee As ADODB.Recordset  '卡费记录集

'-------------------------------------------------------------------------------------------------
'预交票据及打印相关
Private mobjDepositFact As clsFactProperty '预交发票数据
Private mblnDepositStrictly As Boolean '预交是否严格控制
Private mbyt预交票据长度 As Byte   '预交票据长度
Private mblnDepositPrint As Boolean '是否打印
Private mobjDepositPayCards As Cards  '预交支付方式
Private mbytPrepayType As Byte '上次预交类型: 0-门诊住院;1-门诊;2-住院
Private mblnAllowInsureAccDeposit As Boolean  '是否允许医保病人缴预交
Private mstrQRCodeTypeIds_Deposit As String '预交款二维码扫码付
Private mblnDepositLocked As Boolean '预交款部分锁定
Private mobjDepositItems As clsBalanceItems  '预交当前支付信息

'-------------------------------------------------------------------------------------------------
'医保相关变量
Private mcurYBMoney As Currency  '医保个人账户余额
Private mintInsure As Integer  '医保险类
Private mstr医保号 As String    '医保号
Private mstr密码 As String   '医保密码

'-------------------------------------------------------------------------------------------------
'密码键盘相关
Private mobjKeyboard As Object
'-------------------------------------------------------------------------------------------------
'公共事件
Public Event RequestRefreshPatiInf(ByVal strCardNo As String, ByVal strPatiInfoXML As String)
Public Event InputOver()    '输入完成
Public Event ExcuteQRCodePayment() '执行扫码付
Public Event Activate() '子窗体激活
Public Event ExcuteReadQRCode() '扫码读卡
Public Event ControlGotFocus(objControl As Object)
'Public Event ControlLostFocus(objControl As Object)

'-------------------------------------------------------------------------------------------------
'属性变量
Private mblnRealName As Boolean '是否实名认证
 
Public Sub zlSetFocus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:光标定位
    '入参
    '编制:刘兴洪
    '日期:2020-01-13 17:53:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.Enabled And Me.Visible Then Me.SetFocus
    If fra预交.Visible Then
        If txt预交额.Visible And txt预交额.Enabled Then txt预交额.SetFocus
    ElseIf fra磁卡.Visible Then
        If txt卡号.Enabled And txt卡号.Visible Then
            txt卡号.SetFocus
        End If
    End If
End Sub

Public Function zlInit(ByVal frmMain As Object, ByVal lngModule As Long, ByVal bln门诊预交 As Boolean, ByVal bln住院预交 As Boolean, _
    ByVal lngCardTypeID As Long, blnAllowSendCard As Boolean, ByVal blnAllowBoundCard As Boolean, ByVal blnAllowInsureAccDeposit As Boolean, _
    Optional btQRCodePay As Object, Optional objShowTotalMoneyControl As Object, Optional blnView As Boolean = False, _
    Optional objOneCardComLib As Object, Optional ByVal blnCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口
    '入参:frmMain-调用的主窗体
    '     lngModule-模块号
    '     btQRCodePay-扫码付按钮
    '     objShowTotalMoneyControl-显示的总额控件:lable或Text
    '     bln门诊预交-是否缴门诊预交
    '     bln住院预交-是否缴住院预交
    '     lngSendCardTypeID-当前发卡类别ID:传入0时，则参数：blnAllowSendCard及blnAllowBoundCard-无效
    '     blnAllowSendCard-允许发卡
    '     blnAllowBoundCard-允许绑定卡
    '     objOneCardComLib-一卡通公共部件,nothing时，将重新创建一个
    '     blnView-是否查看
    '     strPrivs-当前操作模块权限
    '     blnAllowInsureAccDeposit-是否允许医保账户缴预交
    '     blnCancel-当前是否作废操作
    '出参:
    '返回:初始化成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-23 14:18:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubPatient As clsInterFacePatient
    
    On Error GoTo errHandle
    mblnInited = False
    Set mfrmMain = frmMain: mlngModule = lngModule: mbln门诊预交 = bln门诊预交: mbln住院预交 = bln住院预交
    mblnAllowSendCard = blnAllowSendCard: mblnAllowBoundCard = blnAllowBoundCard
    mblnAllowInsureAccDeposit = blnAllowInsureAccDeposit: mblnCancel = blnCancel
    mlngCardTypeID = lngCardTypeID
    If objOneCardComLib Is Nothing Then
        Set mobjOneCardComLib = New clsOneCardComLib
        If mobjOneCardComLib.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle) = False Then Exit Function
    Else
        Set mobjOneCardComLib = objOneCardComLib
    End If
    Set mobjExseSvr = New clsExpenceSvr
    Call mobjExseSvr.zlInitCommon(glngSys, mlngModule, gcnOracle, gstrDBUser)
    
    Set mobjService = New clsService
    Call mobjService.zlInitCommon(glngSys, mlngModule, gcnOracle, gstrDBUser)
    Set mbtQRCodePay = btQRCodePay
    Call CreateObjectKeyboard
    
    Set mobjThirdSwap = New clsThirdSwapCard '初始化三方交易
    Call mobjThirdSwap.zlInitCompents(Me, mlngModule, mobjOneCardComLib)
    
    If GetPublicPatient(objPubPatient) = False Then Exit Function
    
    Set mobjShowTotalMoneyControl = objShowTotalMoneyControl
    mblnShowDepositAndSendCard = False
    mblnView = blnView
    
   
    If blnCancel Then zlInit = True: mblnInited = True: Exit Function
    
    zlInit = InitFace
     mblnInited = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlRecalcCardFee(ByVal objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算卡费信息
    '入参:
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-23 17:44:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs卡费 As ADODB.Recordset, blnReReadCardFee As Boolean
    
    Set mobjPati = objPati
    If fra磁卡.Visible Then zlRecalcCardFee = True: Exit Function
    
    ' mintPriceGradeStartType As Integer   '启用价格等级类型:'   0-未启用,1-只启用了站点,2-只启用了医疗付款方式,3-站点和医疗款方式都启用了
    If mintPriceGradeStartType >= 2 Then
       Call GetPriceGrade(gstrNodeNo, 0, 0, mobjPati.医疗付款方式, , , mstrPriceGrade)
        '重取价格等级
        If mstrPriceGrade <> mstrPrePriceGrade Then
            '要重新获取价格等级
            Set mrs卡费 = Nothing: blnReReadCardFee = True
                mstrPrePriceGrade = mstrPriceGrade
        End If
    End If
    Set rs卡费 = GetCardFee(blnReReadCardFee, mstrPriceGrade)
    Call InitCardFee '加载卡费数据
    
    Call ReLoadCardFee  '重新计算
End Function
Public Function zlSetCardNo(ByVal strCardNo As String, objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置卡号给卡号文本框
    '入参:objPati-病人信息集
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 18:53:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    txt卡号.Text = strCardNo

    Call zlRecalcCardFee(objPati)
    zlSetCardNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlSetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(9(小五))，1-大(缺省：12(小四)）
    '编制:刘兴洪
    '日期:2014-04-09 11:46:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytFontSize As Byte
    Dim objControl As Control
    On Error GoTo errHandle
    
    bytFontSize = IIf(bytSize = 0, 9, 12)
    Me.Font.Size = bytFontSize
    
    For Each objControl In Me.Controls
        If UCase(TypeName(objControl)) <> UCase("ucQRCodePayButton") Then
            objControl.Font.Size = bytFontSize
            'Debug.Print TypeName(objControl)
            If UCase(TypeName(objControl)) = UCase("TextBox") Or UCase(TypeName(objControl)) = UCase("DTPicker") Then
                 objControl.Height = IIf(bytSize = 0, 300, 360)
            End If
        End If
    Next
    Me.Refresh
    
    fra磁卡.Left = fra预交.Left
    fra预交.Top = Me.ScaleTop + 105
    If mbln门诊预交 Or mbln住院预交 Then
        fra磁卡.Top = fra预交.Top + fra预交.Height + 50
        Me.Height = Me.ScaleTop + fra预交.Height + IIf(mlngCardTypeID <> 0, fra磁卡.Height, 0) + 250
    Else
        fra磁卡.Top = fra预交.Top
        Me.Height = Me.ScaleTop + fra磁卡.Height + 200
    End If
    
    '位置调整
 
    txt帐号.Width = 4800 * (bytFontSize / 12)
    txtFact.Width = 1470 * (bytFontSize / 12)
    txt预交额.Width = 1335 * (bytFontSize / 12)
    txt结算号码.Width = 2445 * (bytFontSize / 12)
    txt开户行.Width = 2805 * (bytFontSize / 12)
    'txt缴款单位.Width = 2745 * (bytFontSize / 12)
    txt卡号.Width = 2625 * (bytFontSize / 12)
    txtPass.Width = 1750 * (bytFontSize / 12)
    txtAudi.Width = 1750 * (bytFontSize / 12)
    chkEndTime.Width = IIf(bytSize = 0, 1415, 1755)
  
    dtpDate.Width = IIf(bytSize = 0, 2100, 2625)

  
    lblDepositMoney.Left = txtFact.Left + txtFact.Width + 100
    txt预交额.Left = lblDepositMoney.Left + lblDepositMoney.Width + 20

    txtFact.Left = lblFact.Left + lblFact.Width + 20
    txt缴款单位.Left = txtFact.Left
    txt卡号.Left = txtFact.Left
    txt卡额.Left = txtFact.Left

    txt帐号.Left = txt结算号码.Left
    
    txt预交额.Top = txtFact.Top
    cbo预交结算.Top = txtFact.Top
    txt结算号码.Top = txtFact.Top
    chk单位缴款.Top = txt结算号码.Top + (txt结算号码.Height - chk单位缴款.Height) \ 2
    lblCode.Top = txt结算号码.Top + (txt结算号码.Height - lblCode.Height) \ 2
    lblStyle.Top = lblCode.Top
    lblDepositMoney.Top = lblCode.Top
    lblFact.Top = lblCode.Top
    
    
    txt缴款单位.Top = txtFact.Top + txtFact.Height + 50
    txt开户行.Top = txt缴款单位.Top
    txt帐号.Top = txt缴款单位.Top
    
    lblAccno.Top = txt帐号.Top + (txt帐号.Height - lblAccno.Height) \ 2
    lblBank.Top = lblAccno.Top
    lblUnit.Top = lblAccno.Top
    
    'lblFact.Left = txtFact.Left - lblFact.Width - 20
    lblUnit.Left = lblFact.Left  'txt缴款单位.Left - lblUnit.Width - 20
    
    lblDepositMoney.Left = txt预交额.Left - lblDepositMoney.Width - 10
    lblStyle.Left = cbo预交结算.Left - lblStyle.Width - 20
    lblAccno.Left = txt帐号.Left - lblAccno.Width - 20
    lblCode.Left = txt结算号码.Left - lblCode.Width - 20
    
    txt开户行.Left = lblBank.Left + lblBank.Width + 20
    cbo预交结算.Left = txt开户行.Left + txt开户行.Width - cbo预交结算.Width
    lblStyle.Left = cbo预交结算.Left - lblStyle.Width - 20
    
    txtPass.Top = txt卡号.Top
    txtAudi.Top = txt卡号.Top
    dtpDate.Top = txt卡号.Top
    
    
    txt卡额.Top = txt卡号.Top + txt卡号.Height + 50
    lbl金额.Top = txt卡额.Top + (txt卡额.Height - lblAccno.Height) \ 2
    lbl金额.Left = lbl卡号.Left
       
       
    cbo发卡结算.Top = txt卡额.Top
    cbo发卡结算.Left = txtPass.Left
    
    lbl结算方式.Top = cbo发卡结算.Top + (cbo发卡结算.Height - lbl结算方式.Height) \ 2
    chk记帐.Top = cbo发卡结算.Top + (cbo发卡结算.Height - chk记帐.Height) \ 2
    
   ' lbl卡号.Left = txt卡号.Left - lbl卡号.Width - 20
    lbl密码.Left = txtPass.Left - lbl密码.Width - 20
    lbl验证.Left = txtAudi.Left - lbl验证.Width - 20
    
    lbl卡号.Top = txt卡号.Top + (txt卡号.Height - lbl卡号.Height) \ 2
    lbl密码.Top = lbl卡号.Top
    lbl验证.Top = lbl卡号.Top
    Call Form_Resize
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

 
Public Sub zlSetUnitInfo(ByVal str单位帐号 As String, ByVal str缴款单位 As String, ByVal str单位开户行 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置单位账号
    '编制:刘兴洪
    '日期:2019-11-26 13:37:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstr单位帐号 = str单位帐号: mstr缴款单位 = str缴款单位: mstr单位开户行 = str单位开户行
End Sub
Public Sub zlSetInsueInfo(ByVal int险类 As Integer, ByVal cur账户余额 As Currency, ByVal str医保号 As String, ByVal str密码 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置医保信息
    '入参:int险类
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-27 20:07:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPay As Card
    mintInsure = int险类: mcurYBMoney = cur账户余额: mstr医保号 = str医保号: mstr密码 = str密码
    lblYBMoney.Caption = "个人帐户余额：" & Format(mcurYBMoney, "0.00")
    lblYBMoney.Visible = True And int险类 <> 0
    Set objPay = GetDepositPayCard
    
    mblnNotClick = True
    Call Load预交结算方式
    Call SetLoaclePayModefromCard(objPay, True)
    mblnNotClick = False
End Sub

Public Sub zlClearControlInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除控件信息
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-26 13:53:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mobjPati = Nothing
     mstr单位帐号 = "": mstr缴款单位 = "": mstr单位开户行 = ""
     txtAudi.Text = "": txtPass.Text = ""
     txt卡额.Text = "": txt卡号.Text = ""
     txtFact.Text = "": txt预交额.Text = "": txt结算号码.Text = "": txt缴款单位.Text = ""
     txt开户行.Text = "": txt帐号.Text = ""
     lblYBMoney.Caption = "个人帐户余额:"
     chk记帐.value = IIf(mbln记账 = True, 1, 0)
     mintInsure = 0: mcurYBMoney = 0: mstr医保号 = "": mstr密码 = ""
     lblYBMoney.Visible = False
    If cbo预交结算.ListCount > 0 Then cbo预交结算.ListIndex = Val(cbo预交结算.Tag)
    If cbo发卡结算.ListCount > 0 Then cbo发卡结算.ListIndex = Val(cbo发卡结算.Tag)
    Set mobjDepositItems = Nothing
    Set mobjCardFeeItems = Nothing
    Call RefreshFactNo
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Function zlGetSendCard(ByRef objSendCard_Out As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前的发卡对象
    '入参:
    '出参:objSendCard_Out-返回当前发卡的对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 15:13:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mCurSendCard.objSendCard Is Nothing Then Exit Function
    Set objSendCard_Out = mCurSendCard.objSendCard
    zlGetSendCard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshFactNo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取不同类别的发票
    '编制:刘兴洪
    '日期:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If mobjDepositFact Is Nothing Then Set mobjDepositFact = New clsFactProperty
    
    If mobjDepositFact.打印方式 = 0 Then txtFact.Text = "": Exit Sub
    If mblnDepositStrictly = False Then
        '松散：取下一个号码
        txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngModule, "")))
        Exit Sub
    End If
    '严格:     取下一个号码
    mobjDepositFact.领用ID = mobjExseSvr.CheckUsedBill(2, IIf(mobjDepositFact.领用ID > 0, mobjDepositFact.领用ID, mobjDepositFact.LastUseID), , Val(Mid(tbDeposit.SelectedItem.Key, 2)))
    If mobjDepositFact.领用ID <= 0 Then
        Select Case mobjDepositFact.领用ID
            Case 0 '操作失败
'            Case -1
'                MsgBox "你没有自用或共用的预交票据,登记病人信息时不能同时缴预交款！" & _
'                    "请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
'            Case -2
'                MsgBox "本地的共用票据已经用完,登记病人信息时不能同时缴预交款！" & _
'                    "请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
        End Select
        txtFact.Text = ""
    Else
        txtFact.Text = mobjExseSvr.GetNextBill(mobjDepositFact.领用ID)
    End If
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数
    '入参:
    '编制:刘兴洪
    '日期:2019-11-25 14:57:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    
    '票据号码长度、就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    mbyt预交票据长度 = Val(Split(strValue, "|")(1))
    

    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    mblnDepositStrictly = Mid(strValue, 2, 1) = "1" '预交严格控制

       
    strValue = zlDatabase.GetPara(21, glngSys, , "01") & "1"
    mbytRegValidDays = Val(Left(strValue, 1))
    If mbytRegValidDays < Val(Mid(strValue, 2, 1)) Then mbytRegValidDays = Val(Mid(strValue, 2, 1))
    
    mbytPrepayType = Val(zlDatabase.GetPara("上次预交类型", glngSys, mlngModule, "0"))
    
    mbln记账 = zlDatabase.GetPara("卡费记帐", glngSys, mlngModule) = "1"
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
 

Private Function InitFace() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化
    '入参:
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-23 14:21:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln院外发卡 As Boolean, blnBoundCard As Boolean
    Dim blnAllowSendCard As Boolean, blnAllowBoundCard As Boolean
    Dim objPubPatient As clsInterFacePatient
    Dim str医疗付款方式 As String, str批次 As String
    Dim objSendCard As Card
    Dim varData As Variant, i As Long, varTemp As Variant
    
    
    On Error GoTo errHandle
    lblYBMoney.Visible = False
    
    Call InitPara '初始化参数值
    Call zlClearControlInfo
    
    If Load支付方式(mbln门诊预交 Or mbln住院预交, mlngCardTypeID > 0) = False Then Exit Function
    
    If Not mobjPati Is Nothing Then str医疗付款方式 = mobjPati.医疗付款方式
    
    mblnSendCardLocked = False
    
    fra预交.Tag = ""
    fra预交.Visible = mbln门诊预交 Or mbln住院预交 Or mblnShowDepositAndSendCard
    fra预交.Tag = IIf(mbln门诊预交 Or mbln住院预交, "", "1")
    tbDeposit.Visible = mbln门诊预交 And mbln住院预交
    mblnNotClick = True
    If mbln门诊预交 And Not mbln住院预交 Then
        '只有门诊预交
        fra预交.Caption = "【门诊预交信息】"
        tbDeposit.Tabs(1).Selected = True
    ElseIf mbln住院预交 And Not mbln门诊预交 Then
        '只有住院预交
        fra预交.Caption = "【住院预交信息】"
         tbDeposit.Tabs(2).Selected = True
    Else
        '两者都有
        fra预交.Caption = "【门诊及住院预交】"
         tbDeposit.Tabs(1).Selected = True
    End If
    
    If mbln门诊预交 Or mbln住院预交 Then
        
        With tbDeposit
            mblnNotClick = True
            .Tabs.Clear
            If mbln门诊预交 Then .Tabs.Add(, "K1", "门诊预交(&M)").Selected = IIf(mbytPrepayType = 1, True, False)
            If mbln住院预交 Then .Tabs.Add(, "K2", "住院预交(&Z)").Selected = IIf(mbytPrepayType = 2, True, False)
            If .Tabs.Count > 0 And .SelectedItem Is Nothing Then
               .Tabs(0).Selected = True
            End If
             
             mblnNotClick = False
            'If Not .SelectedItem Is Nothing Then Call tbDeposit_Click
            
             fra预交.Visible = .Tabs.Count <> 0
            If .Tabs.Count <> 0 And tbDeposit.SelectedItem Is Nothing Then Call RefreshFactNo
         End With
         
    End If
    mblnNotClick = False
    mintPriceGradeStartType = GetPriceGradeStartType()
    If mintPriceGradeStartType <> 0 Then
         Call GetPriceGrade(gstrNodeNo, 0, 0, str医疗付款方式, , , mstrPriceGrade)   '读取价格等级
    End If
    
    Call SetCardEditEnabled(1, True)
    Call SetDepositEditEnabled(1)
    fra磁卡.Visible = mlngCardTypeID <> 0
    tbSendCard.Visible = mlngCardTypeID <> 0
    
    If GetPublicPatient(objPubPatient) = False Then Exit Function
    
    Call SetDepositEditEnabled '设置预
   
    fra磁卡.Tag = ""
    lbl卡名称.Visible = False
    If mlngCardTypeID <> 0 Then
        chk记帐.value = IIf(mbln记账, 1, 0)
        chk记帐.Tag = IIf(mbln记账, 1, 0)
        '发卡及绑定卡设置
        If mobjOneCardComLib.zlGetCard(mlngCardTypeID, False, objSendCard) = False Then
            fra磁卡.Visible = False: tbSendCard.Visible = False
            Exit Function
        End If
        If objSendCard Is Nothing Then
           fra磁卡.Visible = False: tbSendCard.Visible = False
            Exit Function
        End If
        
        Set mCurSendCard.objSendCard = objSendCard
        mCurSendCard.lng共用批次 = 0
        
        str批次 = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModule, "0")
        varData = Split(str批次, "|")
        For i = 0 To UBound(varData)
             varTemp = Split(varData(i), ",")
             If Val(varTemp(0)) <> 0 Then
                If ExistShareBill(Val(varTemp(0)), 5) Then
                    If Val(varTemp(1)) = objSendCard.接口序号 Then
                        mCurSendCard.lng共用批次 = Val(varTemp(0)): Exit For
                    End If
                End If
             End If
        Next
        lbl卡名称.Visible = True
        lbl卡名称.Caption = "【" & objSendCard.名称 & "】"
        
        txt卡号.PasswordChar = IIf(objSendCard.卡号密文规则 <> "", "*", "")
        txt卡号.MaxLength = objSendCard.卡号长度
        
        '有效时间处理
        chkEndTime.value = 0
        If objSendCard.缺省有效时间 <> "" Then
            chkEndTime.value = vbChecked
            dtpDate = Format(objSendCard.zlGetDefaultDate, "yyyy-MM-dd 23:59:59")
        ElseIf objPubPatient.blnRealName Then
             dtpDate = Format(DateAdd("D", mbytRegValidDays, zlDatabase.Currentdate), "yyyy-MM-dd 23:59:59")
        Else
            dtpDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
            dtpDate.Enabled = False
        End If
        mblnNotClick = True
        
        '医疗卡发卡相关控制
        '1.先处理是否允许发卡
        blnAllowSendCard = mblnAllowSendCard And objSendCard.是否发卡
        If blnAllowSendCard = False Then '允许发卡
            '移除发卡页
            Call RemoveSendCardTabFromKey("CardFee")
        ElseIf Not CheckIsExstSendCardTabFromKey("CardFee") Then
            '不存在发卡页签，需要增加
            Call RemoveSendCardTabFromKey   '移除所有选项卡，再增加
            tbSendCard.Tabs.Add , "CardFee", "收费发卡(&1)"
            tbSendCard.Width = GetSendCardTabsWidth
        End If
        
        '2.处理绑定卡s
        
        blnAllowBoundCard = mblnAllowBoundCard And (objSendCard.自制卡 = False Or objSendCard.卡号重复使用)   '允许绑定卡
        If Not blnAllowBoundCard Then
            '不存在绑定卡号
            Call RemoveSendCardTabFromKey("CardBind")
        ElseIf Not CheckIsExstSendCardTabFromKey("CardBind") Then
            tbSendCard.Tabs.Add , "CardBind", "绑定卡号(&2)"
        End If
        tbSendCard.Width = GetSendCardTabsWidth
        
        '缺省定位
        Select Case zlDatabase.GetPara("发卡模式", glngSys, mlngModule, "CardFee")
        Case "CardFee"
              mblnNotClick = True
              If CheckIsExstSendCardTabFromKey("CardFee") Then tbSendCard.Tabs("CardFee").Selected = True
              mblnNotClick = False
        Case "CardBind"
              mblnNotClick = True
              If CheckIsExstSendCardTabFromKey("CardBind") Then tbSendCard.Tabs("CardBind").Selected = True
              mblnNotClick = False
        End Select
        
        If tbSendCard.SelectedItem Is Nothing Then
            If tbSendCard.Tabs.Count > 0 Then
                mblnNotClick = True
                tbSendCard.Tabs(1).Selected = True
                mblnNotClick = False
            End If
        End If
        mblnNotClick = False
        tbSendCard.Width = GetSendCardTabsWidth '设置缺省的宽度
        
        Call InitCardFee '加载卡费数据
        If objSendCard.是否严格控制 Then
            
            mCurSendCard.lng领用ID = mobjExseSvr.CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), , objSendCard.接口序号)
            If mCurSendCard.lng领用ID <= 0 Then
                Select Case mCurSendCard.lng领用ID
                    Case 0 '操作失败
                    Case -1
                        'MsgBox "你没有自用或共用的就诊卡,不能发放！" & vbCrLf & _
                        "请先在本地设置共用批次或领用一批新卡! ", vbExclamation, gstrSysName
                    Case -2
                        ' MsgBox "本地共用的就诊卡已用完,不能发放！" & vbCrLf & _
                        "请重新设置本地共用卡批次或领用一批新卡！", vbExclamation, gstrSysName
                    End Select
            End If
        End If
        '初始化卡费信息值
        Call SetSendCardCtrolVisibled
    ElseIf mblnShowDepositAndSendCard Then
        fra磁卡.Visible = True: fra磁卡.Tag = "1"
        tbSendCard.Visible = False
        Call SetCardEditEnabled(0)
    End If
    
    If Not tbSendCard.SelectedItem Is Nothing Then tbSendCard_Click
    If Not tbDeposit.SelectedItem Is Nothing Then tbDeposit_Click
    InitFace = True
    Exit Function
errHandle:
    mblnNotClick = False
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitCardFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡费值
    '编制:刘兴洪
    '日期:2019-11-25 10:32:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs卡费 As ADODB.Recordset
    Dim str费别 As String, dblMoney As Double
    
    On Error GoTo errHandle
    Set rs卡费 = GetCardFee()
    
    If rs卡费 Is Nothing Then
        txt卡额.Text = "": txt卡额.Tag = ""
        Exit Sub
    End If
    If rs卡费.RecordCount = 0 Then
        txt卡额.Text = "": txt卡额.Tag = ""
        Exit Sub
    End If
    With rs卡费
        str费别 = ""
        If Not mobjPati Is Nothing Then str费别 = mobjPati.费别
        txt卡额.Text = Format(IIf(Nvl(!是否变价, 0) = 1, Val(Nvl(!缺省价格)), Val(Nvl(!现价))), "0.00")
        If Nvl(!是否变价, 0) <> 1 And Nvl(!屏蔽费别, 0) <> 1 Then
            If mobjExseSvr.zl_ExseSvr_Actualmoney(str费别, !收费细目ID, !收入项目ID, Val(txt卡额.Text), dblMoney) Then
                txt卡额.Text = Format(dblMoney, "0.00")
            End If
        End If
        txt卡额.Tag = txt卡额.Text  '保持不变
        txt卡额.Locked = Nvl(!是否变价, 0) <> 1
        txt卡额.TabStop = Nvl(!是否变价, 0) = 1
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetSendCardTabsWidth() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发卡页面的总宽度
    '返回:返回总宽度
    '编制:刘兴洪
    '日期:2019-11-23 16:12:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, lngWidth As Long
    
    For i = 1 To tbSendCard.Tabs.Count
        
        lngWidth = lngWidth + tbSendCard.Tabs(i).Width + Me.TextWidth("刘")
    Next
    GetSendCardTabsWidth = lngWidth
End Function

Private Function CheckIsExstSendCardTabFromKey(ByVal strKey As String, Optional ByRef indIndex_Out As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否存在指定Key值的页面
    '入参:
    '出参:indIndex_Out-存在时，返回该tab的索引
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-23 16:04:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    indIndex_Out = tbSendCard.Tabs(strKey).Index
    If Err <> 0 Then
        Err = 0: On Error GoTo 0: CheckIsExstSendCardTabFromKey = False
        Exit Function
    End If
    CheckIsExstSendCardTabFromKey = True
End Function


Private Function RemoveSendCardTabFromKey(Optional ByVal strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移除指定Key值的页面
    '入参:strKey=""时，表示移除所有卡
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-23 16:04:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Err = 0: On Error Resume Next
    If strKey = "" Then
        tbSendCard.Tabs.Clear
    Else
        Call tbSendCard.Tabs.Remove(strKey)
        If Err <> 0 Then
            Err = 0: On Error GoTo 0
            
        End If
    End If
    RemoveSendCardTabFromKey = True
End Function
 
Private Sub btQRCodeTemp_GotFocus()
    RaiseEvent ControlGotFocus(btQRCodeTemp)
End Sub

 

Private Sub cbo发卡结算_GotFocus()
    RaiseEvent ControlGotFocus(cbo发卡结算)
End Sub

Private Sub cbo预交结算_Click()
    Dim objPayCard As Card
    If mblnNotClick = True Then Exit Sub
    If Not cbo预交结算.Enabled Then Exit Sub
    
    Set objPayCard = GetDepositPayCard
    If objPayCard Is Nothing Then Exit Sub
    
    Call SetDepositEditEnabled
    
    If txt缴款单位.Text <> "" And txt缴款单位.Enabled = True Then
        chk单位缴款.value = 1
    Else
        chk单位缴款.value = 0
    End If
    Call Local结算方式(objPayCard.接口序号, False, IIf(cbo预交结算.ItemData(cbo预交结算.ListIndex) <> 5, cbo预交结算.Text, ""))
    'Call chk单位缴款_Click
End Sub

Private Sub cbo预交结算_GotFocus()
    RaiseEvent ControlGotFocus(cbo预交结算)
End Sub

Private Sub chkEndTime_GotFocus()
    RaiseEvent ControlGotFocus(chkEndTime)
End Sub

Private Sub chk单位缴款_Click()
    If chk单位缴款.value = 1 And cbo预交结算.Enabled Then
        txt缴款单位.Enabled = True
        txt缴款单位.BackColor = &H80000005
    Else
        txt缴款单位.Text = ""
        txt缴款单位.Enabled = False
        txt缴款单位.BackColor = Me.BackColor
    End If
End Sub
Private Sub cbo发卡结算_Click()
    Dim objPayCard As Card
     
    If mblnNotClick = True Then Exit Sub
    
'    Set objPayCard = GetCardFeePayCard
'    If objPayCard Is Nothing Then Exit Sub
'    Call Local结算方式(objPayCard.接口序号, True)
End Sub

Private Sub cbo发卡结算_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo发卡结算.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo发卡结算.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then
        cbo发卡结算.ListIndex = lngIdx
    End If
End Sub

Private Sub chk单位缴款_GotFocus()
    
    RaiseEvent ControlGotFocus(chk单位缴款)
End Sub

Private Sub chk记帐_GotFocus()
    RaiseEvent ControlGotFocus(chk记帐)
End Sub

Private Sub dtpDate_GotFocus()
    RaiseEvent ControlGotFocus(dtpDate)
End Sub

Private Sub Form_Activate()
    RaiseEvent Activate
End Sub
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF6 '扫码付快键
            RaiseEvent ExcuteReadQRCode
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If Me.ActiveControl Is chkEndTime Then
            If Not tbSendCard.SelectedItem Is Nothing Then
                If tbSendCard.SelectedItem.Key = "BoundCard" Then
                    If chkEndTime.value <> 1 Then
                        RaiseEvent InputOver    '输入结束
                        Exit Sub
                    End If
                End If
            End If
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If Me.ActiveControl Is chk记帐 Then
            If cbo发卡结算.Enabled And cbo发卡结算.Visible Then cbo发卡结算.SetFocus
            If chk记帐.value = Checked And Visible Then
                RaiseEvent InputOver    '输入结束
            End If
            Exit Sub
        End If
        
        
        If Me.ActiveControl Is dtpDate Then
            If Not tbSendCard.SelectedItem Is Nothing Then
                If tbSendCard.SelectedItem.Key = "BoundCard" Then
                    RaiseEvent InputOver    '输入结束
                    Exit Sub
                End If
            End If
        End If
        
        If Not (Me.ActiveControl Is txt预交额 Or Me.ActiveControl Is txtPass Or Me.ActiveControl Is txtAudi Or Me.ActiveControl Is txt卡号) Then
            zlCommFun.PressKey vbKeyTab
        End If
        Exit Sub
     End If
     If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call HookDefend(txtPass.hWnd)
    Call HookDefend(txtAudi.hWnd)
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    fra磁卡.Width = Me.ScaleWidth - fra磁卡.Left * 2
    fra预交.Left = fra磁卡.Left
    fra预交.Width = fra磁卡.Width
    
    txt帐号.Left = fra预交.Left + fra预交.Width - txt帐号.Width - fra预交.Left * 2 - 50
    lblAccno.Left = txt帐号.Left - lblAccno.Width - 20
    txt结算号码.Left = txt帐号.Left
    lblCode.Left = txt结算号码.Left - lblCode.Width
    chk单位缴款.Left = txt帐号.Left + txt帐号.Width - chk单位缴款.Width
    
    dtpDate.Left = fra磁卡.Left + fra磁卡.Width - dtpDate.Width - fra磁卡.Left * 2 - 50
    chkEndTime.Top = dtpDate.Top + (dtpDate.Height - chkEndTime.Height) \ 2
    chkEndTime.Left = dtpDate.Left - chkEndTime.Width
    
    txtAudi.Left = chkEndTime.Left - txtAudi.Width - 200
    lbl验证.Left = txtAudi.Left - lbl验证.Width - 20
    
    txtPass.Left = lbl验证.Left - txtPass.Width - 50
    lbl密码.Left = txtPass.Left - lbl密码.Width - 20
    
    cbo发卡结算.Left = txtPass.Left
    lbl结算方式.Left = cbo发卡结算.Left - lbl结算方式.Width - 20
    
    chk记帐.Left = lbl结算方式.Left - chk记帐.Width - 50
    
     lbl卡名称.Left = fra磁卡.Left + fra磁卡.Width - lbl卡名称.Width - 200
End Sub

Private Sub tbDeposit_Click()
    If mblnNotClick Then Exit Sub
    If tbDeposit.SelectedItem Is Nothing Then Exit Sub
    
    Set mobjDepositFact = mobjExseSvr.zl_GetInvoicePreperty(mlngModule, 2, Mid(tbDeposit.SelectedItem.Key, 2))
    mobjDepositFact.领用ID = 0
    Call RefreshFactNo
    If txt预交额.Enabled And txt预交额.Visible Then txt预交额.SetFocus
End Sub

Private Sub tbSendCard_Click()
    If mblnNotClick Then Exit Sub
    Call SetSendCardCtrolVisibled   '调整控件位置及visible属性
End Sub

Private Sub chkEndTime_Click()
    dtpDate.Enabled = chkEndTime.value
End Sub
Private Sub chk记帐_Click()
    
    cbo发卡结算.Enabled = chk记帐.value <> Checked
    Call CalcRQCodePayTotal '计算总额

End Sub

Private Sub SetSendCardCtrolVisibled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整发卡控件的Visibled属性，并调整相应的控件位置
    '编制:刘兴洪
    '日期:2019-11-23 15:29:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSendCard As Boolean
    Dim lngTop As Long
    
'    If Not fra磁卡.Visible Then Exit Sub
    If tbSendCard.SelectedItem Is Nothing Then
        blnSendCard = False
    Else
        blnSendCard = tbSendCard.SelectedItem.Key = "CardFee"
    End If
    lbl金额.Visible = blnSendCard
    txt卡额.Visible = blnSendCard
    chk记帐.Visible = blnSendCard
    cbo发卡结算.Visible = blnSendCard
    lbl结算方式.Visible = blnSendCard
    '调整对应位置
    
    If blnSendCard Then
        lngTop = tbSendCard.Height + 45
    Else
        lngTop = (fra磁卡.Height - txt卡号.Height + tbSendCard.Height \ 2) \ 2
    End If
    
    txt卡号.Top = lngTop: lbl卡号.Top = txt卡号.Top + (txt卡号.Height - lbl卡号.Height) \ 2
    txtPass.Top = lngTop: lbl密码.Top = txtPass.Top + (txtPass.Height - lbl密码.Height) \ 2
    txtAudi.Top = lngTop: lbl验证.Top = txtAudi.Top + (txtAudi.Height - lbl验证.Height) \ 2
    dtpDate.Top = lngTop: chkEndTime.Top = dtpDate.Top + (dtpDate.Height - chkEndTime.Height) \ 2
    
    txt卡额.Top = txt卡号.Top + txt卡号.Height + 80: lbl金额.Top = txt卡额.Top + (txt卡额.Height - lbl金额.Height) \ 2
    cbo发卡结算.Top = txt卡额.Top
    chk记帐.Top = txt卡额.Top + (txt卡额.Height - chk记帐.Height) \ 2
    lbl结算方式.Top = cbo发卡结算.Top + (cbo发卡结算.Height - lbl结算方式.Height) \ 2
End Sub


Private Sub SetCardEditEnabled(Optional bytEnabledType As Byte, Optional blnInit As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置就诊卡编辑属性
    '入参:bytEnabledType-禁用的类型:0-不限制;1-禁用结算信息;2-禁用所有信息
    '       blnInit-是否时初始化界面
    '编制:刘兴洪
    '日期:2019-12-02 11:37:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    Dim bln实名认证 As Boolean
    
    If Not mobjPati Is Nothing Then
        If mobjPati.实名认证 Then bln实名认证 = mobjPati.实名认证
    End If
    
    Select Case bytEnabledType
    Case 2  '2-禁用所有信息
        blnEdit = False
        txt卡号.Enabled = blnEdit
        cbo发卡结算.Enabled = blnEdit
        txt卡额.Enabled = blnEdit
        chk记帐.Enabled = blnEdit
        chkEndTime.Enabled = blnEdit
        dtpDate.Enabled = blnEdit
        lbl卡号.Enabled = blnEdit
        tbSendCard.Enabled = blnEdit
    Case Else   '0-不作限制,1-禁用结算信息
        blnEdit = mlngCardTypeID <> 0 And mint操作状态 <> 2
        txt卡号.Enabled = blnEdit
        blnEdit = Trim(txt卡号.Text) <> "" And mlngCardTypeID <> 0 And mint操作状态 <> 2
        
        cbo发卡结算.Enabled = chk记帐.value = 0 And blnEdit And mint操作状态 <> 1
        txt卡额.Enabled = blnEdit
        dtpDate.Enabled = chkEndTime.value = 1 And (bln实名认证 Or mobjPubPatient.blnRealName = False)
        chkEndTime.Enabled = blnEdit And (bln实名认证 Or mobjPubPatient.blnRealName = False)
        chk记帐.Enabled = blnEdit
        lbl卡号.Enabled = blnEdit
        
        If bytEnabledType = 1 Then
            lbl结算方式.Enabled = False
            cbo发卡结算.Enabled = False
            txt卡额.Enabled = False
            chk记帐.Enabled = False
            tbSendCard.Enabled = blnInit
        End If
    End Select
    txtPass.Enabled = blnEdit: txtAudi.Enabled = blnEdit
    
    '设置颜色
    txtPass.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    txtAudi.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    txt卡号.BackColor = IIf(mlngCardTypeID <> 0, &HEBFFFF, &H8000000F)
    txt卡额.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    cbo发卡结算.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    
End Sub
Private Sub SetDepositEditEnabled(Optional bytEnabledType As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置预交的编辑属性
    '入参:bytEnabledType-禁用的类型:0-不限制;1-禁用结算信息及预交信息;2-禁用所有信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-12-02 11:13:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    Dim objPayCard As Card, int性质 As Integer
    Set objPayCard = GetDepositPayCard
    
    blnEdit = (Not objPayCard Is Nothing)
    If Not objPayCard Is Nothing Then
        int性质 = objPayCard.结算性质
    End If
    
    blnEdit = blnEdit And fra预交.Tag = "" And mint操作状态 <> 2    '无预交信息
    
    Select Case bytEnabledType
    Case 2  '2-禁用所有信息
        blnEdit = False: txtFact.Enabled = False
    
    Case Else   '0-不作限制,1-禁用结算信息及预交信息
         txtFact.Enabled = blnEdit
         If bytEnabledType = 1 Then
            blnEdit = False
         End If
    End Select
    
    txt预交额.Enabled = blnEdit
    cbo预交结算.Enabled = blnEdit And mint操作状态 <> 1
    txt结算号码.Enabled = blnEdit
    tbDeposit.Enabled = blnEdit
    
    If blnEdit Then blnEdit = int性质 <> 3
    If blnEdit Then blnEdit = chk单位缴款.value = 1
    chk单位缴款.Enabled = blnEdit: txt开户行.Enabled = blnEdit
    txt帐号.Enabled = blnEdit
    txt缴款单位.Enabled = blnEdit And chk单位缴款.value = 1
    
    '设置颜色
    txt预交额.BackColor = IIf(txt预交额.Enabled, &H80000005, &H8000000F)
    txtFact.BackColor = IIf(txtFact.Enabled, &H80000005, &H8000000F)
    cbo预交结算.BackColor = IIf(cbo预交结算.Enabled, &H80000005, &H8000000F)
    txt结算号码.BackColor = IIf(txt结算号码.Enabled, &H80000005, &H8000000F)
    txt开户行.BackColor = IIf(txt开户行.Enabled, &H80000005, &H8000000F)
    txt帐号.BackColor = IIf(txt帐号.Enabled, &H80000005, &H8000000F)
    txt缴款单位.BackColor = IIf(txt缴款单位.Enabled, &H80000005, &H8000000F)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If fra磁卡.Visible Then
        zlDatabase.SetPara "发卡模式", tbSendCard.SelectedItem.Key, glngSys, mlngModule
    End If
    
    Set mCurSendCard.objSendCard = Nothing
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    Set mbtQRCodePay = Nothing
    Set mobjOneCardComLib = Nothing
    Set mobjPubPatient = Nothing
    Set mobjService = Nothing
    Set mobjExseSvr = Nothing
    Set mobjPati = Nothing
    Set mobjThirdSwap = Nothing
    Set mfrmMain = Nothing
    Set mobjCommEvents = Nothing
    Set mrs卡费 = Nothing
    Set mobjCardFeePayCards = Nothing
    Set mobjDepositFact = Nothing
    Set mobjDepositPayCards = Nothing
    Set mobjShowTotalMoneyControl = Nothing
    Set mobjCardFeeItems = Nothing
    Set mrsCardFee = Nothing
    Set mobjDepositItems = Nothing
    Set mobjKeyboard = Nothing
    mCurSendCard.lng领用ID = 0
    mblnInited = False: mblnNotClick = False: mblnSendCardLocked = False
    mblnDepositLocked = False
    mblnBoundCarded = False: mblnShowDepositAndSendCard = False
    mint操作状态 = 0
End Sub


Public Function GetCardFee(Optional blnReReadCardFee As Boolean = False, Optional ByVal strPriceGrade As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡费数据
    '入参:blnRereadCardFee-重新读取卡卡费
    '     strPriceGrade-价格等级
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-23 17:31:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card
    
    On Error GoTo errHandle
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Set GetCardFee = Nothing: Exit Function
    If objSendCard.特定项目 = "" Then Set GetCardFee = Nothing: Exit Function
    If Not mrs卡费 Is Nothing Then
        If mrs卡费.State = 1 Then Set GetCardFee = mrs卡费: Exit Function
    End If
    
    Set mrs卡费 = zlGetSpecialItemFee(objSendCard.特定项目, strPriceGrade)
    Set GetCardFee = mrs卡费
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetPublicPatient(ByRef objPubPati_Out As clsInterFacePatient) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建zlPublicPatient部件
    '入参:
    '出参:objPubPati-返回病人公共部件
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 10:11:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjPubPatient Is Nothing Then Set objPubPati_Out = mobjPubPatient: GetPublicPatient = True: Exit Function
    On Error GoTo errHandle
    
    Set mobjPubPatient = New clsInterFacePatient
    If mobjPubPatient.Init(Me, glngSys, glngModul, gcnOracle, gstrDBUser) = False Then Exit Function
    Set objPubPati_Out = mobjPubPatient
    GetPublicPatient = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub lbl卡号_Click()
    Dim strExpand As String, strOutCardNO As String, strPatiInfoXML As String
    Dim objSendCard As Card
    If mblnSendCardLocked Then Exit Sub
    

    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    
    If objSendCard.名称 = "就诊卡" And objSendCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        
        If Not mobjICCard Is Nothing Then
            txt卡号.Text = mobjICCard.Read_Card()
            If txt卡号.Text <> "" Then
                mblnICCard = True
                Call CheckFreeCard(txt卡号.Text)
            End If
        End If
        Exit Sub
    End If
    If (objSendCard.是否接触式读卡 = False And objSendCard.是否非接触式读卡 = False) Or objSendCard.接口序号 <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strPatiInfoXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strPatiInfoXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\

    If mobjOneCardComLib.zlReadCard(Me, mlngModule, objSendCard.接口序号, False, strExpand, strOutCardNO, strPatiInfoXML) = False Then Exit Sub
    txt卡号.Text = strOutCardNO
    If txt卡号.Text <> "" Then
        '问题号:56599
        If strPatiInfoXML <> "" Then RaiseEvent RequestRefreshPatiInf(strOutCardNO, strPatiInfoXML)
        Call CheckFreeCard(txt卡号.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    End If
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt卡号.Text = strCardNo
    If txt卡号.Text <> "" Then
        '问题号:56599
        If strXmlCardInfor <> "" Then RaiseEvent RequestRefreshPatiInf(strCardNo, strXmlCardInfor)
        Call CheckFreeCard(txt卡号.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    End If
End Sub
 
Private Sub CheckFreeCard(ByVal strCardNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对一卡通模式下的卡号，严格控制票号时，检查是否在票据领用范围内，范围之外的卡不收费
    '入参:strCardNo-卡号
    '日期:2019-11-25 12:01:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs卡费 As ADODB.Recordset
    Dim str费别 As String, dblMoney As Double
    Dim objSendCard As Card
    
    If txt卡额.Visible = False Then Exit Sub
    
    Set rs卡费 = GetCardFee()
    If Not mobjPati Is Nothing Then str费别 = mobjPati.费别
    
    If Not rs卡费 Is Nothing And Val(txt卡额.Text) = 0 Then  '先恢复
        txt卡额.Text = Format(IIf(rs卡费!是否变价 = 1, rs卡费!缺省价格, rs卡费!现价), "0.00")
        txt卡额.Tag = txt卡额.Text
    End If
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    If mCurSendCard.blnOneCard And objSendCard.是否严格控制 Then
        mCurSendCard.lng领用ID = mobjExseSvr.CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), strCardNo)
        If mCurSendCard.lng领用ID <= 0 Then txt卡额.Text = "0.00": txt卡额.Tag = txt卡额.Text
    End If

    If Not rs卡费 Is Nothing And Val(txt卡额.Text) <> 0 Then
        If rs卡费!是否变价 = 0 Then
            If mobjExseSvr.zl_ExseSvr_Actualmoney(str费别, rs卡费!收费细目ID, rs卡费!收入项目ID, rs卡费!现价, dblMoney) Then
                txt卡额.Text = Format(dblMoney, "0.00")
                txt卡额.Tag = txt卡额.Text
           End If
        End If
    End If
End Sub


Private Function GetDepositBalanceItems(ByVal dtCurdate As Date, ByRef objBalanceItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前预交结算信息
    '入参:dtCurDate-当前时间
    '出参:objBalanceItems_Out-预交的结算对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 20:16:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, strDepositNo As String
    Dim objCurItem As clsBalanceItem
    Dim lng预交ID As Long
    On Error GoTo errHandle
    
    
    Set objBalanceItems_Out = New clsBalanceItems
    If fra预交.Visible = False Or StrToNum(txt预交额.Text) = 0 Then GetDepositBalanceItems = True: Exit Function
    
    
    Set objCard = GetDepositPayCard()
    If objCard Is Nothing Then Exit Function
    Set objCurItem = New clsBalanceItem
    
    If mobjExseSvr.zl_ExseSvr_GetNextNo(11, strDepositNo) = False Then Exit Function   '预交No
    If mobjExseSvr.zl_ExseSvr_GetNextID("病人预交记录", lng预交ID) = False Then Exit Function
     
    With objCurItem
        Set .objCard = objCard
        .卡类别ID = IIf(objCard.接口序号 < 0, 0, objCard.接口序号)
        .消费卡 = objCard.消费卡
        .结算方式 = objCard.结算方式
        .结算号码 = Trim(txt结算号码.Text)
        .结算金额 = StrToNum(txt预交额.Text)
        .结算性质 = objCard.结算性质
        If .卡类别ID > 0 Then
           .结算类型 = IIf(.消费卡, 5, 3)
        ElseIf objCard.结算性质 = 3 Then
             .结算类型 = 2
        Else
           .结算类型 = 0  '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        End If
        .结算时间 = dtCurdate
        .结算摘要 = ""
        .单据号 = strDepositNo
        .预交ID = lng预交ID
        .是否预交 = True
    End With
    objBalanceItems_Out.AddItem objCurItem
    objBalanceItems_Out.结算金额 = objCurItem.结算金额
    objBalanceItems_Out.单据号 = objCurItem.单据号
    objBalanceItems_Out.类型 = objCurItem.结算类型
    objBalanceItems_Out.结算时间 = Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")
    
    GetDepositBalanceItems = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCardFeeBalanceItems(ByVal dtCurdate As Date, ByRef objBalanceItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前卡费结算信息
    '入参:dtCurDate-当前时间
    '出参:objBalanceItems_Out-卡费的结算对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 20:16:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, strCardFeeNo As String
    Dim objCurItem As clsBalanceItem, lng结帐ID As Long
    On Error GoTo errHandle
    
    Set objBalanceItems_Out = New clsBalanceItems
    If fra磁卡.Visible = False Or tbSendCard.SelectedItem.Key <> "CardFee" Or txt卡号.Text = "" Then GetCardFeeBalanceItems = True: Exit Function
    
    
    If chk记帐.value = 1 Then
        If mobjExseSvr.zl_ExseSvr_GetNextNo(16, strCardFeeNo) = False Then Exit Function     '医疗卡单据号
        objBalanceItems_Out.结算金额 = StrToNum(txt卡额.Text)
        objBalanceItems_Out.单据号 = strCardFeeNo
        objBalanceItems_Out.类型 = gEM_记帐单
        GetCardFeeBalanceItems = True
        Exit Function
    End If
    
    Set objCard = GetCardFeePayCard()
    If objCard Is Nothing Then Exit Function
    Set objCurItem = New clsBalanceItem
    
    If mobjExseSvr.zl_ExseSvr_GetNextNo(16, strCardFeeNo) = False Then Exit Function    '医疗卡单据号
    If mobjExseSvr.zl_ExseSvr_GetNextID("病人结帐记录", lng结帐ID) = False Then Exit Function
    
    With objCurItem
        Set .objCard = objCard
        .卡类别ID = IIf(objCard.接口序号 < 0, 0, objCard.接口序号)
        .消费卡 = objCard.消费卡
        .结算方式 = objCard.结算方式
        .结算号码 = ""
        .结算金额 = StrToNum(txt卡额.Text)
        .结算性质 = objCard.结算性质
        If .卡类别ID > 0 Then
           .结算类型 = IIf(.消费卡, 5, 3)
        Else
           .结算类型 = 0  '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        End If
        .结算时间 = dtCurdate
        .结算摘要 = ""
        .结算ID = lng结帐ID
        .单据号 = strCardFeeNo
        .是否预交 = False
        
    End With
    objBalanceItems_Out.AddItem objCurItem
    objBalanceItems_Out.结算金额 = objCurItem.结算金额
    objBalanceItems_Out.单据号 = objCurItem.单据号
    objBalanceItems_Out.类型 = objCurItem.结算类型
    objBalanceItems_Out.结算时间 = Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")
    
    GetCardFeeBalanceItems = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckDepsoitAndCardFeePayIsSame(ByVal objDepositItems As clsBalanceItems, ByVal objCardFeeItems As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医疗卡及预交款是否同一种支付方式
    '入参:objDepositItems-当前的预交结算
    '     objCardFeeItems-当前的卡费结算
    '出参:
    '返回:同一种返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-26 16:02:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If objDepositItems Is Nothing Or objCardFeeItems Is Nothing Then Exit Function
    If objDepositItems.Count <> objCardFeeItems.Count Then Exit Function
    If objDepositItems.Count = 0 Or objCardFeeItems.Count = 0 Then Exit Function
    
    If objDepositItems(1).卡类别ID = objCardFeeItems(1).卡类别ID And objDepositItems(1).消费卡 = objCardFeeItems(1).消费卡 Then
        CheckDepsoitAndCardFeePayIsSame = True
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetErrDataToColl(ByVal objPati As clsPatientInfo, ByVal lng业务ID As Long, _
    ByVal objCurItems As clsBalanceItems, ByVal int同步标志 As Integer, ByRef lng异常id_Out As Long, ByVal dtCurdate As Date, _
    ByRef cllErrData_out As Collection, Optional ByVal str预交单号 As String, Optional ByVal dbl预交金额 As Double, _
    Optional ByVal str卡费单号 As String, Optional dbl卡费 As Double, Optional blnCancel As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取异常的收费明单数据
    '入参:int操作状态-:0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
    '     objCurItem-当前结算信息
    '     lng异常id-异常ID
    '     dtCurDate-当前日期
    '     int同步标志:   0-正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,4-医疗卡信息更新成功;
    '     blnCancel-是否作废
    '出参:lng异常id_Out-异常ID
    '       cllErrData_Out-返回错误信息集(格式为Array(保存项名称,保存项值)
    '          保存的项目名称包含: 异常ID,操作场景,作废标志,业务id,是否病历费,病人id,主页id,姓名,性别,年龄,门诊号,住院号,预交单号,预交金额,医疗卡单号,卡费,发卡类别id,发卡类别名称,发卡卡号,同步状态,交易信息)
    '          其中交易信息为Json串，格式如下
    '           {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]})
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-26 16:16:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int操作场景 As Integer, i As Long
    Dim objCurItem As clsBalanceItem
    
    On Error GoTo errHandle
    lng异常id_Out = zlDatabase.GetNextId("病人结算异常记录")
    '1-医疗卡发卡;2-病人信息登记;3-病人入院登记;4-预约挂号接收
    int操作场景 = IIf(mlngModule = 1101, 2, 3)
    Set cllErrData_out = New Collection
    cllErrData_out.Add Array("异常ID", lng异常id_Out)
    cllErrData_out.Add Array("操作场景", int操作场景)
    cllErrData_out.Add Array("作废标志", IIf(blnCancel, 1, 0))
    cllErrData_out.Add Array("业务ID", lng业务ID)
    cllErrData_out.Add Array("是否病历费", 0)
    cllErrData_out.Add Array("病人ID", objPati.病人ID)
    cllErrData_out.Add Array("主页ID", objPati.主页ID)
    
    cllErrData_out.Add Array("姓名", objPati.姓名)
    cllErrData_out.Add Array("性别", objPati.性别)
    cllErrData_out.Add Array("年龄", objPati.年龄)
    cllErrData_out.Add Array("门诊号", objPati.门诊号)
    cllErrData_out.Add Array("住院号", objPati.住院号)
    
    cllErrData_out.Add Array("预交单号", str预交单号)
    cllErrData_out.Add Array("预交金额", dbl预交金额)
    cllErrData_out.Add Array("医疗卡单号", str卡费单号)
    cllErrData_out.Add Array("卡费", dbl卡费)
    cllErrData_out.Add Array("操作员姓名", UserInfo.姓名)
    cllErrData_out.Add Array("操作员编号", UserInfo.编号)
    cllErrData_out.Add Array("登记时间", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS"))
    
    
    
    If str卡费单号 <> "" Then
        cllErrData_out.Add Array("发卡类别ID", mCurSendCard.objSendCard.接口序号)
        cllErrData_out.Add Array("发卡类别名称", mCurSendCard.objSendCard.名称)
        cllErrData_out.Add Array("发卡卡号", txt卡号.Text)
    End If
    cllErrData_out.Add Array("同步状态", int同步标志)
    Dim strJson As String
    
    strJson = ""
    If Not objCurItems Is Nothing Then
        objCurItems.异常ID = lng异常id_Out
        objCurItems.业务ID = lng业务ID
        If objCurItems.Count <> 0 Then
            For i = 1 To objCurItems.Count
                objCurItems(i).异常ID = lng异常id_Out
            Next
            Set objCurItem = objCurItems(1)
            strJson = strJson & "" & GetJsonNodeString("card_no", objCurItem.卡号, Json_Text)
            strJson = strJson & "," & GetJsonNodeString("cardtype_id", objCurItem.卡类别ID, Json_num)
            strJson = "{" & strJson & "}"
        End If
    End If
    cllErrData_out.Add Array("交易信息", strJson)
    zlGetErrDataToColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function GetDelErrDataToColl(ByVal lng业务ID As Long, lng异常ID As Long, ByRef cllErrData_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取异常的收费明单数据
    '入参: lng异常ID-异常ID
    '     lng异常id-异常ID
    '出参:lng异常id_Out-异常ID
    '       cllErrData_Out-返回错误信息集(格式为Array(保存项名称,保存项值)
    '          保存的项目名称包含: 异常ID,操作场景,作废标志,业务id,是否病历费,病人id,主页id,姓名,性别,年龄,门诊号,住院号,预交单号,预交金额,医疗卡单号,卡费,发卡类别id,发卡类别名称,发卡卡号,同步状态,交易信息)
    '          其中交易信息为Json串，格式如下
    '           {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]})
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-26 16:16:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int操作场景 As Integer, int同步标志 As Integer
    Dim objCurItem As clsBalanceItem
    On Error GoTo errHandle
   
    '1-医疗卡发卡;2-病人信息登记;3-病人入院 登记;4-预约挂号接收
    int操作场景 = IIf(mlngModule = 1101, 2, 3)
    Set cllErrData_out = New Collection
    cllErrData_out.Add Array("异常ID", lng异常ID)
    cllErrData_out.Add Array("操作场景", int操作场景)
    cllErrData_out.Add Array("业务ID", lng业务ID)
    GetDelErrDataToColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 Private Function GetUpdateErrDataSyncTagToColl(lng异常ID As Long, ByVal int同步标志 As Integer, ByRef cllErrData_out As Collection, Optional ByVal cllSendCard As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取异常的收费明单数据
    '入参: lng异常ID-异常ID
    '     lng异常id-异常ID
    '出参:lng异常id_Out-异常ID
    '       cllErrData_Out-返回错误信息集(格式为Array(保存项名称,保存项值)
    '          保存的项目名称包含: 异常ID,操作场景,作废标志,业务id,是否病历费,病人id,主页id,姓名,性别,年龄,门诊号,住院号,预交单号,预交金额,医疗卡单号,卡费,发卡类别id,发卡类别名称,发卡卡号,同步状态,交易信息)
    '          其中交易信息为Json串，格式如下
    '           {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]})
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-26 16:16:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int操作场景 As Integer
    Dim objCurItem As clsBalanceItem
    Dim varData As Variant
    Dim i As Long
    On Error GoTo errHandle
    Set cllErrData_out = New Collection
    cllErrData_out.Add Array("异常ID", lng异常ID)
    cllErrData_out.Add Array("同步状态", int同步标志)
    If Not cllSendCard Is Nothing Then
        '需要更新卡号信息
        For i = 1 To cllSendCard.Count
            varData = cllSendCard(i)
            Select Case varData(0)
            Case "医疗卡号"
                cllErrData_out.Add Array("发卡卡号", varData(1))
            Case "卡类别ID"
                cllErrData_out.Add Array("发卡类别ID", varData(1))
            End Select
        Next
        
    End If
    GetUpdateErrDataSyncTagToColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function UpdateCardFeeBalanceInfor(ByVal int操作状态 As Integer, ByVal objPati As clsPatientInfo, _
    ByVal cllSendCardInfo As Collection, ByVal objCardFeeItems As clsBalanceItems, ByVal objDepositItems As clsBalanceItems, _
    ByVal cllExpendInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新卡费相关结算信息
    '入参:int操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
    '     objCardFeeItems-当前卡费结算支付信息
    '     objDepositItems-当前预交支付方式
    '     cllSendCardInfo-发卡信息 (卡类别ID,变动类型,卡号,原卡号,IC卡号,密码,加密密码,终止使用时间,卡费,病历费,摘要,卡号重用,领用ID),格式:array(名称,值),"_名称"
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-14 11:49:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllUpdateFeeData As Collection, cllTemp As Collection
    Dim objCurItem As clsBalanceItem, blnTrans As Boolean
    Dim strDepositNo As String, strCardFeeNo As String, strErrMsg As String
    Dim varTemp As Variant, lng预交ID As Long
    Dim cllPro As Collection, strSql As String, int异常状态 As Integer
    Dim cllErrData As Collection, dtCurdate As Date
    
    On Error GoTo errHandle
    
    Set cllUpdateFeeData = New Collection
    Set cllTemp = New Collection
    
    If objCardFeeItems.类型 <> gEM_记帐单 Then
        
        Set objCurItem = objCardFeeItems(1)
        If Not objDepositItems Is Nothing Then
            If objDepositItems.Count <> 0 Then
                strDepositNo = objDepositItems.单据号
                lng预交ID = objDepositItems(1).预交ID
            End If
        End If
    End If
    If Not objDepositItems Is Nothing Then
        strDepositNo = objDepositItems.单据号
    End If
    
    strCardFeeNo = objCardFeeItems.单据号
    cllTemp.Add Array("预交单号", strDepositNo), "_" & "预交单号"
    cllTemp.Add Array("预交ID", lng预交ID), "_" & "预交ID"
    cllTemp.Add Array("收费单号", strCardFeeNo), "_" & "收费单号"
    If Not objCurItem Is Nothing Then
        cllTemp.Add Array("结帐ID", IIf(objCurItem.冲销ID <> 0, objCurItem.冲销ID, objCurItem.结算ID)), "_" & "结帐ID"
    End If
    cllTemp.Add Array("病人ID", objPati.病人ID), "_" & "病人ID"
    cllTemp.Add Array("操作员编号", UserInfo.编号), "_" & "操作员编号"
    cllTemp.Add Array("操作员姓名", UserInfo.姓名), "_" & "操作员姓名"
     
    If Not objCurItem Is Nothing Then
        If Val(objCurItem.结算时间) = 0 Then
            dtCurdate = zlDatabase.Currentdate
            cllTemp.Add Array("收款时间", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")), "_" & "收款时间"
        Else
            cllTemp.Add Array("收款时间", Format(objCurItem.结算时间, "yyyy-mm-dd HH:MM:SS")), "_" & "收款时间"
        End If
    End If
    cllUpdateFeeData.Add cllTemp, "_billinfo"
    
    If Not objCurItem Is Nothing Then
         '结算信息
        Set cllTemp = New Collection
        cllTemp.Add Array("结算方式", objCurItem.结算方式), "_" & "结算方式"
        cllTemp.Add Array("结算号码", objCurItem.结算号码), "_" & "结算号码"
        cllTemp.Add Array("卡类别ID", IIf(objCurItem.消费卡, 0, objCurItem.卡类别ID)), "_" & "卡类别ID"
        cllTemp.Add Array("结算卡序号", IIf(objCurItem.消费卡, objCurItem.卡类别ID, 0)), "_" & "结算卡序号"
        cllTemp.Add Array("卡号", objCurItem.卡号), "_" & "卡号"
        cllTemp.Add Array("交易流水号", objCurItem.交易流水号), "_" & "交易流水号"
        cllTemp.Add Array("交易说明", objCurItem.交易说明), "_" & "交易说明"
        cllTemp.Add Array("摘要", objCurItem.结算摘要), "_" & "摘要"
        cllTemp.Add Array("合作单位", ""), "_" & "合作单位"
        
        If Not cllExpendInfo Is Nothing Then
            cllTemp.Add Array("其他信息集", cllExpendInfo), "_" & "其他信息集"
        End If
        cllUpdateFeeData.Add cllTemp, "_balanceinfo"
    End If
    ' cllUpdateDate-修改的结算数据
    '         |--billinfo-单据信息,"_billinfo"
    '              |-预交单号,预交ID,收费单号,结帐ID,操作员编号,操作员姓名,收款时间)
    '         |--balanceinfo-结算信息,"_balanceinfo"
    '                |--(结算方式,结算号码,卡类别id,结算卡序号,卡号,交易流水号,交易说明,摘要,合作单位)
    '                |--其他信息集,
    '                |-----其他信息:交易名称,交易内容
    
    '同步状态：操作场景=2,3时：0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功"
    If int操作状态 = 0 Then
         int异常状态 = 2
        If Not GetDelErrDataToColl(objCardFeeItems.业务ID, objCardFeeItems.异常ID, cllErrData) Then Exit Function
    ElseIf int操作状态 = 1 Then
        If Not GetUpdateErrDataSyncTagToColl(objCardFeeItems.异常ID, 1, cllErrData) Then Exit Function
        int异常状态 = 1
    Else
        If Not GetUpdateErrDataSyncTagToColl(objCardFeeItems.异常ID, 3, cllErrData) Then Exit Function
        int异常状态 = 1
        '0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    If Zl_病人结算异常记录_Modify(int异常状态, cllErrData) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
        Exit Function
    End If
    
    If mobjExseSvr.Zl_Exsesvr_UpdCardFeeBlncInfo(int操作状态, cllSendCardInfo, cllUpdateFeeData, False, strErrMsg) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    UpdateCardFeeBalanceInfor = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

'Public Function Excute_DepositSaveOver(ByVal objPati As clsPatientInfo, ByVal objBalanceItems As clsBalanceItems, _
'    ByVal cllExpendInfo As Collection) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:完成预交结算
'    '入参:objBalanceItems-当前支付集
'    '返回:成功返回true,否则返回False
'    '编制:刘兴洪
'    '日期:2019-11-14 11:49:36
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cllUpdateFeeData As Collection, cllTemp As Collection
'    Dim objCurItem As clsBalanceItem, blnTrans As Boolean
'    Dim strDepositNo As String, strCardFeeNo As String, strErrMsg As String
'    Dim varTemp As Variant, lng变动ID As Long
'    Dim cllPro As Collection, strSQL As String
'    On Error GoTo errHandle
'
'    Set cllUpdateFeeData = New Collection
'    Set cllTemp = New Collection
'
'
'    Set objCurItem = objBalanceItems(1)
'
'    strCardFeeNo = ""
'
'    strDepositNo = objCurItem.单据号
'
'    cllTemp.Add Array("预交单号", strDepositNo), "_" & "预交单号"
'    cllTemp.Add Array("预交ID", objCurItem.预交ID), "_" & "预交ID"
'    cllTemp.Add Array("病人ID", objPati.病人ID), "_" & "病人ID"
'    cllTemp.Add Array("操作员编号", UserInfo.编号), "_" & "操作员编号"
'    cllTemp.Add Array("操作员姓名", UserInfo.姓名), "_" & "操作员姓名"
'    cllTemp.Add Array("收款时间", Format(objCurItem.结算时间, "yyyy-mm-dd HH:MM:SS")), "_" & "收款时间"
'    cllUpdateFeeData.Add cllTemp, "_billinfo"
'
'     '结算信息
'    Set cllTemp = New Collection
'    Set objCurItem = objBalanceItems(1)
'    cllTemp.Add Array("结算方式", objCurItem.结算方式), "_" & "结算方式"
'    cllTemp.Add Array("结算号码", objCurItem.结算号码), "_" & "结算号码"
'    cllTemp.Add Array("卡类别ID", IIf(objCurItem.消费卡, 0, objCurItem.卡类别ID)), "_" & "卡类别ID"
'    cllTemp.Add Array("结算卡序号", IIf(objCurItem.消费卡, objCurItem.卡类别ID, 0)), "_" & "结算卡序号"
'    cllTemp.Add Array("卡号", objCurItem.卡号), "_" & "卡号"
'    cllTemp.Add Array("交易流水号", objCurItem.交易流水号), "_" & "交易流水号"
'    cllTemp.Add Array("交易说明", objCurItem.交易说明), "_" & "交易说明"
'    cllTemp.Add Array("摘要", objCurItem.结算摘要), "_" & "摘要"
'    cllTemp.Add Array("合作单位", ""), "_" & "合作单位"
'
'    If Not cllExpendInfo Is Nothing Then
'        cllTemp.Add Array("其他信息集", cllExpendInfo), "_" & "其他信息集"
'    End If
'    cllUpdateFeeData.Add cllTemp, "_balanceinfo"
'    ' cllUpdateDate-修改的结算数据
'    '         |--billinfo-单据信息,"_billinfo"
'    '              |-预交单号,预交ID,收费单号,结帐ID,操作员编号,操作员姓名,收款时间)
'    '         |--balanceinfo-结算信息,"_balanceinfo"
'    '                |--(结算方式,结算号码,卡类别id,结算卡序号,卡号,交易流水号,交易说明,摘要,合作单位)
'    '                |--其他信息集,
'    '                |-----其他信息:交易名称,交易内容
'
'    blnTrans = True
'    Set cllTemp = New Collection
'    cllTemp.Add Array("异常ID", objBalanceItems.异常ID), "_异常ID"
'
'    gcnOracle.BeginTrans: blnTrans = True
'    If Zl_病人结算异常记录_Modify(2, cllTemp) = False Then
'        gcnOracle.RollbackTrans: blnTrans = True
'        Exit Function
'    End If
'
'    If mobjExseSvr.Zl_Exsesvr_Upddepositblncinfo(0, cllUpdateFeeData, False, strErrMsg) = False Then
'       gcnOracle.RollbackTrans: blnTrans = False
'       MsgBox strErrMsg, vbInformation, gstrSysName
'       Exit Function
'    End If
'    gcnOracle.CommitTrans: blnTrans = False
'
'    Excute_DepositSaveOver = True
'    Exit Function
'errHandle:
'    If blnTrans Then gcnOracle.RollbackTrans
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function


Public Function Excute_Cancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行异常作废操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-12-02 15:16:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDepositItems As clsBalanceItems, objCardFeeItems As clsBalanceItems, objTempItems As clsBalanceItems
    Dim bln相同 As Boolean, objCard As Card, strErrMsg As String, intSwapStatu As Integer, cllErrData As Collection, cllAddErrData As Collection
    Dim blnTrans As Boolean, dtCurdate As Date, cllDelFeeData As Collection
    Dim lng异常ID As Long, cllSendCardInfo As Collection, lng结帐ID As Long, lng预交ID As Long
    
    On Error GoTo errHandle
    
    dtCurdate = zlDatabase.Currentdate
    
    Set objDepositItems = New clsBalanceItems
    If Not mobjDepositItems Is Nothing Then
        If mobjDepositItems.Count <> 0 Then Set objDepositItems = mobjDepositItems.Clone
    End If
    
    Set objCardFeeItems = New clsBalanceItems
    If Not mobjCardFeeItems Is Nothing Then
       Set objCardFeeItems = mobjCardFeeItems.Clone
       
        If GetSaveSendCardInfotoCollect(mobjPati, dtCurdate, cllSendCardInfo) = False Then Exit Function
    End If
    
    bln相同 = CheckDepsoitAndCardFeePayIsSame(objDepositItems, objCardFeeItems)
    
    '先作废预交
    If objDepositItems.Count <> 0 And Not bln相同 And mobjCardFeeItems Is Nothing Then
        Set objCard = objDepositItems(1).objCard
        If objDepositItems.同步状态 >= 2 Then
            MsgBox objCard.名称 & "已经结算完成，不能进行作废操作,请操作《异常重收》功能!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        If objDepositItems.同步状态 = 1 And objDepositItems.类型 = gEM_一卡通 Then
        
            Set mobjThirdSwap.objPayCards = mobjDepositPayCards
            If mobjThirdSwap.zlThird_IsSwapIsSucces(objDepositItems, intSwapStatu, strErrMsg) = False Then
                '交易失败
                'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                If intSwapStatu = 1 Then
                    MsgBox "原" & objCard.名称 & " 交易正在进行中， 不允许作废操作!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            Else
                MsgBox "原" & objCard.名称 & " 交易已经成功， 不允许作废操作!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
     
        End If
        
        If Not GetDelErrDataToColl(objDepositItems.业务ID, objDepositItems.异常ID, cllErrData) Then Exit Function
        gcnOracle.BeginTrans: blnTrans = True
        If Zl_病人结算异常记录_Modify(2, cllErrData) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        If mobjExseSvr.Zl_Exsesvr_DelDepositErrorRec(0, objDepositItems.单据号, False) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        Excute_Cancel = True: Exit Function
    End If
    
    If Not mobjCardFeeItems Is Nothing Then
         If mobjCardFeeItems.Count = 0 Then
            If mobjCardFeeItems.类型 <> gEM_记帐单 Then Exit Function '未找到异常数据
        End If
    
        If mobjCardFeeItems.同步状态 = 1 And mobjCardFeeItems.类型 = gEM_一卡通 Then
        
            Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
            Set objCard = mobjCardFeeItems(1).objCard
            
            If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjCardFeeItems, intSwapStatu, strErrMsg) = False Then
                '交易失败
                'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                If intSwapStatu = 1 Then
                    MsgBox "原" & objCard.名称 & " 交易正在进行中， 不允许作废操作!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            Else
                MsgBox "原" & objCard.名称 & " 交易已经成功， 不允许作废操作!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
        
        If mobjCardFeeItems.同步状态 = -1 Then
            '未生成费用或预交，直接删除
            If Not GetDelErrDataToColl(objCardFeeItems.业务ID, objCardFeeItems.异常ID, cllErrData) Then Exit Function
            gcnOracle.BeginTrans: blnTrans = True
            If Zl_病人结算异常记录_Modify(2, cllErrData) = False Then
                gcnOracle.RollbackTrans: blnTrans = False: Exit Function
            End If
            
            '删除变动记录
            If mobjService.zl_PatiSvr_DelCardChangeInfo(mobjPati.病人ID, objCardFeeItems.业务ID, Val(cllSendCardInfo("_卡类别ID")(1)), CStr(cllSendCardInfo("_医疗卡号")(1))) = False Then
              gcnOracle.RollbackTrans: blnTrans = False: Exit Function
            End If
            gcnOracle.CommitTrans: blnTrans = False
            Excute_Cancel = True
            Exit Function
        End If
        
        
        
        If Not GetDelErrDataToColl(objCardFeeItems.业务ID, objCardFeeItems.异常ID, cllErrData) Then Exit Function
        If zlGetErrDataToColl(mobjPati, objCardFeeItems.业务ID, objCardFeeItems, 1, lng异常ID, dtCurdate, cllAddErrData, "", 0, objCardFeeItems.单据号, objCardFeeItems.结算金额, True) = False Then Exit Function
        '      cllDelFeeData-退费数据
        '        |-(卡费单号,预交单号,是否退卡费,是否退病历费,操作员姓名,操作员编号,退费时间,结算信息) array(名称,值) ,"_名称)
        '        |-结算信息:(退款金额,结算方式,结算号码,卡类别id,结算卡序号,支付卡号,交易流水号,交易说明,合作单位,关联交易ID) Key="_结算信息"
        
        Set cllDelFeeData = New Collection
        cllDelFeeData.Add Array("卡费单号", objCardFeeItems.单据号)
        If mobjDepositItems Is Nothing Then
            cllDelFeeData.Add Array("预交单号", "")
        Else
            cllDelFeeData.Add Array("预交单号", mobjDepositItems.单据号)
        End If
        cllDelFeeData.Add Array("是否退病历费", 1)
        cllDelFeeData.Add Array("是否退卡费", 1)
        cllDelFeeData.Add Array("操作员姓名", UserInfo.姓名)
        cllDelFeeData.Add Array("操作员编号", UserInfo.编号)
        cllDelFeeData.Add Array("退费时间", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS"))
        gcnOracle.BeginTrans: blnTrans = True
        '1.先删除原异常
         If Zl_病人结算异常记录_Modify(2, cllErrData) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
         End If
        '2.产生作废异常
        If Zl_病人结算异常记录_Modify(0, cllAddErrData) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        
        '3.删除费用及预交
        If mobjExseSvr.Zl_Exsesvr_DelCardfeeInfo(2, cllDelFeeData, lng结帐ID, lng预交ID) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        '4-删除医疗卡变动记录
        
        If Not GetDelErrDataToColl(objCardFeeItems.业务ID, lng异常ID, cllErrData) Then Exit Function
        gcnOracle.BeginTrans: blnTrans = True
        If Zl_病人结算异常记录_Modify(2, cllErrData) = False Then
             gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        
        '删除变动记录
        If mobjService.zl_PatiSvr_DelCardChangeInfo(mobjPati.病人ID, objCardFeeItems.业务ID, Val(cllSendCardInfo("_卡类别ID")(1)), CStr(cllSendCardInfo("_医疗卡号")(1))) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        Excute_Cancel = True
        Exit Function
      
    End If
    Excute_Cancel = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSaveData(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '入参:objPati-病人信息集
    '     blnNewPati-是否新病人
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 13:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDepositPrint As Boolean '是否预交打印
    Dim objDepositItems As clsBalanceItems, objCardFeeItems As clsBalanceItems, objTempItems As clsBalanceItems
    Dim dtCurdate As Date, cllSaveSendCardInfo As Collection, cllDepositAndCardFee As Collection, cllSendCardInfo As Collection
    Dim cllErrData As Collection, cllPro As Collection
    Dim lng变动id As Long, lng异常ID As Long, lng预交ID As Long, i As Long, lng结帐ID As Long
    Dim objCurItem As clsBalanceItem, objItems As clsBalanceItems
    Dim blnTrans As Boolean, dbl帐户余额 As Double, int同步状态 As Integer, int异常操作状态 As Integer
    Dim rsMoney As ADODB.Recordset
    Dim rsExpend As ADODB.Recordset, cllExpend As Collection
    Dim int状态 As Integer, blnSaveed As Boolean
    Dim int变动类型   As Integer
    Dim strDepositNo As String, intSwapStatu As Integer, strErrMsg As String
    
    
    On Error GoTo errHandle
    
    If Trim(txt预交额.Text) = "" And Trim(txt卡号.Text) = "" Then zlSaveData = True: Exit Function
    If mint操作状态 = 2 Then
        '异常作废操作
        zlSaveData = Excute_Cancel
        Exit Function
    End If
    
    dtCurdate = zlDatabase.Currentdate
    If fra磁卡.Visible And Not mblnBoundCarded And mlngCardTypeID <> 0 And Trim(txt卡号.Text) <> "" Then
         int变动类型 = GetCurCard_Statu
         If GetSaveSendCardInfotoCollect(objPati, dtCurdate, cllSendCardInfo) = False Then Exit Function
         If int变动类型 = 11 Then '绑定卡，不存在事务，直接退出
             If mobjService.zlPatisvr_SaveMedcCard(cllSendCardInfo, , True) = False Then Exit Function
             mblnBoundCarded = True
             If fra预交.Visible = False Or StrToNum(txt预交额.Text) = 0 Then zlSaveData = True: Exit Function
         End If
    End If
   
    '获取预交单据及卡费单据结算信息
    '数据组织
    If mobjDepositItems Is Nothing Then
        If GetDepositBalanceItems(dtCurdate, objDepositItems) = False Then Exit Function
    ElseIf mobjDepositItems.Count = 0 Or mobjDepositItems.是否保存 = False Then
        If GetDepositBalanceItems(dtCurdate, objDepositItems) = False Then Exit Function
    Else
        Set objDepositItems = mobjDepositItems.Clone
    End If
    
    If mobjCardFeeItems Is Nothing Then
        If GetCardFeeBalanceItems(dtCurdate, objCardFeeItems) = False Then Exit Function
    ElseIf mint操作状态 = 1 Then
        Set objCardFeeItems = mobjCardFeeItems.Clone
    ElseIf mobjCardFeeItems.是否保存 = False Then
        If GetCardFeeBalanceItems(dtCurdate, objCardFeeItems) = False Then Exit Function
    Else
        Set objCardFeeItems = mobjCardFeeItems.Clone
    End If
    
    
    mbln相同结算 = CheckDepsoitAndCardFeePayIsSame(objDepositItems, objCardFeeItems)
    If GetSaveSendCardInfotoCollect(objPati, dtCurdate, cllSaveSendCardInfo) = False Then Exit Function      '发卡服务数据
    
    If mbln相同结算 = False Or int变动类型 = 11 Then
        '非相同结算，应该分布进行结算
        '第一步:先处理预交款
        Set mobjThirdSwap.objPayCards = mobjDepositPayCards
        If objDepositItems.Count <> 0 And objDepositItems.结算完成 = False Then '未结算完成时，需要重新异常重收
            If objDepositItems.类型 = gEM_医保 And mintInsure = 0 Then
                MsgBox "当前病人非医保病人，不允许使用" & objDepositItems(1).结算方式 & "进行结算.", vbInformation
                Exit Function
            End If
            '1.产生预交异常
            Set objCurItem = objDepositItems(1)
            If objDepositItems.类型 = gEM_一卡通 And objDepositItems.同步状态 < 2 Then '已经调用接口的，不再调用
                '需要先调用检查
                intSwapStatu = 0
                If objDepositItems.同步状态 = 1 Then
                    If mobjThirdSwap.zlThird_IsSwapIsSucces(objDepositItems, intSwapStatu, strErrMsg) = False Then
                        '交易失败
                        'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                        If intSwapStatu = 1 Then
                            MsgBox "原" & mobjDepositItems(1).objCard.名称 & " 交易正在进行中， 请检查!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                            Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                            mblnDepositLocked = True: Call SetDepositEditEnabled(1)    '锁定结算方式
                            Exit Function
                        End If
                    Else
                        intSwapStatu = 1
                    End If
                End If
                If intSwapStatu = 0 Then    '只有交易失败时，才需重刷卡
                    If mobjThirdSwap.zlThird_Payment_IsValid(objPati, objCurItem, objItems, dbl帐户余额) = False Then Exit Function
                    Call objItems.CloneItemsPropertyByItems(objDepositItems)
                    Set objDepositItems = objItems
                End If
                
            ElseIf objDepositItems.类型 = gEM_消费卡 And objDepositItems.同步状态 < 2 Then
                If GetClassMoney(rsMoney) = False Then Exit Function
                If mobjThirdSwap.zlSquare_Payment_IsValid(objPati, objCurItem, objItems, dbl帐户余额, , , , rsMoney) = False Then Exit Function
                If objItems.Count > 1 Then
                    MsgBox objCurItem.objCard.名称 & "不能同时刷多张卡，请只刷一张卡进行结算!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                Call objItems.CloneItemsPropertyByItems(objDepositItems)
                 Set objDepositItems = objItems
            End If
            
            ' 0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功
            If objDepositItems.同步状态 <> 2 Then
                
                'int类型-0-仅卡费;1-仅预交,2-卡费及预交
                If GetAddDepositAndCardFeeDataToCollect(1, objPati, Nothing, objDepositItems, dtCurdate, cllDepositAndCardFee) = False Then Exit Function
                lng预交ID = objDepositItems(1).预交ID
                If objDepositItems.同步状态 = 1 Then   '已经调用接口的，直接删除
                    If GetUpdateErrDataSyncTagToColl(objDepositItems.异常ID, 1, cllErrData) = False Then Exit Function
                    int异常操作状态 = 1
                Else
                    If zlGetErrDataToColl(objPati, lng预交ID, objDepositItems, 1, objDepositItems.异常ID, dtCurdate, cllErrData, objCurItem.单据号, objCurItem.结算金额) = False Then Exit Function
                    int异常操作状态 = IIf(objDepositItems.是否保存, 1, 0)
                    '0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
                End If
                
                int状态 = IIf(objDepositItems.类型 = gEM_一卡通 Or objDepositItems.类型 = gEM_医保, 1, 0)
            
                '------------------------------------------------------------------------------------------------------
                '2.开始数据保存存
                If objDepositItems.是否保存 = False Then
                    
                    gcnOracle.BeginTrans: blnTrans = True
                    If objDepositItems.类型 = gEM_一卡通 Or objDepositItems.类型 = gEM_医保 Then '只有三方卡及医保，才涉及异常
                        If Zl_病人结算异常记录_Modify(int异常操作状态, cllErrData) = False Then
                            gcnOracle.RollbackTrans: blnTrans = False
                            Exit Function
                        End If
                    End If
                    
                     '  2.1 产生预交数据:操作状态:0-正常的预交款 ;1-保存为未生效的预交款
                    If mobjExseSvr.zl_ExseSvr_AddDepositInfo(int状态, cllDepositAndCardFee, lng预交ID) = False Then
                         gcnOracle.RollbackTrans: Exit Function
                    End If
                
                    gcnOracle.CommitTrans: blnTrans = False: mblnDepositLocked = True
                    
                    Call SetDepositEditEnabled(1)   '锁定结算信息
                    '------------------------------------------------------------------------------------------------------
                    objDepositItems.是否保存 = True: objDepositItems.同步状态 = 1
                    For i = 1 To objDepositItems.Count
                        objDepositItems(i).是否保存 = True
                    Next
               End If
            End If
         
            Set mobjDepositItems = objDepositItems
            '3.如果是三方卡需要保存
            If objDepositItems.类型 = gEM_一卡通 Then
                '一卡通扣款
                If objDepositItems.同步状态 <> 2 Then
                    Set objCurItem = objDepositItems(1)
                    If mobjThirdSwap.zlThird_Payment(objCurItem.objCard, objPati, cllPro, objDepositItems, objItems, rsExpend, blnSaveed) = False Then
                        If blnSaveed Then
                            Call objItems.CloneItemsPropertyByItems(objDepositItems)
                            If Not objItems Is Nothing Then
                                If objItems.Count > 0 Then Set mobjDepositItems = objItems
                            End If
                        End If
                        Exit Function
                    End If
                    If objItems.Count > 1 Then
                        MsgBox "预交款，不支持多种结算方式，请检查", vbInformation + vbOKOnly, Me.Caption
                        Exit Function
                    End If
                                    
                                    
                    Call objItems.CloneItemsPropertyByItems(objDepositItems)
                    Set objDepositItems = objItems
                    Set mobjDepositItems = objDepositItems
                    mobjDepositItems.同步状态 = 2 '接口已经调用完成
                    '完成结算
                    Call mobjThirdSwap.zlGetThreeSwapExpendToCollByRecords(rsExpend, cllExpend)
                Else
                    Set cllExpend = Nothing
                End If
                
                If Not mblnDepositLocked Then
                    mblnDepositLocked = True: Call SetDepositEditEnabled(1) '锁定结算信息
                End If
                    
                '修正预交结算信息
                If UpdateDepositBlncInfo(0, objPati, objDepositItems, cllExpend) = False Then Exit Function
                objDepositItems.结算完成 = True
                
                If Not mblnDepositLocked Then
                    mblnDepositLocked = True: Call SetDepositEditEnabled(1) '锁定输结算信息
                End If
                
            ElseIf objDepositItems.类型 = gEM_医保 Then
                '医保结算
                If objDepositItems.同步状态 <> 2 Or mint操作状态 = 1 Then
                    
                    '更新同步标志 '同步状态：操作场景=2,3时：0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功"
                    If Not GetUpdateErrDataSyncTagToColl(objDepositItems.异常ID, 2, cllErrData) Then Exit Function
                    gcnOracle.BeginTrans: blnTrans = True
                    'int操作状态-操作状态:0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
                    If Zl_病人结算异常记录_Modify(1, cllErrData) = False Then
                        gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                    End If
                
                    If Not gclsInsure.TransferSwap(objDepositItems(1).预交ID, objDepositItems.结算金额, mintInsure) Then
                        gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                    End If
                    gcnOracle.CommitTrans
                    mobjDepositItems.同步状态 = 2 '接口已经调用完成
                     
                    If UpdateDepositBlncInfo(0, objPati, objDepositItems, cllExpend) = False Then Exit Function
                    objDepositItems.结算完成 = True
                    If Not mblnDepositLocked Then
                        mblnDepositLocked = True: Call SetDepositEditEnabled(1) '锁定结算信息
                    End If
                
                Else
                     Set cllExpend = Nothing
                End If
            Else
                '其他无处理
                 objDepositItems.结算完成 = True
            End If
            For i = 1 To objDepositItems.Count
                objDepositItems(i).是否结算 = True
                objDepositItems(i).是否允许编辑 = False
                objDepositItems(i).是否允许删除 = False
                objDepositItems(i).是否允许退现 = False
            Next
            Set mobjDepositItems = objDepositItems
        End If
        
        If int变动类型 = 11 Then zlSaveData = True: Exit Function
        
        
        '第二步:再处理发卡数据
        If fra磁卡.Visible And objCardFeeItems.结算完成 = False And Trim(txt卡号.Text) <> "" Then
            Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
            int状态 = GetCurCard_Statu
            If int状态 = 11 Then zlSaveData = True: Exit Function
            
            If GetSaveSendCardInfotoCollect(objPati, dtCurdate, cllSendCardInfo) = False Then Exit Function
            If objCardFeeItems.是否保存 = False Or objCardFeeItems.业务ID = 0 Then
                If mobjService.zlPatiSvr_GetNextID("病人医疗卡变动", lng变动id) = False Then Exit Function
                objCardFeeItems.业务ID = lng变动id
            Else
                lng变动id = objCardFeeItems.业务ID
                lng异常ID = objCardFeeItems.异常ID
            End If
            
            ' 0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功
            'int类型-0-仅卡费;1-仅预交,2-卡费及预交
            If objCardFeeItems.类型 = gEM_一卡通 And objCardFeeItems.同步状态 < 2 Then
                '需要先调用检查
                 Set objCurItem = objCardFeeItems(1)
                 
                intSwapStatu = 0
                If objCardFeeItems.同步状态 = 1 Then
                    If mobjThirdSwap.zlThird_IsSwapIsSucces(objCardFeeItems, intSwapStatu, strErrMsg) = False Then
                        '交易失败
                        'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                        If intSwapStatu = 1 Then
                            MsgBox "原" & objCardFeeItems(1).objCard.名称 & " 交易正在进行中， 请检查!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                            Call SetLoaclePayModefromCard(objCardFeeItems(1).objCard, False, True)
                            mblnDepositLocked = True: Call SetDepositEditEnabled(1)    '锁定结算方式
                            Exit Function
                        End If
                    Else
                        intSwapStatu = 1
                    End If
                End If
                If intSwapStatu = 0 Then    '只有交易失败时，才需重刷卡
                    If mobjThirdSwap.zlThird_Payment_IsValid(objPati, objCurItem, objItems, dbl帐户余额) = False Then Exit Function
                    Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                    Set objCardFeeItems = objItems
                End If
                
            ElseIf objCardFeeItems.类型 = gEM_消费卡 And objCardFeeItems.同步状态 < 2 Then
                If GetClassMoney(rsMoney) = False Then Exit Function
                If mobjThirdSwap.zlSquare_Payment_IsValid(objPati, objCurItem, objItems, dbl帐户余额, , , , rsMoney) = False Then Exit Function
                
                Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                Set objCardFeeItems = objItems
            Else
                '其他
            End If
            
            ' 0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功
            If objCardFeeItems.同步状态 = 0 And objCardFeeItems.是否保存 = False Then '未产生变动记录记录
                  
                If objCardFeeItems.同步状态 = 1 Then   '已经调用接口的，直接删除
                    If GetUpdateErrDataSyncTagToColl(objCardFeeItems.异常ID, 1, cllErrData) = False Then Exit Function
                    int异常操作状态 = 1
                Else
                    If zlGetErrDataToColl(objPati, lng变动id, objCardFeeItems, -1, lng异常ID, dtCurdate, cllErrData, "", 0, objCardFeeItems.单据号, objCardFeeItems.结算金额) = False Then Exit Function
                    int异常操作状态 = IIf(objCardFeeItems.是否保存, 1, 0)
                    '0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
                End If
                                
                                
                '------------------------------------------------------------------------------------------------------
                '数据保存
                '1.保存异常数据及变动记录
                gcnOracle.BeginTrans: blnTrans = True
                If Zl_病人结算异常记录_Modify(int异常操作状态, cllErrData) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False
                    Exit Function
                End If
                ' int操作状态:0-正常记录;1-产生异常数据;2-只产生变动记录
                If mobjService.zlPatisvr_SaveMedcCard(cllSendCardInfo, , , 2, lng变动id) = False Then
                   gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                End If
                gcnOracle.CommitTrans: blnTrans = False
                objCardFeeItems.是否保存 = True
                objCardFeeItems.异常ID = lng异常ID
                objCardFeeItems.业务ID = lng变动id
                objCardFeeItems.同步状态 = -1
                For i = 1 To objCardFeeItems.Count
                    objCardFeeItems(i).是否保存 = True
                    objCardFeeItems(i).异常ID = lng异常ID
                Next
                Set mobjCardFeeItems = objCardFeeItems
                '------------------------------------------------------------------------------------------------------
            Else
                lng变动id = objCardFeeItems.业务ID
                lng异常ID = objCardFeeItems.异常ID
            End If
            
            '2.增加卡费费用数据
            '操作状态:0-正常的预交款或卡费缴款;1-保存为未生效的预交款或异常的卡费;2-保存为记帐单;3-保存为划价单
            If objCardFeeItems.同步状态 = -1 Then
     
                If GetAddDepositAndCardFeeDataToCollect(0, objPati, objCardFeeItems, Nothing, dtCurdate, cllDepositAndCardFee) = False Then Exit Function
                           
                '同步状态：操作场景=2,3时：0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功"
                If GetUpdateErrDataSyncTagToColl(lng异常ID, IIf(objCardFeeItems.类型 = gEM_记帐单, 3, 1), cllErrData) = False Then Exit Function
                gcnOracle.BeginTrans
                
                If Zl_病人结算异常记录_Modify(1, cllErrData) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                End If
                
                int状态 = IIf(objCardFeeItems.类型 = gEM_记帐单, 2, 1)
                If mobjExseSvr.Zl_Exsesvr_AddCardFeeInfo(int状态, cllDepositAndCardFee, lng结帐ID, lng预交ID, True) = False Then
                    '需要删除变动记录及异常记录
                    If GetDelErrDataToColl(lng变动id, lng异常ID, cllErrData) = False Then
                        gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                        Exit Function
                    End If
                    If Zl_病人结算异常记录_Modify(2, cllErrData) = False Then
                          gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                    End If
                    
                    '删除变动记录
                    If mobjService.zl_PatiSvr_DelCardChangeInfo(objPati.病人ID, lng变动id, CLng(cllSendCardInfo("_卡类别ID")(1)), cllSendCardInfo("_医疗卡号")(1), True) = False Then
                       gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                    End If
                    gcnOracle.CommitTrans: blnTrans = False: Exit Function
                    Exit Function
                End If
                gcnOracle.CommitTrans: blnTrans = False
                objCardFeeItems.同步状态 = IIf(objCardFeeItems.类型 = gEM_记帐单, 3, 1)
                For i = 1 To objCardFeeItems.Count
                    objCardFeeItems(i).结算ID = lng结帐ID
                Next
                Set mobjCardFeeItems = objCardFeeItems
            End If
            
            '------------------------------------------------------------------------------------------------------
            '3.一卡能等相关结算数据
            If objCardFeeItems.类型 = gEM_一卡通 Then
                '一卡通扣款
                If objCardFeeItems.同步状态 < 2 Then
                    If mobjThirdSwap.zlThird_Payment(objCurItem.objCard, objPati, cllPro, objCardFeeItems, objItems, rsExpend, blnSaveed) = False Then
                        If blnSaveed Then
                            Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                            If Not objItems Is Nothing Then
                                If objItems.Count > 0 Then
                                    Set objCardFeeItems = objItems
                                    Set mobjCardFeeItems = objCardFeeItems
                                End If
                            End If
                        End If
                        Exit Function
                    End If
                    If objItems.Count > 1 Then
                        MsgBox "一卡通卡费，不支持多种结算方式，请检查", vbInformation + vbOKOnly, Me.Caption
                        Exit Function
                    End If
                    
                    Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                    Set objCardFeeItems = objItems
                            
                    objCardFeeItems.同步状态 = 2
                    Set mobjCardFeeItems = objCardFeeItems
                Else
                    Set mobjCardFeeItems = objCardFeeItems
                End If
                Call mobjThirdSwap.zlGetThreeSwapExpendToCollByRecords(rsExpend, cllExpend)
                'int操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
                If objCardFeeItems.同步状态 < 3 Then
                    If UpdateCardFeeBalanceInfor(2, objPati, cllSendCardInfo, objCardFeeItems, Nothing, cllExpend) = False Then Exit Function
                    objCardFeeItems.同步状态 = 3 '费用结算修正
                    Set mobjCardFeeItems = objCardFeeItems
                End If
                If Not mblnSendCardLocked Then
                    mblnSendCardLocked = True:  Call SetCardEditEnabled(1)  '锁定结算信息
                End If
            ElseIf objCardFeeItems.类型 = gEM_医保 Then    '卡费无医保
                 '医保结算
            Else
                '其他无处理
               
            End If
            
            If objCardFeeItems.同步状态 <= 3 Then
                '4.医疗卡发卡
                '同步状态：操作场景=2,3时：0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功"
                If Not GetUpdateErrDataSyncTagToColl(lng异常ID, 4, cllErrData, cllSendCardInfo) Then Exit Function
                gcnOracle.BeginTrans: blnTrans = True
                'int操作状态-操作状态:0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
                If Zl_病人结算异常记录_Modify(IIf(objCardFeeItems.类型 = gEM_记帐单, 2, 1), cllErrData) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                End If
                If mobjService.zl_PatiSvr_ConfirmCardChange(objPati.病人ID, lng变动id, False, cllSendCardInfo) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                End If
                gcnOracle.CommitTrans: blnTrans = False:
                objCardFeeItems.同步状态 = 4
                If objCardFeeItems.类型 = gEM_记帐单 Then
                    objCardFeeItems.结算完成 = True
                End If
                Set mobjCardFeeItems = objCardFeeItems
                
                If Not mblnSendCardLocked Then
                    mblnSendCardLocked = True:     Call SetCardEditEnabled(1)  '锁定结算信息
                End If
            End If
            
            '5.卡费确认
            'int操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
            If objCardFeeItems.类型 <> gEM_记帐单 Then
                If UpdateCardFeeBalanceInfor(0, objPati, cllSendCardInfo, objCardFeeItems, Nothing, Nothing) = False Then Exit Function
            End If
            If Not mblnSendCardLocked Then
                mblnSendCardLocked = True:  Call SetCardEditEnabled(1)  '锁定结算信息
            End If
            mobjCardFeeItems.结算完成 = True
        End If
        zlSaveData = True
        Exit Function
    End If
    
    '二、卡费及预交同种结算方式收取
    If objCardFeeItems.是否保存 = False Or objCardFeeItems.业务ID = 0 Then
        If mobjService.zlPatiSvr_GetNextID("病人医疗卡变动", lng变动id) = False Then Exit Function
        objCardFeeItems.业务ID = lng变动id
        If Not objDepositItems Is Nothing Then
            objDepositItems.业务ID = lng变动id
        End If
    Else
        lng变动id = objCardFeeItems.业务ID
        lng异常ID = objCardFeeItems.异常ID
    End If
    Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
    ' 0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功
    'int类型-0-仅卡费;1-仅预交,2-卡费及预交
    If objCardFeeItems.类型 = gEM_一卡通 And objCardFeeItems.同步状态 < 2 Then
        '需要先调用检查
        intSwapStatu = 0
        If objCardFeeItems.同步状态 = 1 Then
            Set objItems = objCardFeeItems.Clone
            objItems.结算金额 = objItems.结算金额 + mobjDepositItems.结算金额
            objItems(1).结算金额 = RoundEx(objItems(1).结算金额 + mobjDepositItems.结算金额, 6)
            
            If mobjThirdSwap.zlThird_IsSwapIsSucces(objItems, intSwapStatu, strErrMsg, mobjDepositItems(1).预交ID) = False Then
                '交易失败
                'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                If intSwapStatu = 1 Then
                    MsgBox "原" & objCardFeeItems(1).objCard.名称 & " 交易正在进行中， 请检查!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Call SetLoaclePayModefromCard(objCardFeeItems(1).objCard, True, True)
                    Call SetLoaclePayModefromCard(objCardFeeItems(1).objCard, False, True)
                    mblnSendCardLocked = True: mblnDepositLocked = True
                    Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '锁定结算方式
                    
                    Exit Function
                End If
            Else
                intSwapStatu = 1
            End If
        End If
        
        If intSwapStatu = 0 Then    '只有交易失败时，才需重刷卡
            Set objCurItem = objCardFeeItems(1).Clone
            objCurItem.结算金额 = objCardFeeItems.结算金额 + objDepositItems.结算金额 '支付金额总额
            If mobjThirdSwap.zlThird_Payment_IsValid(objPati, objCurItem, objItems, dbl帐户余额) = False Then Exit Function
            If objItems.结算金额 <> RoundEx(objCardFeeItems.结算金额 + objDepositItems.结算金额, 5) Then
                MsgBox objCurItem.objCard.名称 & "返回的有效金额与本次要结算的金额不一致，可能是因为余额不足造成，请核查!" & vbCrLf & _
                    "  返回金额:" & Format(RoundEx(objItems.结算金额, 5), "####0.00;-####0.00;0.00;0.00") & vbCrLf & _
                    "  本次结算:" & Format(RoundEx(objCardFeeItems.结算金额 + objDepositItems.结算金额, 5), "####0.00;-####0.00;0.00;0.00"), vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            For i = 1 To objCardFeeItems.Count
                  Set objCardFeeItems(i).objCard = objItems(1).objCard
                  objCardFeeItems(i).结算方式 = objItems(1).结算方式
                  objCardFeeItems(i).结算号码 = objItems(1).结算号码
                  objCardFeeItems(i).结算类型 = objItems(1).结算类型
                  objCardFeeItems(i).卡号 = objItems(1).卡号
                  objCardFeeItems(i).卡类别ID = objItems(1).卡类别ID
                  objCardFeeItems(i).密码 = objItems(1).密码
            Next
            If Not objDepositItems Is Nothing Then
                For i = 1 To objDepositItems.Count
                      Set objDepositItems(i).objCard = objItems(1).objCard
                      objDepositItems(i).结算方式 = objItems(1).结算方式
                      objDepositItems(i).结算号码 = objItems(1).结算号码
                      objDepositItems(i).结算类型 = objItems(1).结算类型
                      objDepositItems(i).卡号 = objItems(1).卡号
                      objDepositItems(i).卡类别ID = objItems(1).卡类别ID
                      objDepositItems(i).密码 = objItems(1).密码
                Next
            End If
        End If
    ElseIf objCardFeeItems.类型 = gEM_消费卡 And objCardFeeItems.同步状态 < 2 Then
        If GetClassMoney(rsMoney) = False Then Exit Function
        
        Set objCurItem = objCardFeeItems(1).Clone
        objCurItem.结算金额 = objCardFeeItems.结算金额 + objDepositItems.结算金额 '支付金额总额
        If mobjThirdSwap.zlSquare_Payment_IsValid(objPati, objCurItem, objItems, dbl帐户余额, , , , rsMoney) = False Then Exit Function
        
        If objItems.结算金额 <> RoundEx(objCardFeeItems.结算金额 + objDepositItems.结算金额, 5) Then
            MsgBox objCurItem.objCard.名称 & "返回的有效金额与本次要结算的金额不一致，可能是因为余额不足造成，请核查!" & vbCrLf & _
                "  返回金额:" & Format(RoundEx(objItems.结算金额, 5), "####0.00;-####0.00;0.00;0.00") & vbCrLf & _
                "  本次结算:" & Format(RoundEx(objCardFeeItems.结算金额 + objDepositItems.结算金额, 5), "####0.00;-####0.00;0.00;0.00"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If objItems.Count > 1 Then
             MsgBox objCurItem.objCard.名称 & "不能同时刷多张卡，请只刷一张卡进行结算!", vbInformation + vbOKOnly, gstrSysName
             Exit Function
        End If
        
        For i = 1 To objCardFeeItems.Count
              Set objCardFeeItems(i).objCard = objItems(1).objCard
              objCardFeeItems(i).结算方式 = objItems(1).结算方式
              objCardFeeItems(i).结算号码 = objItems(1).结算号码
              objCardFeeItems(i).结算类型 = objItems(1).结算类型
              objCardFeeItems(i).卡号 = objItems(1).卡号
              objCardFeeItems(i).卡类别ID = objItems(1).卡类别ID
              objCardFeeItems(i).密码 = objItems(1).密码
        Next
        If Not objDepositItems Is Nothing Then
            For i = 1 To objDepositItems.Count
                  Set objDepositItems(i).objCard = objItems(1).objCard
                  objDepositItems(i).结算方式 = objItems(1).结算方式
                  objDepositItems(i).结算号码 = objItems(1).结算号码
                  objDepositItems(i).结算类型 = objItems(1).结算类型
                  objDepositItems(i).卡号 = objItems(1).卡号
                  objDepositItems(i).卡类别ID = objItems(1).卡类别ID
                  objDepositItems(i).密码 = objItems(1).密码
            Next
        End If
    Else
        '其他结算
    End If
        
    If objCardFeeItems.同步状态 = 0 And objCardFeeItems.是否保存 = False Then   '未产生费用
                    
        If objCardFeeItems.同步状态 = 1 Then   '已经调用接口的，直接删除
            If GetUpdateErrDataSyncTagToColl(objCardFeeItems.异常ID, 1, cllErrData) = False Then Exit Function
            int异常操作状态 = 1
        Else
            If zlGetErrDataToColl(objPati, lng变动id, objCardFeeItems, -1, lng异常ID, dtCurdate, cllErrData, objDepositItems.单据号, objDepositItems.结算金额, objCardFeeItems.单据号, objCardFeeItems.结算金额) = False Then Exit Function
            int异常操作状态 = IIf(objCardFeeItems.是否保存, 1, 0)
            '0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
        End If
            
        '------------------------------------------------------------------------------------------------------
        '数据保存
        '1.保存异常数据及变动记录
        gcnOracle.BeginTrans: blnTrans = True
        If Zl_病人结算异常记录_Modify(int异常操作状态, cllErrData) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            Exit Function
        End If
        
        If mobjService.zlPatisvr_SaveMedcCard(cllSendCardInfo, , , 2, lng变动id) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        
        gcnOracle.CommitTrans: blnTrans = False
        objCardFeeItems.是否保存 = True
        objCardFeeItems.异常ID = lng异常ID
        objCardFeeItems.业务ID = lng变动id
        objCardFeeItems.同步状态 = -1
        For i = 1 To objCardFeeItems.Count
            objCardFeeItems(i).是否保存 = True
            objCardFeeItems(i).异常ID = lng异常ID
        Next
        objDepositItems.是否保存 = True
        objDepositItems.异常ID = lng异常ID
        objDepositItems.业务ID = lng变动id
        objDepositItems.同步状态 = -1
        For i = 1 To objDepositItems.Count
            objDepositItems(i).是否保存 = True
            objDepositItems(i).异常ID = lng异常ID
        Next
        
        Set mobjCardFeeItems = objCardFeeItems
        Set mobjDepositItems = objDepositItems
        '------------------------------------------------------------------------------------------------------
    Else
        lng变动id = objCardFeeItems.业务ID
        lng异常ID = objCardFeeItems.异常ID
    End If
    
    '2.增加卡费费用数据
    '操作状态:0-正常的预交款或卡费缴款;1-保存为未生效的预交款或异常的卡费;2-保存为记帐单;3-保存为划价单
    If objCardFeeItems.同步状态 = -1 Then
        'int类型-0-仅卡费;1-仅预交,2-卡费及预交
        If GetAddDepositAndCardFeeDataToCollect(2, objPati, objCardFeeItems, objDepositItems, dtCurdate, cllDepositAndCardFee) = False Then Exit Function
        lng预交ID = objDepositItems(1).预交ID
        If GetUpdateErrDataSyncTagToColl(lng异常ID, 1, cllErrData) = False Then Exit Function
        gcnOracle.BeginTrans
        blnTrans = True:
        If Zl_病人结算异常记录_Modify(1, cllErrData) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        
        int状态 = IIf(objCardFeeItems.类型 = gEM_记帐单, 2, 1)
        If mobjExseSvr.Zl_Exsesvr_AddCardFeeInfo(int状态, cllDepositAndCardFee, lng结帐ID, lng预交ID, True) = False Then
            '需要删除变动记录及异常记录
            If GetDelErrDataToColl(lng变动id, lng异常ID, cllErrData) = False Then
                gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                Exit Function
            End If
            If Zl_病人结算异常记录_Modify(2, cllErrData) = False Then
                  gcnOracle.RollbackTrans: blnTrans = False: Exit Function
            End If
            
            '删除变动记录
            If mobjService.zl_PatiSvr_DelCardChangeInfo(objPati.病人ID, lng变动id, CLng(cllSendCardInfo("_卡类别ID")(1)), cllSendCardInfo("_医疗卡号")(1), True) = False Then
               gcnOracle.RollbackTrans: blnTrans = False: Exit Function
            End If
            gcnOracle.CommitTrans: blnTrans = False: Exit Function
            Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        objCardFeeItems.同步状态 = 1
        objDepositItems.同步状态 = 1
        For i = 1 To objCardFeeItems.Count
            objCardFeeItems(i).结算ID = lng结帐ID
        Next
        For i = 1 To objDepositItems.Count
            objDepositItems(i).结算ID = lng结帐ID
            objDepositItems(i).预交ID = lng预交ID
        Next
        Set mobjCardFeeItems = objCardFeeItems
        Set mobjDepositItems = objDepositItems
    End If
            
    '------------------------------------------------------------------------------------------------------
    '3.一卡通等相关结算数据
    If objCardFeeItems.类型 = gEM_一卡通 Then
        '一卡通扣款
        If objCardFeeItems.同步状态 < 2 Then
            Set objItems = objCardFeeItems.Clone
            If Not objDepositItems Is Nothing Then
                objItems(1).结算金额 = objItems(1).结算金额 + objDepositItems.结算金额
                objItems.结算金额 = objItems.结算金额 + objDepositItems.结算金额
                strDepositNo = objDepositItems.单据号
            End If
            Set objCurItem = objCardFeeItems(1)
            
             If mobjThirdSwap.zlThird_Payment(objCurItem.objCard, objPati, cllPro, objItems, objTempItems, rsExpend, blnSaveed, strDepositNo) = False Then
                If objTempItems Is Nothing Then
                    MsgBox "调用三方接口支付失败，请检查!", vbInformation, gstrSysName
                    Exit Function
                End If
                If objTempItems.Count = 0 Then
                    MsgBox "调用三方接口支付失败，请检查!", vbInformation, gstrSysName
                    Exit Function
                End If
                If objTempItems.Count > 1 Then
                    MsgBox "卡费及预交款暂不支持多种结算方式，请检查!", vbInformation, gstrSysName
                    Exit Function
                End If
                Set objItems = objTempItems.Clone
                
                Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                
                objItems.结算金额 = objCardFeeItems.结算金额
                objItems(1).结算金额 = objCardFeeItems(1).结算金额
                Set objCardFeeItems = objItems
                
                Set objItems = objItems.Clone
                
                objItems.结算金额 = objDepositItems.结算金额
                objItems(1).结算金额 = objDepositItems(1).结算金额
                Set objDepositItems = objItems
                Set mobjCardFeeItems = objCardFeeItems
                Set mobjDepositItems = objDepositItems
                Exit Function
            End If
            
            If RoundEx(objItems.结算金额, 2) <> RoundEx(objTempItems.结算金额, 2) Then
                MsgBox "当前支付总额与本次支付的总额不一致，请检查!", vbInformation, gstrSysName
                Exit Function
            End If
            If objTempItems.Count > 1 Then
                MsgBox "一卡通卡费或预交，不支持多种结算方式，请检查", vbInformation + vbOKOnly, Me.Caption
                Exit Function
            End If
            Set objItems = objTempItems.Clone
            Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
            
            objItems.结算金额 = objCardFeeItems.结算金额
            objItems(1).结算金额 = objCardFeeItems(1).结算金额
            Set objCardFeeItems = objItems
            
            Set objItems = objItems.Clone
            
            Call objItems.CloneItemsPropertyByItems(objDepositItems)
            
            objItems.结算金额 = objDepositItems.结算金额
            objItems(1).结算金额 = objDepositItems(1).结算金额
            objItems(1).单据号 = objDepositItems(1).单据号
            objItems(1).预交ID = objDepositItems(1).预交ID
            objItems(1).是否预交 = objDepositItems(1).是否预交
            
            
            Set objDepositItems = objItems
            '同步状态：操作场景=2,3时：0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功"
            objCardFeeItems.同步状态 = 2
            objDepositItems.同步状态 = 2
            Set mobjCardFeeItems = objCardFeeItems
            Set mobjDepositItems = objDepositItems
            If Not mblnSendCardLocked Then
                mblnSendCardLocked = True: mblnDepositLocked = True
                Call SetCardEditEnabled(1)  '锁定结算信息
                Call SetDepositEditEnabled(1) '锁定结算信息
            End If
            Call mobjThirdSwap.zlGetThreeSwapExpendToCollByRecords(rsExpend, cllExpend)
        Else
            Set mobjCardFeeItems = objCardFeeItems
            Set mobjDepositItems = objDepositItems
            Set cllExpend = Nothing
            If Not mobjCardFeeItems Is Nothing Then Set cllExpend = objCardFeeItems.objTag
            If Not mobjCardFeeItems Is Nothing And cllExpend Is Nothing Then Set cllExpend = objCardFeeItems.objTag
            
        End If
        
        
        'int操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
        If objCardFeeItems.同步状态 <= 3 Then
            If UpdateCardFeeBalanceInfor(2, objPati, cllSendCardInfo, objCardFeeItems, objDepositItems, cllExpend) = False Then Exit Function
            mobjCardFeeItems.同步状态 = 3 '费用结算修正
            mobjDepositItems.同步状态 = 3 '费用结算修正
        End If
        If Not mblnSendCardLocked Then
            mblnSendCardLocked = True: mblnDepositLocked = True
            Call SetCardEditEnabled(1)  '锁定结算信息
            Call SetDepositEditEnabled(1)  '锁定结算信息
        End If
        
    ElseIf objDepositItems.类型 = gEM_医保 Then
         '医保结算
    Else
        '其他无处理
          
    End If
    
    If objCardFeeItems.同步状态 <= 3 Then
        '4.医疗卡发卡
        '同步状态：操作场景=2,3时：0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功"
        If Not GetUpdateErrDataSyncTagToColl(lng异常ID, 4, cllErrData) Then Exit Function
        gcnOracle.BeginTrans: blnTrans = True
        'int操作状态-操作状态:0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
        If Zl_病人结算异常记录_Modify(1, cllErrData) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        If mobjService.zl_PatiSvr_ConfirmCardChange(objPati.病人ID, lng变动id, False, cllSendCardInfo) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        If Not mblnSendCardLocked Then
            mblnSendCardLocked = True: mblnDepositLocked = True
            Call SetCardEditEnabled: Call SetDepositEditEnabled(1) '锁定结算信息
        End If
    End If
    '5.卡费及预交确认
    'int操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
    If UpdateCardFeeBalanceInfor(0, objPati, cllSendCardInfo, objCardFeeItems, objDepositItems, Nothing) = False Then Exit Function
    Set mobjCardFeeItems = objCardFeeItems
    Set mobjDepositItems = objDepositItems
    mobjCardFeeItems.结算完成 = True
    mobjDepositItems.结算完成 = True
    
    zlSaveData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetDepositPayCard(Optional ByVal intIndex As Integer = -1) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前支付的卡对象
    '入参:intIndex-当前支付的索引:-1表示只选择当前选择的支付卡类别
    '返回:返回卡对象
    '编制:刘兴洪
    '日期:2019-11-06 19:21:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If intIndex = -1 Then
        If cbo预交结算.ListIndex < 0 Then Exit Function
        intIndex = cbo预交结算.ListIndex
    End If
    Set GetDepositPayCard = mobjDepositPayCards(intIndex + 1)
    Exit Function
errHandle:
    Set GetDepositPayCard = Nothing
End Function
Private Function GetCardFeePayCard(Optional ByVal intIndex As Integer = -1) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前支付的卡对象
    '入参:intIndex-当前支付的索引:-1表示只选择当前选择的支付卡类别
    '返回:返回卡对象
    '编制:刘兴洪
    '日期:2019-11-06 19:21:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If intIndex = -1 Then
        If cbo发卡结算.ListIndex < 0 Then Exit Function
        intIndex = cbo发卡结算.ListIndex
    End If
    Set GetCardFeePayCard = mobjCardFeePayCards(intIndex + 1)
    Exit Function
errHandle:
    Set GetCardFeePayCard = Nothing
End Function

Public Function zlSaveDataAfter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存后执行
    '入参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 15:08:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card
            
    On Error GoTo errHandle
     
     If Not mobjDepositItems Is Nothing And mblnDepositPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mobjDepositItems.单据号, "收款时间=" & Format(mobjDepositItems(1).结算时间, "yyyy-mm-dd HH:MM:SS"), _
                                "病人ID=" & mobjPati.病人ID, IIf(mobjDepositFact.打印格式 = 0, "", "ReportFormat=" & mobjDepositFact.打印格式), 2)
            
            If mobjDepositFact.严格控制 = False Then
                zlDatabase.SetPara "当前预交票据号", txtFact.Text, glngSys, mlngModule
            End If
     End If
     
    If mbln相同结算 And Trim(txtFact.Text) <> "" And Not mobjDepositItems Is Nothing Then
        Call mobjExseSvr.Zl_Exsesvr_Updatedepositinvinf(mobjDepositItems.单据号, mobjDepositFact.领用ID, txtFact.Text, UserInfo.姓名)
    End If
    mblnSendCardLocked = False
    mbln相同结算 = False
    Call SetDepositEditEnabled
    Call SetCardEditEnabled
    Call RefreshFactNo
    '就诊卡领用检查
    Set objSendCard = mCurSendCard.objSendCard
    If Not objSendCard Is Nothing Then
        If objSendCard.是否严格控制 Then
            mCurSendCard.lng领用ID = mobjExseSvr.CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), , objSendCard.接口序号)
            If mCurSendCard.lng领用ID <= 0 Then
                Select Case mCurSendCard.lng领用ID
                    Case 0 '操作失败
                    Case -1
                        If txt卡号.Text <> "" Then MsgBox "你已没有自用及共用的" & objSendCard.名称 & "卡,不能再发放！" & vbCrLf & _
                            "请先在本地设置共用批次或领用一批新卡！", vbExclamation, gstrSysName
                    Case -2
                        If txt卡号.Text <> "" Then MsgBox "本地共用的" & objSendCard.名称 & "卡已用完,你不能再发放！" & vbCrLf & _
                            "请重新设置本地共用卡批次或领用一批新卡！", vbExclamation, gstrSysName
                End Select
            End If
        End If
        '写卡操作
        If fra磁卡.Visible And objSendCard.是否写卡 Then Call WriteCard(mobjPati.病人ID, objSendCard)
    End If
    Set mobjDepositItems = Nothing
    Set mobjCardFeeItems = Nothing
    zlSaveDataAfter = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function WriteCard(lng病人ID As Long, objSendCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写卡
    '入参:lng病人ID - 病人ID
    '编制:王吉
    '问题:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    If mobjOneCardComLib Is Nothing Then Exit Function
    WriteCard = mobjOneCardComLib.zlBandCardArfter(Me, mlngModule, objSendCard.接口序号, lng病人ID, strExpend)
    Exit Function
ErrHandl:
    WriteCard = False
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

 
Public Function zlSaveDataBeforCheckIsValid(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo, _
    Optional ByVal bln是否自动识别的身份证 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '入参:objPati-病人信息集
    '     blnNewPati-是否新病人
    '     bln是否自动识别的身份证-是否自动认别的身份证号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 13:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If Trim(txt预交额.Text) = "" And Trim(txt卡号.Text) = "" Then zlSaveDataBeforCheckIsValid = True: Exit Function
    If objPati Is Nothing Then
        MsgBox "不能确定病人信息，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    If mint操作状态 = 2 Then zlSaveDataBeforCheckIsValid = True: Exit Function '0-增加;1-异常重收;2-异常作废
    If Not mobjDepositItems Is Nothing Then
        If mobjDepositItems.同步状态 >= 4 Then zlSaveDataBeforCheckIsValid = True: Exit Function
    
    End If
    
    If Not mobjCardFeeItems Is Nothing Then
        If mobjCardFeeItems.同步状态 >= 4 Then zlSaveDataBeforCheckIsValid = True: Exit Function
    
    End If
    
    If CheckSendAndBoudCardIsValid(blnNewPati, objPati, bln是否自动识别的身份证) = False Then Exit Function
    If CheckDepositIsValid(objPati, mblnDepositPrint) = False Then Exit Function
    
    
    '预交及发卡的相关检查
    Dim bln相同 As Boolean, objCard As Card, objItems As clsBalanceItems, strErrMsg As String, intSwapStatu As Integer
    bln相同 = CheckDepsoitAndCardFeePayIsSame(mobjDepositItems, mobjCardFeeItems)
    If bln相同 Then
        '一起结算的，需要判断
        Set objCard = GetCardFeePayCard
        If Not (mobjDepositItems(1).objCard.接口序号 = objCard.接口序号 And objCard.消费卡 = mobjDepositItems(1).消费卡) And mobjDepositItems.类型 = gEM_一卡通 Then
            '一卡通结算，需要检查交易
            Set objItems = mobjCardFeeItems.Clone
            objItems.结算金额 = objItems.结算金额 + mobjDepositItems.结算金额
            objItems(1).结算金额 = RoundEx(objItems(1).结算金额 + mobjDepositItems.结算金额, 6)
             Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
           
            If mobjThirdSwap.zlThird_IsSwapIsSucces(objItems, intSwapStatu, strErrMsg, mobjDepositItems(1).预交ID) = False Then
                '交易失败
                'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                If intSwapStatu = 1 Then
                    MsgBox "原" & mobjDepositItems(1).objCard.名称 & " 交易正在进行中，不允许更改支付方式,请检查!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                    Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, False, True)
                    mblnSendCardLocked = True: mblnDepositLocked = True
                    Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '锁定结算方式
                    Exit Function
                End If
            Else
                MsgBox "原" & mobjDepositItems(1).objCard.名称 & " 交易已经成功，不允许更改支付方式,请检查!", vbInformation + vbOKOnly, gstrSysName
                '交易成功
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, False, True)
                mblnSendCardLocked = True: mblnDepositLocked = True
                Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '锁定结算方式
                Exit Function
            End If
        ElseIf mobjDepositItems(1).消费卡 And Not (mobjDepositItems(1).objCard.接口序号 = objCard.接口序号 And objCard.消费卡 = mobjDepositItems(1).消费卡) Then
            '原为消费卡，因余额退款了，所以只能原样退
            MsgBox "原" & mobjDepositItems(1).objCard.名称 & " 已经扣款成功，不允许更改支付方式,请检查!", vbInformation + vbOKOnly, gstrSysName
            mblnSendCardLocked = True: mblnDepositLocked = True
            Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '锁定结算方式
            Exit Function
        End If
        zlSaveDataBeforCheckIsValid = True: Exit Function
    End If
    '不相同的检查
    '预交检查
    If Not mobjDepositItems Is Nothing Then
    
        If mobjDepositItems.Count <> 0 Then
            Set mobjThirdSwap.objPayCards = mobjDepositPayCards
            Set objCard = GetDepositPayCard
            If Not (mobjDepositItems(1).objCard.接口序号 = objCard.接口序号 And objCard.消费卡 = mobjDepositItems(1).消费卡) And mobjDepositItems.类型 = gEM_一卡通 Then
                If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjDepositItems, intSwapStatu, strErrMsg) = False Then
                    '交易失败
                    'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                    If intSwapStatu = 1 Then
                        MsgBox "原" & mobjDepositItems(1).objCard.名称 & " 交易正在进行中，不允许更改支付方式,请检查!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                        Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                        mblnDepositLocked = True: Call SetDepositEditEnabled(1)  '锁定结算方式
                        Exit Function
                    End If
                Else
                    '交易成功
                     MsgBox "原" & mobjDepositItems(1).objCard.名称 & " 交易已经成功，不允许更改支付方式,请检查!", vbInformation + vbOKOnly, gstrSysName
                    Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                    mblnDepositLocked = True: Call SetDepositEditEnabled(1)  '锁定结算方式
                    Exit Function
                End If
            ElseIf mobjDepositItems(1).消费卡 And Not (mobjDepositItems(1).objCard.接口序号 = objCard.接口序号 And objCard.消费卡 = mobjDepositItems(1).消费卡) Then
                '原为消费卡，因余额退款了，所以只能原样退
                MsgBox "原" & mobjDepositItems(1).objCard.名称 & " 已经扣款成功，不允许更改支付方式,请检查!", vbInformation + vbOKOnly, gstrSysName
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                mblnDepositLocked = True: Call SetDepositEditEnabled(1)  '锁定结算方式
                Exit Function
            End If
        End If
    End If
    
    If Not mobjCardFeeItems Is Nothing Then
        If mobjCardFeeItems.Count <> 0 Then
            Set objCard = GetCardFeePayCard
            Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
           
            If Not (mobjCardFeeItems(1).objCard.接口序号 = objCard.接口序号 And objCard.消费卡 = mobjCardFeeItems(1).消费卡) And mobjCardFeeItems.类型 = gEM_一卡通 Then
                If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjDepositItems, intSwapStatu, strErrMsg) = False Then
                    '交易失败
                    'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                    If intSwapStatu = 1 Then
                        MsgBox "原" & mobjCardFeeItems(1).objCard.名称 & " 交易正在进行中，不允许更改支付方式,请检查!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                        Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                        mblnSendCardLocked = True: Call SetCardEditEnabled(1)  '锁定结算方式
                        Exit Function
                    End If
                Else
                    '交易成功
                     MsgBox "原" & mobjCardFeeItems(1).objCard.名称 & " 交易已经成功，不允许更改支付方式,请检查!", vbInformation + vbOKOnly, gstrSysName
                    Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                    mblnSendCardLocked = True: Call SetCardEditEnabled(1)  '锁定结算方式
                    Exit Function
                End If
            ElseIf mobjCardFeeItems(1).消费卡 And Not (mobjCardFeeItems(1).objCard.接口序号 = objCard.接口序号 And objCard.消费卡 = mobjCardFeeItems(1).消费卡) Then
                '原为消费卡，因余额退款了，所以只能原样退
                MsgBox "原" & mobjCardFeeItems(1).objCard.名称 & " 已经扣款成功，不允许更改支付方式,请检查!", vbInformation + vbOKOnly, gstrSysName
                Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                mblnSendCardLocked = True: Call SetCardEditEnabled(1)  '锁定结算方式
                Exit Function
            End If
        End If
    End If
    zlSaveDataBeforCheckIsValid = True
    
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function




Private Function CheckInputItemIsValid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入项的合法性
    '入参
    '返回:输入合法返回true
    '编制:刘兴洪
    '日期:2019-11-26 11:55:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card
    On Error GoTo errHandle
     
    If Not CheckLen(txt缴款单位, 50, "缴款单位") Then Exit Function
    If Not CheckLen(txtPass, 10, "密码") Then Exit Function
    
    If Not CheckLen(txt开户行, 50, "开户行") Then Exit Function
    If Not CheckLen(txt帐号, 50, "帐号") Then Exit Function
    If Not CheckLen(txt结算号码, 30, "结算号码") Then Exit Function
        
        
    Set objSendCard = mCurSendCard.objSendCard
    If Not objSendCard Is Nothing Then
        If Not CheckLen(txt卡号, CInt(objSendCard.卡号长度), "卡号") Then Exit Function
    End If
    CheckInputItemIsValid = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckSendAndBoudCardIsValid(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo, _
    Optional ByVal bln是否自动识别的身份证 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医疗卡发卡及绑定卡数据的合法性
    '     blnNewPati-是否新病人
    '     bln是否自动识别的身份证-是否自动认别的身份证号
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-27 10:21:41
    '问题:25302
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCard As String, strICCard As String
    Dim objSendCard As Card, rs卡费 As ADODB.Recordset
    Dim dtCurrDate As Date, dbl卡费 As Double
    On Error GoTo errHandle
        
    
    strCard = UCase(txt卡号.Text)
    dbl卡费 = Val(txt卡额.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    '-----------------------------------------------------------------------------------------------------------------
    '1.就诊卡的检查
 
    If Not fra磁卡.Visible Then CheckSendAndBoudCardIsValid = True: Exit Function
    If mblnBoundCarded Then CheckSendAndBoudCardIsValid = True: Exit Function '已经绑定卡或发卡的，就不检查，直接退出
    
    If mlngCardTypeID = 0 Then CheckSendAndBoudCardIsValid = True: Exit Function
    
    
    Set rs卡费 = GetCardFee()
    Set objSendCard = mCurSendCard.objSendCard
    
    Select Case tbSendCard.SelectedItem.Key
    Case "CardFee"
        If mobjPati Is Nothing Then Set mobjPati = New clsPatientInfo
        
        If (mobjPati.费别 <> objPati.费别 Or mobjPati.医疗付款方式 <> objPati.医疗付款方式) And fra磁卡.Visible Then
            If tbSendCard.SelectedItem Is Nothing Then Exit Function
            
            If MsgBox("费别及医疗付款方式发生了改变,是否需要重新计算卡费?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
            Call zlRecalcCardFee(objPati)
        End If
        Set mobjPati = objPati
    
        If Trim(txt卡号.Text) <> "" And Not rs卡费 Is Nothing Then
            If dbl卡费 = 0 Then
                MsgBox objSendCard.名称 & "未输入卡费，请检查！", vbExclamation, gstrSysName
                If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus:  Exit Function
            End If
            If rs卡费!是否变价 = 1 Then
                If rs卡费!现价 <> 0 And Abs(CCur(txt卡额.Text)) > Abs(rs卡费!现价) Then
                    MsgBox objSendCard.名称 & "卡金额绝对值不能大于最高限价：" & Format(Abs(rs卡费!现价), "0.00"), vbExclamation, gstrSysName
                    If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus:  Exit Function
                End If
                If rs卡费!原价 <> 0 And Abs(CCur(txt卡额.Text)) < Abs(rs卡费!原价) Then
                    MsgBox objSendCard.名称 & "卡金额绝对值不能小于最低限价：" & Format(Abs(rs卡费!原价), "0.00"), vbExclamation, gstrSysName
                    If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus: Exit Function
                End If
            End If
        End If
        If cbo发卡结算.Visible And txt卡号.Text <> "" And cbo发卡结算.Enabled And cbo发卡结算.ListIndex = -1 Then
            MsgBox "请确定" & objSendCard.名称 & "的缴款结算方式！", vbExclamation, gstrSysName
            If cbo发卡结算.Enabled And cbo发卡结算.Visible Then cbo发卡结算.SetFocus: Exit Function
        End If
        
        '发卡性质的检查
        If Check发卡性质(objPati.病人ID, objSendCard) = False Then Exit Function
        
    Case Else
         Set mobjPati = objPati
    End Select
    
    If bln是否自动识别的身份证 = False And InStr(",二代身份证,身份证,", "," & objSendCard.名称 & ",") > 0 And txt卡号.Text <> "" Then
            
            MsgBox "绑定身份证只能以自动识别的方式进行，不允许手动输入身份证进行绑定!", vbOKOnly + vbInformation, gstrSysName
            txt卡号.Text = "": txtPass.Text = "": txtAudi.Text = ""
            If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
            Exit Function
    End If
    
    
    If txtPass.Text <> txtAudi.Text And txt卡号.Text <> "" Then
        MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus: Exit Function
    End If
    
    If blnNewPati Then  '新病人
        If Trim(txt卡号.Text) = "" And txt卡号.Visible And mblnNewPatiMustSendCard Then
            MsgBox "请刷卡或输入" & objSendCard.名称 & "卡号！", vbExclamation, gstrSysName
            If txt卡号.Enabled And txt卡号.Enabled Then txt卡号.SetFocus
            Exit Function
        End If
    End If
    
    
     
    If txt卡号.Text <> "" Then
        If mobjPubPatient.blnRealName And mobjPati.实名认证 = False And chkEndTime.value = 0 Then
            If MsgBox("未实名认证的病人只能发放临时卡，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                zlControl.ControlSetFocus chkEndTime
                Exit Function
            End If
            chkEndTime.value = 1
        End If
        
        dtCurrDate = zlDatabase.Currentdate
        If Format(CStr(dtpDate.value), "YYYY-MM-DD HH:MM:SS") < dtCurrDate And chkEndTime.value = vbChecked Then
            MsgBox "请选择大于当前时间的终止使用时间！", vbInformation, gstrSysName
            If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
            Exit Function
        End If
        
        
        If objSendCard.是否严格控制 Then
            '保存前检查就诊卡是否有，是否在范围内
           mCurSendCard.lng领用ID = mobjExseSvr.CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), txt卡号.Text, objSendCard.接口序号)

           If mCurSendCard.lng领用ID <= 0 And Not mCurSendCard.blnOneCard Then
               Select Case mCurSendCard.lng领用ID
                   Case 0 '操作失败
                   Case -1
                           If txt卡号.Text <> "" Then MsgBox "你已没有自用及共用的" & objSendCard.名称 & ",不能发放！" & vbCrLf & _
                               "请先在本地设置共用批次或领用一批新卡! ", vbExclamation, gstrSysName
                   Case -2
                           If txt卡号.Text <> "" Then MsgBox "本地共用的" & objSendCard.名称 & "已用完,不能发放！" & vbCrLf & _
                               "请重新设置本地共用卡批次或领用一批新卡！", vbExclamation, gstrSysName
                   Case -3
                       MsgBox "该张卡号不在有效范围内,请检查是否正确刷卡！", vbExclamation, gstrSysName
                       If txt卡号.Enabled And txt卡号.Enabled Then txt卡号.SetFocus
               End Select
               Exit Function
           End If
        End If
        
                
        If objSendCard.卡号长度 <> zlCommFun.ActualLen(Trim(txt卡号)) And Not objSendCard.是否严格控制 Then
            '104238:李南春，2017/2/15，检查卡号是否满足发卡控制限制
            Select Case objSendCard.发卡控制
                Case 0
                    MsgBox "输入的卡号小于" & objSendCard.名称 & "设定的卡号长度，请重新输入！", vbExclamation, gstrSysName
                    If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                    Exit Function
                Case 2
                    If MsgBox("输入的卡号小于" & objSendCard.名称 & "设定的卡号长度，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                        Exit Function
                    End If
            End Select
        End If
    End If
    
    '密码检查
    If txtPass.Visible Then
        Select Case objSendCard.密码长度限制
        Case 0
        Case 1
            If Len(txtPass.Text) <> objSendCard.密码长度 Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & objSendCard.密码长度 & "位", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        Case Else
            If Len(txtPass.Text) < Abs(objSendCard.密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(objSendCard.密码长度限制) & "位以上.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        End Select
    End If
                              
    If Len(Trim(txtPass.Text)) <= 0 And Len(Trim(txt卡号.Text)) > 0 Then '没有输入密码
        If zl_Get设置默认发卡密码 = False Then Exit Function
    End If
    
    Dim cllCons As Collection
    Set cllCons = New Collection
    '                   保存的项目名称包含:操作状态,病人ID,卡类别ID,卡号,新卡号
    '                    操作状态:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    cllCons.Add Array("操作状态", GetCurCard_Statu)
    cllCons.Add Array("病人ID", objPati.病人ID)
    cllCons.Add Array("卡类别ID", objSendCard.接口序号)
    cllCons.Add Array("卡号", txt卡号.Text)
    If mobjService.ZlPatiSvr_ChkCardChangeValid(cllCons) = False Then Exit Function
    
    CheckSendAndBoudCardIsValid = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Check发卡性质(lng病人ID As Long, ByVal objSendCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发卡时检查是否限制病人的发卡张数
    '入参:lng病人ID - 病人ID;lng卡类别ID  - 医疗卡的类别ID
    '编制:王吉
    '问题:57326
    '日期:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, lngPatiID As Long, blnExisted As Boolean
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
    
    If Trim(txt卡号.Text) = "" Or mlngCardTypeID = 0 Then Check发卡性质 = True: Exit Function
    
    If mobjService.ZlPatisvr_CheckCardExist(lng病人ID, objSendCard.接口序号, "", lngPatiID, blnExisted) = False Then Exit Function
    If Not blnExisted Then Check发卡性质 = True: Exit Function
      
    Select Case objSendCard.发卡性质
    Case 0 '不限制
        Check发卡性质 = True
    Case 1 '同一个病人只允许发一张卡
        MsgBox "该病人已经发过" & objSendCard.名称 & ",不能在进行发卡操作!", vbInformation + vbOKOnly
        Check发卡性质 = False
    Case 2 '同一个病人允许发多张卡,但需要提醒
       Check发卡性质 = MsgBox("该病人已经发过" & objSendCard.名称 & ",是否要进行发卡操作?", vbQuestion + vbYesNo) = vbYes
    End Select
    Exit Function
ErrHandl:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDepositIsValid(ByVal objPati As clsPatientInfo, Optional blnPrint_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查缴预交的合法性
    '入参:objPati-病人信息对象
    '
    '出参:blnPrint_Out-是否打印预交收据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 14:12:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    
    If fra预交.Visible = False Then CheckDepositIsValid = True: Exit Function
    
    If RoundEx(StrToNum(txt预交额.Text), 4) = 0 Then CheckDepositIsValid = True: Exit Function
 
    If cbo预交结算.ListIndex = -1 Then
        MsgBox "请确定病人预交款结算方式！", vbInformation, gstrSysName
        If cbo预交结算.Enabled And cbo预交结算.Visible Then cbo预交结算.SetFocus
        Exit Function
    End If
    
    If cbo预交结算.ItemData(cbo预交结算.ListIndex) = 3 Then
        If mintInsure = 0 Then
            MsgBox "当前病人不是医保病人，不允许使用" & cbo预交结算.Text & "进行预交款缴款.", vbInformation
            Exit Function
        End If
        If mstr医保号 = "" Then
            MsgBox "当前病人不能确定医保号，不允许使用" & cbo预交结算.Text & "进行预交款缴款.", vbInformation
            Exit Function
        End If
        
        If CCur(StrToNum(txt预交额.Text)) > mcurYBMoney Then
            MsgBox "医保个人帐户转入金额不能大于余额:" & Format(mcurYBMoney, "0.00"), vbInformation, gstrSysName
            If txt预交额.Enabled And txt预交额.Visible Then txt预交额.SetFocus: Exit Function
        End If
  
    End If
            
    blnPrint_Out = True
    Select Case mobjDepositFact.打印方式
    Case "0" '不打印预交发票
        blnPrint_Out = False
    Case "1" '自动打印
        blnPrint_Out = True
    Case "2" '打印提醒
        blnPrint_Out = MsgBox("是否打印预交款票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End Select
    
    If blnPrint_Out Then
        If mblnDepositStrictly Then '严格控制
            If Trim(txtFact.Text) = "" Then
                MsgBox "必须输入一个有效的预交票据号码！", vbInformation, gstrSysName
                If txtFact.Enabled And txtFact.Visible Then txtFact.SetFocus
                Exit Function
            End If
            
            mobjDepositFact.领用ID = mobjExseSvr.CheckUsedBill(2, IIf(mobjDepositFact.领用ID > 0, mobjDepositFact.领用ID, mobjDepositFact.LastUseID), txtFact.Text, Val(Mid(tbDeposit.SelectedItem.Key, 2)))
            If mobjDepositFact.领用ID <= 0 Then
                Select Case mobjDepositFact.领用ID
                    Case 0 '操作失败
                    Case -1
                        MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Case -2
                        MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                    Case -3
                        MsgBox "票据号码不在当前有效领用范围内,请重新输入！", vbInformation, gstrSysName
                        If txtFact.Enabled And txtFact.Visible Then txtFact.SetFocus
                End Select
                Exit Function
            End If
        Else
            '非严格控制
            If Len(txtFact.Text) <> mbyt预交票据长度 And txtFact.Text <> "" Then
                MsgBox "预交票据号码长度应该为 " & mbyt预交票据长度 & " 位！", vbInformation, gstrSysName
                If txtFact.Enabled And txtFact.Visible Then txtFact.SetFocus
            End If
        End If
    
    End If
     
    CheckDepositIsValid = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    OpenPassKeyboard txtAudi, True
    RaiseEvent ControlGotFocus(txtAudi)
End Sub
Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    Dim objSendCard As Card
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    If KeyAscii <> 13 Then
        If objSendCard.密码规则 = 1 Then
            Call zlControl.TxtCheckKeyPress(txtAudi, KeyAscii, m数字式)
        End If
    End If
    
    If KeyAscii = 13 Then
        If txtPass.Text <> txtAudi.Text Then
            MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
            Call zlControl.TxtSelAll(txtAudi)
            If txtAudi.Enabled And txtAudi.Visible Then txtAudi.SetFocus
        Else
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub
Private Sub txtAudi_LostFocus()
    Call ClosePassKeyboard(txtAudi)
End Sub

Private Sub txtAudi_Validate(Cancel As Boolean)
    Dim objSendCard As Card
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    Select Case objSendCard.密码长度限制
        Case 0
        Case 1
            If Len(txtAudi.Text) <> objSendCard.密码长度 Then
                MsgBox "注意:" & vbCrLf & "确认密码必须输入" & objSendCard.密码长度 & "位", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Cancel = True
                Exit Sub
             End If
        Case Else
            If Len(txtAudi.Text) < Abs(objSendCard.密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "确密码必须输入" & Abs(objSendCard.密码长度限制) & "位以上.", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Cancel = True
                Exit Sub
             End If
        End Select
End Sub





Private Sub txt预交额_GotFocus()
    If IsNumeric(txt预交额.Text) Then
        txt预交额.Text = StrToNum(txt预交额.Text)
    Else
        txt预交额.Text = ""
    End If
    txt预交额.SelStart = 0: txt预交额.SelLength = Len(txt预交额.Text)
    RaiseEvent ControlGotFocus(txt预交额)
End Sub
Private Sub txt预交额_Validate(Cancel As Boolean)
    Call CalcRQCodePayTotal
End Sub
Private Sub txt预交额_LostFocus()
    
    If IsNumeric(txt预交额.Text) Then
        txt预交额.Text = Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;")
    Else
        txt预交额.Text = ""
    End If
    If txt预交额.MaxLength > 12 Then txt预交额.MaxLength = 12
End Sub

Private Sub txt预交额_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii <> 13 Then
        If InStr(txt预交额.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        '65965:刘鹏飞,2013-09-24,处理预交显示千位位格式
        If (txt预交额.Text <> "" And txt预交额.SelLength <> Len(Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;"))) And _
            (Len(Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;")) >= txt预交额.MaxLength) And _
            InStr(Chr(8), Chr(KeyAscii)) = 0 Then
            If txt预交额.SelLength > 0 And txt预交额.SelLength <= txt预交额.MaxLength Then
            Else
                KeyAscii = 0
            End If
        End If
        Exit Sub
    End If
    
    If IsNumeric(txt预交额.Text) Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If

    '不收取预交款,直接跳过
    txt预交额.Text = ""
    If fra磁卡.Visible Then
       If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
       Exit Sub
    End If
    
    RaiseEvent InputOver '输入完成
End Sub



Private Sub txtFact_GotFocus()
    zlControl.TxtSelAll txtFact
    RaiseEvent ControlGotFocus(txtFact)
End Sub

Private Sub txtFact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtFact.Text) = txtFact.MaxLength And KeyAscii <> 8 And txtFact.SelLength <> Len(txtFact) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


Private Sub txt帐号_GotFocus()
    If StrToNum(txt预交额.Text) <> 0 And txt帐号.Text = "" Then txt帐号.Text = mstr单位帐号
    zlControl.TxtSelAll txt帐号
End Sub

Private Sub txt帐号_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt缴款单位, KeyAscii
End Sub

Private Sub txt帐号_LostFocus()
    Call zlCommFun.OpenIme
End Sub


Private Sub txt缴款单位_GotFocus()
    If StrToNum(txt预交额.Text) <> 0 And txt缴款单位.Text = "" Then txt缴款单位.Text = mstr缴款单位
    zlControl.TxtSelAll txt缴款单位
    Call zlCommFun.OpenIme(True)
    
    RaiseEvent ControlGotFocus(txt缴款单位)
End Sub

Private Sub txt缴款单位_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt缴款单位, KeyAscii
End Sub

Private Sub txt缴款单位_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt结算号码_GotFocus()
    zlControl.TxtSelAll txt结算号码
    RaiseEvent ControlGotFocus(txt结算号码)
End Sub

Private Sub txt结算号码_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt结算号码, KeyAscii
End Sub

  
Private Sub txt开户行_GotFocus()
    If IsNumeric(txt预交额.Text) And txt开户行.Text = "" Then
        txt开户行.Text = mstr单位开户行
    End If
    zlControl.TxtSelAll txt开户行
    Call zlCommFun.OpenIme(True)
    RaiseEvent ControlGotFocus(txt开户行)
End Sub

Private Sub txt开户行_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt开户行, KeyAscii
End Sub

Private Sub txt开户行_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Dim objSendCard As Card
    
    
    If KeyAscii <> 13 Then
        Set objSendCard = mCurSendCard.objSendCard
        If Not objSendCard Is Nothing Then
            If objSendCard.密码规则 = 1 Then
                Call zlControl.TxtCheckKeyPress(txtPass, KeyAscii, m数字式)
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            If Not txt卡额.Locked And txt卡额.TabStop And txt卡额.Enabled Then
                    txt卡额.SetFocus
            ElseIf chk记帐.Visible And chk记帐.Enabled Then
                chk记帐.SetFocus
            ElseIf Me.cbo发卡结算.Enabled And cbo发卡结算.Visible Then
                cbo发卡结算.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
           Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    OpenPassKeyboard txtPass, False
    RaiseEvent ControlGotFocus(txtPass)
End Sub


Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub
Private Sub txtPass_Validate(Cancel As Boolean)
    Dim objSendCard As Card
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    Select Case objSendCard.密码长度限制
    Case 0
    Case 1
        If Len(txtPass.Text) <> objSendCard.密码长度 Then
            MsgBox "注意:" & vbCrLf & "密码必须输入" & objSendCard.密码长度 & "位", vbOKOnly + vbInformation
            If txtPass.Enabled Then txtPass.SetFocus
            Exit Sub
         End If
    Case Else
        If Len(txtPass.Text) < Abs(objSendCard.密码长度限制) Then
            MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(objSendCard.密码长度限制) & "位以上.", vbOKOnly + vbInformation
            If txtPass.Enabled Then txtPass.SetFocus
            Exit Sub
         End If
    End Select
End Sub

Private Function zl_Get设置默认发卡密码() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置默认发卡密码
    '返回:是否继续发卡操作
    '编制:王吉
    '日期:2012-07-06 15:53:14
    '问题号:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card, strID As String
    Dim msgResult As VbMsgBoxResult
 
    Set objSendCard = mCurSendCard.objSendCard
    
    If objSendCard Is Nothing Then Exit Function
    
    If objSendCard.密码长度 = 0 Then  '无限制
        '不控制
    ElseIf objSendCard.是否缺省密码 = 1 Then   '缺省身份证后N位
    
        strID = IIf(mobjPati.身份证号 <> "", Trim(mobjPati.身份证号), Trim(mobjPati.联系人身份证号))
        If Len(strID) > 0 Then    '输入了身份证或联系人身份证号
            txtPass.Text = Right(strID, objSendCard.密码长度)
            zl_Get设置默认发卡密码 = True: Exit Function
        End If
    Else
        zl_Get设置默认发卡密码 = True: Exit Function
    End If
    
    Select Case objSendCard.密码输入限制
        Case 0 '无限制
            zl_Get设置默认发卡密码 = True
            Exit Function
        Case 1 '未输入提醒
            msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续！", vbQuestion + vbYesNo, gstrSysName)
            zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
            Exit Function
        Case 2 '为输入禁止
            MsgBox "未输入卡密码,不能进行发卡？", vbExclamation, gstrSysName
            zl_Get设置默认发卡密码 = False
            Exit Function
    End Select
        

End Function


Private Sub txt卡额_GotFocus()
    zlControl.TxtSelAll txt卡额
    RaiseEvent ControlGotFocus(txt卡额)
End Sub

Private Sub txt卡额_KeyPress(KeyAscii As Integer)
    Dim rs卡费 As ADODB.Recordset
    Dim objSendCard As Card
    
    If txt卡额.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Set rs卡费 = GetCardFee
        If Not rs卡费 Is Nothing Then
            Set objSendCard = mCurSendCard.objSendCard
            If rs卡费!是否变价 = 1 Then
                If rs卡费!现价 <> 0 And Abs(CCur(txt卡额.Text)) > Abs(rs卡费!现价) Then
                    MsgBox objSendCard.名称 & "卡金额绝对值不能大于最高限价：" & Format(Abs(rs卡费!现价), "0.00"), vbExclamation, gstrSysName
                    If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus: Call zlControl.TxtSelAll(txt卡额): Exit Sub
                End If
                If rs卡费!原价 <> 0 And Abs(CCur(txt卡额.Text)) < Abs(rs卡费!原价) Then
                    MsgBox objSendCard.名称 & "卡金额绝对值不能小于最低限价：" & Format(Abs(rs卡费!原价), "0.00"), vbExclamation, gstrSysName
                    If txt卡额.Enabled And txt卡额.Visible Then txt卡额.SetFocus: Call zlControl.TxtSelAll(txt卡额): Exit Sub
                End If
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr(txt卡额.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0:  Exit Sub
    End If
End Sub

Private Sub txt卡额_Validate(Cancel As Boolean)
    If chk记帐.value = 0 Then Call CalcRQCodePayTotal
End Sub


Private Sub txt卡号_GotFocus()
    zlControl.TxtSelAll txt卡号
    Call SetBrushCardObject(True)
    RaiseEvent ControlGotFocus(txt卡号)
End Sub

Private Sub txt卡号_Change()
    Call SetCardEditEnabled(IIf(mblnSendCardLocked, 1, 0))
    Call CalcRQCodePayTotal '计算扫码付总额
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    Dim objSendCard As Card
    
    'mbln是否扫描身份证 = False
    
    Set objSendCard = mCurSendCard.objSendCard
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> 13 Then
        '118070:李南春,2018/4/24,设备只带回车的需要将卡号增长一位
        If objSendCard Is Nothing Then Exit Sub
        
        If txt卡号.SelLength = objSendCard.卡号长度 Then txt卡号.Text = ""
        If Len(txt卡号.Text) = objSendCard.卡号长度 - IIf(objSendCard.设备是否启用回车, 0, 1) And KeyAscii <> 8 Then
            txt卡号.Text = txt卡号.Text & IIf(objSendCard.设备是否启用回车, "", Chr(KeyAscii))
            
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
        
    ElseIf txt卡号.Text = "" Then
        KeyAscii = 0: RaiseEvent InputOver
    Else
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
     
End Sub

Private Sub txt卡号_LostFocus()
    Call SetBrushCardObject(False)
End Sub

Private Sub txt卡号_Validate(Cancel As Boolean)
    Dim lngPatientID As Long, int变动类型 As Integer
    Dim blnCardBind As Boolean  '卡是否进行绑定
    Dim objSendCard As Card
    
    Set objSendCard = mCurSendCard.objSendCard
    
    txt卡号.Text = Trim(txt卡号.Text)
    Call ReLoadCardFee
    Call CheckFreeCard(txt卡号.Text)

    If objSendCard.卡号长度 = Len(Trim(txt卡号.Text)) Then
        
        If mobjOneCardComLib.objOneCardObject.zlGetPatiIDFromCardNo(objSendCard.接口序号, Trim(txt卡号.Text), lngPatientID, False, False) = False Then Exit Sub
         
        If objSendCard.自制卡 And objSendCard.卡号重复使用 And lngPatientID > 0 Then
        
           Call mobjService.zlPatiSvr_GetCardLastChange(lngPatientID, objSendCard.接口序号, txt卡号.Text, int变动类型)
            If int变动类型 = 11 Then
                '如果是绑定
                If MsgBox("卡号为【" & txt卡号.Text & "】的{" & objSendCard.名称 & "}的卡已经与病人标识为【" & lngPatientID & "】的进行了绑定！" & vbCrLf & "是否取消该卡的绑定?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                    Cancel = True
                    txt卡号.Text = ""
                    Exit Sub
                End If
                If BlandCancel(objSendCard.接口序号, Trim(txt卡号.Text), lngPatientID) Then Exit Sub
            End If
        End If

        MsgBox "该卡号已经被绑定,不能绑定该卡号.", vbInformation, gstrSysName
        Cancel = True
        txt卡号.Text = ""
        Exit Sub
   End If
    
End Sub

Private Function BlandCancel(ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消绑定卡
    '入参:intType:0-当前卡号;1-当前类别;2-当前病人所有
    '返回:取消成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-29 11:18:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtCurdate As Date
    Dim cllSaveCard As Collection
    On Error GoTo errHandle

    dtCurdate = zlDatabase.Currentdate
    
    ' 入参 :cllCard-节点包含:操作类型,病人ID,卡类别ID,原卡号,医疗卡号,二维码,变动原因,密码,IC卡号,挂失方式,终止使用时间,单据号,卡费,操作时间,操作员姓名,操作员编号
    '               其中的操作类型:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    '               每项格式:array("名称",值 )
    Set cllSaveCard = New Collection
    cllSaveCard.Add Array("操作类型", 14)
    cllSaveCard.Add Array("病人ID", lng病人ID)
    cllSaveCard.Add Array("卡类别ID", lngCardTypeID)
    cllSaveCard.Add Array("医疗卡号", strCardNo)
    cllSaveCard.Add Array("变动原因", "卡重复自动取消原卡绑定信息")
    cllSaveCard.Add Array("操作时间", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS"))
    cllSaveCard.Add Array("操作员编号", UserInfo.编号)
    cllSaveCard.Add Array("操作员姓名", UserInfo.姓名)
    If mobjService.zlPatisvr_SaveMedcCard(cllSaveCard) = False Then Exit Function
    BlandCancel = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub ReLoadCardFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载卡费性
    '编制:刘兴洪
    '日期:2019-11-25 15:52:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, lng收费细目id As Long
    Dim strSql As String, str年龄 As String
    Dim rsTmp As ADODB.Recordset, rs卡费 As ADODB.Recordset
    Dim objSendCard As Card, dblMoney As Double
    Dim objPati As clsPatientInfo
    On Error GoTo errHandle
    Set rs卡费 = GetCardFee
    
    If rs卡费 Is Nothing Or Trim(txt卡号.Text) = "" Then Exit Sub
    If mobjPati Is Nothing Or rs卡费.RecordCount = 0 Then Exit Sub
    
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    If objSendCard.接口序号 = 0 Then Exit Sub
    
    lng病人ID = mobjPati.病人ID
    str年龄 = mobjPati.年龄
     
    rs卡费.MoveFirst
    
    strSql = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as 收费细目ID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "卡费", mlngModule, objSendCard.接口序号, Trim(txt卡号.Text), lng病人ID, _
                mobjPati.姓名, mobjPati.性别, mobjPati.年龄, mobjPati.身份证号, Val(Nvl(rs卡费!收费细目ID)))
    If rsTmp.EOF Then Exit Sub
    
    lng收费细目id = Val(Nvl(rsTmp!收费细目ID))
    Set rsTmp = zlGetSpecialItemFee(objSendCard.特定项目, mstrPriceGrade, lng收费细目id)
    If Not rsTmp Is Nothing Then Set rs卡费 = rsTmp
    
    With rs卡费
        txt卡额.Text = Format(IIf(Val(Nvl(!是否变价)) = 1, Val(Nvl(!缺省价格)), Val(Nvl(!现价))), "0.00")
        txt卡额.Tag = txt卡额.Text  '保持不变
        txt卡额.Locked = Not (Val(Nvl(!是否变价)) = 1)
        txt卡额.TabStop = (Val(Nvl(!是否变价)) = 1)
        If rs卡费!是否变价 = 0 And Val(txt卡额.Text) <> 0 Then
            If mobjExseSvr.zl_ExseSvr_Actualmoney(mobjPati.费别, rs卡费!收费细目ID, rs卡费!收入项目ID, rs卡费!现价, dblMoney) = False Then Exit Sub
            txt卡额.Text = Format(dblMoney, "0.00")
        End If
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub


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
    If gobjComlib.ErrCenter() = 1 Then
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
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function
 
Private Sub CalcRQCodePayTotal(Optional bln异常 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算汇总扫码付金额
    '日期:2019-11-25 15:35:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    mdblRQCodeMoney = 0
    
    If mobjShowTotalMoneyControl Is Nothing Then Exit Sub
    
    If chk记帐.value = 0 And (txt卡额.Visible Or bln异常) And StrToNum(txt卡额.Text) <> 0 And Trim(txt卡号.Text) <> "" Then
        mdblRQCodeMoney = StrToNum(txt预交额.Text) + StrToNum(txt卡额.Text)
    Else
        mdblRQCodeMoney = StrToNum(txt预交额.Text)
    End If
        
    If UCase(TypeName(mobjShowTotalMoneyControl)) = UCase("TextBox") Then
        mobjShowTotalMoneyControl.Text = Format(mdblRQCodeMoney, "0.00")
    Else
        mobjShowTotalMoneyControl.Caption = "扫码合计：" & Format(mdblRQCodeMoney, "0.00")
    End If
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡接口
    '返回: true-成功，false-失败
    '编制:李南春
    '日期:2016/6/20 13:54:56
    '问题:97634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    Dim objSendCard As Card
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    
    If txt卡号.Locked Then Exit Function
    If mobjOneCardComLib Is Nothing Then Exit Function
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Function
    
    
    If objSendCard.接口序号 <= 0 Or Not (objSendCard.是否扫描 Or objSendCard.是否刷卡) Then Exit Function
    
    If mobjOneCardComLib.zlSetBrushCardObject(objSendCard.接口序号, IIf(blnComm, txt卡号, Nothing), strExpend) Then
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New clsCommEvents
        Call mobjOneCardComLib.zlInitEvents(Me.hWnd, mobjCommEvents)
    End If
End Function



Private Function GetSaveSendCardInfotoCollect(ByVal objPati As clsPatientInfo, ByVal dtCurdate As Date, ByRef cllCardInfo_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保存医疗卡数据集
    '入参:
    '出参:cllCardInfo_Out-返回卡数据集,格式:array(名称,值)
    '         |-操作类型,病人ID,卡类别ID,原卡号,医疗卡号,二维码,变动原因,密码,IC卡号,挂失方式,终止使用时间,单据号,卡费,操作时间,操作员姓名,操作员编号,卡号重用,领用id
    '         操作类型:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    '     cllCardFeeInfo_Out-卡费用信息:
    '
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 18:58:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnInRange   As Boolean, strCardNo As String, strICCard As String
    Dim objSendCard As Card, byt变动类型 As Byte, strEndDate As String
    Dim str变动原因 As String
    
    On Error GoTo errHandle
    
    Set objSendCard = mCurSendCard.objSendCard
    If mlngCardTypeID = 0 Then GetSaveSendCardInfotoCollect = True: Exit Function
    If objSendCard Is Nothing Then Exit Function

    Set cllCardInfo_out = New Collection
    byt变动类型 = GetCurCard_Statu
    strEndDate = ""
    If chkEndTime.value = vbChecked Then
        strEndDate = Format(dtpDate.value, "yyyy-mm-dd HH:MM:SS")
    End If
    
    str变动原因 = Decode(mlngModule, 1101, "病人信息登记发卡", "病人入院登记发卡")
    
    strCardNo = UCase(txt卡号.Text): strICCard = IIf(mblnICCard, strCardNo, "")
    If strCardNo = "" Then GetSaveSendCardInfotoCollect = True: Exit Function
    
    '1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    cllCardInfo_out.Add Array("操作类型", byt变动类型), "_操作类型"
    cllCardInfo_out.Add Array("病人ID", objPati.病人ID), "_病人ID"
    cllCardInfo_out.Add Array("卡类别ID", objSendCard.接口序号), "_卡类别ID"
    cllCardInfo_out.Add Array("原卡号", ""), "_原卡号"
    cllCardInfo_out.Add Array("医疗卡号", strCardNo), "_医疗卡号"
    cllCardInfo_out.Add Array("领用ID", mCurSendCard.lng领用ID), "_领用ID"
    
    cllCardInfo_out.Add Array("二维码", ""), "_二维码"
    cllCardInfo_out.Add Array("变动原因", str变动原因), "_变动原因"
    cllCardInfo_out.Add Array("密码", zlCommFun.zlStringEncode(Trim(txtPass.Text))), "_密码"
    cllCardInfo_out.Add Array("IC卡号", strICCard), "_IC卡号"
    cllCardInfo_out.Add Array("挂失方式", ""), "_挂失方式"
    cllCardInfo_out.Add Array("终止使用时间", strEndDate), "_终止使用时间"
    cllCardInfo_out.Add Array("单据号", ""), "_单据号"
    cllCardInfo_out.Add Array("卡费", StrToNum(txt卡额.Text)), "_卡费"
    cllCardInfo_out.Add Array("操作时间", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")), "_操作时间"
    cllCardInfo_out.Add Array("操作员姓名", UserInfo.姓名), "_操作员姓名"
    cllCardInfo_out.Add Array("操作员编号", UserInfo.编号), "_操作员编号"
    cllCardInfo_out.Add Array("卡号重用", IIf(mCurSendCard.objSendCard.卡号重复使用, 1, 0)), "_卡号重用"
    If mCurSendCard.lng领用ID > 0 Then cllCardInfo_out.Add Array("领用ID", mCurSendCard.lng领用ID), "_领用ID"
    GetSaveSendCardInfotoCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetDepositSaveDataToCollect(ByVal objPati As clsPatientInfo, ByVal objDepositItems As clsBalanceItems, _
    ByRef cllDeposit As Collection, Optional ByVal strCardFeeNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保存预交数据集
    '入参:objPati-病人信息对象
    '     strCardFeeNo-卡费单据号：同时缴款时，传入(预交单据号,发票号,预交类别,病人ID,主页id,姓名,性别,年龄,门诊号,住院号,付款方式编号,付款方式名称,缴款科室id,缴款金额,缴款单位,单位开户行,摘要,领用id)
    '     strDepositNo_Out-预交单据号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-10 19:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney  As Double
    Dim int预交类型 As Integer, int主页id As Long

    On Error GoTo errHandle
        
    dblMoney = StrToNum(txt预交额.Text)
    Set cllDeposit = New Collection
    
    If fra预交.Visible = False Or dblMoney = 0 Then GetDepositSaveDataToCollect = True: Exit Function
    
    If objDepositItems Is Nothing Then Exit Function
    
    int预交类型 = Val(Mid(tbDeposit.SelectedItem.Key, 2))
    int主页id = 0
    If int预交类型 = 2 Then int主页id = objPati.主页ID
    
    'depositinfo:(预交单据号,发票号,预交类别,病人ID,主页id,姓名,性别,年龄,门诊号,住院号,付款方式编号,付款方式名称,缴款科室id,缴款金额,缴款单位,单位开户行,摘要,领用id)
    If objDepositItems.单据号 = "" Then Exit Function
    cllDeposit.Add Array("预交ID", objDepositItems(1).预交ID), "_预交ID"
    cllDeposit.Add Array("预交单据号", objDepositItems.单据号), "_预交单据号"
    cllDeposit.Add Array("发票号", IIf(mblnDepositPrint, txtFact.Text, "")), "_发票号"
    cllDeposit.Add Array("预交类别", Val(Mid(tbDeposit.SelectedItem.Key, 2))), "_预交类别"
    cllDeposit.Add Array("病人ID", objPati.病人ID), "_病人ID"
    cllDeposit.Add Array("主页ID", int主页id), "_主页ID"
    cllDeposit.Add Array("姓名", objPati.姓名), "_姓名"
    cllDeposit.Add Array("性别", objPati.性别), "_性别"
    cllDeposit.Add Array("年龄", objPati.年龄), "_年龄"
    cllDeposit.Add Array("门诊号", objPati.门诊号), "_门诊号"
    cllDeposit.Add Array("住院号", objPati.住院号), "_住院号"
    cllDeposit.Add Array("付款方式编号", objPati.医疗付款方式编码), "_付款方式编号"
    cllDeposit.Add Array("付款方式名称", objPati.医疗付款方式), "_付款方式名称"
    cllDeposit.Add Array("缴款科室ID", Val(txt缴款单位.Tag)), "_缴款科室ID"
    cllDeposit.Add Array("缴款金额", dblMoney), "_缴款金额"
    cllDeposit.Add Array("缴款单位", txt缴款单位.Text), "_缴款单位"
    cllDeposit.Add Array("单位开户行", txt开户行.Text), "_单位开户行"
    cllDeposit.Add Array("开户行账号", txt帐号.Text), "_开户行账号"
    cllDeposit.Add Array("摘要", IIf(strCardFeeNo = "", "", "医疗卡:" & strCardFeeNo)), "_摘要"
    cllDeposit.Add Array("领用ID", mobjDepositFact.领用ID), "_领用ID"
    GetDepositSaveDataToCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCardFeeBalanceSaveDataToColl(ByVal objPati As clsPatientInfo, ByVal objCurBalanceItem As clsBalanceItem, ByRef cllBalanceData_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡结算信息
    '入参:objCurBalanceItem-当前结算信息
    '
    '出参:cllBalanceData_Out-保存结算信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-11 09:26:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    ' balanceinfo:(结算方式,结算号码,卡类别id,结算卡序号,支付卡号,交易流水号,交易说明,合作单位,险类,医保号,医保密码,消费卡ID) Key="_balanceinfo"
    Set cllBalanceData_Out = New Collection
    
    cllBalanceData_Out.Add Array("结算方式", objCurBalanceItem.结算方式), "_" & "结算方式"
    cllBalanceData_Out.Add Array("结算号码", objCurBalanceItem.结算号码), "_" & "结算号码"
    cllBalanceData_Out.Add Array("卡类别ID", IIf(Not objCurBalanceItem.消费卡, objCurBalanceItem.卡类别ID, "")), "_" & "卡类别ID"
    cllBalanceData_Out.Add Array("结算卡序号", IIf(objCurBalanceItem.消费卡, objCurBalanceItem.卡类别ID, "")), "_" & "结算卡序号"
    cllBalanceData_Out.Add Array("支付卡号", objCurBalanceItem.卡号), "_" & "支付卡号"
    cllBalanceData_Out.Add Array("交易流水号", objCurBalanceItem.交易流水号), "_" & "交易流水号"
    cllBalanceData_Out.Add Array("交易说明", objCurBalanceItem.交易说明), "_" & "交易说明"
    cllBalanceData_Out.Add Array("合作单位", ""), "_" & "合作单位"
    cllBalanceData_Out.Add Array("消费卡ID", objCurBalanceItem.消费卡ID), "_" & "消费卡ID"
    
    If objCurBalanceItem.结算性质 = 3 Then
        cllBalanceData_Out.Add Array("险类", mintInsure), "_" & "险类"
        cllBalanceData_Out.Add Array("医保号", mstr医保号), "_" & "医保号"
        cllBalanceData_Out.Add Array("密码", mstr密码), "_" & "密码"
    End If
    GetCardFeeBalanceSaveDataToColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
  
Private Function GetAddDepositAndCardFeeDataToCollect(ByVal int类型 As Integer, ByVal objPati As clsPatientInfo, _
    ByVal objCardFeeItems As clsBalanceItems, ByVal objDepositItems As clsBalanceItems, _
     ByVal dtCurdate As Date, ByRef cllData_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否正确
    '入参:objPati-病人信息集
    '     int类型-0-仅卡费;1-仅预交,2-卡费及预交
    '     objCardFeeItems-当前卡费结算信息
    '     objDepositItems-当前预交结算信息
    '出参:
    '     cllData_Out: 卡数据对象
    '          |--billinfo:(结算合计,操作员编号,操作员姓名,登记时间),Key="_billinfo"
    '          |--patinfo:(病人ID,主页ID,病人姓名,性别,年龄,门诊号,住院号,付款方式编号,费别,险类),Key="_patinfo"
    '          |--cardinfo:发卡信息(卡号,卡类别ID,发卡方式(0-发卡,1-补卡,2-换卡),卡号重用,领用id),key="_cardinfo"
    '          |--cardfeelists:key="_cardfeelists"
    '               |---cardfeelist:(卡费单据号,序号,价格父号,从属父号,收费类别,收费细目id,收入项目id,标准单价,收据费目,应收金额,实收金额,病人科室id,开单部门id,病人病区id,
    '                                 执行部门id,加班标志,是否病历费,保险编码,保险项目否,统筹金额,摘要,发卡卡号,发卡卡类别ID,发卡方式(0-发卡,1-补卡,2-换卡)) ,Key="_" & 序号
    
    '          |--balanceinfo:(结算方式,结算号码,卡类别id,结算卡序号,支付卡号,交易流水号,交易说明,合作单位,险类,医保号,医保密码,消费卡ID) Key="_balanceinfo"
    '          |--depositinfo:(预交单据号,发票号,预交类别,主页id,缴款科室id,缴款金额,缴款单位,单位开户行,摘要,领用id),Key="_depositinfo",无预交时，不传入
    '          以上，格式为:,格式：array(名称,值)
    '          int操作状态=2-保存为记帐单;3-保存为划价单 的，则无"balanceinfo"和"depositinfo"节点
    
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-11 10:37:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllDeposit As Collection, cllCardFee As Collection
    Dim cllBalanceInfo As Collection, strCardFeeNo As String
    Dim cllTemp As Collection
    
    On Error GoTo errHandle
    
    Set cllData_out = New Collection
    
     
    '获取卡费用信息集
    Set cllCardFee = New Collection
    If int类型 = 0 Or int类型 = 2 Then  '卡费处理
        If GetCardFeeSaveDataToCollect(objPati, objCardFeeItems, cllCardFee) = False Then Exit Function
        If Not objCardFeeItems Is Nothing Then strCardFeeNo = objCardFeeItems.单据号
    End If
    
    '获取预交信息集
    If int类型 = 1 Or int类型 = 2 Then
        If GetDepositSaveDataToCollect(objPati, objDepositItems, cllDeposit, strCardFeeNo) = False Then Exit Function
        If objDepositItems Is Nothing Then Set objDepositItems = New clsBalanceItems
    Else
        Set cllDeposit = New Collection
        Set objDepositItems = New clsBalanceItems
    End If
    
    Set cllBalanceInfo = New Collection
    If cllCardFee.Count <> 0 And chk记帐.value = 0 Or cllDeposit.Count <> 0 Then
        '产生结算数据
        If objDepositItems.Count <> 0 Then
            If GetCardFeeBalanceSaveDataToColl(objPati, objDepositItems(1), cllBalanceInfo) = False Then Exit Function
        Else
            If GetCardFeeBalanceSaveDataToColl(objPati, objCardFeeItems(1), cllBalanceInfo) = False Then Exit Function
        End If
    End If
   
   
    '1.构建单据信息数据
    Set cllTemp = New Collection
    If int类型 <> 1 Then
        If objDepositItems Is Nothing Then
            cllTemp.Add Array("结算合计", RoundEx(objCardFeeItems.结算金额, 5)), "_" & "结算合计"
        Else
            cllTemp.Add Array("结算合计", RoundEx(objCardFeeItems.结算金额 + objDepositItems.结算金额, 5)), "_" & "结算合计"
        End If
    End If
    cllTemp.Add Array("操作员编号", UserInfo.编号), "_" & "操作员编号"
    cllTemp.Add Array("操作员姓名", UserInfo.姓名), "_" & "操作员姓名"
    If Not mobjCardFeeItems Is Nothing Then
        
        If mobjCardFeeItems.Count <> 0 Then
            cllTemp.Add Array("结帐ID", mobjCardFeeItems(1).结算ID), "_" & "结帐ID"
        End If
    End If
    cllTemp.Add Array("登记时间", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")), "_" & "登记时间"
    cllData_out.Add cllTemp, "_billinfo"
    
    
    '2.构建病人信息
    If int类型 <> 1 Then
        Set cllTemp = New Collection
        cllTemp.Add Array("病人ID", objPati.病人ID), "_" & "病人ID"
        cllTemp.Add Array("主页ID", objPati.主页ID), "_" & "主页ID"
        cllTemp.Add Array("病人姓名", objPati.姓名), "_" & "病人姓名"
        cllTemp.Add Array("性别", objPati.性别), "_" & "性别"
        cllTemp.Add Array("年龄", objPati.年龄), "_" & "年龄"
        cllTemp.Add Array("门诊号", objPati.门诊号), "_" & "门诊号"
        cllTemp.Add Array("住院号", objPati.住院号), "_" & "住院号"
        cllTemp.Add Array("付款方式编号", objPati.医疗付款方式编码), "_" & "付款方式编号"
        cllTemp.Add Array("付款方式名称", objPati.医疗付款方式), "_" & "付款方式名称"
        cllTemp.Add Array("费别", objPati.费别), "_" & "费别"
        cllTemp.Add Array("险类", 0), "_" & "险类"
        cllData_out.Add cllTemp, "_patinfo"
        '3.构建发卡信息
        If cllCardFee.Count <> 0 Then
            '卡号,卡类别ID,发卡方式(0-发卡,1-补卡,2-换卡),卡号重用,领用id
            '2.发卡信息
            Set cllTemp = New Collection
            cllTemp.Add Array("卡号", Trim(txt卡号.Text)), "_" & "卡号"
            cllTemp.Add Array("卡类别ID", mCurSendCard.objSendCard.接口序号), "_" & "卡类别ID"
            cllTemp.Add Array("发卡方式", 0), "_" & "发卡方式"  '0-发卡,1-补卡,2-换卡
            cllTemp.Add Array("卡号重用", IIf(mCurSendCard.objSendCard.卡号重复使用, 1, 0)), "_" & "卡号重用"
            cllTemp.Add Array("领用ID", mCurSendCard.lng领用ID), "_" & "领用ID"
            cllData_out.Add cllTemp, "_cardinfo"
            
            '卡费
            cllData_out.Add cllCardFee, "_cardfeelists"
        End If
    
    End If
    
    '4.结算信息
    If cllBalanceInfo.Count <> 0 Then
        cllData_out.Add cllBalanceInfo, "_balanceinfo"    '结算信息
        If Not cllDeposit Is Nothing Then
            If cllDeposit.Count <> 0 Then
                cllData_out.Add cllDeposit, "_depositinfo"  '结算信息
            End If
        End If
    End If
    GetAddDepositAndCardFeeDataToCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetCurCard_Statu() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发卡状态
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 21:22:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card, blnInRange As Boolean
    If fra磁卡.Visible = False Then Exit Function
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Set objSendCard = New Card
    
    blnInRange = True
    If mCurSendCard.blnOneCard And objSendCard.是否严格控制 Then blnInRange = mCurSendCard.lng领用ID > 0
    If blnInRange And tbSendCard.SelectedItem.Key = "CardFee" Then
       GetCurCard_Statu = 1
    Else
        GetCurCard_Statu = 11
    End If
End Function
Private Function GetCardFeeSaveDataToCollect(ByVal objPati As clsPatientInfo, ByVal objCardFeeItems As clsBalanceItems, ByRef cllCardFee_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保存卡费数据集
    '入参:objCardFeeItems-当前结算信息
    '出参:cllCardFee_Out-当前卡费数据
    '        |-Row:(卡费单据号,序号,价格父号,从属父号,收费类别,收费细目id,收入项目id,标准单价,收据费目,应收金额,实收金额,病人科室id,开单部门id,病人病区id,
    '                                 执行部门id,是否病历费,保险编码,保险项目否,统筹金额,摘要),Key="_" & 序号
    '
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-10 19:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney  As Double, dbl应收 As Double, dbl实收 As Double, lng执行部门ID As Long
    Dim cllRow As Collection, int序号 As Integer
    Dim rs卡费 As ADODB.Recordset, rs病历费 As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set cllCardFee_Out = New Collection
    
    If fra磁卡.Visible = False Or tbSendCard.SelectedItem.Key <> "CardFee" Then GetCardFeeSaveDataToCollect = True: Exit Function
    
    If objCardFeeItems Is Nothing Then Exit Function
    
    Set rs卡费 = GetCardFee
    If rs卡费 Is Nothing Then Exit Function
    If objCardFeeItems.单据号 = "" Then Exit Function
    '          |--cardfeelists:key="_cardfeelists"
    '               |---cardfeelist:(卡费单据号,序号,价格父号,从属父号,收费类别,收费细目id,收入项目id,标准单价,收据费目,应收金额,实收金额,病人科室id,开单部门id,病人病区id,
    '                                 执行部门id,是否病历费,保险编码,保险项目否,统筹金额,摘要) ,Key="_" & 序号

   
    dbl应收 = IIf(mCurSendCard.bln变价 = False, mCurSendCard.dbl应收金额, StrToNum(txt卡额.Text))
    dbl实收 = StrToNum(txt卡额.Text)
     
    int序号 = 1
    
    '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
     lng执行部门ID = zlGetCardFeeExcuteDeptID(Val(Nvl(rs卡费!收费细目ID)), Val(Nvl(rs卡费!科室标志)), UserInfo.部门ID)
 
    Set cllRow = New Collection
    cllRow.Add Array("卡费单据号", objCardFeeItems.单据号), "_" & "卡费单据号"
    cllRow.Add Array("序号", int序号), "_" & "序号"
    cllRow.Add Array("价格父号", 0), "_" & "价格父号"
    cllRow.Add Array("从属父号", 0), "_" & "从属父号"
    cllRow.Add Array("收费类别", Nvl(rs卡费!收费类别)), "_" & "收费类别"
    cllRow.Add Array("收费细目ID", Nvl(rs卡费!收费细目ID)), "_" & "收费细目ID"
    cllRow.Add Array("收入项目ID", Nvl(rs卡费!收入项目ID)), "_" & "收入项目ID"
    cllRow.Add Array("标准单价", dbl应收), "_" & "标准单价"
    cllRow.Add Array("收据费目", Nvl(rs卡费!收据费目)), "_" & "收据费目"
    
    cllRow.Add Array("应收金额", dbl应收), "_" & "应收金额"
    cllRow.Add Array("实收金额", dbl实收), "_" & "实收金额"
    cllRow.Add Array("病人科室id", UserInfo.部门ID), "_" & "病人科室ID"
    cllRow.Add Array("开单部门id", UserInfo.部门ID), "_" & "开单部门ID"
    cllRow.Add Array("病人病区id", UserInfo.部门ID), "_" & "病人病区ID"
    cllRow.Add Array("执行部门id", lng执行部门ID), "_" & "执行部门ID"
    cllRow.Add Array("是否病历费", 0), "_" & "是否病历费"
    cllRow.Add Array("保险编码", ""), "_" & "保险编码"
    cllRow.Add Array("保险项目否", 0), "_保险项目否" & ""
    cllRow.Add Array("统筹金额", 0), "_" & "统筹金额"
    cllRow.Add Array("摘要", ""), "_" & "摘要"
    cllRow.Add Array("加班标志", IIf(OverTime(), 1, 0)), "_" & "加班标志"
     
    cllCardFee_Out.Add cllRow, "_" & int序号
    GetCardFeeSaveDataToCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetClassMoney(ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        '58322
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        .ActiveConnection = Nothing
        If StrToNum(txt预交额.Text) <> 0 Then
            .AddNew
            !收费类别 = "预交"
            !金额 = StrToNum(txt预交额.Text)
            .Update
        End If
        
        If mCurSendCard.objSendCard.接口序号 <> 0 And cbo发卡结算.Enabled And cbo发卡结算.Visible Then
            .AddNew
            If Not mCurSendCard.rs卡费 Is Nothing Then !收费类别 = mCurSendCard.rs卡费!收费类别
            !金额 = StrToNum(txt卡额.Text)
            .Update
        End If
    End With
    GetClassMoney = True
End Function


Private Sub RestorePayStyle()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:恢复到上次选择的支付方式
    '说明:lbl合计.Tag记录的是上次选择的支付方式
    '       cbo预交结算.Tag记录的是预交款的缺省支付方式
    '       cbo结算方式.Tag记录的是卡费的缺省支付方式
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intDeposit As Integer, intCardFee As Integer
    Dim varTemp As Variant
    
    On Error GoTo errHandle

    If mobjShowTotalMoneyControl.Tag = "" Then Exit Sub
    varTemp = Split(mobjShowTotalMoneyControl.Tag & "|", "|")
    intDeposit = varTemp(0): intCardFee = varTemp(1)
    mobjShowTotalMoneyControl.Tag = ""
    
    '恢复预交款结算方式
        
    If cbo预交结算.Visible And cbo预交结算.Enabled Then
        If intDeposit > cbo预交结算.ListCount - 1 Then
            cbo预交结算.ListIndex = Val(cbo预交结算.Tag)
        Else
            cbo预交结算.ListIndex = intDeposit
        End If
    End If
    '恢复卡费结算方式
    If cbo发卡结算.Visible And cbo发卡结算.Enabled And chk记帐.value = 0 Then
        If intCardFee > cbo发卡结算.ListCount - 1 Then
            cbo发卡结算.ListIndex = Val(cbo发卡结算.Tag)
        Else
            cbo发卡结算.ListIndex = intCardFee
        End If
    End If
    
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub mbtQRCodePay_zlErrShow(ByVal strErrMsg As String, ByVal lngErrNum As Long)
    Call RestorePayStyle '恢复上次选择的支付方式
    If strErrMsg = "" Then Exit Sub
    MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
End Sub

Private Sub mbtQRCodePay_zlGetPayMoney(dblMoney As Double, strExpend As String, blnCancel As Boolean)
    Dim dblDeposit As Double, dblCardFee As Double
    
    Err = 0: On Error GoTo errHandle:
    
    mobjShowTotalMoneyControl.Tag = cbo预交结算.ListIndex & "|" & cbo发卡结算.ListIndex  '记录当前支付方式的Index
    '定位到指定卡类别
    If mbtQRCodePay.Tag = "" Then
        MsgBox "未找到有效的扫码付类别,请检查!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        Exit Sub
    End If

    If fra预交.Visible = False And fra磁卡.Visible = False Then
        MsgBox "没有需要扫码付的费用,不需要进行扫码付款!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True: Exit Sub
    End If
    
    If fra预交.Visible And StrToNum(txt预交额.Text) > 0 Then
        dblDeposit = StrToNum(txt预交额.Text)
    End If
    
    If fra磁卡.Visible And chk记帐.value = 0 And Val(txt卡额.Text) > 0 And txt卡额.Enabled Then
        dblCardFee = StrToNum(txt卡额.Text)
    End If
    
    '获取扫码付金额
    dblMoney = dblDeposit + dblCardFee
    
     If dblMoney < 0 Then
        MsgBox "扫码支付金额为负数，请检查!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        Exit Sub
    End If
    
    If dblMoney = 0 Then
        MsgBox "没有需要扫码付的费用,不需要进行扫码付款!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        zlControl.ControlSetFocus txt预交额
        Exit Sub
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    blnCancel = True
End Sub

Private Sub mbtQRCodePay_zlQRCodePayment(ByVal lngCardTypeID As Long, ByVal strPayMentQRCode As String, ByVal strExpendXML As String, blnCancel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行扫码付款
    '入参:lngCardTypeID-卡类别ID
    '       strPayMentQRCode-二维码付款内码
    '       strExpendXML-暂无
    '出参:strExpendXML-暂无
    '        blnCancel-true表示取消本次扫码付,False-表示本次扫码付成功
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle

    If lngCardTypeID = 0 Or blnCancel Then
        blnCancel = True
        Call RestorePayStyle '恢复上次选择的支付方式
        Exit Sub
    End If

    blnCancel = False
    If LocatePayStyle(lngCardTypeID) = False Then   '定位到扫码付的指定类别上
        blnCancel = True
        MsgBox "不能有效识别当前扫码付的类别，可能本机不支持该类别的扫码付，请与管理员联系！", vbInformation + vbOKOnly, gstrSysName
        Call RestorePayStyle '恢复上次选择的支付方式
        Exit Sub
    End If
    mstrQRcode = strPayMentQRCode
    RaiseEvent ExcuteQRCodePayment
    mstrQRcode = ""
    Call RestorePayStyle  '恢复上次选择的支付方式
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    blnCancel = True
    Call RestorePayStyle '恢复上次选择的支付方式
End Sub

Private Function LocatePayStyle(ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:扫码付时,根据卡类别ID,定位到指定的支付类别上
    '入参:lngCardTypeID-扫码的卡类别ID
    '返回:True-定位到指定的支付类别成功；False-定位到指定的支付类别失败
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFindDeposit As Boolean, blnFindCardFee As Boolean, i As Integer
    Dim objCard As Card
    If lngCardTypeID = 0 Then Exit Function
    
    With cbo预交结算
        If .Visible And .Enabled Then
            For i = 1 To mobjDepositPayCards.Count
                Set objCard = mobjDepositPayCards(i)
                If objCard.接口序号 = lngCardTypeID Then
                    If .ListCount >= i Then .ListIndex = i - 1: blnFindDeposit = True: Exit For
                End If
            Next
        Else
            blnFindDeposit = True
        End If
    End With
    
    With cbo发卡结算
        If .Visible And .Enabled And chk记帐.value = 0 Then
            For i = 1 To mobjCardFeePayCards.Count
                Set objCard = mobjCardFeePayCards(i)
                If objCard.接口序号 = lngCardTypeID Then
                
                    If .ListCount >= i Then .ListIndex = i - 1: blnFindCardFee = True: Exit For
                End If
            Next
        Else
            blnFindCardFee = True
        End If
    End With
    LocatePayStyle = blnFindDeposit And blnFindCardFee
End Function


Private Sub Local结算方式(ByVal lng卡类别ID As Long, Optional bln预交 As Boolean = True, Optional ByVal str结算方式 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定位结算方式
    '编制:刘兴洪
    '日期:2011-07-26 15:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As Cards, cboPay As ComboBox
    Dim i As Long, objCard As Card
    If mblnNotClick Then Exit Sub
    
    If bln预交 Then
       Set objCards = mobjDepositPayCards
        Set cboPay = cbo预交结算
    Else
       Set objCards = mobjCardFeePayCards
        Set cboPay = cbo发卡结算
    End If
    
    If objCards Is Nothing Then Exit Sub
    
    With cboPay
        mblnNotClick = True
        For i = 0 To .ListCount - 1
            Set objCard = objCards(i + 1)
            
            ''短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
            If lng卡类别ID > 0 Then
                If objCard.接口序号 = lng卡类别ID Then
                    .ListIndex = i: Exit For
                End If
            Else
                If objCard.结算方式 = str结算方式 Then
                    .ListIndex = i: Exit For
                End If
            End If
        Next
        mblnNotClick = False
    End With
End Sub
Public Property Get RealName() As Boolean
       RealName = mblnRealName
End Property

Public Property Let RealName(ByVal vNewValue As Boolean)
    mblnRealName = vNewValue
    
    If Not mobjPubPatient.blnRealName Then
        Exit Property
    End If
    If Not mobjPati Is Nothing Then mobjPati.实名认证 = mblnRealName
    '未进行实名证的,只能是零时卡
    chkEndTime.value = IIf(mblnRealName, 0, 1)
    chkEndTime.Enabled = Trim(txt卡号.Text) <> "" And mblnRealName
End Property

Public Property Get GetWidth() As Long
       GetWidth = Me.Width
End Property

Public Property Get GetHeight() As Long
       GetHeight = Me.Height
End Property

Public Sub Load预交结算方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载预交款支付方式
    '编制:刘兴洪
    '日期:2019-11-26 14:27:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, str性质 As String
    Dim i As Long
    
    str性质 = "1,2,8" & IIf(mblnAllowInsureAccDeposit, ",3", "")
    If mobjThirdSwap.zlGetBalanceModeCards(mobjDepositPayCards, , , , mstrQRCodeTypeIds_Deposit, "预交款", str性质) = False Then Set mobjDepositPayCards = New Cards
    With cbo预交结算
        .Clear
        mblnNotClick = True
        For i = 1 To mobjDepositPayCards.Count
            Set objCard = mobjDepositPayCards(i)
            .AddItem objCard.名称
            .ItemData(.NewIndex) = objCard.结算性质
            If objCard.缺省标志 = 1 Then .ListIndex = i
        Next
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        .Enabled = .ListCount > 0
        mblnNotClick = False
    End With
    lblStyle.Tag = mstrQRCodeTypeIds_Deposit
End Sub
Public Sub Load卡费结算方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载卡费支付方式
    '编制:刘兴洪
    '日期:2019-11-26 14:27:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim i As Long, str性质 As String
    
    str性质 = "1,2,8"
    If mobjThirdSwap.zlGetBalanceModeCards(mobjCardFeePayCards, , , , mstrQRCodeTypeIds_CardFee, "就诊卡", str性质) = False Then Set mobjDepositPayCards = New Cards
    With cbo发卡结算
        .Clear
        mblnNotClick = True
        For i = 1 To mobjCardFeePayCards.Count
            Set objCard = mobjCardFeePayCards(i)
            .AddItem objCard.名称
            .ItemData(.NewIndex) = objCard.结算性质
            If objCard.缺省标志 = 1 Then .ListIndex = i
        Next
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        .Enabled = .ListCount > 0
        mblnNotClick = False
    End With
    lbl结算方式.Tag = mstrQRCodeTypeIds_CardFee
End Sub

Private Function Load支付方式(Optional ByVal bln显示预交 As Boolean = True, Optional ByVal bln显示卡费 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载支付方式
    '入参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-26 15:04:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strQRCodeTypeIDs As String, varDeposit As Variant, varCardFee As Variant, varTemp As Variant
    Dim strQRCardTypeIDs As String, strErrMsg As String
    Dim i As Long, j As Long
    Call Load预交结算方式
    Call Load卡费结算方式
    
    If cbo预交结算.ListCount = 0 Then
        MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not mbtQRCodePay Is Nothing Then
        
        If mstrQRCodeTypeIds_Deposit <> mstrQRCodeTypeIds_CardFee Then
            varDeposit = Split(mstrQRCodeTypeIds_Deposit & ",", ",")
            varCardFee = Split(mstrQRCodeTypeIds_CardFee & ",", ",")
            For i = 0 To UBound(varDeposit)
                For j = 0 To UBound(varCardFee)
                    If varCardFee(j) = varDeposit(i) Then
                        strQRCodeTypeIDs = strQRCodeTypeIDs & "," & varDeposit(i)
                        Exit For
                    End If
                Next
            Next
            If strQRCodeTypeIDs <> "" Then strQRCodeTypeIDs = Mid(strQRCodeTypeIDs, 2)
        Else
            strQRCodeTypeIDs = mstrQRCodeTypeIds_Deposit
        End If
        If strQRCodeTypeIDs <> "" Then mbtQRCodePay.Tag = strQRCodeTypeIDs
            
        '初始化扫码控件
        If bln显示预交 And bln显示卡费 Then
            strQRCardTypeIDs = mbtQRCodePay.Tag
        ElseIf bln显示预交 And Not bln显示卡费 Then
            strQRCardTypeIDs = lblStyle.Tag
        ElseIf Not bln显示预交 And bln显示卡费 Then
            strQRCardTypeIDs = lbl结算方式.Tag
        End If
        
        If mbtQRCodePay.zlInit(Me, strQRCardTypeIDs, glngSys, mlngModule, gcnOracle, gstrDBUser, strErrMsg) = False Then strQRCardTypeIDs = ""
        mbtQRCodePay.Tag = strQRCardTypeIDs
        mbtQRCodePay.Visible = strQRCardTypeIDs <> "" Or mblnShowDepositAndSendCard
        mbtQRCodePay.Enabled = strQRCardTypeIDs <> ""
        mobjShowTotalMoneyControl.Visible = strQRCardTypeIDs <> "" Or mblnShowDepositAndSendCard
        
    End If
    Load支付方式 = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function UpdateDepositBlncInfo(ByVal int操作状态 As Integer, ByVal objPati As clsPatientInfo, _
    ByVal objDepositItems As clsBalanceItems, ByVal cllExpendInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新卡费相关结算信息
    '入参:int操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
    '     objDepositItems-当前预交支付方式
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-14 11:49:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllUpdateFeeData As Collection, cllTemp As Collection
    Dim objCurItem As clsBalanceItem, blnTrans As Boolean
    Dim strDepositNo As String, strCardFeeNo As String, strErrMsg As String
    Dim varTemp As Variant, lng变动id As Long, lng预交ID As Long
    Dim cllPro As Collection, strSql As String, int异常状态 As Integer
    Dim cllErrData As Collection
    
    On Error GoTo errHandle
    
    Set cllUpdateFeeData = New Collection
    Set cllTemp = New Collection
     
    If objDepositItems Is Nothing Then Exit Function
    If objDepositItems.Count = 0 Then Exit Function
    
    
    Set objCurItem = objDepositItems(1)
    strDepositNo = objDepositItems.单据号
    lng预交ID = objCurItem.预交ID
    
    
    cllTemp.Add Array("预交单号", strDepositNo), "_" & "预交单号"
    cllTemp.Add Array("预交ID", lng预交ID), "_" & "预交ID"
    cllTemp.Add Array("发票号", txtFact.Text), "_" & "发票号"
    cllTemp.Add Array("领用ID", mobjDepositFact.领用ID), "_" & "领用ID"
    cllTemp.Add Array("病人ID", objPati.病人ID), "_" & "病人ID"
    cllTemp.Add Array("操作员编号", UserInfo.编号), "_" & "操作员编号"
    cllTemp.Add Array("操作员姓名", UserInfo.姓名), "_" & "操作员姓名"
    cllTemp.Add Array("收款时间", Format(objCurItem.结算时间, "yyyy-mm-dd HH:MM:SS")), "_" & "收款时间"
    cllUpdateFeeData.Add cllTemp, "_billinfo"
    
     '结算信息
    Set cllTemp = New Collection
    cllTemp.Add Array("结算方式", objCurItem.结算方式), "_" & "结算方式"
    cllTemp.Add Array("结算号码", objCurItem.结算号码), "_" & "结算号码"
    cllTemp.Add Array("卡类别ID", IIf(objCurItem.消费卡, 0, objCurItem.卡类别ID)), "_" & "卡类别ID"
    cllTemp.Add Array("结算卡序号", IIf(objCurItem.消费卡, objCurItem.卡类别ID, 0)), "_" & "结算卡序号"
    cllTemp.Add Array("卡号", objCurItem.卡号), "_" & "卡号"
    cllTemp.Add Array("交易流水号", objCurItem.交易流水号), "_" & "交易流水号"
    cllTemp.Add Array("交易说明", objCurItem.交易说明), "_" & "交易说明"
    cllTemp.Add Array("摘要", objCurItem.结算摘要), "_" & "摘要"
    cllTemp.Add Array("合作单位", ""), "_" & "合作单位"
    
    If Not cllExpendInfo Is Nothing Then
        cllTemp.Add Array("其他信息集", cllExpendInfo), "_" & "其他信息集"
    End If
    cllUpdateFeeData.Add cllTemp, "_balanceinfo"

    '   cllUpdateDate-修改的结算数据
    '         |--billinfo-单据信息,"_billinfo"
    '              |-预交单号,预交ID,操作员编号,操作员姓名,收款时间,病人ID,发票号，领用ID)
    '         |--balanceinfo-结算信息,"_balanceinfo"
    '                |--(结算方式,结算号码,卡类别id,结算卡序号,卡号,交易流水号,交易说明,摘要,合作单位)
    '                |--其他信息集,
    '                |-----其他信息:交易名称,交易内容
    '     blnShowErrMsg-是否显示错误信息
    
    '同步状态：操作场景=2,3时：0或NULL正常记录;-1-未产生费用;1-未调用接口;2-接口调用成功,3-费用结算修正成功;4-医疗卡信息发卡成功"
    If int操作状态 = 0 Then
         int异常状态 = 2
        If Not GetDelErrDataToColl(objDepositItems.业务ID, objDepositItems.异常ID, cllErrData) Then Exit Function
    ElseIf int操作状态 = 1 Then
        If Not GetUpdateErrDataSyncTagToColl(objDepositItems.异常ID, 1, cllErrData) Then Exit Function
        int异常状态 = 1
    Else
        If Not GetUpdateErrDataSyncTagToColl(objDepositItems.异常ID, 3, cllErrData) Then Exit Function
        int异常状态 = 1
        '0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    If Zl_病人结算异常记录_Modify(int异常状态, cllErrData) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
    End If
    
    If mobjExseSvr.Zl_Exsesvr_Upddepositblncinfo(int操作状态, cllUpdateFeeData, False, strErrMsg) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    UpdateDepositBlncInfo = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SetLoaclePayModefromCard(ByVal objCard As Card, ByVal bln预交 As Boolean, Optional blnAppend As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡对象，设置缺省的支付方式
    '入参:objCard-当前卡对象
    '     blnAppend-未找到，自动增加
    '返回 :定位成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-13 10:10:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objTemp As Card, blnFind As Boolean
    Dim objCombox As ComboBox, objPayCards As Cards
    If objCard Is Nothing Then Exit Function
    
    mblnNotClick = True
    If bln预交 Then
        Set objCombox = cbo预交结算: Set objPayCards = mobjDepositPayCards
    Else
        Set objCombox = cbo发卡结算: Set objPayCards = mobjCardFeePayCards
    End If
    
    blnFind = False
    For i = 0 To objCombox.ListCount - 1
        If bln预交 Then
            Set objTemp = GetDepositPayCard(i)
        Else
            Set objTemp = GetCardFeePayCard(i)
        End If
        
        If objTemp Is Nothing Then Exit Function
        If objTemp.接口序号 = objCard.接口序号 And objTemp.结算方式 = objCard.结算方式 Then
            blnFind = True: objCombox.ListIndex = i: Exit For
        End If
    Next
    If Not blnFind And blnAppend Then
        '未找到
        objCombox.AddItem objCard.结算方式
        objCombox.ItemData(objCombox.NewIndex) = objCombox.ListCount + 1
        objCombox.ListIndex = objCombox.NewIndex
        objPayCards.Add objCard, "K" & objCombox.ListCount + 1
        blnFind = True
    End If
    mblnNotClick = False
    SetLoaclePayModefromCard = blnFind
End Function

Private Function ReadDepositBalanceDataFromDepositNo(ByVal strNO As String, lng异常ID As Long, ByVal lng业务ID As Long, int同步状态 As Integer, _
    ByRef objBalanceItems_Out As clsBalanceItems, Optional bln作废 As Boolean, Optional ByVal str异常交易信息 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取预交结算数据
    '入参:str异常交易信息-当前异常交易信息
    '出参:objBalanceItems_out-预交结算数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-28 10:49:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurItem As clsBalanceItem
    Dim objPayCard As Card, cllDeposit As Collection, int类型 As Integer
    Dim cllSwapinfo As Collection, cllExpends As Collection
    Dim i As Long
    
    On Error GoTo errHandle
    
    Set objBalanceItems_Out = New clsBalanceItems
    objBalanceItems_Out.单据号 = strNO
    
    '加载预交结算方式
    If mobjExseSvr.zl_ExseSvr_GetDepositInfo(strNO, IIf(bln作废, 2, 3), cllDeposit, True, "") = False Then Exit Function
    
    Set objCurItem = New clsBalanceItem
   '出参:cll_Deposit_Out-返回预交票据数据集,key="_"+名称
    '       |-病人ID,主页id,预交ID,预交单据号,发票号,预交类别 ,缴款科室id,缴款金额,缴款单位 ,单位开户行,开户行账号,摘要,操作员姓名,操作员编号,收款时间 ,结算方式,结算号码,
    '       | 卡类别id,结算卡序号 ,消费卡ID,支付卡号,交易流水号,交易说明,合作单位,结算状态 ,关联交易ID,险类,医保号,医保密码
    If Val(cllDeposit("_卡类别ID")) <> 0 Then
        If mobjOneCardComLib.zlGetCard(cllDeposit("_卡类别ID"), False, objPayCard) = False Then Exit Function
        int类型 = 3
    ElseIf Val(cllDeposit("_结算卡序号")) <> 0 Then
         If mobjOneCardComLib.zlGetCard(cllDeposit("_结算卡序号"), True, objPayCard) = False Then Exit Function
         int类型 = 5
    Else
        Set objPayCard = zlGetCardFromBalanceName(cllDeposit("_结算方式"))   '普通的结算方式
        int类型 = 0
        If objPayCard.结算性质 = 3 Then int类型 = 2
    End If
    objBalanceItems_Out.类型 = int类型
    With objCurItem
        Set .objCard = objPayCard
        .结算方式 = cllDeposit("_结算方式")
        .结算号码 = cllDeposit("_结算号码")
        .单据性质 = 1 '' 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
        .关联交易ID = Val(cllDeposit("_关联交易ID"))
        .单据号 = Trim(cllDeposit("_预交单据号"))
        .预交ID = Val(cllDeposit("_预交ID"))
        .异常ID = lng异常ID
        .结算金额 = Val(cllDeposit("_缴款金额"))
        .结算类型 = int类型 ''0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        .结算时间 = CDate(cllDeposit("_收款时间"))
        .结算性质 = objPayCard.结算性质
        .结算摘要 = cllDeposit("_摘要")
        .卡类别ID = Val(cllDeposit("_卡类别ID"))
        If objPayCard.消费卡 Then
            .卡类别ID = Val(cllDeposit("_结算卡序号"))
            .消费卡 = True
            .消费卡ID = Val(cllDeposit("_消费卡ID"))
        End If
        .卡号 = cllDeposit("_支付卡号")
        .交易流水号 = Trim(cllDeposit("_交易流水号"))
        .交易说明 = Trim(cllDeposit("_交易说明"))
        .校对标志 = Trim(cllDeposit("_结算状态"))
        .密码 = cllDeposit("_医保密码")
        .是否预交 = True
        .是否保存 = True
        .是否结算 = (.校对标志 = 2 Or .校对标志 = 0)
        .是否允许编辑 = Not .是否结算
        .是否允许删除 = .是否允许编辑
        .是否允许退现 = .是否允许编辑
    End With
  
    tbDeposit.Visible = False
    If Val(cllDeposit("_预交类别")) <= 1 Then
        mbln门诊预交 = True: mbln住院预交 = False
        fra预交.Caption = "【门诊预交信息】"
    Else
        mbln住院预交 = True: mbln门诊预交 = False
        fra预交.Caption = "【住院预交信息】"
    End If
    
    mintInsure = Val(cllDeposit("_险类"))
    mstr医保号 = Trim(cllDeposit("_医保号"))
    mstr密码 = cllDeposit("_医保密码")
    
    objBalanceItems_Out.AddItem objCurItem
    objBalanceItems_Out.结算金额 = objCurItem.结算金额
    objBalanceItems_Out.是否保存 = True
    objBalanceItems_Out.结算完成 = IIf(objCurItem.校对标志 = 0, True, False)
    objBalanceItems_Out.同步状态 = int同步状态
    objBalanceItems_Out.业务ID = lng业务ID
    objBalanceItems_Out.异常ID = lng异常ID
    objBalanceItems_Out.是否保存 = True
    If str异常交易信息 <> "" And int同步状态 <= 2 Then
         If GetErrSwapInfoByJsonString(str异常交易信息, cllSwapinfo, cllExpends) Then
            '异常数据
             'cllSwapinfo(卡号,卡类别ID,交易流水号,交易说明,交易金额,二维码,支付方式,结算摘要)
            For i = 1 To objBalanceItems_Out.Count
               If cllSwapinfo("_支付方式")(1) <> "" Then objBalanceItems_Out(i).结算方式 = cllSwapinfo("_支付方式")(1)
               objBalanceItems_Out(i).卡类别ID = cllSwapinfo("_卡类别ID")(1)
               objBalanceItems_Out(i).交易流水号 = cllSwapinfo("_交易流水号")(1)
               objBalanceItems_Out(i).交易说明 = cllSwapinfo("_交易说明")(1)
               objBalanceItems_Out(i).卡号 = cllSwapinfo("_卡号")(1)
               objBalanceItems_Out(i).结算摘要 = cllSwapinfo("_结算摘要")(1)
               objBalanceItems_Out(i).QRCode = cllSwapinfo("_二维码")(1)
            Next
         End If
    End If
    
    '初始化数据
    txt预交额.Text = Format(objBalanceItems_Out.结算金额, "0.00")
    
    mblnNotClick = True
    Call SetLoaclePayModefromCard(objCurItem.objCard, True, True): mblnNotClick = False
    
    txt结算号码.Text = objCurItem.结算号码
    If mintInsure = 0 Then
        mblnNotClick = True
        txt缴款单位.Text = Trim(cllDeposit("_缴款单位"))
        txt开户行.Text = Trim(cllDeposit("_单位开户行"))
        txt帐号.Text = Trim(cllDeposit("_开户行账号"))
        chk单位缴款.value = IIf(txt缴款单位.Text <> "", 1, 0)
        mblnNotClick = False
    Else
        chk单位缴款.value = 0
    End If
    Call RefreshFactNo      '刷新发票号
    ReadDepositBalanceDataFromDepositNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFeeDataFromColl(ByVal strNO As String, ByVal cllCardFee As Collection, _
    ByRef rsCardFee_Out As Recordset, Optional ByRef objBalanceItems_Out As clsBalanceItems, Optional ByRef dblMoney_Out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据服务返回的集合,按记录集方式返回信息
    '入参:cllCardFee-当前集合
    '
    '出参:rsCardFee_Out-返回的卡费用集合
    '     objBalanceItems_out-结算信息列表，主要是可能存在记帐，需要给objBalanceItems_out
    '     dblMoney_Out:实收金额
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection
    Dim i As Long, bln记帐 As Boolean
    
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    dblMoney_Out = 0
    Set rsCardFee_Out = New ADODB.Recordset
    With rsCardFee_Out
        If .State = adStateOpen Then .Close
        
        .Fields.Append "单据号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "费用id", adBigInt, , adFldIsNullable
        .Fields.Append "序号", adBigInt, , adFldIsNullable
        .Fields.Append "病人id", adBigInt, , adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "费别", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "收费项目id", adBigInt, , adFldIsNullable
        .Fields.Append "收入项目id", adBigInt, , adFldIsNullable
        .Fields.Append "数次", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "应收金额", adDouble, , adFldIsNullable
        .Fields.Append "实收金额", adDouble, , adFldIsNullable
        
        .Fields.Append "开单人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "操作员编号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "操作员姓名", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "登记时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "发生时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "记录状态", adBigInt, , adFldIsNullable
        
        .Fields.Append "是否病历费", adBigInt, , adFldIsNullable
        .Fields.Append "发票号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "是否记帐", adBigInt, , adFldIsNullable
        .Fields.Append "费用状态", adBigInt, , adFldIsNullable
        .Fields.Append "卡类别ID", adBigInt, , adFldIsNullable
        .Fields.Append "卡号", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "是否挂号发卡", adBigInt, , adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    If cllCardFee Is Nothing Then Exit Function
    '    fee_id  N   1   费用id
    '    fee_num N   1   序号
    '    pati_id N   1   病人id
    '    pati_name   C   1   姓名
    '    pati_sex    C   1   性别
    '    pati_age    C   1   年龄
    '    fee_category    C   1   费别
    '    item_id N   1   收费项目id
    '    income_item_id  N   1   收入项目id
    '    quantity    N   1   数次
    '    fee_amrcvb  N   1   应收金额
    '    fee_ampaid  N   1   实收金额
    '    placer  C   1   开单人
    '    operator_code   C   1   操作员编号
    '    operator_name   C   1   操作员姓名
    '    create_time D   1   登记时间
    '    happen_time D   1   发生时间
    '    rec_status  N   1   记录状态
    '    mrbkfee_sign N   1   是否病历费:1-是病历费;0-不是病历费
    '    invoice_no  N   1   发票号
    '    kpbooks_sign N   1   记帐标志:1-是记帐;0-现收
    '    fee_status   N   1   费用状态:1-异常状态;0-正常费用
    '    cardtype_id N   1   卡类别ID
    '    card_no C   1   卡号
    '    sendcard_reg    N   1   是否挂挂号同步发卡:1-是挂号同时发卡;0-非挂号同时发卡

    For i = 1 To cllCardFee.Count
        Set cllTemp = cllCardFee(i)
        
        If Not bln记帐 Then bln记帐 = Val(Nvl(cllTemp("_kpbooks_sign"))) = 1
        With rsCardFee_Out
            .AddNew
            !单据号 = strNO
            !费用id = Val(Nvl(cllTemp("_fee_id")))
            !序号 = Val(Nvl(cllTemp("_fee_num")))
            !病人ID = Val(Nvl(cllTemp("_pati_id")))
            !姓名 = Nvl(cllTemp("_pati_name"))
            !性别 = Nvl(cllTemp("_pati_sex"))
            !年龄 = Nvl(cllTemp("_pati_age"))
            !费别 = Nvl(cllTemp("_fee_category"))
            !收费项目ID = Val(Nvl(cllTemp("_item_id")))
            !收入项目ID = Val(Nvl(cllTemp("_income_item_id")))
            !数次 = Val(Nvl(cllTemp("_quantity")))
            !应收金额 = Val(Nvl(cllTemp("_fee_amrcvb")))
            !实收金额 = Val(Nvl(cllTemp("_fee_ampaid")))
            !开单人 = Nvl(cllTemp("_placer"))
            !操作员编号 = Nvl(cllTemp("_operator_code"))
            !操作员姓名 = Nvl(cllTemp("_operator_name"))
            !登记时间 = Nvl(cllTemp("_create_time"))
            !发生时间 = Nvl(cllTemp("_happen_time"))
            !记录状态 = Val(Nvl(cllTemp("_rec_status")))
            
            !是否病历费 = Val(Nvl(cllTemp("_mrbkfee_sign")))
            !发票号 = Nvl(cllTemp("_invoice_no"))
            !是否记帐 = Val(Nvl(cllTemp("_kpbooks_sign")))
            !费用状态 = Val(Nvl(cllTemp("_fee_status")))
            !卡类别ID = Val(Nvl(cllTemp("_cardtype_id")))
            !卡号 = Nvl(cllTemp("_card_no"))
            !是否挂号发卡 = Val(Nvl(cllTemp("_sendcard_reg")))
            .Update
            dblMoney_Out = RoundEx(dblMoney_Out + Val(Nvl(rsCardFee_Out!实收金额)), 5)
        End With
    Next
    If bln记帐 Then
        objBalanceItems_Out.类型 = gEM_记帐单
    End If
    objBalanceItems_Out.结算金额 = dblMoney_Out
    objBalanceItems_Out.单据号 = strNO
    objBalanceItems_Out.是否保存 = True
    Set cllTemp = Nothing
    zlGetCardFeeDataFromColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReadCardFeeBalanceDataFromNo(ByVal strNO As String, lng异常ID As Long, ByVal lng业务ID As Long, int同步状态 As Integer, _
    ByRef objBalanceItems_Out As clsBalanceItems, Optional bln作废 As Boolean, Optional ByVal str异常交易信息 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡费结算数据
    '入参:
    '出参:objBalanceItems_out-卡费结算数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-28 10:49:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurItem As clsBalanceItem, cllCardFee As Collection, cllPriceBill As Collection, cllBalance As Collection
    Dim objPayCard As Card, int类型 As Integer, dblMoney As Double
    Dim cllSwapinfo As Collection, cllExpends As Collection
    Dim i As Long
    On Error GoTo errHandle
    
    Set objBalanceItems_Out = New clsBalanceItems
    objBalanceItems_Out.单据号 = strNO
    
    '加载预交结算方式
    '查询类型：0-读取正常单据:1-读取作废单据;2-剩余费用单据
    If mobjExseSvr.zl_ExseSvr_GetCardFeeInfoByNo(strNO, IIf(bln作废, 1, 0), cllCardFee, cllPriceBill, cllBalance, Nothing, , , False) = False Then Exit Function
    If zlGetCardFeeDataFromColl(strNO, cllCardFee, mrsCardFee, objBalanceItems_Out, dblMoney) = False Then Exit Function
    If zlGetBalanceItemsFromCardFeeColl(strNO, cllBalance, lng异常ID, objBalanceItems_Out, IIf(bln作废, True, False)) = False Then Exit Function
    
    objBalanceItems_Out.业务ID = lng业务ID
    objBalanceItems_Out.异常ID = lng异常ID
    objBalanceItems_Out.同步状态 = int同步状态
    
    txt卡额.Text = objBalanceItems_Out.结算金额
    If mrsCardFee.RecordCount <> 0 Then
        txt卡号.Text = Nvl(mrsCardFee!卡号)
    End If

    If objBalanceItems_Out.类型 = gEM_记帐单 Then
        '记帐单
        mblnSendCardLocked = True
        chk记帐.value = 1
        Call SetCardEditEnabled
        ReadCardFeeBalanceDataFromNo = True
        Exit Function
    End If
    
    If str异常交易信息 <> "" And int同步状态 <= 2 Then
         If GetErrSwapInfoByJsonString(str异常交易信息, cllSwapinfo, cllExpends) Then
            '异常数据
             'cllSwapinfo(卡号,卡类别ID,交易流水号,交易说明,交易金额,二维码,支付方式,结算摘要)
            For i = 1 To objBalanceItems_Out.Count
               If cllSwapinfo("_支付方式")(1) <> "" Then objBalanceItems_Out(i).结算方式 = cllSwapinfo("_支付方式")(1)
               objBalanceItems_Out(i).卡类别ID = cllSwapinfo("_卡类别ID")(1)
               objBalanceItems_Out(i).交易流水号 = cllSwapinfo("_交易流水号")(1)
               objBalanceItems_Out(i).交易说明 = cllSwapinfo("_交易说明")(1)
               objBalanceItems_Out(i).卡号 = cllSwapinfo("_卡号")(1)
               objBalanceItems_Out(i).结算摘要 = cllSwapinfo("_结算摘要")(1)
               objBalanceItems_Out(i).QRCode = cllSwapinfo("_二维码")(1)
            Next
            Set objBalanceItems_Out.objTag = cllExpends
         End If
    End If
        
        
    If int同步状态 <> -1 Then
        Set objCurItem = objBalanceItems_Out(1)
        mblnNotClick = True
        Call SetLoaclePayModefromCard(objCurItem.objCard, False, True): mblnNotClick = False
    End If
    
    ReadCardFeeBalanceDataFromNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromCardFeeColl(ByVal strNO As String, ByVal cllCardFeeBalance As Collection, ByVal lng异常ID As Long, _
    ByRef objBalanceItems_Out As clsBalanceItems, _
    Optional ByVal bln查看作废 As Boolean, Optional blnDelFee As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据服务返回的集合,按记录集方式返回结算信息
    '入参:cllCardFeeBalance-当前集合
    '     strNo-费用单据号
    '     bln查看作废-当前查阅的是作废单据
    '     blnDelFee-当前为退费操作
    '出参:objBalanceItems_Out-返回的卡费结算信息集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    Dim objItem As clsBalanceItem, objCard As Card
    Dim dbl误差费 As Double
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    
    If cllCardFeeBalance Is Nothing Then Exit Function
    If objBalanceItems_Out.类型 = gEM_记帐单 Then zlGetBalanceItemsFromCardFeeColl = True: Exit Function
    
    
    
    '    blnc_mode   C   1   结算方式名称
    '    balance_id  N   1   结帐ID
    '    blnc_money  N   1   结帐金额
    '    pay_cardno  N   1   支付卡号
    '    pay_swapno  C   1   交易流水号
    '    pay_swapmemo    C   1   交易说明
    '    relation_id N   1   关联交易id
    '    cardtype_id N   1   卡类别id
    '    consume_card    N   1   是否消费卡:1-是;0-不是
    '    blnc_nature N   1   结算性质:1-现金结算方式,2-其他非医保结算 , 8-结算卡结算 ,9-误差费
    '    blnc_statu  N   1   结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
    '    consume_card_id N   1   消费卡id
    '    blnc_no C   1   结算号码
    '    blnc_memo   C   1   摘要
    
    objBalanceItems_Out.结算金额 = 0
    For i = 1 To cllCardFeeBalance.Count
        Set cllTemp = cllCardFeeBalance(i)
        Set objItem = New clsBalanceItem
        Set objCard = GetCardFromCardType(Val(Nvl(cllTemp("_cardtype_id"))), Val(Nvl(cllTemp("_consume_card"))) = 1, Nvl(cllTemp("_blnc_mode")))
        If Val(Nvl(cllTemp("_blnc_nature"))) = 9 Then
            dbl误差费 = RoundEx(dbl误差费 + Val(Nvl(cllTemp("_blnc_money"))), 6)
        Else
            With objItem
                Set .objCard = objCard
                .单据号 = strNO
                .单据性质 = 5   ' 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
                .结算方式 = Nvl(cllTemp("_blnc_mode"))
                .结算金额 = Val(Nvl(cllTemp("_blnc_money")))
                .关联交易ID = Val(Nvl(cllTemp("_relation_id")))
                .交易流水号 = Nvl(cllTemp("_pay_swapno"))
                .交易说明 = Nvl(cllTemp("_pay_swapmemo"))
                .结算号码 = Nvl(cllTemp("_blnc_no"))
                .结算性质 = Val(Nvl(cllTemp("_blnc_nature")))
                .结算摘要 = Nvl(cllTemp("_blnc_memo"))
                .卡号 = Nvl(cllTemp("_pay_cardno"))
                
                .卡类别ID = Val(Nvl(cllTemp("_cardtype_id")))
                .消费卡ID = Val(Nvl(cllTemp("_consume_card_id")))
                .消费卡 = Val(Nvl(cllTemp("_consume_card"))) = 1
                .是否密文 = objCard.卡号密文规则 <> ""
                .原始金额 = .结算金额
                .未退金额 = .结算金额
                .校对标志 = Val(Nvl(cllTemp("_blnc_statu")))
                .是否结算 = .校对标志 = 2 Or .校对标志 = 0
                .是否允许编辑 = Not .是否结算
                .是否允许删除 = .是否允许编辑
                .是否允许退现 = .是否允许编辑
                .密码 = ""
                .帐户余额 = 0
                If .卡类别ID = 0 Then   '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                      .结算类型 = 0
                ElseIf .卡类别ID <> 0 And .消费卡 = False Then
                      .结算类型 = 3
                ElseIf .卡类别ID <> 0 And .消费卡 Then
                      .结算类型 = 5
                Else
                     .结算类型 = 0
                End If
                .是否退款 = blnDelFee
                If bln查看作废 Then
                    .冲销ID = Val(Nvl(cllTemp("_balance_id")))
                    .结算ID = Val(Nvl(cllTemp("_original_id"))) '原结帐ID
                   
                Else
                    .结算ID = Val(Nvl(cllTemp("_balance_id")))
                    .冲销ID = Val(Nvl(cllTemp("_original_id"))) '原结帐ID
                End If
                .异常ID = lng异常ID
            
                .是否预交 = False
            End With
            objBalanceItems_Out.AddItem objItem
            objBalanceItems_Out.单据号 = objItem.单据号
            objBalanceItems_Out.结算金额 = RoundEx(objBalanceItems_Out.结算金额 + objItem.结算金额, 6)
            
            If objItem.卡类别ID <> 0 Then
                objBalanceItems_Out.类型 = IIf(objItem.消费卡, gEM_消费卡, gEM_一卡通)
            Else
                objBalanceItems_Out.类型 = gEM_普通结算
            End If
        End If
    Next
    objBalanceItems_Out.误差费 = dbl误差费
    objBalanceItems_Out.未退金额 = objBalanceItems_Out.结算金额
    objBalanceItems_Out.原始金额 = objBalanceItems_Out.结算金额 '暂定为未退部分
    objBalanceItems_Out.异常ID = lng异常ID
    zlGetBalanceItemsFromCardFeeColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Private Function GetBalanceItemsFromSwapInfo(ByVal str交易信息 As String, ByVal bln预交 As Boolean, ByVal dblMoney As Double, _
    ByRef objBalanceItems_Out As clsBalanceItems, Optional ByVal blnDefault As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据交易信息,获取当前的结算信息集
    '入参:dblMoney-当前金额
    '     str交易信息-交易信息,病人结算异常记录.交易信息
    '     blnDefault-无str=交易信息时，是否缺省,true-缺省;false-不缺省
    '出参:objBalanceItems_Out-当前结算信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-02-17 15:59:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurItem As clsBalanceItem
    Dim objCard As Card, cllSwapInfor As Collection, cllExpend As Collection
    Dim bln消费卡 As Long, lng卡类别ID As Long, str结算方式 As String
    On Error GoTo errHandle
    
    Set objBalanceItems_Out = New clsBalanceItems
    Set objCurItem = New clsBalanceItem
    If str交易信息 <> "" Then
        ' GetErrSwapInfoByJsonString
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:根据异常交易信息的Json串，获取异常信息
        '入参:
        '出参:cllSwapInfo_out-返回的交易信息:卡号,卡类别ID,交易流水号,交易说明,交易金额,二维码,支付方式,结算摘要
        '     cllExpend_out
        '          |-cllExpend:-交易名称,交易内容
        '           格式:array(名称,值),"_名称"
        '---------------------------------------------------------------------------------------------------------------------------------------------
        If GetErrSwapInfoByJsonString(str交易信息, cllSwapInfor, cllExpend) Then
            bln消费卡 = Val(cllSwapInfor("_是否消费卡")(1)) = 1
            lng卡类别ID = Val(cllSwapInfor("_卡类别ID")(1))
            str结算方式 = Trim(cllSwapInfor("_支付方式")(1))
            
            If lng卡类别ID <> 0 Then
                If mobjOneCardComLib.zlGetCard(lng卡类别ID, bln消费卡, objCard) = False Then Set objCard = Nothing
            ElseIf str结算方式 <> "" Then
               Set objCard = zlGetCardFromBalanceName(str结算方式)
            End If
            
            If Not objCard Is Nothing Then
                With objCurItem
                    Set .objCard = objCard
                    .卡类别ID = IIf(objCard.接口序号 < 0, 0, objCard.接口序号)
                    .消费卡 = objCard.消费卡
                    .结算方式 = IIf(str结算方式 = "", objCard.结算方式, str结算方式)
                    .结算金额 = dblMoney
                    .结算性质 = objCard.结算性质
                    .卡号 = Trim(cllSwapInfor("_卡号")(1))
                    .交易流水号 = Trim(cllSwapInfor("_交易流水号")(1))
                    .交易说明 = Trim(cllSwapInfor("_交易说明")(1))
                    .QRCode = Trim(cllSwapInfor("_二维码")(1))
                    .结算摘要 = Trim(cllSwapInfor("_结算摘要")(1))
                    If .卡类别ID > 0 Then
                       .结算类型 = IIf(.消费卡, 5, 3)
                    ElseIf objCard.结算性质 = 3 Then
                         .结算类型 = 2
                    Else
                       .结算类型 = 0  '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                    End If
                End With
                objBalanceItems_Out.AddItem objCurItem
                objBalanceItems_Out.结算金额 = objCurItem.结算金额
                objBalanceItems_Out.类型 = objCurItem.结算类型
                Set objBalanceItems_Out.objTag = cllExpend
                GetBalanceItemsFromSwapInfo = True
                Exit Function
            End If
        End If
    End If

    If Not blnDefault Then Exit Function
    
    If bln预交 Then
         Set objCard = GetDepositPayCard()
    Else
         Set objCard = GetCardFeePayCard()
    End If
    If objCard Is Nothing Then Exit Function
    
    With objCurItem
        Set .objCard = objCard
        .卡类别ID = IIf(objCard.接口序号 < 0, 0, objCard.接口序号)
        .消费卡 = objCard.消费卡
        .结算方式 = IIf(str结算方式 = "", objCard.结算方式, str结算方式)
        .结算金额 = dblMoney
        .结算性质 = objCard.结算性质
        .卡号 = ""
        .交易流水号 = ""
        .交易说明 = ""
        .QRCode = ""
        .结算摘要 = ""
        If .卡类别ID > 0 Then
           .结算类型 = IIf(.消费卡, 5, 3)
        ElseIf objCard.结算性质 = 3 Then
             .结算类型 = 2
        Else
           .结算类型 = 0  '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        End If
    End With
    objBalanceItems_Out.AddItem objCurItem
    objBalanceItems_Out.结算金额 = objCurItem.结算金额
    objBalanceItems_Out.类型 = objCurItem.结算类型
    GetBalanceItemsFromSwapInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReadCardAndDepositErrData(ByVal int操作状态 As Integer, Optional ByVal lng异常ID As Long, Optional ByRef objPatiInfo_out As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取卡费及预交数
    '入参:int操作状态 -0-增加;1-异常重收;2-异常作废
    '     lng异常ID-异常id
    '出参:objPatiInfo_out-返回的病人信息对象(异常单据的病人信息)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-28 10:20:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim cllDeposit As Collection, rsPatiInfo As ADODB.Recordset
    Dim bln相同 As Boolean, objItems As clsBalanceItems
    Dim strErrMsg As String, intSwapStatu As Integer
    Dim lng卡类别ID As Long, str交易信息 As String
    Dim str发卡卡号 As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    mlng异常ID = lng异常ID: mint操作状态 = int操作状态
    Call zlClearControlInfo '清除界面信息
    mblnShowDepositAndSendCard = True
  
    If int操作状态 = 0 Then
        zlReadCardAndDepositErrData = True: Exit Function
    End If
    strSql = "" & _
    "Select ID, 操作场景, 是否作废, 业务id, 是否病历费, 病人id, 主页id, 预交单号, 医疗卡单号, 卡类别id, 发卡卡号,预交金额,卡费金额, 同步状态, 交易信息, 登记时间, 操作员姓名 " & _
    "     From 病人结算异常记录 " & _
    "     Where ID =[1] "
    
    'int场景:1-医疗卡发卡;2-病人信息登记;3-病人入院 登记;4-预约挂号接收
    Set rsTemp = zlDatabase.OpenSQLRecordLob(strSql, Me.Caption, lng异常ID)
    If rsTemp.EOF Then
        MsgBox "读取异常数据失败，可能因并发原因被他人重收或作废，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    mbyt应用场景 = Val(Nvl(rsTemp!操作场景))
    mlngCardTypeID = Val(Nvl(rsTemp!卡类别ID))
    str交易信息 = Nvl(rsTemp!交易信息)
    str发卡卡号 = Nvl(rsTemp!发卡卡号)
    
    If InitFace = False Then Exit Function
  
    
    If mbyt应用场景 = 3 Then
        ',格式两种:一种是:病人id:主页ID,…;一种：病人id,…
        '查询类型:0-基本信息;1-基本信息的扩展;2-仅取主页
        If mobjService.ZlCissvr_GetPatiPageInfo(1, Val(Nvl(rsTemp!病人ID)) & ":" & Val(Nvl(rsTemp!主页ID)), rsPatiInfo, False) = False Then Exit Function
        If rsPatiInfo.RecordCount = 0 Then
            MsgBox "未获取到病人信息，请检查!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        Set mobjPati = New clsPatientInfo
        With rsPatiInfo
            mobjPati.病人ID = Val(Nvl(!病人ID))
            mobjPati.主页ID = Val(Nvl(!主页ID))
            mobjPati.住院号 = Trim(Nvl(!住院号))
            mobjPati.姓名 = Trim(Nvl(!姓名))
            mobjPati.性别 = Trim(Nvl(!性别))
            mobjPati.年龄 = Trim(Nvl(!年龄))
            mobjPati.费别 = Trim(Nvl(!费别))
            mobjPati.病人性质 = Val(Nvl(!病人性质))
            mobjPati.审核标志 = Val(Nvl(!审核标志))
            mobjPati.住院状态 = Val(Nvl(!住院状态))
            mobjPati.入院日期 = Trim(Nvl(!入院时间))
            mobjPati.出院日期 = Trim(Nvl(!出院时间))
            mobjPati.住院医师 = Trim(Nvl(!住院医师))
            mobjPati.医疗付款方式 = Trim(Nvl(!医疗付款方式名称))
            mobjPati.医疗付款方式编码 = Trim(Nvl(!医疗付款方式编码))
            mobjPati.当前病区id = Val(Nvl(!当前病区id))
            mobjPati.当前科室id = Val(Nvl(!当前科室id))
            mobjPati.医保号 = Trim(Nvl(!医保号))
            mobjPati.险类 = Trim(Nvl(!险类))
            mobjPati.床号 = Trim(Nvl(!当前床号))
            mobjPati.病人类型 = Trim(Nvl(!病人类型))
            mobjPati.学历 = Trim(Nvl(!学历))
            mobjPati.职业 = Trim(Nvl(!职业))
            mobjPati.国籍 = Trim(Nvl(!国籍))
            mobjPati.婚姻状况 = Trim(Nvl(!婚姻状况))
            mobjPati.编目日期 = Trim(Nvl(!编目日期))
            mobjPati.病人备注 = Trim(Nvl(!病人备注))
        End With
    Else
        If mobjOneCardComLib.zlGetPatiInforFromPatiID(Val(Nvl(rsTemp!病人ID)), mobjPati) = False Then Exit Function
    End If
     
    Set objPatiInfo_out = mobjPati  '返回病人信息
    If Nvl(rsTemp!预交单号) <> "" Then  '加载预交单据信息
         If Val(Nvl(rsTemp!同步状态)) = -1 Then
             '未产生费用
            If GetBalanceItemsFromSwapInfo(str交易信息, True, Val(Nvl(rsTemp!预交金额, 0)), mobjDepositItems, True) = False Then Exit Function
            
            For i = 1 To mobjDepositItems.Count
                 mobjDepositItems(i).单据性质 = 1
                 mobjDepositItems(i).单据号 = Nvl(rsTemp!预交单号)
                 mobjDepositItems(i).预交ID = 0
                 mobjDepositItems(i).异常ID = lng异常ID
                 
                 mobjDepositItems(i).结算时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
                 
                 mobjDepositItems(i).校对标志 = 1
                 mobjDepositItems(i).是否预交 = True
                 mobjDepositItems(i).是否保存 = True
                 mobjDepositItems(i).是否结算 = False
                 mobjDepositItems(i).是否允许编辑 = True
                 mobjDepositItems(i).是否允许删除 = True
                 mobjDepositItems(i).是否允许退现 = True
            Next
            mobjDepositItems.同步状态 = -1
            mobjDepositItems.未退金额 = Val(Nvl(rsTemp!预交金额, 0))
            mobjDepositItems.异常ID = lng异常ID
            mobjDepositItems.业务ID = Val(Nvl(rsTemp!业务ID))
            mobjDepositItems.是否保存 = True
            mobjDepositItems.单据号 = Nvl(rsTemp!预交单号)
            txt结算号码.Text = mobjDepositItems(1).结算号码
            If mintInsure = 0 Then
                txt缴款单位.Text = ""
                txt开户行.Text = ""
                txt帐号.Text = ""
                chk单位缴款.value = IIf(txt缴款单位.Text <> "", 1, 0)
            Else
                chk单位缴款.value = 0
            End If
            txt预交额.Text = Format(mobjDepositItems.结算金额, "0.00")
           
         Else
            If ReadDepositBalanceDataFromDepositNo(rsTemp!预交单号, lng异常ID, Val(Nvl(rsTemp!业务ID)), Val(Nvl(rsTemp!同步状态)), mobjDepositItems, Val(Nvl(rsTemp!是否作废)) = 1, str交易信息) = False Then Exit Function
            If Val(Nvl(rsTemp!同步状态)) >= 2 Then
               '接口调用成功的，则需要锁定
               mblnDepositLocked = True: SetDepositEditEnabled (1) '锁定结算信息
            End If
         End If
            
    Else
        tbDeposit.Visible = False
        mblnSendCardLocked = True: SetDepositEditEnabled (2) '锁定所有信息
    End If
    
    tbSendCard.Visible = False
    lbl卡名称.Visible = False
    fra磁卡.Visible = True
    
    If mlngCardTypeID = 0 Or Nvl(rsTemp!医疗卡单号) = "" Then
       mblnSendCardLocked = True: SetCardEditEnabled (2) '禁止所有信息
    Else
        fra磁卡.Caption = "【" & mCurSendCard.objSendCard.名称 & "】发卡"
        If Val(Nvl(rsTemp!同步状态)) = -1 Then
            If mobjDepositItems Is Nothing Then
                     '未产生费用:' 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
                    If GetBalanceItemsFromSwapInfo(str交易信息, True, Val(Nvl(rsTemp!卡费金额, 0)), mobjCardFeeItems, True) = False Then Exit Function
                    For i = 1 To mobjCardFeeItems.Count
                         mobjCardFeeItems(i).单据性质 = 5
                         mobjCardFeeItems(i).单据号 = Nvl(rsTemp!医疗卡单号)
                         mobjCardFeeItems(i).预交ID = 0
                         mobjCardFeeItems(i).异常ID = lng异常ID
                         mobjCardFeeItems(i).结算时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
                         
                         mobjCardFeeItems(i).校对标志 = 1
                         mobjCardFeeItems(i).是否预交 = True
                         mobjCardFeeItems(i).是否保存 = True
                         mobjCardFeeItems(i).是否结算 = False
                         mobjCardFeeItems(i).是否允许编辑 = True
                         mobjCardFeeItems(i).是否允许删除 = True
                         mobjCardFeeItems(i).是否允许退现 = True
                    Next
                
            Else
                    Set mobjCardFeeItems = mobjDepositItems.Clone
                    mobjCardFeeItems.结算金额 = 0
                    For i = 1 To mobjCardFeeItems.Count
                         mobjCardFeeItems(i).单据性质 = 5
                         mobjCardFeeItems(i).结算金额 = Val(Nvl(rsTemp!卡费金额, 0))
                         mobjCardFeeItems(i).单据号 = Nvl(rsTemp!医疗卡单号)
                         mobjCardFeeItems(i).预交ID = 0
                         mobjCardFeeItems(i).是否允许退现 = True
                         mobjCardFeeItems.结算金额 = mobjCardFeeItems.结算金额 + Val(Nvl(rsTemp!卡费金额, 0))
                    Next
            End If
            mobjCardFeeItems.同步状态 = -1
            mobjCardFeeItems.未退金额 = Val(Nvl(rsTemp!卡费金额, 0))
            mobjCardFeeItems.异常ID = lng异常ID
            mobjCardFeeItems.业务ID = Val(Nvl(rsTemp!业务ID))
            mobjCardFeeItems.单据号 = Nvl(rsTemp!医疗卡单号)
            If str发卡卡号 <> "" Then
                txt卡号.Text = str发卡卡号
            End If
            txt卡额.Text = Format(mobjCardFeeItems.结算金额, "0.00")
        Else
            If ReadCardFeeBalanceDataFromNo(rsTemp!医疗卡单号, lng异常ID, Val(Nvl(rsTemp!业务ID)), Val(Nvl(rsTemp!同步状态)), mobjCardFeeItems, Val(Nvl(rsTemp!是否作废)) = 1, str交易信息) = False Then Exit Function
        End If
    End If
    Call CalcRQCodePayTotal(True)  '计算扫码付总额
    If Val(Nvl(rsTemp!同步状态)) >= 4 Then
        Call SetCardEditEnabled(2): Call SetDepositEditEnabled(2)   '锁定结算方式
        
        mbtQRCodePay.Visible = False
        If Not mobjShowTotalMoneyControl Is Nothing Then mobjShowTotalMoneyControl.Visible = False
        
         If Not mobjDepositItems Is Nothing Then
            If mobjDepositItems.Count <> 0 Then
                 Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
            End If
         End If
         
         If Not mobjCardFeeItems Is Nothing Then
            If mobjCardFeeItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
            End If
         End If
         Call RefreshFactNo
         
        zlReadCardAndDepositErrData = True
        Exit Function
    ElseIf Val(Nvl(rsTemp!同步状态)) >= 2 Then
         mblnSendCardLocked = True: mblnDepositLocked = True
         Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '锁定结算方式
         mbtQRCodePay.Visible = False
        If Not mobjShowTotalMoneyControl Is Nothing Then mobjShowTotalMoneyControl.Visible = False
        
         If Not mobjCardFeeItems Is Nothing Then
            If mobjCardFeeItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
            End If
         End If
        
         If Not mobjDepositItems Is Nothing Then
            If mobjDepositItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
            End If
         End If
         Call RefreshFactNo
         zlReadCardAndDepositErrData = True
         Exit Function
    ElseIf Val(Nvl(rsTemp!同步状态)) = -1 Then
        '未产生费用
         If Not mobjCardFeeItems Is Nothing Then
            If mobjCardFeeItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
            End If
         End If
        
         If Not mobjDepositItems Is Nothing Then
            mblnNotClick = True
            If mobjDepositItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
            End If
            mblnNotClick = False
         End If
         Call RefreshFactNo
         zlReadCardAndDepositErrData = True
         Exit Function
    End If
    
    bln相同 = False
    If Nvl(rsTemp!预交单号) <> "" And Nvl(rsTemp!医疗卡单号) <> "" Then
        '医疗卡及预交同时发
        bln相同 = CheckDepsoitAndCardFeePayIsSame(mobjDepositItems, mobjCardFeeItems)
    End If
    
    If bln相同 Then
        '--相同
        If mobjDepositItems.类型 = gEM_一卡通 Then
            Call RefreshFactNo
            Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
            Set objItems = mobjCardFeeItems.Clone
            objItems.结算金额 = mobjDepositItems.结算金额
            objItems(1).结算金额 = RoundEx(objItems(1).结算金额 + mobjDepositItems.结算金额, 6)
            
            If mobjThirdSwap.zlThird_IsSwapIsSucces(objItems, intSwapStatu, strErrMsg, mobjDepositItems(1).预交ID) = False Then
                '交易失败
                'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                If intSwapStatu = 1 Then
                    mblnSendCardLocked = True: mblnDepositLocked = True
                    Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '锁定结算方式
                    Call SetLoaclePayModefromCard(objItems(1).objCard, False, True)
                    Call SetLoaclePayModefromCard(objItems(1).objCard, True, True)
                    
                End If
            Else
                '交易成功
                mblnSendCardLocked = True: mblnDepositLocked = True
                Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '锁定结算方式
                Call SetLoaclePayModefromCard(objItems(1).objCard, False, True)
                Call SetLoaclePayModefromCard(objItems(1).objCard, True, True)
            End If
        End If
    Else
        '1.预交
        If Not mobjDepositItems Is Nothing Then
            Call RefreshFactNo
            If mobjDepositItems.Count <> 0 Then
                If mobjDepositItems.类型 = gEM_一卡通 Then
                    Set mobjThirdSwap.objPayCards = mobjDepositPayCards
                           
                    If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjDepositItems, intSwapStatu, strErrMsg) = False Then
                        '交易失败
                        'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                        If intSwapStatu = 1 Then
                           mblnDepositLocked = True
                           Call SetDepositEditEnabled(1)   '锁定结算方式
                           Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                        End If
                    Else
                        '交易成功
                        mblnDepositLocked = True: Call SetDepositEditEnabled(1)   '锁定结算方式
                        Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                    End If
                End If
            End If
        End If
        '2.发卡
        If Not mobjCardFeeItems Is Nothing Then
            If mobjCardFeeItems.Count <> 0 Then
                If mobjCardFeeItems.类型 = gEM_一卡通 And mobjCardFeeItems.同步状态 <> -1 Then
                    Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
                    If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjCardFeeItems, intSwapStatu, strErrMsg) = False Then
                        '交易失败
                        'intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
                        If intSwapStatu = 1 Then
                            mblnSendCardLocked = True: Call SetCardEditEnabled(1)
                            Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                        End If
                    Else
                        '交易成功
                        mblnSendCardLocked = True:  Call SetCardEditEnabled(1)
                        Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                    End If
                End If
            End If
        End If
    End If
    
    
    Call Form_Resize
    zlReadCardAndDepositErrData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetCardFromCardType(ByVal lng卡类别ID As Long, bln消费卡 As Boolean, ByVal str结算方式 As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡类别ID获取卡对象
    '入参:lng卡类别ID-卡类别ID
    '     bln消费卡-是否消费卡
    '     str结算方式-结算方式
    '出参:
    '返回:成功卡对象
    '编制:刘兴洪
    '日期:2018-04-02 14:29:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As New Card
    On Error GoTo errHandle
    If lng卡类别ID <> 0 Then
        'zlGetCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean,  ByRef objCard As Card) As Boolean
        If mobjOneCardComLib.zlGetCard(lng卡类别ID, bln消费卡, objCard) = False Then
            Set objCard = zlGetCardFromBalanceName(str结算方式)
        End If
    Else
        Set objCard = zlGetCardFromBalanceName(str结算方式)
    End If
    Set GetCardFromCardType = objCard: Exit Function

    GetCardFromCardType = True
    Exit Function
errHandle:
    Set objCard = zlGetCardFromBalanceName(str结算方式)
    Set GetCardFromCardType = objCard: Exit Function
End Function


Private Function GetPatiInfoFromXML(ByVal strPatiXML As String, ByRef int信息更新模式_out As Integer, ByRef cllDrugInfos_Out As Collection, _
    ByRef cllImmuneInfos_Out As Collection, ByRef cllPatiExtInfo_out As Collection, ByRef cllWrangeInfo_out As Collection, _
    ByRef cllOtherPersons_Out As Collection, ByRef cllCertInfos_out As Collection, ByRef dictCardInfo_out As Dictionary) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '出参:cllDrugInfos_Out-过敏药物信息:array(过敏药物名称,过敏反应)
    '     cllImmuneInfos_Out-免疫信息:array(接种时间,疫苗名称)
    '     cllPatiExtInfo_out-从表信息,array(信息名,信息值 ),"_" & 信息名
    '     cllWrangeInfo_out-医学警示信息array(警示名称,信息值),"_警示名称",警示名称：医学警示,其他警示
    '     cllCertInfos_out-证件信息值:array(信息名,信息值 )
    '     dictCardInfo_out-医疗卡属性
    '     cllOtherPersons_Out-其他联系人信息集
    '       |-cllOtherPerson:联系人（姓名,关系,电话,身份证号) array(名称,值),"_名称"
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-08 21:52:04
    '目前未用，待以后扩展，不删除
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long
    Dim str过敏药物 As String, str过敏反应 As String
    Dim str接种日期 As String, str接种名称 As String
    Dim strABO血型 As String
    Dim str信息名 As String, str信息值 As String
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode
    Dim str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, str地址 As String
    Dim objPati As New clsPatientInfo
    Dim cllTemp As Collection
    On Error GoTo errHandle
        
    Set cllDrugInfos_Out = New Collection
    Set cllImmuneInfos_Out = New Collection
    Set cllPatiExtInfo_out = New Collection
    Set cllWrangeInfo_out = New Collection
    Set cllOtherPersons_Out = New Collection
    Set cllCertInfos_out = New Collection
    Set dictCardInfo_out = New Dictionary
    If strPatiXML = "" Then Exit Function
    
    '    信息更新模式 Integer 1 '0-强制更新，1-建档病人不更新，2-建档病人信息补缺
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strPatiXML, False) = False Then Exit Function
    Call zlXML_GetNodeValue("信息更新模式", , strValue): int信息更新模式_out = Val(strValue)
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    Call zlXML_GetNodeValue("卡号", , strValue):  objPati.卡号 = strValue
    '    姓名    Varchar2    64
    Call zlXML_GetNodeValue("姓名", , strValue):  objPati.姓名 = strValue
    '    性别    Varchar2    4
    Call zlXML_GetNodeValue("性别", , strValue):  objPati.性别 = strValue
    '    年龄    Varchar2    10
    Call zlXML_GetNodeValue("年龄", , strValue):  objPati.年龄 = strValue
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("出生日期", , strValue):  objPati.出生日期 = strValue
    '    出生地点    Varchar2    50
    Call zlXML_GetNodeValue("出生地点", , strValue):  objPati.出生地址 = strValue
    '    身份证号    VARCHAR2    18
    Call zlXML_GetNodeValue("身份证号", , strValue):  objPati.身份证号 = strValue
    '    其他证件    Varchar2    20
    Call zlXML_GetNodeValue("其他证件", , strValue):  objPati.其他证件 = strValue
    '    职业    Varchar2    80
    Call zlXML_GetNodeValue("职业", , strValue):  objPati.职业 = strValue
    '    民族    Varchar2    20
    Call zlXML_GetNodeValue("民族", , strValue):  objPati.民族 = strValue
    '    国籍    Varchar2    30
    Call zlXML_GetNodeValue("国籍", , strValue):  objPati.国籍 = strValue
    '    学历    Varchar2    10
    Call zlXML_GetNodeValue("学历", , strValue):  objPati.学历 = strValue
    '    婚姻状况    Varchar2    4
    Call zlXML_GetNodeValue("婚姻状况", , strValue):  objPati.婚姻状况 = strValue
    '    区域    Varchar2    30
    Call zlXML_GetNodeValue("区域", , strValue):  objPati.区域 = strValue
    '    家庭地址    Varchar2    50
    Call zlXML_GetNodeValue("家庭地址", , strValue):  objPati.家庭地址 = strValue
    '    户口地址    Varchar2    50
    Call zlXML_GetNodeValue("户口地址", , strValue):  objPati.户口地址 = strValue
     '    家庭电话    Varchar2    20
    Call zlXML_GetNodeValue("家庭电话", , strValue):  objPati.家庭电话 = strValue
    '    家庭地址邮编    Varchar2    6
    Call zlXML_GetNodeValue("家庭地址邮编", , strValue):  objPati.家庭邮编 = strValue
    '    监护人  Varchar2    64
    Call zlXML_GetNodeValue("监护人", , strValue):  objPati.监护人 = strValue
  
    '    联系人姓名  Varchar2    64
    Call zlXML_GetNodeValue("联系人姓名", , strValue):  objPati.联系人 = strValue
    '    联系人关系  Varchar2    30
    Call zlXML_GetNodeValue("联系人关系", , strValue):  objPati.联系人关系 = strValue
    '    联系人地址  Varchar2    50
    Call zlXML_GetNodeValue("联系人地址", , strValue):  objPati.联系人地址 = strValue
    '    联系人电话  Varchar2    20
    Call zlXML_GetNodeValue("联系人电话", , strValue):  objPati.联系人电话 = strValue
     '   工作单位    Varchar2    100
    Call zlXML_GetNodeValue("工作单位", , strValue):  objPati.工作单位 = strValue
    '    单位电话    Varchar2    20
    Call zlXML_GetNodeValue("单位电话", , strValue):  objPati.工作单位电话 = strValue
   '手机号   Varchar2    20
    Call zlXML_GetNodeValue("手机号", , strValue):  objPati.手机号 = strValue
    '    单位邮编    Varchar2    6
    Call zlXML_GetNodeValue("单位邮编", , strValue):  objPati.工作单位邮编 = strValue
    '    单位开户行  Varchar2    50
    Call zlXML_GetNodeValue("单位开户行", , strValue):  objPati.工作单位开户行帐户 = strValue
    '    单位帐号    Varchar2    20
    Call zlXML_GetNodeValue("单位帐号", , strValue):  objPati.工作单位开户行帐户 = strValue
    
    Call zlXML_GetRows("药物名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("药物名称", i, str过敏药物)
        Call zlXML_GetNodeValue("药物反应", i, str过敏反应)
        cllDrugInfos_Out.Add Array(str过敏药物, str过敏反应)
    Next
    
    lngCount = 0
    '免疫记录
    Call zlXML_GetRows("疫苗名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("疫苗名称", i, str接种名称)
        Call zlXML_GetNodeValue("接种时间", i, str接种日期)
        cllImmuneInfos_Out.Add Array(str接种日期, str接种名称)
    Next
    
    lngCount = 0
    'ABO血型
    Call zlXML_GetNodeValue("ABO血型", , strABO血型): cllPatiExtInfo_out.Add Array("ABO血型", strValue), "_ABO血型"
    'RH
    Call zlXML_GetNodeValue("RH", , strValue): cllPatiExtInfo_out.Add Array("RH", strValue), "_RH"
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
    If strValue <> "" Then strValue = Mid(strValue, 2)
    cllWrangeInfo_out.Add Array("医学警示", strValue), "_医学警示"
    '其他医学警示
    Call zlXML_GetNodeValue("其他医学警示", , strValue): cllWrangeInfo_out.Add Array("其他警示", strValue), "_其他警示"
    '联系信息
    '    联系人地址  Varchar2    50
    Call zlXML_GetNodeValue("联系人地址", , str地址): objPati.联系人地址 = str地址
  
     '    联系人姓名  Varchar2    64
    Call zlXML_GetNodeValue("联系人姓名", , str姓名): objPati.联系人 = str姓名
    '    联系人关系  Varchar2    30
    Call zlXML_GetNodeValue("联系人关系", , str关系): objPati.联系人关系 = str关系
    '    联系人电话  Varchar2    20
    Call zlXML_GetNodeValue("联系人电话", , str电话): objPati.联系人电话 = str电话
    '    联系人身份证 Varchar2   20
    Call zlXML_GetNodeValue("联系人身份证号", , str身份证号): objPati.联系人电话 = str身份证号
    Call zlXML_GetRows("联系信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("联系信息", "姓名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Set cllTemp = New Collection
                Call zlXML_GetChildNodeValue("联系信息", "姓名", i, j, str姓名): cllTemp.Add Array("姓名", str姓名), "_姓名"
                Call zlXML_GetChildNodeValue("联系信息", "关系", i, j, str关系): cllTemp.Add Array("关系", str关系), "_关系"
                Call zlXML_GetChildNodeValue("联系信息", "电话", i, j, str电话): cllTemp.Add Array("电话", str电话), "_电话"
                Call zlXML_GetChildNodeValue("联系信息", "身份证号", i, j, str身份证号): cllTemp.Add Array("身份证号", str身份证号), "_身份证号"
                cllOtherPersons_Out.Add cllTemp
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '其他信息
    '健康档案编号
    Call zlXML_GetNodeValue("健康档案编号", , strValue): cllPatiExtInfo_out.Add Array("健康档案编号", strValue), "_健康档案编号"
    '新农合证号
    Call zlXML_GetNodeValue("新农合证号", , strValue): cllPatiExtInfo_out.Add Array("新农合证号", strValue), "_新农合证号"

    '其他证件
    Call zlXML_GetRows("其他证件", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他证件", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他证件", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他证件", "信息值", i, j, str信息值)
                cllCertInfos_out.Add Array(str信息名, str信息值)
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
                cllPatiExtInfo_out.Add Array(str信息名, str信息值), "_" & str信息值
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
                If dictCardInfo_out.Exists(str信息名) Then
                    dictCardInfo_out.Item(str信息名) = str信息值
                Else
                    dictCardInfo_out.Add str信息名, str信息值
                End If
            Next
        End If
    Next
    GetPatiInfoFromXML = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

