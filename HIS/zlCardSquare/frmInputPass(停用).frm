VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmInputPass 
   BorderStyle     =   0  'None
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPati 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   15
      ScaleHeight     =   765
      ScaleWidth      =   7515
      TabIndex        =   15
      Top             =   0
      Width           =   7515
      Begin VB.Frame fraSplitPati 
         Height          =   90
         Left            =   0
         TabIndex        =   17
         Top             =   675
         Width           =   7515
      End
      Begin VB.CheckBox chk退现 
         Caption         =   "退现"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6270
         TabIndex        =   16
         Top             =   435
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   420
         Width           =   600
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "请刷提货卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   5745
         TabIndex        =   20
         Top             =   15
         Width           =   1635
      End
      Begin VB.Label lblMargin 
         BackStyle       =   0  'Transparent
         Caption         =   "未刷卡额:3000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2955
         TabIndex        =   19
         Top             =   405
         Width           =   1800
      End
      Begin VB.Label lblPatiName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:张三"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   18
         Top             =   75
         Width           =   1155
      End
   End
   Begin VB.PictureBox picPassWord 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   7515
      TabIndex        =   4
      Top             =   795
      Width           =   7515
      Begin VB.CommandButton cmdOK 
         Caption         =   "完成(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   9
         Top             =   45
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txt卡号 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   8
         Top             =   60
         Width           =   4350
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1455
         TabIndex        =   7
         Top             =   1170
         Width           =   4305
      End
      Begin VB.TextBox txtMoney 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1455
         TabIndex        =   5
         Top             =   645
         Width           =   1740
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   6
         Top             =   630
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl密码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   765
         TabIndex        =   14
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label lbl卡号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   735
         TabIndex        =   13
         Top             =   90
         Width           =   570
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "本次消费"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   735
         Width           =   1140
      End
      Begin VB.Label lblBalance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "余额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3495
         TabIndex        =   11
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblBalanceMoney 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   4035
         TabIndex        =   10
         Top             =   645
         Width           =   1740
      End
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   2235
   End
   Begin VB.PictureBox picBlance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   1710
      Left            =   0
      ScaleHeight     =   1680
      ScaleWidth      =   7470
      TabIndex        =   0
      Top             =   2610
      Visible         =   0   'False
      Width           =   7500
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
         Left            =   8745
         TabIndex        =   1
         Top             =   60
         Width           =   1080
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBlance 
         Height          =   1545
         Left            =   0
         TabIndex        =   2
         Top             =   -15
         Width           =   7470
         _cx             =   13176
         _cy             =   2725
         Appearance      =   3
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInputPass.frx":0000
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
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   6555
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7785
      _Version        =   589884
      _ExtentX        =   13732
      _ExtentY        =   11562
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin VB.Shape shpRange 
      BorderStyle     =   6  'Inside Solid
      Height          =   330
      Left            =   7890
      Top             =   4995
      Width           =   210
   End
End
Attribute VB_Name = "frmInputPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'--入口参数
Private mdbl本次消费总额 As Double, mlngModule As Long
Private mlngCardTypeID As Long, mbln消费卡 As Boolean
Private mstrOutCardNo As String, mstrOutPassWord As String
Private mbln退费 As Boolean, mbln退现 As Boolean
Private mstrPatiName As String, mstrSex As String, mstrOld As String
Private mblnShowclsPatientInfo As Boolean
Private mcurCardObject As clsCardObject
Private mstr费用来源 As String, mlng病人ID As Long
'---------------------------------------------------------------------
Private mstrCardNo As String, mstrPassWord As String
Private mstrOldCardNo As String '旧的卡号
Private mblnFirst As Boolean, mblnOk As Boolean
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mintPassInputCount As Integer   '密码输入统计
Private mdbl帐户余额 As Double, mstr限制类别 As String
Private mrsClassMoney As ADODB.Recordset
Private msngOldX As Single, msngOldY As Single
Private mblnReadCard As Boolean
Private mblnPassInputCardNo As Boolean '是否密文输入卡号
Private mobjKeyboard As Object '建盘输入对象
Private mdbl本次扣款额 As Double
Private mbln余额不足禁止 As Boolean
Private mlng消费卡ID As Long
Private mdbl已刷总额 As Double
Private mbln转预交 As Boolean
Private mstrCurFeeType As String '当前的收费类别
Private mblnAllPay As Boolean
Private mblnPosPass As Boolean
'刷卡结算信息,确定时返回(只限本次刷卡的数据);如果是退费时,传入退费的原始数据
Private mVarData As Variant 'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mobjSquare As Object
'---------------------------------------------------------------------
Public Function zlBrushPay(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal objCardObject As clsCardObject, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByRef dbl本次刷卡金额 As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowclsPatientInfo As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varBrushCardData As Variant = Nothing, _
    Optional ByVal bln转预交 As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal str费用来源 As String, _
    Optional ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡处理
    '入参:frmMain-调用的主窗体
    '        lngCardTypeID-卡类别ID(0-表示只刷一卡通)
    '        rsClassMoney-收费类别,实收金额
    '        blnShowclsPatientInfo-是否显示病人信息
    '       dbl金额-传入需要扣款的金额
    '       bln余额不足禁止-余额不足时,禁止继续操作,否则表示提示用余额支付
    '       VarBrushCardData-Collection类型,已经刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文,剩余未退金额(只针对退费)))
    '       blnAllPay-是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
    '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '       lng病人ID - 病人ID(使用消费卡支付时传入)
    '出参:strCardNO-返回卡号
    '        strPassWord-返回输入的密码
    '        bln退现-是否将当前的刷卡金额部分退现
    '        dbl金额-返回本次的扣款金额
    '        str限制类别 -限制类别(消费卡返回)
    '        lng消费卡ID-消费卡信息.ID(消费卡返回)
    '       varBrushCardData-Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-10 12:54:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intMouse As Integer
    Screen.MousePointer = 0: mlngCardTypeID = lngCardTypeID
    mbln消费卡 = bln消费卡: mdbl本次消费总额 = dbl本次刷卡金额
    mbln转预交 = bln转预交: mblnAllPay = blnAllPay
    Set mrsClassMoney = rsClassMoney
    mblnShowclsPatientInfo = blnShowclsPatientInfo: mbln退费 = bln退费
    mstrPatiName = strPatiName
    mstrSex = strSex: mstrOld = strOld: mbln退现 = False
    If IsEmpty(varBrushCardData) Then
        Set mVarData = Nothing
    Else
        Err = 0: On Error Resume Next
        Set mVarData = varBrushCardData '已经刷卡数据
        If Err <> 0 Then
            Set mVarData = Nothing
        End If
    End If
    mstrOldCardNo = strCardNo: mbln余额不足禁止 = bln余额不足禁止
    Set mcurCardObject = objCardObject
    mblnOk = False: intMouse = Screen.MousePointer
    strCardNo = ""
    mlngModule = lngModule
    mstr费用来源 = str费用来源: mlng病人ID = lng病人ID
    
    On Error GoTo 0
    'IC卡对象
    On Error Resume Next
    'Set mobjICCard = CreateObject("zlICCard.clsICCard")
    On Error GoTo 0
    Me.Show 1, frmMain
    zlBrushPay = mblnOk
    strCardNo = mstrCardNo: strPassWord = mstrPassWord
    dbl本次刷卡金额 = mdbl本次扣款额
    Set varBrushCardData = mVarData '返回刷卡数据
    Screen.MousePointer = intMouse
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = intMouse
End Function

Private Sub InitTaskPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化InitTaskPancel
    '编制:刘兴洪
    '日期:2011-06-30 18:20:30
    '问题:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    If mblnShowclsPatientInfo Then
        Call wndTaskPanel.SetGroupInnerMargins(0, 2, 0, 0)
    Else
        picPati.Visible = False
    End If
    
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    Set tkpGroup = wndTaskPanel.Groups.Add(1, "请刷卡后输入密码")
    If mblnShowclsPatientInfo Then
        Set Item = tkpGroup.Items.Add(101, "", xtpTaskItemTypeControl)
        Set Item.Control = picPati
        picPati.BackColor = Item.BackColor
        Call Item.SetMargins(0, -19, 0, IIf(mblnShowclsPatientInfo, -10, -4))
    End If
    Set Item = tkpGroup.Items.Add(102, "", xtpTaskItemTypeControl)
    Set Item.Control = picPassWord
    tkpGroup.CaptionVisible = False
    If mblnShowclsPatientInfo Then
        Call Item.SetMargins(0, 20, 0, 0)
    Else
        Call Item.SetMargins(0, -19, 0, -4)
    End If
    picPassWord.BackColor = Item.BackColor
    'picPati.BackColor = Item.BackColor
    fraSplitPati.BackColor = Item.BackColor
    chk退现.BackColor = Item.BackColor
    tkpGroup.Expandable = False
    vsBlance.BackColor = Item.BackColor
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
End Sub
Private Sub AddWndTaskPancelExpend()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加刷卡信息列表的扩展部分
    '编制:刘兴洪
    '日期:2013-02-22 15:53:55
    '问题:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    On Error GoTo errHandle
    
    Set tkpGroup = wndTaskPanel.Groups.Find(2)
    '存在，就退出
    If Not tkpGroup Is Nothing Then Exit Sub
    picBlance.Visible = True
    Set tkpGroup = wndTaskPanel.Groups.Add(2, "当前刷卡信息")
    Set Item = tkpGroup.Items.Add(201, "", xtpTaskItemTypeControl)
    Set Item.Control = picBlance
    'Call Item.SetMargins(0, -19, 0, IIf(mblnShowclsPatientInfo, -10, -4))
    tkpGroup.CaptionVisible = True
    tkpGroup.Expandable = True
    picBlance.BackColor = Item.BackColor
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
    '设置完成和取消的显示
    cmdOK.Visible = True: cmdCancel.Visible = True
    Call SetWindowHeight
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetWindowHeight()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体的高度
    '编制:刘兴洪
    '日期:2013-02-22 15:55:37
    '问题:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngHeight As Single, sngSplit As Single
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    '调整位置,用On Error resume next 屏蔽可能出现的错误
    Err = 0: On Error Resume Next
    sngSplit = 700
    sngHeight = picPassWord.Height + sngSplit
    If mblnShowclsPatientInfo Then
        sngHeight = sngHeight + picPati.Height
    End If
    Set tkpGroup = wndTaskPanel.Groups.Find(2)
    If tkpGroup Is Nothing Then Me.Height = sngHeight: Exit Sub
    If tkpGroup.Expanded Then
        sngHeight = sngHeight + picBlance.Height + IIf(mblnShowclsPatientInfo = False, 200, 0)
    End If
    sngHeight = sngHeight + 550
    Me.Height = sngHeight
End Sub

Private Sub cmdCancel_Click()
    Dim cllBalance As Collection
    Set mVarData = cllBalance
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txt卡号 <> "" Then
        If Not zlSquareAffirm(True) Then Exit Sub
    End If
    Call SetReturnBrushCardInfor
    mblnOk = True
    Unload Me
End Sub
Private Function SetReturnBrushCardInfor() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置返回的刷卡信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-02-25 15:29:54
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBrushCardInfor As Collection, lngCardTypeID As Long, dblMoney As Double
    Dim lng消费卡ID As Long, strCardNo As String, strPassWord As String, str限制类别 As String
    Dim int密文 As Integer
    Dim i As Long
    On Error GoTo errHandle
    Set cllBrushCardInfor = New Collection
    With vsBlance
        mdbl本次扣款额 = 0
        For i = .Rows - 1 To 1 Step -1
            strCardNo = Trim(.Cell(flexcpData, i, .ColIndex("卡号")))
            If strCardNo <> "" And Val(.RowData(i)) = 0 Then
                lngCardTypeID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                lng消费卡ID = Val(.TextMatrix(i, .ColIndex("消费卡ID")))
                strPassWord = Trim(.TextMatrix(i, .ColIndex("密码")))
                str限制类别 = Trim(.TextMatrix(i, .ColIndex("限制类别")))
                dblMoney = Val(.TextMatrix(i, .ColIndex("刷卡金额")))
                int密文 = Val(.TextMatrix(i, .ColIndex("卡号密文显示")))
                '返回最后一次的刷卡信息
                mlngCardTypeID = lngCardTypeID
                mstrCardNo = strCardNo
                mbln消费卡 = True
                mlng消费卡ID = lng消费卡ID
                mstr限制类别 = str限制类别
                mdbl本次扣款额 = mdbl本次扣款额 + dblMoney
                'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
                cllBrushCardInfor.Add Array(lngCardTypeID, lng消费卡ID, dblMoney, strCardNo, strPassWord, str限制类别, int密文)
            End If
        Next
    End With
    Set mVarData = cllBrushCardInfor
    SetReturnBrushCardInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me: Exit Sub
    If KeyCode = vbKeyF2 And lbl卡号.BorderStyle = 1 Then ClickReadCard: Exit Sub
End Sub

Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始窗体相关控件信息
    '编制:刘兴洪
    '日期:2013-02-22 16:02:31
    '问题:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstr限制类别 = "": mlng消费卡ID = 0
    mdbl本次扣款额 = mdbl本次消费总额   '先初始化
    mstrCardNo = "": mstrPassWord = "":   mstrCurFeeType = Get收费限制类别_名称
    picPati.Visible = mblnShowclsPatientInfo
    lblPatiName.Caption = "姓名:" & mstrPatiName
    lblSex.Caption = "性别:" & mstrSex
    lblMoney.Caption = IIf(mbln退费, "本次退款", "本次消费")
    lblMoney.ForeColor = IIf(mbln退费, vbRed, lbl密码.ForeColor)
    lblMargin.Visible = Not mbln退费
    If Not mcurCardObject Is Nothing Then
        chk退现.Visible = mbln退费 And mcurCardObject.CardPreporty.是否退现 = 1
        lblType.Caption = "请刷" & mcurCardObject.CardPreporty.名称
    Else
        lblType.Caption = "": chk退现.Visible = False
    End If
    If mlngCardTypeID = 0 Then
        txt卡号.Locked = True: lbl卡号.BorderStyle = 1
        lbl密码.Enabled = False: txtPass.Enabled = False
    End If
    cmdOK.Enabled = False
    txtMoney.Enabled = False
    Call LoadBruhCardInfor
    Call ShowMoney
    
    '初始化参数 :276-消费卡刷卡消费须定位到密码框
    mblnPosPass = Val(zlDatabase.GetPara(Val("276-消费卡刷卡消费须定位到密码框"), glngSys, , "1")) = 1
End Sub
Private Sub LoadBruhCardInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载已经刷卡的信息
    '编制:刘兴洪
    '日期:2013-02-25 15:58:01
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As Collection, strCardNo As String
    Dim i As Long
    Dim lngRow As Long
    
    On Error GoTo errHandle
    mdbl已刷总额 = 0
    With vsBlance
        .Rows = 2
        .Clear 1
        .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .RowData(1) = 0
        If IsEmpty(mVarData) Then Exit Sub
        If mVarData Is Nothing Then Exit Sub
        
        Err = 0: On Error Resume Next
        Set cllBalance = mVarData
        If Err <> 0 Then
            Err = 0: On Error GoTo 0: Exit Sub
        End If
        Err = 0: On Error GoTo errHandle:
        lngRow = 1
        For i = 1 To cllBalance.Count
            'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文,剩余退款金额)
            If Val(cllBalance(i)(2)) <> 0 Then
                .RowData(lngRow) = 1
                .TextMatrix(lngRow, .ColIndex("卡类别ID")) = Val(cllBalance(i)(0))
                .TextMatrix(lngRow, .ColIndex("消费卡ID")) = Val(cllBalance(i)(1))
                .TextMatrix(lngRow, .ColIndex("刷卡金额")) = Format(Val(cllBalance(i)(2)), "0.00")
                mdbl已刷总额 = mdbl已刷总额 + Val(cllBalance(i)(2))
                .TextMatrix(lngRow, .ColIndex("卡号密文显示")) = Val(cllBalance(i)(6))
                strCardNo = Trim(cllBalance(i)(3))
                If Val(cllBalance(i)(6)) = 1 Then
                    .TextMatrix(lngRow, .ColIndex("卡号")) = String(Len(strCardNo), "*")
                Else
                    .TextMatrix(lngRow, .ColIndex("卡号")) = strCardNo
                End If
                .Cell(flexcpData, lngRow, .ColIndex("卡号")) = strCardNo
                .TextMatrix(lngRow, .ColIndex("密码")) = Trim(cllBalance(i)(4))
                .TextMatrix(lngRow, .ColIndex("限制类别")) = Trim(cllBalance(i)(5))
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
        Next
        If .Rows > 2 Then
            .Rows = .Rows - 1
        End If
    End With
    
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ShowMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示支付金额
    '编制:刘兴洪
    '日期:2012-02-24 14:19:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double '已刷额
    On Error GoTo errHandle
    dblMoney = GetBruhMoney
    If mdbl本次消费总额 <> dblMoney And mblnAllPay Then
        cmdOK.Visible = False
    End If
    txtMoney.Text = Format(mdbl本次扣款额, "0.00")
    dblMoney = mdbl本次消费总额 + mdbl已刷总额 - mdbl本次扣款额 - dblMoney
    lblMargin.Caption = "剩余未刷卡金额:" & Format(dblMoney, "0.00")
    lblMargin.AutoSize = True
    lblMargin.Visible = dblMoney <> 0
    lblBalanceMoney.Caption = Format(mdbl帐户余额, "0.00")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    msngOldX = 0: msngOldY = 0
    Call InitBalanceGrid '清除刷卡数据
    If mlngCardTypeID = 0 Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    Call InitFace
    txtPass.PasswordChar = "*"
    Call CreateObjectKeyboard
    Call InitTaskPancel
    Err = 0: On Error Resume Next
    If Not mVarData Is Nothing Then
        If mVarData.Count <> 0 Then
            Call AddWndTaskPancelExpend
            Call ShowWndCaption
        End If
    End If
    
    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    
    Call SetWindowHeight
    mblnFirst = True
    'lbl帐户余额.Caption = "帐户余额:"
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With wndTaskPanel
        .Left = ScaleLeft: .Top = ScaleTop
        .Height = ScaleHeight: .Width = ScaleWidth
    End With
End Sub

Private Sub lbl卡号_Click()
    Dim strExpand As String, strCardNo As String, strOutXml As String
    If Not mcurCardObject.CardPreporty.是否接触式读卡 Then Exit Sub
'    If mobjICCard Is Nothing Then
'        Set mobjICCard = CreateObject("zlICCard.clsICCard")
'        Set mobjICCard.gcnOracle = gcnOracle
'    End If
    
'    If Not mobjICCard Is Nothing Then
'        txt卡号.Text = mobjICCard.Read_Card()
'        If txt卡号.Text <> "" Then
'            mblnICCard = True
'            Call CheckFreeCard(txt卡号.Text)
'        End If
'    End If
  
    If mcurCardObject.CardObject Is Nothing Then Exit Sub
    If gobjOneCardComLib.objThirdSwap.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    txt卡号.Text = Trim(strCardNo)
    If txt卡号.Text <> "" Then
        If Not CheckBrush消费卡(Trim(strCardNo)) Then
            Call txt卡号_GotFocus
            If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
            Exit Sub
        End If
    End If
    txt卡号.Tag = Trim(strCardNo)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt卡号.Text = Trim(strCardNo)
    If txt卡号.Text <> "" Then
        If Not CheckBrush消费卡(txt卡号.Text) Then
            Call txt卡号_GotFocus
            If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
            Exit Sub
        Else
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
        End If
    End If
    txt卡号.Tag = Trim(strCardNo)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub picBlance_Click()
    Debug.Print "DD"
End Sub

Private Sub picPassWord_Click()
    Debug.Print "DD"
End Sub

Private Sub picPassWord_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    msngOldX = X: msngOldY = Y
End Sub
Private Sub picPassWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button <> 1 Then Exit Sub
        Me.Left = Me.Left + Me.ScaleLeft - msngOldX + X
        Me.Top = Me.Top + Me.ScaleTop - msngOldY + Y
End Sub
 
Private Sub picPassWord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngOldX = 0: msngOldY = 0
End Sub
'--------------------------------------------------------

Private Function SetBrushObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡对象
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-10 13:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand
    tmrMain.Tag = "": mblnPassInputCardNo = False
    '一卡通,直接退出
    If mlngCardTypeID = 0 Then SetBrushObject = True: Exit Function
    If mcurCardObject.CardObject Is Nothing Then
        MsgBox "注意:" & vbCrLf & "   未找到相关的三方接口,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Not mcurCardObject.InitCompents Then
        If mcurCardObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") Then
              Exit Function
        End If
        mcurCardObject.InitCompents = True
    End If
    mblnPassInputCardNo = mcurCardObject.CardPreporty.卡号密文规则 <> "" And mcurCardObject.CardPreporty.卡号密文规则 <> "0"
    Me.Caption = mcurCardObject.CardPreporty.名称
    With mcurCardObject.CardPreporty
        Me.txt卡号.MaxLength = .卡号长度
        If .是否自动读取 = 1 Then
            tmrMain.Interval = IIf(.自动读取间隔 = 0, 300, .自动读取间隔)
            tmrMain.Tag = 1
        End If
    End With
    '支持刷卡或读卡
    '85565,李南春,2015/7/10:读卡性质
    If mcurCardObject.CardPreporty.是否接触式读卡 Then
        lbl卡号.BorderStyle = 1
    Else
        lbl卡号.BorderStyle = 0
    End If
    txt卡号.Locked = Not (mcurCardObject.CardPreporty.是否刷卡 Or mcurCardObject.CardPreporty.是否扫描)
    'If cmdRead.Visible = False Then txt卡号.Width = txtPass.Width
    SetBrushObject = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证数据的合法性
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-10 14:02:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Trim(txt卡号.Text) = "" Then
        MsgBox "卡号未输入,请输入卡号或刷卡!", vbInformation, gstrSysName
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
        Exit Function
    End If
    
    If Trim(txt卡号.Tag) = "" Then
        MsgBox "还未验证卡片,请在卡号处点击回车重新验卡!", vbInformation, gstrSysName
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
        Exit Function
    End If
    
    If Not mbln消费卡 Then isValied = True: Exit Function
    
    '非自制卡,密码用不着输入,可以在支付接口中进行验证密码(一般来说,都是封装了的)
    If mcurCardObject.自制卡 = False Then Exit Function
    
   ' If CheckBrush消费卡(Trim(txt卡号.Tag)) = False Then Exit Function
    
    If zlCommFun.zlStringEncode(txtPass.Text) <> txtPass.Tag Then
        MsgBox "密码输入错误！", vbExclamation, gstrSysName
        txtPass.Text = "": mintPassInputCount = mintPassInputCount - 1
        If mintPassInputCount > 2 Then Unload Me: Exit Function
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
        Exit Function
    End If
    If Not mbln退费 Then
        '检查可用余额是否够支付
        If Round(mdbl帐户余额, 6) < Round(mdbl本次扣款额, 6) Then
            MsgBox "当前支付金额(" & Format(mdbl本次消费总额, "0.00") & ")大于了帐户余额(" & Format(mdbl帐户余额, "0.00") & "),不能继续!", vbInformation, gstrSysName
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Exit Function
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ClickReadCard()
    If ReadCardNo = False Then
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
        Exit Sub
    End If
    If mlngCardTypeID = 0 Then Unload Me: Exit Sub
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    If SetBrushObject = False Then Unload Me: Exit Sub
    If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    If lbl卡号.BorderStyle = 1 Then txt卡号.ToolTipText = "按F2或回车进行读卡"
'
'    If cmdRead.Visible = False Then
'        Me.Width = Me.Width - cmdRead.Width * 0.3
'        picPassWord.Width = picPassWord.Width - cmdRead.Width * 0.3
'     End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set mobjICCard = Nothing
    UnHookKBD
    Set mcurCardObject = Nothing
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    Set mobjCommEvents = Nothing
    Set mrsClassMoney = Nothing
End Sub

Private Sub picPassWord_Resize()
    Err = 0: On Error Resume Next
 
End Sub

Private Sub picPati_Resize()
    Err = 0: On Error Resume Next
    lblType.Left = picPati.ScaleWidth - lblType.Width - 20
End Sub
'------------------------------------------------------------------------------------------------------------
Private Sub picPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    msngOldX = X: msngOldY = Y
End Sub
Private Sub picPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button <> 1 Then Exit Sub
        Me.Left = Me.Left + Me.ScaleLeft - msngOldX + X
        Me.Top = Me.Top + Me.ScaleTop - msngOldY + Y
End Sub
 
Private Sub picPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngOldX = 0: msngOldY = 0
End Sub
'------------------------------------------------------------------------------------------------------------

Private Sub tmrMain_Timer()
    If mblnReadCard = False Then
        mblnReadCard = True
        Call ReadCardNo
        mblnReadCard = False
    End If
End Sub

Private Sub txtPass_LostFocus()
    Call ClosePassKeyboard(txtPass)
End Sub
Private Sub txt卡号_Change()
    txtPass.Enabled = txt卡号.Text <> ""
    txt卡号.Tag = "": mstrCardNo = "'"   ' lbl帐户余额.Caption = "帐户余额:"
 
    If Not txtPass.Enabled Then txtPass.Text = ""
    tmrMain.Enabled = Val(tmrMain.Tag) <> 0 And Trim(txt卡号.Text) = ""
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txt卡号.Text = "")
End Sub
Private Sub txt卡号_GotFocus()
    Dim strExpend As String
    
    On Error GoTo Errhand
    Call zlControl.TxtSelAll(txt卡号)
    If Not mobjICCard Is Nothing And mlngCardTypeID = 0 Then Call mobjICCard.SetEnabled(True)
    tmrMain.Enabled = Val(tmrMain.Tag) <> 0 And Trim(txt卡号.Text) = ""
    
    If mobjSquare Is Nothing Then
        Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
        '初始化射频卡对象
        If Err <> 0 Then Exit Sub
        mobjSquare.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    End If
    If mcurCardObject.CardPreporty.是否非接触式读卡 Then mobjSquare.SetEnabled True
    '85565:李南春,2015/7/21,调用刷卡接口
    Err = 0: On Error Resume Next
    
    If mcurCardObject.CardPreporty.接口序号 = 0 Or mcurCardObject.CardPreporty.接口程序名 = "" Then Exit Sub
    If Not (mcurCardObject.CardPreporty.是否刷卡 Or mcurCardObject.CardPreporty.是否扫描) Then Exit Sub
    
    Call mobjSquare.zlSetBrushCardObject(mcurCardObject.CardPreporty.接口序号, txt卡号, strExpend, _
                                        mcurCardObject.CardPreporty.消费卡)
    If mobjCommEvents Is Nothing Then Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    mobjSquare.zlInitEvents Me.hWnd, mobjCommEvents
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean, lng病人ID As Long
    Dim strCardNo As String
    
    If KeyAscii = 13 And Trim(txt卡号.Text) = "" Then
        If lbl卡号.BorderStyle = 1 Then
            KeyAscii = 0
            txt卡号.PasswordChar = IIf(mblnPassInputCardNo, "*", "")
            Call ClickReadCard: Exit Sub
        ElseIf cmdOK.Visible And cmdOK.Enabled Then
            cmdOK.SetFocus: Exit Sub
        End If
    End If
    If txt卡号.Locked Or txt卡号.Enabled = False Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    blnCard = zlCommFun.InputIsCard(txt卡号, KeyAscii, mblnPassInputCardNo)
    txt卡号.PasswordChar = IIf(mblnPassInputCardNo, "*", "")
'
'    If lbl卡号.BorderStyle = 1 Then
'        '只能读卡，不能接收输入
'        If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
'    End If
    
    If Not (blnCard And Len(txt卡号.Text) = txt卡号.MaxLength - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txt卡号.Text) <> "") Then
        '问题:51570
        '不是刷卡和回车,则退出
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = timer
            If txt卡号.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txt卡号.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txt卡号.Text = Chr(KeyAscii)
                txt卡号.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
        Exit Sub
    End If
    
    If KeyAscii <> 13 Then
        txt卡号.Text = txt卡号.Text & Chr(KeyAscii)
        txt卡号.SelStart = Len(txt卡号.Text)
    End If
    KeyAscii = 0
    strCardNo = Trim(txt卡号.Text)
    '68927,刘尔旋,2014-01-17,刷卡机刷卡末尾可能存在有回车符的情况
    EnableKBDHook
    If CheckBrush消费卡(strCardNo) = False Then
        txt卡号.Text = ""
        Exit Sub
    End If
    txt卡号.Text = strCardNo ' GetCardNODencode(strCardNo, mlngCardTypeID, mcurCardObject.CardPreporty.卡号密文规则, mbln消费卡)
    txt卡号.Tag = strCardNo
    '由于刷卡后,如果长度到达后,后一位还有回车符,需要取掉该回车符,不能转到密码框后接收该字符
    '所以加上了Doevnts:63335
    DoEvents
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Public Function CheckBrush消费卡(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷消费卡
    '编制:刘兴洪
    '日期:2011-06-23 17:48:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strPassWord As String, dbl失效面额 As Double
    Dim dbl实际支付 As Double, str限制类别 As String, lng消费卡ID As Long
    Dim dbl已刷总额 As Double, dbl当前支付 As Double
    
    If Not mbln消费卡 Then CheckBrush消费卡 = True: Exit Function
    dbl已刷总额 = GetBruhMoney  '获取已经刷卡的总额
    dbl当前支付 = mdbl本次消费总额 + mdbl已刷总额 - dbl已刷总额
    If CheckBrushSquareCard(mlngCardTypeID, strCardNo, mrsClassMoney, _
        dbl当前支付, strPassWord, mdbl帐户余额, dbl失效面额, dbl实际支付, _
        mbln余额不足禁止, str限制类别, mlng消费卡ID, dbl已刷总额) = False Then Exit Function
    '加入多次刷卡数据
    mstr限制类别 = str限制类别
    mdbl本次扣款额 = dbl实际支付:    txtPass.Tag = strPassWord
    Call ShowMoney
    CheckBrush消费卡 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetBruhMoney() As Double
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取已经刷卡的金额
    '返回:返回已经刷卡的金额
    '编制:刘兴洪
    '日期:2013-02-22 16:09:48
    '问题:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    On Error GoTo errHandle
    '显示已刷金额
    With vsBlance
        dblMoney = 0
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, .ColIndex("卡号")) <> "" Then
                dblMoney = dblMoney + Val(.TextMatrix(i, .ColIndex("刷卡金额")))
            End If
        Next
    End With
    GetBruhMoney = dblMoney
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitBalanceGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格信息
    '编制:刘兴洪
    '日期:2013-02-22 16:18:49
    '问题:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
   On Error GoTo errHandle
    With vsBlance
        .Clear
        .Rows = 2: .Cols = 7
        .TextMatrix(0, 0) = "卡号"
        .TextMatrix(0, 1) = "密码"
        .TextMatrix(0, 2) = "卡类别ID"
        .TextMatrix(0, 3) = "消费卡ID"
        .TextMatrix(0, 4) = "限制类别"
        .TextMatrix(0, 5) = "刷卡金额"
        .TextMatrix(0, 6) = "卡号密文显示"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            Select Case .ColKey(i)
            Case "卡类别ID", "消费卡ID", "密码", "限制类别", "卡号密文显示"
                .ColHidden(i) = True
            Case "刷卡金额"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
Private Function CheckTypeValied(ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
    ByVal str限制类别 As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查限制收费类别的合法性
    '入参:lngCardTypeID-卡类别ID
    '        strCardNo-卡号
    '       str限制类别-限制的收费类别
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-02-25 11:07:06
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strName As String, j As Long
    Dim strTemp As String, varData As Variant

    On Error GoTo errHandle
    If mcurCardObject Is Nothing Then
        strName = "消费卡"
    Else
        strName = mcurCardObject.CardPreporty.名称
    End If
    
    With vsBlance
        '先检查是否合法,合法才加入
        For i = 1 To .Rows - 1
            If Trim(.Cell(flexcpData, i, .ColIndex("卡号"))) <> "" Then
                If Trim(.Cell(flexcpData, i, .ColIndex("卡号"))) = strCardNo _
                    And Val(.TextMatrix(i, .ColIndex("卡类别ID"))) = lngCardTypeID Then
                    '此张卡已经刷卡,不能再进行刷卡验证
                    MsgBox "卡号为" & strCardNo & " 的" & strName & ",本次已经刷卡消费,不能再刷卡!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                strTemp = Trim(.TextMatrix(i, .ColIndex("限制类别")))
                
                If (strTemp <> "" Or Trim(str限制类别) <> "") And strTemp <> str限制类别 Then
                    '检查限制类别是否相同
                    If strTemp <> "" Then
                        varData = Split(strTemp, ",")
                        For j = 0 To UBound(varData)
                            If Trim(varData(j)) <> "" Then
                                '当前收费类别,不在限制类别中时,不检查
                                If InStr(1, "," & mstrCurFeeType & ",", "," & varData(j) & ",") > 0 Or mstrCurFeeType = "" Then
                                        If InStr(1, "," & str限制类别 & ",", "," & varData(j) & ",") = 0 Then
                                            MsgBox "卡号为" & strCardNo & " 的" & strName & "的限制类别与已经刷卡的限制类别不同,不能混合刷卡,详情如下:" & vbCrLf & "  已刷卡限制类别:" & strTemp & vbCrLf & "  当前刷卡限制类别:" & str限制类别, vbInformation + vbOKOnly, gstrSysName
                                            Exit Function
                                        End If
                                End If
                            End If
                        Next
                    End If
                    If str限制类别 <> "" Then
                        varData = Split(str限制类别, ",")
                        For j = 0 To UBound(varData)
                            If Trim(varData(j)) <> "" Then
                                '当前收费类别,不在限制类别中时,不检查
                                If InStr(1, "," & mstrCurFeeType & ",", "," & varData(j) & ",") > 0 Or mstrCurFeeType = "" Then
                                    If InStr(1, "," & strTemp & ",", "," & varData(j) & ",") = 0 Then
                                        MsgBox "卡号为" & strCardNo & " 的" & strName & "的限制类别与已经刷卡的限制类别不同,不能混合刷卡,详情如下:" & vbCrLf & "  已刷卡限制类别:" & strTemp & vbCrLf & "  当前刷卡限制类别:" & str限制类别, vbInformation + vbOKOnly, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Next
    End With
    CheckTypeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function AddBrushCardInfor(ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, ByVal strPassWord As String, _
    ByVal str限制类别 As String, ByVal lng消费卡ID As Long, _
    ByVal dblMoney As Double) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向表格中增加刷卡信息
    '入参:lngCardTypeID-当前卡类别ID
    '       strCardNo-当前卡号
    '       strPassWord-当前密码
    '       str限制类别-当前卡的限制类别
    '       lng消费卡ID-当前消费卡的卡ID
    '       dblMoney-本次刷卡金额
    '返回:加入成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-02-22 16:16:01
    '问题:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strName As String
    Dim strTemp As String, varData As Variant, lngRow As Long
    On Error GoTo errHandle
    If mcurCardObject Is Nothing Then
        strName = "消费卡"
    Else
        strName = mcurCardObject.CardPreporty.名称
    End If
    With vsBlance
        '先检查是否合法,合法才加入
        lngRow = 0
        For i = 1 To .Rows - 1
            If Trim(.Cell(flexcpData, i, .ColIndex("卡号"))) = "" Then
                    lngRow = i: Exit For
            End If
        Next
        If lngRow > .Rows - 1 Or lngRow = 0 Then
            .Rows = .Rows + 1
            lngRow = .Rows - 1
        End If
        .TextMatrix(lngRow, .ColIndex("卡号")) = IIf(mblnPassInputCardNo, String(Len(strCardNo), "*"), strCardNo)
        .Cell(flexcpData, lngRow, .ColIndex("卡号")) = strCardNo
        .TextMatrix(lngRow, .ColIndex("密码")) = strPassWord
        .TextMatrix(lngRow, .ColIndex("卡类别ID")) = lngCardTypeID
        .TextMatrix(lngRow, .ColIndex("消费卡ID")) = lng消费卡ID
        .TextMatrix(lngRow, .ColIndex("刷卡金额")) = Format(dblMoney, "0.00")
        .TextMatrix(lngRow, .ColIndex("卡号密文显示")) = IIf(mblnPassInputCardNo, 1, 0)
        .TextMatrix(lngRow, .ColIndex("限制类别")) = str限制类别
        .RowData(lngRow) = 0: .RowPosition(lngRow) = 1
        '77292,冉俊明,2014-8-29,消费卡退款验证时,如果不能退现，则退款额未全部验证则不允许通过
        If mbln消费卡 And mbln退费 And mbln退现 = False And mdbl本次扣款额 <> 0 Then
            cmdOK.Enabled = False
        Else
            cmdOK.Enabled = True
        End If
    End With
    '显示合计金额
    AddBrushCardInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ShowWndCaption()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示已刷金额
    '编制:刘兴洪
    '日期:2013-02-22 17:09:56
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    On Error GoTo errHandle
    Set tkpGroup = wndTaskPanel.Groups.Find(2)
    '存在，就退出
    If Not tkpGroup Is Nothing Then
        tkpGroup.Caption = "当前刷卡信息(已刷:" & Format(GetBruhMoney, "0.00") & ")"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub txt卡号_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt卡号.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt卡号.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt卡号_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt卡号.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPass.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPass_GotFocus()
    '89759:李南春,2017/4/14,刷卡后是否必须确认刷卡信息
    If Not mblnPosPass Then
        If txtPass.Tag = "" And txt卡号.Text <> "" Then Call txtPass_KeyPress(13): Exit Sub
    End If
    EnableKBDHook
    Call zlControl.TxtSelAll(txtPass)
    Call OpenPassKeyboard(txtPass)
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlSquareAffirm
    ElseIf KeyAscii = 22 Then
        KeyAscii = 0 '不允许粘贴
    End If
End Sub
Private Function zlSquareAffirm(Optional blnOK As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费确认
    '入参:blnOk-点完成功能时传入
    '返回:加载合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-02-26 17:34:02
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl剩余额 As Double
    If Not isValied Then Exit Function
    mstrCardNo = txt卡号.Tag: mstrPassWord = Trim(txtPass.Text)
    dbl剩余额 = mdbl本次消费总额 + mdbl已刷总额 - Val(txtMoney.Text) - GetBruhMoney
    Call AddBrushCardInfor(mlngCardTypeID, mstrCardNo, mstrPassWord, mstr限制类别, mlng消费卡ID, Val(txtMoney.Text))
    
    If Round(dbl剩余额, 6) = 0 Or mbln转预交 Then
        '刷卡完成,结束本次刷卡
        mbln退现 = False
        If chk退现.Visible Then mbln退现 = chk退现.value = 1
        mblnOk = True
        zlSquareAffirm = True
        If blnOK Then Exit Function
        Call SetReturnBrushCardInfor
        Unload Me
        Exit Function
    End If
    
    Call AddWndTaskPancelExpend
    mdbl本次扣款额 = mdbl本次消费总额 + mdbl已刷总额 - GetBruhMoney
    Call ShowMoney
    Call ShowWndCaption
    txtPass.Text = "": txt卡号.Text = ""
    If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    zlSquareAffirm = True
End Function
Private Function ReadCardNo() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取卡号
    '返回:成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-22 14:40:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strOutPatiXML As String, strCardNo As String
    On Error GoTo errHandle
    If mlngCardTypeID = 0 Then
        If mobjICCard Is Nothing Then Exit Function
         txt卡号.Text = mobjICCard.Read_Card()
         txt卡号.Tag = txt卡号.Text
         mstrCardNo = txt卡号.Text
         Exit Function
    End If
    
    ' 短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  调用的模块号
    'blnOlnyCardNO   boolean In  仅仅读取卡号
    'strExpand   String  In  扩展参数,暂留，现为空
    'strOutCardNO    String  Out 卡号
    'strOutclsPatientInfoXml  XML Out 见strOutclsPatientInfoXml参数说明
    '参数: blnOlnyCardNO=true时,返回空
    If gobjOneCardComLib.objThirdSwap.zlReadCard(Me, mlngModule, True, "", strCardNo, strOutPatiXML) = True Then
         txt卡号.Text = strCardNo ' GetCardNODencode(strCardNo, mlngCardTypeID, mcurCardObject.CardPreporty.卡号密文规则, mbln消费卡)
         txt卡号.Tag = strCardNo
    End If
    
    'txt卡号.PasswordChar = ""
    '91140:李南春,2015/11/30,消费卡刷卡
    If mbln消费卡 Then
        If CheckBrush消费卡(strCardNo) Then ReadCardNo = True
        Exit Function
    End If
    
    'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long,
       ' strCardTypeID as long ,ByVal strCardNo As String, strExpand As String, dblMoney As Double) As Boolean
    If mcurCardObject.CardObject.zlGetAccountMoney(Me, mlngModule, mlngCardTypeID, strCardNo, "", mdbl帐户余额) Then
         lblBalanceMoney.Caption = Format(mdbl帐户余额, "0.00")
          If mdbl帐户余额 - mdbl本次消费总额 < 0 Then
                If mdbl帐户余额 = 0 Then
                    Call MsgBox("该卡已经没有可用余额,不能继续!", vbInformation + vbOKOnly, gstrSysName)
                    If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
                    Exit Function
                Else
                    If mbln余额不足禁止 Then
                        Call MsgBox("该帐户余额不够支付本次消费额,不能继续!", vbInformation + vbOKOnly, gstrSysName)
                        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
                        Exit Function
                    End If
                    If MsgBox("该帐户余额不够支付本次消费额,是否用帐户余额作为本次消费额?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                          If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
                          Exit Function
                    End If
                    mdbl本次扣款额 = Round(mdbl帐户余额, 2)
                End If
          End If
    End If
    
    Call ShowMoney
    ReadCardNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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

Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
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

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    Dim lngPreIDKind As Long
    If Me.ActiveControl Is txt卡号 Then
        txt卡号.Text = strNo
        If txt卡号.Text = "" Then Call mobjICCard.SetEnabled(False)
        mstrCardNo = strNo
        Unload Me: Exit Sub
    End If
End Sub

Public Function CheckBrushSquareCard(ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal dbl本次支付额 As Double, _
    ByRef strPassWord As String, _
    ByRef dbl帐户金额 As Double, _
    ByRef dbl失效面额 As Double, ByRef dbl实际支付 As Double, _
    Optional bln余额不足禁止 As Boolean = True, _
    Optional ByRef str限制类别Out As String = "", _
    Optional ByRef lng消费卡ID As Long = 0, _
    Optional ByRef dbl已刷总额 As Double) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷消费卡
    '入参:lngCardTypeID-卡类别ID
    '       dbl已刷总额-已经刷卡的总额
    '出参: strPassWord-返回解密的密码
    '         dbl帐户金额-帐户金额
    '         dbl失效金额-失效金额
    '         str限制类别Out-限制的使用类别
    '        dbl实际支付-实际支付金额
    '返回:刷卡合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-02-25 10:57:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, blnFind As Boolean
    Dim dbl余额 As Double, str名称 As String
    Dim strSQL As String, dbl限定额 As Double, dbl合计 As Double, strMsg As String
    Dim intIndex As Integer, rs收费类别 As ADODB.Recordset, str限制类别 As String
    Dim varData As Variant, dblMoney As Double
    Dim str业务场合 As String, var费用来源 As Variant
    
    dbl实际支付 = dbl本次支付额
    
    '短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If lngCardTypeID < 0 Then Exit Function
    
    If str名称 = "" Then
        Set rsTemp = zlGet消费卡接口
        rsTemp.Filter = "编号=" & lngCardTypeID
        If rsTemp.EOF Then
            MsgBox "未找到相关的卡结算接口", vbInformation, gstrSysName
            Exit Function
        End If
        str名称 = NVL(rsTemp!名称)
        rsTemp.Filter = 0
    End If
    
    strSQL = _
        "Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss') as 有效期," & vbNewLine & _
        "       to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间," & vbNewLine & _
        "       decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态," & vbNewLine & _
        "       to_char(a.卡面金额," & gOraFmtString.FM_金额 & ") as 卡面金额," & vbNewLine & _
        "       to_char(a.销售金额," & gOraFmtString.FM_金额 & ") as 销售金额," & vbNewLine & _
        "       to_char(a.充值折扣率," & gOraFmtString.FM_折扣率 & ") as 充值折扣率," & vbNewLine & _
        "       to_char(a.余额," & gOraFmtString.FM_金额 & ") as 余额," & vbNewLine & _
        "       to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & vbNewLine & _
        "       a.限制类别 ,A.密码, b.应用场合, b.是否特定病人, a.病人ID" & vbNewLine & _
        "From 消费卡信息 A, 消费卡类别目录 B" & vbNewLine & _
        "Where a.接口编号 = b.编号 And A.卡号 = [1] and A.接口编号=[2]" & vbNewLine & _
        "      And 序号 = (Select Max(序号) From 消费卡信息 B Where 卡号 = A.卡号 and 接口编号=A.接口编号)" & vbNewLine & _
        "Order by a.序号"
    Err = 0: On Error GoTo Errhand:
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "消费卡检查", strCardNo, lngCardTypeID)
    If rsTemp.EOF Then
       ShowMsgbox "未找到相关的" & str名称 & "信息，请检查！"
        Exit Function
    End If
    
    '检查当前刷卡的合法性
    '是否回收
    If NVL(rsTemp!回收时间, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & str名称 & "已经被" & NVL(rsTemp!当前状态) & "，不能再刷卡！"
        Exit Function
    End If
    
    '是否停用
    If NVL(rsTemp!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & str名称 & "已经被停止使用，不能再刷卡！"
        Exit Function
    End If
    
    If NVL(rsTemp!密码) <> "" Then
        strPassWord = NVL(rsTemp!密码)
    End If
    
    lng消费卡ID = Val(NVL(rsTemp!id))
    str限制类别Out = NVL(rsTemp!限制类别)
    dbl余额 = Val(NVL(rsTemp!余额))
    
    dbl失效面额 = 0
    '检查效期
    If NVL(rsTemp!有效期, "3000-01-01 00:00:00") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
        '到了有效期
        If Val(NVL(rsTemp!可否充值)) = 1 Then
            '允许允值的,到期的,不能消费卡面金额,只能消费允值部分
            dbl失效面额 = zlGet失效面额(Val(NVL(rsTemp!id)))
            dbl余额 = dbl余额 - dbl失效面额
            If dbl余额 <= 0 Then dbl余额 = 0
        ElseIf mbln退费 = False Then
            '不允许允值的,不能再进行消费
            ShowMsgbox "卡号为" & strCardNo & "的" & str名称 & "已经失效，不能再刷卡！"
            Exit Function
        End If
    End If
    
    If mbln退费 Then  '退费
        If Not mVarData Is Nothing Then
            dblMoney = 0
            blnFind = False
            For i = 1 To mVarData.Count
                'arrayarray(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文,剩余未退金额(只针对退费))
                varData = mVarData(i)
                If Val(varData(0)) = lngCardTypeID And lng消费卡ID = varData(1) Then
                    blnFind = True
                    If UBound(varData) >= 7 Then
                        dblMoney = dblMoney + Val(varData(7))
                    Else
                        dblMoney = dblMoney + Val(varData(2))
                    End If
                End If
            Next
            If bln余额不足禁止 Then
                If Round(dblMoney, 6) < dbl本次支付额 Then
                    ShowMsgbox "卡号为 " & strCardNo & " 的" & str名称 & "退款金额超过了剩余未退金额，不能退费！"
                    Exit Function
                End If
            Else
                If Round(dblMoney, 6) < dbl本次支付额 Then dbl实际支付 = dblMoney
            End If
            '77292,冉俊明,2014-8-29,当前卡不在收费时使用的卡列表中，则该卡验证失败
            If Not blnFind Then
                ShowMsgbox "卡号为 " & strCardNo & " 的" & str名称 & "在收费时未使用，不能退费！"
                Exit Function
            End If
            '77292,冉俊明,2014-8-29,刷多次时,不能重复使用同一张卡
            If Not CheckTypeValied(lngCardTypeID, strCardNo, "") Then Exit Function
            '78494,冉俊明,2014-10-10,当前卡未退金额为零,则不允许再退
            If Round(dblMoney, 6) = 0 Then
                ShowMsgbox "卡号为 " & strCardNo & " 的" & str名称 & "未退金额为零，不能再退费！"
                Exit Function
            End If
        End If
        GoTo EndNO:
    End If
    

    If mbln转预交 And NVL(rsTemp!限制类别) <> "" Then
        ShowMsgbox str名称 & "存在限制的收费类别，不允许转预交！"
        Exit Function
    End If
    
    str业务场合 = NVL(rsTemp!应用场合) & "000"
    '应用场合检查
    '共三位数字组成，每一位1表示限制使用，0表示允许使用；第一位门诊业务，第二位住院业务，第三位体检业务；缺省为''111''
    var费用来源 = Split(mstr费用来源, ",")
    For i = 0 To UBound(var费用来源)
        If InStr("123", Val(var费用来源(i))) > 0 Then
            If Val(Mid(str业务场合, Val(var费用来源(i)), 1)) = 1 Then
                ShowMsgbox "卡号为" & strCardNo & "的" & str名称 & _
                    "针对" & Decode(Val(var费用来源(i)), 2, "住院", 3, "体检", "门诊") & "费用不允许使用！"
                Exit Function
            End If
        Else
            ShowMsgbox "卡号为" & strCardNo & "的" & str名称 & "在当前业务场合不允许使用！"
            Exit Function
        End If
    Next
    If Val(NVL(rsTemp!是否特定病人)) = 1 Then
        If mlng病人ID <> Val(NVL(rsTemp!病人ID)) Then
            ShowMsgbox "卡号为" & strCardNo & "的" & str名称 & "只能用于支付持卡人本人的费用！"
            Exit Function
        End If
    End If

    If dbl余额 <= 0 Then
        ShowMsgbox "卡号为" & strCardNo & "的" & str名称 & "已经没有余额，不能再刷卡消费！"
        Exit Function
    End If
    If dbl余额 < dbl本次支付额 Then
        If bln余额不足禁止 Then
            ShowMsgbox "" & str名称 & "的余额(" & Format(dbl余额, "0.00") & ")不够支付本次金额(" & Format(dbl本次支付额, "0.00") & ")！"
            Exit Function
        End If
        '用余额作为本次支付额
        dbl实际支付 = dbl余额
    End If
    '刷多次时,不能存在不同类别的刷卡情况
    If Not CheckTypeValied(lngCardTypeID, strCardNo, NVL(rsTemp!限制类别)) Then Exit Function

    '检查限制刷卡额
    Set rs收费类别 = zlGet收费类别
    str限制类别 = zlGet获取限制类别FromNameToCode(NVL(rsTemp!限制类别))
    
    If rsClassMoney Is Nothing Then GoTo EndNO:
    If rsClassMoney.State <> 1 Then GoTo EndNO:
    
    With rsClassMoney
        dbl合计 = 0: dbl限定额 = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl合计 = dbl合计 + Val(NVL(!金额))
            If InStr(1, str限制类别, "," & !收费类别 & ",") > 0 Then
                rs收费类别.Filter = "编码='" & NVL(!收费类别) & "'"
                If Not rsTemp.EOF Then
                    strMsg = strMsg & vbCrLf & "" & rs收费类别!名称 & ":" & Val(NVL(!金额))
                End If
                dbl限定额 = dbl限定额 + Val(NVL(!金额))
            End If
            .MoveNext
        Loop
        dbl限定额 = Format(dbl限定额, "0.00")
        dbl合计 = Format(dbl合计, "0.00")
        If dbl合计 - dbl限定额 - dbl已刷总额 >= dbl实际支付 Then GoTo EndNO:
        If dbl合计 - dbl限定额 - dbl已刷总额 <= 0 Then
            If dbl限定额 <> 0 Then
                ShowMsgbox "卡号为" & strCardNo & "的" & str名称 & " " & vbCrLf & "存在金额限制，本次不够支付,限制情况如下:" & vbCrLf & strMsg
            Else
                ShowMsgbox "已经刷卡消费完成,不能再刷卡消费!"
            End If
            Exit Function
        End If
        dbl实际支付 = dbl合计 - dbl限定额 - dbl已刷总额
    End With
EndNO:
    dbl帐户金额 = dbl余额
    CheckBrushSquareCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get收费限制类别_名称() As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费限制类别,以名称为主
    '返回:收费限制类别,以名称为主
    '编制:刘兴洪
    '日期:2013-03-07 10:30:39
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If mbln转预交 Then Exit Function
    If mrsClassMoney Is Nothing Then Exit Function
    If mrsClassMoney.State <> 1 Then Exit Function
    If mrsClassMoney.RecordCount = 0 Then Exit Function
    Set rsTemp = zlGet收费类别
    rsTemp.Filter = 0
    mrsClassMoney.Filter = 0
    With mrsClassMoney
        .MoveFirst
        Do While Not .EOF
            rsTemp.Find "编码='" & NVL(!收费类别, "-") & "'", , adSearchForward, 1
            If Not rsTemp.EOF Then strTemp = strTemp & "," & rsTemp!名称
            .MoveNext
        Loop
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    mrsClassMoney.Filter = 0
    Get收费限制类别_名称 = strTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 

End Function


