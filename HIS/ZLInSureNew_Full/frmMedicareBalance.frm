VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMedicareBalance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "医保病人收费结算校对"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   435
      Left            =   5520
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txt缴款 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3840
      Width           =   1755
   End
   Begin VB.TextBox txt找补 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   3840
      Width           =   1755
   End
   Begin VB.TextBox txtMargin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      ToolTipText     =   "仅当改变缺省结算方式的金额时才产生"
      Top             =   120
      Width           =   1755
   End
   Begin VB.TextBox txtTmp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.TextBox txt预交冲款 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   5190
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   600
      Width           =   1755
   End
   Begin VB.TextBox txtPay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   600
      Width           =   1755
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   120
      Width           =   1755
   End
   Begin VB.Frame Frame2 
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
      TabIndex        =   9
      Top             =   4320
      Width           =   7365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确 定(&O)"
      Height          =   435
      Left            =   3840
      TabIndex        =   8
      Top             =   4560
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBalance 
      Height          =   2505
      Left            =   390
      TabIndex        =   4
      Top             =   1200
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4419
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^  结算方式  |^   结算金额   |^      结算号码      "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label lbl缴款 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现金缴款"
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
      Left            =   360
      TabIndex        =   15
      Top             =   3960
      Width           =   960
   End
   Begin VB.Label lbl找补 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现金找补"
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
      Left            =   4080
      TabIndex        =   14
      Top             =   3930
      Width           =   960
   End
   Begin VB.Label lblMargin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费差额"
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
      Left            =   4080
      TabIndex        =   13
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lbl预交冲款 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "预交冲款"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4080
      TabIndex        =   12
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblPay 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应缴金额"
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
      Left            =   360
      TabIndex        =   11
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "实收金额"
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
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmMedicareBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
'    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;3-公共部件增加GetNextNO();4-V10.24及以上版本
'    99-所有交易增加附加参数(最新版)

Private mbytInFun As Byte '0-费用模块调用,1-医保模块调用

Private mlng结帐ID As Long
Private mcur实收金额 As Currency
Private mcur冲预交额 As Currency
Private mstr保险结算 As String
Private mstr收费结算 As String      '其合计为mcur实收金额-mcur预交余额-保险结算合计+mcur收费误差
Private mcur收费误差 As String
Private mcur预交余额 As Currency
Private mintInsure As Integer       '用来判断是否支持分币处理
Private mcur缴款 As Currency

Private mblnOK  As Boolean
Private mintDefault As Integer '缺省结算方式行(为0表示没有)
Private mcurMediCare   As Currency  '医保结算合计,根据[mstr保险结算]计算
Private mblnClickOK As Boolean      '窗体只允许点确定退出
Private mblnCent As Boolean         '医保是否支持分币处理

'1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项
Private Enum PayType
    现金 = 1
    非医保非现金 = 2
    医保个人帐户 = 3
    医保其它结算 = 4
    代收款 = 5
End Enum

'模块参数的私有化
Private Const support分币处理 = 25  '医保病人是否处理分币   ,主要是为了便于医保与医院对帐
Private mstr结算方式 As String
Private mstrDec As String
Private mBytMoney As Byte '收费分币处理方法


Public Function ShowMeFromOut(ByRef frmParent As Object, ByVal lng结帐ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lng病人ID As Long, strTmp As String
    
    On Error GoTo ErrH
    If Not IsZLHIS10 Then
        ShowMeFromOut = frmMedicareBalance9.ShowMeFromOut(frmParent, lng结帐ID)
        Exit Function
    End If
    
    mlng结帐ID = lng结帐ID
    strSQL = "Select Sum(Decode(Nvl(附加标志, 0), 9, 0, 实收金额)) As 实收金额," & _
             "       Sum(Decode(Nvl(附加标志, 0), 9, 实收金额, 0)) As 误差金额" & _
             " From 门诊费用记录" & _
             " Where 记录状态 = 1 And 结帐id = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "保险结算管理", lng结帐ID)
    mcur实收金额 = Val("" & rsTmp!实收金额)
    mcur收费误差 = Val("" & rsTmp!误差金额)
        
    strSQL = "Select a.病人ID,a.记录性质,a.结算方式,a.结算号码,b.性质 结算性质,a.冲预交 " & _
             "   From 病人预交记录 a,结算方式 b " & _
             "   Where a.记录状态 = 1 And a.结算方式 = B.名称 And 结帐id =[1] "
    Set rsTmp = OpenSQLRecord(strSQL, "保险结算管理", lng结帐ID)
    
    If rsTmp.RecordCount > 0 Then lng病人ID = rsTmp!病人ID
    
    mcur冲预交额 = 0
    rsTmp.Filter = "记录性质=1 or 记录性质=11"
    For i = 1 To rsTmp.RecordCount
        mcur冲预交额 = mcur冲预交额 + rsTmp!冲预交
        rsTmp.MoveNext
    Next
        
    mstr收费结算 = "" '结算方式|结算金额|结算号码||
    rsTmp.Filter = "记录性质=3 And 结算性质<>3 And 结算性质<>4"
    For i = 1 To rsTmp.RecordCount
        mstr收费结算 = mstr收费结算 & "||" & rsTmp!结算方式 & "|" & rsTmp!冲预交 & "|" & Nvl(rsTmp!结算号码)
        rsTmp.MoveNext
    Next
    If mstr收费结算 <> "" Then mstr收费结算 = Mid(mstr收费结算, 3)
    
        
    rsTmp.Filter = 0
    strSQL = "Select 结算方式,金额 From 医保核对表 Where 结帐id =[1] And 结算方式<>'现金'"  '医保管控的过程固定写入了一条"现金"
    Set rsTmp = OpenSQLRecord(strSQL, "保险结算管理", lng结帐ID)
    mstr保险结算 = ""   '结算方式|结算金额||
    For i = 1 To rsTmp.RecordCount
        mstr保险结算 = mstr保险结算 & "||" & rsTmp!结算方式 & "|" & rsTmp!金额
        rsTmp.MoveNext
    Next
    If mstr保险结算 <> "" Then mstr保险结算 = Mid(mstr保险结算, 3)
    
    
    
    mcur预交余额 = 0
    mintInsure = 0
    If lng病人ID <> 0 Then
        Set rsTmp = GetMoneyInfo(lng病人ID)
        If Not rsTmp.EOF Then mcur预交余额 = Val("" & rsTmp!预交余额) - Val("" & rsTmp!费用余额)
        
        strSQL = "Select nvl(险类,0) as 险类 From 病人信息 Where 病人ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "保险结算管理", lng病人ID)
        If Not rsTmp.EOF Then mintInsure = rsTmp!险类
    End If
    
    '本地或系统参数
    #If gverControl >= 4 Then
        mstr结算方式 = zlDatabase.GetPara("缺省结算方式", 100, 1121)
        mstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
        strTmp = zlDatabase.GetPara(14, glngSys, , 0)
    #Else
        mstr结算方式 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\zl9OutExse", "缺省结算方式", "")
        mstrDec = "0." & String(Val(GetPara(9, glngSys, , , 2)), "0")
        strTmp = GetPara(14, glngSys, , , 0)
    #End If
    
    mBytMoney = Val(IIf(Len(strTmp) = 1, strTmp, Mid(strTmp, 2, 1)))
    
    mbytInFun = 1
    Me.Show 1, frmParent
    ShowMeFromOut = mblnOK
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ShowME(ByRef frmParent As Object, ByVal lng结帐ID As Long, ByVal cur实收金额 As Currency, _
        ByVal cur冲预交额 As Currency, ByVal str保险结算 As String, ByRef str收费结算 As String, _
        ByVal cur收费误差 As Currency, ByVal cur预交余额 As Currency, ByVal intinsure As Integer, _
        ByVal str缺省结算方式 As String, ByVal str缺省金额位数 As String, ByVal byt缺省分币方式 As Byte, ByRef cur缴款 As Currency) As Boolean
        
    mlng结帐ID = lng结帐ID
    mintInsure = intinsure
    mcur实收金额 = cur实收金额
    mcur冲预交额 = cur冲预交额
    mstr保险结算 = str保险结算
    mstr收费结算 = str收费结算
    mcur收费误差 = cur收费误差
    mcur预交余额 = cur预交余额
    
    mstr结算方式 = str缺省结算方式
    mstrDec = str缺省金额位数
    mBytMoney = byt缺省分币方式
    mcur缴款 = cur缴款
    
    mbytInFun = 0
    Me.Show 1, frmParent
    
    str收费结算 = mstr收费结算  '返回用于缴款累计
    cur缴款 = mcur缴款
    
    ShowME = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    mblnClickOK = True: Unload Me
End Sub

Private Sub cmdOK_Click()
    '检查数据
    Dim i As Long
    
    If Val(txtMargin.Text) <> 0 Then
        If Val(txtMargin.Text) > 0 Then
            MsgBox "病人支付金额不足,请按所显示的差额补款。", vbExclamation, gstrSysName
            mshBalance.SetFocus: Exit Sub
        Else
            MsgBox "病人支付金额过多,请按所显示的差额退款。", vbExclamation, gstrSysName
            mshBalance.SetFocus: Exit Sub
        End If
    End If
    
    '更新数据
    mstr收费结算 = ""
    For i = 1 To mshBalance.Rows - 1
        If Val(mshBalance.TextMatrix(i, 1)) <> 0 Then
            If mshBalance.RowData(i) <> PayType.医保个人帐户 And mshBalance.RowData(i) <> PayType.医保其它结算 Then
                mstr收费结算 = mstr收费结算 & "||" & mshBalance.TextMatrix(i, 0) & "|" & Val(mshBalance.TextMatrix(i, 1)) & _
                    "|" & IIf(mshBalance.TextMatrix(i, 2) = "", " ", mshBalance.TextMatrix(i, 2))
            End If
        End If
    Next
    mstr收费结算 = Mid(mstr收费结算, 3)
    
    gstrSQL = "zl_门诊收费结算_Update(" & mlng结帐ID & ",'" & mstr收费结算 & "'," & mcur冲预交额 & ",'" & mstr保险结算 & "'," & mcur收费误差 & _
        IIf(Val(txt缴款.Text) <> 0, "," & Val(txt缴款.Text) & "," & Val(txt找补.Text), "") & ")"
    On Error GoTo ErrH
    Call ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnOK = True
    mblnClickOK = True: Unload Me
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnClickOK = True: Unload Me
End Sub

Private Sub Form_Activate()
    txt缴款.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim rs应用场合 As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim arrPay As Variant, arrMediCare As Variant, blnExist As Boolean
    Dim curPay As Currency          '付款结算金额合计
    Dim curBalance As Currency      '实收合计减医保结算合计之后的余额
    Dim str可用的医保结算方式 As String
    
    '变量初始
    mblnClickOK = False
    mblnOK = False
    mintDefault = 0
    mcurMediCare = 0
    
    '确定和取消按钮
    If mbytInFun = 0 Then
        cmdOK.Left = cmdCancel.Left
        cmdCancel.Visible = False
    Else
        cmdCancel.Visible = True
    End If
    
    
    mblnCent = gclsInsure.GetCapability(support分币处理, , mintInsure)
    
    arrPay = Array()
    If mstr收费结算 <> "" Then                  '结算方式|结算金额|结算号码||
        arrPay = Split(mstr收费结算, "||")
    End If
    arrMediCare = Array()                       '结算方式|结算金额||
    If mstr保险结算 <> "" Then
        arrMediCare = Split(mstr保险结算, "||")
    End If
    
    On Error GoTo ErrH
    strSQL = _
        " Select Distinct B.编码,B.名称,B.性质,A.缺省标志" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where ((A.应用场合=[1] And B.性质<>3 And B.性质<>4) OR (B.性质=3 OR B.性质=4)) " & _
        " And B.名称=A.结算方式(+) And B.性质<>5" & _
        " Order by B.性质,B.编码"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, "收费")
    
    strSQL = "Select 应用场合,结算方式 From 结算方式应用 Where 应用场合=[1]"
    Set rs应用场合 = OpenSQLRecord(strSQL, Me.Caption, "收费")
    
    mshBalance.ColAlignment(0) = 1
    mshBalance.ColAlignment(1) = 7
    mshBalance.ColAlignment(2) = 1
    mshBalance.Rows = rsTmp.RecordCount + 1
    i = 1
    Do While Not rsTmp.EOF
        mshBalance.RowData(i) = zlCommFun.Nvl(rsTmp!性质, PayType.现金)               '用来判断是否可以修改金额,以及是否是现金
        mshBalance.TextMatrix(i, 0) = rsTmp!名称
        
        '医保结算方式不允许修改,设置不同颜色
        If mshBalance.RowData(i) = PayType.医保个人帐户 Or mshBalance.RowData(i) = PayType.医保其它结算 Then
            
            '保险结算
            blnExist = False
            For j = 0 To UBound(arrMediCare)
                If Split(arrMediCare(j), "|")(0) = rsTmp!名称 Then
                    blnExist = True
                    rs应用场合.Filter = "结算方式='" & rsTmp!名称 & "'"
                    If rs应用场合.EOF Then
                        MsgBox "注意:结算方式[" & rsTmp!名称 & "]未设置应用于[收费]场合,请到[结算方式管理]中设置!", vbInformation, gstrSysName
                    End If
                    
                    mshBalance.TextMatrix(i, 1) = Split(arrMediCare(j), "|")(1)
                    mshBalance.TextMatrix(i, 2) = ""    '无结算号码
                    mcurMediCare = mcurMediCare + Val(mshBalance.TextMatrix(i, 1))
                    Exit For
                End If
            Next
            If blnExist Then
                For j = 0 To mshBalance.Cols - 1
                    mshBalance.Row = i: mshBalance.COL = j
                    mshBalance.CellBackColor = &HE7CFBA
                Next
                i = i + 1                                   '医保结算不允许改,没有金额的不显示
            End If
            
            str可用的医保结算方式 = str可用的医保结算方式 & "," & rsTmp!名称
            
        Else
            If rsTmp!名称 = mstr结算方式 Then mintDefault = i
            If zlCommFun.Nvl(rsTmp!缺省标志, 0) = 1 And mintDefault = 0 Then mintDefault = i
            If zlCommFun.Nvl(rsTmp!性质, 1) = 1 And mintDefault = 0 Then mintDefault = i
        
            '收费结算
            For j = 0 To UBound(arrPay)
                If Split(arrPay(j), "|")(0) = rsTmp!名称 Then
                    mshBalance.TextMatrix(i, 1) = Split(arrPay(j), "|")(1)
                    mshBalance.TextMatrix(i, 2) = Trim(Split(arrPay(j), "|")(2))
                    Exit For
                End If
            Next
            i = i + 1                                      '因为允许改,没有金额的也要显示
        End If
        rsTmp.MoveNext
    Loop
    
    mshBalance.Rows = i     '最后一次加1正好使行数包含列标题行,最后一行是医保且金额为零,i没有加1正好删除
    
    
    '先检查每一种医保结算方式是否都存在
    If mstr保险结算 <> "" Then
        str可用的医保结算方式 = str可用的医保结算方式 & ","
        For j = 0 To UBound(arrMediCare)
            If InStr(str可用的医保结算方式, "," & Split(arrMediCare(j), "|")(0) & ",") <= 0 Then
                MsgBox "医保结算方式[" & Split(arrMediCare(j), "|")(0) & "]未设置,请先到[结算方式管理]中设置!", vbInformation, gstrSysName
                cmdCancel.Visible = True
                cmdOK.Visible = False
            End If
        Next
    End If
    
    
    If mintDefault > 0 Then
        mshBalance.Row = mintDefault: mshBalance.COL = 0
        mshBalance.CellFontBold = True
        mshBalance.COL = 1
    Else        '结算方式没有缺省值,并且无现金方式的情况
        mshBalance.Row = 1: mshBalance.COL = 1
    End If
        
    txt预交冲款.Text = Format(mcur冲预交额, "0.00")
    txt预交冲款.Enabled = mcur预交余额 > 0      '应付合计大于零才允许使用
    If txt预交冲款.Enabled Then txt预交冲款.Enabled = (mcur实收金额 - mcurMediCare > 0)
    
    txtTotal.Text = Format(mcur实收金额, mstrDec)
    txtTotal.ToolTipText = "预结算时,误差金额:" & Format(mcur收费误差, mstrDec)
    
    Call ShowMoney(True)
            
    curPay = 0
    For i = 1 To mshBalance.Rows - 1
        If mshBalance.RowData(i) <> PayType.医保个人帐户 And mshBalance.RowData(i) <> PayType.医保其它结算 Then
            curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
        End If
    Next
    txtPay.Text = Format(curPay, "0.00")
    txt缴款.Text = Format(mcur缴款, "0.00")
            
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtTmp_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim curPay As Currency
    
    If KeyAscii <> 13 Then
        If mshBalance.COL = 1 Then
            If KeyAscii = vbKeyEscape Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
            If InStr(txtTmp.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        Else '结算号码特殊字符限制
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = asc(UCase(Chr(KeyAscii)))
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0

        If mshBalance.COL = 2 Then
            '结算号码防拷贝特殊字符
            If InStr(txtTmp.Text, "'") > 0 Or InStr(txtTmp.Text, "|") > 0 Or InStr(txtTmp.Text, ",") > 0 Then
                Call Beep: Exit Sub
            End If
            mshBalance.TextMatrix(mshBalance.Row, mshBalance.COL) = txtTmp.Text
        ElseIf mshBalance.COL = 1 Then
            If Not IsNumeric(txtTmp.Text) And Trim(txtTmp.Text) <> "" Then
                MsgBox "请输入正确的数值。", vbInformation, gstrSysName
                zlControl.TxtSelAll txtTmp: Exit Sub
            End If
            
            If Val(txtTmp.Text) <> 0 Then   '空字符的val为零
                txtTmp.Text = Format(Val(txtTmp.Text), "0.00")
                If Val(mshBalance.RowData(mshBalance.Row)) = PayType.现金 And mblnCent Then  '如果是在现金栏内输入,则进行分币处理
                    txtTmp.Text = Format(CentMoney(Val(txtTmp.Text)), "0.00")
                End If
                If Val(mshBalance.TextMatrix(mshBalance.Row, mshBalance.COL)) = Val(txtTmp.Text) Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
                If txtTmp.Text = "0.00" Then
                    mshBalance.TextMatrix(mshBalance.Row, mshBalance.COL) = ""
                Else
                    mshBalance.TextMatrix(mshBalance.Row, mshBalance.COL) = Format(Val(txtTmp.Text), "0.00")
                End If
            Else
                If mshBalance.TextMatrix(mshBalance.Row, mshBalance.COL) = "" Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
                mshBalance.TextMatrix(mshBalance.Row, mshBalance.COL) = ""
            End If
            
            Call ShowMoney(mintDefault <> mshBalance.Row)
        End If
        mshBalance.SetFocus
        txtTmp.Visible = False
        
        If mshBalance.Row = mshBalance.Rows - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '下一行处理
            mshBalance.Row = mshBalance.Row + 1
            mshBalance.COL = 1
            If mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(0) - 2) > 1 Then
                mshBalance.TopRow = mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(1) - 2)
            End If
        End If
    End If
End Sub
Private Sub ShowMoney(Optional ByVal blnAutoSet As Boolean)
    Dim curPay As Currency, curBalance As Currency
    Dim blnCent As Boolean, i As Long, bln存在补款 As Boolean
    
    If blnAutoSet And mintDefault > 0 Then      '根据差额自动补平并计算
        For i = 1 To mshBalance.Rows - 1
            If mshBalance.RowData(i) <> PayType.医保个人帐户 And mshBalance.RowData(i) <> PayType.医保其它结算 Then
                curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
            End If
        Next
        curBalance = mcur实收金额 + mcur收费误差 - (curPay + mcurMediCare + mcur冲预交额)
    
        '剩余部份尝试设置到缺省结算方式上
        curBalance = Val(mshBalance.TextMatrix(mintDefault, 1)) + curBalance
        If mshBalance.RowData(mintDefault) = PayType.现金 And mblnCent Then   '现金时要进行分币处理
            mshBalance.TextMatrix(mintDefault, 1) = Format(CentMoney(curBalance), "0.00")
        Else
            mshBalance.TextMatrix(mintDefault, 1) = Format(curBalance, "0.00")
        End If
        If Val(mshBalance.TextMatrix(mintDefault, 1)) = 0 Then mshBalance.TextMatrix(mintDefault, 1) = ""
        
        curPay = 0
        For i = 1 To mshBalance.Rows - 1
            curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
        Next
        
        mcur收费误差 = curPay - mcur实收金额
        txtPay.ToolTipText = "正式结算后,误差金额:" & Format(mcur收费误差, "0.00")
        
    Else
        bln存在补款 = True
    End If
    
    curPay = 0
    For i = 1 To mshBalance.Rows - 1
        If mshBalance.RowData(i) <> PayType.医保个人帐户 And mshBalance.RowData(i) <> PayType.医保其它结算 Then
            curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
        End If
    Next
    
    If bln存在补款 Then
       txtMargin.Text = Format(mcur实收金额 + mcur收费误差 - (curPay + mcurMediCare + mcur冲预交额), "0.00")
    Else
        txtMargin.Text = "0.00"
    End If
    
    If Val(txt缴款.Text) > 0 Then Call txt缴款_Change
End Sub
Private Sub txtTmp_LostFocus()
    txtTmp.Visible = False
End Sub

Private Sub txtTmp_Validate(Cancel As Boolean)
    txtTmp.Visible = False
End Sub


Private Sub txt预交冲款_GotFocus()
    zlControl.TxtSelAll txt预交冲款
    txt预交冲款.Tag = txt预交冲款.Text  '记录原值
End Sub

Private Sub txt预交冲款_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr(txt预交冲款.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
    If InStr("0123456789." & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt预交冲款_Validate(Cancel As Boolean)
    If Trim(txt预交冲款.Text) = "" Then
        txt预交冲款.Text = "0.00"
    ElseIf Not IsNumeric(txt预交冲款.Text) Then
        MsgBox "无效数值！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
        
    ElseIf Val(txt预交冲款.Text) < 0 Then
        MsgBox "预交款冲款金额不能为负！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
        
    ElseIf Val(txt预交冲款.Text) > 0 And (mcur实收金额 - mcurMediCare) < 0 Then
        MsgBox "单据应付金额为负时不能使用预交款！", vbInformation, gstrSysName
        txt预交冲款.Text = "0.00"
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
        
    ElseIf Val(txt预交冲款.Text) > mcur预交余额 Then
        MsgBox "预交款冲款金额不能超过病人的预交余额:" & CStr(mcur预交余额) & " ！", vbInformation, gstrSysName
        txt预交冲款.Text = CStr(mcur预交余额)
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
        
    ElseIf Val(txt预交冲款.Text) > (mcur实收金额 - mcurMediCare) And Val(txt预交冲款.Text) <> 0 Then
        MsgBox "预交款冲款金额不能大于应付金额:" & CStr((mcur实收金额 - mcurMediCare)) & " ！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt预交冲款: Cancel = True: Exit Sub
    Else
        txt预交冲款.Text = Format(txt预交冲款.Text, "0.00")
    End If

    If Val(txt预交冲款.Text) <> Val(txt预交冲款.Tag) Then
        
        mcur冲预交额 = Val(txt预交冲款.Text)
        Call ShowMoney(True)
        
        Dim curPay As Currency, i As Long
        curPay = 0
        For i = 1 To mshBalance.Rows - 1
            If mshBalance.RowData(i) <> PayType.医保个人帐户 And mshBalance.RowData(i) <> PayType.医保其它结算 Then
                curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
            End If
        Next
        txtPay.Text = Format(curPay, "0.00")
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnClickOK Then Cancel = 1
End Sub

Private Sub mshBalance_DblClick()
    If Not txtTmp.Visible And mshBalance.Row >= 1 And mshBalance.COL >= 1 And _
        mshBalance.RowData(mshBalance.Row) <> PayType.医保个人帐户 And mshBalance.RowData(mshBalance.Row) <> PayType.医保其它结算 Then
        With txtTmp
            .MaxLength = IIf(mshBalance.COL = 2, 30, 10)
            .Left = mshBalance.Left + mshBalance.CellLeft + 15
            .Top = mshBalance.Top + mshBalance.CellTop + (mshBalance.CellHeight - txtTmp.Height) / 2 - 15
            .Width = mshBalance.CellWidth - 60
            .ForeColor = mshBalance.CellForeColor
            .BackColor = mshBalance.CellBackColor
            .Alignment = IIf(mshBalance.COL = 1, 1, 0)
            .Text = mshBalance.TextMatrix(mshBalance.Row, mshBalance.COL)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If mshBalance.COL = 0 Then
            mshBalance.COL = 1
        ElseIf mshBalance.Row < mshBalance.Rows - 1 Then
            mshBalance.Row = mshBalance.Row + 1
            mshBalance.COL = 1
            If mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(0) - 2) > 1 Then
                mshBalance.TopRow = mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(1) - 2)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub mshBalance_KeyPress(KeyAscii As Integer)
    If Not txtTmp.Visible And mshBalance.Row >= 1 And mshBalance.COL > 0 And KeyAscii <> 13 And KeyAscii <> vbKeyEscape And _
            mshBalance.RowData(mshBalance.Row) <> PayType.医保个人帐户 And mshBalance.RowData(mshBalance.Row) <> PayType.医保其它结算 Then
        
        If mshBalance.COL = 1 Then
            If InStr("0123456789.-", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        Else '结算号码特殊字符限制
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        
        With txtTmp
            .MaxLength = IIf(mshBalance.COL = 2, 30, 10)
            .Left = mshBalance.Left + mshBalance.CellLeft + 15
            .Top = mshBalance.Top + mshBalance.CellTop + (mshBalance.CellHeight - .Height) / 2 - 15
            .Width = mshBalance.CellWidth - 60
            .ForeColor = mshBalance.CellForeColor
            .BackColor = mshBalance.CellBackColor
            .Alignment = IIf(mshBalance.COL = 1, 1, 0)
            .Text = UCase(Chr(KeyAscii))
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub


Private Sub txt缴款_Change()
    Dim cur现金 As Currency, i As Long
    For i = 1 To mshBalance.Rows - 1
        If mshBalance.RowData(i) = PayType.现金 Then
            cur现金 = Val(mshBalance.TextMatrix(i, 1))
            Exit For
        End If
    Next
    mcur缴款 = Val(txt缴款.Text)
    If mcur缴款 = 0 Then txt找补.Text = "0.00": Exit Sub
    txt找补.Text = Format(mcur缴款 - cur现金, "0.00")
End Sub

Private Sub txt缴款_GotFocus()
    Call zlControl.TxtSelAll(txt缴款)
End Sub

Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
        If txt缴款.Text <> "0.00" Then
            If Val(txt找补.Text) >= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
                txt缴款.SetFocus
                zlControl.TxtSelAll txt缴款
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab) '病人累加缴款
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf KeyAscii = asc(".") And InStr(txt缴款.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt缴款_LostFocus()
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub

Private Sub txt缴款_Validate(Cancel As Boolean)
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub


Private Function GetMoneyInfo(lng病人ID As Long, Optional curModiMoney As Currency) As ADODB.Recordset
'功能：获取指定病人的剩余额
    Dim strSQL As String
        
    #If gverControl >= 5 Then
        If curModiMoney = 0 Then
            strSQL = "Select Nvl(费用余额,0) as 费用余额,Nvl(预交余额,0) as 预交余额 From 病人余额 Where 性质=1 And 类型=1 And 病人ID=[1]"
        Else
            strSQL = "Select Nvl(费用余额,0)-[2] as 费用余额,Nvl(预交余额,0) as 预交余额 From 病人余额 Where 性质=1 And 类型=1 And 病人ID=[1]"
        End If
    #Else
        If curModiMoney = 0 Then
            strSQL = "Select Nvl(费用余额,0) as 费用余额,Nvl(预交余额,0) as 预交余额 From 病人余额 Where 性质=1 And 病人ID=[1]"
        Else
            strSQL = "Select Nvl(费用余额,0)-[2] as 费用余额,Nvl(预交余额,0) as 预交余额 From 病人余额 Where 性质=1 And 病人ID=[1]"
        End If
    #End If
    On Error GoTo ErrH
    Set GetMoneyInfo = OpenSQLRecord(strSQL, "mdlOutExse", lng病人ID, curModiMoney)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'功能：对指定金额按分币处理规则进行处理,返回处理后的金额
'参数：curMoney=要进行分币处理的金额(为应缴金额,2位小数)
'      mBytMoney=
'         0.不处理
'         1.采取四舍五入法,eg:0.51=0.50;0.56=0.60
'         2.补整收法,eg:0.51=0.60,0.56=0.60
'         3.舍分收法,eg:0.51=0.50,0.56=0.50
'         4.四舍六入五成双,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           四舍六入五成双,详见我国科学技术委员会正式颁布的《数字修约规则》,但根据vb的Round函数,若被舍弃的数字包括几位数字时，不对该数字进行连续修约
'           即银行家舍入法:四舍六入五考虑，五后非零就进一，五后皆零看奇偶，五前为偶应舍去，五前为奇要进一
'         5.三七作五、二舍八入,对角进行处理，不需要先对分币进行舍入,即0.29(含)以下都舍掉角，0.80(含)以上都进角，0.3-0.79处理为0.5。
'         6-五舍六入:eg:0.15=0.10:0.16=0.2:    问题:34519

    Dim intSign As Integer, curTmp As Currency

    If mBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf mBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '先取两位金额,再处理分币,如:0.248 得0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf mBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf mBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf mBytMoney = 4 Then
        CentMoney = Format(Round(curMoney, 1), "0.00")
    ElseIf mBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = curMoney - Int(curMoney)
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf mBytMoney = 6 Then
         '刘兴洪 问题:34519 五舍六入:eg:0.15=0.10:0.16=0.2:    日期:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function


Private Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMAX As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMAX Then intMAX = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMAX
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMAX = LenB(StrConv(varValue, vbFromUnicode))
            If intMAX = 0 Or intMAX < 200 Then intMAX = 200
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMAX, varValue)
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMAX = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMAX = 0 Or intMAX < 200 Then intMAX = 200
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMAX, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function


Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"

    Static cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMAX As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '清除原有参数:不然不能重复执行
        cmdData.CommandText = "" '不为空有时清除参数出错
        Do While cmdData.Parameters.Count > 0
            cmdData.Parameters.Delete 0
        Loop
        
        '执行的过程名
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, Val(strPar))
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'") '这种情况绑定变量只需要一个"'"
                        intMAX = LenB(StrConv(strPar, vbFromUnicode))
                        If intMAX = 0 Or intMAX < 200 Then intMAX = 200
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMAX, strPar)
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = zlDatabase.Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULL值当成数字处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                    ElseIf strPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else '可能是其他复杂的表达式，无法处理
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '执行过程
        If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '这句比较慢
            cmdData.CommandType = adCmdStoredProc
        End If
        cmdData.CommandText = strProc
        
        Call SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
        Call SQLTest
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    Call SQLTest(App.ProductName, strFormCaption, strSQL)
    gcnOracle.Execute strSQL, , adCmdStoredProc
    Call SQLTest
End Sub
