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
    Dim strSql As String, i As Long, lng病人ID As Long, strTmp As String
        
    On Error GoTo errH
    mlng结帐ID = lng结帐ID
    
    strSql = "" & _
    " Select Sum(Decode(Nvl(附加标志, 0), 9, 0, 实收金额)) As 实收金额," & _
    "       Sum(Decode(Nvl(附加标志, 0), 9, 实收金额, 0)) As 误差金额" & _
    " From 门诊费用记录" & _
    " Where 记录状态 = 1 And 结帐id = [1] "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng结帐ID)
    mcur实收金额 = Val("" & rsTmp!实收金额)
    mcur收费误差 = Val("" & rsTmp!误差金额)
        
    strSql = "Select a.病人ID,a.记录性质,a.结算方式,a.结算号码,b.性质 结算性质,a.冲预交 " & _
             "   From 病人预交记录 a,结算方式 b " & _
             "   Where a.记录状态 = 1 And a.结算方式 = B.名称 And 结帐id =[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng结帐ID)
    
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
        mstr收费结算 = mstr收费结算 & "||" & rsTmp!结算方式 & "|" & rsTmp!冲预交 & "|" & zlCommFun.Nvl(rsTmp!结算号码)
        rsTmp.MoveNext
    Next
    If mstr收费结算 <> "" Then mstr收费结算 = Mid(mstr收费结算, 3)
    
        
    rsTmp.Filter = 0
    strSql = "Select 结算方式,金额 From 医保核对表 Where 结帐id =[1] And 结算方式<>'现金'"  '医保管控的过程固定写入了一条"现金"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng结帐ID)
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
        
        strSql = "Select nvl(险类,0) as 险类 From 病人信息 Where 病人ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng病人ID)
        If Not rsTmp.EOF Then mintInsure = rsTmp!险类
    End If
    
    '本地或系统参数
    mstr结算方式 = zlDatabase.GetPara("缺省结算方式", glngSys, 1121)
            
    mstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    strTmp = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strTmp) = 1, strTmp, Mid(strTmp, 2, 1)))
    
    mbytInFun = 1
    Me.Show 1, frmParent
    ShowMeFromOut = mblnOK
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ShowMe(ByRef frmParent As Object, ByVal lng结帐ID As Long, ByVal cur实收金额 As Currency, _
        ByVal cur冲预交额 As Currency, ByVal str保险结算 As String, ByRef str收费结算 As String, _
        ByVal cur收费误差 As Currency, ByVal cur预交余额 As Currency, ByVal intInsure As Integer, _
        ByVal str缺省结算方式 As String, ByVal str缺省金额位数 As String, ByVal byt缺省分币方式 As Byte, ByRef cur缴款 As Currency) As Boolean
        
    mlng结帐ID = lng结帐ID
    mintInsure = intInsure
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
    
    ShowMe = mblnOK
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
        IIf(Val(txt缴款.Text) <> 0, "," & Val(txt缴款.Text) & "," & Val(txt找补.Text), "") & ",null,0)"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnOK = True
    mblnClickOK = True: Unload Me
    
    Exit Sub
errH:
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
    Dim strSql As String, i As Long, j As Long
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
    
    On Error GoTo errH
    strSql = _
        " Select Distinct B.编码,B.名称,B.性质,A.缺省标志" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where ((A.应用场合='收费' And B.性质<>3 And B.性质<>4) OR (B.性质=3 OR B.性质=4)) " & _
        "       And B.名称=A.结算方式(+) And B.性质<>5 And a.付款方式 Is Null" & _
        " Order by B.性质,B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    strSql = "Select 应用场合,结算方式 From 结算方式应用 Where 应用场合='收费' And 付款方式 Is Null"
    Set rs应用场合 = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
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
                For j = 0 To mshBalance.COLS - 1
                    mshBalance.Row = i: mshBalance.Col = j
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
        mshBalance.Row = mintDefault: mshBalance.Col = 0
        mshBalance.CellFontBold = True
        mshBalance.Col = 1
    Else        '结算方式没有缺省值,并且无现金方式的情况
        mshBalance.Row = 1: mshBalance.Col = 1
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
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtTmp_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim curPay As Currency
    
    If KeyAscii <> 13 Then
        If mshBalance.Col = 1 Then
            If KeyAscii = vbKeyEscape Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
            If InStr(txtTmp.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        Else '结算号码特殊字符限制
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0

        If mshBalance.Col = 2 Then
            '结算号码防拷贝特殊字符
            If InStr(txtTmp.Text, "'") > 0 Or InStr(txtTmp.Text, "|") > 0 Or InStr(txtTmp.Text, ",") > 0 Then
                Call Beep: Exit Sub
            End If
            mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = txtTmp.Text
        ElseIf mshBalance.Col = 1 Then
            If Not IsNumeric(txtTmp.Text) And Trim(txtTmp.Text) <> "" Then
                MsgBox "请输入正确的数值。", vbInformation, gstrSysName
                zlControl.TxtSelAll txtTmp: Exit Sub
            End If
            
            If Val(txtTmp.Text) <> 0 Then   '空字符的val为零
                txtTmp.Text = Format(Val(txtTmp.Text), "0.00")
                If Val(mshBalance.RowData(mshBalance.Row)) = PayType.现金 And mblnCent Then  '如果是在现金栏内输入,则进行分币处理
                    txtTmp.Text = Format(CentMoney(Val(txtTmp.Text)), "0.00")
                End If
                If Val(mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col)) = Val(txtTmp.Text) Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
                If txtTmp.Text = "0.00" Then
                    mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = ""
                Else
                    mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = Format(Val(txtTmp.Text), "0.00")
                End If
            Else
                If mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = "" Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
                mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = ""
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
            mshBalance.Col = 1
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
    If Not txtTmp.Visible And mshBalance.Row >= 1 And mshBalance.Col >= 1 And _
        mshBalance.RowData(mshBalance.Row) <> PayType.医保个人帐户 And mshBalance.RowData(mshBalance.Row) <> PayType.医保其它结算 Then
        With txtTmp
            .MaxLength = IIf(mshBalance.Col = 2, 30, 10)
            .Left = mshBalance.Left + mshBalance.CellLeft + 15
            .Top = mshBalance.Top + mshBalance.CellTop + (mshBalance.CellHeight - txtTmp.Height) / 2 - 15
            .Width = mshBalance.CellWidth - 60
            .ForeColor = mshBalance.CellForeColor
            .BackColor = mshBalance.CellBackColor
            .Alignment = IIf(mshBalance.Col = 1, 1, 0)
            .Text = mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If mshBalance.Col = 0 Then
            mshBalance.Col = 1
        ElseIf mshBalance.Row < mshBalance.Rows - 1 Then
            mshBalance.Row = mshBalance.Row + 1
            mshBalance.Col = 1
            If mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(0) - 2) > 1 Then
                mshBalance.TopRow = mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(1) - 2)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub mshBalance_KeyPress(KeyAscii As Integer)
    If Not txtTmp.Visible And mshBalance.Row >= 1 And mshBalance.Col > 0 And KeyAscii <> 13 And KeyAscii <> vbKeyEscape And _
            mshBalance.RowData(mshBalance.Row) <> PayType.医保个人帐户 And mshBalance.RowData(mshBalance.Row) <> PayType.医保其它结算 Then
        
        If mshBalance.Col = 1 Then
            If InStr("0123456789.-", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        Else '结算号码特殊字符限制
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        
        With txtTmp
            .MaxLength = IIf(mshBalance.Col = 2, 30, 10)
            .Left = mshBalance.Left + mshBalance.CellLeft + 15
            .Top = mshBalance.Top + mshBalance.CellTop + (mshBalance.CellHeight - .Height) / 2 - 15
            .Width = mshBalance.CellWidth - 60
            .ForeColor = mshBalance.CellForeColor
            .BackColor = mshBalance.CellBackColor
            .Alignment = IIf(mshBalance.Col = 1, 1, 0)
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
        ElseIf KeyAscii = Asc(".") And InStr(txt缴款.Text, ".") > 0 Then
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
    Dim strSql As String
        
    If curModiMoney = 0 Then
        strSql = "Select Nvl(费用余额,0) as 费用余额,Nvl(预交余额,0) as 预交余额 From 病人余额 Where 性质=1 And 类型=1 And 病人ID=[1]"
    Else
        strSql = "Select Nvl(费用余额,0)-[2] as 费用余额,Nvl(预交余额,0) as 预交余额 From 病人余额 Where 性质=1 And 类型=1 And 病人ID=[1]"
    End If
    On Error GoTo errH
    Set GetMoneyInfo = zlDatabase.OpenSQLRecord(strSql, "mdlOutExse", lng病人ID, curModiMoney)
    Exit Function
errH:
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
        curTmp = Format(curMoney - Int(curMoney), "0.0")
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

 

 
