VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMedicareReckoning 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "医保病人结帐校对"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9675
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
   ScaleHeight     =   5745
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消 (&C)"
      Height          =   420
      Left            =   8160
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   1005
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   5040
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
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
      TabIndex        =   8
      Top             =   4680
      Width           =   9885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确 定(&O)"
      Height          =   420
      Left            =   6465
      TabIndex        =   7
      Top             =   5055
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDeposit 
      Height          =   3345
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5900
      _Version        =   393216
      Rows            =   5
      Cols            =   4
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
      _Band(0).Cols   =   4
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   5025
      Width           =   1755
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
      Height          =   3345
      Left            =   5280
      TabIndex        =   3
      Top             =   960
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   5900
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
      FormatString    =   "^ 结算方式 |^ 结算金额 |^   结算号码  "
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
   Begin VB.Label lbl应缴 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应缴:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   7680
      TabIndex        =   16
      Tag             =   "应缴:"
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label lbl医保支付 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医保支付:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   5280
      TabIndex        =   15
      Tag             =   "医保支付:"
      Top             =   4440
      Width           =   1080
   End
   Begin VB.Label lbl冲预交 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "冲预交:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2280
      TabIndex        =   14
      Tag             =   "冲预交:"
      Top             =   4440
      Width           =   840
   End
   Begin VB.Label lbl预交余额 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "预交余额:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Tag             =   "预交余额:"
      Top             =   4440
      Width           =   1080
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
      Left            =   240
      TabIndex        =   4
      Top             =   5160
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
      Left            =   3240
      TabIndex        =   11
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label lblMargin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应补金额"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   360
      Width           =   960
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结帐金额"
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
      TabIndex        =   9
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "frmMedicareReckoning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mbytInFun As Byte '0-费用模块调用,1-医保模块调用
Private mlng结帐ID As Long
Private mlng病人ID As Long
Private mbln中途结帐 As Boolean     '出院结帐,未冲完的预交金额要退为现金
Private mstr保险结算 As String
Private mstr保险信息 As String      '保险类别,保险密码,保险帐号
Private mcur结帐金额 As Currency
Private mcur预交余额 As Currency
Private mintInsure As Integer       '用来判断是否支持分币处理
Private mcur缴款 As Currency
Private mcur应缴金额 As Currency
Private mstr医保号 As String
Private mcur个帐透支 As Currency
Private mintError As Integer
Private mcur收费误差 As Currency
Private mblnOK  As Boolean
Private mintDefault As Integer      '缺省结算方式行(为0表示没有)
Private mcurMediCare   As Currency  '医保结算合计,根据[mstr保险结算]计算
Private mblnClickOK As Boolean      '窗体只允许点确定退出
Private mblnCent As Boolean         '医保是否支持分币处理
Private mcur个帐余额 As Currency
Private mcur冲预交合计 As Currency
Private mcur预交合计 As Currency
Private mstr住院次数 As String  '住院次数:多个用逗号分离
Private mint预交类别 As Integer
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
Private mstrDec As String
Private mBytMoney As Byte '收费分币处理方法
Private mbytMCMode As Byte '医保病人身份证验模式,包括1-门诊,2-住院两种模式,0-表示非医保
Private mblnFirst As Boolean
Private Sub LoadBalance()
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim i As Long
    
    strSQL = "" & _
    "   Select B.性质,A.结算方式,A.冲预交,A.结算号码" & _
    "   From 病人预交记录 A, 结算方式 B " & _
    "   Where A.结算方式=B.名称(+) and A.结帐ID=[1] And Mod(A.记录性质,10)<>1  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结帐ID)
    If rsTemp.EOF Then Exit Sub
    
    With mshMoney
        For i = 1 To .Rows - 1
            If .RowData(i) <> PayType.医保个人帐户 And .RowData(i) <> PayType.医保其它结算 Then
                If i <> mintDefault Then
                    rsTemp.Filter = "结算方式='" & Trim(.TextMatrix(i, 0)) & "'"
                    If rsTemp.EOF = False Then
                        .TextMatrix(i, 1) = Format(Val(NVL(rsTemp!冲预交)), "0.00")
                        .TextMatrix(i, 2) = NVL(rsTemp!结算号码)
                    End If
                End If
            End If
        Next
    End With
    Call ShowMoney(False)
End Sub

Public Function ShowMeFromOut(ByRef frmParent As Object, _
    ByVal lng结帐ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lng病人ID As Long, strValue As String
    
    On Error GoTo errH
    mlng结帐ID = lng结帐ID
    
    strSQL = "Select a.病人ID,a.记录性质,a.结算方式,a.结算号码,b.性质 结算性质,a.冲预交,a.缴款单位,a.单位开户行,a.单位帐号 " & _
             "   From 病人预交记录 a,结算方式 b " & _
             "   Where a.记录状态 = 1 And a.结算方式 = B.名称 And 结帐id = [1]"
             
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "保险结算管理", lng结帐ID)
    mlng病人ID = Val("" & rsTmp!病人ID)
    
    mbln中途结帐 = True     '无法根据数据库信息区分,默认为最常用的方式:中途结帐,如果不是,操作员自已去输入每笔预交的冲款额,以便退现金
        
    rsTmp.Filter = "(记录性质=2 And 结算性质=3) or (记录性质=2 And 结算性质=4)"
    If rsTmp.RecordCount > 0 Then mstr保险信息 = zlCommFun.NVL(rsTmp!缴款单位, " ") & "," & zlCommFun.NVL(rsTmp!单位开户行, " ") & "," & zlCommFun.NVL(rsTmp!单位帐号, " ")
       
    
    rsTmp.Filter = 0    '不能取实收金额,因为结帐作废再结帐时,费用明细没有实收金额
    strSQL = "" & _
    "   Select Sum(nvl(结帐金额,0)) As 结帐金额" & _
    "   From (  Select nvl(结帐金额,0) as 结帐金额 From  From 门诊费用记录  Where 附加标志<>9 And 结帐id = [1]  UNION ALL  " & _
    "           Select nvl(结帐金额,0) as 结帐金额 From  From 住院费用记录  Where 附加标志<>9 And 结帐id = [1] ) "
         
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "保险结算管理", lng结帐ID)
    
    mcur结帐金额 = Val(NVL(rsTmp!结帐金额))
    
    
    '保险信息
    rsTmp.Filter = 0
    strSQL = "Select 结算方式,金额 From 医保核对表 Where 结帐id = [1] And 结算方式<>'现金'"  '医保管控的过程固定写入了一条"现金"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "保险结算管理", lng结帐ID)
    mstr保险结算 = ""   '结算方式|结算金额||
    For i = 1 To rsTmp.RecordCount
        mstr保险结算 = mstr保险结算 & "||" & rsTmp!结算方式 & "|" & rsTmp!金额
        rsTmp.MoveNext
    Next
    If mstr保险结算 <> "" Then mstr保险结算 = Mid(mstr保险结算, 3)
    
    
    mintInsure = 0
    If mlng病人ID <> 0 Then
        strSQL = "Select 险类 From 病案主页 Where 病人id = [1]" & _
                 " And 主页id = (Select Max(主页id) From 病案主页 Where 病人id = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "保险结算管理", mlng病人ID)
        If Not rsTmp.EOF Then mintInsure = zlCommFun.NVL(rsTmp!险类, 0)
        
    End If
            
    mstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 3, 1)))
    
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

Public Function ShowMe(ByRef frmParent As Object, ByVal lng结帐ID As Long, _
    ByVal lng病人ID As Long, ByVal bln中途结帐 As Boolean, _
        ByVal cur结帐金额 As Currency, ByVal str保险结算 As String, ByVal str保险信息 As String, _
        ByVal intInsure As Integer, ByVal str缺省金额位数 As String, ByVal byt缺省分币方式 As Byte, _
        ByVal cur缴款 As Currency, ByVal str医保号 As String, _
        ByVal bytMCMode As Byte, ByVal str住院次数 As String, ByVal int预交类别 As Integer) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保预结算与结算不一致的较对窗体
    '入参:bytMCMode=医保病人身份证验模式,包括1-门诊,2-住院两种模式,0-表示非医保
    '       int预交类别-预交类别:0-门诊和住院;1-门诊;2-住院
    '返回:校对成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-23 10:34:39
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng结帐ID = lng结帐ID: mstr住院次数 = str住院次数: mint预交类别 = int预交类别
    mlng病人ID = lng病人ID: mbln中途结帐 = bln中途结帐
    mstr保险结算 = str保险结算: mstr保险信息 = str保险信息     '用于医保存储:保险类别,保险密码,保险帐号
    mcur结帐金额 = cur结帐金额: mintInsure = intInsure: mstr医保号 = str医保号
    mcur缴款 = cur缴款: mstrDec = str缺省金额位数
    mBytMoney = byt缺省分币方式: mbytMCMode = bytMCMode
    If gblnLED Then 'Led才显示余额
        mcur个帐余额 = gclsInsure.SelfBalance(mlng病人ID, mstr医保号, IIf(mbytMCMode = 1, 10, 40), mcur个帐透支, mintInsure)
    End If
    
    mbytInFun = 0
    Me.Show 1, frmParent
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    mblnClickOK = True: Unload Me
End Sub

Private Sub cmdOK_Click()
    '检查数据
    Dim i As Long
    Dim str结帐结算 As String, str误差NO As String, str冲预交 As String
    
    If Val(txtMargin.Text) <> 0 Then
        If Val(txtMargin.Text) > 0 Then
            MsgBox "病人支付金额不足,请按所显示的差额补款。", vbExclamation, gstrSysName
            mshMoney.SetFocus: Exit Sub
        Else
            MsgBox "病人支付金额过多,请按所显示的差额退款。", vbExclamation, gstrSysName
            mshMoney.SetFocus: Exit Sub
        End If
    End If
    
    '更新数据
    str结帐结算 = ""
    For i = 1 To mshMoney.Rows - 1
        If Val(mshMoney.TextMatrix(i, 1)) <> 0 Then
            str结帐结算 = str结帐结算 & "||" & mshMoney.TextMatrix(i, 0) & "|" & Val(mshMoney.TextMatrix(i, 1)) & "|"
            
            If mshMoney.RowData(i) <> PayType.医保个人帐户 And mshMoney.RowData(i) <> PayType.医保其它结算 Then
                 'Oracle过程根据结算号码字段判断是否医保,所以缴费的结算号码不能含有,号
                 '结算方式|结算金额|结算号码||.....
                str结帐结算 = str结帐结算 & IIf(mshMoney.TextMatrix(i, 2) = "", " ", mshMoney.TextMatrix(i, 2))
            Else
                str结帐结算 = str结帐结算 & mstr保险信息
                '结算方式|结算金额|保险类别,保险密码,保险帐号||.....
            End If
        End If
    Next
    str结帐结算 = Mid(str结帐结算, 3)
    
    For i = 1 To mshDeposit.Rows - 1
        If Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1)) <> 0 Then     'ID|单据号|金额|记录状态||  Id为零表示冲预交余款(非第一次)
            str冲预交 = str冲预交 & "||" & mshDeposit.TextMatrix(i, 0) & "|" & mshDeposit.TextMatrix(i, 1) & "|" & Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1)) & "|" & Val(mshDeposit.RowData(i))
        End If
    Next
    If str冲预交 <> "" Then str冲预交 = Mid(str冲预交, 3)
    
    gstrSQL = "zl_住院收费结算_Update(" & mlng结帐ID & ",'" & IIf(str结帐结算 = "", "", str结帐结算) & "'" & IIf(str冲预交 = "", ",Null", ",'" & str冲预交 & "'") & _
            IIf(Val(txt缴款.Text) <> 0, "," & Val(txt缴款.Text) & "," & Val(txt找补.Text), "") & ")"
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
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
    Call LedDisplayBank
    
End Sub
Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim rs应用场合 As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim arrMediCare As Variant
    Dim bln允许个帐 As Boolean, blnExist As Boolean
    Dim str可用的医保结算方式 As String
    
    '变量初始
    mblnCent = gclsInsure.GetCapability(support分币处理, , mintInsure)
    mcur收费误差 = 0
    mblnOK = False
    mblnClickOK = False
    mintDefault = 0
    mcurMediCare = 0
    
    '确定和取消按钮
    If mbytInFun = 0 Then
        cmdOK.Left = cmdCancel.Left
        cmdCancel.Visible = False
    Else
        cmdCancel.Visible = True
    End If
    
    '显示预交明细
    Call AdjustDepost
    Set rsTmp = GetDepositBefor(mlng病人ID, mstr住院次数, mint预交类别)
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            mshDeposit.Redraw = False
            mshDeposit.Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                mshDeposit.Row = i
                mshDeposit.Col = mshDeposit.Cols - 1: mshDeposit.CellBackColor = txtMoney.BackColor
                mshDeposit.Col = mshDeposit.Cols - 2: mshDeposit.CellBackColor = 12900351
                
                mshDeposit.RowData(i) = IIf(IsNull(rsTmp!记录状态), 0, rsTmp!记录状态)
                mshDeposit.TextMatrix(i, 0) = rsTmp!ID
                mshDeposit.TextMatrix(i, 1) = rsTmp!NO

                mshDeposit.TextMatrix(i, 2) = Format(rsTmp!日期, "yyyy-MM-dd")
                mshDeposit.TextMatrix(i, 3) = IIf(IsNull(rsTmp!结算方式), "", rsTmp!结算方式)
                mshDeposit.TextMatrix(i, 4) = Format(rsTmp!金额, "0.00")
                mshDeposit.TextMatrix(i, 5) = Format(rsTmp!金额, "0.00")
                rsTmp.MoveNext
            Next
            mshDeposit.Row = 1: mshDeposit.Col = mshDeposit.Cols - 1
            mshDeposit.Redraw = True
        End If
    End If
    
    
    '显示保险结算及现付结算方式,即使不支持使用个帐,也更出来,反正医保的不允许改
    arrMediCare = Array()                   '结算方式|结算金额||
    If mstr保险结算 <> "" Then arrMediCare = Split(mstr保险结算, "||")
    
    On Error GoTo errH
    
    strSQL = _
        " Select Distinct B.编码,B.名称,B.性质,A.缺省标志,1 As 位置" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where ((A.应用场合='结帐' And B.性质<>3 And B.性质<>4) OR (B.性质=3 OR B.性质=4)) And B.名称=A.结算方式(+) " & _
        " Union " & _
        " Select 编码,名称,性质,缺省标志,0 As 位置" & _
        " From 结算方式 Where 性质=9 " & _
        " Order By 位置,性质,编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    strSQL = "Select 应用场合,结算方式 From 结算方式应用 Where 应用场合='结帐'"
    Call zlDatabase.OpenRecordset(rs应用场合, strSQL, Me.Caption)
    
    With mshMoney
        .ColAlignment(0) = 1    '结算方式左对齐
        .ColAlignment(1) = 7    '金额右对齐
        .Redraw = False
        .Rows = rsTmp.RecordCount + 1
        i = 1
        Do While Not rsTmp.EOF
            .RowData(i) = zlCommFun.NVL(rsTmp!性质, PayType.现金)                '用来判断是否可以修改金额,以及是否是现金
            .TextMatrix(i, 0) = rsTmp!名称
            .TextMatrix(i, 1) = "0.00"
            
            '缺省结算方式(没有则用现金) 不可能是医保
            If .RowData(i) <> PayType.医保个人帐户 And .RowData(i) <> PayType.医保其它结算 Then
                If zlCommFun.NVL(rsTmp!缺省标志, 0) = 1 Then mintDefault = i
                If zlCommFun.NVL(rsTmp!性质, 1) = 1 And mintDefault = 0 Then mintDefault = i
                If zlCommFun.NVL(rsTmp!性质, 1) = 9 And mintError = 0 Then
                    mintError = i
                    .Row = i
                    .Col = 0
                    .CellForeColor = vbRed
                End If
                i = i + 1
            Else
                '保险结算
                blnExist = False
                For j = 0 To UBound(arrMediCare)
                    If Split(arrMediCare(j), "|")(0) = rsTmp!名称 Then
                        blnExist = True
                        rs应用场合.Filter = "结算方式='" & rsTmp!名称 & "'"
                        
                        If rs应用场合.EOF And NVL(rsTmp!性质, 1) <> 9 Then
                            MsgBox "注意:结算方式[" & rsTmp!名称 & "]未设置应用于[结帐]场合,请到[结算方式管理]中设置!", vbInformation, gstrSysName
                        End If
                        
                        .TextMatrix(i, 1) = Split(arrMediCare(j), "|")(1)
                        .TextMatrix(i, 2) = ""    '无结算号码
                        mcurMediCare = mcurMediCare + Val(.TextMatrix(i, 1))
                        Exit For
                    End If
                Next
                
                If blnExist Then
                     For j = 0 To .Cols - 1
                         .Row = i: .Col = j: .CellBackColor = &HE7CFBA
                     Next
                     i = i + 1
                End If
                
                str可用的医保结算方式 = str可用的医保结算方式 & "," & rsTmp!名称
            End If
            rsTmp.MoveNext
        Loop
        .Rows = i
        .Redraw = True
    End With
    
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
    
    
    '结帐金额
    txtTotal.Text = Format(mcur结帐金额, mstrDec)
    
    '冲预交,结帐金额减去医保正式结帐后的余额
    Call ShowMoney(True)
    Call LoadBalance
    
    If mintDefault > 0 Then
        mshMoney.Row = mintDefault: mshMoney.Col = 0
        mshMoney.CellFontBold = True
        mshMoney.Col = 1
    Else        '结算方式没有缺省值,并且无现金方式的情况
        mshMoney.Row = 1: mshMoney.Col = 1
    End If

    
    txt缴款.Text = Format(mcur缴款, "0.00")
    Call LedDisplayBank
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ShowMoney(Optional ByVal blnAutoSet As Boolean) As String
'功能：设置和显示界面的各种金额

    Dim i As Long, j As Long
    Dim cur结帐合计 As Currency, curMoney As Currency
    Dim cur预交合计 As Currency, cur冲预交合计 As Currency, cur应缴金额 As Currency
    Dim bln存在补款 As Boolean  '只有当没有缺省结算方式,或者修改缺省结算方式的金额时,才有
        
    
    '设置自动冲预交额及余款的结算金额
    '---------------------------------------------------------------------------------------------
    If blnAutoSet Then
        '设置冲预交(结帐合计 - 保险合计)
        cur结帐合计 = mcur结帐金额 - mcurMediCare
        
        If mshDeposit.TextMatrix(1, 0) <> "" Then   '可能没有预交,全部现款
            If Not mbln中途结帐 Then
                '出院结帐全部都冲完(冲多了就退现付)
                For i = 1 To mshDeposit.Rows - 1
                    mshDeposit.TextMatrix(i, mshDeposit.Cols - 1) = Format(Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 2)), "0.00")
                    cur结帐合计 = cur结帐合计 - Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
                Next
            Else
                '中途结帐只冲足够的
                For i = 1 To mshDeposit.Rows - 1
                    If cur结帐合计 = 0 Then
                        mshDeposit.TextMatrix(i, mshDeposit.Cols - 1) = "0.00"
                    Else
                        If Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 2)) <= Format(cur结帐合计, "0.00") Then
                            mshDeposit.TextMatrix(i, mshDeposit.Cols - 1) = Format(Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 2)), "0.00")
                        Else
                            mshDeposit.TextMatrix(i, mshDeposit.Cols - 1) = Format(cur结帐合计, "0.00")
                        End If
                        cur结帐合计 = cur结帐合计 - Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
                    End If
                Next
            End If
        End If
        
        '剩余应缴部份尝试设置到缺省结算方式    '判断是否应该进行分币处理
        If mintDefault <> 0 Then
            If mshMoney.RowData(mintDefault) = PayType.现金 And mblnCent Then '现金时要进行分币处理
                mshMoney.TextMatrix(mintDefault, 1) = Format(CentMoney(cur结帐合计), "0.00")
            Else
                mshMoney.TextMatrix(mintDefault, 1) = Format(cur结帐合计, "0.00")
            End If
        Else
            bln存在补款 = True
        End If
    
    '修改冲预交或结算金额后
    Else
        cur结帐合计 = mcur结帐金额 - GetSumMoney
        
        If mintDefault <> 0 And (Not Me.ActiveControl Is mshMoney Or _
                                Me.ActiveControl Is mshMoney And mintDefault <> mshMoney.Row) Then
            If mshMoney.RowData(mintDefault) = PayType.现金 And mblnCent Then '现金时要进行分币处理
                mshMoney.TextMatrix(mintDefault, 1) = Format(Val(mshMoney.TextMatrix(mintDefault, 1)) + CentMoney(cur结帐合计), "0.00")
            Else
                mshMoney.TextMatrix(mintDefault, 1) = Format(Val(mshMoney.TextMatrix(mintDefault, 1)) + cur结帐合计, "0.00")
            End If
        Else
            bln存在补款 = True
        End If
    End If
    
        
    '显示差额及误差
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetSumMoney(cur预交合计, cur冲预交合计, cur应缴金额)
    mcur收费误差 = Format(curMoney - mcur结帐金额, mstrDec)
    If bln存在补款 Then
        txtMargin.Text = Format(mcur结帐金额 - curMoney, "0.00")
    Else
        txtMargin.Text = "0.00"
    End If
    mshMoney.TextMatrix(mintError, 1) = Format(-1 * mcur收费误差, mstrDec)
    
    lbl预交余额.Caption = lbl预交余额.Tag & Format(cur预交合计, "0.00")
    lbl预交余额.ToolTipText = "本次未冲预交之前的预交余额"
    mcur预交合计 = cur预交合计
    lbl冲预交.Caption = lbl冲预交.Tag & Format(cur冲预交合计, "0.00")
    mcur冲预交合计 = cur冲预交合计
    lbl医保支付.Caption = lbl医保支付.Tag & Format(mcurMediCare, "0.00")
    lbl应缴.Caption = lbl应缴.Tag & Format(cur应缴金额, "0.00")
    mcur应缴金额 = cur应缴金额
    
    lbl预交余额.Left = mshDeposit.Left
    lbl冲预交.Left = lbl预交余额.Left + lbl预交余额.Width + 600
    lbl医保支付.Left = mshMoney.Left
    lbl应缴.Left = lbl医保支付.Left + lbl医保支付.Width + 600
    Call Calc找补
    Call LedDisplayBank
End Function

Private Sub mshMoney_GotFocus()
        Call LedDisplayBank
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    Dim blnCent As Boolean, i As Long
    
    If KeyAscii <> 13 Then        '输入限制
        If InStr(txtMoney.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0: Beep: Exit Sub
        
        If txtMoney.Left > mshMoney.Left Then   '结算输入
            If mshMoney.Col = mshMoney.Cols - 1 Then    '结算号码,逗号用来在过程中判断是否是医保结算方式
                If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
            Else
                If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        Else    '预交输入
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = 0
         '结算输入确认
        If txtMoney.Left > mshMoney.Left Then
            If mshMoney.Col = mshMoney.Cols - 1 Then    '输入结算号
                If InStr(txtMoney.Text, "'") > 0 Or InStr(txtMoney.Text, "|") > 0 Or InStr(txtMoney.Text, ",") > 0 Then
                    Exit Sub
                End If
                
                mshMoney.TextMatrix(mshMoney.Row, mshMoney.Col) = Trim(txtMoney.Text)
                txtMoney.Visible = False
            Else
                If Trim(txtMoney.Text) = "" Or Not IsNumeric(Trim(txtMoney.Text)) Then
                    zlControl.TxtSelAll txtMoney: Call Beep: Exit Sub
                End If
                If mshMoney.RowData(mshMoney.Row) = PayType.现金 And mblnCent Then
                    txtMoney.Text = Format(CentMoney(Val(txtMoney.Text)), "0.00")
                End If
                                
                If Val(mshMoney.TextMatrix(mshMoney.Row, mshMoney.Col)) <> Format(Val(txtMoney.Text), "0.00") Then
                    mshMoney.TextMatrix(mshMoney.Row, mshMoney.Col) = Format(Val(txtMoney.Text), "0.00")
                    txtMoney.Visible = False
                    mshMoney.SetFocus   '必须在先,ShowMoney中以此判断
                    
                    Call ShowMoney
                Else
                    txtMoney.Visible = False
                    mshMoney.SetFocus
                End If
            End If
            
            If mshMoney.Col < mshMoney.Cols - 2 Then
                mshMoney.Col = mshMoney.Col + 1
            Else
                If mshMoney.Row = mshMoney.Rows - 1 Then
                    '下一控件处理
                    If Get应缴 > 0 And txt缴款.Visible Then
                        txt缴款.SetFocus
                    ElseIf cmdOK.Visible And cmdOK.Enabled Then
                        cmdOK.SetFocus
                    End If
                Else
                    '下一行处理
                    If mshMoney.RowData(mshMoney.Row) = PayType.非医保非现金 Then
                       If mshMoney.Col = mshMoney.Cols - 2 Then
                            mshMoney.Col = mshMoney.Cols - 1
                       Else
                            mshMoney.Row = mshMoney.Row + 1
                            mshMoney.Col = mshMoney.Cols - 2
                       End If
                    Else
                        mshMoney.Row = mshMoney.Row + 1
                        mshMoney.Col = mshMoney.Cols - 2
                    End If
                    
                    If mshMoney.Row - (mshMoney.Height \ mshMoney.RowHeight(0) - 2) > 1 Then
                        mshMoney.TopRow = mshMoney.Row - (mshMoney.Height \ mshMoney.RowHeight(1) - 2)
                    End If
                End If
            End If
        
        '预交输入确认
        Else
            If Trim(txtMoney.Text) = "" Or Not IsNumeric(Trim(txtMoney.Text)) Then
                zlControl.TxtSelAll txtMoney: Call Beep: Exit Sub
            End If
            
            '修改不能超过上限
            If Val(txtMoney.Text) > Val(mshDeposit.TextMatrix(mshDeposit.Row, 4)) Then
                txtMoney.Text = Val(mshDeposit.TextMatrix(mshDeposit.Row, 4))
            End If
            
            If Val(mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.Col)) <> Format(Val(txtMoney.Text), "0.00") Then
                mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.Col) = Format(Val(txtMoney.Text), "0.00")
                txtMoney.Visible = False
                mshDeposit.SetFocus '必须在先
                
                Call ShowMoney
            Else
                txtMoney.Visible = False
                mshDeposit.SetFocus
            End If
            
            If mshDeposit.Row = mshDeposit.Rows - 1 Then
                '下一控件处理
                mshMoney.SetFocus
            Else
                '下一行处理
                mshDeposit.Row = mshDeposit.Row + 1
                If mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(0) - 2) > 1 Then
                    mshDeposit.TopRow = mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(1) - 2)
                End If
                mshDeposit.Col = mshDeposit.Cols - 1
            End If
        End If
        
        If Val(txt缴款.Text) > 0 Then Call txt缴款_Change
    End If
End Sub

Private Sub txtMoney_LostFocus()
    txtMoney.Visible = False
End Sub

Private Sub txtMoney_Validate(Cancel As Boolean)
    If txtMoney.Visible Then Call txtMoney_KeyPress(13)
End Sub

Private Sub AdjustDepost()
    Dim bln As Boolean
    With mshDeposit
        bln = .Redraw
        .Redraw = False
        .Clear
        .Rows = 2: .Cols = 6
        
        .TextMatrix(0, 1) = "单据号"
        .TextMatrix(0, 2) = "日期"
        .TextMatrix(0, 3) = "结算方式"
        .TextMatrix(0, 4) = "余额"
        .TextMatrix(0, 5) = "冲预交"
        
        .ColAlignmentFixed(1) = 4: .ColAlignment(1) = 1
        .ColAlignmentFixed(2) = 4: .ColAlignment(2) = 6
        .ColAlignmentFixed(3) = 1: .ColAlignment(3) = 1
        .ColAlignmentFixed(4) = 4: .ColAlignment(4) = 7
        .ColAlignmentFixed(5) = 4: .ColAlignment(5) = 7
        
        .ColWidth(0) = 0
        
        .ColWidth(1) = 1100
        .ColWidth(2) = 1050
        .ColWidth(3) = 620
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        
        .Row = 1: .Col = .Cols - 1
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        
        .Redraw = bln
    End With
End Sub
Private Function GetDepositBefor(ByVal lng病人ID As Long, _
    ByVal str住院次数 As String, ByVal int预交类别 As Integer) As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人本次医保结算之前的剩余预交款明细,包含本次冲销的预交
    '入参:lng病人ID-病人ID
    '       str住院次数-住院次数,如:1,2,3
    '       int预交类别-预交类别:0-门诊和住院;1-门诊;2-住院
    '出参:
    '返回:本次医保结算之前的剩余预交款明细,
    '编制:刘兴洪
    '日期:2013-11-14 11:36:01
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strSub1 As String
    Dim strWherePage As String, strPages As String
    On Error GoTo errH
    strPages = "," & str住院次数 & ","
    
    strWherePage = IIf(str住院次数 = "", "", " And instr([3],','||Nvl(A.主页ID,0)||',')>0")
    If int预交类别 <> 0 Then
        strWherePage = strWherePage & " And A.预交类别 =[4]"
    End If
    
    '该子查询用于消除预交款收费及退费时的一正一负,注意系统允许结过帐的预交款进行预交退费,需要加上记录状态判断
    strSub1 = _
    "   Select NO,Sum(Nvl(A.金额,0)) as 金额 " & _
    "   From 病人预交记录 A" & _
    "   Where (A.结帐ID Is Null " & strWherePage & " Or A.结帐ID=[1]) And Nvl(A.金额, 0)<>0 And A.病人ID=[2]" & _
    "   Group by NO  " & _
    "   Having Sum(Nvl(A.金额,0))<>0"

    strSQL = _
        "   Select A.ID,A.记录状态,A.NO,A.收款时间 as 日期,A.结算方式,Nvl(A.金额,0) as 金额" & _
        "   From 病人预交记录 A,(" & strSub1 & ") B" & _
        "   Where (A.结帐ID Is Null " & strWherePage & " Or A.结帐ID=[1]) And Nvl(A.金额,0)<>0" & _
        "           And A.结算方式 Not IN(Select 名称 From 结算方式 Where 性质=5)" & _
        "           And A.NO=B.NO And A.病人ID=[2]" & _
        "   Union All" & _
        "   Select 0 as ID,记录状态,NO,Min(收款时间) as 日期,结算方式,Sum(Nvl(金额,0)-Nvl(冲预交,0)) as 金额" & _
        "   From 病人预交记录 A" & _
        "   Where 记录性质 IN(1,11) And 结帐ID is Not NULL " & _
        "           And 结帐ID<>[1] And Nvl(金额,0)<>Nvl(冲预交,0) And 病人ID=[2] " & strWherePage & _
        "   Having Sum(Nvl(金额,0)-Nvl(冲预交,0))<>0" & _
        "   Group by 记录状态,NO,结算方式" & _
        "   Order by ID,日期,NO,结算方式"
        
    Set GetDepositBefor = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结帐ID, mlng病人ID, strPages, int预交类别)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetSumMoney(Optional ByRef cur预交合计 As Currency, Optional ByRef cur冲预交合计 As Currency, Optional ByRef cur应缴金额 As Currency) As Currency
    Dim i As Long
    Dim curMoney As Currency
    
    cur预交合计 = 0: cur冲预交合计 = 0: cur应缴金额 = 0
    
    If mshDeposit.TextMatrix(1, 0) <> "" Then
        For i = 1 To mshDeposit.Rows - 1
            curMoney = curMoney + Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
            cur预交合计 = cur预交合计 + Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 2))
            cur冲预交合计 = cur冲预交合计 + Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
        Next
    End If
    For i = 1 To mshMoney.Rows - 1
        If IsNumeric(mshMoney.TextMatrix(i, 1)) Then
            If Not mshMoney.TextMatrix(i, 0) Like "*误差*" Then
                curMoney = curMoney + Val(mshMoney.TextMatrix(i, 1))
            End If
            If mshMoney.RowData(i) <> PayType.医保个人帐户 And mshMoney.RowData(i) <> PayType.医保其它结算 And Not mshMoney.TextMatrix(i, 0) Like "*误差*" Then
                cur应缴金额 = cur应缴金额 + Val(mshMoney.TextMatrix(i, 1))
            End If
        End If
    Next
    
    GetSumMoney = curMoney
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnClickOK Then Cancel = 1
End Sub

Private Sub mshDeposit_DblClick()
    If Not txtMoney.Visible And mshDeposit.Row >= 1 And mshDeposit.Col = mshDeposit.Cols - 1 Then
        With txtMoney
            .Left = mshDeposit.Left + mshDeposit.CellLeft + 15
            .Top = mshDeposit.Top + mshDeposit.CellTop + (mshDeposit.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshDeposit.CellWidth - 60
            .ForeColor = mshDeposit.CellForeColor
            .BackColor = mshDeposit.CellBackColor
            .Alignment = 1
            .Text = mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If mshDeposit.Col = 0 Then
            mshDeposit.Col = mshDeposit.Col + 1
        ElseIf mshDeposit.Row < mshDeposit.Rows - 1 Then
            mshDeposit.Row = mshDeposit.Row + 1
            mshDeposit.Col = mshDeposit.Cols - 1
            If mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(0) - 2) > 1 Then
                mshDeposit.TopRow = mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(1) - 2)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub mshDeposit_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And KeyAscii <> 13 And KeyAscii <> vbKeyEscape Then
        If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        With txtMoney
            .Left = mshDeposit.Left + mshDeposit.CellLeft + 15
            .Top = mshDeposit.Top + mshDeposit.CellTop + (mshDeposit.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshDeposit.CellWidth - 60
            .ForeColor = mshDeposit.CellForeColor
            .BackColor = mshDeposit.CellBackColor
            .Alignment = 1
            .Text = Chr(KeyAscii)
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshMoney_DblClick()
    If Not txtMoney.Visible And mshMoney.Row >= 1 And mshMoney.Col > 0 And _
        mshMoney.RowData(mshMoney.Row) <> PayType.医保个人帐户 And mshMoney.RowData(mshMoney.Row) <> PayType.医保其它结算 And Not mshMoney.TextMatrix(mshMoney.Row, 0) Like "*误差*" Then
        
        With txtMoney
            .MaxLength = IIf(mshMoney.Col = 2, 30, 10)
            .Left = mshMoney.Left + mshMoney.CellLeft + 15
            .Top = mshMoney.Top + mshMoney.CellTop + (mshMoney.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshMoney.CellWidth - 60
            .ForeColor = mshMoney.CellForeColor
            .BackColor = mshMoney.CellBackColor
            .Alignment = IIf(mshMoney.Col = 2, 0, 1)
            .Text = mshMoney.TextMatrix(mshMoney.Row, mshMoney.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mshMoney.Row >= 1 Then
        If mshMoney.Col = 0 Then
            mshMoney.Col = mshMoney.Col + 1
        Else
            If mshMoney.Row < mshMoney.Rows - 1 Then
                
                If mshMoney.RowData(mshMoney.Row) = PayType.非医保非现金 Then
                   If mshMoney.Col = mshMoney.Cols - 2 Then
                        mshMoney.Col = mshMoney.Cols - 1
                   Else
                        mshMoney.Row = mshMoney.Row + 1
                        mshMoney.Col = mshMoney.Cols - 2
                   End If
                Else
                    mshMoney.Row = mshMoney.Row + 1
                    mshMoney.Col = mshMoney.Cols - 2
                End If
                If mshMoney.Row - (mshMoney.Height \ mshMoney.RowHeight(0) - 2) > 1 Then
                    mshMoney.TopRow = mshMoney.Row - (mshMoney.Height \ mshMoney.RowHeight(1) - 2)
                End If
            Else
                If Get应缴 > 0 Then
                    txt缴款.SetFocus
                Else
                    cmdOK.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub mshMoney_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And mshMoney.Row >= 1 And mshMoney.Col > 0 And KeyAscii <> 13 And KeyAscii <> vbKeyEscape And _
         mshMoney.RowData(mshMoney.Row) <> PayType.医保个人帐户 And mshMoney.RowData(mshMoney.Row) <> PayType.医保其它结算 Then
                        
        If mshMoney.Col = 1 Then
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        Else '结算号码特殊字符限制,逗号用来在过程中判断是否是医保结算方式
            If InStr("'||,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        
        With txtMoney
            .MaxLength = IIf(mshMoney.Col = 2, 30, 10)
            .Left = mshMoney.Left + mshMoney.CellLeft + 15
            .Top = mshMoney.Top + mshMoney.CellTop + (mshMoney.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshMoney.CellWidth - 60
            .ForeColor = mshMoney.CellForeColor
            .BackColor = mshMoney.CellBackColor
            .Alignment = IIf(mshMoney.Col = 2, 0, 1)
            .Text = UCase(Chr(KeyAscii))
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Function Get应缴() As Currency
    Dim i As Long
    For i = 1 To mshMoney.Rows - 1
        If mshMoney.RowData(i) = PayType.现金 Then
            Get应缴 = Val(mshMoney.TextMatrix(i, 1))
            Exit Function
        End If
    Next
End Function

 
Private Sub txt缴款_Change()
    Call Calc找补
End Sub
 
Private Sub txt缴款_GotFocus()
    Dim curTotal As Currency
    Call zlControl.TxtSelAll(txt缴款)
    If Not gblnLED Then Exit Sub
    
    curTotal = Get应缴
    '#21 1234.56   --请您付款一千二百三十四点五六元  J
    '#22 1234.56   --预收一千二百三十四点五六元 Y
    '#23 1234.56   --找零一千二百三十四点五六元 Z
    zl9LedVoice.DisplayBank ("")
    If curTotal >= 0 Then
        zl9LedVoice.Speak "#21 " & curTotal
    Else
        zl9LedVoice.Speak "#23 " & Abs(curTotal)
    End If
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
'    If Val(txt缴款.Text) = 0 Then Exit Sub
    
'    If CSng(txt找补.Tag) < 0 Then
'        MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
'        Call zlControl.TxtSelAll(txt缴款): txt缴款.SetFocus
'        Cancel = True: Exit Sub
'    End If
    If Not gblnLED Then Exit Sub
    zl9LedVoice.DispCharge Format(Get应缴, "0.00"), Val(txt缴款.Text), Val(txt找补.Tag)
    zl9LedVoice.Speak "#22 " & txt缴款.Text
    zl9LedVoice.Speak "#23 " & CSng(txt找补.Tag)
    zl9LedVoice.Speak "#3"                  '#3  --请当面点清, 谢谢!
End Sub

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
'         6.五舍六入:eg:0.15=0.10:0.16=0.2:   刘兴洪 问题:34519  日期:2010-12-06 09:58:02
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
        CentMoney = Format(FormatEx(curMoney, 1), "0.00")
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
 
Private Sub LedDisplayBank()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示医保结算信息
    '编制:刘兴洪
    '日期:2013-10-23 14:50:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double, i As Long
    Dim str个人帐户 As String, str医保其他结算 As String, str老一卡通 As String, str普通结算 As String
    Dim varPara  As Variant, str结算方式 As String
    Dim cur个帐余额 As Currency, dbl现金 As Double, dblMoney As Double
    Dim str医保结算 As String
    
    If Not gblnLED Then Exit Sub
    zl9LedVoice.DisplayBank ""
    If mblnFirst = True Then Exit Sub
    
    str医保结算 = "||帐户余额:" & Format(mcur个帐余额, "0.00")
    With mshMoney
        For i = 1 To .Rows - 1
            '医保交易
            str结算方式 = Trim(.TextMatrix(i, 0))
            If str结算方式 <> "" Then
                dblMoney = Val(.TextMatrix(i, 1))
                Select Case .RowData(i)
                Case PayType.医保个人帐户
                    str个人帐户 = str个人帐户 & "||" & str结算方式 & ":" & Format(dblMoney, "0.00")
                Case PayType.医保其它结算
                    str医保其他结算 = str医保其他结算 & "||" & str结算方式 & ":" & Format(dblMoney, "0.00")
                Case PayType.现金
                    dbl现金 = dblMoney
                Case Else
                    str普通结算 = str普通结算 & "||" & str结算方式 & ":" & Format(dblMoney, "0.00")
                End Select
            End If
        Next
    End With
    str结算方式 = ""
    If str个人帐户 <> "" Then str医保结算 = str医保结算 & str个人帐户
    If str医保其他结算 <> "" Then str医保结算 = str医保结算 & str医保其他结算
    
    If str医保结算 <> "" Then str结算方式 = str结算方式 & "||医保结算:" & str医保结算
    If str普通结算 <> "" Then str结算方式 = str结算方式 & "" & str普通结算
    If mcur冲预交合计 <> 0 Then str结算方式 = str结算方式 & "||" & "冲预交:" & Format(mcur冲预交合计, "0.00")
    
    If str结算方式 = "" Then Exit Sub
    str结算方式 = Mid(str结算方式, 3)
    varPara = Split(str结算方式, "||")
    
    dblMoney = Val(txt缴款.Text) - dbl现金
    zl9LedVoice.DisplayBank "总费用" & Format(txtTotal.Text, "0.00"), "预交款" & Format(mcur预交合计, "0.00"), _
            "冲预交" & Format(mcur冲预交合计, "0.00"), IIf(dblMoney > 0, "找补", "应缴") & Format(Abs(dblMoney), "0.00")
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
    'zl9LedVoice.Speak "#21 " & Format(mcur应缴金额, "0.00")
End Sub

 
Private Sub txt找补_Change()
    txt找补.Tag = ""
End Sub
Private Sub Calc找补()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算找补
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-01-12 17:41:47
    '问题:27360
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl找补 As Double
    Dim cur现金 As Currency, i As Long

    If Val(txt缴款.Text) = 0 Then txt找补.Text = "0.00"
    dbl找补 = FormatEx(Val(txt缴款.Text) - Get应缴, 2)
    txt找补.Text = Format(Abs(dbl找补), "0.00")
    txt找补.Tag = dbl找补
    If dbl找补 <= 0 Then
        lbl找补.Caption = "收款"
        lbl找补.ForeColor = &H0&
    Else
        lbl找补.Caption = "找补"
        lbl找补.ForeColor = vbRed   '35830
    End If
    txt找补.ForeColor = lbl找补.ForeColor
End Sub
