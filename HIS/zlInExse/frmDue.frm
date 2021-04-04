VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "应收款登记"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9255
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
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
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   8925
   End
   Begin VB.TextBox txtPay 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
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
      Index           =   1
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox txtLack 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   120
      Width           =   1995
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   120
      Width           =   1995
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
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   4320
      Width           =   1995
   End
   Begin VB.TextBox txt找补 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Enabled         =   0   'False
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
      Left            =   6960
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   4320
      Width           =   1995
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5760
      TabIndex        =   10
      Top             =   5160
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   11
      Top             =   5160
      Width           =   1500
   End
   Begin VB.TextBox txtPay 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
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
      Index           =   0
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1380
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPay 
      Height          =   1665
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   2937
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
      FormatString    =   "^  结算方式  |^  结算金额  |^     结算号码     |^             备注           "
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBalance 
      Height          =   1665
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   2937
      _Version        =   393216
      Rows            =   5
      Cols            =   7
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
      FormatString    =   "^选择|^  单据号  |^  票据号  |^结帐人 |^   结帐日期   |^ 应收金额  |^    冲应收  "
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
      _Band(0).Cols   =   7
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "冲应收合计"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label lbl缴款 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现金缴款(J)"
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
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label lbl找补 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现金找补(&B)"
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
      Left            =   5520
      TabIndex        =   12
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label lblLack 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "付款差额"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmDue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mlng病人ID As Long

Private mlngDefalt As Long '缺省结算方式行
Private mlngCash As Long    '现金结算方式
Private mcurTotal As Currency   '选择和输入的冲应收款合计,小于等于应收款合计
Private mcurInsure As Currency  '医保的冲应收款合计,小于等于应收款合计
Private mstrBalance As String   '医保基金结算方式
Private mcurCheckInsure As Currency

Private Enum PAYCOL
    C0方式 = 0
    C1金额 = 1
    C2号码 = 2
    C3备注 = 3
End Enum
Private Enum BALANCECOL
    C0选择 = 0
    C1单据号 = 1
    C2票据号 = 2
    C3结帐人 = 3
    C4结帐日期 = 4
    C5应收金额 = 5
    C6冲应收 = 6
End Enum

Private Sub cmdCancel_Click()
    gblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strNO As String, Curdate As Currency, curInsure As Currency
    Dim arrSQL() As Variant, blnTrans As Boolean, curOther As Currency
    Dim curTemp As Currency
    
    If Val(txtLack.Text) > 0 Then
        MsgBox "付款金额不足要冲减的结帐款，请输入足够的结算金额!", vbInformation, gstrSysName
        mshPay.SetFocus: Exit Sub
    ElseIf Val(txtLack.Text) < 0 Then
        MsgBox "付款金额多于要冲减的结帐款，请检查结算金额!", vbInformation, gstrSysName
        mshPay.SetFocus: Exit Sub
    End If
    If mcurTotal = 0 Then   '没有输缴款，也没有选择应收款
        MsgBox "请选择并输入应收款的冲款额!", vbInformation, gstrSysName
        mshBalance.SetFocus: Exit Sub
    End If
    
    With mshBalance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, BALANCECOL.C0选择) = "√" And Val(.TextMatrix(i, BALANCECOL.C6冲应收)) <> 0 Then
                curTemp = curTemp + Val(.TextMatrix(i, BALANCECOL.C5应收金额))
            End If
        Next
    End With
    
    With mshPay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, PAYCOL.C1金额)) = 0 Then
                If (.TextMatrix(i, PAYCOL.C2号码) <> "" Or .TextMatrix(i, PAYCOL.C3备注) <> "") Then
                    If MsgBox("注意:第" & i & "行没有输入金额,但输入了" & IIf(.TextMatrix(i, PAYCOL.C2号码) <> "", "结算号码", "备注") & _
                        vbCrLf & "!该信息不会保存!确定要继续吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        .SetFocus: .Row = i: .Col = PAYCOL.C1金额: Exit Sub
                    End If
                End If
            End If
            If .RowData(i) = 4 Then
                curInsure = curInsure + Val(.TextMatrix(i, PAYCOL.C1金额))
            Else
                curOther = curOther + Val(.TextMatrix(i, PAYCOL.C1金额))
            End If
        Next
        If curInsure > mcurInsure Then
            MsgBox "注意:输入的医保金额过大(>" & Format(mcurInsure, "0.00") & "),请检查!", vbInformation, gstrSysName
            mshPay.SetFocus: Exit Sub
        End If
        If curOther > curTemp - mcurCheckInsure And curOther > 0 Then
            MsgBox "注意:输入的非医保金额过大(>" & Format(curTemp - mcurCheckInsure, "0.00") & "),请检查!", vbInformation, gstrSysName
            mshPay.SetFocus: Exit Sub
        End If
    End With
    
    On Error GoTo errH
    arrSQL = Array()
    strNO = zlDatabase.GetNextNo(18)
    Curdate = zlDatabase.Currentdate
    
    With mshPay
        For i = 1 To .Rows - 1
            If .TextMatrix(i, PAYCOL.C0方式) <> "" And Val(.TextMatrix(i, PAYCOL.C1金额)) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人缴款记录_Insert(" & mlng病人ID & ",'" & strNO & "','" & .TextMatrix(i, PAYCOL.C0方式) & "','" & _
                    .TextMatrix(i, PAYCOL.C2号码) & "'," & .TextMatrix(i, PAYCOL.C1金额) & ",'" & .TextMatrix(i, PAYCOL.C3备注) & "'," & _
                    "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.姓名 & "')"
            End If
        Next
    End With
    With mshBalance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, BALANCECOL.C0选择) = "√" And Val(.TextMatrix(i, BALANCECOL.C6冲应收)) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人缴款对照_Insert('" & strNO & "'," & .RowData(i) & "," & .TextMatrix(i, BALANCECOL.C6冲应收) & ")"
            End If
        Next
    End With
    
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
        
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_1", Me, "NO=" & strNO, 2)
       
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    mshPay.SetFocus
End Sub

Private Sub Form_Load()
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim i As Long, j As Long, strBalance() As String
    Dim curInsure As Currency
    
    gblnOK = False
    mcurTotal = 0
    mlngDefalt = 0
    mlngCash = 0
    Call RestoreWinState(Me, App.ProductName)
    
    On Error GoTo errH
    '应收款要求为正数
    strSql = "Select A.ID , A.单据号, A.票据号, A.结帐人, To_Char(A.结帐时间,'YYYY-MM-DD') 结帐日期, 应收金额 - Nvl(Sum(B.金额), 0) 应收金额," & vbNewLine & _
            "       应收金额 - Nvl(Sum(B.金额), 0) 冲应收" & vbNewLine & _
            "From (Select A.NO 单据号, A.实际票号 票据号, A.操作员姓名 结帐人, A.收费时间 结帐时间, A.ID, Sum(B.冲预交) 应收金额" & vbNewLine & _
            "       From 病人结帐记录 A, 病人预交记录 B, 结算方式 C" & vbNewLine & _
            "       Where A.病人id = [1] And A.记录状态 =1 And A.ID = B.结帐id And B.结算方式 = C.名称 And C.应收款 = 1" & vbNewLine & _
            "       Group By A.NO, A.实际票号, A.操作员姓名, A.收费时间, A.ID) A, 病人缴款对照 B" & vbNewLine & _
            "Where A.ID = B.结帐id(+) " & vbNewLine & _
            "Group By A.ID , A.单据号, A.票据号, A.结帐人, A.结帐时间, 应收金额" & vbNewLine & _
            "Having (应收金额 - Nvl(Sum(B.金额), 0))>0 " & vbNewLine & _
            "Order By 结帐日期,单据号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID)
    If rsTmp.RecordCount = 0 Then
        MsgBox "当前病人没有未缴清的应收款!", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    With mshBalance
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            .RowData(i) = rsTmp!ID
            .TextMatrix(i, BALANCECOL.C0选择) = "√"
            .TextMatrix(i, BALANCECOL.C1单据号) = rsTmp!单据号
            .TextMatrix(i, BALANCECOL.C2票据号) = "" & rsTmp!票据号
            .TextMatrix(i, BALANCECOL.C3结帐人) = rsTmp!结帐人
            .TextMatrix(i, BALANCECOL.C4结帐日期) = rsTmp!结帐日期
            .TextMatrix(i, BALANCECOL.C5应收金额) = rsTmp!应收金额
            .TextMatrix(i, BALANCECOL.C6冲应收) = Val("" & rsTmp!冲应收)
            rsTmp.MoveNext
        Next
    End With
    Call SetTotal
    
    strSql = "Select A.名称,A.缺省标志 缺省,A.性质,A.编码" & vbNewLine & _
            "From 结算方式 A, 结算方式应用 B" & vbNewLine & _
            "Where A.名称 = B.结算方式 And B.应用场合 = '结帐' And A.性质 Not In (3, 4, 9) " & vbNewLine & _
            " Union " & _
            "Select A.名称,A.缺省标志 缺省,A.性质,A.编码" & vbNewLine & _
            "From 结算方式 A, 结算方式应用 B" & vbNewLine & _
            "Where A.名称 = B.结算方式 And B.应用场合 = '结帐' And A.性质 = 4 And Nvl(A.应收款, 0) = 1" & vbNewLine & _
            "Order By 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If rsTmp.RecordCount = 0 Then
        MsgBox "没有设置结帐场合的结算方式,不能进行收款!", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    If mstrBalance <> "" Then
        strBalance = Split(mstrBalance, "|")
    End If
    
    curInsure = mcurInsure
    
    With mshPay
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, PAYCOL.C0方式) = rsTmp!名称
            .RowData(i) = rsTmp!性质
            
            If Val("" & rsTmp!缺省) = 1 Then mlngDefalt = i
            If rsTmp!性质 = 1 Then
                mlngCash = i
                If mlngDefalt = 0 Then mlngDefalt = i
            End If
            If rsTmp!性质 = 4 Then
                If mstrBalance <> "" Then
                    For j = 0 To UBound(strBalance)
                        If .TextMatrix(i, PAYCOL.C0方式) = Split(strBalance(j), ",")(0) Then
                            If curInsure > Val(Split(strBalance(j), ",")(1)) Then
                                .TextMatrix(i, PAYCOL.C1金额) = Val(Split(strBalance(j), ",")(1))
                                curInsure = curInsure - Val(Split(strBalance(j), ",")(1))
                            Else
                                If curInsure <> 0 Then .TextMatrix(i, PAYCOL.C1金额) = curInsure
                                curInsure = 0
                            End If
                        End If
                    Next j
                End If
            End If
            rsTmp.MoveNext
        Next
        If mlngDefalt = 0 Then mlngDefalt = 1
        If mcurTotal - mcurInsure <> 0 Then .TextMatrix(mlngDefalt, PAYCOL.C1金额) = mcurTotal - mcurInsure
    End With
    
    txtLack.Text = "0.00"
    If mlngCash = 0 Then txt缴款.Enabled = False
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng病人ID = 0
    Call SaveWinState(Me, App.ProductName)
End Sub

'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------

Private Sub mshPay_DblClick()
    If Not txtPay(0).Visible And mshPay.Row > 0 And mshPay.Col > PAYCOL.C0方式 Then
        Call SetTxtPay(0)
        txtPay(0).Text = mshPay.TextMatrix(mshPay.Row, mshPay.Col)
        txtPay(0).SelStart = 0: txtPay(0).SelLength = Len(txtPay(0).Text)
    End If
End Sub

Private Sub mshPay_KeyDown(KeyCode As Integer, Shift As Integer)
    If mshPay.Row <= 0 Then Exit Sub
    If KeyCode = 13 Then
        Call LocateMshpay
    ElseIf KeyCode = vbKeyDelete Then
        If mshPay.Col > PAYCOL.C0方式 Then
            mshPay.TextMatrix(mshPay.Row, mshPay.Col) = ""
            If mshPay.Col = PAYCOL.C1金额 Then Call SetLack
        End If
    End If
End Sub

Private Sub mshPay_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mshPay.Row <= 0 Then Exit Sub
        If Not txtPay(0).Visible And mshPay.Col > PAYCOL.C0方式 Then
            If mshPay.Col = PAYCOL.C1金额 Then
                '只能输正数
                If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            ElseIf mshPay.Col = PAYCOL.C2号码 Then
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If ZLCommFun.IsCharChinese(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
            End If
            
            Call SetTxtPay(0)
            txtPay(0).Text = Chr(KeyAscii)
            txtPay(0).SelStart = 1
        End If
    End If
End Sub

Private Sub mshPay_LeaveCell()
    txtPay(0).Visible = False
End Sub

Private Sub mshPay_Scroll()
    txtPay(0).Visible = False
End Sub

Private Sub mshPay_EnterCell()
    If mshPay.Col = PAYCOL.C3备注 Then
        txtPay(0).IMEMode = 1
        Call OpenIme(gstrIme)
    Else
        Call OpenIme
        txtPay(0).IMEMode = 3
    End If
End Sub

'--------------------------------------------------------------------------------------------

Private Sub mshBalance_DblClick()
    If mshBalance.Row <= 0 Then Exit Sub
    If Not txtPay(1).Visible And mshBalance.Col = BALANCECOL.C6冲应收 Then
        Call SetTxtPay(1)
        txtPay(1).Text = mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col)
        txtPay(1).SelStart = 0: txtPay(1).SelLength = Len(txtPay(1).Text)
    ElseIf mshBalance.Col = BALANCECOL.C0选择 Then
        Call SetBalanceSelect
    End If
End Sub

Private Sub mshBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If mshBalance.Row <= 0 Then Exit Sub
    If KeyCode = 13 Then
        Call LocateMshBalance
    ElseIf KeyCode = vbKeyDelete Then
        If mshBalance.Col = BALANCECOL.C6冲应收 Then
            mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = ""
            Call SetTotal
            Call SetLack
        End If
    ElseIf KeyCode = vbKeySpace And mshBalance.Col = BALANCECOL.C0选择 Then
        Call SetBalanceSelect
    End If
End Sub

Private Sub SetBalanceSelect()
    If mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = "" Then
        mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = "√"
    Else
        mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = ""
    End If
    If mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C6冲应收) = "" Then
        mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C6冲应收) = mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C5应收金额)
    End If
    
    Call AutoEquate
End Sub

Private Sub mshBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mshBalance.Row <= 0 Then Exit Sub
        If Not txtPay(1).Visible And mshBalance.Col = BALANCECOL.C6冲应收 Then
            '只能输正数,并且不能大于应收款额
            If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            Call SetTxtPay(1)
            txtPay(1).Text = Chr(KeyAscii)
            txtPay(1).SelStart = 1
        End If
    End If
End Sub

Private Sub mshBalance_LeaveCell()
    txtPay(1).Visible = False
End Sub

Private Sub mshBalance_Scroll()
    txtPay(1).Visible = False
End Sub

Private Sub mshBalance_EnterCell()
   Call OpenIme
End Sub



Private Sub LocateMshBalance()
    With mshBalance
        If .Row < .Rows - 1 Then '非末行的最后一列换行
            .Row = .Row + 1
            .Col = BALANCECOL.C6冲应收
            If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
            End If
            Call mshBalance_EnterCell
        Else '末行的最后一列
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub LocateMshpay()
    With mshPay
        '非末行的最后一列换行,非末行金额为零换行
        If .Row < .Rows - 1 And (.Col = .Cols - 1 Or .TextMatrix(.Row, PAYCOL.C1金额) = "" And .Col <> PAYCOL.C0方式) Then
            .Row = .Row + 1
            .Col = PAYCOL.C1金额
            If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
            End If
            mshPay.SetFocus: Call mshPay_EnterCell
        ElseIf .Row = .Rows - 1 And (.Col = .Cols - 1 Or .TextMatrix(.Row, PAYCOL.C1金额) = "" And .Col <> PAYCOL.C0方式) Then
        '末行的最后一列,末行金额为零,Tab
            Call ZLCommFun.PressKey(vbKeyTab)
        Else
             If .RowData(.Row) = 1 And .Col = PAYCOL.C1金额 Then '现金无需输结算号码
                .Col = .Col + 2
            Else
                .Col = .Col + 1
            End If
            mshPay.SetFocus: Call mshPay_EnterCell
        End If
    End With
End Sub

Private Sub SetTxtPay(Index As Integer)
    Dim mshTmp As MSHFlexGrid
    Set mshTmp = IIf(Index = 0, mshPay, mshBalance)
    
    With txtPay(Index)
        If Index = 0 Then
            .MaxLength = Val("" & Choose(mshPay.Col, 10, 30, 25))   '摘要最长50位,所以限制最大只能25个汉字
        Else
            .MaxLength = 10
        End If
        .Left = mshTmp.Left + mshTmp.CellLeft + 15
        .Top = mshTmp.Top + mshTmp.CellTop + (mshTmp.CellHeight - txtPay(Index).Height) / 2 - 15
        .Width = mshTmp.CellWidth - 60
        .ForeColor = mshTmp.CellForeColor
        .BackColor = mshTmp.CellBackColor
        If Index = 0 Then
            .Alignment = IIf(mshTmp.Col = PAYCOL.C1金额, 1, 0)
        Else
            .Alignment = IIf(mshTmp.Col = BALANCECOL.C6冲应收, 1, 0)
        End If
        .ZOrder: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub txtPay_GotFocus(Index As Integer)
    If Index = 0 Then
        If mshPay.Col = PAYCOL.C3备注 Then
            txtPay(Index).IMEMode = 1
            Call OpenIme(gstrIme)
        Else
            txtPay(Index).IMEMode = 3
        End If
    End If
End Sub

Private Sub txtPay_LostFocus(Index As Integer)
    txtPay(Index).Visible = False
    If Index = 0 Then
        txtPay(Index).IMEMode = 3: Call OpenIme
    End If
End Sub

Private Sub txtPay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPay(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPay(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Call SetWindowLong(txtPay(Index).hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtPay_Validate(Index As Integer, Cancel As Boolean)
    txtPay(Index).Visible = False
End Sub

Private Sub txtPay_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim i As Long, strTmp As String
    
    If KeyAscii <> 13 Then
        If IIf(Index = 0, mshPay.Col = PAYCOL.C1金额, mshBalance.Col = BALANCECOL.C6冲应收) Then
            If InStr(txtPay(Index).Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0: Exit Sub
            '只能输正数
            If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        ElseIf Index = 0 And mshPay.Col = PAYCOL.C2号码 Then '结算号码特殊字符限制
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ElseIf Index = 0 And mshPay.Col = PAYCOL.C3备注 Then    '备注
            If InStr("'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        strTmp = txtPay(Index).Text
        If IIf(Index = 0, mshPay.Col = PAYCOL.C1金额, mshBalance.Col = BALANCECOL.C6冲应收) Then
            If Not IsNumeric(strTmp) And strTmp <> "" Then
                MsgBox "请输入正确的数值。", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txtPay(Index)): Exit Sub
            End If
            '冲应收不能大于应收款额
            If Index = 1 Then
                If Val(strTmp) > Val(mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C5应收金额)) Then
                    MsgBox "冲应收金额不能大于应收金额!", vbInformation, gstrSysName
                    txtPay(Index).Text = mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C5应收金额)
                    Call zlControl.TxtSelAll(txtPay(Index)): Exit Sub
                End If
            End If
            
            strTmp = Format(strTmp, "0.00")    '不用进行分币处理
            If Index = 0 Then
                mshPay.TextMatrix(mshPay.Row, mshPay.Col) = IIf(Val(strTmp) = 0, "", strTmp)
                If mshPay.Row <> mlngDefalt Then
                    Call AutoEquate(False)
                Else
                    Call SetLack
                End If
                
            Else
                mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = IIf(Val(strTmp) = 0, "", strTmp)
                Call AutoEquate(True)
            End If
        Else
            mshPay.TextMatrix(mshPay.Row, mshPay.Col) = IIf(Index = 0 And mshPay.Col = PAYCOL.C2号码, UCase(strTmp), strTmp)
        End If
        txtPay(Index).Visible = False
        
        '换行与换列
        If Index = 0 Then
            Call LocateMshpay
        Else
            Call LocateMshBalance
        End If
    End If
End Sub

Private Sub AutoEquate(Optional blnRecalInsure As Boolean = False)
    Dim curPay As Currency, i As Long, j As Long
    Dim curInsure As Currency, blnHave As Boolean
    Dim strBalance() As String
    Dim intBalance As Integer
    '根据结帐单的选择及金额输入的变化，自动将差额补到缺省的结算方式上
    
    Call SetTotal
    
    If blnRecalInsure Then
        curInsure = mcurInsure
        If mstrBalance <> "" Then
            strBalance = Split(mstrBalance, "|")
            For i = 1 To mshPay.Rows - 1
                blnHave = False
                For j = 0 To UBound(strBalance)
                    If mshPay.TextMatrix(i, PAYCOL.C0方式) = Split(strBalance(j), ",")(0) Then
                        blnHave = True
                        intBalance = j
                    End If
                Next j
                If blnHave Then
                    If curInsure > Val(Split(strBalance(intBalance), ",")(1)) Then
                        mshPay.TextMatrix(i, PAYCOL.C1金额) = Format(Val(Split(strBalance(intBalance), ",")(1)), "0.00")
                        curInsure = curInsure - Val(Split(strBalance(intBalance), ",")(1))
                    Else
                        mshPay.TextMatrix(i, PAYCOL.C1金额) = Format(curInsure, "0.00")
                        curInsure = 0
                    End If
                End If
            Next
        End If
    End If
    
    For i = 1 To mshPay.Rows - 1
        If mshPay.TextMatrix(i, PAYCOL.C0方式) <> "" And i <> mlngDefalt Then
            curPay = curPay + Val(mshPay.TextMatrix(i, PAYCOL.C1金额))
        End If
    Next
    curPay = mcurTotal - curPay
    mshPay.TextMatrix(mlngDefalt, PAYCOL.C1金额) = IIf(curPay = 0, "", Format(curPay, "0.00"))
        
    Call SetLack
End Sub


Private Sub SetTotal()
    Dim i As Long, rsTmped As ADODB.Recordset
    Dim strBalanceIDs As String, blnHave As Boolean
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    mcurTotal = 0
    mcurInsure = 0
    mstrBalance = ""
    For i = 1 To mshBalance.Rows - 1
        If mshBalance.TextMatrix(i, BALANCECOL.C0选择) = "√" Then
            mcurTotal = mcurTotal + Val(mshBalance.TextMatrix(i, BALANCECOL.C6冲应收))
            strBalanceIDs = strBalanceIDs & "," & mshBalance.RowData(i)
        End If
    Next
    If strBalanceIDs <> "" Then
        strBalanceIDs = Mid(strBalanceIDs, 2)
        strSql = "Select Sum(A.冲预交) As 金额,A.结算方式" & vbNewLine & _
                "From 病人预交记录 A, 结算方式 B" & vbNewLine & _
                "Where a.结算方式 = b.名称 And b.性质 = 4 And Nvl(b.应收款, 0) = 1 And" & vbNewLine & _
                "      a.结帐id In (Select Column_Value From Table(f_Str2list([1]))) Group By A.结算方式"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalanceIDs)
        strSql = "Select Sum(a.金额) As 金额,a.结算方式" & vbNewLine & _
                "From 病人缴款记录 A, 病人缴款对照 B, 结算方式 C" & vbNewLine & _
                "Where a.No = b.缴款单 And b.结帐id In (Select Column_Value From Table(f_Str2list([1]))) And a.结算方式 = c.名称 And c.性质 = 4 And" & vbNewLine & _
                "      Nvl(c.应收款, 0) = 1 Group By a.结算方式"
        Set rsTmped = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalanceIDs)
        Do While Not rsTmp.EOF
            blnHave = False
            If rsTmped.RecordCount <> 0 Then
                rsTmped.MoveFirst
                Do While Not rsTmped.EOF
                    If Nvl(rsTmped!结算方式) = Nvl(rsTmp!结算方式) Then
                        blnHave = True
                        Exit Do
                    End If
                    rsTmped.MoveNext
                Loop
            End If
            If blnHave Then
                mcurInsure = mcurInsure + Val(Nvl(rsTmp!金额)) - Val(Nvl(rsTmped!金额))
                If Val(Nvl(rsTmp!金额)) - Val(Nvl(rsTmped!金额)) <> 0 Then mstrBalance = mstrBalance & "|" & Nvl(rsTmp!结算方式) & "," & Val(Nvl(rsTmp!金额)) - Val(Nvl(rsTmped!金额))
            Else
                mcurInsure = mcurInsure + Val(Nvl(rsTmp!金额))
                mstrBalance = mstrBalance & "|" & Nvl(rsTmp!结算方式) & "," & Nvl(rsTmp!金额)
            End If
            rsTmp.MoveNext
        Loop
    End If
    If mstrBalance <> "" Then mstrBalance = Mid(mstrBalance, 2)
'    If mcurInsure <> 0 Then
'        If mstrBalance <> "" Then mstrBalance = Mid(mstrBalance, 2)
'        strSql = "Select Sum(a.金额) As 金额" & vbNewLine & _
'                "From 病人缴款记录 A, 病人缴款对照 B, 结算方式 C" & vbNewLine & _
'                "Where a.No = b.缴款单 And b.结帐id In (Select Column_Value From Table(f_Str2list([1]))) And a.结算方式 = c.名称 And c.性质 = 4 And" & vbNewLine & _
'                "      Nvl(c.应收款, 0) = 1"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalanceIDs)
'        If Not rsTmp.EOF Then
'            mcurInsure = mcurInsure - Val(Nvl(rsTmp!金额))
'        End If
'    End If
    mcurCheckInsure = mcurInsure
    If mcurInsure > mcurTotal Then
        mcurInsure = mcurTotal
    End If
    txtTotal.Text = Format(mcurTotal, "0.00")
End Sub

Private Sub SetLack()
    Dim i As Long, curPay As Currency
    
    For i = 1 To mshPay.Rows - 1
        If mshPay.TextMatrix(i, PAYCOL.C0方式) <> "" Then
            curPay = curPay + Val(mshPay.TextMatrix(i, PAYCOL.C1金额))
        End If
    Next
    txtLack.Text = Format(mcurTotal - curPay, "0.00")
End Sub


'-------------------------------------------------------------------------------------------------------------
Private Sub txt缴款_Change()
    If Val(txt缴款.Text) = 0 Then txt找补.Text = "0.00": Exit Sub
    
    txt找补.Text = Format(Val(txt缴款.Text) - Val(mshPay.TextMatrix(mlngCash, PAYCOL.C1金额)), "0.00")
End Sub

Private Sub txt缴款_GotFocus()
     Call zlControl.TxtSelAll(txt缴款)
End Sub

Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txt缴款.Text) <> 0 Then
            If Val(txt找补.Text) >= 0 Then
                Call ZLCommFun.PressKey(vbKeyTab)
            Else
                MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
                txt缴款.SetFocus
                zlControl.TxtSelAll txt缴款
            End If
        Else
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc(".") And InStr(txt缴款.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt缴款_Validate(Cancel As Boolean)
    txt缴款.Text = Format(Trim(txt缴款.Text), "0.00")
End Sub
