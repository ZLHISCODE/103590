VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMultiBalance 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "收费结算"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMultiBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtOwe 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   210
      Width           =   2010
   End
   Begin VB.TextBox txtTmp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   6330
      TabIndex        =   12
      Top             =   4530
      Width           =   1400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   4800
      TabIndex        =   11
      Top             =   4530
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      Height          =   150
      Left            =   -90
      TabIndex        =   13
      Top             =   4215
      Width           =   7845
   End
   Begin VB.TextBox txt找补 
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
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   3720
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
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   1215
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   3720
      Width           =   1995
   End
   Begin VB.TextBox txtPay 
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
      ForeColor       =   &H00C00000&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   135
      Width           =   1995
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPay 
      Height          =   2385
      Left            =   195
      TabIndex        =   6
      Top             =   1080
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   4207
      _Version        =   393216
      Rows            =   7
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
      FormatString    =   "^ 结算方式 |^  结算金额  |^    结算号码    |^          备注    "
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
   Begin VB.Label lblOwe 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "付款差额"
      Height          =   240
      Left            =   3690
      TabIndex        =   2
      Top             =   255
      Width           =   960
   End
   Begin VB.Label lbl找补 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现金找补"
      Height          =   240
      Left            =   3690
      TabIndex        =   9
      Top             =   3810
      Width           =   960
   End
   Begin VB.Label lbl缴款 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现金缴款"
      Height          =   240
      Left            =   195
      TabIndex        =   7
      Top             =   3810
      Width           =   960
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "付款明细"
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblPay 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "付款合计"
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   255
      Width           =   960
   End
End
Attribute VB_Name = "frmMultiBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mintDefault As Integer '缺省结算方式行(为0表示没有)
Private mintInsure As Integer '入:当为医保病人时才传入
Private mlng病人ID As Long '入:当为医保病人时才传入
Private mcurPay As Currency '入:排开保险及冲预交后的应缴合计(未处理分币和小数点)；出:实际支付合计(处理之后)
Private mstrBalance As String '入/出:结算方式|结算金额|结算号码|摘要
Private mcurError As Currency '出:误差金额
Private mrs结算方式 As ADODB.Recordset
Private mlngPayRow As Long, mstr应付结算方式 As String
Private mcur缴款 As Currency    '记录当次结算的缴款和找补金额
Private mcur找补 As Currency
Private mcur现金 As Currency  '35135
Private mcurOneCard As Currency '一卡通余额,当一卡通余额不足缴款金额时才传入
Private mblnHotKey As Boolean

Private Enum COLS
    C0方式 = 0
    C1金额 = 1
    C2号码 = 2
    C3备注 = 3
End Enum


Public Function ShowMe(frmParent As Object, _
    ByVal intInsure As Integer, ByVal lng病人ID As Long, curPay As Currency, _
    strBalance As String, curError As Currency, rs结算方式 As ADODB.Recordset, _
    cur缴款 As Currency, cur找补 As Currency, CurOneCard As Currency, _
    cur现金 As Currency) As Boolean
    
    mintInsure = intInsure
    mlng病人ID = lng病人ID
    mcurPay = curPay
    mstrBalance = strBalance
    mcurError = curError
    mcurOneCard = CurOneCard
    mcur现金 = 0
    Set mrs结算方式 = rs结算方式
    
    Me.Show 1, frmParent
    
    If mblnOK Then
        curPay = mcurPay
        strBalance = mstrBalance
        curError = mcurError
        cur缴款 = mcur缴款
        cur找补 = mcur找补
        cur现金 = mcur现金
    End If
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long
    Dim str支票结算方式 As String
    Dim lngCashRow As Long
    
    If Val(txtOwe.Text) <> 0 Then
        If Val(txtOwe.Text) > 0 Then
            MsgBox "病人支付金额不足,请按所显示的差额补款。", vbExclamation, gstrSysName
            mshPay.SetFocus: Exit Sub
        Else
            MsgBox "病人支付金额过多,请按所显示的差额退款。", vbExclamation, gstrSysName
            mshPay.SetFocus: Exit Sub
        End If
    End If
    '刘兴洪:28947
    If mintInsure <> 0 Then
        If gclsInsure.CheckInsureValid(mintInsure) = False Then
            Exit Sub
        End If
    End If
    

            
    With mshPay
        mcur现金 = 0
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COLS.C1金额)) = 0 And i <> mlngPayRow Then
                If (.TextMatrix(i, COLS.C2号码) <> "" Or .TextMatrix(i, COLS.C3备注) <> "") Then
                    If MsgBox("注意:第" & i & "行没有输入金额,但输入了" & IIf(.TextMatrix(i, COLS.C2号码) <> "", "结算号", "备注") & _
                        vbCrLf & "!该信息不会保存!确定要继续吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        .SetFocus: .Row = i: .Col = COLS.C1金额: Exit Sub
                    End If
                End If
            ElseIf Val(.RowData(i)) = 7 Then
                j = j + 1
            End If
            If .RowData(i) = 1 Then
                mcur现金 = mcur现金 + Val(mshPay.TextMatrix(i, 1))
                lngCashRow = i
            End If
        Next
        
        If j > 1 Then
            MsgBox "不支持一次使用多种一卡通支付！", vbInformation
            Exit Sub
        End If
        '刘兴洪:35204,缴款金额控制
        Select Case gTy_Module_Para.byt缴款控制
        Case 1  '1-代表输入缴款后才结束病人累计
        Case 2  '2-收费时必须要输入缴款金额
            If Val(mcur现金) > 0 And Val(txt缴款.Text) = 0 Then
                MsgBox "注意:" & vbCrLf & _
                "    该病人未输入缴款金额,不能进行收费!", vbInformation + vbDefaultButton1, gstrSysName
                If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                Exit Sub
            End If
        Case Else   ',0-代表不进行缴款输入和累计控制
        End Select
        '37642
        If Val(mcur现金) < 0 Then
            If MsgBox("注意:" & vbCrLf & "   该病人的现金为负数了,你是否真的要退病人现金(" & Format(mcur现金, "####0.00;-####0.00;0.00;0.00") & ") ?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If lngCashRow > 0 Or lngCashRow < .Rows Then
                    .Row = lngCashRow: .Col = COLS.C1金额
                End If
                .SetFocus
                Exit Sub
            End If
        End If
        mcurPay = 0: mstrBalance = ""
        '33722
        str支票结算方式 = ""
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COLS.C1金额)) <> 0 Then
                '支票的结算方式,需要独立出来,主要是在分摊时,应该最后分摊支票,否则有问题
                mcurPay = mcurPay + Val(.TextMatrix(i, COLS.C1金额))
                If (.TextMatrix(i, COLS.C0方式) Like "*支票*" Or i = mlngPayRow) And mlngPayRow > 0 Then
                    str支票结算方式 = str支票结算方式 & "||" & .TextMatrix(i, COLS.C0方式) & "|" & .TextMatrix(i, COLS.C1金额) & _
                        "|" & IIf(.TextMatrix(i, COLS.C2号码) = "", " ", .TextMatrix(i, COLS.C2号码)) & _
                        "|" & IIf(.TextMatrix(i, COLS.C3备注) = "", " ", .TextMatrix(i, COLS.C3备注))
                Else
                    mstrBalance = mstrBalance & "||" & .TextMatrix(i, COLS.C0方式) & "|" & .TextMatrix(i, COLS.C1金额) & _
                        "|" & IIf(.TextMatrix(i, COLS.C2号码) = "", " ", .TextMatrix(i, COLS.C2号码)) & _
                        "|" & IIf(.TextMatrix(i, COLS.C3备注) = "", " ", .TextMatrix(i, COLS.C3备注))
                End If
                '空格填充以区分分隔符
            End If
        Next
    
        '支票最后分摊
        mstrBalance = mstrBalance & str支票结算方式
        mstrBalance = Mid(mstrBalance, 3)
        mcurPay = Format(mcurPay, "0.00")
        mcur缴款 = Val(txt缴款.Text)
        mcur找补 = Val(txt找补.Text)
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mcurOneCard > 0 Then
        If txt缴款.Enabled Then txt缴款.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If txtTmp.Visible Then
            txtTmp.Visible = False
            mshPay.SetFocus
        Else
            Call cmdCancel_Click
        End If
    Case vbKeyF2
        If cmdOK.Enabled And cmdOK.Visible Then
            Call cmdOK.SetFocus
            Call cmdOK_Click
        End If
    Case vbKeyF12
        If Shift = vbCtrlMask Then
            '强制性LED报价,(合计)
            If gblnLED Then
                mblnHotKey = True: txt缴款.SetFocus
                If ActiveControl Is txt缴款 Then txt缴款_GotFocus
            End If
        End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 _
        And Not ActiveControl Is mshPay _
        And Not ActiveControl Is txtTmp _
        And Not ActiveControl Is txt缴款 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("'|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String, i As Long
    Dim arrPay As Variant, j As Long
    
    mblnOK = False
    mintDefault = 0
    mcurError = 0
    
    txtPay.Text = Format(mcurPay, "0.00")
    arrPay = Array()
    If mstrBalance <> "" Then
        arrPay = Split(mstrBalance, "||")
    End If
    
    On Error GoTo errH
    mrs结算方式.Filter = "性质=1 or 性质=2 or 性质=7"
    
    With mshPay
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .Rows = mrs结算方式.RecordCount + 1
        For i = 1 To mrs结算方式.RecordCount
            .RowData(i) = NVL(mrs结算方式!性质, 1)
            .TextMatrix(i, COLS.C0方式) = mrs结算方式!名称
            
            '缺省结算方式(没有则用现金)
            If mrs结算方式!名称 = gstr结算方式 Then mintDefault = i
            If NVL(mrs结算方式!缺省, 0) = 1 And mintDefault = 0 Then mintDefault = i
            If NVL(mrs结算方式!性质, 1) = 1 And mintDefault = 0 Then mintDefault = i
            '缺省值(上一次的)
            For j = 0 To UBound(arrPay)
                If Split(arrPay(j), "|")(0) = mrs结算方式!名称 Then
                    .TextMatrix(i, COLS.C1金额) = Format(Split(arrPay(j), "|")(1), "0.00")
                    .TextMatrix(i, COLS.C2号码) = Trim(Split(arrPay(j), "|")(2))  '去掉人为的空格填充
                    .TextMatrix(i, COLS.C3备注) = Trim(Split(arrPay(j), "|")(3))
                    Exit For
                End If
            Next
            If Val(NVL(mrs结算方式!应付款)) = 1 Then
                mlngPayRow = i: mstr应付结算方式 = mrs结算方式!名称
                .RowHeight(i) = 0
            End If
            mrs结算方式.MoveNext
        Next
        If mintDefault > 0 Then .CellFontBold = True
        '设置应付款的缺省放在支票下的行和字体:33722
        j = -1
        For i = 1 To .Rows - 1
            If .RowData(i) = 2 And InStr(1, .TextMatrix(i, COLS.C0方式), "支票") > 0 And i <> mlngPayRow Then
                '获取最后的支票行
                j = i
            End If
        Next
        
        If j <> -1 And mlngPayRow > 0 Then
             If mlngPayRow <> j And j < .Rows - 1 Then
                '需要将应付款行放在支票后
                .RowPosition(mlngPayRow) = j + 1
                mlngPayRow = j + 1
             End If
        End If
        If mlngPayRow > 0 Then
            .Row = mlngPayRow
            For i = 0 To .COLS - 1
                .Col = i: .CellFontBold = True
            Next
        End If
        .Row = 1: .Col = COLS.C1金额
    End With
    Call ShowMoney(mstrBalance = "", False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mshPay_DblClick()
    With mshPay
        If Not txtTmp.Visible And mshPay.Row > 0 And mshPay.Col > COLS.C0方式 Then
            If mshPay.Row <> mlngPayRow Then
                Call SetTxtTmp
                txtTmp.Text = mshPay.TextMatrix(mshPay.Row, mshPay.Col)
                txtTmp.SelStart = 0: txtTmp.SelLength = Len(txtTmp.Text)
            End If
        End If
    End With
End Sub
Private Sub mshPay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call LocateMshpay
    ElseIf KeyCode = vbKeyDelete Then
        If mshPay.Row > 0 And mshPay.Col > COLS.C0方式 Then
            mshPay.TextMatrix(mshPay.Row, mshPay.Col) = ""
            If mshPay.Col = COLS.C1金额 Then Call ShowMoney(False, mshPay.Row <> mintDefault)
        End If
    End If
End Sub

Private Sub mshPay_KeyPress(KeyAscii As Integer)
    If Not txtTmp.Visible And mshPay.Row > 0 And mshPay.Col > COLS.C0方式 And KeyAscii <> 13 Then
        If mshPay.Col = COLS.C1金额 Then
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        If mshPay.Col <> COLS.C3备注 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If mshPay.Col = COLS.C2号码 Then
            If zlCommFun.IsCharChinese(Chr(KeyAscii)) Then Call Beep: Exit Sub
        End If
        If mshPay.Row <> mlngPayRow Then
            Call SetTxtTmp
            txtTmp.Text = Chr(KeyAscii)
            txtTmp.SelStart = 1
        End If
    End If
End Sub

Private Sub LocateMshpay()
    Dim j As Long
    Dim a As Long
    Dim lngRow As Long
    
    With mshPay
        lngRow = .Row
        If mlngPayRow > 0 And mlngPayRow = .Rows - 1 Then
            If lngRow = .Rows - 2 Then lngRow = .Rows - 1
        End If
        
        '非末行的最后一列换行,非末行金额为零换行
        If lngRow < .Rows - 1 And (.Col = .COLS - 1 Or .TextMatrix(.Row, COLS.C1金额) = "" And .Col <> COLS.C0方式) Then
            a = .Row
            For j = .Row + 1 To .Rows - 1
                If .RowHeight(j) > 0 Then
                    .Row = j: Exit For
                End If
            Next
            .Col = COLS.C1金额
            If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
            End If
            If a = j Then
                
            End If
            Call mshPay_EnterCell
        ElseIf lngRow = .Rows - 1 And (.Col = .COLS - 1 Or .TextMatrix(.Row, COLS.C1金额) = "" And .Col <> COLS.C0方式) Then
        '末行的最后一列,末行金额为零,Tab
            Call zlCommFun.PressKey(vbKeyTab)
        Else
             If .RowData(.Row) = 1 And .Col = COLS.C1金额 Then '现金无需输结算号码
                .Col = .Col + 2
            Else
                .Col = .Col + 1
            End If
            Call mshPay_EnterCell
        End If
    End With
End Sub

Private Sub SetTxtTmp()
    With txtTmp
        .MaxLength = Val("" & Choose(mshPay.Col, 10, 30, 25))   '摘要最长50位,所以限制最大只能25个汉字
        .Left = mshPay.Left + mshPay.CellLeft + 15
        .Top = mshPay.Top + mshPay.CellTop + (mshPay.CellHeight - txtTmp.Height) / 2 - 15
        .Width = mshPay.CellWidth - 60
        .ForeColor = mshPay.CellForeColor
        .BackColor = mshPay.CellBackColor
        .Alignment = IIf(mshPay.Col = COLS.C1金额, 1, 0)
        .ZOrder: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub mshPay_LeaveCell()
    txtTmp.Visible = False
End Sub

Private Sub mshPay_Scroll()
    txtTmp.Visible = False
End Sub


Private Sub txtTmp_KeyPress(KeyAscii As Integer)
    Dim blnCent As Boolean, i As Long
    
    If KeyAscii <> 13 Then
        If mshPay.Col = COLS.C1金额 Then
            If InStr(txtTmp.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        ElseIf mshPay.Col = COLS.C2号码 Then  '结算号码特殊字符限制
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else    '备注
            If InStr("'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If mshPay.Col = COLS.C1金额 Then
            If Not IsNumeric(txtTmp.Text) And txtTmp.Text <> "" Then
                MsgBox "请输入正确的数值。", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txtTmp): Exit Sub
            End If
            If Val(txtTmp.Text) > 100 Then
                If Val(txtTmp.Text) > mcurPay * 2 Then
                    If MsgBox("输入的数字超过了付款合计的两倍(" & Format(mcurPay * 2, "0.00") & ")，你确认要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        Call zlControl.TxtSelAll(txtTmp)
                        Exit Sub
                    End If
                End If
            End If
            
            txtTmp.Text = Format(Val(txtTmp.Text), "0.00")
            If Val(txtTmp.Text) <> 0 Then
                If Val(mshPay.RowData(mshPay.Row)) = 1 Then
                    '如果是在现金栏内输入,则进行分币处理
                    blnCent = True
                    If gBytMoney = 0 Then blnCent = False
                    If blnCent And mintInsure <> 0 And mlng病人ID <> 0 Then
                        If gclsInsure.GetCapability(support门诊预算, mlng病人ID, mintInsure) Then
                            If Not gclsInsure.GetCapability(support分币处理, mlng病人ID, mintInsure) Then
                                blnCent = False
                            End If
                        End If
                    End If
                    If blnCent Then
                        txtTmp.Text = Format(CentMoney(Val(txtTmp.Text)), "0.00")
                    End If
                ElseIf Val(mshPay.RowData(mshPay.Row)) = 7 And mcurOneCard > 0 Then '一卡通
                    If Val(txtTmp.Text) > mcurOneCard Then
                        txtTmp.Text = Format(mcurOneCard, "0.00")
                    End If
                End If
            End If
        
            If Val(txtTmp.Text) = 0 Then
                mshPay.TextMatrix(mshPay.Row, mshPay.Col) = ""
            Else
                mshPay.TextMatrix(mshPay.Row, mshPay.Col) = Format(Val(txtTmp.Text), "0.00")
            End If
            
            '差额及误差计算
            Call ShowMoney(False, mshPay.Row <> mintDefault)
        Else
        '防字符拷贝
            If mshPay.Col = COLS.C2号码 Then
                If InStr(txtTmp.Text, ",") > 0 Then Call Beep: Exit Sub
                If zlCommFun.IsCharChinese(txtTmp.Text) Then Call Beep: Exit Sub
            End If
            If InStr(txtTmp.Text, "'") > 0 Or InStr(txtTmp.Text, "|") > 0 Then Call Beep: Exit Sub
            
            mshPay.TextMatrix(mshPay.Row, mshPay.Col) = txtTmp.Text
        End If
        mshPay.SetFocus
        txtTmp.Visible = False
        '换行与换列
        Call LocateMshpay
    End If
End Sub

Private Sub txtTmp_GotFocus()
    If mshPay.Col = COLS.C3备注 Then
        txtTmp.IMEMode = 1
        zlCommFun.OpenIme True
    Else
        txtTmp.IMEMode = 3
    End If
End Sub

Private Sub txtTmp_LostFocus()
    txtTmp.Visible = False
    If mshPay.Col = COLS.C3备注 Then
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub mshPay_EnterCell()
    If mshPay.Col = COLS.C3备注 Then
       zlCommFun.OpenIme True
       Exit Sub
    End If
    zlCommFun.OpenIme False
End Sub

Private Sub txtTmp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtTmp.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtTmp.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtTmp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Call SetWindowLong(txtTmp.hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtTmp_Validate(Cancel As Boolean)
    txtTmp.Visible = False
End Sub

Private Sub txtPay_GotFocus()
    Call zlControl.TxtSelAll(txtPay)
End Sub

Private Sub txt缴款_Change()
    Dim cur现金 As Currency, i As Long
    
    For i = 1 To mshPay.Rows - 1
        If mshPay.RowData(i) = 1 Then
            cur现金 = Val(mshPay.TextMatrix(i, 1))
            Exit For
        End If
    Next
    If Val(txt缴款.Text) = 0 Then txt找补.Text = "0.00": Exit Sub
    txt找补.Text = Format(Val(txt缴款.Text) - cur现金, "0.00")
End Sub

Private Sub txt缴款_GotFocus()
    Dim cur现金 As Currency
    Dim i As Long
    '35204
    Call zlControl.TxtSelAll(txt缴款)
    cur现金 = 0
    With mshPay
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                cur现金 = cur现金 + Val(mshPay.TextMatrix(i, 1))
            End If
        Next
    End With
    'LED显示:应缴金额
     If gblnLED And cur现金 <> 0 Then
        '自动报价或手工报价时由热键激活
        If (Not gbln手工报价 And ActiveControl Is txt缴款) Or (gbln手工报价 And mblnHotKey) Then
            mblnHotKey = False
            zl9LedVoice.Speak "#21 " & cur现金
        End If
    End If
    
End Sub
Private Function get现金() As Currency
    Dim i As Long, cur现金 As Double
    cur现金 = 0
    With mshPay
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                cur现金 = cur现金 + Val(mshPay.TextMatrix(i, 1))
            End If
        Next
    End With
    get现金 = cur现金
End Function

Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    Dim cur现金 As Currency
    If KeyAscii = 13 Then
        If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
        If txt缴款.Text <> "0.00" Then
            If Val(txt找补.Text) >= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt缴款): txt缴款.SetFocus
                Exit Sub
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab) '病人累加缴款
        End If
        
        'LED显示
        cur现金 = get现金
        If gblnLED Then
            mblnHotKey = False
            Call zl9LedVoice.DisplayBank( _
                "合计:" & mcurPay & "元,应付:" & cur现金 & "元", _
                "收您:" & txt缴款.Text & "元" & IIf(Val(txt找补.Text) = 0, "", ",找您:" & txt找补.Text & "元"))
            zl9LedVoice.Speak "#22 " & txt缴款.Text
            zl9LedVoice.Speak "#23 " & txt找补.Text
            zl9LedVoice.Speak "#3"
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

Private Sub txt找补_GotFocus()
    Call zlControl.TxtSelAll(txt找补)
End Sub

Private Sub ShowMoney(blnFirst As Boolean, blnAutoCalc As Boolean, Optional bln支票 As Boolean = False)
'功能：设置和显示界面的各种金额
'参数：blnFirst=第一次调用时自动设置缺省结算方式及金额
'      blnAutoCalc=是否根据差额自动补平缺省结算方式
    Dim curPay As Currency, curOwn As Currency
    Dim blnCent As Boolean, i As Long, blnSet As Boolean
    
    txt缴款.Text = "0.00"
    '判断是否应该进行分币处理
    blnCent = True
    If gBytMoney = 0 Then blnCent = False
    If blnCent And mintInsure <> 0 And mlng病人ID <> 0 Then
        If gclsInsure.GetCapability(support门诊预算, mlng病人ID, mintInsure) Then
            If Not gclsInsure.GetCapability(support分币处理, mlng病人ID, mintInsure) Then
                blnCent = False
            End If
        End If
    End If
    
    '第一次调用时自动设置缺省结算方式及金额
    '-----------------------------------------------------------------------------------------------------
    If blnFirst Then
        If mcurOneCard > 0 Then
            For i = 1 To mshPay.Rows - 1
                If mshPay.RowData(i) = 7 Then
                    mshPay.TextMatrix(i, COLS.C1金额) = Format(mcurOneCard, "0.00")
                    blnSet = True
                End If
            Next
        End If
        
        If mintDefault > 0 Then
            If mshPay.RowData(mintDefault) = 1 And blnCent Then '现金时要进行分币处理
                mshPay.TextMatrix(mintDefault, COLS.C1金额) = Format(CentMoney(mcurPay - IIf(blnSet, mcurOneCard, 0)), "0.00")
            Else
                mshPay.TextMatrix(mintDefault, COLS.C1金额) = Format(mcurPay - IIf(blnSet, mcurOneCard, 0), "0.00")
            End If
        End If
    End If
    
    Call Calc退支票
    '显示缴款差额
    '-----------------------------------------------------------------------------------------------------
    curPay = 0
    For i = 1 To mshPay.Rows - 1
        curPay = curPay + Val(mshPay.TextMatrix(i, 1))
    Next
    curOwn = mcurPay - curPay
    txtOwe.Text = Format(mcurPay - curPay, "0.00") '这里是差额,不一定用现金,所以不处理分币
    
    '根据差额自动补平并计算
    '-----------------------------------------------------------------------------------------------------
    If blnAutoCalc And Val(txtOwe.Text) <> 0 Then
        '剩余部份尝试设置到缺省结算方式上
        If mlngPayRow >= 0 And bln支票 Then
             mshPay.TextMatrix(mlngPayRow, 1) = Format(Val(mshPay.TextMatrix(mlngPayRow, 1)) + curOwn, "0.00")
             If Val(mshPay.TextMatrix(mlngPayRow, 1)) <> 0 Then
                mshPay.RowHeight(mlngPayRow) = mshPay.RowHeight(0)
             Else
                mshPay.RowHeight(mlngPayRow) = 0
             End If
        Else
            If mintDefault > 0 Then
                If mshPay.RowData(mintDefault) = 1 And blnCent Then '现金时要进行分币处理
                    mshPay.TextMatrix(mintDefault, 1) = _
                        Format(CentMoney(Val(mshPay.TextMatrix(mintDefault, 1)) + curOwn), "0.00")
                Else
                    mshPay.TextMatrix(mintDefault, 1) = _
                        Format(Val(mshPay.TextMatrix(mintDefault, 1)) + curOwn, "0.00")
                End If
                If Val(mshPay.TextMatrix(mintDefault, 1)) = 0 Then
                    mshPay.TextMatrix(mintDefault, 1) = ""
                End If
                txtOwe.Text = "0.00"
            End If
        End If
    End If
    
    '计算误差金额(结算金额-结帐金额)
    '-----------------------------------------------------------------------------------------------------
    curPay = 0
    For i = 1 To mshPay.Rows - 1
        curPay = curPay + Val(mshPay.TextMatrix(i, 1))
    Next
    mcurError = Format(curPay - mcurPay, gstrDec)
    
    '有可能缴款差额正好是处理分币的误差部份,就不显示了(三七作五二舍八入时最大可能有0.29的误差,0.79作0.5,0.29作0)
    If Val(txtOwe.Text) <> 0 And (Abs(Val(Val(txtOwe.Text))) < 0.1 Or gBytMoney = 5 And Abs(Val(Val(txtOwe.Text))) < 0.3) And mintDefault > 0 Then
        If mshPay.RowData(mintDefault) = 1 And blnCent Then
            If CentMoney(Val(mshPay.TextMatrix(mintDefault, 1)) + Val(txtOwe.Text)) = Val(mshPay.TextMatrix(mintDefault, 1)) Then
                txtOwe.Text = "0.00"
            End If
        End If
    End If
    
    '可能缴款差额部份是小数点的正常误差部份,如果四舍五入小于1分,就不补了
    If Val(txtOwe.Text) <> 0 And mcurError + curOwn = 0 And curOwn < 0.005 And curOwn >= -0.005 Then
        txtOwe.Text = "0.00"
    End If
End Sub
Private Sub Calc退支票()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动计算退支票
    '编制:刘兴洪
    '日期:2010-11-08 14:37:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl支票额 As Double, dbl现金 As Double
    '33722
    With mshPay
        '金额栏,又是支票栏的,修正好要重算退支票
     '  If Not (.Col = COLS.C1金额 And InStr(1, .TextMatrix(.Row, COLS.C0方式), "支票") > 0 _
            And .RowData(mshPay.Row) = 2 And .Row <> mlngPayRow) Then Exit Sub
        If mlngPayRow <= 0 Then Exit Sub
        dbl支票额 = 0: dbl现金 = 0
        For i = 1 To .Rows - 1
             If InStr(1, .TextMatrix(i, COLS.C0方式), "支票") > 0 _
                And .RowData(i) = 2 And i <> mlngPayRow Then
                dbl支票额 = dbl支票额 + Val(.TextMatrix(i, COLS.C1金额))
             ElseIf i <> mintDefault And i <> mlngPayRow Then
                    dbl现金 = dbl现金 + Val(.TextMatrix(i, COLS.C1金额))
            End If
        Next
        If RoundEx(mcurPay - dbl现金 - dbl支票额, 2) >= 0 Or dbl支票额 = 0 Then
            .TextMatrix(mlngPayRow, COLS.C1金额) = "": .RowHeight(mlngPayRow) = 0
        Else
            .TextMatrix(mlngPayRow, COLS.C1金额) = Format(RoundEx(mcurPay - dbl现金 - dbl支票额, 2), "0.00"): .RowHeight(mlngPayRow) = .RowHeight(0)
        End If
    End With
End Sub
