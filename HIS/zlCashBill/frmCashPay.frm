VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmCashPay 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增缴款记录"
   ClientHeight    =   6180
   ClientLeft      =   435
   ClientTop       =   720
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCashPay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cbo缴款部门 
      Height          =   360
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4800
      Width           =   2010
   End
   Begin VB.TextBox txtPay 
      Enabled         =   0   'False
      Height          =   360
      Left            =   930
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1245
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   5070
      TabIndex        =   14
      Top             =   5520
      Width           =   1200
   End
   Begin ZL9BillEdit.BillEdit msh 
      Height          =   2205
      Left            =   930
      TabIndex        =   1
      Top             =   1410
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   3889
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   360
      RowHeightMin    =   360
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   930
      TabIndex        =   18
      Top             =   915
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
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
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   190906371
      CurrentDate     =   36904
   End
   Begin VB.TextBox txtSum 
      Enabled         =   0   'False
      Height          =   360
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3780
      Width           =   5325
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   -210
      TabIndex        =   11
      Top             =   5250
      Width           =   7125
   End
   Begin VB.TextBox txtDigest 
      Height          =   360
      Left            =   930
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4290
      Width           =   5325
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印设置(&S)"
      Height          =   420
      Left            =   210
      TabIndex        =   15
      Top             =   5520
      Width           =   1530
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   3720
      TabIndex        =   13
      Top             =   5520
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   2370
      TabIndex        =   12
      Top             =   5520
      Width           =   1200
   End
   Begin VB.TextBox txtHandle 
      Enabled         =   0   'False
      Height          =   360
      Left            =   5130
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1125
   End
   Begin VB.Label lblSum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "合计"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   390
      TabIndex        =   2
      Top             =   3840
      Width           =   480
   End
   Begin VB.Label lblDigest 
      Caption         =   "金额"
      Height          =   240
      Left            =   390
      TabIndex        =   0
      Top             =   1470
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "摘要"
      Height          =   240
      Left            =   390
      TabIndex        =   4
      Top             =   4350
      Width           =   480
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "日期"
      Height          =   240
      Left            =   390
      TabIndex        =   17
      Top             =   975
      Width           =   480
   End
   Begin VB.Label lblHandle 
      BackStyle       =   0  'Transparent
      Caption         =   "经手人"
      Height          =   240
      Left            =   4320
      TabIndex        =   9
      Top             =   4860
      Width           =   720
   End
   Begin VB.Label lblPay 
      BackStyle       =   0  'Transparent
      Caption         =   "缴款人"
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   4860
      Width           =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缴款登记卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2190
      TabIndex        =   16
      Top             =   240
      Width           =   2250
   End
End
Attribute VB_Name = "frmCashPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim colMax As Collection         '存入该结算方式的最大暂存金额

Dim mblnChange As Boolean     '为真时表示已改变了
Dim mdatCurrnet As Date
Dim mblnSuccess As Boolean

Private Sub InitTable()
    dtpDate.Value = mdatCurrnet
    dtpDate.MaxDate = mdatCurrnet
    
    With msh
        .Font.Size = 12
        .CboFont.Size = 12
        .TxtEditFont.Size = 12
        .Cols = 3
        .TextMatrix(0, 0) = "结算方式"
        .TextMatrix(0, 1) = "金额"
        .TextMatrix(0, 2) = "结算号"
        .ColWidth(0) = 1350
        .ColWidth(1) = 2475
        .ColWidth(2) = 1350
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        
        .ColData(0) = 3
        .ColData(1) = 4
        .ColData(2) = 4
        .PrimaryCol = 0
        .Active = True
    End With
    
    '初始化票据打印
    'On Error Resume Next
    'BillInit gcnOracle
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If ValidateContent() = False Then Exit Sub
    If Save() = False Then Exit Sub
    mblnChange = False
    mblnSuccess = True
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then msh.SetFocus
End Sub

Private Sub msh_EnterCell(Row As Long, Col As Long)
    Call ShowSum
End Sub

Private Sub msh_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub msh_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If msh.TxtVisible = False Then
        If msh.Col = 1 Then
            If msh.TextMatrix(msh.Row, 1) = "" And msh.TextMatrix(msh.Row, 2) = "" Then
                txtDigest.SetFocus
            End If
        ElseIf msh.Col = 2 Then
            If msh.TextMatrix(msh.Row, 2) = "" Then msh.TextMatrix(msh.Row, 2) = " "
        End If
        Exit Sub
    End If
    '对输入值进行验证
    msh.Text = Trim(msh.Text)
    If msh.Col = 1 Then
        If Not IsNumeric(msh.Text) Then
            MsgBox "金额输入错误。", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        If Val(msh.Text) > 99999999 Then
            MsgBox "金额过大。", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        If Val(msh.Text) < -999999999 Then
            MsgBox "金额过小。", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        msh.Text = Format(Val(msh.Text), "###########0.00;-###########0.00;0.00;0.00")
    Else
        If LenB(StrConv(msh.Text, vbFromUnicode)) > 10 Then
            MsgBox "结算号的长度不能超过10位。", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        If InStr(msh.Text, "'") > 0 Then
            MsgBox "结算号含有非法字符。", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        If msh.Text = "" Then msh.Text = " "
    End If
    mblnChange = True
End Sub

Private Sub txtDigest_Change()
    mblnChange = True
End Sub

Private Sub txtDigest_GotFocus()
    zlControl.TxtSelAll txtDigest
    zlCommFun.OpenIme True
End Sub

Private Sub txtDigest_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Function ValidateContent() As Boolean
'功能:检查输入内容的是否有效
'返回:有效则返回True,否则返回False

    Dim intTemp As Integer
    Dim intTempSub As Integer
    Dim douSum As Double
    
    Dim strJudged As String
    
    ValidateContent = False
    If LenB(StrConv(txtDigest.Text, vbFromUnicode)) > 50 Then
        MsgBox "摘要的长度不能超过25个汉字或50个字母。", vbInformation, gstrSysName
        zlControl.TxtSelAll txtDigest
        txtDigest.SetFocus
        Exit Function
    End If
    If InStr(txtDigest.Text, "'") > 0 Then
        MsgBox "摘要含有非法字符（'）。", vbInformation, gstrSysName
        zlControl.TxtSelAll txtDigest
        txtDigest.SetFocus
        Exit Function
    End If
    For intTemp = 1 To msh.Rows - 1
        douSum = 0
        If msh.TextMatrix(intTemp, 0) <> "" And msh.TextMatrix(intTemp, 1) <> "" Then
            '这种结算方式在前面判断了没？
            If InStr(strJudged, msh.TextMatrix(intTemp, 0) & ",") = 0 Then
                '统计出该种结算方式的总金额
                For intTempSub = intTemp To msh.Rows - 1
                    If msh.TextMatrix(intTempSub, 0) = msh.TextMatrix(intTemp, 0) Then douSum = douSum + Val(msh.TextMatrix(intTempSub, 1))
                Next
                
                If douSum > colMax(msh.TextMatrix(intTemp, 0)) Then     '该项金额大于暂存金额
                     If MsgBox(msh.TextMatrix(intTemp, 0) & "的缴款金额大于暂存金额，是否继续？", vbYesNo Or vbQuestion Or vbDefaultButton2, Me.Caption) = vbNo Then
                        msh.Row = intTemp
                        msh.Col = 1
                        msh.SetFocus
                        msh.TxtSetFocus
                        Exit Function
                     End If
                End If
            strJudged = strJudged & msh.TextMatrix(intTemp, 0) & ","
            End If
        End If
    Next
    ValidateContent = True
End Function

Private Function Save() As Boolean
'功能:保存编辑的内容
'参数:
'返回值:成功返回True,否则为False
    Dim intTemp As Integer
    Dim strTemp As String
    Dim lngID As Long, lng单据 As Long
    
    On Error GoTo errHandle
    Save = False
    gcnOracle.BeginTrans
    With msh
        lng单据 = zlDatabase.GetNextId("人员缴款记录")
        For intTemp = 1 To .Rows - 1
            If lngID = 0 Then
                lngID = lng单据
            Else
                lngID = zlDatabase.GetNextId("人员缴款记录")
            End If
            
            If .TextMatrix(intTemp, 0) <> "" And .TextMatrix(intTemp, 1) <> "" Then
                gstrSQL = "zl_人员缴款记录_insert(" & lngID & "," & lng单据 & _
                    ",to_date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),'" & txtPay.Text & "','" & txtHandle.Text & _
                    "','" & .TextMatrix(intTemp, 0) & "'," & .TextMatrix(intTemp, 1) & ",'" & .TextMatrix(intTemp, 2) & _
                    "','" & txtDigest.Text & "',Null," & cbo缴款部门.ItemData(cbo缴款部门.ListIndex) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    gcnOracle.CommitTrans
    
    '打印票据
    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me, "单据ID=" & lng单据, 2)  '2表示直接打印
    
    Save = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowSum()
'功能:计算缴款总额
    
    Dim dblTemp As Double
    Dim intTemp As Integer
    
    For intTemp = 1 To msh.Rows - 1
         dblTemp = dblTemp + Val(msh.TextMatrix(intTemp, 1))
    Next
    txtSum.Text = Format(dblTemp, "######0.00;-######0.00;0;") & "元" & IIf(dblTemp = 0, "", " （" & zlCommFun.UppeMoney(dblTemp) & "）")
End Sub

Public Function 编辑缴款记录(ByVal str缴款人 As String, ByVal lng缴款人ID As Long) As Boolean
'功能:用来与调用的财务监控窗口进行通讯的程序,用来增加缴款记录
'参数:str缴款人     缴款人的名字
'返回值:编辑成功返回True,否则为False
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intRow As Integer
    
    On Error GoTo errHandle
    
    mblnSuccess = False
    If UserInfo.姓名 = "" Then
        MsgBox "当前登录用户未指定对应的人员，不能使用本功能。", vbExclamation, gstrSysName
        Set frmCashPay = Nothing
        Exit Function
    End If
    
    mdatCurrnet = Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    
    txtPay.Text = str缴款人
    txtHandle.Text = UserInfo.姓名
    Set rsTmp = GetPersonnelDept(lng缴款人ID)
    Call zlControl.CboAddData(cbo缴款部门, rsTmp, True)
    If cbo缴款部门.ListCount > 0 Then cbo缴款部门.ListIndex = 0
    
    gstrSQL = "Select 结算方式,余额 " & _
            " From 人员缴款余额 Where 收款员 =[1] and 性质=1 and 余额<>0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str缴款人)
    
    If rsTmp.RecordCount = 0 Then
        MsgBox "收费员" & str缴款人 & "没有暂存金额，无须进行缴款操作。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Call InitTable
    msh.Clear
    msh.Rows = rsTmp.RecordCount + 1
    intRow = 1
    Set colMax = New Collection
    Do Until rsTmp.EOF
        msh.TextMatrix(intRow, 0) = rsTmp("结算方式")
        msh.TextMatrix(intRow, 1) = Format(rsTmp("余额"), "###########0.00;-###########0.00;0.00;0.00")
        msh.TextMatrix(intRow, 2) = " "
        '保存最大值
        colMax.Add CDbl(rsTmp("余额")), CStr(rsTmp("结算方式"))
        msh.AddItem rsTmp("结算方式")
        intRow = intRow + 1
        rsTmp.MoveNext
    Loop
    
    mblnChange = False
    frmCashPay.Show vbModal, frmCashSupervise
    编辑缴款记录 = mblnSuccess
    Exit Function
errHandle:
    MsgBox "数据读出失败。", vbExclamation, gstrSysName
    编辑缴款记录 = False
End Function

Private Sub txtDigest_LostFocus()
    zlCommFun.OpenIme False
End Sub
