VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm付款计划 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "付款计划管理"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "0"
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5760
      Left            =   6510
      TabIndex        =   18
      Top             =   495
      Width           =   60
   End
   Begin ZL9BillEdit.BillEdit mshList 
      Height          =   2760
      Left            =   75
      TabIndex        =   0
      Top             =   2595
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   4868
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   5
      RowHeight0      =   315
      RowHeightMin    =   315
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6720
      TabIndex        =   4
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6720
      TabIndex        =   3
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "重置(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6720
      TabIndex        =   2
      Top             =   1530
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6720
      TabIndex        =   1
      Top             =   555
      Width           =   1100
   End
   Begin VB.Frame fraTemp 
      Caption         =   "应付款信息"
      Height          =   1515
      Left            =   75
      TabIndex        =   5
      Top             =   660
      Width           =   6300
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3570
         TabIndex        =   16
         Top             =   1020
         Width           =   1470
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   10
         Top             =   1020
         Width           =   1470
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1050
         TabIndex        =   9
         Top             =   630
         Width           =   5100
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4695
         TabIndex        =   7
         Top             =   255
         Width           =   1455
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1050
         TabIndex        =   6
         Top             =   255
         Width           =   2955
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "未计划金额"
         Height          =   180
         Index           =   4
         Left            =   2610
         TabIndex        =   17
         Top             =   1110
         Width           =   900
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "应付金额"
         Height          =   180
         Index           =   3
         Left            =   285
         TabIndex        =   13
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "发票号"
         Height          =   180
         Index           =   1
         Left            =   4110
         TabIndex        =   12
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "应付款序号"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   11
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "摘    要"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   8
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -45
      TabIndex        =   19
      Top             =   495
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frm付款计划.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   "根据应付记录单，制定经后付款计划。"
      Height          =   165
      Left            =   645
      TabIndex        =   20
      Top             =   330
      Width           =   5655
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      Caption         =   "应付款分期支付计划"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   5
      Left            =   2295
      TabIndex        =   15
      Top             =   2370
      Width           =   1620
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   75
      TabIndex        =   14
      Top             =   2295
      Width           =   6300
   End
End
Attribute VB_Name = "frm付款计划"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlngID As Long, mstrDate As String
Dim mblnOK As Boolean
Dim mblnChange As Boolean
Private Enum HeadCol
        计划支付日期 = 0
        支付金额
        执行
        计划人
        制定日期
End Enum

Public Sub 计划(ByVal FrmMain As Object, lngID As Long, Optional ByRef blnSussces As Boolean)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:制定计划
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------

    mlngID = lngID
    Call initCard
    
    Me.Show vbModal, FrmMain
    blnSussces = mblnOK
End Sub
Private Sub initCard()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始计划
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rslist As New ADODB.Recordset, lngRow As Long, strSQL As String
    Dim lngID As Long
    'by lesfeng 2009-12-2 性能优化
    strSQL = "" & _
            "   Select 单位ID,max(decode(记录状态,1,ID,3,iD,0)) as ID, " & _
            "       max(发票号) 发票号,max(摘要) 摘要,sum(nvl(发票金额,0)) 发票金额 " & _
            "   From 应付记录  a  " & _
            "   Where  exists  (Select 记录性质,NO,nvl(项目id,0),nvl(序号,0) From 应付记录  " & _
            "                   where 记录性质=a.记录性质 and no=a.no and nvl(项目id,0)=nvl(a.项目id,0) and " & _
            "                           nvl(序号,0)=nvl(a.序号,0) and nvl(系统标识,0)=nvl(a.系统标识,0) and ID=[1]" & _
            "                           And 记录性质<>-1 and 审核人 is not null )" & _
            "   Group by 单位ID,记录性质,NO,项目id,序号"
    
    
    Err = 0
    On Error GoTo ErrHand:
    Set rslist = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If rslist.EOF Then
        MsgBox "该计划未找到,可能已经被他人删除!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    lngID = Val(Nvl(rslist!ID))
    txtInfo(0) = IIf(IsNull(rslist!ID), "", rslist!ID)
    txtInfo(1) = IIf(IsNull(rslist!发票号), "", rslist!发票号)
    txtInfo(2) = IIf(IsNull(rslist!摘要), "", rslist!摘要)
    txtInfo(3) = IIf(IsNull(rslist!发票金额), "0.00", Format(rslist!发票金额, "###0.00;-###0.00;0.00;0.00"))
    txtInfo(4) = IIf(IsNull(rslist!发票金额), "0.00", Format(rslist!发票金额, "###0.00;-###0.00;0.00;0.00"))
    txtInfo(0).Tag = rslist!单位ID
    
    Call initGrid
    
    strSQL = "" & _
        "   Select 计划日期,计划金额,Decode(付款序号,Null,' ','√') As 执行,计划人,制定日期 " & _
        "   From 应付记录 " & _
        "   Where ID=[1] And 记录性质=-1 " & _
        "   Order By 计划序号"
    Set rslist = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    Dim dblTmp As Double
    
    With mshList
        .ClearBill
        .Rows = rslist.RecordCount + 2
        lngRow = 1
        While Not rslist.EOF
            .TextMatrix(lngRow, 0) = Format(rslist(0), "yyyy-MM-dd")
            .TextMatrix(lngRow, 1) = Format(rslist(1), "0.00")
            .TextMatrix(lngRow, 2) = Nvl(rslist(2))
            .TextMatrix(lngRow, 3) = Nvl(rslist(3))
            .TextMatrix(lngRow, 4) = Format(rslist(4), "yyyy-MM-dd")
            dblTmp = dblTmp + Nvl(rslist(1), 0)
            rslist.MoveNext
            lngRow = lngRow + 1
        Wend
        
        .Value = Format(zldatabase.Currentdate, "yyyy-mm-dd")
        
    End With
    txtInfo(4) = Format(Val(txtInfo(4)) - dblTmp, "###0.00;-###0.00;0;0")
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub
Private Sub initGrid()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始化网格
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    With mshList
        .Cols = 5
        .TextMatrix(0, HeadCol.计划支付日期) = "计划支付日期"
        .TextMatrix(0, HeadCol.支付金额) = "支付金额"
        .TextMatrix(0, HeadCol.执行) = "执行"
        .TextMatrix(0, HeadCol.计划人) = "计划人"
        .TextMatrix(0, HeadCol.制定日期) = "制定日期"
        
        .ColAlignment(HeadCol.计划支付日期) = flexAlignCenterCenter
        .ColAlignment(HeadCol.支付金额) = flexAlignRightCenter
        .ColAlignment(HeadCol.执行) = flexAlignCenterCenter
        .ColAlignment(HeadCol.计划人) = flexAlignLeftCenter
        .ColAlignment(HeadCol.制定日期) = flexAlignCenterCenter
                
        
        .ColWidth(HeadCol.计划支付日期) = 1600
        .ColWidth(HeadCol.支付金额) = 1200
        .ColWidth(HeadCol.执行) = 700
        .ColWidth(HeadCol.计划人) = 1000
        .ColWidth(HeadCol.制定日期) = 1600
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
        .ColData(HeadCol.计划支付日期) = 2
        .ColData(HeadCol.支付金额) = 4
        .ColData(HeadCol.计划人) = 5
        .ColData(HeadCol.制定日期) = 5
        .ColData(HeadCol.执行) = 5
        .CmdVisible = True
        .PrimaryCol = HeadCol.计划支付日期
        .LocateCol = HeadCol.计划支付日期
        .Active = True
                                
    End With
    
End Sub
Private Sub cmdExit_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdReset_Click()
    Call initCard
    cmdSave.Enabled = False
    cmdReset.Enabled = False
    mblnChange = False
    mshList.SetFocus
End Sub

Private Sub cmdSave_Click()
    '验证数据
    If IsValid = False Then Exit Sub
        
    '保存数据
    If Save() Then
        cmdSave.Enabled = False
        cmdReset.Enabled = False
        mblnChange = False
    End If
    mblnOK = True
    mshList.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And mshList.Col = 4 Then
        Me.Tag = 1
    Else
        Me.Tag = 0
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnSaveFlag As Integer
    Dim blnYes As Boolean
    If mblnChange = False Then
        Exit Sub
    End If
    
    ShowMsgbox "你已经更改了信息,你这样退出的话," & vbCrLf & "所更改的数据将不能保存,真的要退出吗?", True, blnYes
    If blnYes = True Then
        Exit Sub
    End If
    Cancel = 1
    mshList.SetFocus
End Sub
Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long, strSQL As String
    Dim dblTmp As Double
    Dim lngRow As Long
    Dim strTmp As String
    If Val(txtInfo(4)) < 0 Then
        ShowMsgbox "计划付款金额超过了应付款金额，请修改"
        IsValid = False
        Exit Function
    End If
    With mshList
        For lngLoop = 1 To mshList.Rows - 1
            If Trim(mshList.TextMatrix(lngLoop, HeadCol.计划支付日期)) = "" Then Exit For
            
            Err = 0
            On Error Resume Next
            strSQL = Format(CDate(mshList.TextMatrix(lngLoop, 0)), "yyyy-MM-dd")
            If Err.Number <> 0 Then
                ShowMsgbox "日期格式错误，请修改。"
                .Row = lngLoop
                .Col = HeadCol.计划支付日期
                .SetFocus
                Exit Function
            End If
            Err.Clear
            
            dblTmp = Val(mshList.TextMatrix(lngLoop, HeadCol.支付金额))
            If dblTmp = 0 Then
                ShowMsgbox "支付金额必需大于零!"
                .Row = lngLoop
                .Col = HeadCol.支付金额
                .SetFocus
                Exit Function
            End If
            If CDate(mshList.TextMatrix(lngLoop, HeadCol.制定日期)) > CDate(mshList.TextMatrix(lngLoop, HeadCol.计划支付日期)) Then
                If MsgBox("计划付款日期小于制定计划日期，是否忽略？", vbYesNo + vbDefaultButton2 + vbQuestion, Me.Caption) <> vbYes Then
                    mshList.Row = lngLoop
                    mshList.Col = HeadCol.计划支付日期
                    IsValid = False
                    Exit Function
                End If
            End If
            strTmp = Trim(mshList.TextMatrix(lngLoop, HeadCol.计划支付日期))
            For lngRow = lngLoop + 1 To mshList.Rows - 1
                If strTmp = Trim(mshList.TextMatrix(lngRow, HeadCol.计划支付日期)) Then
                    ShowMsgbox "第" & lngLoop & "行与第" & lngRow & "行的计划支付日期相同了," & vbCrLf & "请合并该计划!"
                    .Row = lngLoop
                    .Col = HeadCol.计划支付日期
                    .SetFocus
                    Exit Function
                End If
            Next
        Next
    End With
    IsValid = True
End Function
Private Function Save() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存数据
    '--入参数:
    '--出参数:
    '--返  回:成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------


    Dim lngLoop As Long, strSQL As String, lngNewDate As Boolean
    Dim str计划人 As String
    Dim str制定日期 As String
    Dim lngSN As Long
    Dim rsTmp As ADODB.Recordset
    
    Save = False
    
    On Error GoTo ErrHand:
        
    gcnOracle.BeginTrans
    
    strSQL = "ZL_付款计划_DELETE (" & mlngID & ")"  '已付款的部分，过程中未删除
    
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    
    With mshList
        For lngLoop = 1 To mshList.Rows - 1
            If Trim(.TextMatrix(lngLoop, HeadCol.计划支付日期)) <> "" And .TextMatrix(lngLoop, HeadCol.执行) <> "√" Then
                '过程参数
                '   ID_IN,计划序号_IN,计划金额_IN,计划日期_IN,计划人_IN,制定日期_IN
                str计划人 = Trim(.TextMatrix(lngLoop, HeadCol.计划人))
                str计划人 = "'" & IIf(str计划人 = "", gstrUserName, str计划人) & "'"
                
                str制定日期 = Trim(.TextMatrix(lngLoop, HeadCol.制定日期))
                str制定日期 = IIf(str制定日期 = "", Format(zldatabase.Currentdate, "yyyy-mm-dd"), str制定日期)
                str制定日期 = "to_date('" & str制定日期 & "','yyyy-mm-dd')"
                
                '计划序号
                If lngSN < lngLoop Or lngSN = 0 Then
                    gstrSQL = "Select a.Rec, b.Sn " & _
                              "From " & _
                              "  (Select 1 ID, Count(1) Rec From 应付记录 Where ID = [1] And 记录性质 = -1 And 计划序号 = [2]) A," & _
                              "  (Select 1 ID, Max(计划序号) Sn From 应付记录 Where ID = [1] And 记录性质 = -1) B " & _
                              "Where a.Id = b.Id "
                    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "计划序号", mlngID, lngLoop)
                    If Not rsTmp.EOF Then
                        If rsTmp!rec > 0 Then
                            '计划序号存在，取最大序号，并加1
                            lngSN = rsTmp!sn + 1
                        Else
                            '计划序号不存在，使用lngLoop
                            lngSN = lngLoop
                        End If
                    Else
                        lngSN = lngLoop
                    End If
                    rsTmp.Close
                Else
                    lngSN = lngSN + 1
                End If
                
                strSQL = "ZL_付款计划_INSERT (" & _
                    mlngID & "," & _
                    lngSN & "," & _
                    Val(mshList.TextMatrix(lngLoop, HeadCol.支付金额)) & ",TO_DATE('" & _
                    Format(mshList.TextMatrix(lngLoop, HeadCol.计划支付日期), "yyyy-MM-dd") & "','yyyy-MM-dd')," & _
                    str计划人 & "," & _
                    str制定日期 & ")"
                    
                zldatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
    End With
    gcnOracle.CommitTrans
    Save = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub mshList_AfterDeleteRow()
    cmdSave.Enabled = True
    cmdReset.Enabled = True
    mblnChange = True
    
End Sub

Private Sub mshList_EditChange(curText As String)
    cmdSave.Enabled = True
    cmdReset.Enabled = True
    mblnChange = True
    mshList.TextMatrix(mshList.Row, 3) = UserInfo.姓名
    If mshList.TextMatrix(mshList.Row, 4) = "" Then mshList.TextMatrix(mshList.Row, 4) = Format(zldatabase.Currentdate, "yyyy-MM-dd")
End Sub

Private Sub mshList_EnterCell(Row As Long, Col As Long)
    With mshList
        If Trim(.TextMatrix(Row, 2)) <> "" Then
            .Active = False
        Else
            Select Case Col
                Case 0, 4
                    .TxtCheck = True
                    .TextMask = "-0123456789"
                    .MaxLength = 16
                Case 1
                    .TxtCheck = True
                    .TextMask = ".0123456789"
                    .MaxLength = 10
            End Select
            .Active = True
        End If
    End With
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshList
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        If InStr(1, strKey, "'") <> 0 Then
            MsgBox "含有非常符,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If .ColData(.Col) = 0 Then
            Exit Sub
        End If
        Select Case .Col
            Case HeadCol.支付金额
                If strKey <> "" Then
                        If Not IsNumeric(strKey) And strKey <> "" Then
                            ShowMsgbox "支付金额必须为数字型,请重输！"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If Val(strKey) <= 0 Then
                            ShowMsgbox "支付金额必须大于零,请重输！"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strKey) < 0.001 Then
                                MsgBox "支付金额必须大于0.001,请重输！"
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                            
                        If Val(strKey) >= 10 ^ 11 - 1 Then
                            MsgBox "冲销数量必须小于" & (10 ^ 11 - 1)
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                Else
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    Else
                        .TxtVisible = True
                        .Text = " "
                        Exit Sub
                    End If
                End If
            Case HeadCol.计划支付日期
                
                If strKey <> "" Then
                    
                        If IsNumeric(strKey) Then
                            strKey = TranNumToDate(Val(strKey))
                        End If
                        
                        If Not IsDate(strKey) Then
                            ShowMsgbox "计划支付日期必须为日期格式(如:20030303 或 2003-03-03),请重输！"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                Else
                    If .TxtVisible = True Then
                        .Text = ""
                        Exit Sub
                    Else
                        .TxtVisible = True
                        .Text = ""
                        Exit Sub
                    End If
                End If
        End Select
    End With

End Sub

Private Sub mshList_LeaveCell(Row As Long, Col As Long)
    Dim ingLoop As Integer
    Dim curTemp As Currency
    Dim strTemp As String, strDate As String
    
    If Col = 1 Then
        On Error Resume Next
        txtInfo(4).Text = 0
        For ingLoop = 1 To mshList.Rows - 1
            curTemp = CDbl(Format(mshList.TextMatrix(ingLoop, 1), "0.00"))
            If Err.Number = 0 Then
                mshList.TextMatrix(ingLoop, 1) = Format(mshList.TextMatrix(ingLoop, 1), "0.00")
                txtInfo(4).Text = Format(CDbl(txtInfo(4).Text) + curTemp, "0.00")
            End If
            Err.Clear
        Next
        txtInfo(4).Text = Format(CDbl(txtInfo(3).Text) - CDbl(txtInfo(4).Text), "0.00")
        On Error GoTo 0
    End If
    If Col = 0 Or Col = 4 Then
        On Error Resume Next
        mshList.TextMatrix(Row, Col) = Format(mshList.TextMatrix(Row, Col), "yyyy-MM-dd")
        On Error GoTo 0
    End If
End Sub

