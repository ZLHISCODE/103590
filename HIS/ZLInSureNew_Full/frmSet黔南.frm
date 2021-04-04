VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmSet黔南 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk读卡器 
      Caption         =   "本站点有读卡器."
      Height          =   345
      Left            =   1560
      TabIndex        =   21
      Top             =   4395
      Width           =   2190
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   105
      TabIndex        =   3
      Top             =   210
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "服务器设置(&S)"
      TabPicture(0)   =   "frmSet黔南.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra医保服务器"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "收费类别(&T)"
      TabPicture(1)   =   "frmSet黔南.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "mshBill"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra医保服务器 
         Caption         =   "医院前置医保服务器"
         Height          =   1605
         Left            =   180
         TabIndex        =   12
         Top             =   510
         Width           =   4155
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   16
            Top             =   330
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1260
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   15
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   14
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "测试(&T)"
            Height          =   1095
            Left            =   3000
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "用户名(&U)"
            Height          =   180
            Index           =   0
            Left            =   390
            TabIndex        =   19
            Top             =   390
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "密码(&P)"
            Height          =   180
            Index           =   1
            Left            =   570
            TabIndex        =   18
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "服务器(&S)"
            Height          =   180
            Index           =   2
            Left            =   390
            TabIndex        =   17
            Top             =   1170
            Width           =   810
         End
      End
      Begin VB.Frame fra 
         Caption         =   "医保中心前置医保服务器"
         Height          =   1605
         Left            =   180
         TabIndex        =   4
         Top             =   2280
         Width           =   4155
         Begin VB.CommandButton cmd测试 
            Caption         =   "测试(&4)"
            Height          =   1095
            Left            =   3000
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   330
            Width           =   1005
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   7
            Top             =   1110
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   1260
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   5
            Top             =   330
            Width           =   1635
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "服务器(&3)"
            Height          =   180
            Index           =   3
            Left            =   390
            TabIndex        =   11
            Top             =   1170
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "密码(&2)"
            Height          =   180
            Index           =   4
            Left            =   570
            TabIndex        =   10
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "用户名(&1)"
            Height          =   180
            Index           =   5
            Left            =   390
            TabIndex        =   9
            Top             =   390
            Width           =   810
         End
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   3495
         Left            =   -74940
         TabIndex        =   20
         Top             =   465
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   6165
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
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
   End
   Begin VB.CommandButton cmdCardPara 
      Caption         =   "设置读卡器参数"
      Height          =   660
      Left            =   4905
      TabIndex        =   2
      Top             =   3630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4965
      TabIndex        =   1
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4965
      TabIndex        =   0
      Top             =   930
      Width           =   1100
   End
End
Attribute VB_Name = "frmSet黔南"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcnTest As New ADODB.Connection
Private mcnZxTest   As New ADODB.Connection
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
    text中心用户 = 3
    Text中心密码 = 4
    Text中心服务器 = 5
End Enum
Private Enum mColHead
    收费类别 = 0
    保费项目
End Enum
Private Function LoadCbo() As Boolean
    Dim rsTemp As New ADODB.Recordset
    If gcnOracle_黔南 Is Nothing Then Exit Function
    gstrSQL = " select * from 医保收费类别 "
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle_黔南
    mshBill.Clear
    Do While Not rsTemp.EOF
        mshBill.AddItem rsTemp!编码 & "-" & rsTemp!名称
        rsTemp.MoveNext
    Loop
    LoadCbo = True
End Function
 Private Function iniData() As Boolean
    '初始数据
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTmp As String
    
    '设置页头
    Err = 0
    On Error Resume Next
    '设置表列头
    Call initGrid
    strSQL = "" & _
        "   Select A.类别,b.参数值 From 收费类别 a,(Select * From 保险参数 where 险类=" & TYPE_黔南 & ") b " & _
        "   Where A.类别=b.参数名(+) " & _
        "   order by A.编码 "
        
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    With mshBill
        .ClearBill
        If rsTmp.RecordCount = 0 Then
            .Rows = 2
        Else
            .Rows = rsTmp.RecordCount + 1
        End If
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, mColHead.收费类别) = Nvl(rsTmp!类别)
            strTmp = Nvl(rsTmp!参数值)
            
            If Trim(strTmp) <> "" Then
                .TextMatrix(lngRow, mColHead.保费项目) = strTmp
            End If
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        
    End With
    
End Function
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = 2
        
        .msfObj.FixedCols = 1
        .AllowAddRow = False
        
        .TextMatrix(0, mColHead.收费类别) = "收费类别"
        .TextMatrix(0, mColHead.保费项目) = "中心项目"
        
        
        .ColWidth(mColHead.收费类别) = 1500
        .ColWidth(mColHead.保费项目) = 2000
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(mColHead.收费类别) = 5
        .ColData(mColHead.保费项目) = 3
        
        .ColAlignment(mColHead.收费类别) = flexAlignLeftCenter
        .ColAlignment(mColHead.保费项目) = flexAlignLeftCenter
        .PrimaryCol = mColHead.保费项目
        .LocateCol = mColHead.保费项目
    End With
End Sub


Public Function 参数设置() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    frmSet黔南.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Text) = False Then
        Exit Sub
    End If
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub

Private Sub cmdCardPara_Click()
    If sCard_SetupCardOption_黔南 = False Then Exit Sub
End Sub

Private Sub cmd测试_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnZxTest.State = adStateOpen Then mcnZxTest.Close
    
    If OraDataOpen(mcnZxTest, txtEdit(Text中心服务器).Text, txtEdit(text中心用户).Text, txtEdit(Text中心密码).Text) = False Then
        Exit Sub
    End If
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
     
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_黔南
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "医保用户名"
                  txtEdit(text医保用户).Text = Nvl(!参数值)
            Case "医保用户密码"
                  txtEdit(Text医保密码).Text = Nvl(!参数值)
            Case "医保服务器"
                  txtEdit(Text医保服务器).Text = Nvl(!参数值)
            Case "中心用户名"
                  txtEdit(text中心用户).Text = Nvl(!参数值)
            Case "中心用户密码"
                  txtEdit(Text中心密码).Text = Nvl(!参数值)
            Case "中心服务器"
                  txtEdit(Text中心服务器).Text = Nvl(!参数值)
            End Select
            .MoveNext
        Loop
    End With
    GetRegInFor g公共全局, "医保", "读卡器", strReg
    If Val(strReg) = 1 Then
        chk读卡器.Value = 1
    Else
        chk读卡器.Value = 0
    End If
    Call iniData
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        If gcnOracle_黔南 Is Nothing Then
            If Open中间库() = False Then Exit Sub
        End If
        LoadCbo
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text医保密码 Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Text医保服务器 Or Index = Text医保密码 Or Index = text医保用户 Then
        '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
    If Index = text中心用户 Or Index = Text中心密码 Or Index = Text中心服务器 Then
        If mcnZxTest.State = adStateOpen Then mcnZxTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Text, False) = False Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If mcnZxTest.State = adStateClosed Then
        If OraDataOpen(mcnZxTest, txtEdit(Text中心服务器).Text, txtEdit(text中心用户).Text, txtEdit(Text中心密码).Text, False) = False Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    Dim lngRow As Long
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_黔南 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    With mshBill
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, mColHead.收费类别) <> "" Then
                '新增参数数据
                gstrSQL = "zl_保险参数_Insert(" & TYPE_黔南 & ",null,'" & .TextMatrix(lngRow, mColHead.收费类别) & "' ,'" & .TextMatrix(lngRow, mColHead.保费项目) & "'," & lngRow + 2 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_黔南 & ",null,'医保用户名','" & txtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_黔南 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_黔南 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    gstrSQL = "zl_保险参数_Insert(" & TYPE_黔南 & ",null,'中心用户名','" & txtEdit(text中心用户).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_黔南 & ",null,'中心用户密码','" & txtEdit(Text中心密码).Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_黔南 & ",null,'中心服务器','" & txtEdit(Text中心服务器).Text & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    SaveRegInFor g公共全局, "医保", "读卡器", IIf(chk读卡器.Value = 1, 1, 0)
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
End Sub
Private Function Open中间库() As Boolean
    '连接中间库
    '中间库连接
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strServer As String, strPass As String
    
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "渝北医保", TYPE_黔南)
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_黔南 = New ADODB.Connection

    If OraDataOpen(gcnOracle_黔南, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If


    Open中间库 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
