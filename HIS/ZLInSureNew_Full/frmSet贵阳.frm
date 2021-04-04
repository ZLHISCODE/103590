VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmSet贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险参数设置"
   ClientHeight    =   5640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8850
   Icon            =   "frmSet贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo病历检查 
      ForeColor       =   &H80000012&
      Height          =   300
      Left            =   6630
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1050
      Width           =   2175
   End
   Begin VB.Frame fra 
      Caption         =   "保险参数"
      Height          =   3225
      Left            =   30
      TabIndex        =   8
      Top             =   1740
      Width           =   4365
      Begin VB.ComboBox cbo读卡器端口号 
         Height          =   300
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2610
         Width           =   1995
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   16
         Top             =   1545
         Width           =   1995
      End
      Begin VB.CheckBox chk病种 
         Caption         =   "支持特殊门诊(&T)"
         Height          =   255
         Left            =   990
         TabIndex        =   10
         Top             =   450
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Left            =   3930
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1575
         Width           =   255
      End
      Begin VB.CheckBox chk参保前在院 
         Caption         =   "入院时选择参保前在院(&T)"
         Height          =   255
         Left            =   990
         TabIndex        =   11
         Top             =   690
         Width           =   2385
      End
      Begin VB.CheckBox chk收费 
         Caption         =   "门诊采用连续收费(&L)"
         Height          =   255
         Left            =   990
         TabIndex        =   9
         Top             =   180
         Width           =   2055
      End
      Begin VB.TextBox txt项目数 
         Height          =   300
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   14
         Top             =   1200
         Width           =   1995
      End
      Begin VB.ComboBox cbo出院操作 
         Height          =   300
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1890
         Width           =   1995
      End
      Begin VB.CheckBox chk特殊药品提示 
         Caption         =   "特殊药品提示(&M)"
         Height          =   255
         Left            =   990
         TabIndex        =   12
         Top             =   930
         Width           =   2385
      End
      Begin VB.TextBox txt时间 
         Height          =   300
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   20
         Text            =   "20"
         Top             =   2250
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "读卡器端口号(D)"
         Height          =   180
         Left            =   810
         TabIndex        =   21
         Top             =   2670
         Width           =   1350
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "frmSet贵阳.frx":000C
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lbl服务器 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "医保服务器(&S)"
         Height          =   180
         Left            =   990
         TabIndex        =   15
         Top             =   1620
         Width           =   1170
      End
      Begin VB.Label lbl项目数 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "门诊最大项目数(&N)"
         Height          =   180
         Left            =   630
         TabIndex        =   13
         Top             =   1275
         Width           =   1530
      End
      Begin VB.Label lbl出院操作 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出院操作(&R)"
         Height          =   180
         Left            =   1170
         TabIndex        =   17
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "门诊结算窗口隔           秒自动关闭"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   19
         Top             =   2295
         Width           =   3150
      End
   End
   Begin VB.CheckBox chk本机需要进行医保交易 
      Caption         =   "本机需要进行医保交易"
      Height          =   315
      Left            =   60
      TabIndex        =   33
      Top             =   5220
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -150
      TabIndex        =   32
      Top             =   5040
      Width           =   10395
   End
   Begin VB.Frame Frame3 
      Caption         =   "门诊慢性疾病用药限制"
      Height          =   915
      Left            =   4440
      TabIndex        =   23
      Top             =   90
      Width           =   4365
      Begin VB.CheckBox chk启用用药限制 
         Caption         =   "启用门诊慢性疾病用药限制功能"
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   2835
      End
      Begin VB.ComboBox cbo时间 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   570
         Width           =   795
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "累计用药量按         进行控制"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   2610
      End
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1605
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   4365
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
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
         TabIndex        =   4
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1635
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   30
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7470
      TabIndex        =   31
      Top             =   5190
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3225
      Left            =   4440
      TabIndex        =   29
      Top             =   1740
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   5689
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "住院结帐时未登记病历记录"
      Height          =   180
      Left            =   4440
      TabIndex        =   27
      Top             =   1110
      Width           =   2160
   End
End
Attribute VB_Name = "frmSet贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'只有所有者才能设置收费类别与发票归属科目编码的对应关系
Dim mlng险类 As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
End Enum

Dim mcnTest As New ADODB.Connection

Private Sub Bill_cboClick(ListIndex As Long)
    If Bill.Active = False Then Exit Sub
    Bill.TextMatrix(Bill.Row, 2) = Bill.CboText
End Sub

Private Sub chk启用用药限制_Click()
    cbo时间.Enabled = (chk启用用药限制.Value = 1): lbl时间(1).Enabled = cbo时间.Enabled
End Sub
Private Sub cmdSelect_Click()
    Dim strServer As String
    
    strServer = GetComputer(Me, "选择医保服务器")
    If strServer <> "" Then
        txtServer.Text = strServer
        mblnChange = True
    End If
End Sub

Private Sub cmdTest_Click()
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Text) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "Bill" Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.cbo出院操作
        .Clear
        .AddItem "HIS、医保同时出院"
        .AddItem "仅办理HIS出院"
        .ListIndex = 0
    End With
    
    With Me.cbo读卡器端口号
        .Clear
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .AddItem "COM5"
        .AddItem "COM6"
        .AddItem "COM7"
        .AddItem "COM8"
        .AddItem "COM9"
        .AddItem "USB"
        .ListIndex = 0
    End With
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
    Dim colPara As New Collection
    Dim lngCount As Long
    
    If mcnTest.State = adStateOpen Then mcnTest.Close
    If OraDataOpen(mcnTest, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Text) = False Then
        Exit Sub
    End If
    
    If txtServer.Text = "" Then
        MsgBox "医保服务器名不能为空。", vbInformation, gstrSysName
        txtServer.SetFocus
        Exit Sub
    End If
    If IsNumeric(txt项目数.Text) = False Then
        MsgBox "请输入正确的项目数。", vbInformation, gstrSysName
        txt项目数.SetFocus
        Exit Sub
    End If
    If zlCommFun.StrIsValid(txtServer.Text, txtServer.MaxLength) = False Then
        txtServer.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    '这一部分参数不区分中心
    colPara.Add "null,'医保用户名','" & TxtEdit(text医保用户).Text
    colPara.Add "null,'医保用户密码','" & TxtEdit(Text医保密码).Text
    colPara.Add "null,'医保服务器1','" & TxtEdit(Text医保服务器).Text
    
    colPara.Add "null,'门诊连续收费','" & chk收费.Value
    colPara.Add "null,'支持特殊门诊','" & chk病种.Value
    colPara.Add "null,'入院时选择参保前在院','" & chk参保前在院.Value
    colPara.Add "null,'医保服务器','" & txtServer.Text
    colPara.Add "null,'门诊最大项目数','" & txt项目数.Text
    colPara.Add "null,'特殊药品提示','" & chk特殊药品提示.Value
    colPara.Add "null,'门诊结算窗口关闭时间','" & txt时间.Text
    colPara.Add "null,'启用门诊慢性疾病用药限制功能','" & chk启用用药限制.Value
    colPara.Add "null,'累计用药量计算标准','" & Split(cbo时间.Text, "-")(0)
    colPara.Add "null,'病历检查','" & cbo病历检查.ListIndex
    
    For lngCount = 1 To colPara.Count
        gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & colPara(lngCount) & "'," & lngCount & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    '归属科目编码的保存
    For lngCount = 1 To Bill.Rows - 1
        gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'" & Bill.TextMatrix(lngCount, 0) & "','" & Mid(Bill.TextMatrix(lngCount, 2), 1, 2) & "'," & lngCount + 5 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'出院操作','" & Me.cbo出院操作.ListIndex & "',90)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    SaveSetting "ZLSOFT", "公共全局", "贵阳医保交易", chk本机需要进行医保交易.Value
    gcnOracle.CommitTrans
    
    SaveSetting "ZLSOFT", "公共模块\贵阳市医保", "端口", cbo读卡器端口号.Text
    
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub txtServer_Change()
    mblnChange = True
End Sub

Private Sub txtServer_GotFocus()
    zlControl.TxtSelAll txtServer
End Sub

Public Function 参数设置(ByVal lng险类 As Long) As Boolean
'功能：设置我们中联医保所需要的参数
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    mlng险类 = lng险类
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1] and 中心 is null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng险类)
    cbo时间.Clear
    cbo时间.AddItem "1-年"
    cbo时间.AddItem "2-月"
    cbo时间.AddItem "3-周"
    cbo时间.AddItem "4-天"
    cbo时间.ListIndex = 1
    
    cbo病历检查.Clear
    cbo病历检查.AddItem "0.可以正常结帐"
    cbo病历检查.AddItem "1.提示是否结帐"
    cbo病历检查.AddItem "2.禁止结帐"
    cbo病历检查.ListIndex = 0
    
    cbo读卡器端口号.Text = GetSetting("ZLSOFT", "公共模块\贵阳市医保", "端口", "COM1")
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                Me.TxtEdit(text医保用户).Text = Nvl(rsTemp!参数值)
            Case "医保用户密码"
                Me.TxtEdit(Text医保密码).Text = Nvl(rsTemp!参数值)
            Case "医保服务器1"
                Me.TxtEdit(Text医保服务器).Text = Nvl(rsTemp!参数值)
            Case "门诊连续收费"
                chk收费.Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "支持特殊门诊"
                chk病种.Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "入院时选择参保前在院"
                chk参保前在院.Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "医保服务器"
                txtServer.Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "门诊最大项目数"
                txt项目数.Text = IIf(IsNull(rsTemp("参数值")), "7", rsTemp("参数值"))
            Case "出院操作"
                Me.cbo出院操作.ListIndex = Nvl(rsTemp!参数值, 0)
            Case "特殊药品提示"
                Me.chk特殊药品提示.Value = Nvl(rsTemp!参数值, 0)
            Case "门诊结算窗口关闭时间"
                txt时间.Text = Nvl(rsTemp!参数值, 20)
            Case "启用门诊慢性疾病用药限制功能"
                chk启用用药限制.Value = Nvl(rsTemp!参数值, 0)
            Case "累计用药量计算标准"
                cbo时间.ListIndex = Val(Nvl(rsTemp!参数值, 2)) - 1
            Case "病历检查"
                cbo病历检查.ListIndex = Nvl(rsTemp!参数值, 0)
        End Select
        
        rsTemp.MoveNext
    Loop
    
    '读取出已设置好的归属科目编码对应关系，仅所有都允许修改
    '发票归属科目编码
    '01：西药；02：中成药；03：中草药；04：床位费；05：诊查费；06：检查费；
    '07：治疗费；08：护理费；09：手术费；10：化验费；11：其他
    gstrSQL = "Select 编码,类别,'11-其他' 归属科目编码  " & _
             " From 收费类别 " & _
             " Where 编码 Not IN  " & _
             "     (Select 参数名 From 保险参数 Where 险类=" & lng险类 & " And 序号>=6) " & _
             " union   " & _
             " Select B.编码,B.类别,decode(A.参数值,'01','01-西药','02','02-中成药', " & _
             " '03','03-中草药','04','04-床位费','05','05-诊查费','06','06-检查费','07','07-治疗费', " & _
             " '08','08-护理费','09','09-手术费','10','10-化验费','11-其他') 归属科目编码   " & _
             " From 保险参数 A,收费类别 B " & _
             " Where A.序号>=6 And A.险类=[1] And A.参数名=B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别", lng险类)
    '初始化单据控件
    With Bill
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 0) = "编码"
        .TextMatrix(0, 1) = "收费类别"
        .TextMatrix(0, 2) = "归属科目编码"
        .ColWidth(0) = 500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1800
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColData(0) = 0
        .ColData(2) = 3

        .AddItem "01-西药"
        .AddItem "02-中成药"
        .AddItem "03-中草药"
        .AddItem "04-床位费"
        .AddItem "05-诊查费"
        .AddItem "06-检查费"
        .AddItem "07-治疗费"
        .AddItem "08-护理费"
        .AddItem "09-手术费"
        .AddItem "10-化验费"
        .AddItem "11-其他"
        .ListIndex = 10
        
        .PrimaryCol = 0
        .LocateCol = 2
    End With
    
    With rsTemp
        Do While Not .EOF
            Bill.TextMatrix(.AbsolutePosition, 0) = !编码
            Bill.TextMatrix(.AbsolutePosition, 1) = !类别
            Bill.TextMatrix(.AbsolutePosition, 2) = !归属科目编码
            .MoveNext
            Bill.Rows = Bill.Rows + 1
        Loop
        If .RecordCount <> 0 Then Bill.Rows = Bill.Rows - 1
        Bill.Row = 1
    End With
    
    Bill.AllowAddRow = False
    Bill.Active = OwnerUser(gstrDbUser)
    chk本机需要进行医保交易.Value = GetSetting(appName:="ZLSOFT", Section:="公共全局", Key:="贵阳医保交易", Default:="1")
    cbo时间.Enabled = (chk启用用药限制.Value = 1): lbl时间(1).Enabled = cbo时间.Enabled
    mblnChange = False
    frmSet贵阳.Show vbModal, frm医保类别
    参数设置 = mblnOK
End Function
Private Sub txt时间_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt时间, KeyAscii, m数字式)
End Sub
