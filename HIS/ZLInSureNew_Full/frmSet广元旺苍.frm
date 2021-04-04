VERSION 5.00
Begin VB.Form frmSet广元旺苍 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo适用地区 
      Height          =   300
      Left            =   1200
      TabIndex        =   21
      Top             =   4125
      Width           =   2295
   End
   Begin VB.CheckBox CHK比较 
      Caption         =   "中心与医院结算费用总额不等，不能结算(&Y)"
      Height          =   240
      Left            =   1200
      TabIndex        =   19
      Top             =   3675
      Width           =   3840
   End
   Begin VB.CheckBox chk明细 
      Caption         =   "明细时实上传(&S)"
      Height          =   285
      Left            =   3675
      TabIndex        =   18
      Top             =   3315
      Width           =   1815
   End
   Begin VB.CommandButton cmd下载 
      Caption         =   "下载交易记录(&X)"
      Height          =   350
      Left            =   2010
      TabIndex        =   17
      Top             =   4740
      Width           =   1845
   End
   Begin VB.CheckBox chk个人帐户 
      Caption         =   "操作员直下个人帐户(&L)"
      Height          =   270
      Left            =   1200
      TabIndex        =   8
      Top             =   3330
      Width           =   3120
   End
   Begin VB.ComboBox cbo社保机构 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2955
      Width           =   4755
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   16
      Top             =   4590
      Width           =   7665
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   660
      Width           =   7665
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1980
      Left            =   435
      TabIndex        =   12
      Top             =   825
      Width           =   5835
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   4515
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   555
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1335
         Width           =   3075
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1200
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   945
         Width           =   3075
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   1
         Top             =   555
         Width           =   3075
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   330
         TabIndex        =   4
         Top             =   1395
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   2
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   615
         Width           =   810
      End
   End
   Begin VB.CommandButton cmd社保机构 
      Caption         =   "下载社保机构(&D)"
      Height          =   350
      Left            =   135
      TabIndex        =   11
      Top             =   4740
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4065
      TabIndex        =   9
      Top             =   4740
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5250
      TabIndex        =   10
      Top             =   4740
      Width           =   1100
   End
   Begin VB.Label lbl适用地区 
      AutoSize        =   -1  'True
      Caption         =   "适用地区"
      Height          =   180
      Left            =   390
      TabIndex        =   20
      Top             =   4200
      Width           =   720
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmSet广元旺苍.frx":0000
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "社保机构"
      Height          =   180
      Index           =   1
      Left            =   390
      TabIndex        =   6
      Top             =   3030
      Width           =   720
   End
   Begin VB.Label lbl 
      Caption         =   "配置相关的参数."
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   15
      Top             =   360
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet广元旺苍"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
End Enum



Public Function 参数设置() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    frmSet广元旺苍.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo社保机构_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk个人帐户_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub
Private Sub cmd社保机构_Click()
    Dim strOutput As String, StrInput As String
    Dim strArr, strArr1
    Dim rsTemp As New ADODB.Recordset
    Dim lng序号 As Long
    Dim i As Long
    
    If mcnTest Is Nothing Then
        MsgBox "请先测试中间库是否成常!"
        Exit Sub
    End If
    If mcnTest.State <> 1 Then
        MsgBox "请先测试中间库是否成常!"
        Exit Sub
    End If
    If SaveData = False Then Exit Sub

    If cbo社保机构.ListIndex >= 0 Then
        SaveRegInFor g公共模块, "医保", "社保机构代码", Split(cbo社保机构.Text, "--")(0)
    End If
    
'    If 医保初始化_广元旺苍 = False Then Exit Sub
'
    zlCommFun.ShowFlash "正在下载社保机构,请稍后..."


    If MsgBox("下载门诊的社保机构吗?", vbQuestion + vbDefaultButton1 + vbYesNo) = vbYes Then
        If 业务请求_广元旺苍(获得社保机构_旺苍, StrInput, strOutput) = False Then
            zlCommFun.StopFlash
            Exit Sub
        End If
    Else
        If 业务请求_广元旺苍(获得社保机构_住院_旺苍, StrInput, strOutput) = False Then
            zlCommFun.StopFlash
            Exit Sub
        End If
    
    End If
    If strOutput = "" Then
        zlCommFun.StopFlash
        Exit Sub
    End If
    strArr = Split(strOutput, "@$")
    For i = 0 To UBound(strArr)
        strArr1 = Split(strArr(i), "||")
        gstrSQL = "Select 序号 From 保险中心目录 where 险类=" & TYPE_广元旺苍 & "  and 编码 ='" & strArr1(0) & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If rsTemp.EOF Then
            '新增:
            gstrSQL = "Select nvl(Max(序号),0)+1 as 序号 from 保险中心目录 where 险类=" & TYPE_广元旺苍
            zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
            If rsTemp.EOF Then
                lng序号 = 1
            Else
                lng序号 = Nvl(rsTemp!序号, 1)
            End If
              gstrSQL = "zl_保险中心目录_Insert(" & TYPE_广元旺苍 & "," & lng序号 & ",'" & strArr1(0) & "','" & strArr1(1) & "')"
        Else
              gstrSQL = "zl_保险中心目录_Update(" & TYPE_广元旺苍 & "," & Nvl(rsTemp!序号, 0) & ",'" & strArr1(0) & "','" & strArr1(1) & "')"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    '重新加载数据
    Call LoadCbo
    zlCommFun.StopFlash
End Sub

Private Sub cmd下载_Click()
    Dim StrInput As String, strOutput As String

    If cbo社保机构.ListIndex < 0 Then
        ShowMsgbox "请选择社保机构!"
        Exit Sub
    End If
    
    StrInput = Split(cbo社保机构.Text, "-")(0)
    
    If 业务请求_广元旺苍(下载交易记录_旺苍, StrInput, strOutput) Then Exit Sub
    ShowMsgbox "下载成功!"
    
    
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    If mblnFirst = False Then Exit Sub
    
    
    '增加初始化数据
    Me.cbo适用地区.Clear
    Me.cbo适用地区.AddItem "资阳机车厂"
    Me.cbo适用地区.AddItem "旺苍、眉山"
    Me.cbo适用地区.ListIndex = 0
    
    mblnFirst = False
    Call LoadCbo
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_广元旺苍
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
            Case "操作员直下个人帐户"
                  chk个人帐户.Value = IIf(Nvl(!参数值, 1) = 1, 1, 0)
            Case "明细时实上传"
                  chk明细.Value = IIf(Nvl(!参数值, 1) = 1, 1, 0)
            Case "比较结算数据"
                 CHK比较.Value = IIf(Nvl(!参数值, 1) = 1, 1, 0)
            Case "适用地区"
                 Me.cbo适用地区.ListIndex = Nvl(!参数值)
            End Select
            .MoveNext
        Loop
    End With
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub lbl地区_Click()

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
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    '重置医保初始化状态
    mblnInit = False
   
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
        If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag, False) = False Then
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
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_广元旺苍 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_广元旺苍 & ",null,'医保用户名','" & txtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_广元旺苍 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_广元旺苍 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_保险参数_Insert(" & TYPE_广元旺苍 & ",null,'明细时实上传','" & IIf(chk明细.Value = 1, 1, 0) & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
   gstrSQL = "zl_保险参数_Insert(" & TYPE_广元旺苍 & ",null,'操作员直下个人帐户','" & IIf(chk个人帐户.Value = 1, 1, 0) & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
   
   gstrSQL = "zl_保险参数_Insert(" & TYPE_广元旺苍 & ",null,'比较结算数据','" & IIf(CHK比较.Value = 1, 1, 0) & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
   gstrSQL = "zl_保险参数_Insert(" & TYPE_广元旺苍 & ",null,'适用地区','" & cbo适用地区.ListIndex & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    If cbo社保机构.ListIndex >= 0 Then
        SaveRegInFor g公共模块, "医保", "社保机构代码", Split(cbo社保机构.Text, "--")(0)
    End If
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
Private Sub LoadCbo()
        '加载Grid数据
        Err = 0
        On Error GoTo ErrHand:
        Dim rsTemp As New ADODB.Recordset
        Dim i As Long
        gstrSQL = "Select * From 保险中心目录 where 险类=" & TYPE_广元旺苍 & " and 序号<>0 order by 编码"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取社保机构目录" '  rsTemp.Open gstrSQL, gcnOracle_广元旺苍
        
        With rsTemp
            i = 1
            Me.cbo社保机构.Clear
            Do While Not .EOF
                cbo社保机构.AddItem Nvl(!编码) & "--" & Nvl(!名称)
                cbo社保机构.ItemData(cbo社保机构.NewIndex) = Nvl(rsTemp!序号, 0)
                .MoveNext
            Loop
        End With
        SetDefaultSel
        Exit Sub
ErrHand:
        If ErrCenter = 1 Then Resume
End Sub
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo ErrHand:
    Call GetRegInFor(g公共模块, "医保", "社保机构代码", strReg)
    If cbo社保机构.ListCount = 0 Then Exit Function
    For i = 0 To cbo社保机构.ListCount - 1
        If Split(cbo社保机构.List(i), "--")(0) = strReg Then
            cbo社保机构.ListIndex = i
            Exit For
        End If
    Next
    If cbo社保机构.ListIndex < 0 Then
        cbo社保机构.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
