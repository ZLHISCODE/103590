VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSet成都内江 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab stab 
      Height          =   2865
      Left            =   255
      TabIndex        =   24
      Top             =   765
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5054
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "常规(&A)"
      TabPicture(0)   =   "frmSet成都内江.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEdit(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEdit(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl读卡器"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cbo读卡器"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtPort"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "前置机配置(&Q)"
      TabPicture(1)   =   "frmSet成都内江.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra医保服务器"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra 
         Caption         =   "医保连接配置"
         Height          =   1380
         Index           =   2
         Left            =   150
         TabIndex        =   26
         Top             =   615
         Width           =   5490
         Begin VB.TextBox txt天数 
            Height          =   300
            Left            =   4150
            MaxLength       =   2
            TabIndex        =   28
            Top             =   240
            Width           =   360
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   300
            Left            =   5100
            TabIndex        =   6
            Top             =   930
            Width           =   300
         End
         Begin VB.TextBox txtIP 
            Height          =   300
            Left            =   870
            TabIndex        =   3
            Top             =   600
            Width           =   4545
         End
         Begin VB.TextBox Txt端口号 
            Height          =   300
            Left            =   870
            TabIndex        =   1
            Top             =   240
            Width           =   1290
         End
         Begin VB.TextBox txtFile 
            Height          =   300
            Left            =   870
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "C:\"
            Top             =   945
            Width           =   4545
         End
         Begin VB.Label Lbl天数 
            Caption         =   "允许补办天数"
            Height          =   255
            Left            =   3050
            TabIndex        =   27
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblIn 
            AutoSize        =   -1  'True
            Caption         =   "&IP地址"
            Height          =   180
            Index           =   2
            Left            =   300
            TabIndex        =   2
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblIn 
            AutoSize        =   -1  'True
            Caption         =   "端口号"
            Height          =   180
            Index           =   1
            Left            =   300
            TabIndex        =   0
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lblIn 
            AutoSize        =   -1  'True
            Caption         =   "配置文件"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   1005
            Width           =   720
         End
      End
      Begin VB.TextBox TxtPort 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4305
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "1"
         Top             =   2220
         Width           =   360
      End
      Begin VB.ComboBox cbo读卡器 
         Height          =   300
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2220
         Width           =   1770
      End
      Begin VB.Frame fra医保服务器 
         Caption         =   "医院前置医保服务器"
         Height          =   1545
         Left            =   -74610
         TabIndex        =   25
         Top             =   720
         Width           =   4965
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   13
            Top             =   330
            Width           =   2385
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
            Width           =   2385
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   17
            Top             =   1110
            Width           =   2385
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "测试(&T)"
            Height          =   1095
            Left            =   3870
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "用户名(&U)"
            Height          =   180
            Index           =   0
            Left            =   390
            TabIndex        =   12
            Top             =   390
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "密码(&P)"
            Height          =   180
            Index           =   1
            Left            =   570
            TabIndex        =   14
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "服务器(&S)"
            Height          =   180
            Index           =   2
            Left            =   390
            TabIndex        =   16
            Top             =   1170
            Width           =   810
         End
      End
      Begin VB.Label lbl读卡器 
         AutoSize        =   -1  'True
         Caption         =   "读卡器(&R)"
         Height          =   180
         Left            =   420
         TabIndex        =   7
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "当前串口(&D)"
         Height          =   180
         Index           =   3
         Left            =   3300
         TabIndex        =   9
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "号串口"
         Height          =   180
         Index           =   4
         Left            =   4695
         TabIndex        =   11
         Top             =   2280
         Width           =   540
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   3765
      Width           =   6600
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   15
      TabIndex        =   22
      Top             =   615
      Width           =   6375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3765
      TabIndex        =   20
      Top             =   3915
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5040
      TabIndex        =   19
      Top             =   3915
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   5835
      Top             =   -45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "请设置以下相关医保前置服务器及读卡器的驱动厂商。"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   21
      Top             =   315
      Width           =   7125
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmSet成都内江.frx":0038
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "frmSet成都内江"
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


Private Sub cbo读卡器_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSel_Click()
    Dim strFile As String
    
    Err = 0
    On Error Resume Next
    With dlg
        .Filter = "配置文件(*.ini)|*.ini;*.txt"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Sub
        strFile = .FileName
    End With
    Err = 0
    txtFile.Text = strFile
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Text) = False Then
        Exit Sub
    End If
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub


Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    Dim i As Integer
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
     
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_成都内江
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
            Case "允许补办天数"
                  txt天数.Text = IIf(IsNull(!参数值), 0, !参数值)
            End Select
            .MoveNext
        Loop
    End With
    GetRegInFor g公共全局, "医保", "串口号", strReg
    Me.txtPort.Text = IIf(strReg = "", 1, strReg)
    
    GetRegInFor g公共全局, "医保", "读卡器", strReg
    For i = 0 To cbo读卡器.ListCount - 1
        If i = Val(strReg) Then
            cbo读卡器.ListIndex = i
        End If
    Next
    
    GetRegInFor g公共全局, "医保", "ConfigFileName", strReg
    txtFile.Text = strReg
    GetRegInFor g公共全局, "医保", "HostPort", strReg
    txt端口号.Text = strReg
    GetRegInFor g公共全局, "医保", "IPAddress", strReg
    txtIP.Text = strReg
    
 End Sub
Private Sub Form_Load()
    mblnFirst = True
    Call LoadBaseData
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
    
    If txtFile.Text = "" Then
'        MsgBox "配置文件未选择", vbInformation + vbDefaultButton1, gstrSysName
'        stab.Tab = 0
'        If txtFile.Enabled Then txtFile.SetFocus
'        Exit Function
    End If
    
    If Dir(txtFile.Text) <> "" Then
    Else
        MsgBox "配置文件不存在", vbInformation + vbDefaultButton1, gstrSysName
        stab.Tab = 0
        If txtFile.Enabled Then txtFile.SetFocus
        Exit Function
    End If
    If Trim(txtIP.Text) = "" Then
        MsgBox "IP未输入", vbInformation + vbDefaultButton1, gstrSysName
        stab.Tab = 0
        If txtIP.Enabled Then txtIP.SetFocus
        Exit Function
    End If
    If Trim(txt端口号.Text) = "" Then
        MsgBox "端口号未输入", vbInformation + vbDefaultButton1, gstrSysName
        stab.Tab = 0
        If txt端口号.Enabled Then txt端口号.SetFocus
        Exit Function
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
    gstrSQL = "zl_保险参数_Delete(" & TYPE_成都内江 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都内江 & ",null,'医保用户名','" & txtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都内江 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都内江 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都内江 & ",null,'允许补办天数','" & txt天数.Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    SaveRegInFor g公共全局, "医保", "读卡器", Split(cbo读卡器.Text, "-")(0)
    SaveRegInFor g公共全局, "医保", "串口号", Val(txtPort.Text)
    SaveRegInFor g公共全局, "医保", "ConfigFileName", txtFile.Text
    SaveRegInFor g公共全局, "医保", "HostPort", txt端口号.Text
    SaveRegInFor g公共全局, "医保", "IPAddress", txtIP.Text
    
    
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
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "渝北医保", TYPE_成都内江)
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
    Set gcnOracle_成都内江 = New ADODB.Connection

    If OraDataOpen(gcnOracle_成都内江, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    Open中间库 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub LoadBaseData()
    '加载数据
    With cbo读卡器
        .Clear
        .AddItem "0-明华读卡驱动器"
        .ListIndex = .NewIndex
        .AddItem "1-德森读卡驱动器"
    End With
End Sub



Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Sub txtIP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub TxtPort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        stab.Tab = 1
        If txtEdit(text医保用户).Enabled Then txtEdit(text医保用户).SetFocus
    End If
End Sub

Private Sub TxtPort_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtPort, KeyAscii, m数字式
End Sub

Public Function 参数设置() As Boolean
'功能：设置与东大阿尔派的医保接口

    
    mblnOK = False
    
    On Error GoTo errHandle
    mblnChange = False
    frmSet成都内江.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Txt端口号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txt天数_KeyPress(KeyAscii As Integer)
        zlControl.TxtCheckKeyPress txt天数, KeyAscii, m数字式
End Sub
