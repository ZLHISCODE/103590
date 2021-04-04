VERSION 5.00
Begin VB.Form frmSet昭通 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险参数设置"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtCOM 
      Height          =   300
      Left            =   1275
      TabIndex        =   4
      Top             =   1425
      Width           =   915
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试(&T)"
      Height          =   400
      Left            =   270
      TabIndex        =   5
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   3195
      TabIndex        =   7
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   2085
      TabIndex        =   6
      Top             =   2145
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   4590
   End
   Begin VB.TextBox txtSN 
      Height          =   300
      Left            =   1282
      TabIndex        =   3
      Top             =   1020
      Width           =   3015
   End
   Begin VB.TextBox txtPort 
      Height          =   300
      Left            =   1282
      MaxLength       =   4
      TabIndex        =   2
      Top             =   615
      Width           =   915
   End
   Begin VB.TextBox txtServer 
      Height          =   300
      Left            =   1282
      TabIndex        =   1
      Top             =   210
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "读卡器端口"
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   11
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "许可证编号"
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   1095
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务器端口"
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务器地址"
      Height          =   180
      Index           =   0
      Left            =   277
      TabIndex        =   0
      Top             =   285
      Width           =   900
   End
End
Attribute VB_Name = "frmSet昭通"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnIsOK As Boolean

Private Sub cmdCancel_Click()
    If UCase(txtServer.Text) <> UCase(txtServer.Tag) Then
        If MsgBox("数据进行了修改，未保存就退出吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    If UCase(txtSN.Text) <> UCase(txtSN.Tag) Then
        If MsgBox("数据进行了修改，未保存就退出吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    If UCase(txtPort.Text) <> UCase(txtPort.Tag) Then
        If MsgBox("数据进行了修改，未保存就退出吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    If UCase(txtCom.Text) <> UCase(txtCom.Tag) Then
        If MsgBox("数据进行了修改，未保存就退出吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    blnIsOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    If IsValid() = False Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_昭通 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_昭通 & ",null,'昭通许可证','" & txtSN.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_昭通 & ",null,'昭通服务器','" & txtServer.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_昭通 & ",null,'昭通端口号','" & txtPort.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    '将当前使用的串口写入注册表之中
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", CStr(txtCom.Text)
    blnIsOK = True
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTest_Click()
    txtServer.Text = Trim(txtServer.Text)
    txtPort.Text = Trim(txtPort.Text)
    txtSN.Text = Trim(txtSN.Text)

    If txtServer.Text = "" Or txtPort.Text = "" Or txtSN.Text = "" Then
        MsgBox "参数设置不完整，不能进行连接测试", vbInformation, "参数设置"
        Exit Sub
    End If
    If frmConn昭通.ConnCenter(txtServer.Text, txtPort.Text, txtSN.Text, UserInfo.ID) = True Then
        MsgBox "服务器连接成功", vbInformation, "连接"
        frmConn昭通.ConnClose
    Else
        MsgBox "服务器连接失败", vbInformation, "连接"
    End If
End Sub

Public Function 参数设置() As Boolean
    Dim rsTemp As New ADODB.Recordset, str参数值 As String, strCN As String, strServer As String, lngPort As Long
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_昭通)
        
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "昭通许可证"
                strCN = str参数值
            Case "昭通服务器"
                strServer = str参数值
            Case "昭通端口号"
                lngPort = CLng(str参数值)
        End Select
        rsTemp.MoveNext
    Loop
    
    txtServer.Text = strServer
    txtServer.Tag = strServer
    txtPort.Text = lngPort
    txtPort.Tag = lngPort
    txtSN.Text = strCN
    txtSN.Tag = strCN
    txtCom.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", "1")
    txtCom.Tag = txtCom.Text
    Me.Show vbModal
    参数设置 = blnIsOK
End Function

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    txtServer.Text = Trim(txtServer.Text)
    txtPort.Text = Trim(txtPort.Text)
    txtSN.Text = Trim(txtSN.Text)
    txtCom.Text = Trim(txtCom.Text)

    If txtServer.Text = "" Or txtSN.Text = "" Or txtPort.Text = "" Or txtCom.Text = "" Then
        MsgBox "请输入完整的医保参数", vbInformation, "参数设置"
        IsValid = False
        Exit Function
    End If
    
    '逐步判断字符的合法性
    If zlCommFun.StrIsValid(txtServer.Text, txtServer.MaxLength) = False Then
        zlControl.TxtSelAll txtServer
        txtServer.SetFocus
        Exit Function
    End If
    
    If zlCommFun.StrIsValid(txtSN.Text, txtSN.MaxLength) = False Then
        zlControl.TxtSelAll txtSN
        txtSN.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtCom.Text) Then
        MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        txtCom.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtPort.Text) Then
        MsgBox "请将服务器端口号输入数字信息", vbInformation, gstrSysName
        txtPort.SetFocus
        Exit Function
    End If
    
    '对连接进行测试
    If frmConn昭通.ConnCenter(txtServer.Text, txtPort.Text, txtSN.Text, UserInfo.ID) = False Then
        If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, "连接失败") = vbNo Then Exit Function
    Else
        frmConn昭通.ConnClose
    End If
    IsValid = True
End Function

Private Sub txtCOM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdTest.SetFocus
End Sub

Private Sub TxtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtSN.SetFocus
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtPort.SetFocus
        
End Sub

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCom.SetFocus
End Sub
