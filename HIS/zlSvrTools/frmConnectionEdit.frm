VERSION 5.00
Begin VB.Form frmConnectionEdit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "frmConnectionEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtLinkName 
      Height          =   300
      Left            =   1065
      MaxLength       =   30
      TabIndex        =   0
      Top             =   345
      Width           =   1725
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2445
      MaxLength       =   3
      TabIndex        =   5
      Tag             =   "IP地址"
      Top             =   945
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1995
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "IP地址"
      Top             =   945
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1545
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "IP地址"
      Top             =   945
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1095
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "IP地址"
      Top             =   945
      Width           =   315
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Left            =   1065
      MaxLength       =   20
      TabIndex        =   25
      Tag             =   "IP"
      Text            =   "   ．   ．   ．"
      Top             =   900
      Width           =   1725
   End
   Begin VB.TextBox txtNotes 
      Height          =   1320
      Left            =   1065
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1965
      Width           =   4695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4665
      TabIndex        =   13
      Top             =   3885
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试(&T)"
      Height          =   350
      Left            =   420
      TabIndex        =   11
      Top             =   3885
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3570
      TabIndex        =   12
      Top             =   3885
      Width           =   1100
   End
   Begin VB.Frame fraMain 
      Height          =   30
      Index           =   1
      Left            =   -195
      TabIndex        =   18
      Top             =   3540
      Width           =   6570
   End
   Begin VB.TextBox txtDatabase 
      Height          =   300
      Left            =   4260
      MaxLength       =   50
      TabIndex        =   1
      Top             =   345
      Width           =   1500
   End
   Begin VB.TextBox txtPort 
      Height          =   300
      Left            =   4260
      MaxLength       =   5
      TabIndex        =   7
      Top             =   885
      Width           =   1500
   End
   Begin VB.TextBox txtPasswd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4260
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1425
      Width           =   1500
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1065
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1425
      Width           =   1725
   End
   Begin VB.Label lblLinkName 
      AutoSize        =   -1  'True
      Caption         =   "连接名称"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   405
      Width           =   720
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   2820
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblNotes 
      AutoSize        =   -1  'True
      Caption         =   "说  明"
      Height          =   180
      Left            =   420
      TabIndex        =   24
      Top             =   1965
      Width           =   540
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   5790
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   5790
      TabIndex        =   22
      Top             =   780
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   2820
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   5790
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   2820
      TabIndex        =   19
      Top             =   765
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      Caption         =   "实例名"
      Height          =   180
      Left            =   3585
      TabIndex        =   17
      Top             =   405
      Width           =   540
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "端口号"
      Height          =   180
      Left            =   3585
      TabIndex        =   16
      Top             =   945
      Width           =   540
   End
   Begin VB.Label lblIp 
      AutoSize        =   -1  'True
      Caption         =   "IP地址"
      Height          =   180
      Left            =   420
      TabIndex        =   15
      Top             =   945
      Width           =   540
   End
   Begin VB.Label lblPasswd 
      AutoSize        =   -1  'True
      Caption         =   "密  码"
      Height          =   180
      Left            =   3585
      TabIndex        =   14
      Top             =   1485
      Width           =   540
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   420
      TabIndex        =   6
      Top             =   1485
      Width           =   540
   End
End
Attribute VB_Name = "frmConnectionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'入口变量
Private mstrLinkName As String '连接名称
Private mstrUser As String  '用户名
Private mstrPasswd As String  '密码
Private mstrDatabase As String  '数据库实例名
Private mstrNotes As String  '说明
Private mStrIP As String  'IP地址
Private mstrPort As String  '端口号
Private mlngId As Long  '连接编号，当传入为0时表示新增连接
Private mblnState As Boolean  '标记是否进行了连接测试
Private mblnSaveClick As Boolean  '标记是否保存
Private mclsCiph As clsCipher  '定义一个加解密实例化对象

Private Const lng_新增 As Long = 0

Private Enum CheckTag
    CT_连接名称 = 0
    CT_实例名 = 1
    CT_IP地址 = 2
    CT_端口号 = 3
    CT_用户名 = 4
    CT_密码 = 5
End Enum

Public Function ShowEdit(lngID As Long, strLinkName As String, strUser As String, strPasswd As String, strIp As String, _
                        strPort As String, strDatabase As String, strNotes As String) As Boolean
    '-------------------------------------------------------------------------------
    '--功能：显示和编辑连接信息
    '--参数：strUser:用户名, strIp:服务器IP, strDatabase:实例名, strNotes:备注说明, strCaption:标题, lngId:编号
    '-------------------------------------------------------------------------------
    Set mclsCiph = New clsCipher
    mstrLinkName = strLinkName
    mlngId = lngID
    mStrIP = strIp
    mstrPort = strPort
    mstrUser = strUser
    mstrPasswd = mclsCiph.Decipher(MSTR_DBLINK_KEY, strPasswd)
    mstrDatabase = strDatabase
    mstrNotes = strNotes
    
    Me.Caption = IIf(lngID = lng_新增, "新增连接配置", "修改连接配置")
    Me.Show vbModal, frmMDIMain
    
    If mblnSaveClick Then
        lngID = mlngId
        strUser = mstrUser
        strIp = mStrIP
        strPort = mstrPort
        strDatabase = mstrDatabase
        strNotes = mstrNotes
        strPasswd = mstrPasswd
        strLinkName = mstrLinkName
    End If
    ShowEdit = mblnSaveClick
    '清空数据
    Call ClearDate
End Function

Private Sub FillData(ByVal strLinkName As String, ByVal strIps As String, ByVal strPort As String, _
                    ByVal strUser As String, ByVal strPasswd As String, ByVal strDatabase As String, ByVal strNotes As String)
    '-------------------------------------------------------------------------------
    '--功能：当要修改数据时，将传入的数据显示在对应位置
    '-------------------------------------------------------------------------------
    Dim strIp() As String
    Dim i As Long
    
    On Error GoTo errH:
    strIp = Split(strIps, ".")
    For i = 0 To 3
        txtIp(i) = strIp(i)
    Next
    txtLinkName = strLinkName
    txtPort.Text = strPort
    txtDatabase.Text = strDatabase
    txtUser.Text = strUser
    txtPasswd.Text = strPasswd
    txtNotes.Text = strNotes
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub cmdCancel_Click()
    mblnSaveClick = False
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim rsID As ADODB.Recordset
    Dim strSQL As String
    Dim strNote As String
    
    On Error GoTo errH:
    '初始化标记，防止在点完测试之后又去修改信息，导致保存时发生错误
    mblnState = False
    mblnSaveClick = True
    Call cmdTest_Click
    '将数据加入数据库
    If mblnState Then
        Set mclsCiph = New clsCipher
        If mlngId = lng_新增 Then
            strSQL = "Zl_Zlconnections_Edit(0,Null,'" & Trim(txtLinkName.Text) & "','" & Trim(txtUser.Text) & "','" & _
                                            mclsCiph.Cipher(MSTR_DBLINK_KEY, txtPasswd.Text) & "','" & _
                                            txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & _
                                            "." & txtIp(3).Text & "'," & txtPort.Text & ",'" & txtDatabase.Text & _
                                            "','" & txtNotes.Text & "')"
            Call ExecuteProcedure(strSQL, Me.Caption)
            strSQL = "Select Max(编号) As 编号 From Zlconnections"
            Set rsID = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "查询最大ID")
            mlngId = rsID!编号
            '插入重要操作日志
            Call SaveAuditLog(1, "新增", "添加连接“" & txtLinkName.Text & "”")
        Else
            strSQL = "Zl_Zlconnections_Edit(1," & mlngId & ",'" & Trim(txtLinkName.Text) & "','" & Trim(txtUser.Text) & _
                                            "','" & mclsCiph.Cipher(MSTR_DBLINK_KEY, txtPasswd.Text) & "','" & _
                                            txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & _
                                            "." & txtIp(3).Text & "'," & txtPort.Text & ",'" & txtDatabase.Text & _
                                            "','" & txtNotes.Text & "')"
            Call ExecuteProcedure(strSQL, Me.Caption)
            If mstrLinkName <> Trim(txtLinkName.Text) Then strNote = ",名称由“" & mstrLinkName & "”修改为“" & Trim(txtLinkName.Text) & "”"
            If mstrDatabase <> txtDatabase.Text Then strNote = strNote & ",实例名由" & mstrDatabase & "修改为" & txtDatabase.Text
            If mStrIP <> txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & "." & txtIp(3).Text Then
                strNote = strNote & ",IP地址由" & mStrIP & "修改为" & txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & "." & txtIp(3).Text
            End If
            If mstrPort <> txtPort.Text Then strNote = strNote & ",端口号由" & mstrPort & "修改为" & txtPort.Text
            If mstrUser <> Trim(txtUser.Text) Then strNote = strNote & ",用户名由" & mstrUser & "修改为" & Trim(txtUser.Text)
            '插入重要操作日志
            If strNote <> "" Then
                Call SaveAuditLog(2, "修改", "修改连接“" & mstrLinkName & "”" & strNote)
            End If
        End If
        
        mstrLinkName = Trim(txtLinkName.Text)
        mstrUser = Trim(txtUser.Text)
        mStrIP = txtIp(0).Text & "." & txtIp(1).Text & "." & txtIp(2).Text & "." & txtIp(3).Text
        mstrPort = txtPort.Text
        mstrDatabase = txtDatabase.Text
        mstrNotes = txtNotes.Text
        mstrPasswd = mclsCiph.Cipher(MSTR_DBLINK_KEY, txtPasswd.Text)
        Unload Me
    Else
        mblnSaveClick = False
    End If
    Exit Sub
errH:
    mblnSaveClick = False
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub cmdTest_Click()
    Dim cnOracle As ADODB.Connection
    Dim strServerName As String
    
    mstrUser = Trim(txtUser.Text)
    mstrPasswd = txtPasswd.Text
    If CheckData Then
        'strServerName = "192.168.2.13:1521/dyyy"
        strServerName = Val(txtIp(0).Text) & "." & Val(txtIp(1).Text) & "." & Val(txtIp(2).Text) & "." & Val(txtIp(3).Text) & _
                        ":" & Val(txtPort.Text) & "/" & Trim(txtDatabase.Text)
        Set cnOracle = gobjRegister.GetConnection(strServerName, mstrUser, mstrPasswd, False, OraOLEDB, , False)
        If cnOracle.State = adStateOpen Then
            mblnState = True
            If mblnSaveClick = False Then MsgBox "连接可用！", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Function CheckData() As Boolean
    '检查数据是否填写完整
    Dim blnDataType As Boolean
    Dim i As Integer
    
    '检查IP地址
    For i = 0 To 3
        If txtIp(i).Text = "" Then
            blnDataType = False
        Else
            blnDataType = True
        End If
    Next
    If blnDataType = False Then
        lblCheck(CT_IP地址).Visible = True
    Else
        lblCheck(CT_IP地址).Visible = False
    End If
    
    '检查连接名称
    If txtLinkName.Text = "" Then
        lblCheck(CT_连接名称).Visible = True
    Else
        lblCheck(CT_连接名称).Visible = False
    End If
    
    '检查端口
    If txtPort.Text = "" Then
        lblCheck(CT_端口号).Visible = True
    Else
        lblCheck(CT_端口号).Visible = False
    End If
    
    '检查实例名
    If txtDatabase.Text = "" Then
        lblCheck(CT_实例名).Visible = True
    Else
        lblCheck(CT_实例名).Visible = False
    End If
    
    '检查用户名
    If txtUser.Text = "" Then
        lblCheck(CT_用户名).Visible = True
    Else
        lblCheck(CT_用户名).Visible = False
    End If
    
    '检查密码
    If txtPasswd.Text = "" Then
        lblCheck(CT_密码).Visible = True
    Else
        lblCheck(CT_密码).Visible = False
    End If
    
    '检查连接名称合法性
    If InStr(1, txtLinkName.Text, "'") <> 0 Then
        MsgBox "【连接名称】中不能输入单引号!", vbInformation + vbOKOnly, gstrSysName
        txtLinkName.SetFocus
        Exit Function
    End If
    
    '检查说明合法性
    If InStr(1, txtNotes.Text, "'") <> 0 Then
        MsgBox "【说明】中不能输入单引号!", vbInformation + vbOKOnly, gstrSysName
        txtNotes.SetFocus
        Exit Function
    End If
    For i = 0 To 5
        If lblCheck(i).Visible = True Then
            CheckData = False
            Select Case i
                Case CT_连接名称
                    txtLinkName.SetFocus
                Case CT_IP地址
                    txtIp(0).SetFocus
                Case CT_端口号
                    txtPort.SetFocus
                Case CT_实例名
                    txtDatabase.SetFocus
                Case CT_用户名
                    txtUser.SetFocus
                Case CT_密码
                    txtPasswd.SetFocus
            End Select
            MsgBox "请将信息填写完整！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        Else
            CheckData = True
        End If
    Next
End Function

Private Sub ClearDate()
    mstrUser = ""
    mstrPasswd = ""
    mstrDatabase = ""
    mstrNotes = ""
    mStrIP = ""
    mblnState = False
    mblnSaveClick = False
End Sub

Private Sub Form_Load()
    '填充数据
    If mStrIP <> "" Then
        Call FillData(mstrLinkName, mStrIP, mstrPort, mstrUser, mstrPasswd, mstrDatabase, mstrNotes)
    End If
End Sub

Private Sub txtDatabase_KeyPress(KeyAscii As Integer)
    '只能输入大小写字母或数字
    If Not ((KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtIp_Change(Index As Integer)
    Dim lngLineNo As Long '行号
    Dim lngColNo As Long  '列号
    
    Call GetCursorPos(Me.txtIp(Index).hwnd, lngLineNo, lngColNo)
    If lngColNo > 3 Then
        If Index < 3 Then
            If txtIp(Index + 1).Enabled Then txtIp(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtIp_GotFocus(Index As Integer)
    txtIp(Index).SelStart = 0
    txtIp(Index).SelLength = Len(txtIp(Index).Text)
End Sub

Private Sub txtIp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lngLineNo As Long '行号
    Dim lngColNo As Long  '列号
    err = 0
    On Error Resume Next

    Call GetCursorPos(Me.txtIp(Index).hwnd, lngLineNo, lngColNo)

    Select Case KeyCode
    Case 37     '<-

        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtIp(Index - 1).Enabled Then
                txtIp(Index - 1).SelStart = Len(txtIp(Index - 1))
                txtIp(Index - 1).SetFocus
            End If
        End If
    Case 39     '->
        If Index < 3 Then
            If lngColNo <= Len(txtIp(Index)) Then Exit Sub
            If txtIp(Index + 1).Enabled Then
                txtIp(Index + 1).SelStart = 0
                txtIp(Index + 1).SetFocus
            End If
        End If
    Case 8     'BACKSPACE
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtIp(Index - 1).Enabled Then
                txtIp(Index - 1).SelStart = Len(txtIp(Index - 1))
                txtIp(Index - 1).SetFocus
            End If
        End If
    Case Else
    End Select

End Sub

Private Sub txtIp_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
        If Chr(KeyAscii) = "." Then
            If Index < 3 And Index >= 0 And Trim(txtIp(Index)) <> "" Then
                If txtIp(Index + 1).Enabled Then txtIp(Index + 1).SetFocus
            End If
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Public Sub GetCursorPos(ByVal hwnd5 As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long
    Dim lParam As Long, wParam As Long
    Dim K As Long
    
    i = SendMessage(hwnd5, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16 '取得目前光标所在位置前有多少个Byte
    LineNo = SendMessage(hwnd5, EM_LINEFROMCHAR, j, 0) '取得光标前面有多少行
    LineNo = LineNo + 1
    K = SendMessage(hwnd5, EM_LINEINDEX, -1, 0)
    '取得目前光标所在行前面有多少个Byte
    ColNo = j - K + 1
End Sub

Private Sub txtIp_LostFocus(Index As Integer)
    If txtIp(Index).Text = "" Then Exit Sub
    Select Case Index
    Case 0
        If Val(txtIp(Index).Text) < 1 Or Val(txtIp(Index).Text) > 233 Then
            MsgBox "【" & txtIp(Index).Text & "】不是有效项。请指定一个介于1和233间的值", vbOKOnly + vbInformation, gstrSysName
            txtIp(Index).SetFocus
            txtIp(Index).Text = 233
        End If
    Case 1, 2, 3
        If (Not IsNumeric(txtIp(Index).Text)) Or Val(txtIp(Index).Text) > 255 Then
            MsgBox "【" & txtIp(Index).Text & "】不是有效项。请指定一个介于0和255间的值", vbOKOnly + vbInformation, gstrSysName
            txtIp(Index).SetFocus
            txtIp(Index).Text = 255
        End If
    End Select
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
    If ActualLen(txtNotes.Text) >= 500 Then
        KeyAscii = 0
        MsgBox "最大输入长度为500个字符(250个字)！", vbOKOnly + vbInformation, gstrSysName
    End If
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    '只能输入大小写字母或数字
    If Not ((KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If Not (InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPort_LostFocus()
    If txtPort.Text = "" Then Exit Sub
    If Not IsNumeric(txtPort.Text) Then
        MsgBox "【" & txtPort.Text & "】不是有效项。请输入正确的端口号！", vbOKOnly + vbInformation, gstrSysName
        txtPort.SetFocus
        txtPort.Text = mstrPort
    End If
    
End Sub
