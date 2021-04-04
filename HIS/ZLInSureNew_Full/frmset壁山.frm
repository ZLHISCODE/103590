VERSION 5.00
Begin VB.Form frmset壁山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmset壁山.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1440
      TabIndex        =   19
      Top             =   3510
      Width           =   3345
   End
   Begin VB.ComboBox cbo适用地区 
      Height          =   300
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3090
      Width           =   3345
   End
   Begin VB.Frame fraIC 
      Caption         =   "IC卡操作"
      Height          =   1245
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1320
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   780
         Width           =   990
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   10
         Text            =   "1"
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "卡验证码(&V)"
         Height          =   180
         Index           =   5
         Left            =   210
         TabIndex        =   12
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "当前串口(&D)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   375
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "号串口"
         Height          =   180
         Index           =   4
         Left            =   1740
         TabIndex        =   11
         Top             =   375
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4950
      TabIndex        =   16
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4950
      TabIndex        =   17
      Top             =   780
      Width           =   1100
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1545
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4695
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1110
         Left            =   3450
         TabIndex        =   7
         Top             =   330
         Width           =   1110
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1215
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   2145
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "配置文件位置"
      Height          =   180
      Left            =   315
      TabIndex        =   18
      Top             =   3585
      Width           =   1080
   End
   Begin VB.Label lbl适用地区 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用地区(&Q)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      TabIndex        =   14
      Top             =   3150
      Width           =   990
   End
End
Attribute VB_Name = "frmset壁山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mblnChangePassword As Boolean  '密码被修改过
Private mlngIcdev As Long
Private st%
Private Declare Function IC_InitComm Lib "DCIC32.DLL" (ByVal Port%) As Long
Private Declare Function IC_ExitComm% Lib "DCIC32.DLL" (ByVal icdev As Long)
 
Private Sub cbo适用地区_Change()
    If cbo适用地区.ListIndex = 1 Then
        Text1.Enabled = True
    Else
        Text1.Enabled = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Sub cmdTest_Click()
    If gcn壁山.State = adStateOpen Then gcn壁山.Close
    If Not IsNumeric(TxtEdit(3).Text) Then
        MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error Resume Next
    If cbo适用地区.ListIndex = 1 Then
        gcn壁山.Open "Provider=SQLOLEDB.1;Initial Catalog=hw_interface;Password=" & TxtEdit(1).Text & ";Persist Security Info=True;User ID=" & TxtEdit(0).Text & ";Data Source=" & TxtEdit(2).Text
    Else
        gcn壁山.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
                Trim(TxtEdit(2).Text), Trim(TxtEdit(0).Text), Trim(TxtEdit(1).Tag)
    End If
    
    If Err <> 0 Then
        MsgBox "医保前置服务器连接失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnChangePassword = True Then
        '密码输入成功
        TxtEdit(4).Enabled = True
    End If

    MsgBox "医保前置服务器连接成功", vbInformation, gstrSysName
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    '逐步判断字符的合法性
    For lngCount = TxtEdit.LBound To TxtEdit.UBound
        If zlCommFun.StrIsValid(TxtEdit(lngCount).Text, TxtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll TxtEdit(lngCount)
            TxtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    If TxtEdit(4).Enabled = True And TxtEdit(4).Text = "" Then
        MsgBox "请在询问IC卡供应商后，填写卡验证码。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not IsNumeric(TxtEdit(3).Text) Then
        MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        Exit Function
    End If
    '对连接进行测试
    If gcn壁山.State = adStateClosed Then
        On Error Resume Next
        gcn壁山.Open "Provider=SQLOLEDB.1;Initial Catalog=hw_interface;Password=" & TxtEdit(1).Text & ";Persist Security Info=True;User ID=" & TxtEdit(0).Text & ";Data Source=" & TxtEdit(2).Text
        
'          gcn壁山.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
              Trim(txtEdit(2).Text), Trim(txtEdit(0).Text), Trim(txtEdit(1).Tag)
        
        If Err <> 0 Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    On Error Resume Next
    mlngIcdev = IC_InitComm(TxtEdit(3).Text - 1) 'Init COM2
    If mlngIcdev <= 0 Then
        If MsgBox("串口初始化失败，请检查串口。是否继续保存？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            TxtEdit(3).SetFocus
            Exit Function
        End If
    End If
    st = IC_ExitComm(mlngIcdev)  'Close COM
    IsValid = True
End Function

Public Function 参数设置() As Boolean
'功能：设置与西希公司的医保接口
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    Dim int适用地区 As Integer
    
    mblnOK = False
    On Error GoTo errHandle
    
    
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆壁山)
    
    int适用地区 = 0
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "壁山用户名"
                TxtEdit(0).Text = str参数值
            Case "壁山服务器"
                TxtEdit(2).Text = str参数值
            Case "壁山用户密码"
                TxtEdit(1).Text = "        "    '假密码
                TxtEdit(1).Tag = str参数值
            Case "卡验证码"
                TxtEdit(4).Text = str参数值
            Case "适用地区"
                int适用地区 = Val(str参数值)
            Case "配置文件位置"
                Text1.Text = str参数值
        End Select
        rsTemp.MoveNext
    Loop
    If TxtEdit(4).Text = "" Then TxtEdit(4).Enabled = True
    On Error Resume Next
'    If GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口") = "" Then
    TxtEdit(3).Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口") + 1
    
    'Modified By 朱玉宝 下午 06:07:34
    With cbo适用地区
        .Clear
        .AddItem "壁山"
        .AddItem "黔江"
        .AddItem "西彭"
        .AddItem "秀山"
        .ListIndex = int适用地区
    End With
    
    mblnChange = False
    mblnChangePassword = False
    frmset壁山.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_重庆壁山 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆壁山 & ",null,'壁山用户名','" & TxtEdit(0).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆壁山 & ",null,'壁山用户密码','" & TxtEdit(1).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆壁山 & ",null,'壁山服务器','" & TxtEdit(2).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆壁山 & ",null,'卡验证码','" & TxtEdit(4).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    'Modified By 朱玉宝 下午 06:07:51
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆壁山 & ",null,'适用地区','" & cbo适用地区.ListIndex & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆壁山 & ",null,'配置文件位置','" & Text1.Text & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    gcnOracle.CommitTrans
    '将当前使用的串口写入注册表之中
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", CStr(TxtEdit(3).Text - 1)
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 Then
        TxtEdit(1).Tag = TxtEdit(1).Text
        mblnChangePassword = True
    End If
    
    '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
    If gcn壁山.State = adStateOpen Then gcn壁山.Close
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
    If Index = 3 Then
        If CheckIsInclude(UCase(Chr(KeyAscii)), "正整数") = True Then KeyAscii = 0
    End If
End Sub

Private Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "日期"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "日期时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "可打印字符"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Index = 3 Then
        If Not IsNumeric(TxtEdit(3).Text) Then
            MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        End If
    End If
End Sub
