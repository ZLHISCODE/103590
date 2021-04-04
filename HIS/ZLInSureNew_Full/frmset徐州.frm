VERSION 5.00
Begin VB.Form frmset徐州 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmset徐州.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraIC 
      Caption         =   "IC卡操作"
      Height          =   810
      Left            =   120
      TabIndex        =   8
      Top             =   1815
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   10
         Text            =   "1"
         Top             =   315
         Width           =   360
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
      TabIndex        =   12
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4950
      TabIndex        =   13
      Top             =   780
      Width           =   1100
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   14
         Top             =   1515
         Visible         =   0   'False
         Width           =   2145
      End
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
         Caption         =   "数据库(&D)"
         Height          =   180
         Index           =   5
         Left            =   300
         TabIndex        =   15
         Top             =   1575
         Visible         =   0   'False
         Width           =   810
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
End
Attribute VB_Name = "frmset徐州"
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
    If gcn徐州.State = adStateOpen Then gcn徐州.Close
    If Not IsNumeric(TxtEdit(4).Text) Then
        MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error Resume Next
'    gcn徐州.Open "Provider=SQLOLEDB.1;Initial Catalog=" & Trim(txtEdit(3).Tag) & ";Password=" & Trim(txtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(txtEdit(0).Tag) & ";Data Source=" & Trim(txtEdit(2).Tag)
    gcn徐州.Open "Provider=MSDASQL.1;Password=" & Trim(TxtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(TxtEdit(0).Text) & ";Data Source=" & Trim(TxtEdit(2).Text)
    
    If Err <> 0 Then
        MsgBox "医保前置服务器连接失败！", vbInformation, gstrSysName
        Exit Sub
    End If
'    If mblnChangePassword = True Then
'        '密码输入成功
'        txtEdit(4).Enabled = True
'    End If

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
    
    If Not IsNumeric(TxtEdit(4).Text) Then
        MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        Exit Function
    End If
    '对连接进行测试
    If gcn徐州.State = adStateClosed Then
        On Error Resume Next
'        gcn徐州.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
            Trim(txtEdit(2).Text), Trim(txtEdit(0).Text), Trim(txtEdit(1).Tag)
'        gcn徐州.Open "Provider=SQLOLEDB.1;Initial Catalog=" & Trim(txtEdit(3).Tag) & ";Password=" & Trim(txtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(txtEdit(0).Tag) & ";Data Source=" & Trim(txtEdit(2).Tag)
        gcn徐州.Open "Provider=MSDASQL.1;Password=" & Trim(TxtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(TxtEdit(0).Text) & ";Data Source=" & Trim(TxtEdit(2).Text)
        
        If Err <> 0 Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    On Error Resume Next
    mlngIcdev = init_com(TxtEdit(4).Text - 1) 'Init COM2
    If mlngIcdev <> 0 Then
        If MsgBox("串口初始化失败，请检查串口。是否继续保存？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            TxtEdit(4).SetFocus
            Exit Function
        End If
    End If
    st = close_com()
    IsValid = True
End Function

Public Function 参数设置() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    On Error GoTo errHandle
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_徐州)
    
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "徐州用户名"
                TxtEdit(0).Text = str参数值
            Case "徐州服务器"
                TxtEdit(2).Text = str参数值
            Case "徐州用户密码"
                TxtEdit(1).Text = "        "    '假密码
                TxtEdit(1).Tag = str参数值
            Case "徐州数据库"
                TxtEdit(3).Text = str参数值
        End Select
        rsTemp.MoveNext
    Loop
    On Error Resume Next
    TxtEdit(4).Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口") + 1
    
    mblnChange = False
    mblnChangePassword = False
    frmset徐州.Show vbModal, frm医保类别
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
    gstrSQL = "zl_保险参数_Delete(" & TYPE_徐州 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_徐州 & ",null,'徐州用户名','" & TxtEdit(0).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_徐州 & ",null,'徐州用户密码','" & TxtEdit(1).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_徐州 & ",null,'徐州服务器','" & TxtEdit(2).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_徐州 & ",null,'徐州数据库','" & TxtEdit(3).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    '将当前使用的串口写入注册表之中
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", CStr(TxtEdit(4).Text - 1)
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
    If gcn徐州.State = adStateOpen Then gcn徐州.Close
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
    If Index = 4 Then
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
    If Index = 4 Then
        If Not IsNumeric(TxtEdit(4).Text) Then
            MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        End If
    End If
End Sub
