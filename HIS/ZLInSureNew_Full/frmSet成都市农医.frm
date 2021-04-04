VERSION 5.00
Begin VB.Form frmSet成都市农医 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmSet成都市农医.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frm高级设置 
      Caption         =   "高级设置"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   4515
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   1  'ON
         Index           =   3
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   1  'ON
         Index           =   4
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   12
         Top             =   650
         Width           =   735
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   1  'ON
         Index           =   5
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   11
         Top             =   1050
         Width           =   3135
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "医院名称"
         Height          =   180
         Index           =   3
         Left            =   435
         TabIndex        =   16
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "区县编码"
         Height          =   180
         Index           =   4
         Left            =   435
         TabIndex        =   15
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "区县名称"
         Height          =   180
         Index           =   5
         Left            =   435
         TabIndex        =   14
         Top             =   1125
         Width           =   720
      End
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1605
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   4515
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   3330
         TabIndex        =   6
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1110
         Width           =   1935
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
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   3
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   9
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   8
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   7
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3390
      TabIndex        =   1
      Top             =   3405
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2205
      TabIndex        =   0
      Top             =   3405
      Width           =   1100
   End
End
Attribute VB_Name = "frmSet成都市农医"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
    Text医院名称 = 3
    Text区县编码 = 4
    Text区县名称 = 5
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Dim mblnTest As Boolean
Dim mcnTest As New ADODB.Connection

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    If Not mblnTest Then MsgBox "连接成功！", vbInformation, gstrSysName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    ElseIf KeyAscii = 39 Then
        KeyAscii = 0
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

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = TxtEdit.LBound To TxtEdit.UBound
        If zlCommFun.StrIsValid(TxtEdit(lngCount).Text, TxtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll TxtEdit(lngCount)
            TxtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag, False) = False Then
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
    gstrSQL = "zl_保险参数_Delete(" & TYPE_成都市农医 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都市农医 & ",null,'医保用户名','" & TxtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都市农医 & ",null,'医保用户密码','" & TxtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都市农医 & ",null,'医保服务器','" & TxtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都市农医 & ",null,'医院名称','" & TxtEdit(Text医院名称).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都市农医 & ",null,'区县编码','" & TxtEdit(Text区县编码).Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都市农医 & ",null,'区县名称','" & TxtEdit(Text区县名称).Text & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    
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
    If Index = Text医保密码 Then
        TxtEdit(Index).Tag = TxtEdit(Index).Text
    End If
    
    If Index = Text医保服务器 Or Index = Text医保密码 Or Index = text医保用户 Then
        '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Public Function 参数设置() As Boolean
'功能：设置与东大阿尔派的医保接口
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    Dim str医院名称 As String
    
    mblnOK = False
    
    On Error GoTo errHandle
    
    '取保险参数
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_成都市农医)
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保用户名"
                TxtEdit(text医保用户) = str参数值
            Case "医保服务器"
                TxtEdit(Text医保服务器) = str参数值
            Case "医保用户密码"
                TxtEdit(Text医保密码).Text = "        "    '假密码
                TxtEdit(Text医保密码).Tag = str参数值
            Case "医院名称"
                TxtEdit(Text医院名称).Text = str参数值
            Case "区县编码"
                TxtEdit(Text区县编码).Text = str参数值
            Case "区县名称"
                TxtEdit(Text区县名称).Text = str参数值
        End Select
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet成都市农医.Show vbModal, frm医保类别
    
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


