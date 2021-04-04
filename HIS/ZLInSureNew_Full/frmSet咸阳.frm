VERSION 5.00
Begin VB.Form frmSet咸阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmSet咸阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   2025
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4155
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1500
         Width           =   1635
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1485
         Left            =   3000
         TabIndex        =   9
         Top             =   330
         Width           =   1005
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
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "数据库(&D)"
         Height          =   180
         Index           =   3
         Left            =   390
         TabIndex        =   7
         Top             =   1560
         Width           =   810
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
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3090
      TabIndex        =   11
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1770
      TabIndex        =   10
      Top             =   2400
      Width           =   1100
   End
End
Attribute VB_Name = "frmSet咸阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
    Text医保数据库 = 3
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean

Private Sub cmdTest_Click()
    If gcn咸阳.State = adStateOpen Then gcn咸阳.Close
    
    On Error Resume Next
    
    gcn咸阳.Open "Provider=SQLOLEDB.1;Password=" & TxtEdit(Text医保密码).Tag & ";Persist Security Info=True;User ID=" & TxtEdit(text医保用户).Text & _
                ";Initial Catalog=" & TxtEdit(Text医保数据库).Text & ";Data Source=" & TxtEdit(Text医保服务器).Text
    
    If Err <> 0 Then
        MsgBox "医保前置服务器连接失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
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

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    
    For lngCount = TxtEdit.LBound To TxtEdit.UBound
        If zlCommFun.StrIsValid(TxtEdit(lngCount).Text, TxtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll TxtEdit(lngCount)
            TxtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If gcn咸阳.State = adStateClosed Then
        On Error Resume Next
        gcn咸阳.Open "Provider=SQLOLEDB.1;Password=" & TxtEdit(Text医保密码).Tag & ";Persist Security Info=True;User ID=" & TxtEdit(text医保用户).Text & _
                    ";Initial Catalog=" & TxtEdit(Text医保数据库).Text & ";Data Source=" & TxtEdit(Text医保服务器).Text
        
        If Err <> 0 Then
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
    gstrSQL = "zl_保险参数_Delete(" & TYPE_咸阳市 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_咸阳市 & ",null,'医保用户名','" & TxtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_咸阳市 & ",null,'医保用户密码','" & TxtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_咸阳市 & ",null,'医保服务器','" & TxtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_咸阳市 & ",null,'医保数据库','" & TxtEdit(Text医保数据库).Text & "',4)"
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
    
    '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
    If gcn咸阳.State = adStateOpen Then gcn咸阳.Close
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Public Function 参数设置() As Boolean
'功能：设置与东大阿尔派的医保接口
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    
    On Error GoTo errHandle
    
    
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_咸阳市)
    
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保用户名"
                TxtEdit(text医保用户) = str参数值
            Case "医保服务器"
                TxtEdit(Text医保服务器) = str参数值
            Case "医保数据库"
                TxtEdit(Text医保数据库) = str参数值
            Case "医保用户密码"
                TxtEdit(Text医保密码).Text = "        "    '假密码
                TxtEdit(Text医保密码).Tag = str参数值
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet咸阳.Show vbModal, frm医保类别
    
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
