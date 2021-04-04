VERSION 5.00
Begin VB.Form frmSet自贡 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险参数设置"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmSet自贡.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4170
      TabIndex        =   8
      Top             =   420
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4170
      TabIndex        =   9
      Top             =   870
      Width           =   1100
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医保服务器"
      Height          =   1665
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3765
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   390
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   780
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1170
         Width           =   1305
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   2610
         TabIndex        =   7
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   8
         Left            =   390
         TabIndex        =   1
         Top             =   450
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   9
         Left            =   570
         TabIndex        =   3
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   10
         Left            =   390
         TabIndex        =   5
         Top             =   1230
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmSet自贡"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum文本
    text医保用户 = 0
    Text医保密码
    Text医保服务器
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private mlng险类 As Long

Private Sub cmb级别_Click()
    mblnChange = True
End Sub

Private Sub cmdTest_Click()
    If OraDataOpen(gcn自贡, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
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
    
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保用户名','" & TxtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保用户密码','" & TxtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保服务器','" & TxtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If gcn自贡.State = adStateClosed Then
        If OraDataOpen(gcn自贡, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag, False) = False Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
        End If
    End If
    
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
        If gcn自贡.State = adStateOpen Then gcn自贡.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Public Function 参数设置(ByVal lng险类 As Long, ByVal lng中心 As Long) As Boolean
'功能：设置与东大阿尔派的医保接口
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    mlng险类 = lng险类
    
    On Error GoTo errHandle
    
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng险类)
    
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保用户名"
                TxtEdit(text医保用户) = str参数值
            Case "医保服务器"
                TxtEdit(Text医保服务器) = str参数值
            Case "医保用户密码"
                TxtEdit(Text医保密码).Text = str参数值
                TxtEdit(Text医保密码).Tag = str参数值
        End Select
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet自贡.Show vbModal, frm医保类别
    
    参数设置 = mblnOK
    If mblnOK = False Then
        '保存失败，关闭连接。以免用用户输入的其它用户名进行了连接
        If gcn自贡.State = adStateOpen Then gcn自贡.Close
    End If
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
