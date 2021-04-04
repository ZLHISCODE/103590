VERSION 5.00
Begin VB.Form frmSet松藻 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险参数设置"
   ClientHeight    =   3270
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5745
   Icon            =   "frmSet松藻.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4515
      TabIndex        =   10
      Top             =   990
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4515
      TabIndex        =   11
      Top             =   1470
      Width           =   1100
   End
   Begin VB.Frame fraTop 
      Caption         =   "运行参数"
      Height          =   2400
      Left            =   810
      TabIndex        =   1
      Top             =   690
      Width           =   3525
      Begin VB.CheckBox chk 
         Caption         =   "允许不设置医保项目(&G)"
         Height          =   195
         Index           =   1
         Left            =   570
         TabIndex        =   9
         Top             =   1950
         Width           =   2265
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1485
         Width           =   645
      End
      Begin VB.CheckBox chk 
         Caption         =   "需要密码验证(&V)"
         Height          =   195
         Index           =   0
         Left            =   570
         TabIndex        =   6
         Top             =   1170
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   5
         Top             =   735
         Width           =   645
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "医保卡号长度(&R)"
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   2
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "密码长度(&P)"
         Height          =   180
         Index           =   1
         Left            =   930
         TabIndex        =   7
         Top             =   1545
         Width           =   990
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "退休证号长度(&T)"
         Height          =   180
         Index           =   2
         Left            =   570
         TabIndex        =   4
         Top             =   795
         Width           =   1350
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmSet松藻.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "当选择指定保险类别，程序将按参数要求运行。"
      Height          =   180
      Left            =   795
      TabIndex        =   0
      Top             =   240
      Width           =   3780
   End
End
Attribute VB_Name = "frmSet松藻"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum编辑
    Text卡号长度 = 0
    Text密码长度 = 1
    Text退休证号 = 2
End Enum

Private Enum enum选择
    Check密码 = 0
    Check允许不设置医保项目 = 1
End Enum

Dim mlng险类 As Long, mlng中心 As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

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
    
    If Val(TxtEdit(Text卡号长度).Text) > 26 Then
        MsgBox "卡号长度不能超过26位。", vbInformation, gstrSysName
        TxtEdit(Text卡号长度).SetFocus
        Exit Sub
    End If
    
    If Val(TxtEdit(Text密码长度).Text) > 8 Then
        MsgBox "密码长度不能超过8位。", vbInformation, gstrSysName
        TxtEdit(Text密码长度).SetFocus
        Exit Sub
    End If
    If Val(TxtEdit(Text退休证号).Text) > 26 Then
        MsgBox "退休证号长度不能超过26位。", vbInformation, gstrSysName
        TxtEdit(Text退休证号).SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & "," & mlng中心 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    colPara.Add mlng中心 & ",'卡号长度','" & Int(Val(TxtEdit(Text卡号长度).Text))
    colPara.Add mlng中心 & ",'退休证长度','" & Int(Val(TxtEdit(Text退休证号).Text))
    colPara.Add mlng中心 & ",'密码验证','" & chk(Check密码).Value
    colPara.Add mlng中心 & ",'密码长度','" & IIf(chk(Check密码).Value = 1, Int(Val(TxtEdit(Text密码长度).Text)), 0)
    '这一部分参数不区分中心
    colPara.Add "null,'允许不设置医保项目','" & chk(Check允许不设置医保项目).Value
    
    For lngCount = 1 To colPara.Count
        gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & colPara(lngCount) & "'," & lngCount & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub chk_Click(Index As Integer)
    mblnChange = True
    If Index = Check密码 Then
        TxtEdit(Text密码长度).Enabled = chk(Check密码).Value
        lblEdit(Text密码长度).Enabled = chk(Check密码).Value
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
    zlCommFun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Public Function 参数设置(ByVal lng险类 As Long, ByVal lng中心 As Long) As Boolean
'功能：设置我们中联医保所需要的参数
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    mlng险类 = lng险类
    mlng中心 = lng中心
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1] and (中心 is null or 中心=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng险类, lng中心)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "卡号长度"
                TxtEdit(Text卡号长度).Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "退休证长度"
                TxtEdit(Text退休证号).Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "密码验证"
                chk(Check密码).Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "允许不设置医保项目"
                chk(Check允许不设置医保项目).Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "密码长度"
                TxtEdit(Text密码长度).Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet松藻.Show vbModal, frm医保类别
    参数设置 = mblnOK
End Function

