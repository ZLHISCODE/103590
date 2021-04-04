VERSION 5.00
Begin VB.Form frmSet中联 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险参数设置"
   ClientHeight    =   3870
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   Icon            =   "frmSet中联.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3285
      TabIndex        =   19
      Top             =   3390
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4575
      TabIndex        =   20
      Top             =   3390
      Width           =   1100
   End
   Begin VB.Frame fraButtom 
      Caption         =   "个人帐户支付范围"
      Height          =   2550
      Left            =   2940
      TabIndex        =   11
      Top             =   690
      Width           =   2925
      Begin VB.CheckBox chk 
         Caption         =   "全自费部分(&A)"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   13
         Top             =   590
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "首先自付部分(&F)"
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   14
         Top             =   925
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Caption         =   "全自费部分(&L)"
         Height          =   255
         Index           =   3
         Left            =   660
         TabIndex        =   16
         Top             =   1520
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "首先自付部分(&I)"
         Height          =   255
         Index           =   4
         Left            =   660
         TabIndex        =   17
         Top             =   1855
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Caption         =   "超限部分(&V)"
         Height          =   255
         Index           =   5
         Left            =   660
         TabIndex        =   18
         Top             =   2190
         Width           =   1365
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "收费时的使用范围"
         Height          =   180
         Index           =   2
         Left            =   270
         TabIndex        =   12
         Top             =   330
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "结算时的使用范围"
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   15
         Top             =   1260
         Width           =   1440
      End
   End
   Begin VB.Frame fraTop 
      Caption         =   "运行参数"
      Height          =   2550
      Left            =   180
      TabIndex        =   1
      Top             =   690
      Width           =   2565
      Begin VB.CheckBox chk97 
         Caption         =   "超过门诊限额不交易"
         Height          =   225
         Left            =   345
         TabIndex        =   23
         Top             =   2130
         Width           =   1965
      End
      Begin VB.TextBox txt97 
         Height          =   270
         Left            =   1110
         TabIndex        =   21
         Top             =   1800
         Width           =   780
      End
      Begin VB.CheckBox chk 
         Caption         =   "先扣起付线"
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   10
         Top             =   1500
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         Caption         =   "收费使用医保基金(&G)"
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   9
         Top             =   1170
         Width           =   2025
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1485
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox chk 
         Caption         =   "需要密码验证(&V)"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   1170
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   5
         Top             =   735
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "元"
         Height          =   180
         Left            =   2025
         TabIndex        =   24
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "门诊限额"
         Height          =   285
         Left            =   315
         TabIndex        =   22
         Top             =   1830
         Width           =   750
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "医保卡号长度(&R)"
         Height          =   180
         Index           =   0
         Left            =   300
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
         Left            =   660
         TabIndex        =   7
         Top             =   1545
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "退休证号长度(&T)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   4
         Top             =   795
         Width           =   1350
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmSet中联.frx":000C
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
Attribute VB_Name = "frmSet中联"
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
    Check收费全自费 = 1
    Check收费首先自付 = 2
    Check结算全自费 = 3
    Check结算首先自付 = 4
    Check结算超限 = 5
    Check收费医保基金 = 6
    Check先扣起付线 = 7
End Enum

Dim mlng险类 As Long, mlng中心 As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub Form_Load()
    '20051220 陈东 丰县医院增加
    If mlng险类 = TYPE_徐州六院 Then
        txt97.Visible = True
        chk97.Visible = True
        Label1.Visible = True
        Label2.Visible = True
    Else
        txt97.Visible = False
        chk97.Visible = False
        Label1.Visible = False
        Label2.Visible = False
    End If
End Sub

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
    
    If Val(txtEdit(Text卡号长度).Text) > 26 Then
        MsgBox "卡号长度不能超过26位。", vbInformation, gstrSysName
        txtEdit(Text卡号长度).SetFocus
        Exit Sub
    End If
    
    If Val(txtEdit(Text密码长度).Text) > 8 Then
        MsgBox "密码长度不能超过8位。", vbInformation, gstrSysName
        txtEdit(Text密码长度).SetFocus
        Exit Sub
    End If
    If Val(txtEdit(Text退休证号).Text) > 26 Then
        MsgBox "退休证号长度不能超过26位。", vbInformation, gstrSysName
        txtEdit(Text退休证号).SetFocus
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
    colPara.Add mlng中心 & ",'卡号长度','" & Int(Val(txtEdit(Text卡号长度).Text))
    colPara.Add mlng中心 & ",'退休证长度','" & Int(Val(txtEdit(Text退休证号).Text))
    colPara.Add mlng中心 & ",'密码验证','" & chk(Check密码).Value
    colPara.Add mlng中心 & ",'密码长度','" & IIf(chk(Check密码).Value = 1, Int(Val(txtEdit(Text密码长度).Text)), 0)
    '这一部分参数不区分中心
    colPara.Add "null,'收费使用医保基金','" & chk(Check收费医保基金).Value
    colPara.Add "null,'收费个人帐户使用范围','" & _
                chk(Check收费全自费).Value & chk(Check收费首先自付).Value
    colPara.Add "null,'结算个人帐户使用范围','" & _
                chk(Check结算全自费).Value & chk(Check结算首先自付).Value & chk(Check结算超限).Value
    colPara.Add "null,'先扣起付线','" & chk(Check先扣起付线).Value
    
    For lngCount = 1 To colPara.Count
        gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & colPara(lngCount) & "'," & lngCount & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    '20051220 陈东 丰县医院增加
    If mlng险类 = TYPE_徐州六院 Then
        If Val(txt97) > 0 Then
            gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",0,'门诊限额','" & Val(txt97) & "'," & lngCount + 1 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        
        If chk97.Value = 1 Then
            gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",0,'超额禁止','" & 1 & "'," & lngCount + 2 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End If
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
        txtEdit(Text密码长度).Enabled = chk(Check密码).Value
        lblEdit(Text密码长度).Enabled = chk(Check密码).Value
    End If
    If Index = Check收费全自费 Or Index = Check结算全自费 Then
        If chk(Index).Value = 1 Then
            chk(Index + 1).Value = 1
            chk(Index + 1).Enabled = False
        Else
            chk(Index + 1).Enabled = True
        End If
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
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
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1] and (中心 is null or 中心=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng险类, lng中心)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "卡号长度"
                txtEdit(Text卡号长度).Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "退休证长度"
                txtEdit(Text退休证号).Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "密码验证"
                chk(Check密码).Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "收费使用医保基金"
                chk(Check收费医保基金).Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "密码长度"
                txtEdit(Text密码长度).Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "收费个人帐户使用范围"
                str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                chk(Check收费全自费).Value = IIf(Left(str参数值, 1) = "1", 1, 0)
                chk(Check收费首先自付).Value = IIf(Mid(str参数值, 2, 1) = "1", 1, 0)
                '全自费优先
                If chk(Check收费全自费).Value = 1 Then
                    chk(Check收费首先自付).Value = 1
                    chk(Check收费首先自付).Enabled = False
                End If
            Case "结算个人帐户使用范围"
                str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                chk(Check结算全自费).Value = IIf(Left(str参数值, 1) = "1", 1, 0)
                chk(Check结算首先自付).Value = IIf(Mid(str参数值, 2, 1) = "1", 1, 0)
                chk(Check结算超限).Value = IIf(Mid(str参数值, 3, 1) = "1", 1, 0)
                '全自费优先
                If chk(Check结算全自费).Value = 1 Then
                    chk(Check结算首先自付).Value = 1
                    chk(Check结算首先自付).Enabled = False
                End If
            Case "先扣起付线"
                str参数值 = IIf(IsNull(rsTemp("参数值")), "0", rsTemp("参数值"))
                chk(Check先扣起付线).Value = IIf(Left(str参数值, 1) = "1", 1, 0)
            '20051220 陈东 丰县增加
            Case "门诊限额"
                txt97 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "超额禁止"
                chk97.Value = IIf(rsTemp("参数值") = 1, 1, 0)
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet中联.Show vbModal, frm医保类别
    参数设置 = mblnOK
End Function
