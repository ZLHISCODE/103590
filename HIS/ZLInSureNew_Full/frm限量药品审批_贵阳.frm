VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm限量药品审批_贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "限量药品审批"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   Icon            =   "frm限量药品审批_贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fma 
      Caption         =   "已审批"
      Height          =   3525
      Index           =   2
      Left            =   0
      TabIndex        =   22
      Top             =   2850
      Width           =   10815
      Begin VB.CommandButton cmd操作日志 
         Caption         =   "操作日志"
         Height          =   315
         Left            =   8280
         TabIndex        =   45
         Top             =   3150
         Width           =   1065
      End
      Begin VB.CommandButton cmd取消审批 
         Caption         =   "取消审批"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9690
         TabIndex        =   43
         Top             =   3150
         Width           =   1065
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
         Height          =   2895
         Left            =   60
         TabIndex        =   19
         Top             =   210
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fma 
      Caption         =   "待审批"
      Height          =   1815
      Index           =   1
      Left            =   0
      TabIndex        =   21
      Top             =   1020
      Width           =   10815
      Begin VB.TextBox txt药品名称 
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   180
         Width           =   3495
      End
      Begin VB.TextBox txt金额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   6450
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1020
         Width           =   1515
      End
      Begin VB.TextBox txt限量 
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txt售价 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txt产地 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   7125
      End
      Begin VB.TextBox txt规格 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   2685
      End
      Begin VB.TextBox txt单位 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txt备注 
         Height          =   285
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1448
         Width           =   7125
      End
      Begin VB.CommandButton cmd药品 
         Caption         =   "…"
         Height          =   315
         Left            =   4350
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
      End
      Begin VB.CommandButton cmd审批 
         Caption         =   "通过审批"
         Height          =   315
         Left            =   8040
         TabIndex        =   18
         Top             =   1410
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker dtp效期 
         Height          =   300
         Left            =   8700
         TabIndex        =   16
         Top             =   1020
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   105971715
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "有效期"
         Enabled         =   0   'False
         Height          =   180
         Index           =   18
         Left            =   8130
         TabIndex        =   44
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品名称"
         Height          =   180
         Index           =   17
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "规格"
         Enabled         =   0   'False
         Height          =   180
         Index           =   16
         Left            =   4890
         TabIndex        =   41
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "产地"
         Enabled         =   0   'False
         Height          =   180
         Index           =   15
         Left            =   480
         TabIndex        =   40
         Top             =   660
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "准许超量"
         Height          =   180
         Index           =   14
         Left            =   3840
         TabIndex        =   39
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售价单位"
         Enabled         =   0   'False
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售价"
         Enabled         =   0   'False
         Height          =   180
         Index           =   12
         Left            =   1800
         TabIndex        =   37
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售价金额"
         Enabled         =   0   'False
         Height          =   180
         Index           =   11
         Left            =   5760
         TabIndex        =   36
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   180
         Index           =   10
         Left            =   450
         TabIndex        =   35
         Top             =   1500
         Width           =   360
      End
   End
   Begin VB.CommandButton cmd退出 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   345
      Left            =   9720
      TabIndex        =   23
      Top             =   6510
      Width           =   1065
   End
   Begin VB.Frame fma 
      Caption         =   "基本信息"
      Height          =   945
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   60
      Width           =   10815
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   9000
         TabIndex        =   8
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   5940
         TabIndex        =   7
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   3480
         TabIndex        =   6
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1770
         TabIndex        =   5
         Top             =   540
         Width           =   705
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   4
         Top             =   540
         Width           =   465
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   9000
         TabIndex        =   3
         Top             =   203
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   5940
         TabIndex        =   2
         Top             =   203
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   1
         Top             =   203
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   203
         Width           =   1635
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上次就诊时间"
         Enabled         =   0   'False
         Height          =   180
         Index           =   8
         Left            =   7920
         TabIndex        =   32
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         Enabled         =   0   'False
         Height          =   180
         Index           =   7
         Left            =   2910
         TabIndex        =   31
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Enabled         =   0   'False
         Height          =   180
         Index           =   6
         Left            =   450
         TabIndex        =   30
         Top             =   592
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Enabled         =   0   'False
         Height          =   180
         Index           =   5
         Left            =   1410
         TabIndex        =   29
         Top             =   585
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Enabled         =   0   'False
         Height          =   180
         Index           =   4
         Left            =   5400
         TabIndex        =   28
         Top             =   592
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡号"
         Enabled         =   0   'False
         Height          =   180
         Index           =   3
         Left            =   2730
         TabIndex        =   27
         Top             =   585
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Enabled         =   0   'False
         Height          =   180
         Index           =   2
         Left            =   8460
         TabIndex        =   26
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Enabled         =   0   'False
         Height          =   180
         Index           =   1
         Left            =   5400
         TabIndex        =   25
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   24
         Top             =   255
         Width           =   360
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "###"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   9
      Left            =   60
      TabIndex        =   33
      Top             =   6480
      Width           =   270
   End
End
Attribute VB_Name = "frm限量药品审批_贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Public Sub ShowME(ByVal intinsure As Integer)
    mint险类 = intinsure
    dtp效期.Value = zlDatabase.Currentdate + 7
    lbl(9).Caption = "提示：1、在姓名处可以通过输入-病人ID:+住院号;*门诊号;/医保号或直接录入姓名的方式提取病人信息" & vbNewLine & _
                    "       2、所有涉及到“限量”或“准许超量”时，均指按月限量，如果程序设置成按其他方式限量时，系统会自动换算！"
    Me.Show 1
End Sub

Private Sub cmd操作日志_Click()
     Call zl9Report.ReportOpen(gcnOracle, 0, "SYB_LOG1", Me)
End Sub

Private Sub cmd取消审批_Click()
    On Error GoTo errHand
    With mshBill
        If .Row = 0 Or .Rows = 1 Then ShowMsgbox "请选择要取消审批的记录！": Exit Sub
        If MsgBox("要取消“[" & .TextMatrix(.Row, 2) & "]" & .TextMatrix(.Row, 3) & "”的限量审批信息吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_用药限量审批_贵阳_Delete(" & mint险类 & ",'" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, 1) & "','" & UserInfo.姓名 & "',Sysdate)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call Get已审批
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd审批_Click()
    On Error GoTo errHand
    If txt药品名称.Tag = "" Then ShowMsgbox "请确定药品信息！": txt药品名称.SetFocus: Exit Sub
    If Val(txt限量.Text) <= 0 Then ShowMsgbox "药品“准许超量”必须大于0！": txt限量.SetFocus: Exit Sub
    If Val(txt(0).Tag) = 0 Then ShowMsgbox "请先确定病人身份信息！": txt(0).SetFocus: Exit Sub
    If Format(dtp效期.Value, "yyyy-mm-dd") < Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
        ShowMsgbox "药品超量使用有效期不能小于今天！": dtp效期.SetFocus: Exit Sub
    End If
    gstrSQL = "Zl_用药限量审批_贵阳_Insert(" & mint险类 & ",'" & txt(0).Tag & "','" & txt药品名称.Tag & "','" & txt限量.Text & "'," & _
              "To_Date('" & Format(dtp效期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd'),'" & txt备注.Text & "','" & UserInfo.姓名 & "',Sysdate)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call Get已审批
    ShowMsgbox "药品限量审批信息设置成功！"
    txt药品名称.Text = "": txt备注.Text = "": txt产地.Text = "": txt金额.Text = ""
    txt规格.Text = ""
    txt规格.Text = ""
    txt单位.Text = "": txt售价.Text = ""
    txt限量.Text = 1: txt药品名称.Tag = "": txt药品名称.SetFocus
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd退出_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshBill.Cols > 3 Then Call SaveFlexState(mshBill, Me.Caption)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo errHand
    Dim strCode As String, rsTemp As New ADODB.Recordset
    Select Case Index
        Case 0
            If KeyAscii <> vbKeyReturn Then Exit Sub
                strCode = Trim(txt(Index).Text)
                If Left(strCode, 1) = "-" And IsNumeric(Mid(strCode, 2)) Then  '病人ID
                   gstrSQL = "Select A.病人id As ID, A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.姓名, A.性别, A.年龄, B.医保号," & _
                           "      B.就诊时间 As 最后就诊时间 " & _
                           "From 病人信息 A, 保险帐户 B " & _
                           "Where A.险类 = B.险类 And A.病人ID = B.病人ID And B.险类=" & mint险类 & " And A.病人ID='" & Mid(strCode, 2) & "'"
                ElseIf Left(strCode, 1) = "+" And IsNumeric(Mid(strCode, 2)) Then  '住院号
                    gstrSQL = "Select A.病人id As ID, A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.姓名, A.性别, A.年龄, B.医保号," & _
                           "      B.就诊时间 As 最后就诊时间 " & _
                           "From 病人信息 A, 保险帐户 B " & _
                           "Where A.险类 = B.险类 And A.病人ID = B.病人ID And B.险类=" & mint险类 & " And A.住院号='" & Mid(strCode, 2) & "'"
                ElseIf (Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then      '门诊号
                    gstrSQL = "Select A.病人id As ID, A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.姓名, A.性别, A.年龄, B.医保号," & _
                           "      B.就诊时间 As 最后就诊时间 " & _
                           "From 病人信息 A, 保险帐户 B " & _
                           "Where A.险类 = B.险类 And A.病人ID = B.病人ID And B.险类=" & mint险类 & " And A.门诊号='" & Mid(strCode, 2) & "'"
                ElseIf (Left(strCode, 1) = "/") And IsNumeric(Mid(strCode, 2)) Then      '医保号
                    gstrSQL = "Select A.病人id As ID, A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.姓名, A.性别, A.年龄, B.医保号," & _
                           "      B.就诊时间 As 最后就诊时间 " & _
                           "From 病人信息 A, 保险帐户 B " & _
                           "Where A.险类 = B.险类 And A.病人ID = B.病人ID And B.险类=" & mint险类 & " And B.医保号='" & Mid(strCode, 2) & "'"
                Else
                    '全当成姓名
                    gstrSQL = "Select A.病人id As ID, A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.姓名, A.性别, A.年龄, B.医保号," & _
                           "      B.就诊时间 As 最后就诊时间 " & _
                           "From 病人信息 A, 保险帐户 B " & _
                           "Where A.险类 = B.险类 And A.病人ID = B.病人ID And B.险类=" & mint险类 & " And A.姓名='" & strCode & "'"
                End If
                Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病人", , txt(Index).Text)
                If Not rsTemp Is Nothing Then
                    If rsTemp.State = 1 Then
                        txt(0).Tag = Nvl(rsTemp!ID): txt(0).Text = Nvl(rsTemp!姓名)
                        txt(1).Text = Nvl(rsTemp!病人ID): txt(2).Text = Nvl(rsTemp!门诊号)
                        txt(3).Text = Nvl(rsTemp!住院号): txt(4).Text = Nvl(rsTemp!性别)
                        txt(5).Text = Nvl(rsTemp!年龄): txt(6).Text = Nvl(rsTemp!就诊卡号)
                        txt(7).Text = Nvl(rsTemp!医保号): txt(8).Text = Nvl(rsTemp!最后就诊时间)
                        Call zlCommFun.PressKey(vbKeyTab)
                    Else
                       txt(0).Tag = "": txt(0).SetFocus
                    End If
                Else
                    txt(0).Tag = "": txt(0).SetFocus
                End If
                Call Get已审批
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Get已审批()
   '提已有的审批信息
   On Error GoTo errHand
   Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.病人id, A.药品id, B.编码, B.名称, B.规格, B.计算单位 As 售价单位, C.现价 As 售价, A.准许超量, A.有效期," & _
              "       A.备注  From 用药限量审批_贵阳 A, 收费细目 B, 收费价目 C " & _
              "Where A.药品id = B.ID And B.ID = C.收费细目id And (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) " & _
              "      And A.险类=[1] And A.病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, CLng(Val(txt(0).Tag)))
    Set mshBill.DataSource = rsTemp
    Call CenterTableCaption(mshBill)
    mshBill.ColWidth(0) = 0: mshBill.ColWidth(1) = 0
    cmd取消审批.Enabled = mshBill.Rows > 1 And mshBill.Row <> 0
    Call RestoreFlexState(mshBill, Me.Caption)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub cmd药品_Click()
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select  ID, Null As 上级id, 0 As 末级, 名称 As 分类,  编码, 名称, Null As 规格, Null As 产地," & _
              "      Null As 售价单位, -null As 售价 From 药品用途分类 Union All " & _
              "Select B.药品id As ID, A.用途分类id As 上级id, 1 As 末级, D.名称 As 分类, B.编码, B.名称, B.规格, B.产地, B.售价单位," & _
              "      C.现价 As 售价 From 药品信息 A, 药品目录 B, 收费价目 C, 药品用途分类 D " & _
              "Where A.用途分类id = D.ID And A.药名id = B.药名id And B.药品id = C.收费细目id And " & _
              "     (B.撤档时间 Is Null Or B.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
              "     (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) "
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 2, "药品目录", , txt药品名称.Text)
    If Not rsTemp Is Nothing Then
        If rsTemp.State = 1 Then
            txt药品名称.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
            txt规格.Text = Nvl(rsTemp!规格): txt产地.Text = Nvl(rsTemp!产地)
            txt单位.Text = Nvl(rsTemp!售价单位): txt售价.Text = Format(Nvl(rsTemp!售价), "0.00000")
            txt限量.Text = 1: txt药品名称.Tag = Nvl(rsTemp!ID)
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt限量.Text = 0: txt药品名称.Tag = ""
        End If
    Else
        txt限量.Text = 0: txt药品名称.Tag = ""
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub txt限量_Change()
    txt金额.Text = Format(Val(txt售价.Text) * Val(txt限量.Text), "0.00000")
End Sub

Private Sub txt限量_GotFocus()
     zlControl.TxtSelAll txt限量
End Sub

Private Sub txt限量_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt限量, KeyAscii, m金额式)
End Sub

Private Sub txt药品名称_GotFocus()
    zlControl.TxtSelAll txt药品名称
End Sub

Private Sub txt药品名称_KeyPress(KeyAscii As Integer)
     On Error GoTo errHand
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Distinct B.药品id As ID,Null As 上级ID,D.名称 As 分类, B.编码, B.名称, B.规格, B.产地," & _
              "       B.售价单位, C.现价 As 售价 From 药品信息 A, 药品目录 B, 收费价目 C, 药品用途分类 D, 收费别名 E " & _
              "Where A.用途分类id = D.ID And A.药名id = B.药名id And B.药品id = C.收费细目id And " & _
              "     (B.撤档时间 Is Null Or B.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
              "     (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) And B.药品id = E.收费细目id " & _
              "     And (Upper(B.编码) Like '%" & UCase(txt药品名称.Text) & "%' Or  Upper(B.名称) Like '%" & UCase(txt药品名称.Text) & "%' " & _
              "     Or Upper(E.简码) Like '%" & UCase(txt药品名称.Text) & "%')"
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "药品目录", , txt药品名称.Text)
    If Not rsTemp Is Nothing Then
        If rsTemp.State = 1 Then
            txt药品名称.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
            txt规格.Text = Nvl(rsTemp!规格): txt产地.Text = Nvl(rsTemp!产地)
            txt单位.Text = Nvl(rsTemp!售价单位): txt售价.Text = Format(Nvl(rsTemp!售价), "0.00000")
            txt限量.Text = 1: txt药品名称.Tag = Nvl(rsTemp!ID)
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt限量.Text = 0: txt药品名称.Tag = ""
            txt药品名称.SetFocus
        End If
    Else
        txt限量.Text = 0: txt药品名称.Tag = ""
        txt药品名称.SetFocus
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub


