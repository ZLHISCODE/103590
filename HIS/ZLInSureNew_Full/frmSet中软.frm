VERSION 5.00
Begin VB.Form frmSet中软 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   Icon            =   "frmSet中软.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra医保服务器 
      Caption         =   "医保服务器"
      Height          =   1545
      Left            =   150
      TabIndex        =   12
      Top             =   3000
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   19
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   18
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   14
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   10
         Left            =   390
         TabIndex        =   17
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   9
         Left            =   570
         TabIndex        =   15
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   8
         Left            =   390
         TabIndex        =   13
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.Frame fra中心参数 
      Caption         =   "中心参数"
      Height          =   4365
      Left            =   4440
      TabIndex        =   20
      Top             =   180
      Width           =   4605
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   24
         Top             =   718
         Width           =   1635
      End
      Begin VB.CheckBox chk 
         Caption         =   "是否定点机构(&9)"
         Height          =   255
         Index           =   0
         Left            =   1590
         TabIndex        =   39
         Top             =   3434
         Width           =   1785
      End
      Begin VB.CheckBox chk 
         Caption         =   "是否传输数据(&0)"
         Height          =   225
         Index           =   1
         Left            =   1590
         TabIndex        =   40
         Top             =   3780
         Width           =   1665
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   22
         Top             =   330
         Width           =   2835
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1590
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   1106
         Width           =   1635
      End
      Begin VB.CommandButton cmd目录 
         Caption         =   "…"
         Height          =   240
         Index           =   1
         Left            =   4140
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3090
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1590
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   1494
         Width           =   1635
      End
      Begin VB.CommandButton cmd目录 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   4140
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2685
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   37
         Top             =   3046
         Width           =   2835
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   34
         Top             =   2658
         Width           =   2835
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   32
         Top             =   2270
         Width           =   2835
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   30
         Top             =   1882
         Width           =   2835
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "FTP登陆用户(&2)"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   23
         Top             =   778
         Width           =   1260
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "远程主机(&1)"
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   21
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "FTP用户密码(&3)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   25
         Top             =   1166
         Width           =   1260
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "远程上传目录(&5)"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   29
         Top             =   1942
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "远程下载目录(&6)"
         Height          =   180
         Index           =   5
         Left            =   210
         TabIndex        =   31
         Top             =   2330
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "本地上传目录(&7)"
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   33
         Top             =   2718
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "本地下载目录(&8)"
         Height          =   180
         Index           =   7
         Left            =   210
         TabIndex        =   36
         Top             =   3106
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "FTP密码确认(&4)"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   27
         Top             =   1554
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7440
      TabIndex        =   42
      Top             =   4710
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   41
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Frame fra医院 
      Caption         =   "医院信息"
      Height          =   2745
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4155
      Begin VB.CheckBox chk 
         Caption         =   "超限部分(&V)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   630
         TabIndex        =   11
         Top             =   2400
         Width           =   1365
      End
      Begin VB.CheckBox chk 
         Caption         =   "首先自付部分(&I)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   2370
         TabIndex        =   10
         Top             =   2070
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Caption         =   "全自费部分(&L)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   630
         TabIndex        =   9
         Top             =   2070
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "首先自付部分(&F)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   2370
         TabIndex        =   7
         Top             =   1380
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Caption         =   "全自费部分(&A)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   630
         TabIndex        =   6
         Top             =   1410
         Width           =   1485
      End
      Begin VB.ComboBox cmb装钱 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   690
         Width           =   1785
      End
      Begin VB.ComboBox cmb级别 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1785
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "结算时个人帐户的使用范围(&B)"
         Height          =   180
         Index           =   3
         Left            =   330
         TabIndex        =   8
         Top             =   1770
         Width           =   2430
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "收费时个人帐户的使用范围(&R)"
         Height          =   180
         Index           =   2
         Left            =   330
         TabIndex        =   5
         Top             =   1110
         Width           =   2430
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "装钱模式(&N)"
         Height          =   180
         Index           =   1
         Left            =   330
         TabIndex        =   3
         Top             =   765
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "医院级别(&G)"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   390
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmSet中软"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum文本
    Text远程主机 = 0
    Text登陆用户 = 1
    Text用户密码 = 2
    Text确认密码 = 3
    Text远程上传 = 4
    Text远程下载 = 5
    Text本地上传 = 6
    Text本地下载 = 7
    text医保用户 = 8
    Text医保密码 = 9
    Text医保服务器 = 10
End Enum

Private Enum enum选择
    Check定点机构 = 0
    Check传输数据 = 1
    Check收费全自费 = 2
    Check收费首先自付 = 3
    Check结算全自费 = 4
    Check结算首先自付 = 5
    Check结算超限 = 6
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private mlng险类 As Long, mlng中心 As Long

Private Sub cmb级别_Click()
    mblnChange = True
End Sub

Private Sub cmb装钱_Change()
    mblnChange = True
End Sub

Private Sub cmdTest_Click()
    If OraDataOpen(gcn中软, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
    
    If TxtEdit(Text医保密码).Tag = TxtEdit(Text医保密码).Text Then
        cmb级别.Enabled = True
        cmb装钱.Enabled = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    cmb级别.Enabled = False
    cmb装钱.Enabled = False
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
    
    If cmb级别.Text = "" Then
        MsgBox "请设置医院级别。", vbInformation, gstrSysName
        If cmb级别.Enabled = True Then cmb级别.SetFocus
        Exit Function
    End If
    If cmb装钱.Text = "" Then
        MsgBox "请设置装钱模式。", vbInformation, gstrSysName
        If cmb装钱.Enabled = True Then cmb装钱.SetFocus
        Exit Function
    End If
    
    For lngCount = TxtEdit.LBound To TxtEdit.UBound
        If zlCommFun.StrIsValid(TxtEdit(lngCount).Text, TxtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll TxtEdit(lngCount)
            TxtEdit(lngCount).SetFocus
            Exit Function
        End If
        
        If lngCount < Text远程上传 Then
            If Len(TxtEdit(lngCount).Text) = 0 Then
                strTitle = Mid(lblEdit(lngCount).Caption, 1, InStr(lblEdit(lngCount).Caption, "(") - 1)
                If MsgBox("“" & strTitle & "”项长度为空可能使上传下载无法正常工作，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    zlControl.TxtSelAll TxtEdit(lngCount)
                    TxtEdit(lngCount).SetFocus
                    Exit Function
                End If
            End If
        End If
    Next
    
    '密码正确性
    If TxtEdit(Text用户密码).Text <> TxtEdit(Text确认密码).Text Then
        MsgBox "请确保两次输入的密码一致。", vbInformation, gstrSysName
        zlControl.TxtSelAll TxtEdit(Text用户密码)
        TxtEdit(Text用户密码).SetFocus
        Exit Function
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
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & "," & mlng中心 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医院级别','" & cmb级别.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'装钱模式','" & cmb装钱.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'收费个人帐户使用范围','" & _
                chk(Check收费全自费).Value & chk(Check收费首先自付).Value & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'结算个人帐户使用范围','" & _
                chk(Check结算全自费).Value & chk(Check结算首先自付).Value & chk(Check结算超限).Value & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保用户名','" & TxtEdit(text医保用户).Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保用户密码','" & _
            IIf(TxtEdit(Text医保密码).Tag = "", "", EncryptStr(TxtEdit(Text医保密码).Tag, 256, True)) & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保服务器','" & TxtEdit(Text医保服务器).Text & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'远程主机','" & TxtEdit(Text远程主机).Text & "',8)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'FTP登陆用户','" & TxtEdit(Text登陆用户).Text & "',9)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'FTP用户密码','" & _
            IIf(TxtEdit(Text用户密码).Tag = "", "", EncryptStr(TxtEdit(Text用户密码).Tag, 256, True)) & "',10)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'远程上传目录','" & TxtEdit(Text远程上传).Text & "',11)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'远程下载目录','" & TxtEdit(Text远程下载).Text & "',12)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'本地上传目录','" & TxtEdit(Text本地上传).Text & "',13)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'本地下载目录','" & TxtEdit(Text本地下载).Text & "',14)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'定点医疗机构','" & chk(Check定点机构).Value & "',15)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & mlng中心 & ",'传输数据','" & chk(Check传输数据).Value & "',16)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If gcn中软.State = adStateClosed Then
        If OraDataOpen(gcn中软, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag, False) = False Then
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

Private Sub chk_Click(Index As Integer)
    mblnChange = True
    
    If Index = Check收费全自费 Or Index = Check结算全自费 Then
        If chk(Index).Value = 1 Then
            chk(Index + 1).Value = 1
            chk(Index + 1).Enabled = False
        Else
            chk(Index + 1).Enabled = True
        End If
    End If
End Sub

Private Sub cmd目录_Click(Index As Integer)
    Dim strTitle As String
    Dim strPath As String
    
    If Index = 0 Then
        strTitle = "请选择保存上传文件的目录："
    Else
        strTitle = "请选择保存下载文件的目录："
    End If
    
    strPath = zlCommFun.OpenDir(Me.hwnd, strTitle)
    If strPath <> "" Then
        '保存目录名
        TxtEdit(Index + 6).Text = strPath
        TxtEdit(Index + 6).SetFocus
    End If
    
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text用户密码 Or Index = Text医保密码 Then
        TxtEdit(Index).Tag = TxtEdit(Index).Text
    End If
    
    If Index = Text医保服务器 Or Index = Text医保密码 Or Index = text医保用户 Then
        '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
        If gcn中软.State = adStateOpen Then gcn中软.Close
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
    mlng中心 = lng中心
    
    On Error GoTo errHandle
    
    cmb装钱.AddItem "0.不能装钱"
    cmb装钱.AddItem "1.在线装钱"
    cmb装钱.AddItem "2.离线装钱"
    
    cmb级别.AddItem "33.三等甲级"
    cmb级别.AddItem "32.三等乙级"
    cmb级别.AddItem "23.二等甲级"
    cmb级别.AddItem "22.二等乙级"
    cmb级别.AddItem "13.一等甲级"
    cmb级别.AddItem "12.一等乙级"
    cmb级别.AddItem "0.无级 "
    
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1] and (中心 is null or 中心=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng险类, lng中心)
    
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医院级别"
                SetComboByText cmb级别, str参数值, False, " "
            Case "装钱模式"
                SetComboByText cmb装钱, str参数值, False, " "
            Case "医保用户名"
                TxtEdit(text医保用户) = str参数值
            Case "医保服务器"
                TxtEdit(Text医保服务器) = str参数值
            Case "医保用户密码"
                TxtEdit(Text医保密码).Text = "        "    '假密码
                If str参数值 = "" Then
                    TxtEdit(Text医保密码).Tag = ""
                Else
                    TxtEdit(Text医保密码).Tag = EncryptStr(str参数值, 256, False)
                End If
            Case "远程主机"
                TxtEdit(Text远程主机) = str参数值
            Case "FTP登陆用户"
                TxtEdit(Text登陆用户) = str参数值
            Case "FTP用户密码"
                TxtEdit(Text用户密码).Text = "        "    '假密码
                TxtEdit(Text确认密码).Text = "        "    '假密码
                If str参数值 = "" Then
                    TxtEdit(Text用户密码).Tag = ""
                Else
                    TxtEdit(Text用户密码).Tag = EncryptStr(str参数值, 256, False)
                End If
            Case "远程上传目录"
                TxtEdit(Text远程上传) = str参数值
            Case "远程下载目录"
                TxtEdit(Text远程下载) = str参数值
            Case "本地上传目录"
                TxtEdit(Text本地上传) = str参数值
            Case "本地下载目录"
                TxtEdit(Text本地下载) = str参数值
            Case "定点医疗机构"
                chk(Check定点机构).Value = IIf(str参数值 = "1", 1, 0)
            Case "传输数据"
                chk(Check传输数据).Value = IIf(str参数值 = "1", 1, 0)
'            Case "收费个人帐户使用范围"
'                chk(Check收费全自费).Value = IIf(Left(str参数值, 1) = "1", 1, 0)
'                chk(Check收费首先自付).Value = IIf(Mid(str参数值, 2, 1) = "1", 1, 0)
'                '全自费优先
'                If chk(Check收费全自费).Value = 1 Then
'                    chk(Check收费首先自付).Value = 1
'                    chk(Check收费首先自付).Enabled = False
'                End If
'            Case "结算个人帐户使用范围"
'                chk(Check结算全自费).Value = IIf(Left(str参数值, 1) = "1", 1, 0)
'                chk(Check结算首先自付).Value = IIf(Mid(str参数值, 2, 1) = "1", 1, 0)
'                chk(Check结算超限).Value = IIf(Mid(str参数值, 3, 1) = "1", 1, 0)
'                '全自费优先
'                If chk(Check结算全自费).Value = 1 Then
'                    chk(Check结算首先自付).Value = 1
'                    chk(Check结算首先自付).Enabled = False
'                End If
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet中软.Show vbModal, frm医保类别
    
    参数设置 = mblnOK
    If mblnOK = False Then
        '保存失败，关闭连接。以免用用户输入的其它用户名进行了连接
        If gcn中软.State = adStateOpen Then gcn中软.Close
    End If
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
