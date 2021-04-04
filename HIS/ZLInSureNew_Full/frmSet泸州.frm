VERSION 5.00
Begin VB.Form frmSet泸州 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmSet泸州.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txt地址 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1890
      TabIndex        =   13
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CheckBox chk远程 
      Caption         =   "通过医保中心完成身份验证(&M)"
      Height          =   285
      Left            =   210
      TabIndex        =   11
      Top             =   3030
      Width           =   2775
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医保服务器"
      Height          =   1545
      Left            =   150
      TabIndex        =   3
      Top             =   1320
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   10
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   9
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
         TabIndex        =   7
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   5
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   8
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   6
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   4
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4470
      TabIndex        =   15
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4470
      TabIndex        =   14
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame fra医院 
      Caption         =   "医院信息"
      Height          =   945
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4155
      Begin VB.ComboBox cmb级别 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   2595
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医保中心地址(&A)"
      Height          =   180
      Left            =   510
      TabIndex        =   12
      Top             =   3420
      Width           =   1350
   End
End
Attribute VB_Name = "frmSet泸州"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private mlng险类 As Long
Private mlng中心 As Long

Private Sub chk远程_Click()
    If chk远程.Value = 1 Then
        txt地址.BackColor = TxtEdit(Text医保服务器).BackColor
        txt地址.Enabled = True
    Else
        txt地址.BackColor = Me.BackColor
        txt地址.Enabled = False
    End If
End Sub

Private Sub cmb级别_Click()
    mblnChange = True
End Sub

Private Sub cmdTest_Click()
    If OraDataOpen(gcn泸州, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
    
    If TxtEdit(Text医保密码).Tag = TxtEdit(Text医保密码).Text Then
        cmb级别.Enabled = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    cmb级别.Enabled = False
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
    
    For lngCount = TxtEdit.LBound To TxtEdit.UBound
        If zlCommFun.StrIsValid(TxtEdit(lngCount).Text, TxtEdit(lngCount).MaxLength, TxtEdit(lngCount).hwnd) = False Then
            zlControl.TxtSelAll TxtEdit(lngCount)
            Exit Function
        End If
    Next
    
    If txt地址.Enabled = True Then
        If zlCommFun.StrIsValid(txt地址.Text, , txt地址.hwnd, "医保中心地址") = False Then
            Exit Function
        End If
        If Trim(txt地址.Text) = "" Then
            MsgBox "请输入医保中心地址。", vbInformation, gstrSysName
            zlControl.TxtSelAll txt地址
            txt地址.SetFocus
            Exit Function
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
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医院级别','" & cmb级别.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保用户名','" & TxtEdit(text医保用户).Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保用户密码','" & _
            IIf(TxtEdit(Text医保密码).Tag = "", "", EncryptStr(TxtEdit(Text医保密码).Tag, 256, True)) & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保服务器','" & TxtEdit(Text医保服务器).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'中心身份验证','" & IIf(chk远程.Value = 1, "是", "") & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'医保中心地址','" & IIf(chk远程.Value = 1, txt地址.Text, "") & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If gcn泸州.State = adStateClosed Then
        If OraDataOpen(gcn泸州, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag, False) = False Then
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
        If gcn泸州.State = adStateOpen Then gcn泸州.Close
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
    
    cmb级别.AddItem "33.三等甲级"
    cmb级别.AddItem "32.三等乙级"
    cmb级别.AddItem "23.二等甲级"
    cmb级别.AddItem "22.二等乙级"
    cmb级别.AddItem "13.一等甲级"
    cmb级别.AddItem "12.一等乙级"
    cmb级别.AddItem "03.社区医疗"
    cmb级别.AddItem "0.无级 "
    
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1] and (中心 is null or 中心=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng险类, lng中心)
    
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医院级别"
                SetComboByText cmb级别, str参数值, False, " "
            Case "中心身份验证"
                chk远程.Value = IIf(str参数值 = "是", 1, 0)
            Case "医保中心地址"
                txt地址.Text = str参数值
            Case "医保用户名"
                TxtEdit(text医保用户).Text = str参数值
            Case "医保服务器"
                TxtEdit(Text医保服务器).Text = str参数值
            Case "医保用户密码"
                TxtEdit(Text医保密码).Text = "        "    '假密码
                If str参数值 = "" Then
                    TxtEdit(Text医保密码).Tag = ""
                Else
                    TxtEdit(Text医保密码).Tag = EncryptStr(str参数值, 256, False)
                End If
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet泸州.Show vbModal, frm医保类别
    
    参数设置 = mblnOK
    If mblnOK = False Then
        '保存失败，关闭连接。以免用用户输入的其它用户名进行了连接
        If gcn泸州.State = adStateOpen Then gcn泸州.Close
    End If
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt地址_GotFocus()
    zlControl.TxtSelAll txt地址
End Sub
