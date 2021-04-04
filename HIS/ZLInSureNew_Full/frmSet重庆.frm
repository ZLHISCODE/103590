VERSION 5.00
Begin VB.Form frmSet重庆 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmSet重庆.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk中途结算仅处理已上传部分 
      Caption         =   "中途结算仅结算已上传的费用"
      Height          =   210
      Left            =   420
      TabIndex        =   16
      Top             =   4320
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.ComboBox cbo个人帐户 
      Height          =   300
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3420
      Width           =   2415
   End
   Begin VB.CheckBox chk离休干部门诊就诊录入诊断 
      Caption         =   "离休干部门诊就诊时需要录入诊断"
      Height          =   210
      Left            =   420
      TabIndex        =   15
      Top             =   4050
      Width           =   3975
   End
   Begin VB.CheckBox chk保存医保号 
      Caption         =   "保存上次就诊病人的医保号"
      Height          =   210
      Left            =   420
      TabIndex        =   14
      Top             =   3780
      Width           =   3975
   End
   Begin VB.Frame fra医院等级 
      Caption         =   "医院等级"
      Height          =   1365
      Left            =   150
      TabIndex        =   8
      Top             =   1980
      Width           =   4155
      Begin VB.ComboBox cmb等级 
         Height          =   300
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   870
         Width           =   2415
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "医院等级(&L)"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   10
         Top             =   930
         Width           =   990
      End
      Begin VB.Label lbl说明 
         Caption         =   "    该等级用于计算部分按医院等级进行限价的诊疗项目的实际价格。"
         Height          =   480
         Left            =   390
         TabIndex        =   9
         Top             =   330
         Width           =   3450
      End
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1605
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   7
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
      Left            =   4560
      TabIndex        =   18
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4560
      TabIndex        =   17
      Top             =   300
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "个人帐户(&A)"
      Height          =   180
      Index           =   4
      Left            =   450
      TabIndex        =   12
      Top             =   3480
      Width           =   990
   End
End
Attribute VB_Name = "frmSet重庆"
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

Dim mcnTest As New ADODB.Connection

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, TxtEdit(Text医保服务器).Text, TxtEdit(text医保用户).Text, TxtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    On Error Resume Next
    rsTemp.Open "select * from YPML where rownum<1", mcnTest, adOpenStatic, adLockReadOnly
    If Err <> 0 Then
        MsgBox "在该用户未发现有医保接口的相关表。", vbInformation, gstrSysName
        mcnTest.Close
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
        On Error Resume Next
        rsTemp.Open "select * from YPML where rownum<1", mcnTest, adOpenStatic, adLockReadOnly
        If Err <> 0 Then
            If MsgBox("在该用户未发现有医保接口的相关表，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                mcnTest.Close
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
    gstrSQL = "zl_保险参数_Delete(" & TYPE_重庆市 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆市 & ",null,'医保用户名','" & TxtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆市 & ",null,'医保用户密码','" & TxtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆市 & ",null,'医保服务器','" & TxtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆市 & ",null,'医院等级','" & cmb等级.Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆市 & ",null,'保存医保号','" & chk保存医保号.Value & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆市 & ",null,'离休就诊录入诊断','" & chk离休干部门诊就诊录入诊断.Value & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆市 & ",null,'个人帐户','" & Me.cbo个人帐户.ListIndex & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆市 & ",null,'中途结算','" & Me.chk中途结算仅处理已上传部分.Value & "',8)"
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
    
    mblnOK = False
    
    On Error GoTo errHandle
    
    cmb等级.AddItem "一级"
    cmb等级.AddItem "二级"
    cmb等级.AddItem "三级"
    
    With Me.cbo个人帐户
        .Clear
        .AddItem "使用个人帐户支付"
        .AddItem "住院时询问"
        .AddItem "门诊时询问"
        .AddItem "不使用个人帐户"
        .ListIndex = 0
    End With
    
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆市)
    
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
            Case "医院等级"
                On Error Resume Next
                cmb等级.Text = str参数值
            Case "保存医保号"
                chk保存医保号.Value = Val(str参数值)
            Case "离休就诊录入诊断"
                chk离休干部门诊就诊录入诊断.Value = Val(str参数值)
            Case "中途结算"
                '为了保持与以前一致，缺省是结算已上传部分
                str参数值 = Nvl(rsTemp("参数值"), 1)
                chk中途结算仅处理已上传部分.Value = Val(str参数值)
            Case "个人帐户"
                Me.cbo个人帐户.ListIndex = Val(str参数值)
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet重庆.Show vbModal, frm医保类别
    
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
