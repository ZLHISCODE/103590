VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSet重庆渝北 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmSet重庆渝北.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComCtl2.UpDown upDown 
      Height          =   300
      Left            =   1890
      TabIndex        =   16
      Top             =   3555
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   3
      BuddyControl    =   "txt提示"
      BuddyDispid     =   196622
      OrigLeft        =   2100
      OrigTop         =   3600
      OrigRight       =   2340
      OrigBottom      =   3945
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt提示 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "3"
      Top             =   3555
      Width           =   360
   End
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   285
      Left            =   5280
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3090
      Width           =   255
   End
   Begin VB.TextBox txt项目 
      Height          =   300
      Left            =   1455
      TabIndex        =   12
      Top             =   3075
      Width           =   4095
   End
   Begin VB.ComboBox cbo收入项目 
      Height          =   300
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2685
      Width           =   2415
   End
   Begin VB.Frame fra 
      Caption         =   "其他参数确定"
      Height          =   630
      Left            =   180
      TabIndex        =   20
      Top             =   1890
      Width           =   5655
      Begin VB.TextBox Txt限额 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1260
         TabIndex        =   8
         Text            =   "200.00"
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "处方单价大于                  元弹出审批信息"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   3960
      End
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1605
      Left            =   165
      TabIndex        =   19
      Top             =   195
      Width           =   4155
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   1
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1110
         Width           =   1635
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   0
         Top             =   390
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   2
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   4
         Top             =   1170
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "享受待遇时间小于       月提示!"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   3630
      Width           =   2700
   End
   Begin VB.Label lbl 
      Caption         =   "默认项目编码"
      Height          =   285
      Index           =   1
      Left            =   255
      TabIndex        =   11
      Top             =   3135
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "个人帐户支付"
      Height          =   180
      Left            =   255
      TabIndex        =   9
      Top             =   2760
      Width           =   1080
   End
End
Attribute VB_Name = "frmSet重庆渝北"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Dim mblnFirst As Boolean
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
End Enum

 
Public Function 参数设置() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select * From 收入项目 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.cbo收入项目.Clear
        Do While Not .EOF
             cbo收入项目.AddItem !名称
             cbo收入项目.ItemData(cbo收入项目.NewIndex) = Nvl(!ID)
            .MoveNext
        Loop
    End With
    
    
    frmSet重庆渝北.Show vbModal, frm医保类别

    参数设置 = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo收入项目_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub



Private Sub cmd病种_Click()
        '刘兴宏:20040706
        Dim strCode As String
        Dim STRNAME As String
        
        On Error Resume Next
        If frm保险项目选择重庆渝北.GetCode(Me, strCode, STRNAME, True) = True Then
            Me.txt项目.Text = strCode & "-" & STRNAME
            Me.txt项目.Tag = strCode
        End If
    
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_重庆渝北
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "医保用户名"
                  txtEdit(text医保用户).Text = Nvl(!参数值)
            Case "医保用户密码"
                  txtEdit(Text医保密码).Text = Nvl(!参数值)
            Case "医保服务器"
                  txtEdit(Text医保服务器).Text = Nvl(!参数值)
            Case "诊疗项目编码"
                  txt项目.Text = Nvl(!参数值)
                  txt项目.Tag = txt项目.Text
            Case "处方单价限制"
                 Txt限额.Text = Nvl(!参数值)
            Case "个人帐户"
                Dim i As Long
                For i = 0 To cbo收入项目.ListCount - 1
                    If cbo收入项目.ItemData(i) = Val(Nvl(!参数值)) Then
                        cbo收入项目.ListIndex = i
                        Exit For
                    End If
                Next
            Case "享受待遇提醒月数"
                txt提示.Text = Nvl(!参数值)
            End Select
            .MoveNext
        Loop
    End With
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text医保密码 Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Text医保服务器 Or Index = Text医保密码 Or Index = text医保用户 Then
        '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
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
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag, False) = False Then
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
    gstrSQL = "zl_保险参数_Delete(" & TYPE_重庆渝北 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆渝北 & ",null,'医保用户名','" & txtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆渝北 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆渝北 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆渝北 & ",null,'处方单价限制','" & Val(Txt限额.Text) & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If Me.cbo收入项目.ListIndex < 0 Then
        gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆渝北 & ",null,'个人帐户','" & 0 & "',5)"
    Else
        gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆渝北 & ",null,'个人帐户','" & Me.cbo收入项目.ItemData(Me.cbo收入项目.ListIndex) & "',5)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆渝北 & ",null,'诊疗项目编码','" & txt项目.Tag & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆渝北 & ",null,'享受待遇提醒月数','" & Val(txt提示.Text) & "',7)"
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
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub txt提示_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Txt限额_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt项目_Change()
    txt项目.Tag = ""
End Sub

Private Sub txt项目_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim rsTemp As New ADODB.Recordset
    Dim strLeft As String
    Dim strTemp As String
    Dim blnReturn As Boolean
    If txt项目.Text = "" Then Exit Sub
    strLeft = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    strTemp = "'" & strLeft & txt项目.Text & "%'"
    
    gstrSQL = " select  商品代码 as 医保编码,  医院大类编码, 药品通用中文名, 药品通用英文名,商品名, 商品曾用名, 服务项目结算方式, 报销标识, 医保标识, 是否处方用药, 药品适应症, 限制医生, 审批权限, 别名, 包装规格, " & _
             "         最小包装单位, 最小计量单位, 每日最大用量, 指导价格, 招标价格, 基金支付限价1, 基金支付限价2, 基金支付限价3, 实际执行价格, 自付比例1, 自付比例2, 自付比例3, 自付比例4, 自付比例5, 自付比例6, 自付比例7, 自付比例8,  " & _
             "         自付比例9, 自付比例10, 自付比例11, 自付比例12, 医院使用状态, 中心使用状态, 标准编号,  " & _
             "         五笔助记码1, 五笔助记码2, 五笔助记码3, 拼音助记码1, 拼音助记码2, 拼音助记码3, 备注, 医保经办机构,机构标准编号, 医疗机构编号, " & _
             "          修改时间, 目录分类  " & _
             "  from 医保服务项目目录" & _
             "  where 医院大类编码='61' and ( 商品代码 like " & strTemp & " Or 商品名 like " & strTemp & " Or " & _
             "        五笔助记码1 like " & UCase(strTemp) & " Or " & _
             "        拼音助记码1 like " & UCase(strTemp) & ")"
    
    If gcnOracle_CQYB.State = adStateOpen Then
        rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    Else
        '强制使记录集为打开状态
        gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    End If
                   
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_重庆渝北, rsTemp, "医保编码", "医保项目选择", "请选择对应的医保项目：")
        Else
            blnReturn = True
        End If
    Else
        MsgBox "无此项目!"
        Exit Sub
    End If
    
    If blnReturn = False Then Exit Sub

    '肯定是有记录集的
    txt项目.Text = rsTemp("医保编码") & "-" & Nvl(rsTemp!商品名)
    txt项目.Tag = rsTemp("医保编码")

End Sub

Private Sub txt项目_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt项目, KeyAscii, m文本式
End Sub
