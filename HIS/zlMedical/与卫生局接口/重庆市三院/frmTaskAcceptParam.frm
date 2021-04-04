VERSION 5.00
Begin VB.Form frmTaskAcceptParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "本地参数"
   ClientHeight    =   2595
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4830
   Icon            =   "frmTaskAcceptParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3525
      TabIndex        =   8
      Top             =   2010
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2370
      TabIndex        =   7
      Top             =   2010
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   4560
      Begin VB.PictureBox picCmd 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4110
         ScaleHeight     =   255
         ScaleWidth      =   300
         TabIndex        =   5
         Top             =   615
         Width           =   300
         Begin VB.CommandButton cmd 
            Caption         =   "…"
            Height          =   240
            Left            =   15
            TabIndex        =   6
            Top             =   15
            Width           =   270
         End
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1185
         TabIndex        =   2
         Top             =   585
         Width           =   3240
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   3240
      End
      Begin VB.Label Label1 
         Caption         =   "这里的合约单位指的发送任务包的机构，即卫生局，如果没有信息，请在合约单位中建立。"
         Height          =   570
         Left            =   1170
         TabIndex        =   9
         Top             =   1020
         Width           =   3345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合约单位(&U)"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Top             =   660
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检部门(&D)"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   255
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmTaskAcceptParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean

Private Type Items
    团体名称 As String
    ID As Long
End Type

Private usrSaveGroup As Items

Private Function InitActivate() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化数据，发生在窗体的Activate事件
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT A.编码||'-'||A.名称,ID FROM 部门表 A,部门性质说明 B WHERE (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质='体检' ORDER BY A.编码||'-'||A.名称"
    
    Call OpenRecordset(rs, Me.Caption)
    If rs.BOF Then
        ShowSimpleMsg "没有体检性质的部门，请在部门管理中设置！"
        Exit Function
    End If
    
    '绑定数据到控件中
    Call AddComboData(cboDept, rs)
    
    '初始选择数据处理
    CboLocate cboDept, UserInfo.部门ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    
    InitActivate = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
End Function

Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  装载数据
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    CboLocate cboDept, Val(GetSetting("ZLSOFT", "公共全局\干保接口", "体检部门", "0")), True

    cmd.Tag = Val(GetSetting("ZLSOFT", "公共全局\干保接口", "合约单位", "0"))
    
    gstrSQL = "Select 名称 From 合约单位 Where ID=" & Val(cmd.Tag)
    Call OpenRecordset(rs, Me.Caption)
    If rs.BOF = False Then
        txt.Text = NVL(rs("名称"))
    Else
        cmd.Tag = ""
    End If
    
    LoadData = True
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
End Function

Private Function SaveData() As Boolean

    On Error GoTo errHand
    
    
    If cboDept.ListIndex = -1 Then
        Call SaveSetting("ZLSOFT", "公共全局\干保接口", "体检部门", "0")
    Else
        Call SaveSetting("ZLSOFT", "公共全局\干保接口", "体检部门", cboDept.ItemData(cboDept.ListIndex))
    End If
    
    Call SaveSetting("ZLSOFT", "公共全局\干保接口", "合约单位", cmd.Tag)
    
    SaveData = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
    
End Function

Private Sub cmd_Click()
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT -1 AS ID,NULL+0 AS 上级id,'0' AS 编码,'所有' AS 名称,'' as 简码,'' as 地址,0 AS 末级,'' AS 联系人,'' AS 电话,'' AS 电子邮件,'' AS 开户银行,'' AS 帐号,'' AS 地址,'' AS 说明 from dual " & _
                        "Union All " & _
                        "SELECT ID,DECODE(上级id,NULL,-1,0,-1,上级id) AS 上级id,编码,名称,简码,地址,0 AS 末级,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位  where 末级<>1 " & _
                        "Start With 上级id is null connect by prior ID=上级id " & _
                        "Union All " & _
                        "SELECT ID,DECODE(上级id,NULL,-1,0,-1,上级id) AS 上级id,编码,名称,简码,地址,1 AS 末级,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位  where 末级=1"
                        
    If ShowTxtSelectDialog(Me, txt, "编码,900,0,1;名称,1500,0,1;简码,900,0,1;地址,3000,0,1", Me.Name & "\体检团体选择", "请在下表中选择一个团体/单位。", strSQL, rs, 8790, 5100) Then
        
        txt.Text = NVL(rs("名称").Value)
        cmd.Tag = NVL(rs("ID").Value, 0)
        
        usrSaveGroup.团体名称 = txt.Text

    End If
    txt.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then Unload Me
End Sub

Private Sub Form_Activate()

    If mblnStartUp = False Then Exit Sub
    DoEvents
        
    If InitActivate = False Then
        mblnStartUp = False
        Unload Me
        Exit Sub
    End If
    
    mblnStartUp = False
    
    Call LoadData
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True

End Sub

Private Sub txt_Change()

    txt.Tag = "Changed"
    cmd.Tag = ""

End Sub

Private Sub txt_GotFocus()

    TxtSelAll txt

End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim strFilter As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If txt.Tag = "Changed" Then
            
            strFilter = "'%" & UCase(txt.Text) & "%'"
            
            gstrSQL = "select ID,编码,名称,简码,地址,联系人,电话,电子邮件,开户银行,帐号,地址,说明,末级 from 合约单位  where 末级=1 " & _
                " AND (编码 Like " & strFilter & " or 名称 Like " & strFilter & " OR 简码 Like " & strFilter & ")"
                
            If ShowTxtFilterDialog(Me, txt, "名称,1800,0,0;编码,900,0,0;简码,900,0,0;联系人,900,0,0;电话,1200,0,0", Me.Name & "\团体过滤选择", "请从下面选择一个团体单位", gstrSQL, rs) Then
                
                txt.Text = NVL(rs("名称"))
                cmd.Tag = NVL(rs("ID"))
                txt.Tag = ""
                
                usrSaveGroup.团体名称 = txt.Text
                
                                
            Else
                txt.Text = usrSaveGroup.团体名称
                Exit Sub
            End If
        End If

        PressKey vbKeyTab
        PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
    
    If txt.Tag = "Changed" Then
        txt.Text = usrSaveGroup.团体名称
        txt.Tag = ""
    End If

End Sub

