VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm保险病种编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险病种编辑"
   ClientHeight    =   5280
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   7950
   Icon            =   "frm保险病种编辑.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdClear 
      Caption         =   "全清(&A)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   6660
      TabIndex        =   20
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "清除(&D)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5460
      TabIndex        =   19
      Top             =   4110
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   360
      Left            =   3930
      TabIndex        =   13
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "特准大类(&0)"
      TabPicture(0)   =   "frm保险病种编辑.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "特准明细(&1)"
      TabPicture(1)   =   "frm保险病种编辑.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "选择(&S)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4110
      TabIndex        =   18
      Top             =   4110
      Width           =   1100
   End
   Begin VB.Frame Fra类别 
      Caption         =   "类别"
      Height          =   1545
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   3675
      Begin VB.OptionButton opt类别 
         Caption         =   "慢性病(&M)"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   780
         Width           =   1155
      End
      Begin VB.OptionButton opt类别 
         Caption         =   "普通病(&G)"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   420
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt类别 
         Caption         =   "特种病(&T)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   1140
         Width           =   1155
      End
   End
   Begin VB.Frame fra基本 
      Caption         =   "基本"
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3675
      Begin VB.TextBox txt封顶线 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2070
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox chk封顶线 
         Caption         =   "使用特殊封顶线(&T)"
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   825
         MaxLength       =   6
         TabIndex        =   2
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   825
         MaxLength       =   50
         TabIndex        =   4
         Top             =   780
         Width           =   2715
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   825
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1185
         Width           =   1095
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&E)"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   1245
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&U)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   450
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   195
      TabIndex        =   23
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5460
      TabIndex        =   21
      Top             =   4770
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6660
      TabIndex        =   22
      Top             =   4770
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2940
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种编辑.frx":0044
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种编辑.frx":035E
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种编辑.frx":0678
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种编辑.frx":0992
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种编辑.frx":0CAC
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种编辑.frx":1246
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra大类 
      Height          =   4245
      Left            =   3930
      TabIndex        =   14
      Top             =   420
      Width           =   3885
      Begin MSComctlLib.ListView lvw大类 
         Height          =   3300
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   5821
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "名称"
            Text            =   "名称"
            Object.Width           =   3933
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "性质"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Fra明细 
      Height          =   4245
      Left            =   3930
      TabIndex        =   16
      Top             =   420
      Width           =   3885
      Begin MSComctlLib.ListView Lvw明细 
         Height          =   3300
         Left            =   90
         TabIndex        =   17
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   5821
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "名称"
            Text            =   "名称"
            Object.Width           =   3933
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "规格"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "性质"
            Object.Width           =   1764
         EndProperty
      End
   End
End
Attribute VB_Name = "frm保险病种编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum编辑
    text编码 = 0
    Text名称 = 1
    Text简码 = 2
End Enum

Dim mlng险类 As Long
Dim mstrID As String         '当前编辑的医保大类ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub chk封顶线_Click()
    mblnChange = True
    If chk封顶线.Value = 1 Then
        txt封顶线.Enabled = True
        txt封顶线.BackColor = txtEdit(1).BackColor
    Else
        txt封顶线.Text = ""
        txt封顶线.Enabled = False
        txt封顶线.BackColor = fra基本.BackColor
    End If
End Sub

Private Sub chk封顶线_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 '使之不响
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub CmdClear_Click()
    Dim objLvw As ListView
    
    If SSTab.Tab = 0 Then
        Set objLvw = lvw大类
    Else
        Set objLvw = Lvw明细
    End If
    
    objLvw.ListItems.Clear
    
    CmdDel.Enabled = (objLvw.ListItems.Count <> 0)
    CmdClear.Enabled = CmdDel.Enabled
End Sub

Private Sub CmdDel_Click()
    Dim lngItem As Long
    Dim objLvw As ListView
    
    If SSTab.Tab = 0 Then
        Set objLvw = lvw大类
    Else
        Set objLvw = Lvw明细
    End If
    
    With objLvw
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        
        For lngItem = 1 To .ListItems.Count
            If lngItem > .ListItems.Count Then Exit For
            If .ListItems(lngItem).Selected Then
                .ListItems.Remove .ListItems(lngItem).Key
                lngItem = lngItem - 1
            End If
        Next
        
        If .ListItems.Count <> 0 Then .ListItems(1).Selected = True
    End With
    
    CmdDel.Enabled = (objLvw.ListItems.Count <> 0)
    CmdClear.Enabled = CmdDel.Enabled
End Sub

Private Sub Form_Load()
    If mlng险类 = TYPE_重庆银海版 Then
        Load opt类别(3)
        opt类别(3).Top = opt类别(0).Top
        opt类别(3).Left = opt类别(0).Left + opt类别(0).Width + 150
        opt类别(3).Visible = True
        
        Load opt类别(4)
        opt类别(4).Top = opt类别(1).Top
        opt类别(4).Left = opt类别(1).Left + opt类别(1).Width + 150
        opt类别(4).Visible = True
        
        '修改名称
        opt类别(1).Caption = "普通病"
        opt类别(1).Caption = "特殊病"
        opt类别(2).Caption = "急诊病"
        opt类别(3).Caption = "恶性肿瘤"
        opt类别(4).Caption = "精神病"
        opt类别(0).Value = True
    End If
End Sub

Private Sub lvw大类_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvw大类, ColumnHeader.Index)
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    Dim objLvw As ListView
    Select Case SSTab.Tab
    Case 0
        Fra明细.Visible = True
        Set objLvw = lvw大类
        fra大类.ZOrder
    Case 1
        fra大类.Visible = True
        Set objLvw = Lvw明细
        Fra明细.ZOrder
    End Select
    
    SSTab.ZOrder
    cmdADD.ZOrder
    CmdClear.ZOrder
    CmdDel.ZOrder
    
    CmdDel.Enabled = (objLvw.ListItems.Count <> 0)
    CmdClear.Enabled = CmdDel.Enabled
End Sub

Private Sub cmdADD_Click()
    With frm特准项目选择
        .lng险类 = mlng险类
        .bln明细 = (SSTab.Tab = 1)
        Set .frmParent = Me
        .Show 1, Me
    End With
    Call SSTab_Click(SSTab.Tab)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
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
    Dim lngIndex As Long
    
    If IsValid() = False Then Exit Sub
    If Save项目() = False Then Exit Sub
    
    If mstrID = "" Then
        '连续新增
        'Modified by 朱玉宝 20031218 地区：福州
        If mlng险类 = TYPE_福建巨龙 Or mlng险类 = TYPE_福建省 Or mlng险类 = TYPE_福州市 Or mlng险类 = TYPE_南平市 Then
            txtEdit(text编码).Text = GetMaxCode
        Else
            txtEdit(text编码).Text = zlDatabase.GetMax("保险病种", "编码", 6, " where 险类=" & mlng险类)
        End If
        For lngIndex = Text名称 To Text简码
            txtEdit(lngIndex).Text = ""
        Next
        lvw大类.ListItems.Clear
        Lvw明细.ListItems.Clear
        
        mblnChange = False
        txtEdit(text编码).SetFocus
    Else
        mblnChange = False
        Unload Me
    End If
End Sub

Private Function Save项目() As Boolean
    Dim lngID As Long, lng类别 As Long, lng病种ID As Long
    Dim lngIndex As Long, lst As ListItem
    Dim strCode As String
    Dim rsTmp As New ADODB.Recordset
    
    For lngIndex = opt类别.LBound To opt类别.UBound
        If opt类别(lngIndex).Value = True Then
            lng类别 = lngIndex
            Exit For
        End If
    Next
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    If mstrID = "" Then
        '新增
        lngID = zlDatabase.GetNextId("保险病种")
        'Modified by 朱玉宝 20031218 地区：福州
        If mlng险类 = TYPE_福建巨龙 Or mlng险类 = TYPE_福建省 Or mlng险类 = TYPE_福州市 Or mlng险类 = TYPE_南平市 Then
            If CheckCode(txtEdit(text编码)) = False Then Exit Function
            '获取保险编码
            strCode = zlDatabase.GetMax("保险病种", "编码", 6, " Where 险类=" & mlng险类)
            gstrSQL = "zl_保险病种_INSERT(" & lngID & "," & mlng险类 & ",'" & strCode & "','" & _
                    Trim(txtEdit(text编码).Text) & "@@" & Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng类别 & ",null,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else
            lng病种ID = lngID
            gstrSQL = "zl_保险病种_INSERT(" & lngID & "," & mlng险类 & ",'" & Trim(txtEdit(text编码).Text) & "','" & _
                    Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng类别 & _
                    "," & chk封顶线.Value & "," & IIf(chk封顶线.Value = 0, "null", IIf(txt封顶线.Text = "", "null", txt封顶线.Text)) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Else
        'Modified by 朱玉宝 20031218 地区：福州
        If mlng险类 = TYPE_福建巨龙 Or mlng险类 = TYPE_福建省 Or mlng险类 = TYPE_福州市 Or mlng险类 = TYPE_南平市 Then
            If CheckCode(txtEdit(text编码), False) = False Then Exit Function
            '获取保险编码
            gstrSQL = "Select 编码 From 保险病种 Where 险类=" & mlng险类 & " And ID=" & mstrID
            Call OpenRecordset(rsTmp, "获取当前保险病种的编码")
            strCode = rsTmp!编码
            
            gstrSQL = "zl_保险病种_Update(" & mstrID & ",'" & strCode & "','" & _
                    Trim(txtEdit(text编码).Text) & "@@" & Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng类别 & ",null,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else
            lng病种ID = mstrID
            gstrSQL = "zl_保险病种_Update(" & mstrID & ",'" & Trim(txtEdit(text编码).Text) & "','" & _
                    Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng类别 & _
                    "," & chk封顶线.Value & "," & IIf(chk封顶线.Value = 0, "null", IIf(txt封顶线.Text = "", "null", txt封顶线.Text)) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        gstrSQL = "zl_保险特准项目_INSERT(" & mstrID & ",NULL,0,0,1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    '更新特准项目（大类）
    For Each lst In lvw大类.ListItems
        gstrSQL = "zl_保险特准项目_INSERT(" & lng病种ID & "," & Mid(lst.Key, 2) & ",1," & Mid(lst.SubItems(1), 1, 1) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    '更新特准项目（明细）
    For Each lst In Lvw明细.ListItems
        gstrSQL = "zl_保险特准项目_INSERT(" & lng病种ID & "," & Mid(lst.Key, 2) & ",0," & Mid(lst.SubItems(2), 1, 1) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    '更新主界面
    If mstrID = "" Then
        Set lst = frm保险病种.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text编码), "Disease", "Disease")
    Else
        Set lst = frm保险病种.lvwItem.SelectedItem
    End If
    lst.SubItems(1) = Trim(txtEdit(Text名称).Text)
    lst.SubItems(2) = Trim(txtEdit(Text简码).Text)
    '调试重庆医保银海版 204-03-31
    If mlng险类 = TYPE_重庆银海版 Then
        lst.SubItems(3) = IIf(lng类别 = 1, "特殊病", IIf(lng类别 = 2, "急诊病", IIf(lng类别 = 3, "恶性肿瘤", IIf(lng类别 = 4, "精神病", "普通病"))))
    Else
        lst.SubItems(3) = IIf(lng类别 = 0, "普通病", IIf(lng类别 = 1, "慢性病", "特种病"))
    End If
    lst.SubItems(4) = IIf(chk封顶线.Value = 1, "有", "")
    lst.SubItems(5) = IIf(chk封顶线.Value = 0, "", IIf(txt封顶线.Text = "", "无封顶线", txt封顶线.Text))
    
    Save项目 = True
    mblnOK = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
'功能:分析输入有关医保类别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim lngIndex As Integer
    For lngIndex = text编码 To Text简码
        If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
            txtEdit(lngIndex).SetFocus
            zlControl.TxtSelAll txtEdit(lngIndex)
            Exit Function
        End If
        
        If lngIndex = text编码 Or lngIndex = Text名称 Then
            If Len(Trim(txtEdit(lngIndex).Text)) = 0 Then
                txtEdit(lngIndex).Text = ""
                MsgBox "编码或名称都不能为空。", vbExclamation, gstrSysName
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    If txt封顶线.Enabled = True Then
        If txt封顶线.Text <> "" Then
            If zlCommFun.IntIsValid(txt封顶线.Text, 8, True, True, txt封顶线.hwnd, "封顶线金额") = False Then
                Exit Function
            End If
        End If
    End If
    
    IsValid = True
End Function

Private Sub opt类别_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt类别_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text名称 Then
        txtEdit(Text简码).Text = zlCommFun.SpellCode(txtEdit(Text名称).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text名称
            zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 '使之不响
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = text编码 Then
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Public Function 编辑病种(ByVal lng险类 As Long, ByVal strID As String) As Boolean
'功能:用来与调用的医保类别管理窗口进行通讯的程序
'参数:str序号           当前编辑的医保类别的的序号
'返回值:编辑成功返回True,否则为False
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer, lst As ListItem
    
    mblnOK = False
    mlng险类 = lng险类
    mstrID = strID
    
    rsTemp.CursorLocation = adUseClient
    If mstrID <> "" Then
        '修改医保病种
        'Modified by 朱玉宝 20031218 地区：福州
        If mlng险类 = TYPE_福建巨龙 Or mlng险类 = TYPE_福建省 Or mlng险类 = TYPE_福州市 Or mlng险类 = TYPE_南平市 Then
            'Modified by 朱玉宝 20031218 地区：福州
            txtEdit(text编码).MaxLength = 20
            gstrSQL = "select substr(名称,1,instr(名称,'@@')-1) 编码,substr(名称,instr(名称,'@@')+2) 名称,简码,nvl(类别,'0') as 类别,特殊封顶线,封顶线金额 from 保险病种 where ID=" & mstrID
        Else
            gstrSQL = "select 编码,名称,简码,nvl(类别,'0') as 类别,特殊封顶线,封顶线金额 from 保险病种 where ID=" & mstrID
        End If
        Call OpenRecordset(rsTemp, Me.Caption)
        
        txtEdit(text编码).Text = rsTemp("编码")
        txtEdit(Text名称).Text = rsTemp("名称")
        txtEdit(Text简码).Text = IIf(IsNull(rsTemp("简码")), "", rsTemp("简码"))
        opt类别(rsTemp("类别")).Value = True
        
        If rsTemp("特殊封顶线") = 1 Then
            chk封顶线.Value = 1
            If Not IsNull(rsTemp("封顶线金额")) Then
                txt封顶线.Text = rsTemp("封顶线金额")
            End If
        End If
        
        '修改特准大类
        gstrSQL = "select A.ID,A.编码,A.名称,decode(B.性质,1,'1-允许',2,'2-排斥','0-不限') 性质 from 保险支付大类 A,保险特准项目 B where A.ID=B.收费细目ID and B.大类=1 and B.病种ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        Do Until rsTemp.EOF
            Set lst = lvw大类.ListItems.Add(, "K" & rsTemp("ID"), "[" & rsTemp!编码 & "]" & rsTemp("名称"), "Limit", "Limit")
            lst.SubItems(1) = rsTemp("性质")
            rsTemp.MoveNext
        Loop
    
        '修改特准项目
        gstrSQL = "select A.ID,A.编码,A.名称,A.规格,decode(B.性质,1,'1-允许',2,'2-排斥','0-不限') 性质 from 收费细目 A,保险特准项目 B where A.ID=B.收费细目ID and B.大类=0 and B.病种ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        Do Until rsTemp.EOF
            Set lst = Lvw明细.ListItems.Add(, "K" & rsTemp("ID"), "[" & rsTemp!编码 & "]" & rsTemp("名称"), "Fix", "Fix")
            lst.SubItems(1) = Nvl(rsTemp("规格"))
            lst.SubItems(2) = Nvl(rsTemp("性质"))
            rsTemp.MoveNext
        Loop
    
    Else
        '新增医保大类
        txtEdit(text编码).Text = zlDatabase.GetMax("保险病种", "编码", 6, " where 险类=" & mlng险类)
    End If
    
    '调试重庆医保银海版 204-03-31
    'Modified by 朱玉宝 20031218 地区：福州
    If mlng险类 = TYPE_福建巨龙 Or mlng险类 = TYPE_福建省 Or mlng险类 = TYPE_福州市 Or mlng险类 = TYPE_南平市 Or mlng险类 = TYPE_重庆银海版 Then
        'Modified by 朱玉宝 20031218 地区：福州
        txtEdit(text编码).MaxLength = 20
        cmdADD.Enabled = False
        CmdClear.Enabled = False
        CmdDel.Enabled = False
    End If
    
    mblnChange = False
    frm保险病种编辑.Show vbModal, frm保险病种
    编辑病种 = mblnOK
End Function

Private Sub txt封顶线_Change()
    mblnChange = True
End Sub

Private Sub txt封顶线_GotFocus()
    zlControl.TxtSelAll txt封顶线
End Sub

Private Sub txt封顶线_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 '使之不响
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Function GetMaxCode() As String
'功能：读取指定表的本级编码的最大值
'返回：成功返回 下级最大编码; 否者返回 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim lngLengh As Long
    
    On Error GoTo ErrHand
    With rsTemp
        gstrSQL = "SELECT max(length(substr(名称,1,instr(名称,'@@')-1))) as 最长值 FROM 保险病种 where 险类=" & mlng险类
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            GetMaxCode = "1"
            Exit Function
        Else
            lngLengh = Nvl(rsTemp("最长值"), "1")
        End If
        
        gstrSQL = "SELECT MAX(LPAD(substr(名称,1,instr(名称,'@@')-1)," & lngLengh & ",' ')) as 最大值 FROM 保险病种 where 险类=" & mlng险类
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF Then
            GetMaxCode = Format(1, String(lngLengh, "0"))
            Exit Function
        End If
        
        varTemp = Nvl(rsTemp("最大值"), "0")
        If IsNumeric(varTemp) Then
            GetMaxCode = CStr(Val(varTemp) + 1)
            GetMaxCode = Format(GetMaxCode, String(lngLengh, "0"))
        Else
            GetMaxCode = Mid(varTemp, 1, Len(varTemp) - 1) & Chr(asc(Right(varTemp, 1)) + 1)
            GetMaxCode = Trim(GetMaxCode)
        End If
        .Close
    End With
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function CheckCode(ByVal strCode As String, Optional ByVal blnNew As Boolean = True) As Boolean
    Dim rsCode As New ADODB.Recordset
    '因为编码超长，只有将编码与名称保存在名称列，而编码列实际保存的是记录数，在用户修改编码时，需要判断编码是否重复
    
    CheckCode = False
    gstrSQL = "Select 1 From 保险病种 Where 险类=" & mlng险类 & " And substr(名称,1,instr(名称,'@@')-1)='" & strCode & "'" & IIf(blnNew, "", " And ID<>" & mstrID)
    Call OpenRecordset(rsCode, "判断编码是否重复")
    
    If Not rsCode.EOF Then
        MsgBox "保险病种编码重复！", vbInformation, gstrSysName
        txtEdit(text编码).SetFocus
        Exit Function
    End If
    CheckCode = True
End Function
